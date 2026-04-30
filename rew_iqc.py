from __future__ import annotations
"""
REW IQC Pass/Fail Tool
======================
Factory floor incoming quality control for speakers via REW 5.40+ API.

Connects to REW's localhost REST API, pulls frequency response and
distortion data, and evaluates against configurable limit masks for
magnitude, THD, and HOHD (Higher-Order Harmonic Distortion).

Requirements:
    pip3 install requests numpy matplotlib

Usage:
    # Auto-measure mode (recommended, requires REW Pro license):
    python3 rew_iqc.py --limits limits/speaker.json --auto

    # Manual mode (sweep in REW, evaluate here):
    python3 rew_iqc.py --limits limits/speaker.json

    # Single measurement check by index or UUID:
    python3 rew_iqc.py --measurement 1 --limits limits/speaker.json

    # Batch mode (evaluate all loaded measurements):
    python3 rew_iqc.py --batch --limits limits/speaker.json --report

    # Generate a starter limit mask:
    python3 rew_iqc.py --create-example-mask limits/speaker.json

Architecture:
    This script is a thin REST client. All audio I/O, sweep generation,
    signal processing, and FFT computation happen inside REW. The script
    reads processed frequency response and distortion data via the API
    and applies limit-mask pass/fail logic on top.

    REW API docs (Swagger UI): http://localhost:4735 (when API is running)
    REW API spec: http://localhost:4735/doc.json
"""

__version__ = "1.3.1"

import argparse
import base64
import csv
import json
import logging
import os
import struct
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union

import numpy as np
import requests

try:
    import matplotlib
    # Use non-interactive backend so plots save to disk without a GUI window.
    # Change to "TkAgg" if you want interactive pop-up plots.
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.patches import FancyBboxPatch
    HAS_MATPLOTLIB = True
except ImportError:
    HAS_MATPLOTLIB = False

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# REW API defaults — override with --host / --port CLI args
REW_HOST = "http://127.0.0.1"
REW_PORT = 4735
REW_BASE = "{}:{}".format(REW_HOST, REW_PORT)

# Output directories (created automatically)
PLOT_DIR = Path("iqc_plots")
REPORT_DIR = Path("iqc_reports")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("rew_iqc")


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------

@dataclass
class LimitMask:
    """Upper/lower frequency-dependent limit envelope with optional THD/HOHD limits.

    The mask defines an acceptable corridor for a speaker's frequency
    response. Frequencies and dB values are specified at anchor points;
    the tool interpolates linearly between them when comparing against
    a measurement.

    THD limits define the maximum acceptable total harmonic distortion
    (in percent relative to the fundamental) at anchor frequencies. A single
    value applies as a flat ceiling; multiple points define a frequency-
    dependent limit that the tool interpolates between.

    HOHD limits (Higher-Order Harmonic Distortion) work the same way as
    THD limits but cover a separate, user-selected band of harmonics
    (typically H10-H15). HOHD is computed by sqrt-sum-of-squares of the
    selected harmonic columns from REW's per-harmonic data.

    Attributes:
        name:            Display name (shown on plots and reports)
        version:         Version string for traceability
        freq_hz:         Anchor frequencies in Hz (monotonically increasing)
        upper_db:        Upper dB SPL limit at each anchor frequency
        lower_db:        Lower dB SPL limit at each anchor frequency
        smoothing:       Smoothing to request from REW (e.g. "1/12", "1/3")
        ppo:             Points per octave for the frequency response data
        freq_range:      (low_hz, high_hz) — evaluation window for magnitude
        thd_freq_hz:     Anchor frequencies for THD limits (Hz)
        thd_max_pct:     Maximum THD (%) at each anchor frequency
                         (e.g. 3.0 = THD must be below 3%)
        thd_freq_range:  (low_hz, high_hz) — evaluation window for THD
        thd_ppo:         Points per octave for distortion data from REW
        thd_harmonics:   Harmonics expected in REW's THD aggregate, info-only
                         (default ["H2".."H9"]). REW already aggregates THD
                         in its column 2 output, so this is metadata.
        hohd_freq_hz:    Anchor frequencies for HOHD limits (Hz)
        hohd_max_pct:    Maximum HOHD (%) at each anchor frequency
        hohd_freq_range: (low_hz, high_hz) — evaluation window for HOHD
        hohd_ppo:        Points per octave for distortion data (HOHD)
        hohd_harmonics:  Harmonic columns to aggregate into HOHD via
                         sqrt(sum(Hn^2)). Default ["H10".."H15"].
        metadata:        Optional dict for part numbers, spec docs, notes
    """
    name: str
    version: str
    freq_hz: np.ndarray
    upper_db: np.ndarray
    lower_db: np.ndarray
    smoothing: str = "1/12"
    ppo: int = 48
    freq_range: Tuple[float, float] = (100, 20000)
    # THD limits (optional — if thd_freq_hz is empty, THD is not checked)
    thd_freq_hz: np.ndarray = field(default_factory=lambda: np.array([]))
    thd_max_pct: np.ndarray = field(default_factory=lambda: np.array([]))
    thd_freq_range: Tuple[float, float] = (200, 10000)
    thd_ppo: int = 12
    thd_harmonics: List[str] = field(
        default_factory=lambda: ["H2", "H3", "H4", "H5", "H6", "H7", "H8", "H9"]
    )
    # HOHD limits (optional — if hohd_freq_hz is empty, HOHD is not checked)
    hohd_freq_hz: np.ndarray = field(default_factory=lambda: np.array([]))
    hohd_max_pct: np.ndarray = field(default_factory=lambda: np.array([]))
    hohd_freq_range: Tuple[float, float] = (200, 8000)
    hohd_ppo: int = 12
    hohd_harmonics: List[str] = field(
        default_factory=lambda: ["H10", "H11", "H12", "H13", "H14", "H15"]
    )
    metadata: Dict = field(default_factory=dict)

    @property
    def has_thd_limits(self) -> bool:
        return len(self.thd_freq_hz) > 0

    @property
    def has_hohd_limits(self) -> bool:
        return len(self.hohd_freq_hz) > 0

    def check_magnitude(self, freq_hz: np.ndarray, mag_db: np.ndarray):
        """Check if a frequency response falls within the magnitude envelope.

        Returns:
            (passed: bool, details: dict)
        """
        f_lo, f_hi = self.freq_range

        # Only evaluate within the defined frequency range
        idx = (freq_hz >= f_lo) & (freq_hz <= f_hi)
        f = freq_hz[idx]
        m = mag_db[idx]

        # Interpolate the mask anchor points onto the measurement's
        # frequency axis for point-by-point comparison
        upper = np.interp(f, self.freq_hz, self.upper_db)
        lower = np.interp(f, self.freq_hz, self.lower_db)

        over = m > upper   # points above the upper limit
        under = m < lower  # points below the lower limit

        violations = []
        if np.any(over):
            worst_idx = np.argmax(m - upper)
            violations.append({
                "type": "UPPER",
                "freq_hz": float(f[np.where(over)[0][0]]),
                "worst_freq_hz": float(f[worst_idx]),
                "worst_delta_db": float((m - upper)[worst_idx]),
                "count": int(np.sum(over)),
            })
        if np.any(under):
            worst_idx = np.argmin(m - lower)
            violations.append({
                "type": "LOWER",
                "freq_hz": float(f[np.where(under)[0][0]]),
                "worst_freq_hz": float(f[worst_idx]),
                "worst_delta_db": float((m - lower)[worst_idx]),
                "count": int(np.sum(under)),
            })

        passed = len(violations) == 0
        return passed, {
            "passed": passed,
            "violations": violations,
            "eval_range": (f_lo, f_hi),
            "points_evaluated": int(np.sum(idx)),
        }

    def check_thd(self, thd_freq_hz: np.ndarray, thd_pct: np.ndarray):
        """Check if THD falls below the maximum limit at each frequency.

        Args:
            thd_freq_hz: Frequency axis from REW distortion data (Hz)
            thd_pct:     THD values in percent (e.g. 3.0 = 3%)

        Returns:
            (passed: bool, details: dict)
        """
        if not self.has_thd_limits:
            return True, {"passed": True, "violations": [], "points_evaluated": 0}

        f_lo, f_hi = self.thd_freq_range

        # Only evaluate within the THD frequency range
        idx = (thd_freq_hz >= f_lo) & (thd_freq_hz <= f_hi)
        f = thd_freq_hz[idx]
        t = thd_pct[idx]

        # Interpolate the THD limit onto the measurement's frequency axis
        limit = np.interp(f, self.thd_freq_hz, self.thd_max_pct)

        # A violation means THD exceeds the limit
        # Example: THD = 5.0%, limit = 3.0% → 5.0 > 3.0 → FAIL
        over = t > limit

        violations = []
        if np.any(over):
            worst_idx = np.argmax(t - limit)
            violations.append({
                "type": "THD",
                "freq_hz": float(f[np.where(over)[0][0]]),
                "worst_freq_hz": float(f[worst_idx]),
                "worst_thd_pct": float(t[worst_idx]),
                "worst_limit_pct": float(limit[worst_idx]),
                "worst_delta_pct": float(t[worst_idx] - limit[worst_idx]),
                "count": int(np.sum(over)),
            })

        passed = len(violations) == 0
        return passed, {
            "passed": passed,
            "violations": violations,
            "eval_range": (f_lo, f_hi),
            "points_evaluated": int(np.sum(idx)),
        }

    def check_hohd(self, hohd_freq_hz: np.ndarray, hohd_pct: np.ndarray):
        """Check if HOHD falls below the maximum limit at each frequency.

        HOHD (Higher-Order Harmonic Distortion) is computed upstream
        from REW's per-harmonic data via sqrt(sum(Hn^2)) over the
        harmonics listed in mask.hohd_harmonics.

        Args:
            hohd_freq_hz: Frequency axis from REW distortion data (Hz)
            hohd_pct:     Aggregate HOHD values in percent

        Returns:
            (passed: bool, details: dict)
        """
        if not self.has_hohd_limits:
            return True, {"passed": True, "violations": [], "points_evaluated": 0}

        f_lo, f_hi = self.hohd_freq_range
        idx = (hohd_freq_hz >= f_lo) & (hohd_freq_hz <= f_hi)
        f = hohd_freq_hz[idx]
        h = hohd_pct[idx]

        limit = np.interp(f, self.hohd_freq_hz, self.hohd_max_pct)
        over = h > limit

        violations = []
        if np.any(over):
            worst_idx = np.argmax(h - limit)
            violations.append({
                "type": "HOHD",
                "freq_hz": float(f[np.where(over)[0][0]]),
                "worst_freq_hz": float(f[worst_idx]),
                "worst_hohd_pct": float(h[worst_idx]),
                "worst_limit_pct": float(limit[worst_idx]),
                "worst_delta_pct": float(h[worst_idx] - limit[worst_idx]),
                "count": int(np.sum(over)),
            })

        passed = len(violations) == 0
        return passed, {
            "passed": passed,
            "violations": violations,
            "eval_range": (f_lo, f_hi),
            "points_evaluated": int(np.sum(idx)),
        }


@dataclass
class IQCResult:
    """Container for a single DUT evaluation result.

    Holds everything needed for reporting and plotting: the raw
    measurement data, the pass/fail verdict, violation details,
    and the path to any saved plot image.
    """
    timestamp: str
    measurement_name: str
    measurement_uuid: str
    serial_number: str
    limit_mask_name: str
    passed: bool
    mag_passed: bool
    thd_passed: bool
    mag_details: Dict
    thd_details: Dict
    freq_hz: np.ndarray
    mag_db: np.ndarray
    thd_freq_hz: np.ndarray      # empty if no distortion data
    thd_pct: np.ndarray           # THD in percent; empty if no distortion data
    hohd_passed: bool = True
    hohd_details: Dict = field(default_factory=lambda: {"passed": True, "violations": [], "points_evaluated": 0})
    hohd_freq_hz: np.ndarray = field(default_factory=lambda: np.array([]))
    hohd_pct: np.ndarray = field(default_factory=lambda: np.array([]))
    plot_path: Optional[str] = None


# ---------------------------------------------------------------------------
# REW API Client
# ---------------------------------------------------------------------------

class REWClient:
    """Thin wrapper around the REW V5.40+ REST API.

    REW's API runs on localhost (default port 4735) and provides full
    access to measurements, audio settings, and measurement control.

    API notes discovered during development:
        - GET /measurements returns a dict keyed by index strings
          ("1", "2", ...), NOT a list
        - GET /measurements/{id}/frequency-response returns magnitude
          data under the key "magnitude" (singular), not "magnitudes"
        - GET /measurements/selected-uuid returns a bare quoted string
        - GET /measurements/{id}/distortion returns columnHeaders and
          a 2D data array; columns include Freq, Fundamental, THD,
          Noise, and H2-H9 (all in the requested unit)
        - POST /measure/command accepts {"command": "SPL"} to trigger
          a sweep (requires Pro license)
        - Blocking mode may not reliably block for measurement commands,
          so we poll for completion by watching the measurement count
    """

    def __init__(self, base_url: str = REW_BASE, timeout: float = 10.0):
        self.base = base_url.rstrip("/")
        self.timeout = timeout
        self.session = requests.Session()

    def _get(self, path: str, **params):
        """Send a GET request to the REW API."""
        url = "{}{}".format(self.base, path)
        r = self.session.get(url, params=params, timeout=self.timeout)
        r.raise_for_status()
        return r.json() if r.content else None

    def _post(self, path: str, data=None):
        """Send a POST request to the REW API."""
        url = "{}{}".format(self.base, path)
        r = self.session.post(url, json=data, timeout=self.timeout)
        r.raise_for_status()
        return r.json() if r.content else None

    # -- Connection ---------------------------------------------------------

    def ping(self) -> bool:
        """Check if the REW API is reachable."""
        try:
            self._get("/application/commands")
            return True
        except Exception:
            return False

    # -- Measurements -------------------------------------------------------

    def list_measurements(self) -> Dict:
        """Return the measurements dict keyed by index string ("1", "2", ...).

        Each value is a measurement summary dict with fields like
        title, uuid, date, startFreq, endFreq, sampleRate, etc.
        """
        return self._get("/measurements") or {}

    def get_measurement(self, id_or_uuid: str) -> Dict:
        """Get the summary for a single measurement by index or UUID."""
        return self._get("/measurements/{}".format(id_or_uuid))

    def get_selected_uuid(self) -> str:
        """Get the UUID of the currently selected measurement in REW.

        The API returns a bare JSON string (e.g. '"abc-123"'),
        so we strip any extra quotes.
        """
        raw = self._get("/measurements/selected-uuid")
        if isinstance(raw, str):
            return raw.strip('"')
        return str(raw)

    def get_measurement_count(self) -> int:
        """Return the number of measurements currently loaded in REW."""
        return len(self.list_measurements())

    def get_latest_measurement(self) -> Tuple[str, Dict]:
        """Return (index_str, summary_dict) for the highest-numbered measurement."""
        measurements = self.list_measurements()
        if not measurements:
            return ("", {})
        highest_key = max(measurements.keys(), key=lambda k: int(k))
        return (highest_key, measurements[highest_key])

    def get_frequency_response(
        self,
        id_or_uuid: str,
        smoothing: str = "1/12",
        ppo: int = 48,
        unit: str = "SPL",
    ) -> Tuple[np.ndarray, np.ndarray, np.ndarray]:
        """Fetch frequency response data from REW.

        Returns:
            (freq_hz, magnitude_db, phase_deg) as numpy arrays.
            Phase may be empty if not available for this measurement.

        REW returns magnitude data as a Base64-encoded big-endian
        float32 array. The frequency axis is reconstructed from the
        startFreq and ppo (or freqStep) values in the response.
        """
        params = {"smoothing": smoothing, "ppo": ppo, "unit": unit}
        data = self._get("/measurements/{}/frequency-response".format(id_or_uuid), **params)

        # REW V5.40 betas use "magnitude" (singular); handle both forms
        mag_key = "magnitude" if "magnitude" in data else "magnitudes"
        mag = self._decode_base64_floats(data.get(mag_key, ""))

        # Phase may or may not be present in the response
        phase_key = "phase" if "phase" in data else "phases"
        phase_raw = data.get(phase_key, "")
        phase = self._decode_base64_floats(phase_raw) if phase_raw else np.array([])

        # Reconstruct the frequency axis from API metadata
        start_freq = data["startFreq"]
        if "ppo" in data and data["ppo"]:
            n = len(mag)
            ratio = 2.0 ** (1.0 / data["ppo"])
            freq = start_freq * (ratio ** np.arange(n))
        elif "freqStep" in data and data["freqStep"]:
            freq = start_freq + data["freqStep"] * np.arange(len(mag))
        else:
            freq = start_freq * (2.0 ** (np.arange(len(mag)) / ppo))

        return freq, mag, phase

    def get_distortion(
        self,
        id_or_uuid: str,
        unit: str = "percent",
        ppo: int = 12,
    ) -> Tuple[np.ndarray, np.ndarray, List[str]]:
        """Fetch THD distortion data from REW (legacy convenience wrapper).

        Args:
            id_or_uuid: Measurement index or UUID
            unit:       "percent" or "dBr" (relative to fundamental)
            ppo:        Points per octave for sweep distortion data

        Returns:
            (freq_hz, thd_values, column_headers) where:
                - freq_hz:    frequency axis from column 0
                - thd_values: THD from column 2 (in requested unit)
                - column_headers: all column names for reference

        For per-harmonic access (needed for HOHD), use get_distortion_full().
        """
        freq, harmonics, headers = self.get_distortion_full(
            id_or_uuid, unit=unit, ppo=ppo
        )
        thd = harmonics.get("THD", np.array([]))
        return freq, thd, headers

    def get_distortion_full(
        self,
        id_or_uuid: str,
        unit: str = "percent",
        ppo: int = 12,
    ) -> Tuple[np.ndarray, Dict[str, np.ndarray], List[str]]:
        """Fetch full distortion data including per-harmonic columns.

        Returns:
            (freq_hz, harmonics, column_headers) where:
                - freq_hz:    frequency axis (Hz)
                - harmonics:  dict keyed by short column name like
                              "Fundamental", "THD", "Noise", "H2", "H3", ...
                              Each value is a numpy array aligned with freq_hz.
                              If REW didn't supply a harmonic at a particular
                              frequency, that entry is NaN.
                - column_headers: original column header strings from REW

        REW returns ragged rows (some frequencies may lack higher harmonics).
        We allocate the full grid as NaN and fill in available cells.
        """
        params = {"unit": unit, "ppo": ppo}
        data = self._get("/measurements/{}/distortion".format(id_or_uuid), **params)

        headers = data.get("columnHeaders", [])
        rows = data.get("data", [])

        if not rows or not headers:
            return np.array([]), {}, headers

        # Map full header strings to short keys (e.g. "H2 (dBr)" -> "H2",
        # "THD (%)" -> "THD", "Fundamental (dB)" -> "Fundamental")
        short_keys = []
        for h in headers:
            # Take the part before any space-paren (units block)
            if "(" in h:
                key = h.split("(")[0].strip()
            else:
                key = h.strip()
            short_keys.append(key)

        # Allocate freq + per-column buffers
        freq_list = []
        n_cols = len(headers)
        # First pass: collect valid frequencies
        valid_rows = []
        for row in rows:
            if row is None or len(row) < 1:
                continue
            f = row[0]
            if f is None:
                continue
            try:
                freq_list.append(float(f))
                valid_rows.append(row)
            except (TypeError, ValueError):
                continue

        if not freq_list:
            return np.array([]), {}, headers

        freq = np.array(freq_list, dtype=np.float64)
        n_pts = len(freq)

        # Initialize per-column arrays as NaN; skip the first column (freq)
        harmonics: Dict[str, np.ndarray] = {}
        for c in range(1, n_cols):
            harmonics[short_keys[c]] = np.full(n_pts, np.nan, dtype=np.float64)

        # Fill in values where present
        for i, row in enumerate(valid_rows):
            for c in range(1, n_cols):
                if c >= len(row):
                    continue
                v = row[c]
                if v is None:
                    continue
                try:
                    harmonics[short_keys[c]][i] = float(v)
                except (TypeError, ValueError):
                    pass

        return freq, harmonics, headers

    # -- Measurement control (requires REW Pro license) ---------------------

    def enable_blocking(self, enable: bool = True):
        """Enable/disable blocking mode for API commands."""
        self._post("/application/blocking", enable)

    def measure_spl(self, timeout: float = 60.0) -> str:
        """Trigger an SPL sweep and wait for it to complete.

        Uses the sweep settings already configured in REW's Measure dialog.
        Polls the measurement count to detect completion.

        Returns:
            UUID of the newly created measurement

        Raises:
            TimeoutError: if no new measurement appears within timeout
        """
        count_before = self.get_measurement_count()
        self._post("/measure/command", {"command": "SPL"})

        poll_interval = 0.5
        elapsed = 0.0
        while elapsed < timeout:
            time.sleep(poll_interval)
            elapsed += poll_interval
            count_after = self.get_measurement_count()
            if count_after > count_before:
                _, newest = self.get_latest_measurement()
                return newest.get("uuid", "")

        raise TimeoutError(
            "Measurement did not complete within {} seconds".format(timeout)
        )

    # -- Helpers ------------------------------------------------------------

    @staticmethod
    def _decode_base64_floats(b64: str) -> np.ndarray:
        """Decode REW's Base64-encoded big-endian float32 arrays."""
        if not b64:
            return np.array([], dtype=np.float32)
        raw = base64.b64decode(b64)
        count = len(raw) // 4
        return np.array(
            struct.unpack(">{}f".format(count), raw),
            dtype=np.float32,
        )


# ---------------------------------------------------------------------------
# Distortion math helpers
# ---------------------------------------------------------------------------

def aggregate_harmonics_pct(
    harmonics: Dict[str, np.ndarray],
    selected: List[str],
) -> Tuple[np.ndarray, List[str], List[str]]:
    """Aggregate selected harmonic columns into a single distortion %.

    Combines harmonics by sqrt(sum of squares), the standard total-distortion
    formula, when the values are already in percent.

    Args:
        harmonics: dict from REWClient.get_distortion_full(), keyed like
                   "H2", "H3", ..., values are np.ndarray of % per frequency
                   (with NaN where REW didn't supply that harmonic)
        selected:  list of harmonic short names the user wants aggregated,
                   e.g. ["H10", "H11", "H12", "H13", "H14", "H15"]

    Returns:
        (aggregate_pct, found, missing) where:
            - aggregate_pct: np.ndarray of aggregate distortion % per
                             frequency point (NaN where no selected
                             harmonics were available)
            - found:   list of harmonic names actually used
            - missing: list of harmonic names that were requested but not
                       present in REW's data (may be empty)

    Notes:
        - When a particular harmonic has NaN at some frequencies (REW
          ragged data), those NaNs are treated as 0 in the sum so the
          aggregate uses whatever is available.
        - If no requested harmonic exists in the data at all, returns an
          empty array.
    """
    found = [h for h in selected if h in harmonics]
    missing = [h for h in selected if h not in harmonics]

    if not found:
        return np.array([]), found, missing

    # Stack the selected harmonic % values; replace NaN with 0 for the sum.
    stack = np.vstack([np.nan_to_num(harmonics[h], nan=0.0) for h in found])
    # sqrt(sum of squares) = sqrt(stack ** 2).sum(axis=0)
    aggregate = np.sqrt(np.sum(stack ** 2, axis=0))

    # If every selected harmonic was NaN at a given frequency, mark NaN.
    presence = np.vstack([~np.isnan(harmonics[h]) for h in found])
    no_data = ~np.any(presence, axis=0)
    aggregate[no_data] = np.nan

    return aggregate, found, missing


# ---------------------------------------------------------------------------
# Limit mask I/O
# ---------------------------------------------------------------------------

def load_limit_mask(path: Union[str, Path]) -> LimitMask:
    """Load a limit mask from a JSON file.

    The JSON can optionally include a "thd_limits" section for THD
    evaluation, and an "hohd_limits" section for HOHD (higher-order
    harmonic distortion). If a section is omitted, that check is skipped.
    """
    p = Path(path)
    with open(p, "r") as f:
        d = json.load(f)

    pts = d["limits"]
    # Upper/lower at each freq are individually optional. A point with only
    # `upper_db` means "no lower bound here" (effectively -inf); a point with
    # only `lower_db` means "no upper bound here" (effectively +inf). Missing
    # both is equivalent to "no constraint at this frequency" — harmless but
    # pointless. check_magnitude treats inf bounds as never-violated.
    freq = np.array([pt["freq_hz"] for pt in pts], dtype=np.float64)
    upper = np.array(
        [pt.get("upper_db", np.inf) for pt in pts], dtype=np.float64
    )
    lower = np.array(
        [pt.get("lower_db", -np.inf) for pt in pts], dtype=np.float64
    )

    # THD limits are optional
    thd_freq = np.array([])
    thd_max = np.array([])
    thd_range = (200, 10000)
    thd_ppo = 12
    thd_harmonics = ["H2", "H3", "H4", "H5", "H6", "H7", "H8", "H9"]
    if "thd_limits" in d:
        thd_section = d["thd_limits"]
        thd_pts = thd_section.get("limits", [])
        thd_freq = np.array([pt["freq_hz"] for pt in thd_pts], dtype=np.float64)
        thd_max = np.array(
            [pt.get("max_thd_pct", np.inf) for pt in thd_pts], dtype=np.float64
        )
        thd_range = tuple(thd_section.get("freq_range_hz", [200, 10000]))
        thd_ppo = thd_section.get("ppo", 12)
        thd_harmonics = thd_section.get("harmonics", thd_harmonics)

    # HOHD limits are optional
    hohd_freq = np.array([])
    hohd_max = np.array([])
    hohd_range = (200, 8000)
    hohd_ppo = 12
    hohd_harmonics = ["H10", "H11", "H12", "H13", "H14", "H15"]
    if "hohd_limits" in d:
        hohd_section = d["hohd_limits"]
        hohd_pts = hohd_section.get("limits", [])
        hohd_freq = np.array([pt["freq_hz"] for pt in hohd_pts], dtype=np.float64)
        hohd_max = np.array(
            [pt.get("max_hohd_pct", np.inf) for pt in hohd_pts], dtype=np.float64
        )
        hohd_range = tuple(hohd_section.get("freq_range_hz", [200, 8000]))
        hohd_ppo = hohd_section.get("ppo", 12)
        hohd_harmonics = hohd_section.get("harmonics", hohd_harmonics)

    return LimitMask(
        name=d.get("name", p.stem),
        version=d.get("version", "0"),
        freq_hz=freq,
        upper_db=upper,
        lower_db=lower,
        smoothing=d.get("smoothing", "1/12"),
        ppo=d.get("ppo", 48),
        freq_range=tuple(d.get("freq_range_hz", [100, 20000])),
        thd_freq_hz=thd_freq,
        thd_max_pct=thd_max,
        thd_freq_range=thd_range,
        thd_ppo=thd_ppo,
        thd_harmonics=thd_harmonics,
        hohd_freq_hz=hohd_freq,
        hohd_max_pct=hohd_max,
        hohd_freq_range=hohd_range,
        hohd_ppo=hohd_ppo,
        hohd_harmonics=hohd_harmonics,
        metadata=d.get("metadata", {}),
    )


def create_example_limit_mask(path: Union[str, Path]):
    """Generate an example speaker limit mask JSON file with THD + HOHD limits.

    Creates a generic envelope centered around ~75 dB SPL at 1 kHz
    with +/-5 dB tolerance in the midband, a THD ceiling of 3% in the
    midband relaxing to 10% at the frequency extremes, and a HOHD ceiling
    that's tighter (since higher-order distortion is typically lower).
    """
    freqs = [200, 300, 500, 700, 1000, 1500, 2000, 3000, 4000, 5000,
             6000, 8000, 10000, 12000, 14000, 16000]
    nominal = [68, 71, 74, 75, 76, 76, 76, 75, 74, 73, 72, 70, 68, 65, 62, 58]
    tolerance_upper = [8, 7, 6, 6, 5, 5, 5, 6, 6, 7, 7, 8, 9, 10, 12, 14]
    tolerance_lower = [8, 7, 6, 6, 5, 5, 5, 6, 6, 7, 7, 8, 9, 10, 12, 14]

    limits = []
    for i, f in enumerate(freqs):
        limits.append({
            "freq_hz": f,
            "upper_db": nominal[i] + tolerance_upper[i],
            "lower_db": nominal[i] - tolerance_lower[i],
        })

    mask = {
        "name": "Speaker IQC - Example",
        "version": "1.2",
        "smoothing": "1/12",
        "ppo": 48,
        "freq_range_hz": [200, 16000],
        "limits": limits,
        "thd_limits": {
            "freq_range_hz": [200, 10000],
            "ppo": 12,
            "harmonics": ["H2", "H3", "H4", "H5", "H6", "H7", "H8", "H9"],
            "limits": [
                {"freq_hz": 200,   "max_thd_pct": 10.0},
                {"freq_hz": 500,   "max_thd_pct": 5.0},
                {"freq_hz": 1000,  "max_thd_pct": 3.0},
                {"freq_hz": 2000,  "max_thd_pct": 3.0},
                {"freq_hz": 5000,  "max_thd_pct": 5.0},
                {"freq_hz": 10000, "max_thd_pct": 10.0}
            ]
        },
        "hohd_limits": {
            "freq_range_hz": [200, 8000],
            "ppo": 12,
            "harmonics": ["H10", "H11", "H12", "H13", "H14", "H15"],
            "limits": [
                {"freq_hz": 200,   "max_hohd_pct": 2.0},
                {"freq_hz": 500,   "max_hohd_pct": 1.0},
                {"freq_hz": 1000,  "max_hohd_pct": 0.5},
                {"freq_hz": 2000,  "max_hohd_pct": 0.5},
                {"freq_hz": 5000,  "max_hohd_pct": 1.0},
                {"freq_hz": 8000,  "max_hohd_pct": 2.0}
            ]
        },
        "metadata": {
            "part_number": "SPK-EXAMPLE",
            "notes": "Example mask with magnitude, THD, and HOHD limits. "
                     "THD uses REW's pre-aggregated THD column (H2-H9 by "
                     "default in REW). HOHD is computed by sqrt-sum-of-"
                     "squares of H10-H15; this requires REW to be configured "
                     "to report higher harmonics. Adjust to your DUT population."
        }
    }

    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    with open(p, "w") as f:
        json.dump(mask, f, indent=2)
    log.info("Example limit mask written to {}".format(p))


# ---------------------------------------------------------------------------
# Plotting
# ---------------------------------------------------------------------------

def plot_result(
    result: IQCResult,
    mask: LimitMask,
    save_path: Optional[Union[str, Path]] = None,
    show: bool = False,
) -> Optional[str]:
    """Generate a pass/fail plot with magnitude and optional THD panels.

    If the mask has THD limits and distortion data is available, the
    plot has two panels: magnitude on top, THD on the bottom. Otherwise
    it's a single-panel magnitude plot.
    """
    if not HAS_MATPLOTLIB:
        log.warning("matplotlib not available, skipping plot")
        return None

    has_thd = mask.has_thd_limits and len(result.thd_freq_hz) > 0 and len(result.thd_pct) > 0
    has_hohd = mask.has_hohd_limits and len(result.hohd_freq_hz) > 0 and len(result.hohd_pct) > 0

    n_panels = 1 + (1 if has_thd else 0) + (1 if has_hohd else 0)

    if n_panels == 3:
        fig, (ax_mag, ax_thd, ax_hohd) = plt.subplots(
            3, 1, figsize=(14, 11), height_ratios=[3, 2, 2], sharex=True,
        )
    elif n_panels == 2 and has_thd:
        fig, (ax_mag, ax_thd) = plt.subplots(
            2, 1, figsize=(14, 9), height_ratios=[3, 2], sharex=True,
        )
        ax_hohd = None
    elif n_panels == 2 and has_hohd:
        # HOHD-only second panel (rare, but possible)
        fig, (ax_mag, ax_hohd) = plt.subplots(
            2, 1, figsize=(14, 9), height_ratios=[3, 2], sharex=True,
        )
        ax_thd = None
    else:
        fig, ax_mag = plt.subplots(figsize=(14, 7))
        ax_thd = None
        ax_hohd = None

    # =====================================================================
    # TOP PANEL: Magnitude frequency response
    # =====================================================================

    # Limit mask envelope
    # Bounds may contain +/-inf at frequencies where one side wasn't defined
    # (e.g. an upper-only mask, or a sparse-data mask where one curve ran
    # out before the other). For plotting we substitute "off-screen" values
    # so fill_between still works and auto-scaling isn't pulled to infinity.
    mask_freq = mask.freq_hz
    finite_mask_vals = []
    if np.any(np.isfinite(mask.upper_db)):
        finite_mask_vals.extend(mask.upper_db[np.isfinite(mask.upper_db)].tolist())
    if np.any(np.isfinite(mask.lower_db)):
        finite_mask_vals.extend(mask.lower_db[np.isfinite(mask.lower_db)].tolist())
    if finite_mask_vals:
        finite_lo = min(finite_mask_vals)
        finite_hi = max(finite_mask_vals)
    else:
        finite_lo, finite_hi = 0.0, 100.0
    plot_upper = np.where(np.isfinite(mask.upper_db), mask.upper_db, finite_hi + 50)
    plot_lower = np.where(np.isfinite(mask.lower_db), mask.lower_db, finite_lo - 50)

    ax_mag.fill_between(mask_freq, plot_lower, plot_upper,
                         alpha=0.15, color="#2196F3", label="Limit mask")
    # Only draw the boundary line where it's actually defined
    if np.any(np.isfinite(mask.upper_db)):
        ax_mag.plot(mask_freq, np.where(np.isfinite(mask.upper_db),
                                         mask.upper_db, np.nan),
                    color="#1565C0", linewidth=1.5, linestyle="--", alpha=0.7)
    if np.any(np.isfinite(mask.lower_db)):
        ax_mag.plot(mask_freq, np.where(np.isfinite(mask.lower_db),
                                         mask.lower_db, np.nan),
                    color="#1565C0", linewidth=1.5, linestyle="--", alpha=0.7)

    # Measurement trace (green = pass, red = fail for overall result)
    trace_color = "#4CAF50" if result.passed else "#F44336"
    ax_mag.semilogx(result.freq_hz, result.mag_db, color=trace_color,
                     linewidth=2, label=result.measurement_name)

    # Magnitude violation markers
    if not result.mag_passed:
        for v in result.mag_details.get("violations", []):
            ax_mag.axvline(x=v["worst_freq_hz"], color="#F44336", alpha=0.3,
                           linestyle=":")
            if v["type"] == "UPPER":
                vlabel = "UPPER +{:.1f} dB".format(v["worst_delta_db"])
            else:
                vlabel = "LOWER {:.1f} dB".format(v["worst_delta_db"])
            ax_mag.annotate(
                vlabel,
                xy=(v["worst_freq_hz"], 0), xycoords=("data", "axes fraction"),
                fontsize=9, color="#F44336", ha="center", va="bottom",
                fontweight="bold",
            )

    # X axis: 100 Hz to 20 kHz
    ax_mag.set_xlim(100, 20000)

    # Y axis: auto-scale to fit both measurement and mask
    x_lo, x_hi = 100, 20000
    in_range = (result.freq_hz >= x_lo) & (result.freq_hz <= x_hi)
    data_in_range = result.mag_db[in_range]

    mask_in_range = (mask.freq_hz >= x_lo) & (mask.freq_hz <= x_hi)
    mask_upper_in_range = mask.upper_db[mask_in_range]
    mask_lower_in_range = mask.lower_db[mask_in_range]
    # Drop infinities before computing y-range
    mask_upper_finite = mask_upper_in_range[np.isfinite(mask_upper_in_range)]
    mask_lower_finite = mask_lower_in_range[np.isfinite(mask_lower_in_range)]

    all_mins = []
    all_maxs = []
    if len(mask_lower_finite) > 0:
        all_mins.append(np.min(mask_lower_finite))
    if len(mask_upper_finite) > 0:
        all_maxs.append(np.max(mask_upper_finite))
    if len(data_in_range) > 0:
        all_mins.append(np.min(data_in_range))
        all_maxs.append(np.max(data_in_range))

    y_lo = min(all_mins) - 5 if all_mins else 0
    y_hi = max(all_maxs) + 5 if all_maxs else 100
    ax_mag.set_ylim(y_lo, y_hi)

    # Large solid PASS/FAIL badge (top-left corner)
    verdict = "PASS" if result.passed else "FAIL"
    box_color = "#4CAF50" if result.passed else "#F44336"

    box = FancyBboxPatch(
        (0.02, 0.72), 0.36, 0.25,
        boxstyle="round,pad=0.02",
        facecolor=box_color,
        edgecolor="white",
        linewidth=3,
        alpha=0.92,
        transform=ax_mag.transAxes,
        zorder=10,
    )
    ax_mag.add_patch(box)
    ax_mag.text(
        0.20, 0.845, verdict,
        transform=ax_mag.transAxes, fontsize=64, fontweight="bold",
        color="white", ha="center", va="center",
        fontfamily="monospace", zorder=11,
    )

    # Build title with optional serial number
    if result.serial_number:
        title_str = "IQC: {} [SN: {}]  |  Mask: {} v{}  |  {}".format(
            result.measurement_name, result.serial_number,
            mask.name, mask.version, result.timestamp
        )
    else:
        title_str = "IQC: {}  |  Mask: {} v{}  |  {}".format(
            result.measurement_name, mask.name, mask.version, result.timestamp
        )

    # Magnitude panel labels
    ax_mag.set_ylabel("Magnitude (dB SPL)", fontsize=12)
    ax_mag.set_title(title_str, fontsize=12, fontweight="bold")
    ax_mag.legend(loc="lower left", fontsize=10)
    ax_mag.grid(True, which="both", alpha=0.3)

    if n_panels == 1:
        ax_mag.set_xlabel("Frequency (Hz)", fontsize=12)

    # =====================================================================
    # MIDDLE PANEL: THD (only if mask has THD limits and data exists)
    # =====================================================================

    if has_thd and ax_thd is not None:
        # THD limit line
        ax_thd.semilogx(mask.thd_freq_hz, mask.thd_max_pct,
                         color="#1565C0", linewidth=2, linestyle="--",
                         label="THD limit", alpha=0.8)

        # Fill above the limit to show the fail zone
        thd_fill_top = max(np.max(mask.thd_max_pct) * 3, 20)
        ax_thd.fill_between(mask.thd_freq_hz, mask.thd_max_pct, thd_fill_top,
                             alpha=0.08, color="#F44336")

        # THD measurement trace
        thd_color = "#4CAF50" if result.thd_passed else "#F44336"
        ax_thd.semilogx(result.thd_freq_hz, result.thd_pct,
                         color=thd_color, linewidth=2,
                         label="THD ({})".format(result.measurement_name))

        # THD violation markers
        if not result.thd_passed:
            for v in result.thd_details.get("violations", []):
                ax_thd.axvline(x=v["worst_freq_hz"], color="#F44336",
                               alpha=0.3, linestyle=":")
                ax_thd.annotate(
                    "THD {:.1f}% (limit {:.1f}%)".format(
                        v["worst_thd_pct"], v["worst_limit_pct"]
                    ),
                    xy=(v["worst_freq_hz"], v["worst_thd_pct"]),
                    fontsize=9, color="#F44336", fontweight="bold",
                    ha="center", va="bottom",
                )

        # THD panel formatting
        ax_thd.set_xlim(100, 20000)
        # Bottom panel gets the x-label only if there's no HOHD panel below
        if not has_hohd:
            ax_thd.set_xlabel("Frequency (Hz)", fontsize=12)
        ax_thd.set_ylabel("THD (%)", fontsize=12)
        ax_thd.legend(loc="upper left", fontsize=10)
        ax_thd.grid(True, which="both", alpha=0.3)

        # Auto-scale THD Y axis (0 to a sensible max)
        thd_lo, thd_hi = mask.thd_freq_range
        thd_in_range = (result.thd_freq_hz >= thd_lo) & (result.thd_freq_hz <= thd_hi)
        if np.any(thd_in_range):
            thd_data_max = np.max(result.thd_pct[thd_in_range])
            thd_limit_max = np.max(mask.thd_max_pct)
            y_top = max(thd_data_max, thd_limit_max) * 1.3
            ax_thd.set_ylim(0, max(y_top, 1.0))

        # Small pass/fail badge for THD panel
        thd_verdict = "THD PASS" if result.thd_passed else "THD FAIL"
        thd_badge_color = "#4CAF50" if result.thd_passed else "#F44336"
        ax_thd.text(
            0.98, 0.95, thd_verdict,
            transform=ax_thd.transAxes, fontsize=14, fontweight="bold",
            color=thd_badge_color, ha="right", va="top",
            fontfamily="monospace",
            bbox=dict(boxstyle="round,pad=0.3", facecolor="white",
                      edgecolor=thd_badge_color, alpha=0.9),
        )

    # =====================================================================
    # BOTTOM PANEL: HOHD (only if mask has HOHD limits and data exists)
    # =====================================================================

    if has_hohd and ax_hohd is not None:
        # HOHD limit line
        ax_hohd.semilogx(mask.hohd_freq_hz, mask.hohd_max_pct,
                          color="#6A1B9A", linewidth=2, linestyle="--",
                          label="HOHD limit", alpha=0.8)

        # Fill above the limit to show the fail zone
        hohd_fill_top = max(np.max(mask.hohd_max_pct) * 3, 5)
        ax_hohd.fill_between(mask.hohd_freq_hz, mask.hohd_max_pct, hohd_fill_top,
                              alpha=0.08, color="#F44336")

        # HOHD measurement trace
        hohd_color = "#4CAF50" if result.hohd_passed else "#F44336"
        harmonics_label = ", ".join(mask.hohd_harmonics)
        ax_hohd.semilogx(result.hohd_freq_hz, result.hohd_pct,
                          color=hohd_color, linewidth=2,
                          label="HOHD [{}]".format(harmonics_label))

        # HOHD violation markers
        if not result.hohd_passed:
            for v in result.hohd_details.get("violations", []):
                ax_hohd.axvline(x=v["worst_freq_hz"], color="#F44336",
                                 alpha=0.3, linestyle=":")
                ax_hohd.annotate(
                    "HOHD {:.2f}% (limit {:.2f}%)".format(
                        v["worst_hohd_pct"], v["worst_limit_pct"]
                    ),
                    xy=(v["worst_freq_hz"], v["worst_hohd_pct"]),
                    fontsize=9, color="#F44336", fontweight="bold",
                    ha="center", va="bottom",
                )

        # HOHD panel formatting
        ax_hohd.set_xlim(100, 20000)
        ax_hohd.set_xlabel("Frequency (Hz)", fontsize=12)
        ax_hohd.set_ylabel("HOHD (%)", fontsize=12)
        ax_hohd.legend(loc="upper left", fontsize=10)
        ax_hohd.grid(True, which="both", alpha=0.3)

        # Auto-scale HOHD Y axis
        hohd_lo, hohd_hi = mask.hohd_freq_range
        hohd_in_range = (result.hohd_freq_hz >= hohd_lo) & (result.hohd_freq_hz <= hohd_hi)
        if np.any(hohd_in_range):
            hohd_data_max = np.max(result.hohd_pct[hohd_in_range])
            hohd_limit_max = np.max(mask.hohd_max_pct)
            y_top = max(hohd_data_max, hohd_limit_max) * 1.3
            ax_hohd.set_ylim(0, max(y_top, 0.5))

        # Small pass/fail badge for HOHD panel
        hohd_verdict = "HOHD PASS" if result.hohd_passed else "HOHD FAIL"
        hohd_badge_color = "#4CAF50" if result.hohd_passed else "#F44336"
        ax_hohd.text(
            0.98, 0.95, hohd_verdict,
            transform=ax_hohd.transAxes, fontsize=14, fontweight="bold",
            color=hohd_badge_color, ha="right", va="top",
            fontfamily="monospace",
            bbox=dict(boxstyle="round,pad=0.3", facecolor="white",
                      edgecolor=hohd_badge_color, alpha=0.9),
        )

    fig.tight_layout()

    if save_path:
        p = Path(save_path)
        p.parent.mkdir(parents=True, exist_ok=True)
        fig.savefig(p, dpi=150, bbox_inches="tight")
        log.info("Plot saved: {}".format(p))

    if show:
        plt.show()
    else:
        plt.close(fig)

    return str(save_path) if save_path else None


def open_plot(plot_path: str):
    """Open a plot image in the default viewer (Preview on macOS)."""
    full_path = os.path.abspath(plot_path)
    log.info("Opening plot: {}".format(full_path))
    os.system('/usr/bin/open "{}"'.format(full_path))


# ---------------------------------------------------------------------------
# Report generation
# ---------------------------------------------------------------------------

def write_csv_report(results: List[IQCResult], path: Union[str, Path]):
    """Append IQC results to a daily CSV report file.

    Columns: timestamp, serial_number, measurement_name, uuid, limit_mask,
             result, mag_result, thd_result, hohd_result, violations_summary, plot_file
    """
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    file_exists = p.exists()

    with open(p, "a", newline="") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow([
                "timestamp", "serial_number", "measurement_name", "uuid",
                "limit_mask", "result", "mag_result", "thd_result", "hohd_result",
                "violations_summary", "plot_file",
            ])
        for r in results:
            # Combine magnitude, THD, and HOHD violations into one summary
            all_violations = (
                r.mag_details.get("violations", []) +
                r.thd_details.get("violations", []) +
                r.hohd_details.get("violations", [])
            )
            vsummary_parts = []
            for v in all_violations:
                if v["type"] in ("UPPER", "LOWER"):
                    vsummary_parts.append(
                        "{type} worst={wf:.0f}Hz delta={wd:.1f}dB (n={c})".format(
                            type=v["type"], wf=v["worst_freq_hz"],
                            wd=v["worst_delta_db"], c=v["count"]
                        )
                    )
                elif v["type"] == "THD":
                    vsummary_parts.append(
                        "THD worst={wf:.0f}Hz thd={td:.1f}% limit={tl:.1f}% (n={c})".format(
                            wf=v["worst_freq_hz"], td=v["worst_thd_pct"],
                            tl=v["worst_limit_pct"], c=v["count"]
                        )
                    )
                elif v["type"] == "HOHD":
                    vsummary_parts.append(
                        "HOHD worst={wf:.0f}Hz hohd={hd:.2f}% limit={hl:.2f}% (n={c})".format(
                            wf=v["worst_freq_hz"], hd=v["worst_hohd_pct"],
                            hl=v["worst_limit_pct"], c=v["count"]
                        )
                    )
            vsummary = "; ".join(vsummary_parts) or "none"

            writer.writerow([
                r.timestamp, r.serial_number, r.measurement_name,
                r.measurement_uuid, r.limit_mask_name,
                "PASS" if r.passed else "FAIL",
                "PASS" if r.mag_passed else "FAIL",
                "PASS" if r.thd_passed else "FAIL",
                "PASS" if r.hohd_passed else "FAIL",
                vsummary, r.plot_path or "",
            ])
    log.info("Report updated: {}".format(p))


# ---------------------------------------------------------------------------
# Core IQC engine
# ---------------------------------------------------------------------------

class IQCEngine:
    """Orchestrates the evaluation pipeline: fetch -> compare -> report."""

    def __init__(self, rew: REWClient, mask: LimitMask):
        self.rew = rew
        self.mask = mask
        self.results = []  # type: List[IQCResult]

    def check_measurement(
        self,
        id_or_uuid: str,
        serial_number: str = "",
        save_plot: bool = True,
        show_plot: bool = False,
    ) -> IQCResult:
        """Evaluate a single measurement against magnitude, THD, and HOHD limits."""
        summary = self.rew.get_measurement(id_or_uuid)
        name = summary.get("title", "meas_{}".format(id_or_uuid))
        uuid = summary.get("uuid", str(id_or_uuid))

        label = "{} [SN: {}]".format(name, serial_number) if serial_number else name
        log.info("Checking: {} (UUID: {})".format(label, uuid))

        # --- Fetch magnitude data ---
        freq, mag, _phase = self.rew.get_frequency_response(
            id_or_uuid,
            smoothing=self.mask.smoothing,
            ppo=self.mask.ppo,
        )

        # --- Evaluate magnitude ---
        mag_passed, mag_details = self.mask.check_magnitude(freq, mag)

        # --- Fetch distortion data once if any distortion check is needed ---
        thd_freq = np.array([])
        thd_pct = np.array([])
        thd_passed = True
        thd_details = {"passed": True, "violations": [], "points_evaluated": 0}

        hohd_freq = np.array([])
        hohd_pct = np.array([])
        hohd_passed = True
        hohd_details = {"passed": True, "violations": [], "points_evaluated": 0}

        need_distortion = self.mask.has_thd_limits or self.mask.has_hohd_limits
        if need_distortion:
            # Use the higher of the two PPO values so a single fetch serves both.
            dist_ppo = max(self.mask.thd_ppo, self.mask.hohd_ppo) if (
                self.mask.has_thd_limits and self.mask.has_hohd_limits
            ) else (self.mask.hohd_ppo if self.mask.has_hohd_limits else self.mask.thd_ppo)
            try:
                dist_freq, harmonics, _headers = self.rew.get_distortion_full(
                    id_or_uuid, unit="percent", ppo=dist_ppo,
                )

                # --- Evaluate THD using REW's pre-aggregated THD column ---
                if self.mask.has_thd_limits:
                    thd_col = harmonics.get("THD", np.array([]))
                    if len(dist_freq) > 0 and len(thd_col) > 0:
                        # Strip any NaN points from THD (REW returns ragged
                        # data — some frequencies have valid THD, some don't)
                        valid = ~np.isnan(thd_col)
                        thd_freq = dist_freq[valid]
                        thd_pct = thd_col[valid]
                        if len(thd_freq) > 0:
                            thd_passed, thd_details = self.mask.check_thd(
                                thd_freq, thd_pct
                            )
                        else:
                            log.warning("  No valid THD data points available")
                    else:
                        log.warning("  THD column missing from REW response")

                # --- Evaluate HOHD by aggregating selected harmonic columns ---
                if self.mask.has_hohd_limits:
                    aggregate, found, missing = aggregate_harmonics_pct(
                        harmonics, self.mask.hohd_harmonics
                    )
                    if missing:
                        log.warning(
                            "  HOHD: {} not present in REW data (configure REW "
                            "to report higher harmonics for full HOHD coverage)"
                            .format(", ".join(missing))
                        )
                    if len(aggregate) > 0:
                        # Drop any all-NaN points (no harmonic data at all)
                        valid = ~np.isnan(aggregate)
                        hohd_freq = dist_freq[valid]
                        hohd_pct = aggregate[valid]
                        if len(hohd_freq) > 0 and found:
                            log.info("  HOHD aggregating: {}".format(
                                ", ".join(found)
                            ))
                            hohd_passed, hohd_details = self.mask.check_hohd(
                                hohd_freq, hohd_pct
                            )
                        else:
                            log.warning("  No HOHD data available — skipping HOHD check")
                            hohd_passed = True  # don't fail if no data
                    else:
                        log.warning(
                            "  HOHD: none of the requested harmonics ({}) "
                            "are present in REW data".format(
                                ", ".join(self.mask.hohd_harmonics)
                            )
                        )
                        hohd_passed = True
            except Exception as e:
                log.warning("  Could not fetch distortion data: {}".format(e))

        # --- Overall verdict: must pass magnitude AND THD AND HOHD ---
        passed = mag_passed and thd_passed and hohd_passed
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        result = IQCResult(
            timestamp=ts,
            measurement_name=name,
            measurement_uuid=uuid,
            serial_number=serial_number,
            limit_mask_name="{} v{}".format(self.mask.name, self.mask.version),
            passed=passed,
            mag_passed=mag_passed,
            thd_passed=thd_passed,
            mag_details=mag_details,
            thd_details=thd_details,
            freq_hz=freq,
            mag_db=mag,
            thd_freq_hz=thd_freq,
            thd_pct=thd_pct,
            hohd_passed=hohd_passed,
            hohd_details=hohd_details,
            hohd_freq_hz=hohd_freq,
            hohd_pct=hohd_pct,
        )

        # Generate and save the plot
        if save_plot or show_plot:
            safe_name = "".join(
                c if c.isalnum() or c in "-_ " else "_" for c in name
            )
            if serial_number:
                safe_sn = "".join(
                    c if c.isalnum() or c in "-_" else "_" for c in serial_number
                )
                plot_name = "{}_{}_{}".format(
                    safe_sn, safe_name, ts.replace(":", "").replace(" ", "_")
                )
            else:
                plot_name = "{}_{}".format(
                    safe_name, ts.replace(":", "").replace(" ", "_")
                )
            plot_file = PLOT_DIR / "{}.png".format(plot_name)
            result.plot_path = plot_result(
                result, self.mask, plot_file, show=show_plot
            )

        self.results.append(result)

        # Log the verdict
        log.info("  Magnitude: {}".format("PASS" if mag_passed else "FAIL"))
        if not mag_passed:
            for v in mag_details["violations"]:
                log.info(
                    "    {} violation: worst at {:.0f} Hz, "
                    "delta {:+.1f} dB ({} points)".format(
                        v["type"], v["worst_freq_hz"],
                        v["worst_delta_db"], v["count"]
                    )
                )
        if self.mask.has_thd_limits:
            log.info("  THD: {}".format("PASS" if thd_passed else "FAIL"))
            if not thd_passed:
                for v in thd_details["violations"]:
                    log.info(
                        "    THD violation: worst at {:.0f} Hz, "
                        "THD={:.1f}%, limit={:.1f}% ({} points)".format(
                            v["worst_freq_hz"], v["worst_thd_pct"],
                            v["worst_limit_pct"], v["count"]
                        )
                    )
        if self.mask.has_hohd_limits:
            log.info("  HOHD: {}".format("PASS" if hohd_passed else "FAIL"))
            if not hohd_passed:
                for v in hohd_details["violations"]:
                    log.info(
                        "    HOHD violation: worst at {:.0f} Hz, "
                        "HOHD={:.2f}%, limit={:.2f}% ({} points)".format(
                            v["worst_freq_hz"], v["worst_hohd_pct"],
                            v["worst_limit_pct"], v["count"]
                        )
                    )
        # Log the verdict
        log.info("  Overall: {}{}".format(
            "PASS" if passed else "FAIL",
            " [SN: {}]".format(serial_number) if serial_number else ""
        ))

        return result

    def check_all(self, save_plots: bool = True) -> List[IQCResult]:
        """Evaluate all measurements currently loaded in REW."""
        measurements = self.rew.list_measurements()
        log.info("Found {} measurements in REW".format(len(measurements)))
        results = []
        for idx_str, m in measurements.items():
            uuid = m.get("uuid", idx_str)
            r = self.check_measurement(uuid, save_plot=save_plots)
            results.append(r)
        return results

    def save_report(self, path: Optional[Union[str, Path]] = None):
        """Write accumulated results to a CSV report."""
        if not self.results:
            log.warning("No results to report")
            return
        if path is None:
            ts = datetime.now().strftime("%Y%m%d")
            path = REPORT_DIR / "iqc_report_{}.csv".format(ts)
        write_csv_report(self.results, path)


# ---------------------------------------------------------------------------
# Operator workflow (interactive mode)
# ---------------------------------------------------------------------------

def operator_loop(mask_path: str, show_plots: bool = False, auto_measure: bool = False):
    """Interactive operator loop for factory floor use."""
    rew = REWClient()
    if not rew.ping():
        log.error("Cannot reach REW API at {}".format(REW_BASE))
        log.error("Make sure REW 5.40+ is running with -api flag")
        sys.exit(1)

    log.info("Connected to REW at {}".format(REW_BASE))

    if auto_measure:
        try:
            rew.enable_blocking(True)
            log.info("Blocking mode enabled (auto-measure active)")
        except Exception as e:
            log.error("Could not enable blocking mode: {}".format(e))
            log.error("Auto-measure requires REW Pro license")
            sys.exit(1)

    mask = load_limit_mask(mask_path)
    log.info("Loaded limit mask: {} v{}".format(mask.name, mask.version))
    log.info("  Magnitude eval range: {}-{} Hz".format(
        mask.freq_range[0], mask.freq_range[1]))
    if mask.has_thd_limits:
        log.info("  THD eval range: {}-{} Hz".format(
            mask.thd_freq_range[0], mask.thd_freq_range[1]))
    else:
        log.info("  THD limits: not defined")
    if mask.has_hohd_limits:
        log.info("  HOHD eval range: {}-{} Hz (harmonics: {})".format(
            mask.hohd_freq_range[0], mask.hohd_freq_range[1],
            ", ".join(mask.hohd_harmonics)))
    else:
        log.info("  HOHD limits: not defined")

    engine = IQCEngine(rew, mask)
    ts = datetime.now().strftime("%Y%m%d")
    report_path = REPORT_DIR / "iqc_report_{}.csv".format(ts)

    mode_label = "AUTO-MEASURE" if auto_measure else "MANUAL"
    check_parts = ["Magnitude"]
    if mask.has_thd_limits:
        check_parts.append("THD")
    if mask.has_hohd_limits:
        check_parts.append("HOHD")
    checks = " + ".join(check_parts)
    print("\n" + "=" * 60)
    print("  REW IQC TOOL -- OPERATOR MODE ({})".format(mode_label))
    print("  Limit mask: {} v{}".format(mask.name, mask.version))
    print("  Checks:     {}".format(checks))
    print("  Report:     {}".format(report_path))
    print("=" * 60)
    if auto_measure:
        print("\nWorkflow:")
        print("  1. Load DUT into fixture")
        print("  2. Enter serial number (or press ENTER to skip)")
        print("  3. Sweep runs automatically, then evaluates")
        print("  4. Type 'q' to quit\n")
    else:
        print("\nWorkflow:")
        print("  1. Load DUT into fixture, run sweep in REW")
        print("  2. Enter serial number (or press ENTER to skip)")
        print("  3. Evaluates the selected measurement")
        print("  4. Type 'q' to quit\n")

    while True:
        user = input(">> Enter serial number (or ENTER to skip, 'q' to quit): ").strip()
        if user.lower() == "q":
            break
        serial_number = user

        try:
            if auto_measure:
                log.info("Triggering SPL measurement...")
                print("  Measuring...")
                new_uuid = rew.measure_spl(timeout=60)
                log.info("Measurement complete (UUID: {})".format(new_uuid))
                uuid = new_uuid
            else:
                uuid = rew.get_selected_uuid()

            result = engine.check_measurement(uuid, serial_number=serial_number,
                                                save_plot=True, show_plot=show_plots)

            # Print big PASS/FAIL to console
            sn_label = " [SN: {}]".format(serial_number) if serial_number else ""
            if result.passed:
                print("\n" + "=" * 40)
                print("       PASS  PASS  PASS  PASS  PASS")
                print("       ====  ====  ====  ====  ====")
                if sn_label:
                    print("      {}".format(sn_label))
                print("=" * 40 + "\n")
            else:
                print("\n" + "=" * 40)
                print("       FAIL  FAIL  FAIL  FAIL  FAIL")
                print("       ====  ====  ====  ====  ====")
                if sn_label:
                    print("      {}".format(sn_label))
                print("=" * 40)
                # Print magnitude violations
                for v in result.mag_details.get("violations", []):
                    print("  {}: worst {:.0f} Hz ({:+.1f} dB), {} pts".format(
                        v["type"], v["worst_freq_hz"],
                        v["worst_delta_db"], v["count"]
                    ))
                # Print THD violations
                for v in result.thd_details.get("violations", []):
                    print("  THD: worst {:.0f} Hz ({:.1f}%, limit {:.1f}%), {} pts".format(
                        v["worst_freq_hz"], v["worst_thd_pct"],
                        v["worst_limit_pct"], v["count"]
                    ))
                # Print HOHD violations
                for v in result.hohd_details.get("violations", []):
                    print("  HOHD: worst {:.0f} Hz ({:.2f}%, limit {:.2f}%), {} pts".format(
                        v["worst_freq_hz"], v["worst_hohd_pct"],
                        v["worst_limit_pct"], v["count"]
                    ))
                print()

            if result.plot_path:
                open_plot(result.plot_path)

            engine.save_report(report_path)

        except requests.exceptions.ConnectionError:
            log.error("Lost connection to REW -- is it still running?")
        except Exception as e:
            log.error("Error: {}".format(e))

    print("\nSession complete. {} DUTs tested.".format(len(engine.results)))
    if engine.results:
        pass_count = sum(1 for r in engine.results if r.passed)
        fail_count = len(engine.results) - pass_count
        print("  PASS: {}  |  FAIL: {}  |  Yield: {:.1f}%".format(
            pass_count, fail_count, pass_count / len(engine.results) * 100
        ))
        print("  Report: {}".format(report_path))


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="REW IQC Pass/Fail Tool v{} -- speaker incoming quality control "
                    "via REW 5.40+ REST API (magnitude + THD + serial tracking)".format(__version__)
    )
    parser.add_argument(
        "--limits", "-l", type=str, default=None,
        help="Path to limit mask JSON file",
    )
    parser.add_argument(
        "--measurement", "-m", type=str, default=None,
        help="Measurement index or UUID to check (default: selected)",
    )
    parser.add_argument(
        "--batch", "-b", action="store_true",
        help="Check all loaded measurements",
    )
    parser.add_argument(
        "--report", "-r", action="store_true",
        help="Write CSV report",
    )
    parser.add_argument(
        "--auto", "-a", action="store_true",
        help="Auto-trigger measurements (requires REW Pro license)",
    )
    parser.add_argument(
        "--show-plots", action="store_true",
        help="Display plots interactively (default: save only)",
    )
    parser.add_argument(
        "--create-example-mask", type=str, default=None,
        metavar="PATH",
        help="Generate an example limit mask JSON and exit",
    )
    parser.add_argument(
        "--host", type=str, default=None,
        help="REW API host (default: {})".format(REW_HOST),
    )
    parser.add_argument(
        "--port", type=int, default=None,
        help="REW API port (default: {})".format(REW_PORT),
    )
    args = parser.parse_args()

    global REW_BASE
    host = args.host or REW_HOST
    port = args.port or REW_PORT
    REW_BASE = "{}:{}".format(host, port)

    if args.create_example_mask:
        create_example_limit_mask(args.create_example_mask)
        return

    if args.measurement is None and not args.batch:
        if args.limits is None:
            print("No limit mask specified. Generate an example with:")
            print("  python3 rew_iqc.py --create-example-mask limits/speaker.json")
            print("\nThen run:")
            print("  python3 rew_iqc.py --limits limits/speaker.json")
            return
        operator_loop(args.limits, show_plots=args.show_plots, auto_measure=args.auto)
        return

    if args.limits is None:
        parser.error("--limits is required for measurement checks")

    rew = REWClient(REW_BASE)
    if not rew.ping():
        log.error("Cannot reach REW API at {}".format(REW_BASE))
        sys.exit(1)

    mask = load_limit_mask(args.limits)
    engine = IQCEngine(rew, mask)

    if args.batch:
        engine.check_all(save_plots=True)
    else:
        engine.check_measurement(args.measurement, save_plot=True, show_plot=args.show_plots)

    if args.report:
        engine.save_report()

    if any(not r.passed for r in engine.results):
        sys.exit(1)


if __name__ == "__main__":
    main()