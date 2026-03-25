from __future__ import annotations
"""
REW IQC Pass/Fail Tool
======================
Factory floor incoming quality control for speakers via REW 5.40+ API.

Connects to REW's localhost REST API, pulls frequency response and
distortion data, and evaluates against configurable limit masks
for both magnitude and THD.

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

__version__ = "1.1.0"

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
    """Upper/lower frequency-dependent limit envelope with optional THD limits.

    The mask defines an acceptable corridor for a speaker's frequency
    response. Frequencies and dB values are specified at anchor points;
    the tool interpolates linearly between them when comparing against
    a measurement.

    THD limits define the maximum acceptable total harmonic distortion
    (in percent relative to the fundamental) at anchor frequencies. A single
    value applies as a flat ceiling; multiple points define a frequency-
    dependent limit that the tool interpolates between.

    Attributes:
        name:           Display name (shown on plots and reports)
        version:        Version string for traceability
        freq_hz:        Anchor frequencies in Hz (monotonically increasing)
        upper_db:       Upper dB SPL limit at each anchor frequency
        lower_db:       Lower dB SPL limit at each anchor frequency
        smoothing:      Smoothing to request from REW (e.g. "1/12", "1/3")
        ppo:            Points per octave for the frequency response data
        freq_range:     (low_hz, high_hz) — evaluation window for magnitude
        thd_freq_hz:    Anchor frequencies for THD limits (Hz)
        thd_max_pct:    Maximum THD (%) at each anchor frequency
                        (e.g. 3.0 = THD must be below 3%)
        thd_freq_range: (low_hz, high_hz) — evaluation window for THD
        thd_ppo:        Points per octave for distortion data from REW
        metadata:       Optional dict for part numbers, spec docs, notes
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
    metadata: Dict = field(default_factory=dict)

    @property
    def has_thd_limits(self) -> bool:
        return len(self.thd_freq_hz) > 0

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
        """Fetch distortion data from REW.

        Args:
            id_or_uuid: Measurement index or UUID
            unit:       "percent" or "dBr" (relative to fundamental)
            ppo:        Points per octave for sweep distortion data

        Returns:
            (freq_hz, thd_values, column_headers) where:
                - freq_hz:    frequency axis from column 0
                - thd_values: THD from column 2 (in requested unit)
                - column_headers: all column names for reference

        The full data array also contains Fundamental, Noise, and
        individual harmonics H2-H9, but we extract just THD here.
        Additional columns can be accessed by extending this method.
        """
        params = {"unit": unit, "ppo": ppo}
        data = self._get("/measurements/{}/distortion".format(id_or_uuid), **params)

        headers = data.get("columnHeaders", [])
        rows = data.get("data", [])

        if not rows:
            return np.array([]), np.array([]), headers

        # Some rows may have fewer columns (missing harmonics at certain
        # frequencies) or contain None values. Extract just the columns
        # we need (0=freq, 2=THD) and skip any rows with missing data.
        freq_list = []
        thd_list = []
        for row in rows:
            if row is None or len(row) < 3:
                continue
            f = row[0]
            t = row[2]
            if f is not None and t is not None:
                freq_list.append(float(f))
                thd_list.append(float(t))

        freq = np.array(freq_list, dtype=np.float64)
        thd = np.array(thd_list, dtype=np.float64)

        return freq, thd, headers

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
# Limit mask I/O
# ---------------------------------------------------------------------------

def load_limit_mask(path: Union[str, Path]) -> LimitMask:
    """Load a limit mask from a JSON file.

    The JSON can optionally include a "thd_limits" section for THD
    evaluation. If omitted, only magnitude is checked.
    """
    p = Path(path)
    with open(p, "r") as f:
        d = json.load(f)

    pts = d["limits"]
    freq = np.array([pt["freq_hz"] for pt in pts], dtype=np.float64)
    upper = np.array([pt["upper_db"] for pt in pts], dtype=np.float64)
    lower = np.array([pt["lower_db"] for pt in pts], dtype=np.float64)

    # THD limits are optional
    thd_freq = np.array([])
    thd_max = np.array([])
    thd_range = (200, 10000)
    thd_ppo = 12
    if "thd_limits" in d:
        thd_section = d["thd_limits"]
        thd_pts = thd_section.get("limits", [])
        thd_freq = np.array([pt["freq_hz"] for pt in thd_pts], dtype=np.float64)
        thd_max = np.array([pt["max_thd_pct"] for pt in thd_pts], dtype=np.float64)
        thd_range = tuple(thd_section.get("freq_range_hz", [200, 10000]))
        thd_ppo = thd_section.get("ppo", 12)

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
        metadata=d.get("metadata", {}),
    )


def create_example_limit_mask(path: Union[str, Path]):
    """Generate an example speaker limit mask JSON file with THD limits.

    Creates a generic envelope centered around ~75 dB SPL at 1 kHz
    with +/-5 dB tolerance in the midband, plus a THD ceiling of
    3% in the midband relaxing to 10% at the frequency extremes.
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
        "version": "1.1",
        "smoothing": "1/12",
        "ppo": 48,
        "freq_range_hz": [200, 16000],
        "limits": limits,
        "thd_limits": {
            "freq_range_hz": [200, 10000],
            "ppo": 12,
            "limits": [
                {"freq_hz": 200,   "max_thd_pct": 10.0},
                {"freq_hz": 500,   "max_thd_pct": 5.0},
                {"freq_hz": 1000,  "max_thd_pct": 3.0},
                {"freq_hz": 2000,  "max_thd_pct": 3.0},
                {"freq_hz": 5000,  "max_thd_pct": 5.0},
                {"freq_hz": 10000, "max_thd_pct": 10.0}
            ]
        },
        "metadata": {
            "part_number": "SPK-EXAMPLE",
            "notes": "Example mask with magnitude and THD limits. "
                     "THD limit is relaxed at LF and HF where distortion "
                     "is naturally higher (10% at extremes, 3% midband). "
                     "Adjust to your DUT population."
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

    if has_thd:
        # Two-panel plot: magnitude on top (taller), THD on bottom
        fig, (ax_mag, ax_thd) = plt.subplots(
            2, 1, figsize=(14, 9), height_ratios=[3, 2],
            sharex=True,
        )
    else:
        fig, ax_mag = plt.subplots(figsize=(14, 7))
        ax_thd = None

    # =====================================================================
    # TOP PANEL: Magnitude frequency response
    # =====================================================================

    # Limit mask envelope
    mask_freq = mask.freq_hz
    ax_mag.fill_between(mask_freq, mask.lower_db, mask.upper_db,
                         alpha=0.15, color="#2196F3", label="Limit mask")
    ax_mag.plot(mask_freq, mask.upper_db, color="#1565C0", linewidth=1.5,
                linestyle="--", alpha=0.7)
    ax_mag.plot(mask_freq, mask.lower_db, color="#1565C0", linewidth=1.5,
                linestyle="--", alpha=0.7)

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

    all_mins = [np.min(mask_lower_in_range)]
    all_maxs = [np.max(mask_upper_in_range)]
    if len(data_in_range) > 0:
        all_mins.append(np.min(data_in_range))
        all_maxs.append(np.max(data_in_range))

    y_lo = min(all_mins) - 5
    y_hi = max(all_maxs) + 5
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

    # Magnitude panel labels
    ax_mag.set_ylabel("Magnitude (dB SPL)", fontsize=12)
    ax_mag.set_title(
        "IQC: {}  |  Mask: {} v{}  |  {}".format(
            result.measurement_name, mask.name, mask.version, result.timestamp
        ),
        fontsize=12, fontweight="bold",
    )
    ax_mag.legend(loc="lower left", fontsize=10)
    ax_mag.grid(True, which="both", alpha=0.3)

    if not has_thd:
        ax_mag.set_xlabel("Frequency (Hz)", fontsize=12)

    # =====================================================================
    # BOTTOM PANEL: THD (only if mask has THD limits and data exists)
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

    Columns: timestamp, measurement_name, uuid, limit_mask,
             result, mag_result, thd_result, violations_summary, plot_file
    """
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    file_exists = p.exists()

    with open(p, "a", newline="") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow([
                "timestamp", "measurement_name", "uuid", "limit_mask",
                "result", "mag_result", "thd_result",
                "violations_summary", "plot_file",
            ])
        for r in results:
            # Combine magnitude and THD violations into one summary
            all_violations = (
                r.mag_details.get("violations", []) +
                r.thd_details.get("violations", [])
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
            vsummary = "; ".join(vsummary_parts) or "none"

            writer.writerow([
                r.timestamp, r.measurement_name, r.measurement_uuid,
                r.limit_mask_name,
                "PASS" if r.passed else "FAIL",
                "PASS" if r.mag_passed else "FAIL",
                "PASS" if r.thd_passed else "FAIL",
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
        save_plot: bool = True,
        show_plot: bool = False,
    ) -> IQCResult:
        """Evaluate a single measurement against magnitude and THD limits."""
        summary = self.rew.get_measurement(id_or_uuid)
        name = summary.get("title", "meas_{}".format(id_or_uuid))
        uuid = summary.get("uuid", str(id_or_uuid))

        log.info("Checking: {} (UUID: {})".format(name, uuid))

        # --- Fetch magnitude data ---
        freq, mag, _phase = self.rew.get_frequency_response(
            id_or_uuid,
            smoothing=self.mask.smoothing,
            ppo=self.mask.ppo,
        )

        # --- Evaluate magnitude ---
        mag_passed, mag_details = self.mask.check_magnitude(freq, mag)

        # --- Fetch and evaluate THD (if mask has THD limits) ---
        thd_freq = np.array([])
        thd_pct = np.array([])
        thd_passed = True
        thd_details = {"passed": True, "violations": [], "points_evaluated": 0}

        if self.mask.has_thd_limits:
            try:
                thd_freq, thd_pct, _headers = self.rew.get_distortion(
                    id_or_uuid,
                    unit="percent",
                    ppo=self.mask.thd_ppo,
                )
                if len(thd_freq) > 0:
                    thd_passed, thd_details = self.mask.check_thd(thd_freq, thd_pct)
                else:
                    log.warning("  No distortion data available for this measurement")
            except Exception as e:
                log.warning("  Could not fetch distortion data: {}".format(e))

        # --- Overall verdict: must pass BOTH magnitude and THD ---
        passed = mag_passed and thd_passed
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        result = IQCResult(
            timestamp=ts,
            measurement_name=name,
            measurement_uuid=uuid,
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
        )

        # Generate and save the plot
        if save_plot or show_plot:
            safe_name = "".join(
                c if c.isalnum() or c in "-_ " else "_" for c in name
            )
            plot_file = PLOT_DIR / "{}_{}.png".format(
                safe_name, ts.replace(":", "").replace(" ", "_")
            )
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
        log.info("  Overall: {}".format("PASS" if passed else "FAIL"))

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
        log.info("  THD limits: not defined (magnitude only)")

    engine = IQCEngine(rew, mask)
    ts = datetime.now().strftime("%Y%m%d")
    report_path = REPORT_DIR / "iqc_report_{}.csv".format(ts)

    mode_label = "AUTO-MEASURE" if auto_measure else "MANUAL"
    checks = "Magnitude + THD" if mask.has_thd_limits else "Magnitude only"
    print("\n" + "=" * 60)
    print("  REW IQC TOOL -- OPERATOR MODE ({})".format(mode_label))
    print("  Limit mask: {} v{}".format(mask.name, mask.version))
    print("  Checks:     {}".format(checks))
    print("  Report:     {}".format(report_path))
    print("=" * 60)
    if auto_measure:
        print("\nWorkflow:")
        print("  1. Load DUT into fixture")
        print("  2. Press ENTER -- sweep runs automatically, then evaluates")
        print("  3. Type 'q' to quit\n")
    else:
        print("\nWorkflow:")
        print("  1. Load DUT into fixture, run sweep in REW")
        print("  2. Press ENTER to evaluate the selected measurement")
        print("  3. Type 'q' to quit\n")

    while True:
        user = input(">> Press ENTER to evaluate (or 'q' to quit): ").strip()
        if user.lower() == "q":
            break

        try:
            if auto_measure:
                log.info("Triggering SPL measurement...")
                print("  Measuring...")
                new_uuid = rew.measure_spl(timeout=60)
                log.info("Measurement complete (UUID: {})".format(new_uuid))
                uuid = new_uuid
            else:
                uuid = rew.get_selected_uuid()

            result = engine.check_measurement(uuid, save_plot=True, show_plot=show_plots)

            # Print big PASS/FAIL to console
            if result.passed:
                print("\n" + "=" * 40)
                print("       PASS  PASS  PASS  PASS  PASS")
                print("       ====  ====  ====  ====  ====")
                print("=" * 40 + "\n")
            else:
                print("\n" + "=" * 40)
                print("       FAIL  FAIL  FAIL  FAIL  FAIL")
                print("       ====  ====  ====  ====  ====")
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
                    "via REW 5.40+ REST API (magnitude + THD)".format(__version__)
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