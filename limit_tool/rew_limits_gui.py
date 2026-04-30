#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AudioMacGyver's REW Limit Tool
================================
A FabFilter-inspired GUI for building factory IQC PASS/FAIL limit masks
from REW (Room EQ Wizard) measurements. Designed to integrate with the
rew-iqc factory tool (https://github.com/audiomacgyver/rew-iqc).

Features:
  - Three tabs (FR / THD / HOHD) for building all three limit types in
    one session and exporting to a single combined JSON
  - Load REW text exports (.txt) or capture live from REW REST API
  - Four limit creation methods: anchor points, sigma, fixed offset, manual table
  - Three normalization modes: absolute, normalize-to-frequency, floating (FR only)
  - Limit shapes: upper+lower, upper-only, lower-only
  - Fractional-octave smoothing with energy-domain averaging
  - Build limits around the mean of all curves OR a specific reference unit
  - Test any DUT measurement against the limits with PASS/FAIL evaluation
  - Hybrid workflow: compute statistical limits then convert to draggable anchors
  - User-selectable harmonics for THD and HOHD (default H2-H9 / H10-H15)
  - Export combined FR + THD + HOHD limits to JSON (rew-iqc) or REW-importable .txt

Requirements:
  pip install PyQt5 numpy

Usage:
  python3 rew_limits_gui.py

Author: Jesse Lippert (audiomacgyver)
"""

__version__ = "1.1.0"
APP_TITLE = "AudioMacGyver's REW Limit Tool"

import sys
import os
import json
import math
import base64
import struct
import urllib.request
import urllib.error
import numpy as np
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLabel, QPushButton, QLineEdit, QGroupBox, QComboBox,
    QMessageBox, QFrame, QSizePolicy, QSpinBox, QDoubleSpinBox,
    QFileDialog, QTableWidget, QTableWidgetItem, QHeaderView,
    QSplitter, QListWidget, QListWidgetItem, QCheckBox, QRadioButton,
    QButtonGroup, QDialog, QDialogButtonBox, QAbstractItemView,
    QStackedWidget, QScrollArea, QTabWidget, QToolBar, QAction,
    QStatusBar
)
from PyQt5.QtCore import Qt, QRectF, pyqtSignal, QPointF
from PyQt5.QtGui import (
    QPainter, QColor, QLinearGradient, QPen, QFont, QBrush,
    QPainterPath, QPolygonF
)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Color palette (FabFilter-inspired dark theme)
TEXT_PRIMARY  = QColor(210, 210, 215)
TEXT_DIM      = QColor(120, 120, 130)
ACCENT        = QColor(0, 180, 220)
GRID_LINE     = QColor(50, 50, 60)
GRID_MAJOR    = QColor(70, 70, 80)
UPPER_LIMIT   = QColor(240, 80, 80)     # red
LOWER_LIMIT   = QColor(80, 180, 240)    # blue
MEAN_LINE     = QColor(240, 200, 40)    # yellow
ANCHOR_COLOR  = QColor(255, 255, 255)
RANGE_FILL    = QColor(0, 180, 220, 25)

# Per-measurement curve colors
CURVE_COLORS = [
    QColor(120, 200, 120),
    QColor(200, 120, 200),
    QColor(120, 200, 200),
    QColor(200, 200, 120),
    QColor(200, 160, 100),
    QColor(160, 100, 200),
    QColor(100, 160, 200),
    QColor(200, 100, 160),
]

# Plot frequency range (Hz)
FREQ_MIN = 20
FREQ_MAX = 20000

# Anchor click tolerance (pixels)
ANCHOR_HIT_RADIUS = 10

# Always show at least this much range on the value axis
MIN_SPL_RANGE_DB = 50      # FR uses dB units, ~50 dB span looks right
MIN_PCT_RANGE = 30         # THD/HOHD use % units, default ~30% span
MIN_HOHD_RANGE = 10        # HOHD typically much lower than THD

# REW REST API
DEFAULT_REW_HOST = 'http://localhost:4735'

# ---------------------------------------------------------------------------
# Tab kind configuration
# ---------------------------------------------------------------------------
# Each tab in the main window represents a different limit type. The kind
# string is used throughout the workspace to drive UI decisions: which data
# column to plot, which units to label, what default ranges to use, whether
# to show normalization controls, and so on.

KIND_FR   = 'FR'      # Frequency response (magnitude, dB SPL)
KIND_THD  = 'THD'     # Total Harmonic Distortion (%, REW-aggregated)
KIND_HOHD = 'HOHD'    # Higher-Order Harmonic Distortion (%, sqrt-sum-of-squares)

# Per-kind configuration. Mirrors the rew-iqc schema; values here populate
# defaults in the workspace UI and the export JSON.
KIND_CONFIG = {
    KIND_FR: {
        'unit_label': 'dB SPL',
        'unit_short': 'dB',
        'y_span': MIN_SPL_RANGE_DB,
        'y_default_min': -25,
        'y_floor': None,            # FR has no fixed lower bound (centered on data)
        'default_freq_range': (100, 20000),
        'default_shape': 'both',
        'show_normalization': True,
        'show_harmonics': False,
        'default_harmonics': [],
        'json_section_key': None,    # FR is at top level (limits, freq_range_hz, etc.)
        'json_max_field': None,      # not used (uses upper_db/lower_db)
        'default_smoothing': '1/12 octave',
        'default_ppo': 48,
    },
    KIND_THD: {
        'unit_label': '%',
        'unit_short': '%',
        'y_span': MIN_PCT_RANGE,
        'y_default_min': 0,
        'y_floor': 0,                # THD% can't go negative
        'default_freq_range': (200, 10000),
        'default_shape': 'upper',
        'show_normalization': False,
        'show_harmonics': True,
        'default_harmonics': ['H2', 'H3', 'H4', 'H5', 'H6', 'H7', 'H8', 'H9'],
        'json_section_key': 'thd_limits',
        'json_max_field': 'max_thd_pct',
        'default_smoothing': 'None',
        'default_ppo': 12,
    },
    KIND_HOHD: {
        'unit_label': '%',
        'unit_short': '%',
        'y_span': MIN_HOHD_RANGE,
        'y_default_min': 0,
        'y_floor': 0,
        'default_freq_range': (200, 8000),
        'default_shape': 'upper',
        'show_normalization': False,
        'show_harmonics': True,
        'default_harmonics': ['H10', 'H11', 'H12', 'H13', 'H14', 'H15'],
        'json_section_key': 'hohd_limits',
        'json_max_field': 'max_hohd_pct',
        'default_smoothing': 'None',
        'default_ppo': 12,
    },
}

STYLESHEET = """
QMainWindow { background-color: #19191e; }
QDialog { background-color: #19191e; }
QWidget { background-color: transparent; color: #d2d2d7; font-family: 'Segoe UI'; }
QGroupBox {
    background-color: #1e1e26;
    border: 1px solid #2a2a34;
    border-radius: 8px;
    margin-top: 14px;
    padding-top: 18px;
    font-size: 11px;
    font-weight: bold;
    color: #888;
}
QGroupBox::title { subcontrol-origin: margin; left: 14px; padding: 0 6px; }
QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox {
    background-color: #2a2a34;
    border: 1px solid #3a3a46;
    border-radius: 4px;
    padding: 4px 8px;
    color: #d2d2d7;
    font-size: 12px;
    selection-background-color: #00b4dc;
    min-height: 18px;
}
QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus, QComboBox:focus {
    border: 1px solid #00b4dc;
}
QPushButton {
    background-color: #2a2a34;
    border: 1px solid #3a3a46;
    border-radius: 5px;
    padding: 6px 14px;
    color: #d2d2d7;
    font-size: 11px;
    font-weight: bold;
}
QPushButton:hover { background-color: #353542; border-color: #00b4dc; }
QPushButton:pressed { background-color: #00b4dc; color: #111; }
QPushButton:disabled { color: #555; border-color: #2a2a34; }
QPushButton#accent {
    background-color: #00789c;
    border-color: #00b4dc;
    color: white;
}
QPushButton#accent:hover { background-color: #009abb; }
QPushButton#side_upper {
    background-color: #802020;
    border-color: #f05050;
    color: white;
}
QPushButton#side_lower {
    background-color: #205080;
    border-color: #50b4f0;
    color: white;
}
QRadioButton, QCheckBox { color: #ccc; font-size: 11px; }
QTableWidget {
    background-color: #1a1a20;
    alternate-background-color: #20202a;
    color: #d2d2d7;
    gridline-color: #2a2a34;
    border: 1px solid #2a2a34;
    border-radius: 4px;
    font-size: 11px;
}
QTableWidget::item:selected { background-color: #00789c; color: white; }
QHeaderView::section {
    background-color: #2a2a34;
    color: #aaa;
    padding: 4px;
    border: none;
    border-right: 1px solid #1a1a20;
    border-bottom: 1px solid #1a1a20;
    font-weight: bold;
}
QListWidget {
    background-color: #1a1a20;
    border: 1px solid #2a2a34;
    border-radius: 4px;
    color: #d2d2d7;
    font-size: 11px;
}
QListWidget::item:selected { background-color: #00789c; color: white; }
QSplitter::handle { background-color: #2a2a34; }
QScrollArea { border: none; background-color: transparent; }
QScrollBar:vertical {
    background: #1a1a20; width: 10px; margin: 0;
}
QScrollBar::handle:vertical {
    background: #3a3a46; border-radius: 5px; min-height: 20px;
}
QScrollBar::handle:vertical:hover { background: #4a4a56; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }

/* --- Tab widget (FR / THD / HOHD) --- */
QTabWidget::pane {
    border: 1px solid #2a2a34;
    border-radius: 6px;
    background-color: #1e1e26;
    top: -1px;
}
QTabBar {
    background-color: transparent;
    qproperty-drawBase: 0;
}
QTabBar::tab {
    background-color: #1a1a20;
    color: #888;
    border: 1px solid #2a2a34;
    border-bottom: none;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    padding: 8px 24px;
    margin-right: 4px;
    min-width: 80px;
    font-size: 12px;
    font-weight: bold;
}
QTabBar::tab:selected {
    background-color: #1e1e26;
    color: #00b4dc;
    border: 1px solid #00789c;
    border-bottom: 2px solid #1e1e26;
}
QTabBar::tab:hover:!selected {
    background-color: #232330;
    color: #d2d2d7;
}

/* --- Status bar --- */
QStatusBar {
    background-color: #15151a;
    color: #888;
    border-top: 1px solid #2a2a34;
    font-family: 'Consolas', monospace;
    font-size: 10px;
}
QStatusBar::item { border: none; }
"""

# ---------------------------------------------------------------------------
# REW File Parsing (text exports)
# ---------------------------------------------------------------------------

def parse_rew_file(filepath):
    """Parse a REW text export file."""
    name = os.path.splitext(os.path.basename(filepath))[0]
    headers = []
    rows = []
    column_labels = None

    with open(filepath, 'r', errors='replace') as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            if line.startswith('*'):
                headers.append(line)
                if 'Freq' in line or 'freq' in line.lower():
                    parts = line.lstrip('*').strip().split()
                    if len(parts) >= 2:
                        column_labels = parts
                continue
            parts = line.split()
            try:
                vals = [float(p) for p in parts]
                rows.append(vals)
            except ValueError:
                continue

    if not rows:
        return None

    arr = np.array(rows)
    freqs = arr[:, 0]
    n_cols = arr.shape[1]

    data = {}
    if column_labels and len(column_labels) >= n_cols:
        for i in range(1, n_cols):
            label = column_labels[i].split('(')[0].strip()
            data[label or 'col{}'.format(i)] = arr[:, i]
    else:
        if n_cols == 2:
            data['SPL'] = arr[:, 1]
        elif n_cols == 3:
            data['SPL'] = arr[:, 1]
            data['Phase'] = arr[:, 2]
        else:
            data['SPL'] = arr[:, 1]
            for i in range(2, n_cols):
                data['col{}'.format(i)] = arr[:, i]

    header_text = ' '.join(headers).lower()
    if 'thd' in header_text or 'distortion' in header_text:
        data_type = 'thd' if 'thd' in header_text else 'distortion'
    else:
        data_type = 'fr'

    return {
        'name': name,
        'freqs': freqs,
        'data': data,
        'data_type': data_type,
        'filepath': filepath,
    }


# ---------------------------------------------------------------------------
# REW REST API Client
# ---------------------------------------------------------------------------

def rew_api_get(host, path, timeout=10):
    url = host.rstrip('/') + path
    req = urllib.request.Request(url, headers={'Accept': 'application/json'})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode('utf-8'))


def rew_list_measurements(host):
    """Fetch list of measurements from REW."""
    data = rew_api_get(host, '/measurements')
    items = []
    if isinstance(data, dict):
        for mid, m in data.items():
            if mid == 'message':
                continue
            items.append({
                'id': mid,
                'title': m.get('title', 'Measurement {}'.format(mid)) if isinstance(m, dict) else str(m),
            })
    elif isinstance(data, list):
        for m in data:
            items.append({
                'id': str(m.get('id', '')),
                'title': m.get('title', 'Untitled'),
            })
    return items


def _decode_b64_floats(s):
    """Decode REW base64-encoded big-endian float32 array to numpy."""
    raw = base64.b64decode(s)
    n = len(raw) // 4
    return np.array(struct.unpack('>{}f'.format(n), raw), dtype=float)


def _build_rew_freq_axis(start_freq, ppo, count):
    """Reconstruct REW's log-spaced frequency axis."""
    return float(start_freq) * (2.0 ** (np.arange(count) / float(ppo)))


def _extract_fr(resp):
    """Extract (freqs, magnitude, phase) from a REW FR response."""
    if not isinstance(resp, dict):
        return None, None, None

    raw_mag = resp.get('magnitude') or resp.get('SPL') or resp.get('mag')
    if isinstance(raw_mag, str):
        mag = _decode_b64_floats(raw_mag)
    elif isinstance(raw_mag, list):
        mag = np.array(raw_mag, dtype=float)
    else:
        return None, None, None

    raw_phase = resp.get('phase')
    phase = None
    if isinstance(raw_phase, str):
        phase = _decode_b64_floats(raw_phase)
    elif isinstance(raw_phase, list):
        phase = np.array(raw_phase, dtype=float)

    freqs = None
    if 'freqs' in resp and isinstance(resp['freqs'], list):
        freqs = np.array(resp['freqs'], dtype=float)
    elif 'startFreq' in resp and 'ppo' in resp:
        freqs = _build_rew_freq_axis(resp['startFreq'], resp['ppo'], len(mag))
    elif 'startFreq' in resp and 'freqStep' in resp:
        start = float(resp['startFreq'])
        step = float(resp['freqStep'])
        freqs = start + step * np.arange(len(mag))

    if freqs is None or len(freqs) != len(mag):
        return None, None, None
    return freqs, mag, phase


def _extract_distortion(resp):
    """Extract distortion data from a REW /distortion response.

    REW V5.40+ returns:
        {
          "columnHeaders": ["Freq (Hz)", "Fundamental (dB)", "THD (%)",
                            "Noise (%)", "H2 (%)", "H3 (%)", ...],
          "data": [
              [100.0, 62.42, 0.85, 0.12, 0.4, 0.2, ...],   # one row per freq
              ...
          ]
        }

    Some older or alternative response shapes use top-level keys
    (`thd`, `H2`, etc.) with base64-encoded float arrays — we still
    handle those as a fallback.

    Returns:
        (freqs, data_dict) where data_dict has keys like 'THD%', 'H2',
        'H3', ... (always uppercased H-keys, THD always renamed 'THD%').
        Returns (None, {}) if extraction fails.

    NOTE: REW's headers indicate the unit (e.g. "THD (%)"); we always
    rename the THD column to "THD%" regardless of the unit so the rest
    of the GUI can find it. The URL caller is expected to request
    `unit=percent` so the actual values are in percent — see
    rew_fetch_measurement.
    """
    if not isinstance(resp, dict):
        return None, {}

    # ---- Path 1: columnHeaders + data rows (REW V5.40+ standard) ----
    headers = resp.get('columnHeaders')
    rows = resp.get('data')
    if isinstance(headers, list) and isinstance(rows, list) and rows:
        # Map each column to a short key:
        #   "Freq (Hz)"        -> "Freq"
        #   "Fundamental (dB)" -> "Fundamental"
        #   "THD (%)"          -> "THD"
        #   "H2 (%)"           -> "H2"
        short_keys = []
        for h in headers:
            if not isinstance(h, str):
                short_keys.append(None)
                continue
            key = h.split('(')[0].strip() if '(' in h else h.strip()
            short_keys.append(key)

        # Find the freq column index (usually 0)
        freq_idx = None
        for i, k in enumerate(short_keys):
            if k and k.lower().startswith('freq'):
                freq_idx = i
                break
        if freq_idx is None:
            freq_idx = 0  # assume column 0 is freq if not labeled

        # Allocate per-column arrays as NaN, then fill in valid rows
        n_pts = len(rows)
        n_cols = len(headers)
        per_col = {}
        freq_list = np.full(n_pts, np.nan, dtype=float)

        for c, k in enumerate(short_keys):
            if c == freq_idx or not k:
                continue
            per_col[k] = np.full(n_pts, np.nan, dtype=float)

        for row_i, row in enumerate(rows):
            if not isinstance(row, (list, tuple)):
                continue
            if freq_idx < len(row) and row[freq_idx] is not None:
                try:
                    freq_list[row_i] = float(row[freq_idx])
                except (TypeError, ValueError):
                    pass
            for c, k in enumerate(short_keys):
                if c == freq_idx or not k or c >= len(row):
                    continue
                if row[c] is None:
                    continue
                try:
                    per_col[k][row_i] = float(row[c])
                except (TypeError, ValueError):
                    pass

        # Drop rows where freq is NaN
        valid = ~np.isnan(freq_list)
        if not np.any(valid):
            return None, {}
        freqs = freq_list[valid]

        # Rename THD -> THD% for GUI consistency, uppercase H-keys
        out = {}
        for k, arr in per_col.items():
            arr = arr[valid]
            if k.upper() == 'THD':
                out['THD%'] = arr
            elif k.upper().startswith('H') and k[1:].isdigit():
                out[k.upper()] = arr
            else:
                out[k] = arr

        if out:
            return freqs, out
        return None, {}

    # ---- Path 2: legacy top-level keys with base64-encoded arrays ----
    out = {}
    sample_len = None

    field_map = [
        ('thdPercent', 'THD%'),
        ('thd', 'THD%'),
        ('thdN', 'THD+N'),
        ('thdDb', 'THDdB'),
        ('fundamental', 'Fundamental'),
        ('noise', 'Noise'),
    ]

    for src_key, dst_key in field_map:
        if src_key in resp:
            val = resp[src_key]
            if isinstance(val, str):
                arr = _decode_b64_floats(val)
            elif isinstance(val, list):
                arr = np.array(val, dtype=float)
            else:
                continue
            out[dst_key] = arr
            sample_len = len(arr)

    for k, v in resp.items():
        kl = k.lower()
        if kl.startswith('h') and kl[1:].isdigit():
            if isinstance(v, str):
                arr = _decode_b64_floats(v)
            elif isinstance(v, list):
                arr = np.array(v, dtype=float)
            else:
                continue
            out[k.upper()] = arr
            sample_len = len(arr)

    if not out or sample_len is None:
        return None, {}

    freqs = None
    if 'freqs' in resp and isinstance(resp['freqs'], list):
        freqs = np.array(resp['freqs'], dtype=float)
    elif 'startFreq' in resp and 'ppo' in resp:
        freqs = _build_rew_freq_axis(resp['startFreq'], resp['ppo'], sample_len)

    if freqs is None or len(freqs) != sample_len:
        return None, {}
    return freqs, out


def rew_fetch_measurement(host, mid, ppo=48):
    """Fetch FR + distortion for a measurement, trying multiple endpoints."""
    try:
        meta = rew_api_get(host, '/measurements/{}'.format(mid))
        title = meta.get('title', 'Measurement {}'.format(mid))
    except Exception:
        title = 'Measurement {}'.format(mid)

    data = {}
    freqs = None
    data_type = 'fr'
    errors = []

    fr_endpoints = [
        '/measurements/{}/frequency-response?ppo={}'.format(mid, ppo),
        '/measurements/{}/frequency-response'.format(mid),
        '/measurements/{}/spl?ppo={}'.format(mid, ppo),
        '/measurements/{}/spl'.format(mid),
    ]
    for ep in fr_endpoints:
        try:
            fr = rew_api_get(host, ep)
            f, m_arr, p_arr = _extract_fr(fr)
            if f is not None and m_arr is not None:
                freqs = f
                data['SPL'] = m_arr
                if p_arr is not None:
                    data['Phase'] = p_arr
                break
        except Exception as e:
            errors.append("{}: {}".format(ep, e))

    # Distortion endpoints. We always request unit=percent so harmonic values
    # are directly usable for THD/HOHD limit math (rew-iqc reads percent).
    # Distortion uses a coarser default PPO (12) than FR — REW's distortion
    # analysis is much sparser than its FR; asking for ppo=48 wastes effort
    # and won't return more rows.
    dist_ppo = 12
    dist_freqs = None
    dist_endpoints = [
        '/measurements/{}/distortion?unit=percent&ppo={}'.format(mid, dist_ppo),
        '/measurements/{}/distortion?ppo={}'.format(mid, dist_ppo),
        '/measurements/{}/distortion'.format(mid),
    ]
    for ep in dist_endpoints:
        try:
            dist = rew_api_get(host, ep)
            f, dist_data = _extract_distortion(dist)
            if f is not None:
                # Distortion data has its OWN freq axis (different length and
                # range than FR) — keep it separate so the plotting code can
                # pair distortion columns with the correct x-values.
                dist_freqs = f
                if freqs is None:
                    freqs = f
                for key, vals in dist_data.items():
                    data[key] = vals
                if 'THD%' in dist_data or 'THDdB' in dist_data:
                    data_type = 'thd'
                if any(k.startswith('H') and k[1:].isdigit() for k in dist_data):
                    data_type = 'distortion'
                break
        except Exception as e:
            errors.append("{}: {}".format(ep, e))

    if freqs is None or not data:
        raise RuntimeError(
            "Could not fetch data for measurement {}. Tried:\n{}".format(
                mid, '\n'.join(errors[:6])))

    out = {
        'name': title,
        'freqs': freqs,
        'data': data,
        'data_type': data_type,
        'filepath': 'rew://measurement/{}'.format(mid),
    }
    if dist_freqs is not None:
        out['dist_freqs'] = dist_freqs
    return out


class RewCaptureDialog(QDialog):
    def __init__(self, parent=None, host=DEFAULT_REW_HOST):
        super().__init__(parent)
        self.setWindowTitle('Capture from REW')
        self.setMinimumSize(540, 480)
        self.setStyleSheet(STYLESHEET)
        self.host = host
        self.selected = []

        lay = QVBoxLayout(self)

        hl = QHBoxLayout()
        hl.addWidget(QLabel("REW API URL:"))
        self.edit_host = QLineEdit(host)
        hl.addWidget(self.edit_host)
        btn_refresh = QPushButton("Refresh")
        btn_refresh.clicked.connect(self._refresh)
        hl.addWidget(btn_refresh)
        lay.addLayout(hl)

        self.lbl_status = QLabel("Click Refresh to fetch measurement list.")
        self.lbl_status.setStyleSheet("color: #888;")
        lay.addWidget(self.lbl_status)

        self.list_meas = QListWidget()
        self.list_meas.setSelectionMode(QAbstractItemView.ExtendedSelection)
        lay.addWidget(self.list_meas, stretch=1)

        ppo_row = QHBoxLayout()
        ppo_row.addWidget(QLabel("Points/Octave:"))
        self.spin_ppo = QSpinBox()
        self.spin_ppo.setRange(1, 96)
        self.spin_ppo.setValue(48)
        ppo_row.addWidget(self.spin_ppo)
        ppo_row.addStretch()
        lay.addLayout(ppo_row)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.button(QDialogButtonBox.Ok).setText("Import Selected")
        btns.accepted.connect(self._import_selected)
        btns.rejected.connect(self.reject)
        lay.addWidget(btns)

        self._refresh()

    def _refresh(self):
        self.host = self.edit_host.text().strip()
        self.list_meas.clear()
        self.lbl_status.setText("Fetching from {}...".format(self.host))
        self.lbl_status.setStyleSheet("color: #888;")
        QApplication.processEvents()
        try:
            items = rew_list_measurements(self.host)
            if not items:
                self.lbl_status.setText("Connected, but no measurements in REW.")
                return
            for it in items:
                qi = QListWidgetItem("[{}] {}".format(it['id'], it['title']))
                qi.setData(Qt.UserRole, it['id'])
                self.list_meas.addItem(qi)
            self.lbl_status.setText(
                "Found {} measurement(s). Select and click Import.".format(len(items)))
        except urllib.error.URLError:
            self.lbl_status.setText(
                "Cannot reach REW at {}. Is REW running with API enabled?".format(self.host))
            self.lbl_status.setStyleSheet("color: #f66;")
        except Exception as e:
            self.lbl_status.setText("Error: {}".format(e))
            self.lbl_status.setStyleSheet("color: #f66;")

    def _import_selected(self):
        items = self.list_meas.selectedItems()
        if not items:
            QMessageBox.warning(self, 'Capture', 'Select at least one measurement.')
            return
        ppo = self.spin_ppo.value()
        self.selected = []
        self.lbl_status.setText("Fetching {} measurement(s)...".format(len(items)))
        self.lbl_status.setStyleSheet("color: #888;")
        QApplication.processEvents()

        for qi in items:
            mid = qi.data(Qt.UserRole)
            try:
                m = rew_fetch_measurement(self.host, mid, ppo)
                if m is not None:
                    self.selected.append(m)
            except Exception as e:
                QMessageBox.warning(self, 'Fetch Error',
                                     "Failed to fetch measurement {}:\n{}".format(mid, e))
        if self.selected:
            self.accept()
        else:
            self.lbl_status.setText("No measurements could be fetched.")
            self.lbl_status.setStyleSheet("color: #f66;")


# ---------------------------------------------------------------------------
# Frequency / coordinate mapping helpers
# ---------------------------------------------------------------------------

def freq_to_x(f, plot_w, padding_l):
    if f <= 0:
        return padding_l
    norm = (np.log10(f) - np.log10(FREQ_MIN)) / (np.log10(FREQ_MAX) - np.log10(FREQ_MIN))
    return padding_l + norm * plot_w


def x_to_freq(x, plot_w, padding_l):
    norm = (x - padding_l) / max(1, plot_w)
    norm = max(0, min(1, norm))
    log_f = np.log10(FREQ_MIN) + norm * (np.log10(FREQ_MAX) - np.log10(FREQ_MIN))
    return 10 ** log_f


def db_to_y(db, plot_h, padding_t, db_min, db_max):
    rng = db_max - db_min
    if rng == 0:
        return padding_t + plot_h / 2
    norm = (db - db_min) / rng
    return padding_t + plot_h * (1 - norm)


def y_to_db(y, plot_h, padding_t, db_min, db_max):
    norm = 1 - (y - padding_t) / max(1, plot_h)
    norm = max(0, min(1, norm))
    return db_min + norm * (db_max - db_min)


# ---------------------------------------------------------------------------
# Limit Computation
# ---------------------------------------------------------------------------

def smooth_fractional_octave(freqs, values, fraction):
    """Apply fractional-octave smoothing in the energy domain (REW-style).

    Implementation notes:
    - Uses energy-domain averaging: dB -> linear power -> mean -> dB.
      This is the conventional acoustic averaging method and produces
      smoother results than dB-domain (arithmetic) averaging.
    - Internally resamples onto a uniform log-frequency grid before
      smoothing. Without this, irregular input grids produce variable
      window densities at different frequencies (more samples per octave
      at high freqs in linearly-spaced data) which causes the visible
      "step" artifacts in the high-frequency region.
    - Uses a Hann (raised cosine) weighting window instead of a flat
      rectangular average, which further reduces aliasing artifacts at
      the smoothing window edges.

    Args:
        freqs: array of frequencies (Hz), need not be uniform
        values: array of values in dB
        fraction: smoothing fraction (e.g., 3 for 1/3 octave). 0 disables.

    Returns:
        Smoothed values array on the original frequency grid.
    """
    if not fraction or fraction <= 0:
        return values.copy()

    # Build a dense uniform log-frequency grid covering the input range.
    # ~100 samples per octave gives plenty of resolution while keeping
    # the smoothing window large enough to be meaningful.
    f_min = max(1e-3, np.min(freqs[freqs > 0])) if np.any(freqs > 0) else 1.0
    f_max = np.max(freqs)
    if f_max <= f_min:
        return values.copy()
    n_octaves = np.log2(f_max / f_min)
    n_uniform = max(256, int(n_octaves * 100))
    uniform_log_f = np.linspace(np.log2(f_min), np.log2(f_max), n_uniform)
    uniform_f = 2.0 ** uniform_log_f

    # Interpolate input values onto uniform log grid (in dB).
    sort_idx = np.argsort(freqs)
    f_sorted = freqs[sort_idx]
    v_sorted = values[sort_idx]
    valid = np.isfinite(v_sorted)
    uniform_db = np.interp(uniform_f, f_sorted[valid], v_sorted[valid])

    # Convert to linear power for energy-domain averaging.
    uniform_power = 10.0 ** (uniform_db / 10.0)

    # Smoothing window in log-frequency samples.
    # Half-width = (1/(2*fraction)) octaves * (samples per octave)
    samples_per_octave = (n_uniform - 1) / n_octaves
    half_w = max(1, int(round(samples_per_octave * 0.5 / fraction)))
    win_len = 2 * half_w + 1

    # Hann window for smooth, anti-aliased averaging.
    window = np.hanning(win_len)
    window /= window.sum()

    # Convolve in linear power domain. Use 'edge' padding so edges don't
    # droop toward zero.
    padded = np.pad(uniform_power, half_w, mode='edge')
    smoothed_power = np.convolve(padded, window, mode='valid')

    # Back to dB.
    smoothed_power = np.maximum(smoothed_power, 1e-30)  # avoid log(0)
    smoothed_db = 10.0 * np.log10(smoothed_power)

    # Interpolate back to the original frequency grid.
    out = np.interp(freqs, uniform_f, smoothed_db)
    # Preserve any non-finite values from input
    out = np.where(np.isfinite(values), out, values)
    return out


def normalize_curve_to_freq(freqs, values, ref_freq):
    """Shift a curve so its value at ref_freq becomes 0 dB."""
    idx = np.argmin(np.abs(freqs - ref_freq))
    return values - values[idx]


def normalize_measurements(measurements, mode, ref_freq, primary_col,
                            smoothing=0, reference_idx=None):
    """Apply smoothing + normalization to measurements.

    Returns (ref_freqs, normalized_stack, basis_curve) where:
      - ref_freqs: shared frequency grid
      - normalized_stack: shape (n_measurements, n_freqs) of all curves
      - basis_curve: the curve to compute limits from. Either the selected
        reference measurement, or the mean across all if reference_idx is None.
    """
    if not measurements:
        return None, None, None
    ref_freqs = measurements[0]['freqs']

    stacks = []
    for m in measurements:
        col = primary_col if primary_col in m['data'] else list(m['data'].keys())[0]
        vals = np.interp(ref_freqs, m['freqs'], m['data'][col])
        if smoothing and smoothing > 0:
            vals = smooth_fractional_octave(ref_freqs, vals, smoothing)
        if mode == 'normalize' and ref_freq is not None:
            vals = normalize_curve_to_freq(ref_freqs, vals, ref_freq)
        stacks.append(vals)

    stack = np.array(stacks)

    if reference_idx is not None and 0 <= reference_idx < len(stacks):
        basis = stacks[reference_idx]
    else:
        basis = np.mean(stack, axis=0) if len(stacks) > 1 else stacks[0]

    return ref_freqs, stack, basis


def compute_sigma_limits(stack, basis, sigma_mult, shape='both'):
    """Compute basis +/- N*sigma.

    Args:
        stack: shape (n_measurements, n_freqs) - used for std calculation
        basis: 1-d array - the curve to build limits around (mean or reference)
        sigma_mult: multiplier for sigma
        shape: 'both', 'upper', 'lower'

    Returns:
        (basis, upper, lower)
    """
    if stack is None or stack.shape[0] < 2:
        return basis, None, None
    std = np.std(stack, axis=0, ddof=1)
    upper = basis + sigma_mult * std if shape in ('both', 'upper') else None
    lower = basis - sigma_mult * std if shape in ('both', 'lower') else None
    return basis, upper, lower


def compute_offset_limits(basis, offset_type, offset_up, offset_down, shape='both'):
    """Compute basis +/- offset (no statistics).

    Args:
        basis: 1-d array - the curve to build limits around
        offset_type: 'dB' or '%'
        offset_up: positive headroom above basis
        offset_down: positive amount below basis
        shape: 'both', 'upper', 'lower'

    Returns:
        (basis, upper, lower)
    """
    if basis is None:
        return None, None, None

    if offset_type == 'dB':
        up_db = offset_up
        down_db = offset_down
    else:
        up_db = 20 * math.log10(1 + offset_up / 100.0) if offset_up > -100 else 0
        down_db = 20 * math.log10(1 + offset_down / 100.0) if offset_down > -100 else 0

    upper = basis + up_db if shape in ('both', 'upper') else None
    lower = basis - down_db if shape in ('both', 'lower') else None
    return basis, upper, lower


def anchors_to_curves(anchors, freqs, shape='both'):
    """Convert anchors to continuous upper/lower curves."""
    upper_anchors = sorted([a for a in anchors if a['side'] == 'upper'],
                            key=lambda a: a['freq'])
    lower_anchors = sorted([a for a in anchors if a['side'] == 'lower'],
                            key=lambda a: a['freq'])

    def interp_anchors(anchor_list):
        if len(anchor_list) < 2:
            return np.full_like(freqs, np.nan, dtype=float)
        af = np.array([a['freq'] for a in anchor_list])
        av = np.array([a['value'] for a in anchor_list])
        log_freqs = np.log10(np.clip(freqs, 1, None))
        log_af = np.log10(np.clip(af, 1, None))
        return np.interp(log_freqs, log_af, av, left=np.nan, right=np.nan)

    upper = interp_anchors(upper_anchors) if shape in ('both', 'upper') else None
    lower = interp_anchors(lower_anchors) if shape in ('both', 'lower') else None
    return upper, lower


def evaluate_dut_against_limits(dut_freqs, dut_values, limit_freqs,
                                  upper, lower, range_start, range_stop):
    """Check a DUT measurement against limit curves.

    Returns a result dict with pass/fail info and details on any failures.
    """
    in_range = (limit_freqs >= range_start) & (limit_freqs <= range_stop)
    eval_freqs = limit_freqs[in_range]
    if len(eval_freqs) == 0:
        return {'pass': False, 'n_evaluated': 0, 'n_fail': 0,
                'fail_freqs': np.array([]), 'fail_reasons': [],
                'max_overshoot_db': 0.0, 'max_undershoot_db': 0.0,
                'error': 'No frequency range to evaluate.'}

    dut_on_grid = np.interp(eval_freqs, dut_freqs, dut_values)
    upper_eval = upper[in_range] if upper is not None else None
    lower_eval = lower[in_range] if lower is not None else None

    fail_mask = np.zeros(len(eval_freqs), dtype=bool)
    over_db = np.zeros(len(eval_freqs))
    under_db = np.zeros(len(eval_freqs))

    if upper_eval is not None:
        valid = np.isfinite(upper_eval) & np.isfinite(dut_on_grid)
        over_db = np.where(valid, np.maximum(0, dut_on_grid - upper_eval), 0)
        fail_mask |= (over_db > 0)

    if lower_eval is not None:
        valid = np.isfinite(lower_eval) & np.isfinite(dut_on_grid)
        under_db = np.where(valid, np.maximum(0, lower_eval - dut_on_grid), 0)
        fail_mask |= (under_db > 0)

    fail_idx = np.where(fail_mask)[0]
    fail_freqs = eval_freqs[fail_mask]
    reasons = []
    for i in fail_idx:
        if over_db[i] > 0:
            reasons.append("{:.0f} Hz: {:+.2f} dB over upper".format(
                eval_freqs[i], over_db[i]))
        elif under_db[i] > 0:
            reasons.append("{:.0f} Hz: {:+.2f} dB under lower".format(
                eval_freqs[i], under_db[i]))

    return {
        'pass': len(fail_idx) == 0,
        'n_evaluated': len(eval_freqs),
        'n_fail': len(fail_idx),
        'fail_freqs': fail_freqs,
        'fail_reasons': reasons,
        'max_overshoot_db': float(np.max(over_db)) if len(over_db) > 0 else 0.0,
        'max_undershoot_db': float(np.max(under_db)) if len(under_db) > 0 else 0.0,
    }


# ---------------------------------------------------------------------------
# Limit Plot Widget
# ---------------------------------------------------------------------------

class LimitPlot(QWidget):
    """Interactive frequency-domain plot with editable limit lines.

    Kind-aware: the plot adapts its y-axis range, default values, and
    column-extraction logic based on the workspace's kind (FR/THD/HOHD).
    For HOHD, it aggregates the user-selected harmonics from each
    measurement using sqrt-sum-of-squares.
    """

    anchors_changed = pyqtSignal()

    def __init__(self, parent=None, kind=KIND_FR):
        super().__init__(parent)
        self.setMinimumHeight(380)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setMouseTracking(True)

        self.kind = kind
        self.kind_config = KIND_CONFIG[kind]
        self.harmonics = list(self.kind_config['default_harmonics'])  # editable list

        self.measurements = []
        self.measurement_curves = []  # parallel: post-normalization curves to draw
        self.anchors = []
        self.anchor_side = 'upper'  # which side new anchors go to

        self.limit_freqs = None
        self.upper_limit = None
        self.lower_limit = None
        self.mean_curve = None

        # DUT test mode: an overlay measurement to evaluate against limits
        self.dut_freqs = None
        self.dut_values = None
        self.dut_pass = None  # None = no test run, True/False = result
        self.dut_fail_freqs = None  # array of failing frequencies for highlighting

        self.range_start = 20.0
        self.range_stop = 20000.0

        # Y-axis defaults from kind config
        self.db_min = self.kind_config['y_default_min']
        self.db_max = self.db_min + self.kind_config['y_span']

        self._dragging_anchor = None
        self._padding = (52, 20, 30, 36)  # left, right, top, bottom (extra bottom for legend)

    def set_harmonics(self, harmonics):
        """Set which harmonics this plot aggregates (HOHD/THD overrides).

        Has no effect for FR or when REW-aggregated THD is used. For HOHD,
        triggers a redraw because the displayed curves change.
        """
        self.harmonics = list(harmonics)
        if self.kind == KIND_HOHD:
            self._auto_scale_y()
            self.update()

    def set_measurements(self, measurements, normalized_curves=None):
        self.measurements = measurements
        self.measurement_curves = normalized_curves
        self._auto_scale_y()
        self.update()

    def set_anchor_side(self, side):
        self.anchor_side = side
        self.update()

    def set_range(self, start_hz, stop_hz):
        self.range_start = start_hz
        self.range_stop = stop_hz
        self.update()

    def set_limit_curves(self, freqs, upper, lower, mean=None):
        self.limit_freqs = freqs
        self.upper_limit = upper
        self.lower_limit = lower
        self.mean_curve = mean
        self.update()

    def clear_limits(self):
        self.limit_freqs = None
        self.upper_limit = None
        self.lower_limit = None
        self.mean_curve = None
        self.update()

    def set_dut(self, freqs, values, passed=None, fail_freqs=None):
        """Set the DUT (Device Under Test) curve to display as an overlay.

        Args:
            freqs: DUT frequencies (or None to clear)
            values: DUT values in dB
            passed: True/False/None - test result
            fail_freqs: array of frequencies that failed (for highlight markers)
        """
        self.dut_freqs = freqs
        self.dut_values = values
        self.dut_pass = passed
        self.dut_fail_freqs = fail_freqs
        self.update()

    def clear_dut(self):
        self.dut_freqs = None
        self.dut_values = None
        self.dut_pass = None
        self.dut_fail_freqs = None
        self.update()

    def _auto_scale_y(self):
        """Scale Y to data, then snap to the kind's standard span.

        Each kind has a fixed span in y_span (e.g. 50 dB for FR, 30% for THD)
        so the visual scale stays consistent across measurements within a tab.
        """
        all_vals = []
        if self.measurement_curves is not None:
            for entry in self.measurement_curves:
                # Tuples of (freqs, values); also tolerate plain arrays for
                # any callers that haven't been updated.
                if isinstance(entry, tuple):
                    _xs, ys = entry
                    if ys is not None:
                        all_vals.extend(ys.tolist())
                elif entry is not None:
                    all_vals.extend(entry.tolist())
        else:
            for m in self.measurements:
                primary = self._primary_column(m)
                if primary is not None:
                    all_vals.extend(primary.tolist())

        span = self.kind_config['y_span']
        floor = self.kind_config['y_floor']

        if all_vals:
            arr = np.array(all_vals)
            arr = arr[np.isfinite(arr)]
            if len(arr) > 0:
                if floor is not None:
                    # Bottom-anchored kinds (THD, HOHD): start at floor, span up
                    # to cover the data peak with some headroom.
                    self.db_min = floor
                    peak = float(np.max(arr))
                    # Round up to next 5 (or 1) above peak * 1.2
                    target_top = max(peak * 1.2, span)
                    step = 5 if span >= 20 else 1
                    self.db_max = math.ceil(target_top / step) * step
                else:
                    # Centered kinds (FR): center the span on the data median.
                    center = np.median(arr)
                    self.db_min = math.floor((center - span / 2) / 10) * 10
                    self.db_max = self.db_min + span
                return

        # Default range if no data
        if floor is not None:
            self.db_min = floor
            self.db_max = floor + span
        else:
            self.db_min = self.kind_config['y_default_min']
            self.db_max = self.db_min + span

    def _primary_column(self, m):
        """Pick the right data column from a measurement based on kind.

        For HOHD, aggregates the selected harmonics on the fly via
        sqrt-sum-of-squares; returns None if none of the requested
        harmonics are present in the measurement.

        NOTE: returns the raw values only. Use _primary_xy() if you need
        the matched frequency axis (FR data and distortion data have
        different axes).
        """
        data = m['data']
        if self.kind == KIND_FR:
            for col in ('SPL', 'mag', 'Magnitude'):
                if col in data:
                    return data[col]
            return list(data.values())[0]
        elif self.kind == KIND_THD:
            # Prefer REW-aggregated THD%; fall back to aggregating selected
            # harmonics if THD% column isn't present.
            for col in ('THD%', 'THD'):
                if col in data:
                    return data[col]
            return self._aggregate_harmonics(m)
        elif self.kind == KIND_HOHD:
            return self._aggregate_harmonics(m)
        return None

    def _primary_xy(self, m):
        """Return (freqs, values) for the active kind, paired correctly.

        FR columns use m['freqs']; distortion columns use m['dist_freqs']
        if it exists (it's set when REW returned distortion data on a
        different frequency axis than the FR data). Returns (None, None)
        if no usable column for this kind exists in the measurement.
        """
        values = self._primary_column(m)
        if values is None:
            return None, None
        # FR uses the main freq axis; THD/HOHD use the distortion axis when
        # present (REW returns distortion on a sparser axis than FR).
        if self.kind == KIND_FR:
            freqs = m.get('freqs')
        else:
            freqs = m.get('dist_freqs')
            if freqs is None:
                freqs = m.get('freqs')
        if freqs is None or len(freqs) != len(values):
            return None, None
        return freqs, values

    def _aggregate_harmonics(self, m):
        """Aggregate the configured harmonics in measurement m via sqrt-sum-of-squares.

        Returns None if none of the requested harmonics are present.
        Uppercase the harmonic keys when looking up so 'H10' matches 'H10'.
        """
        data = m['data']
        present = []
        ref_len = None
        for h in self.harmonics:
            key = h.upper()
            if key in data:
                arr = np.nan_to_num(np.asarray(data[key], dtype=float), nan=0.0)
                if ref_len is None:
                    ref_len = len(arr)
                if len(arr) != ref_len:
                    continue   # skip mismatched-length harmonics defensively
                present.append(arr)
        if not present:
            return None
        # sqrt(sum of squares) — assumes harmonic values are in % (REW
        # exports/distortion API in percent mode), matching rew_iqc behavior.
        stack = np.vstack(present)
        return np.sqrt(np.sum(stack ** 2, axis=0))

    def _plot_rect(self):
        pl, pr, pt, pb = self._padding
        return pl, pt, self.width() - pl - pr, self.height() - pt - pb

    def _f2x(self, f):
        pl, pt, pw, ph = self._plot_rect()
        return freq_to_x(f, pw, pl)

    def _x2f(self, x):
        pl, pt, pw, ph = self._plot_rect()
        return x_to_freq(x, pw, pl)

    def _db2y(self, db):
        pl, pt, pw, ph = self._plot_rect()
        return db_to_y(db, ph, pt, self.db_min, self.db_max)

    def _y2db(self, y):
        pl, pt, pw, ph = self._plot_rect()
        return y_to_db(y, ph, pt, self.db_min, self.db_max)

    # -- Painting --

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        w, h = self.width(), self.height()
        pl, pt, pw, ph = self._plot_rect()

        # Background
        grad = QLinearGradient(0, 0, 0, h)
        grad.setColorAt(0, QColor(22, 22, 28))
        grad.setColorAt(1, QColor(18, 18, 22))
        p.fillRect(0, 0, w, h, grad)

        # Range highlight
        x0 = self._f2x(self.range_start)
        x1 = self._f2x(self.range_stop)
        p.fillRect(QRectF(x0, pt, x1 - x0, ph), RANGE_FILL)

        # Vertical grid (frequency) - standard semilog with 1-9 per decade
        # Major lines at decades (10, 100, 1k, 10k); minor at 2,3,4,5,6,7,8,9
        major_freqs = [10, 100, 1000, 10000, 100000]
        for decade_start in [10, 100, 1000, 10000]:
            for mult in range(1, 10):
                f = decade_start * mult
                if not (FREQ_MIN <= f <= FREQ_MAX):
                    continue
                x = self._f2x(f)
                is_major = (f in major_freqs)
                if is_major:
                    p.setPen(QPen(GRID_MAJOR, 0.8))
                else:
                    p.setPen(QPen(GRID_LINE, 0.3))
                p.drawLine(int(x), pt, int(x), pt + ph)

        # Frequency labels
        p.setPen(TEXT_DIM)
        p.setFont(QFont('Segoe UI', 8))
        for f, lbl in [(20, '20'), (50, '50'), (100, '100'), (200, '200'),
                        (500, '500'), (1000, '1k'), (2000, '2k'), (5000, '5k'),
                        (10000, '10k'), (20000, '20k')]:
            if FREQ_MIN <= f <= FREQ_MAX:
                x = self._f2x(f)
                p.drawText(int(x) - 12, pt + ph + 14, lbl)

        # Horizontal grid (dB)
        db_step = 10 if (self.db_max - self.db_min) <= 80 else 20
        db = math.ceil(self.db_min / db_step) * db_step
        while db <= self.db_max:
            y = self._db2y(db)
            p.setPen(QPen(GRID_MAJOR if db == 0 else GRID_LINE,
                           1.0 if db == 0 else 0.5))
            p.drawLine(pl, int(y), pl + pw, int(y))
            p.setPen(TEXT_DIM)
            p.setFont(QFont('Segoe UI', 8))
            p.drawText(4, int(y) + 4, "{}".format(int(db)))
            db += db_step

        # Plot measurements
        for i, m in enumerate(self.measurements):
            color = CURVE_COLORS[i % len(CURVE_COLORS)]
            if self.measurement_curves is not None and i < len(self.measurement_curves):
                # measurement_curves is a list of (freqs, values) tuples so
                # FR and distortion data can have different x-axes
                fr_x, fr_y = self.measurement_curves[i]
                if fr_x is not None and fr_y is not None:
                    self._draw_array(p, fr_x, fr_y, color, width=1.2)
            else:
                xs, ys = self._primary_xy(m)
                if xs is not None and ys is not None:
                    self._draw_array(p, xs, ys, color, width=1.2)

        # Mean curve
        if self.mean_curve is not None and self.limit_freqs is not None:
            self._draw_array(p, self.limit_freqs, self.mean_curve,
                              MEAN_LINE, width=1.5, dashed=True)

        # Limit lines
        if self.upper_limit is not None and self.limit_freqs is not None:
            self._draw_array(p, self.limit_freqs, self.upper_limit,
                              UPPER_LIMIT, width=2.2)
        if self.lower_limit is not None and self.limit_freqs is not None:
            self._draw_array(p, self.limit_freqs, self.lower_limit,
                              LOWER_LIMIT, width=2.2)

        # DUT overlay (if test is loaded)
        if self.dut_freqs is not None and self.dut_values is not None:
            dut_color = QColor(80, 240, 80) if self.dut_pass else QColor(240, 80, 80)
            if self.dut_pass is None:
                dut_color = QColor(240, 240, 240)  # white when untested
            self._draw_array(p, self.dut_freqs, self.dut_values, dut_color, width=2.6)

            # Mark failing frequencies with red dots
            if self.dut_fail_freqs is not None and len(self.dut_fail_freqs) > 0:
                p.setPen(Qt.NoPen)
                p.setBrush(QColor(255, 60, 60))
                for ff in self.dut_fail_freqs:
                    fy_idx = np.argmin(np.abs(self.dut_freqs - ff))
                    fy = self.dut_values[fy_idx]
                    if np.isfinite(fy):
                        p.drawEllipse(QPointF(self._f2x(ff), self._db2y(fy)), 4, 4)

            # PASS/FAIL banner
            banner_text = "PASS" if self.dut_pass else "FAIL" if self.dut_pass is False else ""
            if banner_text:
                p.setFont(QFont('Segoe UI', 14, QFont.Bold))
                p.setPen(dut_color)
                p.drawText(int(pl + pw - 80), int(pt + 24), banner_text)

        # Anchors
        for a in self.anchors:
            x = self._f2x(a['freq'])
            y = self._db2y(a['value'])
            color = UPPER_LIMIT if a['side'] == 'upper' else LOWER_LIMIT
            p.setPen(QPen(color, 1.5))
            p.setBrush(ANCHOR_COLOR)
            p.drawEllipse(QPointF(x, y), 5, 5)

        # Range labels
        p.setPen(ACCENT)
        p.setFont(QFont('Segoe UI', 8, QFont.Bold))
        p.drawText(int(x0) + 2, pt + 12, "{:g} Hz".format(self.range_start))
        p.drawText(int(x1) - 60, pt + 12, "{:g} Hz".format(self.range_stop))

        # Anchor side indicator
        side_color = UPPER_LIMIT if self.anchor_side == 'upper' else LOWER_LIMIT
        p.setPen(side_color)
        p.setFont(QFont('Segoe UI', 9, QFont.Bold))
        p.drawText(pl + 6, pt + 14,
                    "Adding to: {}".format(self.anchor_side.upper()))

        p.end()

    def _draw_array(self, painter, freqs, values, color, width=1.5, dashed=False):
        pen = QPen(color, width)
        if dashed:
            pen.setStyle(Qt.DashLine)
        painter.setPen(pen)
        path = QPainterPath()
        first = True
        for f, v in zip(freqs, values):
            if not (FREQ_MIN <= f <= FREQ_MAX) or not np.isfinite(v):
                first = True
                continue
            x = self._f2x(f)
            y = self._db2y(v)
            if first:
                path.moveTo(x, y)
                first = False
            else:
                path.lineTo(x, y)
        painter.drawPath(path)

    # -- Mouse interaction --

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            anchor = self._anchor_at(e.x(), e.y())
            if anchor is not None:
                self._dragging_anchor = anchor
            else:
                f = self._x2f(e.x())
                v = self._y2db(e.y())
                self.anchors.append({'freq': f, 'value': v, 'side': self.anchor_side})
                self.anchors.sort(key=lambda a: (a['side'], a['freq']))
                self.anchors_changed.emit()
                self.update()
        elif e.button() == Qt.RightButton:
            anchor = self._anchor_at(e.x(), e.y())
            if anchor is not None:
                self.anchors.remove(anchor)
                self.anchors_changed.emit()
                self.update()

    def mouseMoveEvent(self, e):
        if self._dragging_anchor:
            self._dragging_anchor['freq'] = self._x2f(e.x())
            self._dragging_anchor['value'] = self._y2db(e.y())
            self.anchors.sort(key=lambda a: (a['side'], a['freq']))
            self.anchors_changed.emit()
            self.update()

    def mouseReleaseEvent(self, e):
        self._dragging_anchor = None

    def _anchor_at(self, x, y):
        for a in self.anchors:
            ax = self._f2x(a['freq'])
            ay = self._db2y(a['value'])
            if (ax - x) ** 2 + (ay - y) ** 2 <= ANCHOR_HIT_RADIUS ** 2:
                return a
        return None


# ---------------------------------------------------------------------------
# Legend Widget (color swatch + filename for each loaded measurement)
# ---------------------------------------------------------------------------

class LegendWidget(QWidget):
    """Color-swatch legend mapping each loaded measurement to its plot color.

    Also shows the standard limit-line colors (upper, lower, mean/basis)
    so the user can map every line in the plot back to its meaning.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.entries = []  # list of (name, color, style) where style is 'fill' or 'dash'
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setFixedHeight(34)

    def set_entries(self, measurements):
        # Standard limit-line entries always shown first
        self.entries = [
            ('Upper Limit', UPPER_LIMIT, 'line'),
            ('Lower Limit', LOWER_LIMIT, 'line'),
            ('Mean / Basis', MEAN_LINE, 'dash'),
        ]
        # Then per-measurement entries
        for i, m in enumerate(measurements):
            color = CURVE_COLORS[i % len(CURVE_COLORS)]
            self.entries.append((m['name'], color, 'fill'))
        self.update()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        # Background with subtle border
        p.fillRect(0, 0, self.width(), self.height(), QColor(26, 26, 32))
        p.setPen(QPen(QColor(60, 60, 72), 1))
        p.drawRect(0, 0, self.width() - 1, self.height() - 1)

        # "LEGEND:" label on the left
        p.setPen(TEXT_DIM)
        p.setFont(QFont('Segoe UI', 8, QFont.Bold))
        p.drawText(10, 21, "LEGEND:")

        # Entries
        x = 78
        y = self.height() / 2
        p.setFont(QFont('Segoe UI', 8))
        for name, color, style in self.entries:
            # Color swatch - filled rectangle for measurements,
            # solid line for upper/lower limits, dashed line for mean
            if style == 'fill':
                p.setPen(Qt.NoPen)
                p.setBrush(color)
                p.drawRect(int(x), int(y) - 5, 16, 10)
            elif style == 'line':
                p.setPen(QPen(color, 2.2))
                p.drawLine(int(x), int(y), int(x) + 16, int(y))
            elif style == 'dash':
                pen = QPen(color, 1.5)
                pen.setStyle(Qt.DashLine)
                p.setPen(pen)
                p.drawLine(int(x), int(y), int(x) + 16, int(y))
            x += 22

            # Filename or label
            display = name if len(name) <= 32 else name[:30] + '..'
            p.setPen(TEXT_PRIMARY)
            text_w = p.fontMetrics().width(display)
            p.drawText(int(x), int(y) + 4, display)
            x += text_w + 18

            # Stop if running out of room
            if x > self.width() - 40:
                p.setPen(TEXT_DIM)
                p.drawText(int(x), int(y) + 4, "...")
                break
        p.end()


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

def export_json_for_rew_iqc(filepath, freqs, upper, lower,
                             range_start, range_stop, normalization, metadata=None):
    """Export limits as JSON in the format expected by rew-iqc.

    Schema (matches rew-iqc's --create-example-mask output, version 1.1):
        {
            "name": "...",
            "version": "1.1",
            "smoothing": "1/12",       # human-readable fraction
            "ppo": 48,                  # points per octave
            "freq_range_hz": [start, stop],
            "limits": [
                {"freq_hz": 200, "upper_db": 76, "lower_db": 60},
                ...
            ],
            "metadata": { ... custom fields ... }
        }

    Note: the limit tool's other metadata (method, normalization, basis,
    sigma_mult, etc.) is preserved inside the metadata dict so factory
    operators have full provenance.
    """
    mask = (freqs >= range_start) & (freqs <= range_stop)
    points = []
    for i in np.where(mask)[0]:
        pt = {'freq_hz': float(freqs[i])}
        if upper is not None and np.isfinite(upper[i]):
            pt['upper_db'] = float(upper[i])
        if lower is not None and np.isfinite(lower[i]):
            pt['lower_db'] = float(lower[i])
        if 'upper_db' in pt or 'lower_db' in pt:
            points.append(pt)

    metadata = metadata or {}

    # Convert "1/12 octave" -> "1/12" for the top-level field
    smoothing_short = 'None'
    if 'smoothing' in metadata and metadata['smoothing']:
        s = metadata['smoothing']
        if 'octave' in s:
            smoothing_short = s.replace(' octave', '').strip()
        else:
            smoothing_short = s

    # Best name from source files / method
    name = metadata.get('name') or "REW IQC Limits"
    if metadata.get('source_files'):
        name = "Limits from {} measurements".format(len(metadata['source_files']))

    # Augment metadata with our extra fields and the floating-mode flag
    extended_meta = dict(metadata)
    extended_meta['exported_by'] = "AudioMacGyver's REW Limit Tool"
    extended_meta['normalization_mode'] = normalization

    data = {
        'name': name,
        'version': '1.1',
        'smoothing': smoothing_short,
        'ppo': 48,  # default; rew-iqc uses this as the resampling grid
        'freq_range_hz': [float(range_start), float(range_stop)],
        'limits': points,
        'metadata': extended_meta,
    }
    with open(filepath, 'w') as f:
        json.dump(data, f, indent=2)


def export_rew_limit_file(filepath, freqs, values, label='Limit'):
    """Export a single limit line as REW-importable text."""
    with open(filepath, 'w') as f:
        f.write("* {} - exported by rew_limits_gui\n".format(label))
        f.write("* Freq(Hz)  Value(dB)\n")
        for fr, v in zip(freqs, values):
            if np.isfinite(v):
                f.write("{:.3f}\t{:.3f}\n".format(fr, v))


# ---------------------------------------------------------------------------
# Main Window
# ---------------------------------------------------------------------------

class LimitWorkspace(QWidget):
    """One tab's worth of UI: plot, table, controls, all kind-aware.

    Each instance handles one of FR / THD / HOHD. Owns its own anchors,
    limits, table, and method state but shares the measurements list with
    its sibling workspaces via the parent LimitsWindow.

    Signals:
        status_message(str): emitted when the workspace wants to update
                             the parent window's status bar
    """

    LIMIT_METHODS = ['anchors', 'sigma', 'offset', 'manual']
    NORM_MODES = ['absolute', 'normalize', 'floating']
    SHAPES = ['both', 'upper', 'lower']

    status_message = pyqtSignal(str)

    def __init__(self, parent=None, kind=KIND_FR):
        super().__init__(parent)
        self.kind = kind
        self.kind_config = KIND_CONFIG[kind]

        self.measurements = []
        self._extra_dut_files = []  # extra files loaded only for DUT testing
        self._current_method = 'anchors'
        self._current_norm = 'absolute'
        self._current_shape = self.kind_config['default_shape']
        self._suppress_table_updates = False

        outer = QVBoxLayout(self)
        outer.setContentsMargins(8, 8, 8, 8)
        outer.setSpacing(6)

        # Main split
        splitter = QSplitter(Qt.Horizontal)
        outer.addWidget(splitter, stretch=1)

        # Left: plot + legend + table
        left = QWidget()
        left_lay = QVBoxLayout(left)
        left_lay.setContentsMargins(0, 0, 0, 0)
        left_lay.setSpacing(4)

        self.plot = LimitPlot(kind=kind)
        self.plot.anchors_changed.connect(self._on_anchors_changed)
        left_lay.addWidget(self.plot, stretch=2)

        self.legend = LegendWidget()
        left_lay.addWidget(self.legend)

        left_lay.addWidget(self._build_table_group(), stretch=1)
        splitter.addWidget(left)

        # Right: scrollable controls
        right_scroll = QScrollArea()
        right_scroll.setWidgetResizable(True)
        right_scroll.setFixedWidth(340)
        right_widget = QWidget()
        right_lay = QVBoxLayout(right_widget)
        right_lay.setContentsMargins(2, 2, 2, 2)
        right_lay.setSpacing(6)

        right_lay.addWidget(self._build_files_group())
        right_lay.addWidget(self._build_range_group())
        right_lay.addWidget(self._build_processing_group())
        # Harmonics selector for THD/HOHD only (FR doesn't use harmonics)
        if self.kind_config['show_harmonics']:
            right_lay.addWidget(self._build_harmonics_group())
        # Normalization only makes sense for FR (% values aren't normalized)
        self.norm_group = self._build_normalization_group()
        right_lay.addWidget(self.norm_group)
        if not self.kind_config['show_normalization']:
            self.norm_group.hide()
        right_lay.addWidget(self._build_method_group())
        right_lay.addWidget(self._build_method_settings_stack())
        right_lay.addWidget(self._build_anchor_utilities())
        right_lay.addWidget(self._build_test_group())
        right_lay.addWidget(self._build_export_group())
        right_lay.addStretch()

        right_scroll.setWidget(right_widget)
        splitter.addWidget(right_scroll)
        splitter.setSizes([940, 340])

        # Apply kind-specific defaults to the range spinners (set after
        # builders so the widgets exist)
        f_lo, f_hi = self.kind_config['default_freq_range']
        self.spin_start.setValue(int(f_lo))
        self.spin_stop.setValue(int(f_hi))
        self.plot.set_range(f_lo, f_hi)

        # Apply kind-specific default shape (THD/HOHD = upper-only)
        if self.kind_config['default_shape'] == 'upper':
            self.combo_shape.setCurrentText('Upper Only')
        elif self.kind_config['default_shape'] == 'lower':
            self.combo_shape.setCurrentText('Lower Only')

        # Column-picker only makes sense for FR (which can show SPL vs Phase
        # vs other FR-domain columns). For THD/HOHD the kind itself decides
        # the column (THD% or aggregated harmonics) so hide the dropdown.
        if self.kind != KIND_FR:
            self.lbl_col_sigma.hide()
            self.combo_column_sigma.hide()
            self.lbl_col_offset.hide()
            self.combo_column_offset.hide()

        # Adjust offset spinner units/labels for distortion tabs (% only —
        # there's no dB-relative offset for absolute % values)
        if self.kind != KIND_FR:
            self.spin_offset_up.setSuffix(' %')
            self.spin_offset_dn.setSuffix(' %')
            # And make the offset-type combo show only the % options
            self.combo_offset_type.blockSignals(True)
            self.combo_offset_type.clear()
            self.combo_offset_type.addItems(['Symmetric %', 'Asymmetric %'])
            self.combo_offset_type.setCurrentIndex(0)
            self.combo_offset_type.blockSignals(False)

        self._update_method_visibility()
        self._update_normalization_visibility()

    # ---- Builders ----

    def _build_table_group(self):
        g = QGroupBox("LIMIT TABLE")
        lay = QVBoxLayout(g)
        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(['Freq (Hz)', 'Upper (dB)', 'Lower (dB)'])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setAlternatingRowColors(True)
        self.table.itemChanged.connect(self._on_table_edited)
        lay.addWidget(self.table)

        btns = QHBoxLayout()
        btn_add = QPushButton("+ Add Row")
        btn_add.clicked.connect(self._table_add_row)
        btns.addWidget(btn_add)
        btn_del = QPushButton("- Delete Selected")
        btn_del.clicked.connect(self._table_del_row)
        btns.addWidget(btn_del)
        btns.addStretch()
        lay.addLayout(btns)
        return g

    def _build_files_group(self):
        g = QGroupBox("LOADED FILES")
        lay = QVBoxLayout(g)
        self.list_files = QListWidget()
        self.list_files.setFixedHeight(90)
        lay.addWidget(self.list_files)
        return g

    def _build_harmonics_group(self):
        """Multi-checkbox harmonic selector for THD/HOHD tabs.

        For HOHD: the curve plotted is sqrt-sum-of-squares of the checked
        harmonics, computed from REW's per-harmonic data.

        For THD: the plotted curve is REW's pre-aggregated THD% column;
        the harmonic list here is metadata that gets written into the
        exported JSON so the operator knows which harmonics REW was
        configured to include in the THD aggregate.
        """
        g = QGroupBox("HARMONICS")
        lay = QVBoxLayout(g)
        if self.kind == KIND_THD:
            lay.addWidget(QLabel(
                "Metadata: which harmonics REW aggregates into THD.\n"
                "REW's THD column is used directly for the curve."
            ))
        else:  # HOHD
            lay.addWidget(QLabel(
                "Aggregated as sqrt(sum of squares).\n"
                "Requires REW configured to report these harmonics."
            ))

        # Two columns of checkboxes (H2-H8 left, H9-H15 right) for compactness
        all_h = ['H{}'.format(i) for i in range(2, 16)]
        self.harmonic_checks = {}
        grid = QGridLayout()
        defaults = set(self.kind_config['default_harmonics'])
        for idx, h in enumerate(all_h):
            cb = QCheckBox(h)
            cb.setChecked(h in defaults)
            cb.toggled.connect(self._on_harmonics_changed)
            self.harmonic_checks[h] = cb
            row = idx // 2
            col = idx % 2
            grid.addWidget(cb, row, col)
        lay.addLayout(grid)

        # Quick presets
        btn_row = QHBoxLayout()
        btn_thd = QPushButton("THD (H2-H9)")
        btn_thd.clicked.connect(lambda: self._set_harmonics(
            ['H{}'.format(i) for i in range(2, 10)]))
        btn_row.addWidget(btn_thd)
        btn_hohd = QPushButton("HOHD (H10-H15)")
        btn_hohd.clicked.connect(lambda: self._set_harmonics(
            ['H{}'.format(i) for i in range(10, 16)]))
        btn_row.addWidget(btn_hohd)
        lay.addLayout(btn_row)

        for lbl in g.findChildren(QLabel):
            lbl.setStyleSheet("color: #888; font-size: 10px;")
            lbl.setWordWrap(True)
        return g

    def _selected_harmonics(self):
        """Return the list of currently-checked harmonic names."""
        if not self.kind_config['show_harmonics']:
            return list(self.kind_config['default_harmonics'])
        return [h for h, cb in self.harmonic_checks.items() if cb.isChecked()]

    def _set_harmonics(self, names):
        """Programmatically set checked state and trigger redraw."""
        names_set = set(names)
        for h, cb in self.harmonic_checks.items():
            cb.blockSignals(True)
            cb.setChecked(h in names_set)
            cb.blockSignals(False)
        self._on_harmonics_changed()

    def _on_harmonics_changed(self):
        # Push selection to plot and re-render. For HOHD this changes the
        # actual curve shape; for THD it's metadata-only but we still want
        # auto-scale to update.
        sel = self._selected_harmonics()
        self.plot.set_harmonics(sel)
        self._refresh_plot_with_normalization()
        self._log("{} harmonics: {}".format(self.kind, ", ".join(sel) or "(none)"))

    def _build_range_group(self):
        g = QGroupBox("FREQUENCY RANGE")
        lay = QGridLayout(g)
        lay.addWidget(QLabel("Start (Hz):"), 0, 0)
        self.spin_start = QDoubleSpinBox()
        self.spin_start.setRange(1, 100000)
        self.spin_start.setValue(20)
        self.spin_start.setDecimals(0)
        self.spin_start.valueChanged.connect(self._on_range_changed)
        lay.addWidget(self.spin_start, 0, 1)
        lay.addWidget(QLabel("Stop (Hz):"), 1, 0)
        self.spin_stop = QDoubleSpinBox()
        self.spin_stop.setRange(1, 100000)
        self.spin_stop.setValue(20000)
        self.spin_stop.setDecimals(0)
        self.spin_stop.valueChanged.connect(self._on_range_changed)
        lay.addWidget(self.spin_stop, 1, 1)
        return g

    def _build_processing_group(self):
        g = QGroupBox("DATA PROCESSING")
        lay = QGridLayout(g)

        # Smoothing
        lay.addWidget(QLabel("Smoothing:"), 0, 0)
        self.combo_smoothing = QComboBox()
        self.combo_smoothing.addItems([
            'None',
            '1/48 octave',
            '1/24 octave',
            '1/12 octave',
            '1/6 octave',
            '1/3 octave',
            '1/2 octave',
            '1 octave',
        ])
        # Default: 1/12 for FR, None for THD/HOHD (distortion is rarely smoothed)
        default_sm = self.kind_config['default_smoothing']
        idx = self.combo_smoothing.findText(default_sm)
        if idx >= 0:
            self.combo_smoothing.setCurrentIndex(idx)
        else:
            self.combo_smoothing.setCurrentIndex(0)
        self.combo_smoothing.currentIndexChanged.connect(self._on_processing_changed)
        lay.addWidget(self.combo_smoothing, 0, 1)

        # Basis curve selector
        lay.addWidget(QLabel("Limits around:"), 1, 0)
        self.combo_basis = QComboBox()
        self.combo_basis.addItem("Mean of all", -1)
        self.combo_basis.currentIndexChanged.connect(self._on_processing_changed)
        lay.addWidget(self.combo_basis, 1, 1)

        return g

    def _build_normalization_group(self):
        g = QGroupBox("NORMALIZATION")
        lay = QVBoxLayout(g)

        self.norm_buttons = QButtonGroup(self)
        self.rb_norm_abs = QRadioButton("Absolute (no normalization)")
        self.rb_norm_ref = QRadioButton("Normalize to reference frequency")
        self.rb_norm_float = QRadioButton("Floating (rew-iqc shifts at test time)")
        self.rb_norm_abs.setChecked(True)
        self.norm_buttons.addButton(self.rb_norm_abs, 0)
        self.norm_buttons.addButton(self.rb_norm_ref, 1)
        self.norm_buttons.addButton(self.rb_norm_float, 2)
        for rb in (self.rb_norm_abs, self.rb_norm_ref, self.rb_norm_float):
            rb.toggled.connect(self._on_norm_changed)
            lay.addWidget(rb)

        ref_row = QHBoxLayout()
        ref_row.addWidget(QLabel("    Ref freq:"))
        self.spin_norm_freq = QDoubleSpinBox()
        self.spin_norm_freq.setRange(20, 20000)
        self.spin_norm_freq.setValue(1000)
        self.spin_norm_freq.setDecimals(0)
        self.spin_norm_freq.valueChanged.connect(self._recompute_limits)
        ref_row.addWidget(self.spin_norm_freq)
        ref_row.addWidget(QLabel("Hz"))
        ref_row.addStretch()
        lay.addLayout(ref_row)

        # Quick-pick buttons
        quick = QHBoxLayout()
        for hz in (1000, 250, 500, 2000):
            b = QPushButton(str(hz))
            b.setFixedWidth(48)
            b.clicked.connect(lambda _, f=hz: self.spin_norm_freq.setValue(f))
            quick.addWidget(b)
        quick.addStretch()
        lay.addLayout(quick)

        return g

    def _build_method_group(self):
        g = QGroupBox("LIMIT METHOD")
        lay = QVBoxLayout(g)

        self.method_buttons = QButtonGroup(self)
        self.rb_anchors = QRadioButton("Anchor Points (click & drag)")
        self.rb_sigma = QRadioButton("Sigma (mean +/- N*sigma)")
        self.rb_offset = QRadioButton("Offset from Data (mean +/- offset)")
        self.rb_manual = QRadioButton("Manual Table Entry")
        self.rb_anchors.setChecked(True)
        for i, rb in enumerate((self.rb_anchors, self.rb_sigma, self.rb_offset, self.rb_manual)):
            self.method_buttons.addButton(rb, i)
            rb.toggled.connect(self._on_method_changed)
            lay.addWidget(rb)

        # Shape selector
        shape_row = QHBoxLayout()
        shape_row.addWidget(QLabel("Shape:"))
        self.combo_shape = QComboBox()
        self.combo_shape.addItems(['Upper + Lower', 'Upper Only', 'Lower Only'])
        self.combo_shape.currentIndexChanged.connect(self._on_shape_changed)
        shape_row.addWidget(self.combo_shape)
        lay.addLayout(shape_row)

        return g

    def _build_method_settings_stack(self):
        """Stacked widget showing settings specific to the chosen limit method."""
        self.method_stack = QStackedWidget()

        # 0: Anchors
        anchor_panel = QGroupBox("ANCHOR DRAWING")
        ap_lay = QVBoxLayout(anchor_panel)
        ap_lay.addWidget(QLabel("Click in plot to add anchors.\nRight-click anchor to delete."))

        side_row = QHBoxLayout()
        side_row.addWidget(QLabel("Adding to:"))
        self.btn_side_upper = QPushButton("UPPER")
        self.btn_side_upper.setObjectName("side_upper")
        self.btn_side_upper.setCheckable(True)
        self.btn_side_upper.setChecked(True)
        self.btn_side_upper.clicked.connect(lambda: self._set_anchor_side('upper'))
        side_row.addWidget(self.btn_side_upper)
        self.btn_side_lower = QPushButton("LOWER")
        self.btn_side_lower.setObjectName("side_lower")
        self.btn_side_lower.setCheckable(True)
        self.btn_side_lower.clicked.connect(lambda: self._set_anchor_side('lower'))
        side_row.addWidget(self.btn_side_lower)
        ap_lay.addLayout(side_row)
        self.method_stack.addWidget(anchor_panel)

        # 1: Sigma
        sigma_panel = QGroupBox("SIGMA SETTINGS")
        sg_lay = QGridLayout(sigma_panel)
        sg_lay.addWidget(QLabel("Multiplier:"), 0, 0)
        self.spin_sigma = QDoubleSpinBox()
        self.spin_sigma.setRange(0.1, 10)
        self.spin_sigma.setValue(3.0)
        self.spin_sigma.setSingleStep(0.5)
        self.spin_sigma.setDecimals(1)
        self.spin_sigma.setSuffix(" sigma")
        self.spin_sigma.valueChanged.connect(self._recompute_limits)
        sg_lay.addWidget(self.spin_sigma, 0, 1)

        self.lbl_col_sigma = QLabel("Column:")
        sg_lay.addWidget(self.lbl_col_sigma, 1, 0)
        self.combo_column_sigma = QComboBox()
        self.combo_column_sigma.currentTextChanged.connect(self._recompute_limits)
        sg_lay.addWidget(self.combo_column_sigma, 1, 1)
        sg_lay.setRowStretch(2, 1)
        self.method_stack.addWidget(sigma_panel)

        # 2: Offset
        offset_panel = QGroupBox("OFFSET SETTINGS")
        of_lay = QGridLayout(offset_panel)
        of_lay.addWidget(QLabel("Type:"), 0, 0)
        self.combo_offset_type = QComboBox()
        self.combo_offset_type.addItems(['Symmetric dB', 'Symmetric %', 'Asymmetric dB', 'Asymmetric %'])
        self.combo_offset_type.currentIndexChanged.connect(self._on_offset_type_changed)
        of_lay.addWidget(self.combo_offset_type, 0, 1)

        of_lay.addWidget(QLabel("Upper:"), 1, 0)
        self.spin_offset_up = QDoubleSpinBox()
        self.spin_offset_up.setRange(0, 200)
        self.spin_offset_up.setValue(3.0)
        self.spin_offset_up.setSingleStep(0.5)
        self.spin_offset_up.setDecimals(1)
        self.spin_offset_up.setSuffix(' dB')
        self.spin_offset_up.valueChanged.connect(self._recompute_limits)
        of_lay.addWidget(self.spin_offset_up, 1, 1)

        of_lay.addWidget(QLabel("Lower:"), 2, 0)
        self.spin_offset_dn = QDoubleSpinBox()
        self.spin_offset_dn.setRange(0, 200)
        self.spin_offset_dn.setValue(3.0)
        self.spin_offset_dn.setSingleStep(0.5)
        self.spin_offset_dn.setDecimals(1)
        self.spin_offset_dn.setSuffix(' dB')
        self.spin_offset_dn.valueChanged.connect(self._recompute_limits)
        of_lay.addWidget(self.spin_offset_dn, 2, 1)

        self.lbl_col_offset = QLabel("Column:")
        of_lay.addWidget(self.lbl_col_offset, 3, 0)
        self.combo_column_offset = QComboBox()
        self.combo_column_offset.currentTextChanged.connect(self._recompute_limits)
        of_lay.addWidget(self.combo_column_offset, 3, 1)
        of_lay.setRowStretch(4, 1)
        self.method_stack.addWidget(offset_panel)

        # 3: Manual
        manual_panel = QGroupBox("MANUAL ENTRY")
        mp_lay = QVBoxLayout(manual_panel)
        mp_lay.addWidget(QLabel("Edit values directly in the\nLimit Table at the bottom."))
        mp_lay.addStretch()
        self.method_stack.addWidget(manual_panel)

        return self.method_stack

    def _build_anchor_utilities(self):
        g = QGroupBox("ANCHOR UTILITIES")
        lay = QVBoxLayout(g)
        lay.addWidget(QLabel(
            "Tip: Compute with Sigma or Offset, then\n"
            "click 'Convert to Anchors' to fine-tune."))
        btn_clear = QPushButton("Clear All Anchors")
        btn_clear.clicked.connect(self._clear_anchors)
        lay.addWidget(btn_clear)
        btn_seed_sigma = QPushButton("Convert to Anchors (from Sigma)")
        btn_seed_sigma.clicked.connect(lambda: self._seed_anchors_from('sigma'))
        lay.addWidget(btn_seed_sigma)
        btn_seed_offset = QPushButton("Convert to Anchors (from Offset)")
        btn_seed_offset.clicked.connect(lambda: self._seed_anchors_from('offset'))
        lay.addWidget(btn_seed_offset)
        btn_seed_current = QPushButton("Convert Current Limits to Anchors")
        btn_seed_current.setObjectName("accent")
        btn_seed_current.setToolTip(
            "Take whatever limits are currently displayed and convert them\n"
            "into editable anchor points (the most flexible workflow).")
        btn_seed_current.clicked.connect(self._seed_anchors_from_current)
        lay.addWidget(btn_seed_current)
        return g

    def _seed_anchors_from_current(self):
        """Convert the currently displayed limit curves (whatever method) to anchors."""
        if self.plot.limit_freqs is None or (
            self.plot.upper_limit is None and self.plot.lower_limit is None):
            QMessageBox.information(self, 'Convert to Anchors',
                'No limits to convert. Compute limits first.')
            return
        log_anchors = np.logspace(
            np.log10(self.spin_start.value()),
            np.log10(self.spin_stop.value()), 12)
        new_anchors = []
        for f in log_anchors:
            idx = np.argmin(np.abs(self.plot.limit_freqs - f))
            if self.plot.upper_limit is not None:
                u = self.plot.upper_limit[idx]
                if np.isfinite(u):
                    new_anchors.append({'freq': float(f), 'value': float(u), 'side': 'upper'})
            if self.plot.lower_limit is not None:
                l = self.plot.lower_limit[idx]
                if np.isfinite(l):
                    new_anchors.append({'freq': float(f), 'value': float(l), 'side': 'lower'})
        self.plot.anchors = new_anchors
        self.rb_anchors.setChecked(True)
        self._compute_from_anchors()
        self._log("Converted current limits to {} anchors. Drag to fine-tune.".format(len(new_anchors)))

    def _build_test_group(self):
        g = QGroupBox("TEST DUT AGAINST LIMITS")
        lay = QVBoxLayout(g)

        # DUT source
        lay.addWidget(QLabel("Pick a measurement to test:"))
        self.combo_dut = QComboBox()
        self.combo_dut.addItem("(none)", -1)
        lay.addWidget(self.combo_dut)

        # Buttons row
        btn_row = QHBoxLayout()
        btn_test_file = QPushButton("Load File...")
        btn_test_file.setToolTip("Load a DUT measurement from a REW text file")
        btn_test_file.clicked.connect(self._load_dut_file)
        btn_row.addWidget(btn_test_file)

        btn_test_run = QPushButton("Run Test")
        btn_test_run.setObjectName("accent")
        btn_test_run.clicked.connect(self._run_test)
        btn_row.addWidget(btn_test_run)
        lay.addLayout(btn_row)

        btn_clear = QPushButton("Clear Test")
        btn_clear.clicked.connect(self._clear_test)
        lay.addWidget(btn_clear)

        # Result display
        self.lbl_test_result = QLabel("No test run.")
        self.lbl_test_result.setWordWrap(True)
        self.lbl_test_result.setStyleSheet(
            "background-color: #1a1a20; padding: 6px; border-radius: 4px; color: #888;")
        self.lbl_test_result.setMinimumHeight(60)
        lay.addWidget(self.lbl_test_result)

        return g

    def _build_export_group(self):
        g = QGroupBox("EXPORT (this tab)")
        lay = QVBoxLayout(g)
        lbl_help = QLabel(
            "JSON export combining all 3 tabs (FR/THD/HOHD) is at the\n"
            "top of the window. Below: this tab's curves only as a REW-\n"
            "importable text file."
        )
        lbl_help.setStyleSheet("color: #888; font-size: 10px;")
        lbl_help.setWordWrap(True)
        lay.addWidget(lbl_help)
        btn_rew = QPushButton("Export REW Limit Files")
        btn_rew.clicked.connect(self._export_rew)
        lay.addWidget(btn_rew)
        return g

    # ---- Logging ----

    def _log(self, msg):
        self.status_message.emit(msg)

    # ---- Measurements API (driven by parent window) ----

    def set_measurements(self, measurements):
        """Replace the measurement list with a new one and refresh UI.

        The parent LimitsWindow calls this on every tab when files are
        loaded, captured, or cleared at the window level so all three
        tabs stay in sync.
        """
        self.measurements = list(measurements)
        # Refresh the (read-only display) file list with type tags
        self.list_files.clear()
        for m in self.measurements:
            tag = m.get('data_type', 'fr').upper()
            self.list_files.addItem(QListWidgetItem(
                "[{}] {}".format(tag, m['name'])))
        if not self.measurements:
            self._extra_dut_files = []
            self.plot.clear_dut()
        self._on_data_changed()

    def _on_data_changed(self):
        self.legend.set_entries(self.measurements)
        self._refresh_columns()
        self._refresh_basis_combo()
        self._refresh_dut_combo()
        self._refresh_plot_with_normalization()
        self._recompute_limits()

    def _refresh_dut_combo(self):
        """Repopulate DUT picker with all loaded measurements + extra files."""
        current = self.combo_dut.currentData()
        self.combo_dut.blockSignals(True)
        self.combo_dut.clear()
        self.combo_dut.addItem("(none)", -1)
        for i, m in enumerate(self.measurements):
            display = m['name'] if len(m['name']) <= 30 else m['name'][:28] + '..'
            self.combo_dut.addItem(display, i)
        # Extra DUT files (loaded just for testing)
        for i, m in enumerate(self._extra_dut_files):
            display = m['name'] if len(m['name']) <= 28 else m['name'][:26] + '..'
            self.combo_dut.addItem("[file] " + display, 1000 + i)
        if current is not None:
            for i in range(self.combo_dut.count()):
                if self.combo_dut.itemData(i) == current:
                    self.combo_dut.setCurrentIndex(i)
                    break
        self.combo_dut.blockSignals(False)

    def _load_dut_file(self):
        """Load an extra REW file just for testing (not added to limit basis)."""
        path, _ = QFileDialog.getOpenFileName(
            self, "Load DUT File", "", "Text Files (*.txt);;All Files (*)")
        if not path:
            return
        try:
            m = parse_rew_file(path)
            if m is None:
                raise RuntimeError("Could not parse file.")
            self._extra_dut_files.append(m)
            self._refresh_dut_combo()
            # Auto-select the newly loaded file
            idx = self.combo_dut.findData(1000 + len(self._extra_dut_files) - 1)
            if idx >= 0:
                self.combo_dut.setCurrentIndex(idx)
            self._log("Loaded DUT file: {}".format(m['name']))
        except Exception as e:
            QMessageBox.warning(self, 'Load Error', "Failed to load: {}".format(e))

    def _get_dut_measurement(self):
        idx = self.combo_dut.currentData()
        if idx is None or idx == -1:
            return None
        if idx >= 1000:
            extra_idx = idx - 1000
            if 0 <= extra_idx < len(self._extra_dut_files):
                return self._extra_dut_files[extra_idx]
        elif 0 <= idx < len(self.measurements):
            return self.measurements[idx]
        return None

    def _run_test(self):
        if self.plot.limit_freqs is None:
            QMessageBox.warning(self, 'Test', 'Build limits first.')
            return
        m = self._get_dut_measurement()
        if m is None:
            QMessageBox.warning(self, 'Test', 'Pick a measurement to test.')
            return

        # Use the kind-aware (freqs, values) extractor so DUT test for THD/HOHD
        # uses the distortion freq axis, not the FR one
        xs, vals = self.plot._primary_xy(m)
        if xs is None or vals is None:
            QMessageBox.warning(
                self, 'Test',
                "DUT measurement doesn't contain data for the {} tab.\n\n"
                "Make sure the measurement includes the relevant column "
                "(e.g. distortion data for THD/HOHD).".format(self.kind))
            return
        vals = vals.copy()

        smoothing = self._current_smoothing()
        if smoothing > 0:
            vals = smooth_fractional_octave(xs, vals, smoothing)
        if self._current_norm == 'normalize':
            vals = normalize_curve_to_freq(xs, vals, self.spin_norm_freq.value())
        elif self._current_norm == 'floating':
            # Floating: shift DUT to best-fit limit window in the test range
            vals = self._best_fit_shift(xs, vals)

        result = evaluate_dut_against_limits(
            xs, vals,
            self.plot.limit_freqs,
            self.plot.upper_limit, self.plot.lower_limit,
            self.spin_start.value(), self.spin_stop.value())

        # Display DUT on plot
        self.plot.set_dut(xs, vals,
                          passed=result['pass'],
                          fail_freqs=result['fail_freqs'])

        # Display result
        if result.get('error'):
            text = "ERROR: " + result['error']
            color = "#888"
        elif result['pass']:
            text = "PASS\n{}/{} points within limits.\nMax overshoot: {:+.2f} dB\nMax undershoot: {:+.2f} dB".format(
                result['n_evaluated'] - result['n_fail'],
                result['n_evaluated'],
                result['max_overshoot_db'],
                result['max_undershoot_db'])
            color = "#5fcf7f"
        else:
            n_show = min(5, len(result['fail_reasons']))
            sample = '\n'.join(result['fail_reasons'][:n_show])
            extra = "" if len(result['fail_reasons']) <= n_show else \
                "\n... and {} more".format(len(result['fail_reasons']) - n_show)
            text = "FAIL\n{}/{} points failed.\nMax overshoot: {:+.2f} dB\nMax undershoot: {:+.2f} dB\n{}{}".format(
                result['n_fail'], result['n_evaluated'],
                result['max_overshoot_db'], result['max_undershoot_db'],
                sample, extra)
            color = "#f06060"

        self.lbl_test_result.setText(text)
        self.lbl_test_result.setStyleSheet(
            "background-color: #1a1a20; padding: 6px; border-radius: 4px; "
            "color: {}; font-family: Consolas; font-size: 10px;".format(color))
        self._log("Test {}: {} of {} points failed.".format(
            'PASS' if result['pass'] else 'FAIL',
            result['n_fail'], result['n_evaluated']))

    def _dut_primary_col(self, m):
        """Legacy: pick a representative column name for DUT testing.
        The actual DUT-test path uses _primary_xy now; this is kept as a
        helper for places that still want a column name string.
        """
        if self.kind == KIND_FR:
            for col in ('SPL', 'mag', 'Magnitude'):
                if col in m['data']:
                    return col
        elif self.kind == KIND_THD:
            for col in ('THD%', 'THD'):
                if col in m['data']:
                    return col
        return list(m['data'].keys())[0]

    def _best_fit_shift(self, freqs, values):
        """Shift DUT vertically to minimize squared error against the limit
        midline in the active range. Used for 'floating' normalization mode."""
        if self.plot.upper_limit is None and self.plot.lower_limit is None:
            return values
        in_range = (freqs >= self.spin_start.value()) & (freqs <= self.spin_stop.value())
        if not np.any(in_range):
            return values

        # Build target midline (mean of upper and lower where both exist)
        limit_freqs = self.plot.limit_freqs
        if self.plot.upper_limit is not None and self.plot.lower_limit is not None:
            mid = (self.plot.upper_limit + self.plot.lower_limit) / 2
        elif self.plot.upper_limit is not None:
            mid = self.plot.upper_limit
        else:
            mid = self.plot.lower_limit

        target = np.interp(freqs[in_range], limit_freqs, mid)
        dut_in = values[in_range]
        valid = np.isfinite(target) & np.isfinite(dut_in)
        if np.any(valid):
            shift = np.mean(target[valid] - dut_in[valid])
            return values + shift
        return values

    def _clear_test(self):
        self.plot.clear_dut()
        self.lbl_test_result.setText("No test run.")
        self.lbl_test_result.setStyleSheet(
            "background-color: #1a1a20; padding: 6px; border-radius: 4px; color: #888;")

    def _refresh_basis_combo(self):
        """Repopulate the basis-curve combo with all loaded measurements."""
        current = self.combo_basis.currentData()
        self.combo_basis.blockSignals(True)
        self.combo_basis.clear()
        self.combo_basis.addItem("Mean of all", -1)
        for i, m in enumerate(self.measurements):
            display = m['name'] if len(m['name']) <= 30 else m['name'][:28] + '..'
            self.combo_basis.addItem(display, i)
        # Try to restore previous selection
        if current is not None:
            for i in range(self.combo_basis.count()):
                if self.combo_basis.itemData(i) == current:
                    self.combo_basis.setCurrentIndex(i)
                    break
        self.combo_basis.blockSignals(False)

    def _refresh_columns(self):
        cols = set()
        for m in self.measurements:
            cols.update(m['data'].keys())
        for combo in (self.combo_column_sigma, self.combo_column_offset):
            current = combo.currentText()
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(sorted(cols))
            if current in cols:
                combo.setCurrentText(current)
            elif 'SPL' in cols:
                combo.setCurrentText('SPL')
            combo.blockSignals(False)

    # ---- Mode/method changes ----

    SMOOTHING_FRACTIONS = [0, 48, 24, 12, 6, 3, 2, 1]

    def _current_smoothing(self):
        """Return the selected smoothing fraction (0 = none)."""
        idx = self.combo_smoothing.currentIndex()
        if 0 <= idx < len(self.SMOOTHING_FRACTIONS):
            return self.SMOOTHING_FRACTIONS[idx]
        return 0

    def _current_basis_idx(self):
        """Return the selected basis measurement index (-1 = mean)."""
        return self.combo_basis.currentData()

    def _on_processing_changed(self):
        self._refresh_plot_with_normalization()
        self._recompute_limits()

    def _on_method_changed(self):
        if self.rb_anchors.isChecked():
            self._current_method = 'anchors'
        elif self.rb_sigma.isChecked():
            self._current_method = 'sigma'
        elif self.rb_offset.isChecked():
            self._current_method = 'offset'
        else:
            self._current_method = 'manual'
        self._update_method_visibility()
        self._recompute_limits()

    def _update_method_visibility(self):
        idx = {'anchors': 0, 'sigma': 1, 'offset': 2, 'manual': 3}[self._current_method]
        self.method_stack.setCurrentIndex(idx)

    def _on_norm_changed(self):
        if self.rb_norm_abs.isChecked():
            self._current_norm = 'absolute'
        elif self.rb_norm_ref.isChecked():
            self._current_norm = 'normalize'
        else:
            self._current_norm = 'floating'
        self._update_normalization_visibility()
        self._refresh_plot_with_normalization()
        self._recompute_limits()

    def _update_normalization_visibility(self):
        is_norm = (self._current_norm == 'normalize')
        self.spin_norm_freq.setEnabled(is_norm)

    def _on_shape_changed(self):
        idx = self.combo_shape.currentIndex()
        self._current_shape = ('both', 'upper', 'lower')[idx]
        # Update table headers to reflect shape
        labels = {
            'both': ['Freq (Hz)', 'Upper (dB)', 'Lower (dB)'],
            'upper': ['Freq (Hz)', 'Upper (dB)', '(unused)'],
            'lower': ['Freq (Hz)', '(unused)', 'Lower (dB)'],
        }[self._current_shape]
        self.table.setHorizontalHeaderLabels(labels)
        self._recompute_limits()

    def _on_offset_type_changed(self):
        t = self.combo_offset_type.currentText()
        suffix = ' dB' if 'dB' in t else ' %'
        is_sym = 'Symmetric' in t
        self.spin_offset_up.setSuffix(suffix)
        self.spin_offset_dn.setSuffix(suffix)
        self.spin_offset_dn.setEnabled(not is_sym)
        if is_sym:
            self.spin_offset_dn.setValue(self.spin_offset_up.value())
            self.spin_offset_up.valueChanged.connect(
                self.spin_offset_dn.setValue)
        self._recompute_limits()

    def _on_range_changed(self):
        self.plot.set_range(self.spin_start.value(), self.spin_stop.value())

    def _set_anchor_side(self, side):
        self.plot.set_anchor_side(side)
        self.btn_side_upper.setChecked(side == 'upper')
        self.btn_side_lower.setChecked(side == 'lower')

    # ---- Normalization application ----

    def _refresh_plot_with_normalization(self):
        """Apply smoothing + normalization to displayed measurement curves.

        Each output entry is (freqs, values) so FR data on the FR axis and
        distortion data on the distortion axis don't get crossed.
        """
        if not self.measurements:
            self.plot.set_measurements([], None)
            return

        smoothing = self._current_smoothing()
        ref_freq = self.spin_norm_freq.value() if self._current_norm == 'normalize' else None

        curves = []
        valid_measurements = []
        for m in self.measurements:
            xs, ys = self.plot._primary_xy(m)
            if xs is None or ys is None:
                continue   # skip measurements that don't have data for this kind
            vals = ys.copy()
            if smoothing > 0:
                vals = smooth_fractional_octave(xs, vals, smoothing)
            if ref_freq is not None:
                vals = normalize_curve_to_freq(xs, vals, ref_freq)
            curves.append((xs, vals))
            valid_measurements.append(m)

        # Always pass curves so smoothing is visible even without normalization
        self.plot.set_measurements(valid_measurements, curves)

    # ---- Limit recomputation ----

    def _get_normalized_stack(self, primary_col=None):
        """Build a stack of curves on a common frequency axis.

        Uses the plot's _primary_xy() so each measurement is paired with
        the correct freq axis (FR axis for FR, distortion axis for THD/HOHD).
        The legacy `primary_col` arg is ignored — the kind decides the column.
        """
        if not self.measurements:
            return None, None, None

        # Extract per-measurement (freqs, values) tuples for this kind
        raw_curves = []
        kept_indices = []
        for i, m in enumerate(self.measurements):
            xs, ys = self.plot._primary_xy(m)
            if xs is not None and ys is not None and len(xs) > 0:
                raw_curves.append((xs, ys))
                kept_indices.append(i)
        if not raw_curves:
            return None, None, None

        # Use the first valid measurement's freq axis as the reference grid.
        # Other measurements get interpolated onto it.
        ref_freqs = raw_curves[0][0]

        smoothing = self._current_smoothing()
        mode = self._current_norm
        ref_freq = self.spin_norm_freq.value() if mode == 'normalize' else None

        stacks = []
        for xs, ys in raw_curves:
            if len(xs) == len(ref_freqs) and np.allclose(xs, ref_freqs):
                vals = ys.copy()
            else:
                vals = np.interp(ref_freqs, xs, ys)
            if smoothing and smoothing > 0:
                vals = smooth_fractional_octave(ref_freqs, vals, smoothing)
            if mode == 'normalize' and ref_freq is not None:
                vals = normalize_curve_to_freq(ref_freqs, vals, ref_freq)
            stacks.append(vals)
        stack = np.array(stacks)

        basis_idx = self._current_basis_idx()
        # Translate the original measurement index to the kept-index list, if needed
        if basis_idx is not None and basis_idx >= 0 and basis_idx in kept_indices:
            local_idx = kept_indices.index(basis_idx)
            basis = stacks[local_idx]
        else:
            basis = np.mean(stack, axis=0) if len(stacks) > 1 else stacks[0]

        return ref_freqs, stack, basis

    def _recompute_limits(self):
        method = self._current_method
        if method == 'sigma':
            self._compute_sigma()
        elif method == 'offset':
            self._compute_offset()
        elif method == 'anchors':
            self._compute_from_anchors()
        else:
            self._compute_from_table()
        self._refresh_table_from_limits()

    def _compute_sigma(self):
        if len(self.measurements) < 2:
            self.plot.clear_limits()
            self._log("Sigma mode needs at least 2 measurements.")
            return
        result = self._get_normalized_stack()
        if result is None or result[1] is None:
            self.plot.clear_limits()
            self._log("No usable data for {} on this tab.".format(self.kind))
            return
        ref_freqs, stack, basis = result
        _, upper, lower = compute_sigma_limits(
            stack, basis, self.spin_sigma.value(), self._current_shape)
        self.plot.set_limit_curves(ref_freqs, upper, lower, basis)
        basis_label = self._basis_label()
        self._log("Sigma: {} sigma around {}, shape={}".format(
            self.spin_sigma.value(), basis_label, self._current_shape))

    def _compute_offset(self):
        if not self.measurements:
            self.plot.clear_limits()
            return
        result = self._get_normalized_stack()
        if result is None or result[2] is None:
            self.plot.clear_limits()
            self._log("No usable data for {} on this tab.".format(self.kind))
            return
        ref_freqs, stack, basis = result
        offset_t = self.combo_offset_type.currentText()
        otype = 'dB' if 'dB' in offset_t else '%'
        _, upper, lower = compute_offset_limits(
            basis, otype,
            self.spin_offset_up.value(),
            self.spin_offset_dn.value(),
            self._current_shape)
        self.plot.set_limit_curves(ref_freqs, upper, lower, basis)
        basis_label = self._basis_label()
        self._log("Offset: +{}/-{} {} around {}, shape={}".format(
            self.spin_offset_up.value(), self.spin_offset_dn.value(),
            otype, basis_label, self._current_shape))

    def _basis_label(self):
        idx = self._current_basis_idx()
        if idx is None or idx == -1:
            return 'mean'
        if 0 <= idx < len(self.measurements):
            return self.measurements[idx]['name']
        return 'mean'

    def _compute_from_anchors(self):
        if len(self.plot.anchors) == 0:
            self.plot.clear_limits()
            self._log("Click on the plot to add anchor points.")
            return
        if self.measurements:
            freqs = self.measurements[0]['freqs']
        else:
            freqs = np.logspace(np.log10(FREQ_MIN), np.log10(FREQ_MAX), 512)
        upper, lower = anchors_to_curves(self.plot.anchors, freqs, self._current_shape)
        self.plot.set_limit_curves(freqs, upper, lower)
        n_up = sum(1 for a in self.plot.anchors if a['side'] == 'upper')
        n_lo = sum(1 for a in self.plot.anchors if a['side'] == 'lower')
        self._log("Anchors: {} upper, {} lower (shape={})".format(
            n_up, n_lo, self._current_shape))

    def _compute_from_table(self):
        if self.table.rowCount() == 0:
            self.plot.clear_limits()
            return
        rows = []
        for r in range(self.table.rowCount()):
            try:
                f_item = self.table.item(r, 0)
                if not f_item:
                    continue
                f = float(f_item.text())
                u_item = self.table.item(r, 1)
                l_item = self.table.item(r, 2)
                u = float(u_item.text()) if u_item and u_item.text() not in ('', '(unused)') else np.nan
                l = float(l_item.text()) if l_item and l_item.text() not in ('', '(unused)') else np.nan
                rows.append((f, u, l))
            except (ValueError, AttributeError):
                continue
        if not rows:
            self.plot.clear_limits()
            return
        rows.sort(key=lambda x: x[0])
        freqs = np.array([r[0] for r in rows])
        upper = np.array([r[1] for r in rows]) if self._current_shape in ('both', 'upper') else None
        lower = np.array([r[2] for r in rows]) if self._current_shape in ('both', 'lower') else None
        self.plot.set_limit_curves(freqs, upper, lower)

    def _on_anchors_changed(self):
        if self._current_method == 'anchors':
            self._compute_from_anchors()
            self._refresh_table_from_limits()

    # ---- Table sync ----

    def _refresh_table_from_limits(self):
        if self._suppress_table_updates:
            return
        self._suppress_table_updates = True
        try:
            self.table.setRowCount(0)
            if self.plot.limit_freqs is None:
                return
            freqs = self.plot.limit_freqs
            upper = self.plot.upper_limit if self.plot.upper_limit is not None else \
                np.full_like(freqs, np.nan, dtype=float)
            lower = self.plot.lower_limit if self.plot.lower_limit is not None else \
                np.full_like(freqs, np.nan, dtype=float)
            in_range = (freqs >= self.spin_start.value()) & (freqs <= self.spin_stop.value())
            valid = (np.isfinite(upper) | np.isfinite(lower)) & in_range

            if self._current_method == 'anchors':
                anchor_freqs = sorted(set(a['freq'] for a in self.plot.anchors))
                for f in anchor_freqs:
                    idx = np.argmin(np.abs(freqs - f))
                    if valid[idx]:
                        self._add_table_row(freqs[idx], upper[idx], lower[idx])
            else:
                valid_idx = np.where(valid)[0]
                if len(valid_idx) > 30:
                    step = len(valid_idx) // 30
                    valid_idx = valid_idx[::step]
                for i in valid_idx:
                    self._add_table_row(freqs[i], upper[i], lower[i])
        finally:
            self._suppress_table_updates = False

    def _add_table_row(self, freq, upper, lower):
        r = self.table.rowCount()
        self.table.insertRow(r)
        self.table.setItem(r, 0, QTableWidgetItem("{:.1f}".format(freq)))
        if np.isfinite(upper):
            self.table.setItem(r, 1, QTableWidgetItem("{:.2f}".format(upper)))
        else:
            self.table.setItem(r, 1, QTableWidgetItem(""))
        if np.isfinite(lower):
            self.table.setItem(r, 2, QTableWidgetItem("{:.2f}".format(lower)))
        else:
            self.table.setItem(r, 2, QTableWidgetItem(""))

    def _on_table_edited(self, item):
        if self._suppress_table_updates:
            return
        if self._current_method == 'manual':
            self._compute_from_table()

    def _table_add_row(self):
        if self._current_method != 'manual':
            self.rb_manual.setChecked(True)
        r = self.table.rowCount()
        self.table.insertRow(r)
        self.table.setItem(r, 0, QTableWidgetItem("1000"))
        self.table.setItem(r, 1, QTableWidgetItem("3"))
        self.table.setItem(r, 2, QTableWidgetItem("-3"))

    def _table_del_row(self):
        rows = sorted(set(idx.row() for idx in self.table.selectedIndexes()),
                      reverse=True)
        for r in rows:
            self.table.removeRow(r)
        if self._current_method == 'manual':
            self._compute_from_table()

    # ---- Anchor utilities ----

    def _clear_anchors(self):
        self.plot.anchors = []
        self.plot.update()
        if self._current_method == 'anchors':
            self._compute_from_anchors()

    def _seed_anchors_from(self, source):
        """Seed anchors from current sigma or offset limits."""
        if source == 'sigma':
            self._compute_sigma()
        else:
            self._compute_offset()
        if self.plot.upper_limit is None and self.plot.lower_limit is None:
            QMessageBox.information(self, 'Seed Anchors',
                'Could not compute limits to seed from. Check inputs.')
            return
        log_anchors = np.logspace(
            np.log10(self.spin_start.value()),
            np.log10(self.spin_stop.value()), 10)
        self.plot.anchors = []
        for f in log_anchors:
            idx = np.argmin(np.abs(self.plot.limit_freqs - f))
            if self.plot.upper_limit is not None:
                u = self.plot.upper_limit[idx]
                if np.isfinite(u):
                    self.plot.anchors.append(
                        {'freq': float(f), 'value': float(u), 'side': 'upper'})
            if self.plot.lower_limit is not None:
                l = self.plot.lower_limit[idx]
                if np.isfinite(l):
                    self.plot.anchors.append(
                        {'freq': float(f), 'value': float(l), 'side': 'lower'})
        self.rb_anchors.setChecked(True)
        self._compute_from_anchors()
        self._log("Seeded {} anchors from {}.".format(len(self.plot.anchors), source))

    # ---- Export ----

    def get_export_dict(self):
        """Return the kind-specific portion of the limit mask JSON.

        For FR (kind=KIND_FR), returns:
            {'kind': 'FR', 'has_data': bool, 'freq_range_hz': [lo, hi],
             'smoothing': str, 'ppo': int,
             'limits': [{'freq_hz': f, 'upper_db': u, 'lower_db': l}, ...],
             'metadata': {...}}

        For THD/HOHD, returns:
            {'kind': 'THD'|'HOHD', 'has_data': bool, 'freq_range_hz': [lo, hi],
             'ppo': int, 'harmonics': ['H2', ...],
             'limits': [{'freq_hz': f, 'max_*_pct': v}, ...],
             'metadata': {...}}

        `has_data` is False if the user didn't build any limits in this tab,
        in which case the parent window omits that tab from the combined JSON.
        """
        out = {
            'kind': self.kind,
            'has_data': False,
            'freq_range_hz': [float(self.spin_start.value()), float(self.spin_stop.value())],
            'ppo': self.kind_config['default_ppo'],
            'smoothing': self.combo_smoothing.currentText(),
            'metadata': self._build_metadata(),
            'limits': [],
        }
        if self.kind_config['show_harmonics']:
            out['harmonics'] = self._selected_harmonics()

        if self.plot.limit_freqs is None:
            return out

        freqs = self.plot.limit_freqs
        upper = self.plot.upper_limit
        lower = self.plot.lower_limit
        mask = (freqs >= self.spin_start.value()) & (freqs <= self.spin_stop.value())

        points = []
        max_field = self.kind_config['json_max_field']
        for i in np.where(mask)[0]:
            pt = {'freq_hz': float(freqs[i])}
            if self.kind == KIND_FR:
                if upper is not None and np.isfinite(upper[i]):
                    pt['upper_db'] = float(upper[i])
                if lower is not None and np.isfinite(lower[i]):
                    pt['lower_db'] = float(lower[i])
                if 'upper_db' in pt or 'lower_db' in pt:
                    points.append(pt)
            else:
                # THD/HOHD: max value is the upper curve. Lower is unused for
                # %-based distortion limits — there's no minimum harmful THD.
                if upper is not None and np.isfinite(upper[i]):
                    pt[max_field] = float(upper[i])
                    points.append(pt)

        out['limits'] = points
        out['has_data'] = len(points) > 0
        return out

    def _build_metadata(self):
        """Per-tab metadata block for the JSON export."""
        return {
            'method': self._current_method,
            'shape': self._current_shape,
            'normalization': self._current_norm,
            'norm_ref_freq': self.spin_norm_freq.value() if self._current_norm == 'normalize' else None,
            'smoothing': self.combo_smoothing.currentText(),
            'basis': self._basis_label(),
            'sigma_mult': self.spin_sigma.value() if self._current_method == 'sigma' else None,
            'offset_type': self.combo_offset_type.currentText() if self._current_method == 'offset' else None,
            'offset_up': self.spin_offset_up.value() if self._current_method == 'offset' else None,
            'offset_down': self.spin_offset_dn.value() if self._current_method == 'offset' else None,
            'source_files': [m['name'] for m in self.measurements],
        }

    def _export_rew(self):
        if self.plot.limit_freqs is None:
            QMessageBox.warning(self, 'Export', 'No limits to export.')
            return
        base_path, _ = QFileDialog.getSaveFileName(
            self, 'Export REW Limit Files (base name)', '',
            'Text Files (*.txt)')
        if not base_path:
            return
        try:
            base, ext = os.path.splitext(base_path)
            ext = ext or '.txt'
            mask = (self.plot.limit_freqs >= self.spin_start.value()) & \
                   (self.plot.limit_freqs <= self.spin_stop.value())
            written = []
            if self.plot.upper_limit is not None:
                p = base + '_upper' + ext
                export_rew_limit_file(p, self.plot.limit_freqs[mask],
                                       self.plot.upper_limit[mask], 'Upper Limit')
                written.append(p)
            if self.plot.lower_limit is not None:
                p = base + '_lower' + ext
                export_rew_limit_file(p, self.plot.limit_freqs[mask],
                                       self.plot.lower_limit[mask], 'Lower Limit')
                written.append(p)
            self._log("Exported REW files: {}".format(', '.join(os.path.basename(p) for p in written)))
        except Exception as e:
            QMessageBox.critical(self, 'Export Error', str(e))


# ---------------------------------------------------------------------------
# Combined export: assemble all three tabs into one rew-iqc JSON
# ---------------------------------------------------------------------------

def export_combined_json(filepath, fr_dict, thd_dict, hohd_dict, source_files=None):
    """Write a single combined limit mask JSON matching the rew-iqc schema.

    Args:
        filepath:    output path
        fr_dict:     dict from LimitWorkspace.get_export_dict() for FR tab
        thd_dict:    same for THD tab (may have has_data=False to skip)
        hohd_dict:   same for HOHD tab (may have has_data=False to skip)
        source_files: optional list of filenames the user loaded — written
                      into the top-level metadata for traceability

    Schema (matches rew_iqc.py v1.3.0):
        {
            "name": "...",
            "version": "1.2",
            "smoothing": "1/12",
            "ppo": 48,
            "freq_range_hz": [start, stop],
            "limits": [...],            # FR (always present — the FR tab is mandatory)
            "thd_limits": {...},        # only if THD tab has data
            "hohd_limits": {...},       # only if HOHD tab has data
            "metadata": {...}
        }

    The FR tab's data populates the top-level keys. THD/HOHD go in their
    nested sections. Per-tab metadata is preserved inside each section's
    metadata block.
    """
    if not fr_dict.get('has_data'):
        raise ValueError(
            "FR tab has no limits. Build at least the magnitude (FR) "
            "limits before exporting; THD and HOHD are optional."
        )

    # Smoothing string normalization: "1/12 octave" -> "1/12"
    smoothing_short = 'None'
    sm = fr_dict.get('smoothing', '')
    if sm:
        smoothing_short = sm.replace(' octave', '').strip() or 'None'

    name = "REW IQC Limits"
    if source_files:
        name = "Limits from {} measurements".format(len(source_files))

    data = {
        'name': name,
        'version': '1.2',
        'smoothing': smoothing_short,
        'ppo': fr_dict.get('ppo', 48),
        'freq_range_hz': fr_dict['freq_range_hz'],
        'limits': fr_dict['limits'],
        'metadata': {
            'exported_by': "AudioMacGyver's REW Limit Tool v{}".format(__version__),
            'source_files': source_files or [],
            'fr': fr_dict.get('metadata', {}),
        },
    }

    if thd_dict.get('has_data'):
        thd_smoothing = thd_dict.get('smoothing', '').replace(' octave', '').strip() or 'None'
        thd_section = {
            'freq_range_hz': thd_dict['freq_range_hz'],
            'ppo': thd_dict.get('ppo', 12),
            'smoothing': thd_smoothing,
            'harmonics': thd_dict.get('harmonics', []),
            'limits': thd_dict['limits'],
        }
        data['thd_limits'] = thd_section
        data['metadata']['thd'] = thd_dict.get('metadata', {})

    if hohd_dict.get('has_data'):
        hohd_smoothing = hohd_dict.get('smoothing', '').replace(' octave', '').strip() or 'None'
        hohd_section = {
            'freq_range_hz': hohd_dict['freq_range_hz'],
            'ppo': hohd_dict.get('ppo', 12),
            'smoothing': hohd_smoothing,
            'harmonics': hohd_dict.get('harmonics', []),
            'limits': hohd_dict['limits'],
        }
        data['hohd_limits'] = hohd_section
        data['metadata']['hohd'] = hohd_dict.get('metadata', {})

    with open(filepath, 'w') as f:
        json.dump(data, f, indent=2)


# ---------------------------------------------------------------------------
# Main window: tabs + shared sources + combined export
# ---------------------------------------------------------------------------

class LimitsWindow(QMainWindow):
    """Top-level window. Owns the shared measurements list and three tabs.

    Each tab is a LimitWorkspace for one of FR / THD / HOHD. The window
    handles file loading, REW capture, and combined JSON export. When the
    user loads or clears measurements at the window level, all three tabs
    receive the update.
    """

    def __init__(self):
        super().__init__()
        self.setWindowTitle('{} v{}'.format(APP_TITLE, __version__))
        self.setMinimumSize(1300, 860)

        # Shared state (one source of truth across tabs)
        self.measurements = []

        central = QWidget()
        self.setCentralWidget(central)
        outer = QVBoxLayout(central)
        outer.setContentsMargins(8, 8, 8, 8)
        outer.setSpacing(6)

        # --- Top toolbar: Load / Capture / Clear / Combined Export ---
        outer.addLayout(self._build_top_bar())

        # --- Tab widget ---
        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        outer.addWidget(self.tabs, stretch=1)

        self.workspaces = {}
        for kind, label in [
            (KIND_FR,   'FR (Magnitude)'),
            (KIND_THD,  'THD'),
            (KIND_HOHD, 'HOHD'),
        ]:
            ws = LimitWorkspace(parent=self, kind=kind)
            ws.status_message.connect(self._on_workspace_status)
            self.tabs.addTab(ws, label)
            self.workspaces[kind] = ws

        # --- Status bar ---
        sb = QStatusBar()
        self.setStatusBar(sb)
        self._status_default = "Ready. Load REW files or capture from REW to begin."
        sb.showMessage(self._status_default)

    # ---- Top toolbar ----

    def _build_top_bar(self):
        top = QHBoxLayout()
        top.addWidget(QLabel("Sources:"))
        btn_load = QPushButton("Load REW File(s)...")
        btn_load.clicked.connect(self._load_files)
        top.addWidget(btn_load)
        btn_capture = QPushButton("Capture from REW")
        btn_capture.setObjectName("accent")
        btn_capture.clicked.connect(self._capture_from_rew)
        top.addWidget(btn_capture)
        btn_clear = QPushButton("Clear All")
        btn_clear.clicked.connect(self._clear_files)
        top.addWidget(btn_clear)

        top.addSpacing(20)

        # The big-deal button: combined export (FR + THD + HOHD in one JSON)
        btn_export = QPushButton("Export Combined JSON (rew-iqc)")
        btn_export.setObjectName("accent")
        btn_export.clicked.connect(self._export_combined_json)
        top.addWidget(btn_export)

        top.addStretch()

        self.lbl_file_count = QLabel("0 measurements loaded")
        self.lbl_file_count.setStyleSheet("color: #888;")
        top.addWidget(self.lbl_file_count)
        return top

    # ---- File operations (write through to all tabs) ----

    def _load_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Load REW Files", "", "Text Files (*.txt);;All Files (*)")
        loaded = 0
        for path in paths:
            try:
                m = parse_rew_file(path)
                if m is not None:
                    self.measurements.append(m)
                    loaded += 1
            except Exception as e:
                QMessageBox.warning(self, 'Load Error',
                                    "Failed to parse {}:\n{}".format(path, e))
        if loaded:
            self._broadcast_measurements()
        self._set_status("Loaded {} file(s). Total: {}.".format(
            loaded, len(self.measurements)))

    def _capture_from_rew(self):
        dlg = RewCaptureDialog(self)
        if dlg.exec_() == QDialog.Accepted and dlg.selected:
            for m in dlg.selected:
                self.measurements.append(m)
            self._broadcast_measurements()
            self._set_status("Captured {} measurement(s) from REW. Total: {}.".format(
                len(dlg.selected), len(self.measurements)))

    def _clear_files(self):
        self.measurements = []
        self._broadcast_measurements()
        self._set_status("All measurements cleared.")

    def _broadcast_measurements(self):
        """Push the shared measurements list to all three tabs."""
        for ws in self.workspaces.values():
            ws.set_measurements(self.measurements)
        self.lbl_file_count.setText(
            "{} measurement(s) loaded".format(len(self.measurements)))

    # ---- Status forwarding ----

    def _on_workspace_status(self, msg):
        # Prefix with the active tab name so the user knows which tab spoke
        kind = self._active_kind()
        self._set_status("[{}] {}".format(kind, msg))

    def _set_status(self, msg):
        self.statusBar().showMessage(msg)

    def _active_kind(self):
        idx = self.tabs.currentIndex()
        if idx < 0:
            return ''
        widget = self.tabs.widget(idx)
        for k, ws in self.workspaces.items():
            if ws is widget:
                return k
        return ''

    # ---- Combined export ----

    def _export_combined_json(self):
        # Snapshot each tab's current state
        fr   = self.workspaces[KIND_FR].get_export_dict()
        thd  = self.workspaces[KIND_THD].get_export_dict()
        hohd = self.workspaces[KIND_HOHD].get_export_dict()

        if not fr.get('has_data'):
            QMessageBox.warning(
                self, 'Export',
                "FR (Magnitude) tab has no limits to export.\n\n"
                "Build the FR limits first — THD and HOHD are optional, "
                "but the FR section is always required.")
            return

        path, _ = QFileDialog.getSaveFileName(
            self, 'Export Combined JSON for rew-iqc', '', 'JSON (*.json)')
        if not path:
            return

        try:
            source_files = [m['name'] for m in self.measurements]
            export_combined_json(path, fr, thd, hohd, source_files=source_files)
            tabs_with_data = ['FR']
            if thd.get('has_data'):
                tabs_with_data.append('THD')
            if hohd.get('has_data'):
                tabs_with_data.append('HOHD')
            self._set_status("Exported {} -> {}".format(
                ' + '.join(tabs_with_data), os.path.basename(path)))
        except Exception as e:
            QMessageBox.critical(self, 'Export Error', str(e))


def main():
    app = QApplication(sys.argv)
    app.setStyleSheet(STYLESHEET)
    win = LimitsWindow()
    win.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()