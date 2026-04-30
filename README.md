# REW IQC Pass/Fail Tool

**v1.3.1** — Speaker incoming quality control via the REW 5.40+ REST API.

Automated pass/fail evaluation of loudspeaker frequency response, THD, and HOHD against configurable limit masks. Designed for factory floor IQC where an operator needs a fast, unambiguous PASS/FAIL with full traceability.

![Example IQC result](limit_tool/screenshots/iqc_example_fail.png)

*Example FAIL result — magnitude null at 4 kHz (-5.4 dB below the lower bound) and a THD spike to 9.4% against a 6.5% limit at ~250 Hz.*

## What it does

1. Connects to a running instance of REW via its localhost HTTP API (port 4735)
2. Optionally triggers a sweep measurement via the API (auto-measure mode, Pro license required)
3. Pulls frequency response magnitude and per-harmonic distortion data for the measurement
4. Evaluates magnitude against an upper/lower dB SPL envelope
5. Evaluates THD against a frequency-dependent percent ceiling (optional)
6. Evaluates HOHD (Higher-Order Harmonic Distortion) against a separate ceiling using a user-selected band of harmonics aggregated by sqrt-sum-of-squares (optional)
7. Displays a large, color-coded PASS/FAIL result (console + plot image)
8. Logs every result with serial number to a daily CSV report for traceability

A DUT must pass **all configured checks** (magnitude AND THD AND HOHD where defined) to receive an overall PASS.

## How it works

The Python script is a thin REST client. All audio I/O, sweep generation, signal processing, and FFT computation happen inside REW. The script only reads processed data and applies limit-mask logic on top.

```
                              HTTP GET/POST
   rew_iqc.py  <----------------------------->  REW 5.40+
   (Python)         localhost:4735               (Java)
       |                                            |
       v                                            v
   Limit Mask                                  Audio Interface
   (.json)                                     + Mic + DUT
       |
       v
   iqc_plots/*.png
   iqc_reports/*.csv
```

The companion **limit_tool/** GUI builds the limit-mask JSON from a batch of golden-sample measurements. See [limit_tool/README.md](limit_tool/README.md).

## Requirements

| Component | Version | Notes |
|-----------|---------|-------|
| REW | V5.40 beta or later | REST API added in V5.40. Download from [AV Nirvana](https://www.avnirvana.com/resources/categories/rew-room-eq-wizard-beta-downloads.1/) |
| REW Pro License | Required for auto-measure only | Free version supports manual mode. Purchase at [roomeqwizard.com/upgrades.html](https://www.roomeqwizard.com/upgrades.html) |
| Python | 3.9 or later | macOS: `xcode-select --install` or `brew install python` |
| pip packages | requests, numpy, matplotlib | `pip3 install requests numpy matplotlib` |

## Quick start

```bash
# 1. Clone this repo
git clone https://github.com/audiomacgyver/rew-iqc.git
cd rew-iqc

# 2. Install dependencies
pip3 install requests numpy matplotlib

# 3. Generate example limit mask (covers all three: FR, THD, HOHD)
python3 rew_iqc.py --create-example-mask limits/speaker.json

# 4. Start REW with API enabled
open -a REW.app --args -api

# 5. Run (auto-measure, requires Pro license)
python3 rew_iqc.py --limits limits/speaker.json --auto

# Or run in manual mode (sweep in REW, evaluate here)
python3 rew_iqc.py --limits limits/speaker.json
```

## Starting the REW API Server

The API server must be running before the IQC tool can connect.

**Option A — From the REW GUI:**
Go to Preferences, then the API tab, and click the button to start the API server.

**Option B — From the terminal:**
```bash
open -a REW.app --args -api
```

**Verify it's running:**
```bash
curl -s http://127.0.0.1:4735/application/commands
```

API documentation (Swagger UI) is at `http://localhost:4735` when the server is running.

## Configuring REW for Measurement

Before using the IQC tool, configure REW's measurement settings. The API uses whatever is currently set in the REW GUI:

1. **Audio I/O**: Select your audio interface input and output in Soundcard preferences
2. **Input calibration**: Load your microphone calibration file
3. **Sweep settings**: Set start frequency, end frequency, and sweep length on the Measure dialog
4. **Output level**: Set an appropriate drive level for your DUT and fixture
5. **Distortion settings**: For HOHD coverage, configure REW to report higher harmonics (Distortion settings panel). REW only computes harmonics it can isolate without aliasing — for a 192 kHz sample rate you can get clean H9 up to ~10 kHz. Higher harmonics need either a lower sweep stop frequency or a higher sample rate.
6. **Timing reference**: Configure if using acoustic or loopback timing

These settings persist across REW sessions.

## Usage modes

### Auto-measure (recommended for production)

Requires REW Pro license. Enter the serial number (or press Enter to skip), and the script triggers the sweep, waits for completion, then evaluates.

```bash
python3 rew_iqc.py --limits limits/speaker.json --auto
```

Operator workflow:

```
>> Enter serial number (or ENTER to skip, 'q' to quit): SN-00142
  Measuring...

========================================
       PASS  PASS  PASS  PASS  PASS
       ====  ====  ====  ====  ====
       [SN: SN-00142]
========================================
```

Load DUT, enter serial, see PASS/FAIL, repeat. Type `q` to quit and see yield summary.

### Manual

No Pro license needed. Run sweeps in REW, enter the serial number in the script to evaluate the selected measurement.

```bash
python3 rew_iqc.py --limits limits/speaker.json
```

### Batch

Evaluate all measurements currently loaded in REW:

```bash
python3 rew_iqc.py --limits limits/speaker.json --batch --report
```

### Single measurement

By index or UUID:

```bash
python3 rew_iqc.py --limits limits/speaker.json --measurement 1
```

## Limit mask format

Limit masks are JSON files with magnitude limits (always required), plus optional THD and HOHD limits.

### Example

```json
{
  "name": "Speaker IQC",
  "version": "1.2",
  "smoothing": "1/12",
  "ppo": 48,
  "freq_range_hz": [200, 16000],
  "limits": [
    { "freq_hz": 200,   "upper_db": 78, "lower_db": 58 },
    { "freq_hz": 1000,  "upper_db": 81, "lower_db": 71 },
    { "freq_hz": 8000,  "upper_db": 78, "lower_db": 62 },
    { "freq_hz": 16000, "upper_db": 82, "lower_db": 44 }
  ],
  "thd_limits": {
    "freq_range_hz": [200, 10000],
    "ppo": 12,
    "harmonics": ["H2", "H3", "H4", "H5", "H6", "H7", "H8", "H9"],
    "limits": [
      { "freq_hz": 200,   "max_thd_pct": 10.0 },
      { "freq_hz": 1000,  "max_thd_pct": 3.0 },
      { "freq_hz": 10000, "max_thd_pct": 10.0 }
    ]
  },
  "hohd_limits": {
    "freq_range_hz": [200, 8000],
    "ppo": 12,
    "harmonics": ["H10", "H11", "H12", "H13", "H14", "H15"],
    "limits": [
      { "freq_hz": 200,  "max_hohd_pct": 2.0 },
      { "freq_hz": 1000, "max_hohd_pct": 0.5 },
      { "freq_hz": 8000, "max_hohd_pct": 2.0 }
    ]
  },
  "metadata": {
    "part_number": "SPK-001",
    "notes": "Derived from 10 golden samples"
  }
}
```

### Magnitude fields (required)

| Field | Description |
|-------|-------------|
| `name` | Display name for plots and reports |
| `version` | Version string for traceability |
| `smoothing` | Smoothing to request from REW (e.g. `"1/12"`) |
| `ppo` | Points per octave for frequency response data |
| `freq_range_hz` | Evaluation window — violations outside are ignored |
| `limits[].freq_hz` | Anchor frequency in Hz |
| `limits[].upper_db` | Upper limit in dB SPL (optional — see below) |
| `limits[].lower_db` | Lower limit in dB SPL (optional — see below) |

`upper_db` and `lower_db` are individually optional. A point with only `upper_db` means no lower bound is enforced at that frequency; a point with only `lower_db` means no upper bound is enforced. This matches what the limit-tool GUI emits when a sigma-computed envelope's upper or lower curve doesn't extend across the entire frequency range.

### THD fields (optional)

If `thd_limits` is omitted, THD is not checked.

| Field | Description |
|-------|-------------|
| `thd_limits.freq_range_hz` | Evaluation window for THD |
| `thd_limits.ppo` | Points per octave for distortion data |
| `thd_limits.harmonics` | Which harmonics REW aggregates into THD (informational; default `H2`–`H9`) |
| `thd_limits.limits[].freq_hz` | Anchor frequency in Hz |
| `thd_limits.limits[].max_thd_pct` | Maximum THD in percent (e.g. 3.0 = 3%) |

THD evaluation uses REW's pre-aggregated THD column directly. The `harmonics` field documents which harmonics REW was configured to include for traceability.

### HOHD fields (optional)

If `hohd_limits` is omitted, HOHD is not checked.

| Field | Description |
|-------|-------------|
| `hohd_limits.freq_range_hz` | Evaluation window for HOHD |
| `hohd_limits.ppo` | Points per octave for distortion data |
| `hohd_limits.harmonics` | Which harmonics to aggregate (default `H10`–`H15`) |
| `hohd_limits.limits[].freq_hz` | Anchor frequency in Hz |
| `hohd_limits.limits[].max_hohd_pct` | Maximum HOHD in percent |

HOHD is computed by sqrt-sum-of-squares of the listed harmonic columns from REW's per-harmonic data: `HOHD = sqrt(H_a² + H_b² + ... + H_n²)`. If REW didn't compute one of the requested harmonics (typically because it would alias above Nyquist), the tool logs a warning and continues with whatever harmonics are present.

The tool interpolates linearly between anchor points, so you only need to define points where the slope changes.

### Creating a mask from golden samples

Use the [limit_tool/](limit_tool/) GUI for this — it handles the math, smoothing, and JSON format for you. Workflow:

1. Measure 5–10 known-good units under production conditions
2. Open the limit tool, **Capture from REW**, select your golden samples
3. **FR tab** — pick a method (sigma, offset, or hand-drawn anchors), generate the magnitude envelope
4. **THD tab** — set an upper-only limit for THD %
5. **HOHD tab** — same as THD but for the higher harmonics (only useful if REW computes them)
6. Click **Export Combined JSON (rew-iqc)** at the top to save a single mask file with all three sections

Or build the JSON by hand using the schema above.

## Serial number tracking

The operator is prompted for a serial number before each test. Serial numbers appear in:

- **Console output**: shown under the PASS/FAIL banner
- **Plot title**: `IQC: measurement_name [SN: xxx] | Mask: ...`
- **Plot filename**: `SN-00142_measurement_name_2026-03-25_143000.png`
- **CSV report**: dedicated `serial_number` column

Pressing Enter without typing a serial number skips it — nothing breaks, the field is just left blank.

## Output files

### Plots

Saved to `iqc_plots/`. The plot adapts to which limit sections the mask defines:

- **3 panels** (Magnitude / THD / HOHD) when all three limits are configured
- **2 panels** (Magnitude / THD) when HOHD is not defined
- **1 panel** (Magnitude only) when neither THD nor HOHD is defined

Each panel has its own pass/fail badge. The HOHD panel includes the harmonic list in its legend so it's clear what's being aggregated.

### CSV reports

Appended to `iqc_reports/iqc_report_YYYYMMDD.csv` with columns:

| Column | Description |
|--------|-------------|
| `timestamp` | Date and time of evaluation |
| `serial_number` | DUT serial number entered by operator |
| `measurement_name` | REW measurement name |
| `uuid` | REW measurement UUID |
| `limit_mask` | Mask name and version |
| `result` | Overall PASS or FAIL |
| `mag_result` | Magnitude PASS or FAIL |
| `thd_result` | THD PASS or FAIL |
| `hohd_result` | HOHD PASS or FAIL |
| `violations_summary` | Details of all violations |
| `plot_file` | Path to saved plot PNG |

A new report file is created each day. Results are appended within the day.

## Command-line reference

```
python3 rew_iqc.py [OPTIONS]

Options:
  -l, --limits PATH           Limit mask JSON file (required for evaluation)
  -m, --measurement ID        Evaluate specific measurement by index or UUID
  -b, --batch                 Evaluate all loaded measurements
  -r, --report                Write CSV report (automatic in operator mode)
  -a, --auto                  Auto-trigger sweeps (requires REW Pro)
  --show-plots                Display plots interactively
  --create-example-mask PATH  Generate example limit mask and exit
  --host HOST                 REW API host (default: http://127.0.0.1)
  --port PORT                 REW API port (default: 4735)
```

Exit codes: 0 = all passed, 1 = any failure.

## REW API notes

Implementation details discovered during development:

- `GET /measurements` returns a **dict keyed by index strings** (`"1"`, `"2"`, ...), not a list
- `GET /measurements/{id}/frequency-response` uses `"magnitude"` (singular) as the key
- `GET /measurements/{id}/distortion` returns `columnHeaders` and a 2D `data` array; some rows may have fewer columns than others (missing harmonics at certain frequencies, especially for higher harmonics that would alias above Nyquist)
- `GET /measurements/selected-uuid` returns a bare quoted string
- `POST /measure/command` with `{"command": "SPL"}` triggers a sweep; valid commands listed at `GET /measure/commands`
- **Blocking mode** may not reliably block for measurement commands — the tool polls measurement count to detect completion
- Distortion data requested with `unit=percent` returns harmonics as a percentage of the fundamental — required for THD/HOHD math

## Troubleshooting

| Problem | Solution |
|---------|----------|
| `Cannot reach REW API` | Start the API: Preferences then API then Start, or `open -a REW.app --args -api` |
| `400 Bad Request` on auto-measure | Check `curl -s http://127.0.0.1:4735/measure/commands`. Ensure Pro license is active. |
| Auto-measure fires but evaluates old data | Verify sweep runs (hear test signal, see new measurement in REW). Check audio I/O config. |
| `inhomogeneous shape` error on distortion | Fixed in v1.1.0. The parser handles ragged rows where harmonics are missing. |
| `KeyError: 'upper_db'` on mask load | Fixed in v1.3.1. Update `rew_iqc.py` — the loader now treats individual `upper_db` / `lower_db` as optional. |
| `HOHD: H10, H11... not present in REW data` | REW isn't reporting that harmonic. Either configure REW to report higher harmonics in Distortion settings, or change the mask's `hohd_limits.harmonics` list to include only what REW returns. |
| Plot doesn't open in Preview | Test: `open ~/iqc_tool/iqc_plots/somefile.png`. Check log for "Opening plot:" line. |
| `SyntaxError: from __future__` | Must be the very first line in the file. Re-copy the script from the repo. |
| urllib3 OpenSSL warning | Cosmetic. macOS system Python 3.9 uses LibreSSL. Install Python 3.10+ via Homebrew to resolve. |
| `TimeoutError: Measurement did not complete` | Default is 60s. Adjust in `measure_spl()` if your sweep is longer. |

## File structure

```
rew-iqc/
  rew_iqc.py                              # Main script (v1.3.1)
  README.md                               # This file
  CHANGELOG.md                            # Version history for both tools
  LICENSE                                  # MIT License
  limits/
    speaker.json                           # Example limit mask (customize per DUT)
  limit_tool/
    rew_limits_gui.py                      # Limit-builder GUI (v1.1.0)
    README.md                              # Limit tool docs
    requirements.txt
    screenshots/                           # GUI screenshots + IQC example
  iqc_plots/                               # (gitignored — output)
    SN-00142_L_R Unit A_2026-03-25.png    # One plot per evaluation
  iqc_reports/                             # (gitignored — output)
    iqc_report_20260325.csv                # Daily CSV report
```

## License

MIT License. See LICENSE file.

## Disclaimer

I built this for my own production use and I'm sharing it as-is. I don't have bandwidth to maintain it as an ongoing project, but I'll push updates to the repo if I make improvements on my end.

## Acknowledgments

Built on the REW V5.40 REST API by John Mulcahy. REW handles all audio I/O, sweep generation, signal processing, and frequency response computation. This tool adds limit-mask evaluation, pass/fail reporting, and operator workflow automation as an external client.

Code and documentation developed with HEAVY assistance from Claude (Anthropic).

REW is free software at [roomeqwizard.com](https://www.roomeqwizard.com).

Pro upgrade required for auto-measure.
