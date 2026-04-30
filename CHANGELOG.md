# Changelog

All notable changes to the rew-iqc factory tool and the bundled limit-builder GUI.

The two tools are versioned separately because they're typically updated together but can drift slightly between releases. Both follow [Semantic Versioning](https://semver.org/).

---

## rew_iqc.py — v1.3.1 (2026-04-29)

### Fixed
- `load_limit_mask` now tolerates FR limit points that have only `upper_db` or only `lower_db` instead of crashing with `KeyError`. Missing bounds are treated as no-constraint at that frequency (mathematically: +∞ for missing upper, −∞ for missing lower). This matches what the limit-tool GUI emits when a sigma-computed envelope's upper or lower curve doesn't extend across the full frequency range.

## rew_iqc.py — v1.3.0 (2026-04-29)

### Added
- HOHD (Higher-Order Harmonic Distortion) limit support, parallel to the existing THD support. Masks may now include an optional `hohd_limits` JSON section with `freq_range_hz`, `ppo`, `harmonics`, and `limits` fields. HOHD is computed by sqrt-sum-of-squares aggregation of the user-selected harmonic columns from REW's distortion data (default H10–H15).
- Per-harmonic data access via `REWClient.get_distortion_full()`, returning a dict keyed by short column name (`Fundamental`, `THD`, `Noise`, `H2`, `H3`, …). The legacy `REWClient.get_distortion()` is retained as a thin wrapper for backward compatibility.
- New module-level helper `aggregate_harmonics_pct(harmonics, selected)` for sqrt-sum-of-squares aggregation with graceful handling of NaN values from REW's ragged data.
- `LimitMask.check_hohd()` method mirroring `check_thd()`.
- `IQCResult` carries new `hohd_passed`, `hohd_details`, `hohd_freq_hz`, and `hohd_pct` fields.
- 3-panel result plot (Magnitude / THD / HOHD), automatically reduced to 2 panels when HOHD limits aren't defined and 1 panel when no distortion limits are defined.
- New `hohd_result` column in the per-day CSV report.
- `thd_harmonics` and `hohd_harmonics` lists on `LimitMask` to record which harmonics each limit covers, included in the JSON for traceability.
- `create_example_limit_mask()` now generates an example with all three limit types (FR, THD, HOHD).

### Changed
- `IQCEngine.check_measurement` now performs a single distortion fetch per measurement and runs both THD and HOHD evaluation against it (previously THD-only).
- Operator console banner and FAIL output include HOHD verdict and harmonic list.
- Default example mask version bumped to 1.2 to reflect the new schema sections.

### Backward compatibility
- Masks without `hohd_limits` are loaded normally (HOHD check is skipped).
- Masks without `thd_limits` continue to work (THD check is skipped, magnitude only).
- The CSV report adds the new `hohd_result` column at position 9, before `violations_summary` and `plot_file`. Downstream parsers that read by column name are unaffected; parsers that read by index need updating.

---

## limit_tool (rew_limits_gui.py) — v1.1.0 (2026-04-29)

### Added
- Three tabs at the top of the window: **FR (Magnitude) / THD / HOHD**. Each tab is an independent workspace for building one limit type, with its own plot, anchors, table, and method state. The shared measurement list is loaded once at the window level and pushed to all three tabs.
- Top-of-window toolbar containing **Load REW File(s)…**, **Capture from REW**, **Clear All**, and the new **Export Combined JSON (rew-iqc)** button.
- The Export Combined JSON action assembles whatever limits are present on each tab into a single JSON file matching the rew-iqc schema. FR is required; THD and HOHD are optional (omitted from the JSON if not built).
- Per-tab **HARMONICS** group on THD and HOHD tabs with H2–H15 checkboxes plus quick-preset buttons. THD defaults to H2–H9 (REW's standard THD aggregation), HOHD defaults to H10–H15. The selected harmonics are recorded in the exported JSON for traceability.
- Status bar at the bottom of the window with active-tab prefix on messages.
- Tab styling matching the dark theme — active tab in cyan with accent border.

### Changed
- Y-axis ranges adapt per tab: FR centers a 50 dB span on the data (as before), THD floors at 0% with ~30% headroom, HOHD floors at 0% with ~10% headroom.
- Default frequency ranges per tab: FR 100–20000 Hz, THD 200–10000 Hz, HOHD 200–8000 Hz.
- Default limit shape per tab: FR uses Upper + Lower, THD/HOHD default to Upper Only (% values have no meaningful lower bound).
- Default smoothing per tab: FR 1/12-octave, THD/HOHD None (distortion data is rarely smoothed).
- Offset spinner units flip to `%` on THD/HOHD tabs.
- Normalization controls hidden on THD/HOHD tabs (% values aren't normalized).
- Column-picker dropdowns inside the Sigma and Offset method panels are hidden on THD/HOHD tabs (the tab kind already determines the column).

### Fixed
- `_extract_distortion` rewritten to handle REW V5.40+'s actual response format (`columnHeaders` + `data` row arrays). Prior versions silently dropped distortion data because they were looking for an unused legacy shape with top-level `thd`/`H2`/etc. keys. The legacy shape is retained as a fallback.
- Distortion API requests now include `unit=percent` so harmonic values are usable directly for THD/HOHD math.
- Distortion data is stored on its own frequency axis (`m['dist_freqs']`) separate from the FR axis, since REW's distortion analysis runs at a different (typically coarser, narrower) PPO and freq range than its FR analysis. Plotting, smoothing, normalization, sigma/offset compute, and DUT testing all thread the correct axis through based on the active tab's kind.
