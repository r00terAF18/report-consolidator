# Tool Pusher DDR Consolidator

PySide6 desktop app that copies source report data into the supplied Tool Pusher
DDR workbook. Source and target fields are located by printed labels, section boundaries,
merged anchors and table headings. No Excel cell addresses are used as mappings.
The fixed reference folder for the current aggregation work is `1405.07/1405.07.03`.
Its source-to-target decisions are recorded in [the reference comparison](docs/reference-1405.07.03.md).

```bash
uv sync
uv run python -m my_qt_app.main
```

Drop a daily folder (for example `1405.07/1405.07.22`) or a month folder
(`1405.07`) onto the window, or use **Browse Folder**. A month folder processes
its dated daily subfolders independently and reports any failed days.

Each daily folder should contain:

| Report | Supported input | Copied data |
| --- | --- | --- |
| Tool Pusher DDR | `*TOOL POOSHER*.xlsx` or `*TOOL PUSHER*.xlsx` | Target with `DDR P1` and `DDR P2` sheets |
| Company DDR | Other `*DDR*.xlsx`, sheets `DDR P 1` and `DDR 2` | Depths, BHA, bit/nozzles, drilling parameters, surveys, formation, casing, BOP details/drill dates, mud properties, fuel/drill-water balances, mud materials, requests/receipts, daily and morning operations |
| Personnel | `*POB*.xlsx` | Named personnel counts, principal rig roles and department totals |
| Mechanical | `*MDR*.xlsx` or text PDF | All remarks and cumulative engine/compressor hours; parts and oil details in the companion file, explicit engine-oil consumption in the equipment table |
| Electrical | `*EDR*.xlsx` or text PDF | All remarks and cumulative engine/compressor hours; full PDF text or workbook cells in the companion file |
| Mud service | `*Mud Report*.pdf` | Remarks, requests, daily shaker hours, equipment make/type and final screen stock |

The highest numbered `REV` wins. For the same revision, Excel takes priority
over PDF. Equal-priority duplicate reports are reported as ambiguous. Temporary
Excel lock files and earlier consolidated outputs are ignored.

The original target is preserved. A `.xlsx.bak` is created once, and a separate
`*_CONSOLIDATED.xlsx` is saved beside it. Subsequent runs use `_CONSOLIDATED_2`,
`_CONSOLIDATED_3`, etc. Existing merged ranges and target formulas are retained.
Open the output in Excel or LibreOffice to recalculate formulas.
Use **Open Saved DDR** after processing to open the newly generated workbook.
The original workbook remains the input template on subsequent runs.

Only existing value slots are filled. The workbook retains its original sheets,
labels, formulas, merged ranges, styles, row heights, column widths, drawings,
images and print layout. The original Excel package is copied and only changed
cell values are patched; no worksheet is added. Formulas are never overwritten.

Each output has a separate **`*.details.json`** companion file containing the
selected report names, parsed details, all readable source workbook cells and
PDF text, alternate revisions, forecasts, notes, materials and conflicts. This
keeps data without a matching destination available without changing the DDR.
An unresolved or ambiguous label is logged and its target value is retained.
The companion's `target_values_from_sources` records every assigned destination,
including imported values that already matched the template. Company mud
readings also include their source sheet/cell in `field_sources`, and the log
shows each source-to-target mud assignment. Blank source readings are reported
explicitly. An unreadable daily/morning operation table no longer discards the
other company DDR sections.

Review the log and output before issuing a DDR:

- Source filenames determine the reporting day. Conflicting dates inside source
  workbooks/PDFs are logged; those sources are still imported for review.
- Mechanical reports own engine/compressor runtime; electrical disagreements
  are logged with both readings. Source well-identity mismatches are also logged.
- POB is counted from populated name cells, without relying on cached formulas.
  Personnel outside the six mapped principal rig roles remain in the rig crew
  total. Explicit trainee designations are counted separately.
- Operations retain their descriptions, case, subcodes and time columns. Main
  codes are translated by activity meaning into the target legend, whose category
  totals are summed separately. Drilling and drill-out hours are distinguished
  from reported bit hours. Target-only notes and pending requests are retained. Duration
  comes from the time interval where available, including crossing midnight.
  If needed, untimed continuation paragraphs are joined to the preceding operation.
  Timed rows that exceed the existing target capacity retain the section and
  are reported in the companion file.
- Mud properties use the last populated reading under the Report Time headings.
  KCl reported in lb/bbl is converted to mg/L; solids percent becomes a fraction.
  Department remarks follow the target's merged section labels, including the
  different mud section positions in early and later templates.
- Missing, unreadable or unmapped sections retain the target's existing values.
  These may need manual updating. Ambiguous oil stocks/consumption, equipment
  service schedules and pit-volume definitions are retained in the companion file
  for manual review. The MDR oil figures appear
  under a unit column and are ambiguous.
- Row/column shifts are resolved from labels and merged boundaries. Renamed,
  missing or ambiguous headings require review; the importer does not guess.
  Scanned PDFs need OCR before importing.

## Themes

Synthwave and Liquid Glass are selectable in the window.

## Compile a standalone binary

```bash
uv run python build.py
```

The executable will be in `dist/main.dist/`.
