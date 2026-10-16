# Aggregation reference: 1405.07.03

This folder is the reference for the current implementation. The original Tool
Pusher Excel file and its PDF were compared with the company DDR Excel/PDF,
POB workbook, mechanical PDF, electrical PDF and mud service PDF. The target PDF
makes the intended operation legend, totals column and existing manual notes
visible. No other daily folder was used for this implementation pass.

## Sources and destinations

| Destination in Tool Pusher DDR | Source and aggregation rule | Reference-day result |
| --- | --- | --- |
| Depths | Company DDR's midnight, end-of-day and morning MD labels; target labels determine placement | Previous 401 m; current and morning 423 m |
| Daily footage | Written new-hole intervals total 21 m; reported bit footage is 53 m, including drill-out | 32 m D/O + 21 m DRLG |
| Daily hours / Drilling Data hours | Sum timed drilling and drill-out activities; keep bit hours separate | 18 h: 12 D/O + 6 DRLG |
| Daily ROP / Bit ROP | Import their separately reported company fields | 3.67 m/h daily; 8.83 m/h bit |
| Left operation legend totals | Classify each timed activity, then sum into the matching target legend label | Trips 5 h; drill-out 12 h; drilling 6 h; circulation 0.5 h; FIT 0.5 h; total 24 h |
| Timed operation codes | Read code numbers from the target legend rather than reuse company code numbers | Drill-out becomes target code 14, although company code says 4/Coring; FIT becomes 18, although company code says 27/Well Control |
| Daily and morning remarks | Copy all activity descriptions, times and subcodes; preserve target-only untimed notes | Eight daily description rows; two morning rows totaling 6 h |
| Bit / BHA | Company drilling data and BHA table | Bit 3, BHA 4; three individual size-16 nozzles become 3*16; two HWDP stands become six joints, as shown in the target |
| Mud properties | Latest populated company Report Time reading | Bentonite, MW 67, viscosity 41, PV 13, YP 18, pH 11, chloride 7000 mg/L, calcium 340 mg/L, solids 6%; filtrate is explicitly N.C |
| Mud material stock | Match each product printed on the target with company Mud Chemical On Hand and Unit | Salt 61; Soda Ash 155; Caustic Soda 109; Bentonite 21, with their source packaging units |
| Installed shaker screens / make / type | Mud service deck-size rows and equipment headings take priority over company screen descriptions | Shakers 1/2: 2*120 + 2*100 API; shaker 3: 4*100 API; Drrick / Hyperpool |
| Shaker hours / warehouse screen inventory | Mud service daily working hours and final stock, respectively | Shakers 1/2: 18 h each; MONGOOSE 100 final stock 10, correcting target stock 14 |
| Mechanical/electrical/mud remarks | Each department's own source; wrapped numbered mechanical items stay together | Three mechanical items, five electrical items, three mud items |
| Engine and compressor runtime | Mechanical report owns mechanical cumulative runtime; electrical disagreements are logged | Compressors 7980, 8975, 6668 h; electrical report has different values |
| POB | Count actual names and designations in the personnel workbook | 129 personnel; department/role cells are filled independently of cached totals |
| Rig diesel | Company consumption plus the explicit current-target rig-to-camp transfer | 6000 L consumed + 2000 L transferred = 8000 L total used; opening 30000 L and closing 22000 L |
| Drill water | Company Used/Received/On Hand and balance equation | Opening 560 bbl; received 420; used 780; closing 200 |
| Urgent mud requests | Copy the three source requests into existing department slots | Keep the additional pending hose request and its MR number already recorded on the target |

All destinations use printed labels, section boundaries, table headers and merged
anchors. Value updates preserve the original workbook package, sheet list,
formatting, drawings, merged ranges, formulas and print settings.

Operation totals use the existing lane immediately before the From header;
blank unmerged cells following a legend description belong to its text area.
Decimal durations in Hour cells with clock formatting are stored as text so
they display as hours without changing the template's number formats. Unknown
activity categories are logged, while their times, subcodes and remarks are
still copied with a blank target code.

## Data requiring review

- The mechanical PDF's location identifies S_36, while the target, company DDR
  and electrical report identify S_31. The filename/day match, and the original
  target's engine readings match the mechanical report. This discrepancy is
  logged; the source identity still needs human review before issuing the DDR.
- Mechanical and electrical equipment hours disagree. Mechanical readings are
  selected explicitly, and both reports remain available in the companion file.
- The company casing table contains later strings without shoe depths. The
  importer retains the target casing entries rather than substituting an older
  complete casing row.
- Source pit-volume categories do not establish the same active/reserve grouping
  used on the target. These totals remain unchanged rather than guessed.
- KCl, mud alkalinity, weather, service due dates, oil consumption, filter changes,
  pump spare inventory and several other sections have no usable source values
  on this day. Existing target values are retained. Blank consumption is not zero.
- The company DDR's 22 m depth increase differs from its 21 m explicitly written
  new-hole footage. The former remains represented by the depth fields; the
  daily footage breakdown follows the operations and reported bit footage.
- Source summaries, forecast and other details without a matching existing
  destination remain in the separate details JSON. No sheet or form row is added.

Tests are deferred at the user's request. The real report is generated from this
folder; the generation log records filled sections and unresolved source issues.
