"""Fill existing Tool Pusher value slots found from report labels."""
import re
from collections import defaultdict

from openpyxl.cell.cell import MergedCell
from openpyxl.styles.numbers import is_date_format

from my_qt_app.report_layout import ReportLayout, LayoutError, label_key
from my_qt_app.operation_mapping import CATEGORY_LABELS, activity_category

# Each field names its printed label, with section limits where labels repeat.
FIELDS = {
    'planned_td': 'Planned TD (m)', 'rig_days': 'Rig Days', 'report_no': 'Report NO',
    'rig_name': 'Rig Name', 'field_name': 'Field Name', 'supervisor': 'Well Site Supervisor',
    'geologist': 'Geologist', 'night_supervisor': 'Drilling Supervisor', 'well_name': 'Well Name', 'previous_depth': 'Previous Depth MD (00:00) (m)',
    'current_depth': 'Current Depth MD (24:00) (m)', 'morning_depth': 'Morning Depth MD (m) (06:00)',
    'daily_footage': 'Daily Footage (m)', 'daily_hours': 'Daily Drilling Hours (HR)',
    'bit_type': 'Type / Model', 'bit_size': 'Size ( in )', 'bit_serial': 'Serial No.',
    'iadc': 'IADC Code', 'bit_drilled': 'Total Drilled', 'bit_hours': 'Total Hours',
    'bit_condition': 'Condition', 'nozzles': 'Nozzle Jets (/32")',
    'drilled_interval': 'Drilled Interval (m)', 'drilling_hours': 'Drilling Hours',
    'wob': 'WOB ( klb )', 'rpm': 'RPM', 'torque': 'On BTM TORQ (kft.lb)',
    'spp': 'SPP (psi)', 'flow': 'Flow Rate (gpm)',
    'mud_type': 'Mud Type', 'mud_weight': 'Mud Weight', 'viscosity': 'Viscosity',
    'pv': 'PV', 'yp': 'YP', 'filtrate': 'Filtrate (mL/30min)', 'ph': 'PH',
    'chlorides': 'Chlorides (mg/L)', 'calcium': 'Calcium (mg/L)', 'solids': 'Solid', 'kcl': 'KCl (mg/L)',
    'alkalinity': 'Alkal Mud (Pm)',
    'hole_volume': 'Hole Vol (bbl)', 'mud_loss': 'Tot. loss (bbl)',
    'wind': 'Wind SPD. Knot', 'temperature': 'temperature', 'visibility': 'Visibility (KM)',
    'formation': 'Formation Name', 'formation_top': 'Formation Top MD', 'lithology': 'Lithology',
    'casing': 'Last Casing/ Window Depth (m)', 'casing_shoe': 'Last Casing Shoe(m)', 'kop': 'KOP (m)',
    'survey_md': 'MD (m)', 'survey_tvd': 'TVD (m)', 'survey_inc': 'INC. (deg)', 'survey_az': 'AZ. (deg)',
    'bop_test': 'Last BOP Test', 'bop_size': 'BOP Size / WP', 'annular_bop': 'ANN. BOP SIZE/TYPE',
    'single_bop': 'SINGLE BOP SIZE/TYPE', 'bottom_bop': 'BOTTOM Ram SIZE/TYPE', 'shear_bop': 'SHEAR Ram SIZE/TYPE',
}
POB = {'tool_pusher': 'TOOL PUSHER', 'driller': 'DRILLER', 'assistant_driller': 'ASS.DRILLER',
       'mud_engineer': 'Mud Engineer', 'electrical': 'ELECTRECAL', 'mechanical': 'MECANIC',
       'trainee': 'TRAINEE', 'rig_crew': 'RIG CREW', 'casual_lab': 'CASUAL LAB', 'catering': 'CATRING LAB',
       'services': 'SERVICE COMP', 'client': 'CLIENT', 'security': 'SECURITY'}


class TargetReport:
    def __init__(self, workbook, log):
        layouts = [ReportLayout(s) for s in workbook]
        self.p1 = next((s for s in layouts if s.matches('Operation Time Brake Down & Remarks')), None)
        self.p2 = next((s for s in layouts if s.matches('DAILY REMARKS')), None)
        if not self.p1 or not self.p2:
            raise LayoutError('Tool Pusher DDR pages could not be identified by their headings')
        self.log = log
        self.changes = {}
        self.assignments = {}
        self.section_results = []
        self.write_count = 0
        self.count = 0
        self.labels = set()
        for label in (*FIELDS.values(), *POB.values(), 'BIT', 'BHA', 'Morning', 'Mechanic', 'Electrician', 'MUD',
                      'welder', 'Operation Time Brake Down & Remarks', 'From', 'To', 'Hour', 'Code',
                      'Oil', 'Fuel', 'Equipment', 'Working', 'Last Date', 'Next Due Date', 'Description'):
            self.labels.add(label_key(label))

    def put(self, layout, cell, value, *, clear=False):
        if value is None and not clear or isinstance(value, str) and not value.strip() and not clear:
            return
        if isinstance(cell, MergedCell):
            raise LayoutError('Refusing to write into a merged cell interior')
        if cell.data_type == 'f':
            self.log(f'Warning: formula retained in {layout.sheet.title}; source value {value!r} not assigned.')
            return
        if isinstance(cell.value, str) and label_key(cell.value) in self.labels:
            raise LayoutError(f'Refusing to replace target label {cell.value!r}')
        if isinstance(value, str) and value.startswith('='):
            raise LayoutError('Unresolved source formula cannot be used as a report value')
        updates = self.changes.setdefault(layout.sheet.title, {})
        self.assignments.setdefault(layout.sheet.title, {})[cell.coordinate] = value
        self.write_count += 1
        if cell.value == value:
            updates.pop(cell.coordinate, None)
        else:
            updates[cell.coordinate] = value
        self.count = sum(len(values) for values in self.changes.values())
        # Keep the label index and original workbook entirely unchanged in memory.

    def attempt(self, description, action):
        previous = {name: cells.copy() for name, cells in self.changes.items()}
        assignments = {name: cells.copy() for name, cells in self.assignments.items()}
        count = self.count
        writes = self.write_count
        try:
            result = action()
            self.section_results.append({'section': description, 'status': 'applied' if self.write_count > writes else 'no_assignment',
                                         'assigned_values': self.write_count - writes})
            return result
        except (LayoutError, StopIteration, ValueError) as exc:
            self.changes, self.count = previous, count
            self.assignments = assignments
            self.write_count = writes
            self.section_results.append({'section': description, 'status': 'mapping_failed', 'reason': str(exc)})
            self.log(f'Warning: {description} retained: {exc}')

    def field(self, key, value):
        if value is None or isinstance(value, str) and not value.strip(): return
        def fill():
            s = self.p1
            if key in ('daily_rop', 'bit_rop'):
                start = s.find('WELL INFORMATION' if key == 'daily_rop' else 'DRILLING DATA').row
                stop = s.find('BHA' if key == 'daily_rop' else 'Survey Data').row
                label = s.find('ROP (m/h)', start=start, end=stop)
                cell = s.right(label)
            elif key == 'bit_no':
                labels = s.matches('BIT')
                if len(labels) != 2: raise LayoutError('BIT section and number labels missing or ambiguous')
                label = max(labels, key=lambda c: c.row)
                cell = s.right(label)
            elif key == 'bha_no':
                cell = s.below(s.find('BHA'))
            elif key in ('fire_drill', 'h2s_drill', 'kick_drill'):
                labels = {'fire_drill': 'Fire', 'h2s_drill': 'H2S Gas Drill', 'kick_drill': 'Kick Drill'}
                row = s.find(labels[key]).row
                col = s.find('LAST DATE').column
                cell = s.sheet.cell(row, col)
            elif key in ('salt', 'soda_ash', 'caustic_soda', 'limestone'):
                labels = {'salt': 'Salt', 'soda_ash': 'Soda Ash', 'caustic_soda': 'Caustic Soda', 'limestone': 'Limestone Powder'}
                label = s.find(labels[key], start=s.find('Mud Mat. Inventory').row, end=s.find('SHALE SHAKERS & MUD CLEANERS').row)
                unit = s.right(label)
                cell = s.right(unit)
            elif key in ('previous_depth', 'current_depth', 'morning_depth'):
                labels = {'previous_depth': 'Previous Depth', 'current_depth': 'Current Depth', 'morning_depth': 'Morning Depth'}
                cell = s.right(s.find(labels[key], start=s.find('WELL INFORMATION').row, end=s.find('BHA').row, prefix=True))
            elif key in FIELDS:
                scope = {"start": s.find('Mud Properties').row, "end": s.find('MUD VOLUME').row} if key in (
                    'mud_type', 'mud_weight', 'viscosity', 'pv', 'yp', 'filtrate', 'ph', 'chlorides', 'calcium', 'solids', 'kcl', 'alkalinity') else {}
                cell = s.right(s.find(FIELDS[key], **scope))
            else:
                raise LayoutError(f'No scalar target mapping for {key}')
            self.put(s, cell, value)
            if cell.coordinate in self.assignments.get(s.sheet.title, {}):
                return f'{s.sheet.title}!{cell.coordinate}'
        return self.attempt(key, fill)

    def company(self, data):
        if not data:
            return
        s = self.p1
        destinations = {key: self.field(key, value) for key, value in data.get('fields', {}).items()
                        if not key.startswith(('diesel_', 'drill_water_'))}
        for key in ('mud_type', 'mud_weight', 'viscosity', 'pv', 'yp', 'filtrate',
                    'ph', 'chlorides', 'calcium', 'solids', 'kcl', 'alkalinity'):
            value = data.get('fields', {}).get(key)
            source = data.get('field_sources', {}).get(key)
            if value is None:
                self.log(f'Mud properties: {FIELDS[key]} has no company reading; existing target value retained.')
            elif source:
                destination = destinations.get(key)
                if destination:
                    self.log(f'Mud properties: company {source["sheet"]}!{source["cell"]} -> {destination}: {FIELDS[key]} = {value!r}.')
                else:
                    self.log(f'Warning: mud property {FIELDS[key]} = {value!r} could not be placed in the target.')
        for warning in data.get('warnings', []): self.log(f'Warning: source mapping: {warning}')
        if data.get('pump_liners'):
            def liners():
                heading = s.find('MUD CIRCULATION')
                stop = s.find('SLOW CIRC . RATE').row
                header = s.find('Mud Pump', start=heading.row, end=stop)
                row = s.find('Liner ID', start=header.row, end=stop).row
                for n, value in data['pump_liners'].items():
                    column = s.find(f'#{n}', start=header.row, end=header.row + 1).column
                    self.put(s, s.sheet.cell(row, column), value)
            self.attempt('Mud pump liner sizes', liners)
        if data.get('bha'):
            def bha():
                heading = s.find('BHA')
                stop = s.find('Total BHA Length (M)').row
                columns = [s.find(label, start=heading.row, end=stop).column for label in ('Item', 'OD', 'QTY', 'Length')]
                header_row = s.find('QTY', start=heading.row, end=stop).row
                rows = [r for r in range(header_row + 1, stop) if all(not isinstance(s.sheet.cell(r, c), MergedCell) for c in columns)]
                if len(data['bha']) > len(rows): raise LayoutError('BHA source exceeds existing target capacity; see companion details')
                for index, row in enumerate(rows):
                    values = data['bha'][index] if index < len(data['bha']) else (None,) * len(columns)
                    for col, value in zip(columns, values): self.put(s, s.sheet.cell(row, col), value, clear=True)
            self.attempt('BHA', bha)
        if data.get('scr'):
            def scr():
                heading = s.find('SLOW CIRC . RATE')
                stop = s.find('BOP Status').row
                pump = s.find('Mud Pump', start=heading.row, end=stop)
                for n, spm, pressure in data['scr']:
                    col = s.find(f'#{n}', start=pump.row, end=pump.row + 1).column
                    row = s.find(f'SPM {spm:g}' if isinstance(spm, (int, float)) else f'SPM {spm}', start=pump.row, end=stop).row
                    self.put(s, s.sheet.cell(row, col), pressure)
            self.attempt('Slow circulation readings', scr)
        self.inventory(data.get('inventory', []))
        self.attempt('BHA total length', lambda: self.put(s, s.right(s.find('Total BHA Length (M)')),
                     sum(entry[3] for entry in data['bha']) if data.get('bha') and all(isinstance(entry[3], (int, float)) for entry in data['bha']) else None))
        transfers = self.diesel_transfers()
        for n, screen, hours in data.get('shakers', []):
            self.shaker(n, screen=screen, hours=hours)
        for name, label in (('diesel', 'RIG Diesel (L)'), ('drill_water', 'Drill Water (BBL)')):
            def bulk(name=name, label=label):
                heading = s.find('BULK & LIQUID (TON)')
                stop = s.find('Total Variable Load').row
                row = s.find(label.split(' (')[0], start=heading.row, end=stop, prefix=True).row
                for suffix, header in (('opening', 'Opening bal.'), ('received', "Rec'd"), ('used', 'Total Used'), ('stock', 'Stock')):
                    col = s.find(header, start=heading.row, end=row).column
                    value = data.get('fields', {}).get(f'{name}_{suffix}')
                    if name == 'diesel' and suffix in ('used', 'opening') and isinstance(value, (int, float)):
                        value += transfers
                    self.put(s, s.sheet.cell(row, col), value)
            self.attempt(name, bulk)
        for key in ('operations', 'morning'):
            if data.get(key): self.attempt(key, lambda key=key: self.operations(key, data[key], data.get('operation_categories', {})))
        self.requests(data.get('requests', []))
        if data.get('received'):
            def receipts():
                heading = self.p2.find('بارنامه دریافتی')
                stop = self.p2.find('بارنامه ارسالی').row
                rows = [r for r in range(self.p2.bounds(heading)[2] + 1, stop)
                        if not isinstance(self.p2.sheet.cell(r, heading.column), MergedCell)]
                self.lines(self.p2, [self.p2.sheet.cell(r, heading.column) for r in rows], data['received'])
            self.attempt('Receipts', receipts)

    def headers(self, report_date, rig_name):
        for layout, boundary in ((self.p1, 'WELL INFORMATION'), (self.p2, 'DAILY REMARKS')):
            for label, value in (('Date', report_date), ('Rig Name', rig_name)):
                if value:
                    self.attempt(f'{layout.sheet.title} header {label}',
                                 lambda layout=layout, label=label, value=value, boundary=boundary:
                                 self.put(layout, layout.right(layout.find(label, end=layout.find(boundary).row)), value))

    def lines(self, layout, cells, values):
        if not cells: raise LayoutError('Existing value slots missing')
        for index, cell in enumerate(cells):
            if index >= len(values):
                break  # No source value means no instruction to erase a target entry.
            value = '\n'.join(map(str, values[index:])) if index == len(cells) - 1 else values[index]
            self.put(layout, cell, value)

    def operations(self, key, records, source_categories):
        s = self.p1
        heading = s.find('Operation Time Brake Down & Remarks')
        morning = s.find('Morning', start=heading.row)
        end = s.find('Related Departments Description of NPT').row
        headers = [s.find(label, start=heading.row, end=morning.row) for label in ('From', 'To', 'Hour')]
        code = [c for c in s.matches('Code', start=headers[0].row, end=headers[0].row + 1)
                if c.column > s.bounds(headers[-1])[3]]
        if len(code) != 1: raise LayoutError('Operation code header missing or ambiguous')
        headers.append(code[0])
        subcode = s.right(headers[-1])
        description = s.right(subcode)
        cols = [c.column for c in (*headers, subcode, description)]
        start = max(s.bounds(h)[2] for h in headers) + 1 if key == 'operations' else s.bounds(morning)[2] + 1
        stop = morning.row if key == 'operations' else end
        rows = [r for r in range(start, stop) if all(not isinstance(s.sheet.cell(r, c), MergedCell) for c in cols)]
        if len(records) > len(rows):
            compact = []
            for record in records:
                if compact and all(v is None for v in record[:2]) and record[2] in (None, 0) and all(v is None for v in record[3:5]):
                    compact[-1] = (*compact[-1][:5], f'{compact[-1][5]}\n{record[5]}')
                else: compact.append(record)
            records = compact
        if len(records) > len(rows): raise LayoutError('Timed operation rows exceed existing target capacity; see companion details')
        # Notes entered only on the target (for example a fuel transfer) are
        # independent report data. Reserve their existing rows rather than erase them.
        source_text = '\n'.join(str(record[5]) for record in records)
        reserved = set()
        for row in rows:
            cells = [s.sheet.cell(row, col) for col in cols]
            note = cells[-1].value
            if isinstance(note, str) and note.strip() and all(c.value is None for c in cells[:-1]):
                if label_key(note) not in label_key(source_text):
                    reserved.add(row)
        available = [row for row in rows if row not in reserved]
        if len(records) > len(available):
            raise LayoutError('Operation records plus target-only notes exceed existing capacity')
        operation_header = s.find('Operation', start=headers[0].row, end=headers[0].row + 1)
        code_header = s.find('Code', start=operation_header.row, end=operation_header.row + 1,
                             column=min(c.column for c in s.matches('Code', start=operation_header.row, end=operation_header.row + 1)))
        legend = {}
        for cell in s.cells:
            if operation_header.row < cell.row < end and cell.column == operation_header.column and cell.value:
                code_value = s.sheet.cell(cell.row, code_header.column).value
                if isinstance(code_value, (int, float)):
                    legend[label_key(cell.value)] = (cell, int(code_value))
        mapped = []
        totals = defaultdict(float)
        for record in records:
            category = activity_category(record[5], source_categories.get(str(record[3]), ''))
            options = [legend[label_key(label)] for label in CATEGORY_LABELS.get(category, ()) if label_key(label) in legend]
            if len(options) != 1:
                if isinstance(record[2], (int, float)) and record[2] > 0:
                    self.log(f'Warning: {record[2]:g} operation hours have no unique target category; times and remarks copied without a target code: {str(record[5])[:80]}')
                mapped.append((*record[:3], None, record[4], record[5]))
                continue
            label, code = options[0]
            mapped.append((*record[:3], code if record[3] is not None else None, record[4], record[5]))
            if isinstance(record[2], (int, float)):
                totals[label_key(label.value)] += record[2]
        for index, row in enumerate(available):
            values = mapped[index] if index < len(mapped) else (None,) * len(cols)
            for col, value in zip(cols, values):
                cell = s.sheet.cell(row, col)
                # Some empty Hour slots carry clock formats in the original
                # form. Store decimal hours as text there to keep both the
                # displayed duration and the original formatting intact.
                if col == headers[2].column and isinstance(value, (int, float)) and is_date_format(cell.number_format):
                    value = f'{value:g}'
                self.put(s, cell, value, clear=True)
        if key == 'operations':
            # Legend descriptions overflow across unmerged blank cells. The
            # totals lane ends immediately before From, rather than immediately
            # after the description's anchor.
            for legend_key, (label, _) in legend.items():
                slot = s.sheet.cell(label.row, headers[0].column - 1)
                if isinstance(slot, MergedCell) or slot.column <= s.bounds(label)[3]:
                    raise LayoutError('Operation total slot missing between legend and From')
                self.put(s, slot, round(totals.get(legend_key, 0), 3))
            self.log(f'Operation categories: {dict(totals)}; daily total {sum(totals.values()):g} hours.')
        if reserved: self.log(f'Preserved {len(reserved)} target-only {key} note rows.')
        self.log(f'Operation time breakdown: {len(records)} {key} rows filled with target codes and remarks.')

    def inventory(self, products):
        if not products: return
        def fill():
            s = self.p1
            heading = s.find('Mud Mat. Inventory')
            stop = s.find('SHALE SHAKERS & MUD CLEANERS').row
            by_name = {label_key(p['product']): p for p in products}
            for label in s.cells:
                if not heading.row < label.row < stop or label.column != heading.column or not label.value:
                    continue
                product = by_name.get(label_key(label.value))
                if not product:
                    self.log(f'Warning: no source inventory product matching {label.value!r}; target retained.')
                    continue
                unit = s.right(label)
                stock = s.right(unit)
                self.put(s, unit, product['unit'])
                self.put(s, stock, product['stock'])
                self.log(f"Mud material {label.value}: {product['stock']} {product['unit']} on hand.")
        self.attempt('Mud material inventory', fill)

    def diesel_transfers(self):
        notes = set()
        total = 0.0
        for layout in (self.p1, self.p2):
            for cell in layout.cells:
                text = str(cell.value or '')
                key = label_key(text)
                if key in notes or not all(word in key for word in ('ارسال', 'گازوئیل', 'دکل', 'اردوگاه')):
                    continue
                if not re.search(r'از\s*دکل.*به\s*اردوگاه', text):
                    continue
                match = re.search(r'(\d[\d,،]*)\s*لیتر', text)
                if match:
                    amount = float(match[1].replace(',', '').replace('،', ''))
                    total += amount
                    notes.add(key)
        if total:
            self.log(f'Rig diesel: adding {total:g} L of explicitly recorded rig-to-camp transfer to consumption and opening balance; transfer notes retained.')
        return total

    def personnel(self, data):
        for key, value in data.get('fields', {}).items():
            def fill(key=key, value=value):
                start = self.p1.find('POB').row
                end = self.p1.find('GEOLOGY').row
                self.put(self.p1, self.p1.right(self.p1.find(POB[key], start=start, end=end)), value)
            self.attempt(f'POB {key}', fill)
        for key, label in {'senior_tool_pusher': 'S.Tool Pusher', 'junior_tool_pusher': 'J.Tool Pusher',
                           'top_client_rep': 'Top Client. Rep.'}.items():
            if data.get('staff', {}).get(key):
                self.attempt(label, lambda key=key, label=label:
                             self.put(self.p1, self.p1.right(self.p1.find(label, end=self.p1.find('WELL INFORMATION').row)),
                                      data['staff'][key]))
        if data: self.log(f"POB: {data['total']} named personnel counted; target total formula retained.")

    def shaker(self, n, *, screen=None, hours=None, model=None, cleaner=False):
        def fill():
            s = self.p1
            heading = s.find('SHALE SHAKERS & MUD CLEANERS')
            stop = s.find('BULK & LIQUID (TON)').row
            name = f'Mud Cleaner No.{n}' if cleaner else f'Shaker No.{n}'
            row = s.find(name, start=heading.row, end=stop).row
            for label, value in (('Screen Size', screen), ('HRs', hours), ('Make/Type', model)):
                col = s.find(label, start=heading.row, end=row).column
                self.put(s, s.sheet.cell(row, col), value)
        self.attempt(f"{'Mud Cleaner' if cleaner else 'Shaker'} {n}", fill)

    def remarks(self, label, values):
        if not values: return
        def fill():
            heading = self.p2.find(label, start=self.p2.find('DAILY REMARKS').row,
                                   end=self.p2.find('EQUIPMENT ROUTINE SERVICE AND USED MATERIAL').row)
            self.lines(self.p2, self.p2.slots(heading), values)
            self.log(f'{label}: {len(values)} remark lines filled.')
        self.attempt(f'{label} remarks', fill)

    def runtimes(self, reports):
        s = self.p2
        for key, label in {'gen_1': '#1', 'gen_2': '#2', 'gen_3': '#3', 'gen_4': '#4',
                           'emergency': 'Emergency', 'comp_1': 'Main comp #1', 'comp_2': 'Main comp #2', 'comp_3': 'Main comp #3'}.items():
            readings = [(source, report['runtimes'][key]) for source, report in reports.items()
                        if source in ('mdr', 'edr') and key in report.get('runtimes', {})]
            if not readings: continue
            owner, selected = next(((source, value) for source, value in readings if source == 'mdr'), readings[0])
            if len({v for _, v in readings}) > 1:
                self.log(f'Warning: conflicting {key} hours {readings}; using {owner.upper()} mechanical equipment reading {selected:g}.')
            if readings:
                def fill(label=label, value=selected):
                    start = s.find('EQUIPMENT STATUS').row
                    end = s.find('LOGISTIC').row
                    header = s.find('Equipment', start=start, end=end, column=s.find('EQUIPMENT STATUS').column)
                    sub = s.right(s.find('Engines', start=header.row + 1, end=end, column=header.column))
                    row = s.find(label, start=header.row + 1, end=end, column=sub.column).row
                    col = s.find('WORKING', start=header.row, end=header.row + 1).column
                    cell = s.sheet.cell(row, col)
                    if label_key(cell.value) in ('off', 'on', 'standby', 'outofservice'):
                        # Cumulative hours do not establish the operating status.
                        value = f'{value:g}\n{cell.value}'
                        self.log(f'{label}: imported cumulative hours and retained target status {cell.value!r}.')
                    self.put(s, cell, value)
                self.attempt(f'{key} hours', fill)

    def requests(self, values):
        if not values: return
        def fill():
            s = self.p2
            heading = s.find('OPERATIONAL OUTSTANDING REQUESTS')
            header = s.find('Description', start=heading.row, column=s.right(s.find('Item', start=heading.row, column=heading.column)).column)
            item = s.find('Item', start=heading.row, column=heading.column)
            rows = [c.row for c in s.cells if c.row > header.row and c.column == item.column and isinstance(c.value, (int, float))]
            self.lines(s, [s.sheet.cell(r, header.column) for r in sorted(rows)], values)
        self.attempt('Operational requests', fill)

    def mud(self, data):
        if not data:
            return
        self.remarks('MUD', data.get('remarks', []))
        equipment = data.get('equipment', {})
        engineers = [equipment.get(k) for k in ('senior_mud_engineer', 'junior_mud_engineer')]
        if any(engineers):
            self.attempt('Mud engineer names', lambda: self.put(self.p1, self.p1.right(self.p1.find('Mud Engineer', end=self.p1.find('WELL INFORMATION').row)), ' / '.join(n for n in engineers if n)))
        for n in (1, 2, 3):
            makers, types = equipment.get('manufacturer', []), equipment.get('type', [])
            model = f'{makers[n - 1]} / {types[n - 1]}' if len(makers) >= n and len(types) >= n else None
            self.shaker(n, screen=data.get('installed_screens', {}).get(str(n)),
                        hours=data.get('runtimes', {}).get(f'shaker_{n}'), model=model)
        makers, types = equipment.get('manufacturer', []), equipment.get('type', [])
        if len(makers) > 3 and len(types) > 3:
            self.shaker(1, cleaner=True, model=f'{makers[3]} / {types[3]}',
                        screen=data.get('installed_screens', {}).get('cleaner_1'))
        for size, value in data.get('screens', {}).items():
            self.attempt(f'MONGOOSE {size} inventory', lambda size=size, value=value:
                         self.put(self.p2, self.p2.right(self.p2.find(f'MONGOOSE {size}')), value))
        if data.get('requests'):
            def requests():
                s = self.p2
                heading = s.find('MUD', start=s.find('DAILY REMARKS').row,
                                 end=s.find('EQUIPMENT ROUTINE SERVICE AND USED MATERIAL').row)
                # The request column is named at the top of the parallel request table.
                column = s.find('Urgent Request & Equipment Failure Report').column
                start, _, end, _ = s.bounds(heading)
                cells = [s.sheet.cell(r, column) for r in range(start, end + 1)
                         if not isinstance(s.sheet.cell(r, column), MergedCell)]
                self.lines(s, cells, data['requests'])
            self.attempt('Mud department requests', requests)

    def oil(self, data):
        for oil in data.get('oils', []):
            if oil.get('used') is None: continue
            label = label_key(oil.get('description'))
            if '1500' in label or '15w40' in label:
                def fill():
                    s = self.p2
                    start = s.find('EQUIPMENT ROUTINE SERVICE AND USED MATERIAL').row
                    stop = s.find('Fishing Tools Inventory').row
                    row = s.find('Main Generators', start=start, end=stop).row
                    col = s.find('Refill Oil(lit)', start=start, end=row).column
                    self.put(s, s.sheet.cell(row, col), oil['used'])
                self.attempt('Engine oil consumption', fill)
