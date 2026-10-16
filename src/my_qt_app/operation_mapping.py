"""Translate activity descriptions into the categories printed on the target DDR."""
import re

from my_qt_app.report_layout import label_key

# Source and Tool Pusher code numbers differ. Match their meanings, then read
# the actual code from the target's operation legend.
CATEGORY_LABELS = {
    'drilling': ('Drilling',),
    'drill_out': ('WOC & Drill Out / Milling',),
    'fit': ('FIT/ Leak-off Test/ Casing Integrity Test',),
    'circulation': ('Condition Mud / Hole & Circulate',),
    'trips': ('Round Trips /BHA',),
    'reaming': ('Wash & Ream',),
    'coring': ('Coring',),
    'moving': ('Rig up / Rig Down / Moving',),
    'service': ('Lubricate Rig / Service',),
    'repair': ('Repair Rig',),
    'drill_line': ('Slip & Cut Drilling Line',),
    'survey': ('Deviation Survey',),
    'logging': ('Well Logging',),
    'casing': ('Run Casing & Liner',),
    'cementing': ('Cementing / Squeeze Cement',),
    'bop': ('Nipple Up / Down',),
    'bop_test': ('BOP Testing / Press. Test',),
    'well_control': ('Kick Control / Well Killing',),
}
SOURCE_CATEGORIES = {
    'Rig Up/ Tear Down / Move': 'moving', 'Drilling': 'drilling', 'Reaming': 'reaming',
    'Coring': 'coring', 'Circulate & Condition': 'circulation', 'Trips': 'trips',
    'Service/ Maintain Rig': 'service', 'Repair Rig': 'repair', 'Replacing Drill Line': 'drill_line',
    'Deviation Survey': 'survey', 'Logging': 'logging', 'Run Casing/ Liner': 'casing',
    'Cementing': 'cementing', 'Wait on Cement': 'drill_out', 'Rig Up/Down BOP': 'bop',
    'Test BOP': 'bop_test', 'Well Control': 'well_control',
}


def activity_category(description, source_name=''):
    text = str(description).upper()
    # Explicit activities take precedence over a misleading source code. In
    # the reference day, drill-out is coded Coring and FIT is coded Well Control.
    if re.search(r'\b(?:FIT|LOT)\s+(?:TEST|AS)|\bPERFORM\s+(?:FIT|LOT)\b|LEAK[- ]?OFF', text):
        return 'fit'
    if re.search(r'\bD\s*/\s*O\b|\bDRILL(?:ING)?\s+OUT\b|\bMILL(?:ING)?\b', text):
        return 'drill_out'
    if re.search(r'\b(?:CONT(?:INUE)?\.?|COMMENCE|START|RESUME)\s+(?:DRLG|DRILLING)\b', text):
        return 'drilling'
    if re.search(r'\bCBU\b|\bCIRC(?:ULATE|ULATING)?\b', text) and not re.search(r'\b(?:POOH|RIH)\b', text):
        return 'circulation'
    if re.search(r'\b(?:RIH|POOH|TRIP(?:PING)?)\b', text):
        return 'trips'
    key = label_key(source_name)
    return next((category for label, category in SOURCE_CATEGORIES.items() if label_key(label) == key), None)


def drilling_summary(records, source_categories, bit_footage):
    hours = {'drilling': 0.0, 'drill_out': 0.0}
    new_hole = 0.0
    intervals = 0
    for record in records:
        category = activity_category(record[5], source_categories.get(str(record[3]), ''))
        if category not in hours or not isinstance(record[2], (int, float)):
            continue
        hours[category] += record[2]
        if category == 'drilling':
            # Depth intervals explicitly written as F/402 TO 405M, etc.
            matches = re.findall(r'\bF(?:ROM|/)?\s*(\d+(?:\.\d+)?)\s*M?\s*TO\s*(\d+(?:\.\d+)?)', str(record[5]), re.I)
            for low, high in matches:
                if float(high) >= float(low):
                    new_hole += float(high) - float(low)
                    intervals += 1
    result = {'drilling_hours': hours['drilling'], 'drill_out_hours': hours['drill_out'],
              'rotating_hours': sum(hours.values())}
    if hours['drill_out'] and intervals and isinstance(bit_footage, (int, float)) and bit_footage >= new_hole:
        result.update(new_hole_footage=new_hole, drill_out_footage=bit_footage - new_hole)
    return result
