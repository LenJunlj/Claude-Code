#!/usr/bin/env python3
"""Validate generated DBC file."""
import os, re

path = 'd:/冷俊/Claude Code/work/CX835EX_VP_LDCANFD_FSCM_20251216_Fix.dbc'
print('=== Self-Check Report ===')
print()

print('1. File:', os.path.basename(path), '(', os.path.getsize(path), 'bytes)')
with open(path, 'r', encoding='windows-1252') as f:
    content = f.read()
print('   Encoding: ANSI (Windows-1252) - OK')

non_latin1 = sum(1 for ch in content if ord(ch) >= 256)
print('2. Non-ASCII bytes:', non_latin1, '(0 = pass)')

empty_cm = re.findall(r'CM_ SG_ \w+ \w+ "";', content)
print('3. Empty CM_ SG_ comments:', len(empty_cm), '(0 = pass)')

nodes = re.findall(r'BU_:\s+(.*)', content)
hyphen_nodes = [n for n in nodes[0].split() if '-' in n] if nodes else []
print('4. Node names with hyphen:', len(hyphen_nodes), '(0 = pass)')
print('   Nodes:', nodes[0] if nodes else 'N/A')

bo = len(re.findall(r'^BO_ ', content, re.MULTILINE))
sg = len(re.findall(r'^ SG_ ', content, re.MULTILINE))
print('5. Frames:', bo, '| Signals:', sg)

cm_sg = len(re.findall(r'^CM_ SG_ ', content, re.MULTILINE))
cm_bo = len(re.findall(r'^CM_ BO_ ', content, re.MULTILINE))
print('6. CM_ SG_ comments:', cm_sg, '| CM_ BO_ comments:', cm_bo)

val_count = len(re.findall(r'^VAL_ ', content, re.MULTILINE))
print('7. VAL_ entries:', val_count)

ba_cycle = len(re.findall(r'BA_ "GenMsgCycleTime"', content))
print('8. BA_ GenMsgCycleTime:', ba_cycle)

print()
print('=== Spot-check: Start Bits vs Excel ===')
checks = [
    ('VCU_GearSelectorReq', '15'),
    ('ABS_VehSpdLgt', '35'),
    ('ABS_VehSpdLgtStatus', '55'),
    ('LFSeatForeAftPosCtrl', '7'),
    ('LFSeatUpDownPosCtrl', '15'),
    ('SecondRPassengerSittingSts', '67'),
    ('IC_OdometerMasterValue', '7'),
    ('CarMode', '7'),
]
all_ok = True
for name, expected in checks:
    match = re.search(r'SG_ ' + name + r' : (\d+)', content)
    if match:
        found = match.group(1)
        status = 'OK' if found == expected else 'MISMATCH'
        if status != 'OK':
            all_ok = False
        print('   ', name, ':', found, '(expected', expected + ')', status)
    else:
        print('   ', name, ': NOT FOUND')
        all_ok = False
print()
if all_ok:
    print('All spot-checks passed!')
else:
    print('Some checks failed!')

print()
print('=== Last 10 lines of DBC ===')
lines = content.strip().split('\n')
for l in lines[-10:]:
    print(' ', l)
