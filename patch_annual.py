# -*- coding: utf-8 -*-
path = r'E:\ideaFile\schedule-app\backend\main.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

original = content

# 1. SHIFT_TEXT_MAP 加年假
old = '"休": "休息", "休息": "休息",\n}'
new = '"休": "休息", "休息": "休息",\n    "年": "年假", "年假": "年假",\n}'
if old in content:
    content = content.replace(old, new, 1)
    print('SHIFT_TEXT_MAP patched')
else:
    print('SHIFT_TEXT_MAP: not found!')

# 2. SHIFT_HOURS 加年假
old = 'SHIFT_HOURS = {"值班": 12, "白班": 8, "休息": 0}'
new = 'SHIFT_HOURS = {"值班": 12, "白班": 8, "休息": 0, "年假": 8}'
if old in content:
    content = content.replace(old, new, 1)
    print('SHIFT_HOURS patched')
else:
    print('SHIFT_HOURS: not found!')

# 3. SHIFT_FILL 加年假
old = '    "休息": PatternFill(fill_type="solid", fgColor="70AD47"),\n}'
new = '    "休息": PatternFill(fill_type="solid", fgColor="70AD47"),\n    "年假": PatternFill(fill_type="solid", fgColor="FFC000"),\n}'
if old in content:
    content = content.replace(old, new, 1)
    print('SHIFT_FILL patched')
else:
    print('SHIFT_FILL: not found!')

# 4. 返回值加 annual_leave
old = '        "holiday_rest": counts["节假日休息"],\n        "mode"'
new = '        "holiday_rest": counts["节假日休息"],\n        "annual_leave": counts["年假"],\n        "mode"'
if old in content:
    content = content.replace(old, new, 1)
    print('return dict patched')
else:
    print('return dict: not found!')

if content != original:
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print('File saved.')
else:
    print('No changes made.')
