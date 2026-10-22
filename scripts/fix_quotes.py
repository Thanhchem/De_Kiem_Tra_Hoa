# -*- coding: utf-8 -*-
import sys

with open(r'c:\Antigravity_Thanh\scripts\hoa12_to36_data.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for line in lines:
    if '"content": "' in line:
        parts = line.split('"content": "')
        prefix = parts[0]
        rest = '"content": "'.join(parts[1:])
        rest_stripped = rest.rstrip('\r\n')
        if rest_stripped.endswith('",'):
            inner = rest_stripped[:-2]
            suffix = '",\n'
        elif rest_stripped.endswith('"'):
            inner = rest_stripped[:-1]
            suffix = '"\n'
        else:
            inner = rest_stripped
            suffix = '\n'
        inner = inner.replace('"', "'")
        new_lines.append(prefix + '"content": "' + inner + suffix)
    elif '"text": "' in line:
        parts = line.split('"text": "')
        prefix = parts[0]
        rest = '"text": "'.join(parts[1:])
        rest_stripped = rest.rstrip('\r\n')
        if rest_stripped.endswith('",'):
            inner = rest_stripped[:-2]
            suffix = '",\n'
        elif rest_stripped.endswith('"'):
            inner = rest_stripped[:-1]
            suffix = '"\n'
        else:
            inner = rest_stripped
            suffix = '\n'
        inner = inner.replace('"', "'")
        new_lines.append(prefix + '"text": "' + inner + suffix)
    else:
        new_lines.append(line)

with open(r'c:\Antigravity_Thanh\scripts\hoa12_to36_data.py', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Quotes fixed successfully.")
