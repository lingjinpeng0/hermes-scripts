#!/usr/bin/env python3
"""Write current timestamp to a file every minute."""
from datetime import datetime
import os

out_dir = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
out_dir = os.path.join(out_dir, 'hermes')
os.makedirs(out_dir, exist_ok=True)

now = datetime.now().strftime('%Y-%m-%d %H:%M:%S %A')
with open(os.path.join(out_dir, 'current_time.txt'), 'w') as f:
    f.write(now)
print(now)
