import os
import subprocess
import sys
import time
from pathlib import Path

root = Path(__file__).resolve().parent
app = root / 'app.py'

cmd = [sys.executable, '-m', 'streamlit', 'run', str(app), '--server.address', '0.0.0.0', '--server.port', '8501']

print('Starting Streamlit server...')
proc = subprocess.Popen(cmd, cwd=root, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)

try:
    while True:
        line = proc.stdout.readline()
        if not line:
            break
        print(line, end='')
        if 'Local URL' in line or 'Network URL' in line:
            break
except KeyboardInterrupt:
    proc.terminate()
    proc.wait()
