import os
import subprocess
import sys
from pathlib import Path

root = Path(__file__).resolve().parent
cmd = [sys.executable, '-m', 'streamlit', 'run', str(root / 'app.py'), '--server.address', '0.0.0.0', '--server.port', '8501']
print('Launching app...')
print('Command:', ' '.join(cmd))
subprocess.call(cmd, cwd=root)
