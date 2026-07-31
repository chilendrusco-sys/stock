"""Arranca la app Streamlit y crea un tunel publico con Cloudflare (sin cuenta necesaria)."""
import re
import subprocess
import sys
import threading
from pathlib import Path

root = Path(__file__).resolve().parent
cloudflared = root / "cloudflared.exe"

streamlit_cmd = [
    sys.executable, "-m", "streamlit", "run", str(root / "app.py"),
    "--server.address", "0.0.0.0", "--server.port", "8501",
    "--server.headless", "true",
]

print("Iniciando Streamlit...")
streamlit_proc = subprocess.Popen(
    streamlit_cmd, cwd=root, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
)

if not cloudflared.exists():
    print("No se encontro cloudflared.exe en la carpeta del proyecto.")
    streamlit_proc.wait()
    sys.exit(1)

tunnel_cmd = [str(cloudflared), "tunnel", "--url", "http://localhost:8501"]
print("Creando tunel publico...")
tunnel_proc = subprocess.Popen(
    tunnel_cmd, cwd=root, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True
)

url_pattern = re.compile(r"https://[a-zA-Z0-9\-]+\.trycloudflare\.com")
found_url = threading.Event()

def watch_output():
    for line in tunnel_proc.stdout:
        match = url_pattern.search(line)
        if match and not found_url.is_set():
            found_url.set()
            print("\n=========================================")
            print(f" URL publica: {match.group(0)}")
            print("=========================================\n")

watcher = threading.Thread(target=watch_output, daemon=True)
watcher.start()

try:
    tunnel_proc.wait()
except KeyboardInterrupt:
    tunnel_proc.terminate()
    streamlit_proc.terminate()
