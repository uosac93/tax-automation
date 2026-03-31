import subprocess
import os
import threading
import urllib.request
import time
import webview

app_dir = os.path.dirname(os.path.abspath(__file__))

# node 서버 시작 (백그라운드)
server = subprocess.Popen(
    ['node', 'server.js'],
    cwd=app_dir,
    stdout=subprocess.DEVNULL,
    stderr=subprocess.DEVNULL,
    creationflags=0x08000000
)

def load_when_ready(window):
    for _ in range(50):
        try:
            urllib.request.urlopen('http://localhost:3100/api/health', timeout=1)
            window.load_url('http://localhost:3100')
            return
        except:
            time.sleep(0.15)

window = webview.create_window(
    'TAX AI',
    html='<body style="background:#191919"></body>',
    width=1129,
    height=750,
    x=150,
    y=30,
    background_color='#191919'
)

webview.start(load_when_ready, window)
server.terminate()
