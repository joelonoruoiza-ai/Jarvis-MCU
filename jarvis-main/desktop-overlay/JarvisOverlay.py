import webview
import win32gui
import win32con
import win32api
import threading
import time

# The HTML content from your Swift code (abbreviated for clarity, keep your JS logic)
ORB_HTML = """
<!DOCTYPE html>
<html>
<head>
    <style>
        * { margin: 0; padding: 0; }
        html, body { 
            width: 100%; height: 100%; 
            overflow: hidden; 
            background: transparent !important; 
        }
        canvas { position: fixed; top: 0; left: 0; width: 100%; height: 100%; }
    </style>
</head>
<body>
    <canvas id="c"></canvas>
    <script type="importmap">
    { "imports": { "three": "https://cdn.jsdelivr.net/npm/three@0.170.0/build/three.module.js" } }
    </script>
    <script type="module">
        import * as THREE from 'three';
        // ... PASTE YOUR ENTIRE THREE.JS SCRIPT LOGIC HERE ...
        // (Ensure you use ws://localhost:8340 for the WebSocket)
    </script>
</body>
</html>
"""

def set_desktop_parent(window_title):
    """
    Windows Magic: Finds the worker window behind desktop icons 
    and parents our JARVIS window to it.
    """
    time.sleep(1) # Wait for window to create
    
    hwnd = win32gui.FindWindow(None, window_title)
    if not hwnd:
        return

    # 1. Make the window click-through (Transparent to Mouse)
    ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
    win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, 
                           ex_style | win32con.WS_EX_LAYERED | win32con.WS_EX_TRANSPARENT)

    # 2. Find the Desktop "WorkerW" or "Progman"
    progman = win32gui.FindWindow("Progman", None)
    win32gui.SendMessageTimeout(progman, 0x052C, 0, 0, win32con.SMTO_NORMAL, 1000)
    
    def callback(hwnd_child, list_hwnds):
        shell = win32gui.FindWindowEx(hwnd_child, 0, "SHELLDLL_DefView", None)
        if shell:
            list_hwnds.append(win32gui.FindWindowEx(0, hwnd_child, "WorkerW", None))
        return True

    worker_windows = []
    win32gui.EnumWindows(callback, worker_windows)
    
    if worker_windows[0]:
        # Set our window as a child of the desktop worker
        win32gui.SetParent(hwnd, worker_windows[0])

def run_orb():
    # Create the window with a unique title
    title = "JARVIS_ORB_OVERLAY"
    window = webview.create_window(
        title,
        html=ORB_HTML,
        transparent=True,
        frameless=True,
        width=win32api.GetSystemMetrics(0),
        height=win32api.GetSystemMetrics(1),
        on_top=False # Set to false because we parent it to desktop
    )

    # Start the "Desktop Pinning" logic in a separate thread
    threading.Thread(target=set_desktop_parent, args=(title,), daemon=True).start()
    
    # Start the webview
    webview.start()

if __name__ == "__main__":
    run_orb()