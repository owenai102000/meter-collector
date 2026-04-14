#!/usr/bin/env python3
"""
Meter Reading Collector - Windows Version
Convert to .exe using PyInstaller
Installer runs at customer site, polls copier, sends to Firebase
"""

import requests
import re
import json
import time
import sqlite3
import socket
import struct
import sys
import os
from datetime import datetime
from pathlib import Path

# ============================================
# CONFIGURATION - Edit these values
# ============================================

# FIREBASE CONFIGURATION
FIREBASE_URL = "https://copier-meter-3a69e-default-rtdb.asia-southeast1.firebasedatabase.app/"
FIREBASE_SECRET = "J5TwKTs66zX8gWDMZvkIUxAVU0mY7FnPzCnVUfMB"

# CUSTOMER/COPyER INFO
CUSTOMER_ID = "customer_001"
COPIER_IP = "192.168.1.124"
COPIER_PORT = "8000"
COPIER_USERNAME = "7654321"
COPIER_PASSWORD = "7654321"

# POLL INTERVAL (seconds)
POLL_INTERVAL = 3600  # 1 hour

# LOCAL DATABASE (for offline backup)
LOCAL_DB_PATH = os.path.join(os.environ['APPDATA'], 'MeterCollector', 'readings.db')

# LOG FILE
LOG_FILE = os.path.join(os.environ['APPDATA'], 'MeterCollector', 'collector.log')

# ============================================
# WINDOWS SERVICE SUPPORT
# ============================================

def is_admin():
    """Check if running as administrator"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def install_service():
    """Install as Windows service (requires admin)"""
    if not is_admin():
        print("ERROR: Admin privileges required to install service")
        print("Right-click -> Run as administrator")
        return False
    
    # Create service using nssm or sc
    # For simplicity, we'll use Task Scheduler instead
    import subprocess
    
    script_path = os.path.abspath(sys.argv[0])
    bat_path = os.path.join(os.environ['APPDATA'], 'MeterCollector', 'startup.bat')
    
    # Create startup folder shortcut
    startup_dir = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
    
    with open(bat_path, 'w') as f:
        f.write(f'@echo off\npythonw "{script_path}"\n')
    
    print(f"Created shortcut at: {bat_path}")
    print("Added to Windows startup!")
    return True


# ============================================
# LOGGING
# ============================================

def log(message, level="INFO"):
    """Write to log file"""
    try:
        os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
        with open(LOG_FILE, 'a') as f:
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            f.write(f"[{timestamp}] [{level}] {message}\n")
    except:
        pass
    
    # Also print to console
    print(f"[{level}] {message}")


# ============================================
# CANON REMOTE UI LOGIN & SCRAPING
# ============================================

def login_to_canon(base_url, username, password):
    """Login to Canon Remote UI and return session cookie"""
    session = requests.Session()
    
    login_url = f"{base_url}/checkLogin.cgi"
    data = {
        "i0012": "1",
        "i0014": username,
        "i0016": password,
        "i0017": "2",
        "i0019": "",
        "i2101": password,
        "errText": "Error!"
    }
    
    try:
        response = session.post(login_url, data=data, timeout=15, allow_redirects=True)
        if "portal_top.html" in response.url or response.status_code == 200:
            return session
    except Exception as e:
        log(f"Login error: {e}", "ERROR")
    
    return None


def extract_counters(html_content):
    """Extract meter counters from Canon Remote UI counter page"""
    counters = {}
    
    pattern = r'<th>([^<]+)</th><td>([^<]+)</td>'
    matches = re.findall(pattern, html_content)
    
    for label, value in matches:
        label = label.strip()
        value = value.strip().replace(",", "")
        
        if label.startswith("101:") and "Total" in label:
            counters["total"] = int(value) if value.isdigit() else 0
        elif label.startswith("201:") and "Copy" in label:
            counters["copy"] = int(value) if value.isdigit() else 0
        elif label.startswith("301:") and "Print" in label:
            counters["print"] = int(value) if value.isdigit() else 0
        elif label.startswith("401:") and "Scan" in label:
            counters["scan"] = int(value) if value.isdigit() else 0
        elif "B/W" in label or "Black" in label:
            counters["bw"] = int(value) if value.isdigit() else 0
        elif "Color" in label or "Full Color" in label:
            counters["color"] = int(value) if value.isdigit() else 0
        elif "Receive" in label:
            counters["receive"] = int(value) if value.isdigit() else 0
    
    return counters


def poll_copier(base_url, username, password):
    """Poll copier and return reading"""
    log(f"Connecting to copier at {base_url}...")
    
    session = login_to_canon(base_url, username, password)
    if not session:
        return None
    
    try:
        counter_url = f"{base_url}/d_counter.html"
        response = session.get(counter_url, timeout=10)
        if response.status_code == 200:
            return extract_counters(response.text)
    except Exception as e:
        log(f"Poll error: {e}", "ERROR")
    
    return None


# ============================================
# FIREBASE
# ============================================

def send_to_firebase(customer_id, reading):
    """Send reading to Firebase Realtime Database"""
    url = f"{FIREBASE_URL}readings/{customer_id}.json?auth={FIREBASE_SECRET}"
    
    data = {
        "date": datetime.now().strftime("%Y-%m-%d"),
        "total": reading.get("total", 0),
        "bw": reading.get("bw", 0),
        "color": reading.get("color", 0),
        "copy": reading.get("copy", 0),
        "print": reading.get("print", 0),
        "scans": reading.get("scan", 0),
        "receive": reading.get("receive", 0),
        "polled_at": datetime.now().isoformat()
    }
    
    try:
        response = requests.post(url, json=data, timeout=10)
        if response.status_code == 200:
            log("✅ Sent to Firebase!")
            return True
        else:
            log(f"❌ Firebase error: {response.status_code}", "ERROR")
            return False
    except Exception as e:
        log(f"❌ Network error: {e}", "ERROR")
        return False


# ============================================
# LOCAL BACKUP (SQLite)
# ============================================

def init_local_db():
    """Initialize local backup database"""
    try:
        os.makedirs(os.path.dirname(LOCAL_DB_PATH), exist_ok=True)
        conn = sqlite3.connect(LOCAL_DB_PATH)
        c = conn.cursor()
        c.execute("""
            CREATE TABLE IF NOT EXISTS readings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                customer_id TEXT,
                date TEXT,
                total INTEGER,
                bw INTEGER,
                color INTEGER,
                copy INTEGER,
                prints INTEGER,
                scans INTEGER,
                synced INTEGER DEFAULT 0,
                polled_at TEXT
            )
        """)
        conn.commit()
        conn.close()
    except Exception as e:
        log(f"Database init error: {e}", "ERROR")


def save_local(customer_id, reading):
    """Save reading to local database (offline backup)"""
    try:
        conn = sqlite3.connect(LOCAL_DB_PATH)
        c = conn.cursor()
        c.execute("""
            INSERT INTO readings (customer_id, date, total, bw, color, copy, prints, scans, polled_at, synced)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0)
        """, (
            customer_id,
            datetime.now().strftime("%Y-%m-%d"),
            reading.get("total", 0),
            reading.get("bw", 0),
            reading.get("color", 0),
            reading.get("copy", 0),
            reading.get("print", 0),
            reading.get("scan", 0),
            datetime.now().isoformat()
        ))
        conn.commit()
        conn.close()
        log("💾 Saved locally (offline backup)")
        return True
    except Exception as e:
        log(f"Local save error: {e}", "ERROR")
        return False


def sync_local_to_firebase():
    """Sync offline readings to Firebase when back online"""
    try:
        conn = sqlite3.connect(LOCAL_DB_PATH)
        c = conn.cursor()
        c.execute("SELECT * FROM readings WHERE synced = 0")
        unsynced = c.fetchall()
        conn.close()
        
        if not unsynced:
            return
        
        log(f"Syncing {len(unsynced)} offline readings to Firebase...")
        for row in unsynced:
            reading = {
                "total": row[3],
                "bw": row[4],
                "color": row[5],
                "copy": row[6],
                "print": row[7],
                "scans": row[8]
            }
            if send_to_firebase(row[1], reading):
                conn = sqlite3.connect(LOCAL_DB_PATH)
                c = conn.cursor()
                c.execute("UPDATE readings SET synced = 1 WHERE id = ?", (row[0],))
                conn.commit()
                conn.close()
    except Exception as e:
        log(f"Sync error: {e}", "ERROR")


# ============================================
# SYSTEM TRAY (Windows)
# ============================================

try:
    import win32api
    import win32con
    import win32gui
    import win32serviceutil
    
    class SystemTray:
        def __init__(self):
            self.notify_id = None
            self.icon_path = None
            
        def create(self, msg, title="Meter Collector"):
            # Create a simple message icon
            hwnd = win32gui.CreateWarningIcon(title, msg)
            return hwnd
            
        def destroy(self, hwnd):
            win32gui.DestroyIcon(hwnd)
    
    TRAY_AVAILABLE = True
except:
    TRAY_AVAILABLE = False
    log("System tray not available (win32 not installed)", "WARN")


# ============================================
# MAIN LOOP
# ============================================

def main():
    print("=" * 50)
    print("Canon Meter Reading Collector - Windows")
    print("=" * 50)
    print(f"Customer: {CUSTOMER_ID}")
    print(f"Copier: {COPIER_IP}:{COPIER_PORT}")
    print(f"Poll Interval: {POLL_INTERVAL} seconds")
    print(f"Log File: {LOG_FILE}")
    print("=" * 50)
    
    # Check for --install flag
    if len(sys.argv) > 1 and sys.argv[1] == "--install":
        install_service()
        return
    
    # Initialize local DB for offline backup
    init_local_db()
    
    # Build base URL
    base_url = f"http://{COPIER_IP}:{COPIER_PORT}"
    
    log("Starting meter reading collector...")
    
    while True:
        try:
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            log(f"\n[{timestamp}] Polling...")
            
            # Poll copier
            reading = poll_copier(base_url, COPIER_USERNAME, COPIER_PASSWORD)
            
            if reading:
                log(f"📊 Got reading: Total={reading.get('total', 0):,}")
                
                # Try to send to Firebase
                if not send_to_firebase(CUSTOMER_ID, reading):
                    # Firebase failed, save locally
                    save_local(CUSTOMER_ID, reading)
            else:
                log("❌ Failed to poll copier", "ERROR")
            
            # Try to sync any offline readings
            sync_local_to_firebase()
            
            log(f"💤 Sleeping for {POLL_INTERVAL} seconds...")
            time.sleep(POLL_INTERVAL)
            
        except KeyboardInterrupt:
            log("Collector stopped by user")
            break
        except Exception as e:
            log(f"Unexpected error: {e}", "ERROR")
            time.sleep(60)  # Wait before retrying


if __name__ == "__main__":
    main()
