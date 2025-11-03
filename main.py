import json
import os
import sys
import threading
import queue
import tkinter as tk
from tkinter import ttk, scrolledtext
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright, Playwright, TimeoutError, expect
import warnings
import pyxlsb
import csv
import tempfile
import shutil
import time
from datetime import date, timedelta

from Tasks import Login_and_Navigation


warnings.filterwarnings("ignore", category=UserWarning)


def get_playwright_browser_path():
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
        chromium_path = os.path.join(base_path, "ms-playwright", "chromium-1187", "chrome-win", "chrome.exe")
    else:
        base_path = r"C:\Users\perna\AppData\Local"

        # Join the rest of the Playwright folder path
        chromium_path = os.path.join(
            base_path,
            "ms-playwright",
            "chromium-1187",
            "chrome-win",
            "chrome.exe"
        )
   
    if chromium_path and not os.path.exists(chromium_path):
        raise FileNotFoundError(f"Chromium executable not found at {chromium_path}")

    return chromium_path


# --- GUI UPDATE FUNCTION ---
def update_gui(queue_instance, status_label, progress_bar, log_text):
    """Checks the queue for messages from the worker thread and updates the GUI."""
    try:
        while True:
            message_type, value = queue_instance.get_nowait()
            if message_type == "status":
                status_label.config(text=value)
                log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {value}\n")
                log_text.see(tk.END)
            elif message_type == "progress":
                progress_bar['value'] = value
            elif message_type == "done":
                status_label.config(text="Processo Concluído!")
                progress_bar['value'] = 100
                return # Stop checking
    except queue.Empty:
        pass
    status_label.after(100, lambda: update_gui(queue_instance, status_label, progress_bar, log_text))


def load_credentials():
    """Loads Credencial.json from the same directory as the running script or executable."""
    base_path = os.path.dirname(os.path.abspath(sys.argv[0]))
    cred_path = os.path.join(base_path, "credencial.json")

    if not os.path.exists(cred_path):
        raise FileNotFoundError(f"Credencial.json not found in: {cred_path}")

    with open(cred_path, "r", encoding="utf-8") as f:
        return json.load(f)



# list of profile items we want to copy (small & useful)
PROFILE_ITEMS = [
    "Cookies",                 # cookies sqlite
    "Preferences",             # prefs JSON
    "Local State",             # local state
    "Local Storage",           # local storage folder
    "Network",                 # may contain network tokens (sometimes)
    "Web Data",                # form data / autofill
    "Login Data",              # saved logins (may be locked)
    "Extensions",              # extensions (optional)
    "IndexedDB",               # optional
    "Service Worker",          # optional
    "Session Storage"          # optional
]

def safe_copy(src_root: str, dst_root: str, items: list):
    """
    Copy only specific files/folders from src_root into dst_root safely,
    skipping locked files and ignoring errors.
    """
    os.makedirs(dst_root, exist_ok=True)
    for name in items:
        src_path = os.path.join(src_root, name)
        dst_path = os.path.join(dst_root, name)
        try:
            if os.path.isdir(src_path):
                # copytree with ignore errors: walk and copy files individually
                os.makedirs(dst_path, exist_ok=True)
                for dirpath, dirnames, filenames in os.walk(src_path):
                    rel = os.path.relpath(dirpath, src_path)
                    target_dir = os.path.join(dst_path, rel) if rel != '.' else dst_path
                    os.makedirs(target_dir, exist_ok=True)
                    for fname in filenames:
                        sfile = os.path.join(dirpath, fname)
                        tfile = os.path.join(target_dir, fname)
                        try:
                            shutil.copy2(sfile, tfile)
                        except Exception:
                            # skip locked or unreadable files
                            continue
            elif os.path.isfile(src_path):
                os.makedirs(os.path.dirname(dst_path), exist_ok=True)
                try:
                    shutil.copy2(src_path, dst_path)
                except Exception:
                    # skip locked files
                    continue
        except Exception:
            continue

def run_automation(playwright: Playwright, q: queue.Queue):
    try:
        q.put(("status", "Carregando credenciais..."))
        q.put(("progress", 5))
        credentials = load_credentials()
        url, username, password = credentials['url'], credentials['user'], credentials['password']

        q.put(("status", "Iniciando navegador..."))

        # real Chrome profile (Default folder). No trailing space.
        real_profile_root = r"C:\Users\perna\AppData\Local\Google\Chrome\User Data\Default"
        if not os.path.exists(real_profile_root):
            # Try without Default subfolder (some setups use "User Data" folder directly)
            real_profile_root = r"C:\Users\perna\AppData\Local\Google\Chrome\User Data"
            if not os.path.exists(real_profile_root):
                raise FileNotFoundError(f"Chrome profile not found at expected locations.")

        # Create a lightweight temp profile and copy only useful files (safe while Chrome is open)
        temp_profile = tempfile.mkdtemp(prefix="chrome_profile_")
        dst_profile = os.path.join(temp_profile, "Default")
        os.makedirs(dst_profile, exist_ok=True)

        q.put(("status", "Copiando arquivos do perfil (somente essenciais)..."))
        safe_copy(real_profile_root, dst_profile, PROFILE_ITEMS)

        # Launch Chromium with the copied profile (uses Chrome user-data content)
        # Note: Using playwright.chromium.launch_persistent_context on the browser-type
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=temp_profile,
            headless=False,
            args=[
                "--start-minimized",
                "--disable-blink-features=AutomationControlled",
                "--disable-infobars",
                "--no-sandbox",
                "--disable-dev-shm-usage",
            ],
        )

        page = context.new_page()
        # small wait so profile files are loaded
        time.sleep(1)

        Login_and_Navigation(page, url, q, username, password)

    except Exception as e:
        q.put(("status", f"Ocorreu um erro inesperado: {e}"))
        # If you want extra debugging:
        # import traceback; q.put(("status", traceback.format_exc()))
    finally:
        q.put(("status", "Fechando navegador..."))
        if 'context' in locals():
            try:
                context.close()
            except Exception:
                pass
        # cleanup the temp profile folder
        if 'temp_profile' in locals() and os.path.exists(temp_profile):
            try:
                shutil.rmtree(temp_profile, ignore_errors=True)
            except Exception:
                pass
        q.put(("done", True))

def main_process(q: queue.Queue):
    with sync_playwright() as playwright:
        # for i in range(0,3):
        run_automation(playwright, q)

# --- TKINTER APP SETUP ---
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Ferramenta de Automação e Processamento")
        self.root.geometry("600x400")

        self.queue = queue.Queue()

        # --- Widgets ---
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.status_label = ttk.Label(main_frame, text="Pronto para iniciar. Clique em 'Processar'.", font=("Helvetica", 12))
        self.status_label.pack(pady=5, padx=5, fill=tk.X)

        self.progress_bar = ttk.Progressbar(main_frame, orient='horizontal', length=400, mode='determinate')
        self.progress_bar.pack(pady=10, padx=5, fill=tk.X)

        self.process_button = ttk.Button(main_frame, text="Processar", command=self.start_processing_thread)
        self.process_button.pack(pady=10)
        
        log_frame = ttk.LabelFrame(main_frame, text="Log de Atividades", padding="10")
        log_frame.pack(pady=10, padx=5, fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, width=70, height=15)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def start_processing_thread(self):
        self.process_button.config(state="disabled")
        self.progress_bar['value'] = 0
        self.log_text.delete('1.0', tk.END)
        self.status_label.config(text="Iniciando processo...")
        
        self.thread = threading.Thread(target=main_process, args=(self.queue,))
        self.thread.daemon = True
        self.thread.start()
        
        # Start checking the queue for updates
        update_gui(self.queue, self.status_label, self.progress_bar, self.log_text)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
