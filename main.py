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

from Tasks import Login_and_Navigation ,Order_datas_from_sharepoint
from Pega_retorno_cliente import run_retorno_automation


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
def update_gui(queue_instance, status_label, progress_bar, log_text, process_button=None):
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
                status_label.config(text="Processo Concluído! Clique em 'Processar' para executar novamente.")
                progress_bar['value'] = 100
                if process_button:
                    process_button.config(state="normal")  # Re-enable the button
                return # Stop checking
    except queue.Empty:
        pass
    status_label.after(100, lambda: update_gui(queue_instance, status_label, progress_bar, log_text, process_button))


def load_credentials():
    """Loads Credencial.json from the same directory as the running script or executable."""
    base_path = os.path.dirname(os.path.abspath(sys.argv[0]))
    cred_path = os.path.join(base_path, "credencial.json")

    if not os.path.exists(cred_path):
        raise FileNotFoundError(f"Credencial.json not found in: {cred_path}")

    with open(cred_path, "r", encoding="utf-8") as f:
        return json.load(f)

def run_automation(playwright: Playwright, q: queue.Queue):
    try:
        q.put(("status", "Carregando credenciais..."))
        q.put(("progress", 3))
        credentials = load_credentials()
        url, username, password = credentials['url'], credentials['user'], credentials['password']

        q.put(("status", "Iniciando navegador..."))
        q.put(("progress", 5))

        # Create fresh temporary profile (works better with Cloudflare protection)
        temp_profile = tempfile.mkdtemp(prefix="chrome_profile_")
        q.put(("status", "Criando perfil temporário limpo..."))
        q.put(("progress", 8))
        
        
        chromium_path = get_playwright_browser_path()
        
        if chromium_path:
            browser = playwright.chromium.launch(
            headless=False,
            executable_path=chromium_path,
            args=[
                "--start-maximized",
                "--disable-blink-features=AutomationControlled",
                "--disable-infobars",
                "--no-sandbox",
                "--disable-dev-shm-usage",
            ]
                )
            
        else : 
            browser = playwright.chromium.launch(
                headless=False,
                args=[
                    "--start-maximized",
                    "--disable-blink-features=AutomationControlled",
                    "--disable-infobars",
                    "--no-sandbox",
                    "--disable-dev-shm-usage",
                ],
            )
        
        context = browser.new_context(
            no_viewport=True,  # Use full browser window size
        )
        
        
        
        page = context.new_page()
        time.sleep(1)


        # Order_datas_from_sharepoint(q)
        Login_and_Navigation(page, url, q, username, password)

    except Exception as e:
        q.put(("status", f"Ocorreu um erro inesperado: {e}"))
        # If you want extra debugging:
        # import traceback; q.put(("status", traceback.format_exc()))
    finally:
        q.put(("status", "Fechando navegador..."))
        if 'browser' in locals():
            try:
                browser.close()
            except Exception:
                pass
        # cleanup the temp profile folder (if used)
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


def retorno_process(q: queue.Queue):
    """Process for Pegar Retorno - fetches client responses and updates SharePoint"""
    with sync_playwright() as playwright:
        run_retorno_automation(playwright, q)



# --- TKINTER APP SETUP ---
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Ferramenta de Automação e Processamento RPA")
        self.root.geometry("700x550")
        self.root.resizable(True, True)
        
        # DHL & STELLANTIS Colors
        # DHL: Red (#FF0000), Yellow (#FFCC00)
        # STELLANTIS: Blue (#003DA5), Orange (#FF6600)
        dhl_red = "#FF0000"
        dhl_yellow = "#FFCC00"
        stellantis_blue = "#003DA5"
        stellantis_orange = "#FF6600"
        
        # Set modern color scheme with DHL/STELLANTIS theme
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure button style with STELLANTIS blue
        style.configure('TButton', background=stellantis_blue, foreground="white", relief="flat", padding=6)
        style.map('TButton', background=[('active', stellantis_orange)])
        
        # Configure progressbar with gradient effect using STELLANTIS colors
        style.configure('TProgressbar', background=stellantis_blue, troughcolor='#E8E8E8', bordercolor='#CCCCCC', lightcolor=stellantis_orange, darkcolor=stellantis_blue)
        
        # Configure labels with theme colors
        style.configure('Title.TLabel', font=("Segoe UI", 16, "bold"), foreground=stellantis_blue)
        
        self.queue = queue.Queue()

        # --- Main container ---
        container = tk.Frame(root, bg="white")
        container.pack(fill=tk.BOTH, expand=True)

        # --- Header with DHL/STELLANTIS accent ---
        header_frame = tk.Frame(container, bg=stellantis_blue, height=80)
        header_frame.pack(fill=tk.X, padx=0, pady=0)
        header_frame.pack_propagate(False)
        
        # Title section with colored background
        title_label = tk.Label(header_frame, text="🤖 Automação de Processo ..", font=("Segoe UI", 16, "bold"), fg="white", bg=stellantis_blue)
        title_label.pack(anchor="w", padx=15, pady=(10, 2))
        
        subtitle_label = tk.Label(header_frame, text="Processamento Inteligente de Processos Manuais", font=("Segoe UI", 9), fg=dhl_yellow, bg=stellantis_blue)
        subtitle_label.pack(anchor="w", padx=15, pady=(0, 10))

        # --- Main content frame ---
        main_frame = ttk.Frame(container, padding="13")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Status section
        self.status_label = ttk.Label(main_frame, text="Pronto para iniciar. Clique em 'Processar'.", font=("Segoe UI", 11), foreground=stellantis_blue)
        self.status_label.pack(pady=(2, 5), padx=1, fill=tk.X)

        # Progress bar with accent color
        self.progress_bar = ttk.Progressbar(main_frame, orient='horizontal', length=400, mode='determinate')
        self.progress_bar.pack(pady=10, padx=5, fill=tk.X)

        # Button section with modern styling
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=4, fill=tk.X)
        
        self.process_button = ttk.Button(button_frame, text="▶ Processar", command=self.start_processing_thread, style='TButton')
        self.process_button.pack(side=tk.LEFT, padx=5)
        
        self.retorno_button = ttk.Button(button_frame, text="📥 Pegar Retorno", command=self.start_retorno_thread, style='TButton')
        self.retorno_button.pack(side=tk.RIGHT, padx=5)
        
        # Log section with accent
        log_frame = ttk.LabelFrame(main_frame, text="📋 Log de Atividades", padding="13")
        log_frame.pack(pady=0, padx=2, fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, width=80, height=10, font=("Consolas", 11), bg="#F5F5F5", fg="#333333")
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Footer section with DHL/STELLANTIS branding
        footer_frame = tk.Frame(container, bg=stellantis_blue, height=34)
        footer_frame.pack(fill=tk.X, padx=0, pady=0, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        # Left side - DHL -> STELLANTIS
        left_footer = tk.Frame(footer_frame, bg=stellantis_blue)
        left_footer.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=15, pady=10)
        
        # DHL Logo/Text (DHL Yellow)
        dhl_label = tk.Label(left_footer, text="🚚 DHL", font=("Segoe UI", 11, "bold"), fg=dhl_yellow, bg=stellantis_blue)
        dhl_label.pack(side=tk.LEFT, padx=1)
        
        arrow_label = tk.Label(left_footer, text="→", font=("Segoe UI", 12, "bold"), fg=dhl_yellow, bg=stellantis_blue)
        arrow_label.pack(side=tk.LEFT, padx=3)
        
        # STELLANTIS Logo/Text (STELLANTIS Orange accent)
        stellantis_label = tk.Label(left_footer, text="ALPARGATAS 🏢", font=("Segoe UI", 11, "bold"), fg=stellantis_orange, bg=stellantis_blue)
        stellantis_label.pack(side=tk.LEFT, padx=3)
        
        # Right side - Developer credit
        right_footer = tk.Frame(footer_frame, bg=stellantis_blue)
        right_footer.pack(side=tk.RIGHT, padx=15, pady=10)
        
        footer_label = tk.Label(right_footer, text="Desenvolvido por: Vincent Pernarh", font=("Segoe UI", 9), fg="white", bg=stellantis_blue)
        footer_label.pack(anchor="e")

    def start_processing_thread(self):
        self.process_button.config(state="disabled")
        self.progress_bar['value'] = 0
        self.log_text.delete('1.0', tk.END)
        self.status_label.config(text="Iniciando processo...")
        
        self.thread = threading.Thread(target=main_process, args=(self.queue,))
        self.thread.daemon = True
        self.thread.start()
        
        # Start checking the queue for updates (pass button reference to re-enable it)
        update_gui(self.queue, self.status_label, self.progress_bar, self.log_text, self.process_button)
    
    def start_retorno_thread(self):
        """Start the Pegar Retorno process in a separate thread"""
        self.retorno_button.config(state="disabled")
        self.progress_bar['value'] = 0
        self.log_text.delete('1.0', tk.END)
        self.status_label.config(text="Iniciando processo de Pegar Retorno...")
        
        self.retorno_thread = threading.Thread(target=retorno_process, args=(self.queue,))
        self.retorno_thread.daemon = True
        self.retorno_thread.start()
        
        # Start checking the queue for updates (pass button reference to re-enable it)
        update_gui(self.queue, self.status_label, self.progress_bar, self.log_text, self.retorno_button)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
