import json
import os
import sys
import subprocess
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

import time
from datetime import date, timedelta


warnings.filterwarnings("ignore", category=UserWarning)

import requests
from playwright.sync_api import Page, TimeoutError

from playwright.sync_api import Page, TimeoutError
import time

import random

from playwright_stealth.stealth import stealth_sync



def human_like_delay(min_delay=0.1, max_delay=0.6):
    time.sleep(random.uniform(min_delay, max_delay))

def Login_and_Navigation(page: Page, url, q, username, password):
    try:
        

        q.put(("status", "Navigating to login page..."))
        page.goto(url, timeout=60000)
        page.wait_for_load_state("domcontentloaded")

        q.put(("progress", 2))
        q.put(("status", "Performing login..."))

        # Simulate human-like typing
        page.get_by_role("textbox", name="E-mail ou telefone").click()
        for char in username:
            page.keyboard.insert_text(char)
            human_like_delay(0.02, 0.07)

        human_like_delay(0.1, 0.2)
        page.get_by_role("textbox", name="Senha").click()
        for char in password:
            page.keyboard.insert_text(char)
            human_like_delay(0.05, 0.07)

        human_like_delay(0.1, 0.2)

        # --- Handle Cloudflare Turnstile if present ---
        q.put(("status", "Checking for Cloudflare verification..."))
        from playwright.sync_api import TimeoutError

        # Set a specific, reasonable timeout for this operation, e.g., 5 seconds (5000 ms)
        WAIT_TIMEOUT_MS = 5000 

        try:
            # 1. Create the locator
            success_locator = page.locator('span#success-text')

            # 2. Tell the locator to wait until it is visible (auto-waits up to WAIT_TIMEOUT_MS)
            success_locator.wait_for(state='visible', timeout=WAIT_TIMEOUT_MS)

            print("Success found")

        except TimeoutError:
            # This exception is raised only if the element doesn't appear within the timeout
            q.put(("status", "Success element not detected within timeout. Continuing..."))
            
        except Exception as e:
            # Handle any other unexpected errors during the wait
            q.put(("status", f"Error while checking for success element: {e}"))
            

        # --- Human-like activity before clicking login ---
        page.mouse.wheel(0, 200)
        human_like_delay(0.3, 0.8)
        page.mouse.move(random.randint(100, 300), random.randint(400, 500))

        # --- Click Login Button ---
        q.put(("status", "Submitting login..."))
        page.get_by_role("button", name="Entrar").hover()
        human_like_delay(0.1, 0.5)
        page.get_by_role("button", name="Entrar").click()
        

        # # --- Wait for successful login indicator ---
        try:
            page.get_by_role("button", name="Gestão de Pedidos Gestão da")
            q.put(("status", "✅ Login successful"))
            q.put(("progress", 5))
        except TimeoutError:
            q.put(("status", "⚠️ Login attempt failed "))
            page.screenshot(path="login_failed.png")

        page.wait_for_timeout(100000)

        page.pause()

    except Exception as e:
        q.put(("status", f"❌ Error during login: {e}"))
