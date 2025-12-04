
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
import xlwings as xw
import re

import time
from datetime import date, timedelta

import asyncio
import os
import pandas as pd
from dotenv import load_dotenv
import aiohttp
import datetime as dt
import requests
from playwright.sync_api import Page, TimeoutError

from playwright.sync_api import Page, TimeoutError
import time

import random

from azure.identity.aio import ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from msgraph.generated.sites.item.drives.drives_request_builder import DrivesRequestBuilder
from msgraph.generated.drives.item.items.item.workbook.worksheets.item.used_range.used_range_request_builder import UsedRangeRequestBuilder



from Azure_Access import main, update_excel_rows, update_protocol_async, update_response_async

base_path = os.getcwd()

warnings.filterwarnings("ignore", category=UserWarning)


def load_static_data():
    """Load static data from static_data.json"""
    static_data_path = os.path.join(base_path, "static_data.json")
    try:
        with open(static_data_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        raise FileNotFoundError(f"static_data.json not found at {static_data_path}")
    except json.JSONDecodeError:
        raise ValueError("static_data.json is not valid JSON")



def human_like_delay(min_delay=0.1, max_delay=0.6):
    time.sleep(random.uniform(min_delay, max_delay))




# +++++++++ HELPER FUNCTION TO RUN ASYNC IN A THREAD +++++++++
def azure_main_in_thread(result_queue: queue.Queue):
   
    try:
        # Create and set a new event loop for this thread
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        
        result = loop.run_until_complete(main()) 
        
        # Put the successful result into the queue
        result_queue.put(result)
    except Exception as e:
        # If anything goes wrong, put the exception in the queue
        result_queue.put(e)
    finally:
        # Clean up the loop
        loop.close()



def Order_datas_from_sharepoint(q):
    
    q.put(("status", "Obtendo dados do Azure..."))
    result_queue = queue.Queue() # A new queue just for this thread's result
    
    # Create and start the thread, targeting our new helper function
    azure_thread = threading.Thread(target=azure_main_in_thread, args=(result_queue,))
    azure_thread.start()
    
    # Wait for the thread to finish its work
    azure_thread.join() # <--- You are here. The thread is finished.
                     
    try:
        # 1. Get the item from the queue
        result = result_queue.get_nowait() 
        if isinstance(result, Exception):
            # If the thread sent back an error, handle it
            q.put(("status", f"❌ Erro ao obter dados do Azure: {result}"))
            raise result # Re-raise the error
        
        df, drive_id, file_id = result
        # df['chave_pedido_loja'] =df['Nº Pedido Cliente'].astype(str) + '-' + df['CÓD LOJA'].astype(str).str.split('-').str[0]
        
        df['chave_pedido_loja'] =df['Nº Pedido Cliente'].astype(str) + '-' + df['CÓD LOJA'].astype(str).str.split('-').str[0]

        q.put(("status", "✅ Dados do Azure obtidos com sucesso."))
        
        return df, drive_id, file_id

    except queue.Empty:
        # This shouldn't happen if join() worked, but it's safe to have
        q.put(("status", "❌ Thread do Azure finalizou sem resultado."))
        return None, None, None
    except Exception as e:
        # Handle any other error
        q.put(("status", f"❌ Falha ao processar resultado do Azure: {e}"))
        return None, None, None





def Login_and_Navigation(page: Page, url, q, username, password):
    
    try:
        q.put(("status", "Navegando para página de login..."))
        page.goto(url, timeout=60000)
        page.wait_for_load_state("domcontentloaded")

        q.put(("progress", 2))
        q.put(("status", "Realizando login..."))

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

        q.put(("status", "Verificando autenticação Cloudflare..."))
        from playwright.sync_api import TimeoutError

        # Set a specific, reasonable timeout for this operation, e.g., 5 seconds (5000 ms)
        WAIT_TIMEOUT_MS = 5000 

        try:
            # 1. Create the locator
            success_locator = page.locator('span#success-text')
            success_locator.wait_for(state='visible', timeout=WAIT_TIMEOUT_MS)

            print("Success found")

        except TimeoutError:
            # This exception is raised only if the element doesn't appear within the timeout
            q.put(("status", "Elemento de sucesso não detectado no tempo limite. Continuando..."))
            
        except Exception as e:
            # Handle any other unexpected errors during the wait
            q.put(("status", f"Erro ao verificar elemento de sucesso: {e}"))
            
        # --- Human-like activity before clicking login ---
        page.mouse.wheel(0, 200)
        human_like_delay(0.3, 0.8)
        page.mouse.move(random.randint(100, 300), random.randint(400, 500))

        # --- Click Login Button ---
        q.put(("status", "Enviando login..."))
        page.get_by_role("button", name="Entrar").hover()
        human_like_delay(0.1, 0.5)
        page.get_by_role("button", name="Entrar").click()
        
        q.put(("status", "✅ Login realizado com sucesso"))
        q.put(("progress", 10))
        
        # Pause here for manual navigation definition
        q.put(("status", "⏸️ Pausado para definição manual de navegação..."))
        page.pause()

    except Exception as e:
        q.put(("status", f"❌ Erro durante o login: {e}"))


def process_protocol_responses(page: Page, df, drive_id, file_id, q):
    """
    Extract unique protocols by group, search web for each,
    collect responses, and update SharePoint Excel.
    """
    try:
        q.put(("status", "Processando protocolos únicos..."))
        q.put(("progress", 15))
        
        # Get unique protocols grouped by chave_pedido_loja
        unique_groups = df.groupby('chave_pedido_loja').first().reset_index()
        
        q.put(("status", f"Encontrados {len(unique_groups)} grupos únicos para processar"))
        
        response_data_list = []
        not_found_protocols = []
        
        total_groups = len(unique_groups)
        
        for idx, row in unique_groups.iterrows():
            try:
                chave = row['chave_pedido_loja']
                protocol = str(row.get('PROTOCOLO DA SOLICITAÇÃO', '')).strip()
                
                # Skip if protocol is empty, None, NaN, or contains error messages
                if not protocol or protocol.lower() in ['nan', 'none', ''] or 'erro' in protocol.lower():
                    q.put(("status", f"⚠️ Grupo {chave} não possui protocolo válido"))
                    continue
                
                q.put(("status", f"Buscando resposta para protocolo {protocol} ({idx+1}/{total_groups})..."))
                progress_value = 15 + int((idx / total_groups) * 70)
                q.put(("progress", progress_value))
                
                # ============================================
                # TODO: Add web search logic here
                # Example structure:
                # 1. Navigate to search page (if needed)
                # 2. Fill search input with protocol
                # 3. Submit search
                # 4. Extract response text from results
                # ============================================
                
                # Placeholder - will be implemented after page.pause()
                response_text = "Pendente Implementação"  # Replace with actual search result
                
                response_data_list.append({
                    "chave": chave,
                    "protocol": protocol,
                    "response": response_text
                })
                
                q.put(("status", f"✅ Resposta obtida para {chave}: {response_text}"))
                human_like_delay(0.3, 0.8)  # Human-like delay between searches
                
            except Exception as e:
                q.put(("status", f"❌ Erro ao processar {chave}: {e}"))
                not_found_protocols.append(chave)
                continue
        
        # Update SharePoint with collected responses
        if response_data_list:
            q.put(("status", f"Atualizando {len(response_data_list)} respostas no SharePoint..."))
            q.put(("progress", 85))
            
            # Run async update in thread
            update_queue = queue.Queue()
            update_thread = threading.Thread(
                target=azure_update_response_in_thread,
                args=(drive_id, file_id, response_data_list, update_queue)
            )
            update_thread.start()
            update_thread.join()
            
            try:
                result = update_queue.get_nowait()
                if isinstance(result, Exception):
                    q.put(("status", f"❌ Erro ao atualizar respostas: {result}"))
                else:
                    q.put(("status", "✅ Respostas atualizadas com sucesso no SharePoint"))
            except queue.Empty:
                q.put(("status", "⚠️ Atualização de respostas sem retorno"))
        
        q.put(("progress", 95))
        
        if not_found_protocols:
            q.put(("status", f"⚠️ {len(not_found_protocols)} protocolos não processados"))
        
        q.put(("status", "✅ Processo de busca de retornos concluído!"))
        q.put(("progress", 100))
        
    except Exception as e:
        q.put(("status", f"❌ Erro durante processamento de protocolos: {e}"))
        import traceback
        traceback.print_exc()


def azure_update_response_in_thread(drive_id: str, file_id: str, response_data_list: list, result_queue: queue.Queue):
    """
    Helper function to run async update_response_async in a thread.
    """
    try:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        
        loop.run_until_complete(
            update_response_async(drive_id, file_id, response_data_list)
        )
        
        result_queue.put("success")
    except Exception as e:
        result_queue.put(e)
    finally:
        loop.close()


def run_retorno_automation(playwright: Playwright, q: queue.Queue):
    """
    Main automation function for Pegar Retorno process.
    Orchestrates login, data fetch, and protocol response collection.
    """
    try:
        q.put(("status", "Carregando credenciais..."))
        q.put(("progress", 1))
        
        base_path = os.path.dirname(os.path.abspath(sys.argv[0]))
        cred_path = os.path.join(base_path, "credencial.json")
        
        if not os.path.exists(cred_path):
            raise FileNotFoundError(f"Credencial.json não encontrado em: {cred_path}")
        
        with open(cred_path, "r", encoding="utf-8") as f:
            credentials = json.load(f)
        
        url = credentials['url']
        username = credentials['user']
        password = credentials['password']
        
        q.put(("status", "Iniciando navegador..."))
        q.put(("progress", 2))
        
        # Get Chromium path
        if getattr(sys, 'frozen', False):
            base_path_browser = sys._MEIPASS
            chromium_path = os.path.join(base_path_browser, "ms-playwright", "chromium-1187", "chrome-win", "chrome.exe")
        else:
            base_path_browser = r"C:\Users\perna\AppData\Local"
            chromium_path = os.path.join(
                base_path_browser,
                "ms-playwright",
                "chromium-1187",
                "chrome-win",
                "chrome.exe"
            )
        
        if chromium_path and os.path.exists(chromium_path):
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
        else:
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
        
        context = browser.new_context(no_viewport=True)
        page = context.new_page()
        time.sleep(1)
        
        # Step 1: Login
        Login_and_Navigation(page, url, q, username, password)
        
        # Step 2: Get data from SharePoint
        df, drive_id, file_id = Order_datas_from_sharepoint(q)
        
        if df is None or df.empty:
            q.put(("status", "❌ Nenhum dado obtido do SharePoint"))
            return
        
        q.put(("status", f"✅ {len(df)} registros carregados do SharePoint"))
        
        # Step 3: Process protocols and fetch responses
        process_protocol_responses(page, df, drive_id, file_id, q)
        
    except Exception as e:
        q.put(("status", f"❌ Erro inesperado: {e}"))
        import traceback
        traceback.print_exc()
    finally:
        q.put(("status", "Fechando navegador..."))
        if 'browser' in locals():
            try:
                browser.close()
            except Exception:
                pass
        q.put(("done", True))
