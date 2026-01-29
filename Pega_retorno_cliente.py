
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



from RestAPIHelper import main, update_excel_rows, update_protocol_async, update_agenda_async

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
        
        
        df = df[df['Nº Pedido Cliente'].notna()].copy()
        df['chave_pedido_loja'] = df['Nº Pedido Cliente'].astype(str) + '-' + df['CÓD LOJA'].astype(str).str.split('-').str[0]

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
        
        q.put(("progress", 10))
       
        try:
            page.get_by_role("button", name="Demandas Gestão e controle de")
            q.put(("status", "✅ Login realizado com sucesso"))
            q.put(("progress", 5))
        except TimeoutError:
            q.put(("status", "⚠️ Tentativa de login falhou "))
            page.screenshot(path="login_failed.png")
      
        page.get_by_role("button", name="Demandas Gestão e controle de").click()
       

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
        df=df[df["Nº Pedido Cliente"].notna()].copy()
        
        # grouped_by_carro = loja_df.groupby('CARRO')
        unique_groups = df.groupby('CARRO').first().reset_index()
        
        q.put(("status", f"Encontrados {len(unique_groups)} grupos únicos para processar"))
        
        agenda_data_list = []
        not_found_protocols = []
        
        total_groups = len(unique_groups)
        
        for idx, row in unique_groups.iterrows():
            try:
                chave = row['chave_pedido_loja']
                
                # Check if agenda columns are already filled for this group
                agenda_cols = ['AGENDA CONFIRMADA', 'PROTOCOLO AGENDA', 'HORÁRIO']
                group_filled = True
                
                # Get all rows for this chave to check if agenda is filled
                group_rows = df[df['chave_pedido_loja'] == chave]
                for _, group_row in group_rows.iterrows():
                    for col in agenda_cols:
                        if pd.isna(group_row.get(col)) or str(group_row.get(col)).strip() == '':
                            group_filled = False
                            break
                    if not group_filled:
                        break
                
                if group_filled:
                    q.put(("status", f"Grupo {chave} já processado (agenda preenchida). Pulando."))
                    continue  # Skip to next group
                
                protocol = str(row.get('PROTOCOLO DA SOLICITAÇÃO', '')).strip()
                
                # Skip if protocol is empty, None, NaN, or contains error messages
                if not protocol or protocol.lower() in ['nan', 'none', ''] or 'erro' in protocol.lower():
                    q.put(("status", f"⚠️ Grupo {chave} não possui protocolo válido"))
                    continue
                
                q.put(("status", f"Buscando resposta para protocolo {protocol} ({idx+1}/{total_groups})..."))
                progress_value = 15 + int((idx / total_groups) * 70)
                q.put(("progress", progress_value))
                
                page.get_by_role("textbox", name="Buscar demandas...").fill(protocol)
                               
                page.get_by_role("link", name=f"Demanda #{protocol}").click(timeout=4000)
                page.get_by_role("button", name="Histórico").click()
                
                page.wait_for_timeout(2000)  # Wait for history to load
                
                
                page.pause()
                # Find the status in the gridcell
                status_element = page.get_by_role("gridcell").filter(has_text=re.compile(r"^Recebimento - ")).first
                if status_element.is_visible():
                    status_text = status_element.text_content()
                    
                    if "Recebimento - Aprovada" in status_text and "No-show" not in status_text:
                        # Proceed to extract details
                        page.locator(".history-item-container").first.click()
                        human_like_delay(0.2, 0.5)
                        
                        # Extract demand number and date/time
                        try:
                            # Define a locator for the visible details panel.
                            details_panel = page.locator("div[id^='panel-']:has-text('Data efetiva entrega'):visible")
                            
                            # Ensure panel is ready before proceeding
                            details_panel.wait_for(timeout=5000)

                            # --- Extract demand number (scoped to the panel) ---
                            demand_number = None
                            demand_candidates = details_panel.locator("text=/Agendamento\\d{5,}/").all_text_contents()
                            if demand_candidates:
                                demand_number = demand_candidates[0]
                            else:
                                # Fallback
                                demand_candidates = details_panel.locator("text=Agendamento").all_text_contents()
                                if demand_candidates:
                                    demand_number = demand_candidates[0]

                            # --- Extract date and time (from parent container) ---
                            date_text = None
                            date_time = None
                            datetime_pattern = re.compile(r"(\d{2}/\d{2}/\d{4})[\s\S]*?(\d{2}:\d{2})")
                            
                            # Find the parent container of the "Data efetiva entrega" label.
                            date_container = details_panel.locator("text=/Data efetiva entrega.*/").first.locator("..")
                            
                            # Get text from that container and parse it
                            full_date_text = date_container.text_content()
                            
                            date_text_parts = full_date_text.split('\n')
                            if date_text_parts:
                                date_text = date_text_parts[0].strip()

                            datetime_match = datetime_pattern.search(full_date_text)
                            if datetime_match:
                                date_value = datetime_match.group(1)
                                time_value = datetime_match.group(2)
                                date_time = f"{date_value} {time_value}"
                                
                                print("date Value here : ",date_value,"Time Value : ", time_value)
                            else:
                                date_value = ""
                                time_value = ""
                        except Exception as e:
                            demand_number = ""
                            date_value = ""
                            time_value = ""
                            q.put(("status", f"⚠️ Falha ao extrair detalhes para protocolo {protocol}: {e}"))

                        # Parse demand_number to extract only the number part (remove "Agendamento" prefix)
                        if demand_number and demand_number.startswith("Agendamento"):
                            demand_number = demand_number.replace("Agendamento", "").strip()

                        # Only append if all required values are present and not empty
                        if date_value and demand_number and time_value:
                            agenda_data_list.append({
                                "chave": chave,
                                "protocol": protocol,
                                "agenda_confirmada": date_value,
                                "protocolo_agenda": demand_number,
                                "horario": time_value
                            })

                            q.put(("status", f"✅ Dados extraídos para {chave}: Agenda={date_value}, Protocolo Agenda={demand_number}, Horário={time_value}"))
                        else:
                            q.put(("status", f"⚠️ Dados incompletos para {chave}, pulando atualização"))
                            not_found_protocols.append(chave)
                        
                    elif "Recebimento - Remanejamento" in status_text  and "No-show" not in status_text:
                        # Proceed to extract details for remanejamento
                        page.locator(".history-item-container").nth(1).click()
                        human_like_delay(0.2, 0.5)
                        
                        # Extract date and time from "Data sugerida entrega"
                        try:
                            # Define a locator for the visible details panel.
                            details_panel = page.locator("div[id^='panel-']:has-text('Data sugerida entrega'):visible")
                            
                            # Ensure panel is ready before proceeding
                            details_panel.wait_for(timeout=5000)

                            # --- Extract date and time (from parent container) ---
                            datetime_pattern = re.compile(r"(\d{2}/\d{2}/\d{4})[\s\S]*?(\d{2}:\d{2})")
                            
                            # Find the parent container of the "Data sugerida entrega" label.
                            date_container = details_panel.locator("text=/Data sugerida entrega.*/").first.locator("..")
                            
                            # Get text from that container and parse it
                            full_date_text = date_container.text_content()
                            
                            datetime_match = datetime_pattern.search(full_date_text)
                            if datetime_match:
                                date_value = datetime_match.group(1)
                                time_value = datetime_match.group(2)
                                
                            else:
                                date_value = ""
                                time_value = ""
                        except Exception as e:
                            date_value = ""
                            time_value = ""
                            q.put(("status", f"⚠️ Falha ao extrair detalhes de remanejamento para protocolo {protocol}: {e}"))

                        # Only append if date and time are present
                        if date_value and time_value:
                            agenda_data_list.append({
                                "chave": chave,
                                "protocol": protocol,  # Use the protocol number for matching
                                "agenda_confirmada": date_value,
                                "protocolo_agenda": "",  # Empty as requested
                                "horario": time_value
                            })

                            q.put(("status", f"✅ Dados extraídos para remanejamento {chave}: Agenda={date_value}, Horário={time_value}"))
                        else:
                            q.put(("status", f"⚠️ Dados incompletos para remanejamento {chave}, pulando atualização"))
                            not_found_protocols.append(chave)
                        
                    else:
                        # For other statuses (canceled, pendente, etc.), use the status text
                        agenda_data_list.append({
                            "chave": chave,
                            "protocol": protocol,  # Use the protocol number for matching
                            "agenda_confirmada": status_text,
                            "protocolo_agenda": status_text,
                            "horario": status_text
                        })
                        
                        q.put(("status", f"✅ Protocolo {protocol} status: {status_text}"))
                    
                    # Go back to search page
                    page.locator("section").get_by_role("button").filter(has_text=re.compile(r"^$")).click()
                    
                else:
                    q.put(("status", f"⚠️ Nenhum status 'Recebimento - ' encontrado para protocolo {protocol}. Pulando..."))
                    # Go back to search page
                    page.locator("section").get_by_role("button").filter(has_text=re.compile(r"^$")).click()
                    continue
                
                human_like_delay(0.3, 0.8)  # Human-like delay between searches
               
                
            except Exception as e:
                q.put(("status", f"❌ Erro ao processar {chave}: {e}"))
                not_found_protocols.append(chave)
                # Go back to search page
                page.locator("section").get_by_role("button").filter(has_text=re.compile(r"^$")).click()
                continue
        
        # Update SharePoint with collected agenda data
        if agenda_data_list:
            q.put(("status", f"Atualizando {len(agenda_data_list)} dados de agenda no SharePoint..."))
            q.put(("progress", 85))
            
            # Run async update in thread
            update_queue = queue.Queue()
            update_thread = threading.Thread(
                target=azure_update_agenda_in_thread,
                args=(drive_id, file_id, agenda_data_list, update_queue)
            )
            update_thread.start()
            update_thread.join()
            
            try:
                result = update_queue.get_nowait()
                if isinstance(result, Exception):
                    q.put(("status", f"❌ Erro ao atualizar dados de agenda: {result}"))
                else:
                    q.put(("status", "✅ Dados de agenda atualizados com sucesso no SharePoint"))
            except queue.Empty:
                q.put(("status", "⚠️ Atualização de agenda sem retorno"))
        
        q.put(("progress", 95))
        
        if not_found_protocols:
            q.put(("status", f"⚠️ {len(not_found_protocols)} protocolos não processados"))
        
        q.put(("status", "✅ Processo de busca de retornos concluído!"))
        q.put(("progress", 100))
        
    except Exception as e:
        q.put(("status", f"❌ Erro durante processamento de protocolos: {e}"))
        import traceback
        traceback.print_exc()


def azure_update_agenda_in_thread(drive_id: str, file_id: str, agenda_data_list: list, result_queue: queue.Queue):
    """
    Helper function to run async update_agenda_async in a thread.
    """
    try:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        
        loop.run_until_complete(
            update_agenda_async(drive_id, file_id, agenda_data_list)
        )
        
        result_queue.put("success")
    except Exception as e:
        result_queue.put(e)
    finally:
        loop.close()


def run_retorno_automation(playwright: Playwright, q: queue.Queue):
    
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
