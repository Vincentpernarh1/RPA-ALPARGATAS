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

import time
from datetime import date, timedelta

import asyncio
import os
import pandas as pd
from dotenv import load_dotenv
import aiohttp
import datetime as dt

from azure.identity.aio import ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from msgraph.generated.sites.item.drives.drives_request_builder import DrivesRequestBuilder
from msgraph.generated.drives.item.items.item.workbook.worksheets.item.used_range.used_range_request_builder import UsedRangeRequestBuilder



from Azure_Access import main

base_path = os.getcwd()

warnings.filterwarnings("ignore", category=UserWarning)

import requests
from playwright.sync_api import Page, TimeoutError

from playwright.sync_api import Page, TimeoutError
import time

import random

from playwright_stealth.stealth import stealth_sync



def human_like_delay(min_delay=0.1, max_delay=0.6):
    time.sleep(random.uniform(min_delay, max_delay))


def load_filepath():
    CARTEURA_GRUPO_PATH = None
    CARTEURA_GRUPO_folder = os.path.join(base_path,"Dados")

    for file in os.listdir(CARTEURA_GRUPO_folder):
        
        if "CARTEIRA GRUPO" in file and file.endswith(".xlsx") and not file.startswith("~$"):
            CARTEURA_GRUPO_PATH = os.path.join(CARTEURA_GRUPO_folder, file)
            break

    return CARTEURA_GRUPO_PATH




# +++++++++ HELPER FUNCTION TO RUN ASYNC IN A THREAD +++++++++
def azure_main_in_thread(result_queue: queue.Queue):
   
    try:
        # Create and set a new event loop for this thread
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        
        # Run the async main function and wait for its result
        # The 'main()' function is the one imported from Azure_Access
        result = loop.run_until_complete(main()) 
        
        # Put the successful result into the queue
        result_queue.put(result)
    except Exception as e:
        # If anything goes wrong, put the exception in the queue
        result_queue.put(e)
    finally:
        # Clean up the loop
        loop.close()
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++



def Order_datas_from_sharepoint(q):
    
    q.put(("status", "Fetching Azure data..."))
    result_queue = queue.Queue() # A new queue just for this thread's result
    
    # Create and start the thread, targeting our new helper function
    azure_thread = threading.Thread(target=azure_main_in_thread, args=(result_queue,))
    azure_thread.start()
    print(result_queue) 
    
    # Wait for the thread to finish its work
    azure_thread.join() # <--- You are here. The thread is finished.
                     
    try:
        # 1. Get the item from the queue
        result = result_queue.get_nowait() 
        if isinstance(result, Exception):
            # If the thread sent back an error, handle it
            q.put(("status", f"❌ Error fetching Azure data: {result}"))
            raise result # Re-raise the error
        
        df, drive_id, file_id = result
        df['chave_pedido_loja'] =df['Nº Pedido Cliente'].astype(str) + '-' + df['CÓD LOJA'].astype(str).str.split('-').str[0]
        # print(df.head(3))
        
        q.put(("status", "✅ Azure data fetched successfully."))
        
        # 3. Return the values so the function that called 'teste_azure' can use them
        return df, drive_id, file_id

    except queue.Empty:
        # This shouldn't happen if join() worked, but it's safe to have
        q.put(("status", "❌ Azure thread finished with no result."))
        return None, None, None
    except Exception as e:
        # Handle any other error
        q.put(("status", f"❌ Failed to process Azure result: {e}"))
        return None, None, None


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

        q.put(("status", "Checking for Cloudflare verification..."))
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
        
        try:
            page.get_by_role("button", name="Gestão de Pedidos Gestão da")
            q.put(("status", "✅ Login successful"))
            q.put(("progress", 5))
        except TimeoutError:
            q.put(("status", "⚠️ Login attempt failed "))
            page.screenshot(path="login_failed.png")
        page.get_by_role("button", name="Gestão de Pedidos Gestão da").click()
        page.locator("#iframe-servico").content_frame.get_by_role("button", name="AÇÕES").click()
        page.locator("#iframe-servico").content_frame.get_by_role("menuitem", name="CONSUMIR ITENS").click()
       
        process_orders(page, q)
           
        page.wait_for_timeout(100000)

    except Exception as e:
        q.put(("status", f"❌ Error during login: {e}"))



def process_orders(page: Page, q):

    not_found_items = []
    try:
        df, drive_id, file_id = Order_datas_from_sharepoint(q)

        if df is None or df.empty:
            q.put(("status", "⚠️ No order data found to process."))
            return

        df['PRODUTO INTERNO CLIENTE'] = df['PRODUTO INTERNO CLIENTE'].astype(str)
        df['chave_pedido_loja'] = df['chave_pedido_loja'].astype(str)

        grouped_orders = df.groupby('chave_pedido_loja')
        total_groups = len(grouped_orders)
        q.put(("status", f"Found {len(df)} total items in {total_groups} unique groups."))
        
        frame_locator = page.locator("#iframe-servico").first.content_frame
        found_items = []  # <- put this BEFORE the loop
        # --- LOOP 1: By Unique 'chave_pedido_loja' ---
        for group_index, (chave, group_df) in enumerate(grouped_orders):
            
            q.put(("status", f"--- Processing Group {group_index + 1}/{total_groups}: {chave} ---"))
            
            try:
                chave_input = frame_locator.locator(".dx-texteditor-input").first
                chave_input.fill(chave)
                
                page.wait_for_timeout(1500) # 1.5 seconds
                # ++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                q.put(("status", "    -> Waiting for filter result..."))
                
                data_locator = frame_locator.locator(".dx-row.dx-data-row > td:nth-child(8)").first
                sem_dados_locator = frame_locator.get_by_text("Sem dados")

                try:
                    # 1. Check for data (with a shorter timeout, as the page is stable)
                    data_locator.wait_for(state="visible", timeout=3000)
                    q.put(("status", "    -> Data found. Proceeding to product loop."))

                except TimeoutError:
                    # 2. No data found. Check for "Sem dados"
                    if sem_dados_locator.is_visible():
                        q.put(("status", "    -> 'Sem dados' found for this group. Skipping."))
                        for _, row in group_df.iterrows():
                            not_found_items.append({
                                "chave": chave,
                                "produto": row['PRODUTO INTERNO CLIENTE'],
                                "motivo": "Chave principal não encontrada"
                            })
                        continue 
                    else:
                        q.put(("status", "    -> ERROR: No data row OR 'Sem dados' text found. Skipping group."))
                        continue 

                # --- LOOP 2: By 'PRODUTO INTERNO CLIENTE' for this group ---
                for _, row in group_df.iterrows():
                    produto_interno_cliente = row['PRODUTO INTERNO CLIENTE']
                    quantidade = row['Qtd. Item']
                    numero_cliente = row['Nº Pedido Cliente']
                
                    q.put(("status", f"    -> Filtering for product: {produto_interno_cliente}"))
                   
                    product_filter_input = frame_locator.locator("input[aria-label='Filtro de célula']").nth(3)
                                      
                    product_filter_input.fill("")
                    page.wait_for_timeout(300) # Short pause for clear
                    
                    product_filter_input.fill(produto_interno_cliente)
                    
                    page.wait_for_timeout(1000) # 1 second

                    q.put(("status", "-> Waiting for product filter result..."))
                    
                    try:
                        data_locator.wait_for(state="visible", timeout=3000)
                        q.put(("status", "    -> Product found."))
                        page.locator("#iframe-servico").content_frame.get_by_role("gridcell", name="Selecionar linha").get_by_role("checkbox").click()

                        found_items.append({
                                "numero_cliente": numero_cliente,
                                "produto_interno_cliente": produto_interno_cliente,
                                "quantidade": quantidade
                            })
                       
                        page.wait_for_timeout(1000) # 1 second

                    except TimeoutError:
                        if sem_dados_locator.is_visible():
                            q.put(("status", f"    -> 'Sem dados' for product {produto_interno_cliente}. Skipping item."))
                            not_found_items.append({
                                "chave": chave,
                                "produto": produto_interno_cliente,
                                "motivo": "Produto específico não encontrado"
                            })
                            continue
                        else:
                            q.put(("status", "    -> ERROR: No product row OR 'Sem dados' text found. Skipping item."))
                            continue
                        
                    q.put(("status", "    -> Item found. (Add your action here)"))
                    
                    # --- Clear product filter ---
                    product_filter_input.fill("")
                    page.wait_for_timeout(200)

                q.put(("status", f"✅ Finished group: {chave}"))
                q.put(("progress", group_index + 1))

                # --- Clear main 'chave' search ---
                chave_input.fill("")
                page.wait_for_timeout(500) 

            except Exception as e:
                q.put(("status", f"❌ Error on group {chave}: {e}. Skipping to next group."))
                try:
                    frame_locator.locator(".dx-texteditor-input").first.fill("")
                except Exception as e_clear:
                    q.put(("status", f"    -> Failed to clear input: {e_clear}"))

        q.put(("status", "🎉 All order groups processed successfully."))

        processar_e_Fazer_upload_Arquivos(page,found_items)

        if not_found_items:
            q.put(("status", f"⚠️ Found {len(not_found_items)} items that were not on the site:"))
            # for item in not_found_items:
            #     q.put(("status", f"    -> Chave: {item['chave']}, Produto: {item['produto']}, Motivo: {item['motivo']}"))

    except Exception as e:
        q.put(("status", f"❌ A critical error occurred: {e}"))


def processar_e_Fazer_upload_Arquivos(page: Page ,  items : list):
   
    page.locator("#iframe-servico").content_frame.get_by_role("button", name="DOWNLOAD PLANILHA").click()

    with page.expect_download() as download_info:
        page.locator("#iframe-servico").content_frame.get_by_role("menuitem", name="APENAS SELECIONADOS").click()

    download = download_info.value

    # Ensure folder exists
    os.makedirs("Arquivos", exist_ok=True)

    # Correct saving method
    save_path = f"Arquivos/{dt.datetime.now().strftime('%Y-%m-%d')}.xlsx"
    download.save_as(save_path)


    print(items)

    # for item in items:
    #         numero_cliente = item["numero_cliente"]
    #         produto_interno_cliente = item["produto_interno_cliente"]
    #         quantidade = item["quantidade"]




        
    time.sleep(2)







