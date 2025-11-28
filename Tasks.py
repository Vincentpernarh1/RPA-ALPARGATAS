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



from Azure_Access import main

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


def load_filepath():
    CARTEURA_GRUPO_PATH = None
    CARTEURA_GRUPO_folder = os.path.join(base_path,"Dados")

    for file in os.listdir(CARTEURA_GRUPO_folder):
        
        if "CARTEIRA GRUPO" in file and file.endswith(".xlsx") and not file.startswith("~$"):
            CARTEURA_GRUPO_PATH = os.path.join(CARTEURA_GRUPO_folder, file)
            break

    return CARTEURA_GRUPO_PATH


 
def convert_excel_date(excel_value):
    """Convert Excel date serial number to DD/MM/YYYY format"""
    try:
        if isinstance(excel_value, (int, float)):
            
            print(f"Converting Excel date serial: {excel_value}")
            # Excel date serial number (days since 1899-12-30)
            # Use 1899-12-30 as the base date (Excel's actual epoch)
            from datetime import datetime, timedelta
            excel_epoch = datetime(1899, 12, 30)
            result_date = excel_epoch + timedelta(days=excel_value)
            return result_date.strftime('%d/%m/%Y')
        elif isinstance(excel_value, str):
            # Already a string, try to parse and reformat
            from datetime import datetime
            # Try common date formats
            for fmt in ['%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y']:
                try:
                    parsed = datetime.strptime(excel_value, fmt)
                    return parsed.strftime('%d/%m/%Y')
                except:
                    continue
            return str(excel_value)
        else:
            return str(excel_value)
    except Exception as e:
        print(f"Error converting date {excel_value}: {e}")
        
        
        return str(excel_value)
    

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
        print(df["PREVISÃO DE ENTREGA"].head(3))
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
    static_data = load_static_data()  # Load static data at the start
    try:
        q.put(("progress", 10))
        q.put(("status", "Fetching order data from SharePoint..."))
        df, drive_id, file_id = Order_datas_from_sharepoint(q)

        if df is None or df.empty:
            q.put(("status", "⚠️ No order data found to process."))
            return

        df['PRODUTO INTERNO CLIENTE'] = df['PRODUTO INTERNO CLIENTE'].astype(str)
        df['chave_pedido_loja'] = df['chave_pedido_loja'].astype(str)

        grouped_orders = df.groupby('chave_pedido_loja')
        group_carro = df.groupby('CARRO')
        
        total_groups = len(grouped_orders)
        total_items = len(df)
        q.put(("progress", 15))
        q.put(("status", f"Found {total_items} total items in {total_groups} unique groups."))
        
        frame_locator = page.locator("#iframe-servico").first.content_frame
        
        # --- LOOP 1: By Unique 'chave_pedido_loja' ---
        for group_index, (chave, group_df) in enumerate(grouped_orders):
            
            q.put(("status", f"--- Processing Group {group_index + 1}/{total_groups}: {chave} ---"))
            
            # Calculate progress: 15% (initial) + 70% (processing) = 85% max before upload
            progress_per_group = 70 / total_groups if total_groups > 0 else 0
            current_progress = 15 + (group_index * progress_per_group)
            q.put(("progress", int(current_progress)))
            
            try:
                chave_input = frame_locator.locator(".dx-texteditor-input").first
                chave_input.fill(chave)
                
                page.wait_for_timeout(2000) # 2 seconds

                q.put(("status", "    -> Waiting for filter result..."))
                
                data_locator = frame_locator.locator(".dx-row.dx-data-row > td:nth-child(8)").first
                sem_dados_locator = frame_locator.get_by_text("Sem dados")

                try:
                    # 1. Check for data (with a shorter timeout, as the page is stable)
                    data_locator.wait_for(state="visible", timeout=5000)
                    q.put(("status", "    -> Data found. Proceeding to CARRO groups."))

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

                # --- LOOP 2: Group by 'CARRO' within this chave_pedido_loja ---
                grouped_by_carro = group_df.groupby('CARRO')
                total_carros = len(grouped_by_carro)
                q.put(("status", f"    -> Found {total_carros} CARRO groups for {chave}"))
                
                for carro_index, (carro, carro_df) in enumerate(grouped_by_carro):
                    q.put(("status", f"    --- Processing CARRO {carro_index + 1}/{total_carros}: {carro} for {chave} ---"))
                    
                    found_items = []  # Reset for each CARRO group
                    
                    # --- LOOP 3: By 'PRODUTO INTERNO CLIENTE' for this CARRO ---
                    for _, row in carro_df.iterrows():
                        produto_interno_cliente = row['PRODUTO INTERNO CLIENTE']
                        numero_cliente = row['Nº Pedido Cliente']
                        data_deprevisao_de_entrega = row['PREVISÃO DE ENTREGA']

                        if "TRADICIONAL" in row['Descrição']:
                            quantidade = (row['Qtd. Faturada']/24)
                        else:
                            quantidade = (row['Qtd. Faturada']/12)
                    
                        q.put(("status", f"        -> Filtering for product: {produto_interno_cliente}"))
                       
                        product_filter_input = frame_locator.locator("input[aria-label='Filtro de célula']").nth(3)
                                          
                        product_filter_input.fill("")
                        page.wait_for_timeout(300) # Short pause for clear
                        
                        product_filter_input.fill(produto_interno_cliente)
                        
                        page.wait_for_timeout(1000) # 1 second

                        q.put(("status", "        -> Waiting for product filter result..."))
                        
                        try:
                            data_locator.wait_for(state="visible", timeout=4000)
                            q.put(("status", "        -> Product found."))

                            # Click all visible checkboxes for this product
                            checkboxes = page.locator("#iframe-servico").content_frame.get_by_role("gridcell", name="Selecionar linha").get_by_role("checkbox")
                            checkbox_count = checkboxes.count()
                            
                            if checkbox_count > 0:
                                for idx in range(checkbox_count):
                                    try:
                                        checkboxes.nth(idx).click()
                                        page.wait_for_timeout(100)  # Small delay between clicks
                                    except Exception as click_err:
                                        q.put(("status", f"        -> Warning: Could not click checkbox {idx}: {click_err}"))
                                q.put(("status", f"        -> Clicked {checkbox_count} checkboxes"))
                            else:
                                q.put(("status", "        -> ⚠️ No checkboxes found to select"))

                            found_items.append({
                                    "numero_cliente": numero_cliente,
                                    "produto_interno_cliente": produto_interno_cliente,
                                    "quantidade": quantidade,
                                    "data_deprevisao_de_entrega": data_deprevisao_de_entrega,
                                    "caracteristica": static_data["caracteristica"],
                                    "caracteristica_do_veiculo": static_data["caracteristica_do_veiculo"],
                                    "chave_pedido_loja": chave,  # Full key for Excel matching (Order#-Store#)
                                    "carro": carro  # Add CARRO identifier
                                })
                           
                            page.wait_for_timeout(1000) # 1 second

                        except TimeoutError:
                            if sem_dados_locator.is_visible():
                                q.put(("status", f"        -> 'Sem dados' for product {produto_interno_cliente}. Skipping item."))
                                not_found_items.append({
                                    "chave": chave,
                                    "produto": produto_interno_cliente,
                                    "carro": carro,
                                    "motivo": "Produto específico não encontrado"
                                })
                                continue
                            else:
                                q.put(("status", "        -> ERROR: No product row OR 'Sem dados' text found. Skipping item."))
                                continue
                            
                        # --- Clear product filter ---
                        product_filter_input.fill("")
                        page.wait_for_timeout(200)

                    # --- Process and upload Excel file for this CARRO group ---
                    if found_items:
                        q.put(("status", f"    -> Processing and uploading Excel file for {chave}-CARRO{carro}..."))
                        processar_e_Fazer_upload_Arquivos(page, found_items, q)
                        q.put(("status", f"    ✅ Finished {chave}-CARRO{carro} (Group {carro_index + 1}/{total_carros})"))
                    else:
                        q.put(("status", f"    ⚠️ No items found for {chave}-CARRO{carro}. Skipping upload."))

                q.put(("status", f"✅ Finished all CARROs for group: {chave}"))
                
                # Update progress at end of group
                progress_per_group = 70 / total_groups if total_groups > 0 else 0
                completed_progress = 15 + ((group_index + 1) * progress_per_group)
                q.put(("progress", int(completed_progress)))

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
        q.put(("progress", 95))

        q.put(("status", "Finalizing..."))
        if not_found_items:
            q.put(("status", f"⚠️ Found {len(not_found_items)} items that were not on the site:"))
        
        q.put(("progress", 100))
           
    except Exception as e:
        q.put(("status", f"❌ A critical error occurred: {e}"))


def processar_e_Fazer_upload_Arquivos(page: Page, items: list, q):
    """
    Download Excel file, process it with xlwings, fill data from items,
    save and prepare for upload to page.
    """
    page.locator("#iframe-servico").content_frame.get_by_role("button", name="DOWNLOAD PLANILHA").click()

    with page.expect_download() as download_info:
        page.locator("#iframe-servico").content_frame.get_by_role("menuitem", name="APENAS SELECIONADOS").click()

    download = download_info.value

    # Ensure folder exists
    os.makedirs("Arquivos", exist_ok=True)

    # Save the downloaded file
    save_path = f"Arquivos/{dt.datetime.now().strftime('%Y-%m-%d')}.xlsx"
    if os.path.exists(save_path):
        os.remove(save_path)
        q.put(("status", f"Existing file {save_path} removed."))
    download.save_as(save_path)

    # Process the Excel file with xlwings
    pedidos = processar_excel_com_dados(save_path, items,q)
    if pedidos == 'Done':
        print("Excel processing completed successfully.")
        q.put(("status", "    ✅ Excel processing completed successfully."))

        try:
            q.put(("status", "    -> Preparing to upload file..."))
            page.wait_for_timeout(1000)  # Wait for page stability
            
            # Click the "Upload da planilha" button
            upload_button = page.locator("#iframe-servico").content_frame.get_by_role("button", name="Upload da planilha")
            upload_button.wait_for(state="visible", timeout=5000)
            upload_button.click()
            q.put(("status", "    -> Clicked 'Upload da planilha' button"))
            page.wait_for_timeout(1000)
            
            # Wait for the modal/upload dialog to appear
            q.put(("status", "    -> Waiting for upload dialog..."))
            page.wait_for_timeout(1000)
            
            frame = page.locator("#iframe-servico").content_frame
            file_input = frame.locator("input[type='file']")
            
            # Wait for the file input to be ready
            file_input.wait_for(state="attached", timeout=5000)
            q.put(("status", "    -> Found file input element"))
            
            # Set the file directly to the input element
            file_input.set_input_files(save_path)
            q.put(("status", f"    -> ✅ File uploaded: {os.path.basename(save_path)}"))
            page.wait_for_timeout(2000)  # Wait for upload to process
            
            # Look for a confirm/submit button after upload
            try:
                confirm_button = frame.get_by_role("button", name=re.compile(r"(Enviar|Confirmar|OK|Upload)", re.IGNORECASE))
                confirm_button.wait_for(state="visible", timeout=3000)
                confirm_button.click()
                q.put(("status", "    -> Clicked confirm button"))
                page.locator("#iframe-servico").content_frame.get_by_role("button", name="ir para a lista de logs").click()
                page.pause()

                # Extrair_logs_de_upload_e_Atualizar_sharepoint(page, q)

                # page.locator("#iframe-servico").content_frame.get_by_role("row", name="Expandir 21/11/2025 12:26:").locator("div").first.click()
                # page.locator("#iframe-servico").content_frame.get_by_role("gridcell", name="Sucesso").click()
                # page.locator("#iframe-servico").content_frame.get_by_role("gridcell", name="Erro", exact=True).click()
                # page.locator("#iframe-servico").content_frame.get_by_role("gridcell", name="Erro ao gerar a demanda \"").click()
                
                page.wait_for_timeout(1000)
            except TimeoutError:
                q.put(("status", "    -> ℹ️ No explicit confirm button found (may auto-submit)"))
            
            # Try to proceed to next step if available
            try:
                data_button = frame.get_by_text("Data sugerida para entrega")
                data_button.wait_for(state="visible", timeout=3000)
                data_button.click()
                q.put(("status", "    -> Proceeding to next step (Data sugerida)"))
            except TimeoutError:
                q.put(("status", "    -> ℹ️ Next step not available yet"))
            
            q.put(("status", "    ✅ File upload process completed"))
            
        except TimeoutError as e:
            q.put(("status", f"    ❌ Upload timeout: {e}"))
            print(f"Upload timeout: {e}")
        except Exception as e:
            q.put(("status", f"    ❌ Upload error: {e}"))
            print(f"Upload error: {e}")
            import traceback
            traceback.print_exc()

    time.sleep(2)




def processar_excel_com_dados(file_path: str, items: list, q):
   
   
    try:
        # Open the workbook with xlwings
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = True
        wb = app.books.open(file_path,update_links=False, read_only=False)
        ws = wb.sheets['Planilha1']
        
        # Get column headers from row 3
        headers = {}
        for col_idx, cell in enumerate(ws.range(3, 1).expand('right').value, start=1):
            if cell:
                headers[cell] = col_idx
        
        # Map column names to their indices (exact matching)
        col_mapping = {
            'quantidade_entrega': None,
            'data_sugerida': None,
            'caracteristica_veiculo': None,
            'caracteristica_carga': None,
            'observacao_fornecedor': None,
            'demanda': None,
            'codigo_pedido': None,
            'codigo_produto': None
        }
        
        # Search for columns (exact matching first, then fuzzy)
        for header, col_idx in headers.items():
            header_lower = header.lower() if header else ""
            
            # Exact matches first
            if header_lower == 'quantidade entrega':
                col_mapping['quantidade_entrega'] = col_idx
            elif header_lower == 'data sugerida de entrega':
                col_mapping['data_sugerida'] = col_idx
            elif header_lower == 'característica do veículo':
                col_mapping['caracteristica_veiculo'] = col_idx
            elif header_lower == 'característica da carga':
                col_mapping['caracteristica_carga'] = col_idx
            elif header_lower == 'observação/ fornecedor (opcional)':
                col_mapping['observacao_fornecedor'] = col_idx
            elif header_lower == 'demanda':
                col_mapping['demanda'] = col_idx
            elif header_lower == 'código do pedido cliente':
                col_mapping['codigo_pedido'] = col_idx
            elif header_lower == 'código produto cliente':
                col_mapping['codigo_produto'] = col_idx
        
        # Log found columns for debugging
        print(f"Headers found: {headers}")
        q.put(("status", f"    -> Excel columns found: {len(headers)} columns detected"))
        print(f"Column mapping: {col_mapping}")
        print(f"Items to match: {len(items)} items")
        
        # Process data rows (starting from row 4, since headers are in row 3)
        row_num = 4
        matched_count = 0
        
        while True:
            # Get values from current row
            code_pedido = ws.range(row_num, col_mapping['codigo_pedido']).value if col_mapping['codigo_pedido'] else None
            code_produto = ws.range(row_num, col_mapping['codigo_produto']).value if col_mapping['codigo_produto'] else None
            
            # Stop if we've reached empty rows
            if not code_pedido or not code_produto:
                print(f"Reached end of data at row {row_num}")
                break
            
            # Convert to string for comparison
            code_pedido = str(code_pedido).strip()
            code_produto = str(code_produto).strip()
            
            print(f"Row {row_num}: Looking for pedido={code_pedido}, produto={code_produto}")
            
            # Find matching item in the items list
            matching_item = None
            for item in items:
                # Match by chave_pedido_loja (Order#-Store#) and produto_interno_cliente
                item_chave = str(item.get('chave_pedido_loja', '')).strip()
                item_produto = str(item.get('produto_interno_cliente', '')).strip()

                print(f"  Checking item: chave={item_chave}, produto={item_produto}")
                
                if item_chave == code_pedido and item_produto == code_produto:
                    matching_item = item
                    print(f"  ✅ Match found: {item_chave} - {item_produto}")
                    break
            
            # If matching item found, fill the columns
            if matching_item:
                # Fill Quantidade entrega
                if col_mapping['quantidade_entrega']:
                    quantidade_val = matching_item.get('quantidade')
                    ws.range(row_num, col_mapping['quantidade_entrega']).value = quantidade_val
                    print(f"    -> Set Quantidade entrega: {quantidade_val}")
                
                # Fill Data sugerida de entrega (convert to Excel date format)
                if col_mapping['data_sugerida']:
                    data_val = matching_item.get('data_deprevisao_de_entrega')
                    
                    print(f"    -> Raw Data sugerida: {data_val}")
                    # Convert to datetime and set as date format
                    try:
                        from datetime import datetime, timedelta
                        if isinstance(data_val, (int, float)):
                            # Already an Excel serial, convert back to date
                            # Use 1899-12-30 as the base date (Excel's actual epoch)
                            excel_epoch = datetime(1899, 12, 30)
                            result_date = excel_epoch + timedelta(days=data_val)
                        else:
                            # Parse string date
                            result_date = None
                            for fmt in ['%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y']:
                                try:
                                    result_date = datetime.strptime(str(data_val), fmt)
                                    break
                                except:
                                    continue
                            if result_date is None:
                                raise ValueError(f"Could not parse date: {data_val}")
                        
                        
                        print("result_date : ",result_date  , "After transformation : ", result_date.strftime('%m-%d-%Y'))
                        
                        # Set the cell value as a string in YYYY-DD-MM format
                        cell = ws.range(row_num, col_mapping['data_sugerida'])
                        cell.value = result_date.strftime('%m-%d-%Y')  # Store as string in YYYY-MM-DD format
                        
                        print(f"    -> Set Data sugerida: {result_date.strftime('%Y-%m-%d')} (as string)")
                    except Exception as date_err:
                        # Fallback: set as string
                        ws.range(row_num, col_mapping['data_sugerida']).value = convert_excel_date(data_val)
                        print(f"    -> Set Data sugerida: {data_val} (as string, fallback)")
                
                # Fill Característica do veículo
                if col_mapping['caracteristica_veiculo']:
                    carac_val = matching_item.get('caracteristica_do_veiculo')
                    ws.range(row_num, col_mapping['caracteristica_veiculo']).value = carac_val
                    print(f"    -> Set Característica do veículo: {carac_val}")
                
                # Fill Característica da carga
                if col_mapping['caracteristica_carga']:
                    carga_val = matching_item.get('caracteristica')
                    ws.range(row_num, col_mapping['caracteristica_carga']).value = carga_val
                    print(f"    -> Set Característica da carga: {carga_val}")
                
                # Fill Demanda with chave_pedido_loja-CARRO
                if col_mapping['demanda']:
                    chave_val = matching_item.get('chave_pedido_loja')
                    carro_val = matching_item.get('carro', '')
                    demanda_val = f"{chave_val}-{carro_val}" if carro_val else chave_val
                    ws.range(row_num, col_mapping['demanda']).value = demanda_val
                    print(f"    -> Set Demanda: {demanda_val}")
                
                # Fill Observação/ Fornecedor with chave_pedido_loja-CARRO
                if col_mapping['observacao_fornecedor']:
                    chave_val = matching_item.get('chave_pedido_loja')
                    carro_val = matching_item.get('carro', '')
                    observacao_val = f"{chave_val}-CARRO{carro_val}" if carro_val else chave_val
                    ws.range(row_num, col_mapping['observacao_fornecedor']).value = observacao_val
                    print(f"    -> Set Observação/Fornecedor: {observacao_val}")
                
                q.put(("status", f"    -> Row {row_num}: ✅ Filled all fields for {code_pedido}"))
                matched_count += 1
            else:
                print(f"  ⚠️ No match found for: {code_pedido} - {code_produto}")
                q.put(("status", f"    -> Row {row_num}: ⚠️ No match for {code_pedido}"))
            
            row_num += 1
        
        print(f"✅ Matched {matched_count} out of {len(items)} items")
        q.put(("status", f"    -> Matched {matched_count} out of {len(items)} items"))
        
        # Save the modified workbook
        wb.save()
        wb.close()
        app.quit()
        
        q.put(("status", f"    ✅ Excel file processed and saved"))
        print(f"✅ Excel file processed and saved: {file_path}")
        return 'Done'
        
    except Exception as e:
        print(f"❌ Error processing Excel file: {e}")
        q.put(("status", f"    ❌ Error processing Excel: {e}"))
        import traceback
        traceback.print_exc()



# def Extrair_logs_de_upload_e_Atualizar_sharepoint(page: Page, q):
#     print("Function 'Extrair_logs_de_upload_e_Atualizar_sharepoint' called.")




