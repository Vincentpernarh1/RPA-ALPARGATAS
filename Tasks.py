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

from RestAPIHelper import main, update_excel_rows, update_protocol_async

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
def sharepoint_main_in_thread(result_queue: queue.Queue):
   
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


# +++++++++ HELPER FUNCTION TO UPDATE SHAREPOINT PROTOCOL IN BACKGROUND +++++++++
def update_sharepoint_protocol_in_thread(drive_id: str, file_id: str, protocol_data: dict, q: queue.Queue):
    
    try:
        # Create and set a new event loop for this thread
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        
        # Run the async update function
        protocol_data_list = [protocol_data]  # Wrap in list as function expects list
        loop.run_until_complete(update_protocol_async(drive_id, file_id, protocol_data_list))
        
        q.put(("status", f"    -> ✅ SharePoint atualizado: {protocol_data['chave']}-{protocol_data['carro']} = {protocol_data['protocol']}"))
    except Exception as e:
        q.put(("status", f"    -> ⚠️ Erro ao atualizar SharePoint: {e}"))
        import traceback
        traceback.print_exc()
    finally:
        # Clean up the loop
        loop.close()


def Order_datas_from_sharepoint(q):
    
    q.put(("status", "Obtendo dados do SharePoint..."))
    result_queue = queue.Queue() # A new queue just for this thread's result

    # Create and start the thread, targeting our new helper function
    sharepoint_thread = threading.Thread(target=sharepoint_main_in_thread, args=(result_queue,))
    sharepoint_thread.start()

    # Wait for the thread to finish its work
    sharepoint_thread.join() # <--- You are here. The thread is finished.

    try:
        # 1. Get the item from the queue
        result = result_queue.get_nowait()
        if isinstance(result, Exception):
            # If the thread sent back an error, handle it
            q.put(("status", f"❌ Erro ao obter dados do SharePoint: {result}"))
            raise result # Re-raise the error
       
        df, drive_id, file_id = result
        
        # Cleaning up the dataframe to remove none values
        df = df[df['Nº Pedido Cliente'].notna()].copy()
        
        df['chave_pedido_loja'] = df['Nº Pedido Cliente'].astype(str) + '-' + df['CÓD LOJA'].astype(str).str.split('-').str[0]
      
        q.put(("status", "✅ Dados do SharePoint obtidos com sucesso."))

        return df, drive_id, file_id

    except queue.Empty:
        # This shouldn't happen if join() worked, but it's safe to have
        q.put(("status", "❌ Thread do SharePoint finalizou sem resultado."))
        return None, None, None
    except Exception as e:
        # Handle any other error
        q.put(("status", f"❌ Falha ao processar resultado do SharePoint: {e}"))
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
        
        try:
            page.get_by_role("button", name="Gestão de Pedidos Gestão da")
            q.put(("status", "✅ Login realizado com sucesso"))
            q.put(("progress", 5))
        except TimeoutError:
            q.put(("status", "⚠️ Tentativa de login falhou "))
            page.screenshot(path="login_failed.png")
        page.get_by_role("button", name="Gestão de Pedidos Gestão da").click()
        page.locator("#iframe-servico").content_frame.get_by_role("button", name="AÇÕES").click()
        page.locator("#iframe-servico").content_frame.get_by_role("menuitem", name="CONSUMIR ITENS").click()
       
        process_orders(page, q)
           
        page.wait_for_timeout(100000)

    except Exception as e:
        q.put(("status", f"❌ Erro durante o login: {e}"))

def process_orders(page: Page, q):

    not_found_items = []
    static_data = load_static_data()  # Load static data at the start
    try:
        q.put(("progress", 10))

        df, drive_id, file_id = Order_datas_from_sharepoint(q)
        
    
        
        if df is None or df.empty:
            q.put(("status", "⚠️ Nenhum dado de pedido encontrado para processar."))
            return

        df['PRODUTO INTERNO CLIENTE'] = df['PRODUTO INTERNO CLIENTE'].astype(str)
        df['chave_pedido_loja'] = df['chave_pedido_loja'].astype(str)
        
        # Extract loja (store) from CÓD LOJA column (before the dash)
        df['loja'] = df['CÓD LOJA'].astype(str).str.split('-').str[0]
        
        # Group by loja first
        grouped_by_loja = df.groupby('loja')
        
        total_lojas = len(grouped_by_loja)
        total_items = len(df)
        q.put(("progress", 15))
        q.put(("status", f"Encontrados {total_items} itens no total em {total_lojas} lojas únicas."))
        
        # Calculate total CARRO groups across all lojas for progress
        total_all_carros = sum(len(loja_df.groupby('CARRO')) for _, loja_df in grouped_by_loja)
        
        frame_locator = page.locator("#iframe-servico").first.content_frame
        
        # --- LOOP 1: By Unique 'loja' (Store) ---
        loja_index = 0
        carro_global_index = 0
        for loja, loja_df in grouped_by_loja:
            q.put(("status", f"=== Processando Loja {loja_index + 1}/{total_lojas}: {loja} ==="))
            
            # Within each loja, group by CARRO
            grouped_by_carro = loja_df.groupby('CARRO')
            total_carros_in_loja = len(grouped_by_carro)
            q.put(("status", f"    Encontrados {total_carros_in_loja} grupos CARRO para loja {loja}"))
            
            # --- LOOP 2: By Unique 'CARRO' within this loja ---
            for carro_index, (carro, carro_df) in enumerate(grouped_by_carro):
                q.put(("status", f"    --- Processando CARRO {carro_index + 1}/{total_carros_in_loja}: {carro} para loja {loja} ---"))
                
                # Check if protocol column indicates processing is needed (error or null)
                needs_processing = False
                for _, row in carro_df.iterrows():
                    protocol_val = str(row.get('PROTOCOLO DA SOLICITAÇÃO', '')).strip()
                    if pd.isna(row.get('PROTOCOLO DA SOLICITAÇÃO')) or protocol_val == '' or 'ERRO' in protocol_val.upper() or 'Erro ao gerar a demanda' in protocol_val:
                        needs_processing = True
                        break
                
                if not needs_processing:
                    q.put(("status", f"    --- CARRO {carro} já possui protocolo válido. Pulando. ---"))
                    carro_global_index += 1
                    continue  # Skip to next CARRO
                
                # Calculate progress: 15% (initial) + 70% (processing) = 85% max before upload
                progress_per_carro = 70 / total_all_carros if total_all_carros > 0 else 0
                current_progress = 15 + (carro_global_index * progress_per_carro)
                q.put(("progress", int(current_progress)))
                
                found_items = []  # Reset for each CARRO group
                
                # Get unique chaves in this CARRO
                unique_chaves = carro_df['chave_pedido_loja'].unique()
                q.put(("status", f"        -> Encontradas {len(unique_chaves)} chaves únicas para CARRO {carro}"))
                
                try:
                    # --- LOOP 3: By unique 'chave_pedido_loja' within this CARRO ---
                    for chave in unique_chaves:
                        q.put(("status", f"        --- Buscando chave {chave} para CARRO {carro} ---"))
                        
                        chave_input = frame_locator.locator(".dx-texteditor-input").first
                        chave_input.fill(chave.strip())
                        
                        page.wait_for_timeout(2000)  # 2 seconds
                        
                        data_locator = frame_locator.locator(".dx-row.dx-data-row > td:nth-child(8)").first
                        sem_dados_locator = frame_locator.get_by_text("Sem dados")
                        
                        try:
                            # Check for data
                            data_locator.wait_for(state="visible", timeout=5000)
                            q.put(("status", f"        -> Dados encontrados para chave {chave}."))
                            
                        except TimeoutError:
                            # No data found
                            if sem_dados_locator.is_visible():
                                q.put(("status", f"        -> 'Sem dados' para chave {chave}. Pulando chave."))
                                for _, row in carro_df[carro_df['chave_pedido_loja'] == chave].iterrows():
                                    not_found_items.append({
                                        "chave": chave,
                                        "produto": row['PRODUTO INTERNO CLIENTE'],
                                        "carro": carro,
                                        "motivo": "Chave não encontrada"
                                    })
                                continue
                            else:
                                q.put(("status", f"        -> ERRO: Nenhuma linha de dados para chave {chave}. Pulando chave."))
                                continue
                        
                        # --- LOOP 4: By 'PRODUTO INTERNO CLIENTE' for this chave in CARRO ---
                        for _, row in carro_df[carro_df['chave_pedido_loja'] == chave].iterrows():
                            produto_interno_cliente = row['PRODUTO INTERNO CLIENTE']
                            numero_cliente = row['Nº Pedido Cliente']
                            data_deprevisao_de_entrega = row['PREVISÃO DE ENTREGA']
                            
                            # Convert Qtd. Faturada to numeric
                            qtd_faturada = pd.to_numeric(row['Qtd. Faturada'], errors='coerce')
                            if pd.isna(qtd_faturada):
                                q.put(("status", f"            -> ⚠️ Qtd. Faturada inválida para produto {produto_interno_cliente}, pulando item."))
                                continue
                            
                            if "TRADICIONAL" in row['Descrição']:
                                quantidade = (qtd_faturada / 24)
                            else:
                                quantidade = (qtd_faturada / 12)
                            
                            q.put(("status", f"            -> Filtrando produto: {produto_interno_cliente} para chave {chave}"))
                            
                            product_filter_input = frame_locator.locator("input[aria-label='Filtro de célula']").nth(3)
                            
                            product_filter_input.fill("")
                            page.wait_for_timeout(300)  # Short pause for clear
                            
                            product_filter_input.fill(produto_interno_cliente)
                            
                            page.wait_for_timeout(1000)  # 1 second
                            
                            try:
                                data_locator.wait_for(state="visible", timeout=4000)
                                q.put(("status", "            -> Produto encontrado."))
                                
                                # Click all visible checkboxes for this product
                                checkboxes = page.locator("#iframe-servico").content_frame.get_by_role("gridcell", name="Selecionar linha").get_by_role("checkbox")
                                checkbox_count = checkboxes.count()
                                
                                if checkbox_count > 0:
                                    successful_clicks = 0
                                    for idx in range(checkbox_count):
                                        checkbox = checkboxes.nth(idx)
                                        clicked = False
                                        
                                        # Try up to 2 times
                                        for attempt in range(2):
                                            try:
                                                checkbox.wait_for(state="visible", timeout=2000)
                                                checkbox.scroll_into_view_if_needed()
                                                checkbox.click(timeout=1500)
                                                successful_clicks += 1
                                                clicked = True
                                                page.wait_for_timeout(100)
                                                break
                                                
                                            except TimeoutError:
                                                if attempt == 1:
                                                    break
                                                page.wait_for_timeout(300)
                                                
                                            except Exception as click_err:
                                                if attempt == 1:
                                                    break
                                                page.wait_for_timeout(300)
                                        
                                        if not clicked and idx > 0:
                                            break
                                    
                                    q.put(("status", f"            -> Clicado em {successful_clicks} checkboxes"))
                                else:
                                    q.put(("status", "            -> ⚠️ Nenhuma checkbox encontrada para selecionar"))
                                
                                found_items.append({
                                    "numero_cliente": numero_cliente,
                                    "produto_interno_cliente": produto_interno_cliente,
                                    "quantidade": quantidade,
                                    "data_deprevisao_de_entrega": data_deprevisao_de_entrega,
                                    "caracteristica": static_data["caracteristica"],
                                    "caracteristica_do_veiculo": static_data["caracteristica_do_veiculo"],
                                    "chave_pedido_loja": chave,
                                    "carro": carro
                                })
                                
                                page.wait_for_timeout(1000)  # 1 second
                                
                            except TimeoutError:
                                if sem_dados_locator.is_visible():
                                    q.put(("status", f"            -> 'Sem dados' para produto {produto_interno_cliente}. Pulando item."))
                                    not_found_items.append({
                                        "chave": chave,
                                        "produto": produto_interno_cliente,
                                        "carro": carro,
                                        "motivo": "Produto específico não encontrado"
                                    })
                                    continue
                                else:
                                    q.put(("status", "            -> ERRO: Nenhuma linha de produto encontrado. Pulando item."))
                                    continue
                            
                            # Clear product filter
                            product_filter_input.fill("")
                            page.wait_for_timeout(200)
                        
                        # Clear main 'chave' search after processing all products for this chave
                        chave_input.fill("")
                        page.wait_for_timeout(500)
                    
                    # --- Process and upload Excel file for this CARRO group ---
                    if found_items:
                        q.put(("status", f"    -> Processando e enviando arquivo Excel para loja {loja} - CARRO {carro}..."))
                        processar_e_Fazer_upload_Arquivos(page, found_items, q, drive_id, file_id)
                        q.put(("status", f"    ✅ Finalizado loja {loja} - CARRO {carro}"))
                    else:
                        q.put(("status", f"    ⚠️ Nenhum item encontrado para loja {loja} - CARRO {carro}. Pulando envio."))
                    
                    carro_global_index += 1
                    
                except Exception as e:
                    q.put(("status", f"❌ Erro no CARRO {carro}: {e}. Pulando para o próximo CARRO."))
                    carro_global_index += 1
                    try:
                        frame_locator.locator(".dx-texteditor-input").first.fill("")
                    except Exception as e_clear:
                        q.put(("status", f"    -> Falha ao limpar campo: {e_clear}"))
            
            # Increment loja index after processing all CARRO groups for this loja
            loja_index += 1
            q.put(("status", f"=== ✅ Finalizada Loja: {loja} ==="))

        q.put(("status", "🎉 Todos os grupos de pedidos processados com sucesso."))
        q.put(("progress", 95))

        q.put(("status", "Finalizando..."))
        # if not_found_items:
        #     q.put(("status", f"⚠️ Encontrados {len(not_found_items)} itens que não estavam no site:"))
        
        q.put(("progress", 100))
           
    except Exception as e:
        q.put(("status", f"❌ Ocorreu um erro crítico: {e}"))


def processar_e_Fazer_upload_Arquivos(page: Page, items: list, q, drive_id: str = None, file_id: str = None):
    
    page.locator("#iframe-servico").content_frame.get_by_role("button", name="DOWNLOAD PLANILHA").click()

    with page.expect_download() as download_info:
        page.locator("#iframe-servico").content_frame.get_by_role("menuitem", name="APENAS SELECIONADOS").click()

    download = download_info.value

    # Ensure folder exists
    os.makedirs("Arquivos", exist_ok=True)

    # Save the downloaded file
    save_path = f"Arquivos/{dt.datetime.now().strftime('%d-%m-%Y')}.xlsx"
    
    if os.path.exists(save_path):
        os.remove(save_path)
        q.put(("status", f"Arquivo existente {save_path} removido."))
    download.save_as(save_path)

    # Process the Excel file with xlwings
    matched_count = processar_excel_com_dados(save_path, items,q)
    if matched_count > 0:
        print("Excel processing completed successfully.")
        q.put(("status", "    ✅ Processamento do Excel concluído com sucesso."))

        try:
            q.put(("status", "    -> Preparando para enviar arquivo..."))
            page.wait_for_timeout(1000)  # Wait for page stability
            
            # Click the "Upload da planilha" button
            upload_button = page.locator("#iframe-servico").content_frame.get_by_role("button", name="Upload da planilha")
            upload_button.wait_for(state="visible", timeout=5000)
            upload_button.click()
            q.put(("status", "    -> Clicado no botão 'Upload da planilha'"))
            page.wait_for_timeout(1000)
            
            # Wait for the modal/upload dialog to appear
            q.put(("status", "    -> Aguardando janela de upload..."))
            page.wait_for_timeout(1000)
            
            frame = page.locator("#iframe-servico").content_frame
            file_input = frame.locator("input[type='file']")
            
            # Wait for the file input to be ready
            file_input.wait_for(state="attached", timeout=5000)
            q.put(("status", "    -> Elemento de entrada de arquivo encontrado"))
            
            # Set the file directly to the input element
            file_input.set_input_files(save_path)
            q.put(("status", f"    -> ✅ Arquivo enviado: {os.path.basename(save_path)}"))
            page.wait_for_timeout(3000)  # Wait for upload to process
            
            # Look for a confirm/submit button after upload
            try:
                confirm_button = frame.get_by_role("button", name=re.compile(r"(Enviar|Confirmar|OK|Upload)", re.IGNORECASE))
                confirm_button.wait_for(state="visible", timeout=3000)
                confirm_button.click()
                q.put(("status", "    -> Clicado no botão confirmar"))
                page.locator("#iframe-servico").content_frame.get_by_role("button", name="ir para a lista de logs").click()
                page.wait_for_timeout(6000)  # Wait for logs page to load
                q.put(("status", "    -> Navegado para página de logs"))

                # Extract protocol and get chave/carro from first item
                if items:
                    chave = items[0].get('chave_pedido_loja', '')
                    carro = items[0].get('carro', '')
                    protocol_data = Extrair_logs_de_upload_e_Atualizar_sharepoint(page, chave, carro, q)
                    
                    if protocol_data and drive_id and file_id:
                        q.put(("status", f"    ✅ Dados do protocolo: {protocol_data}"))
                        
                        # Launch background thread to update SharePoint (NON-BLOCKING)
                        # This allows automation to continue immediately to next group
                        q.put(("status", "    -> 🚀 Iniciando atualização do SharePoint em segundo plano..."))
                        update_thread = threading.Thread(
                            target=update_sharepoint_protocol_in_thread,
                            args=(drive_id, file_id, protocol_data, q),
                            daemon=True  # Daemon thread won't block program exit
                        )
                        update_thread.start()
                        q.put(("status", "    -> ✅ Atualização em segundo plano iniciada, continuando para o próximo grupo..."))
                    elif protocol_data:
                        q.put(("status", f"    ✅ Dados do protocolo: {protocol_data}"))
                        q.put(("status", "    -> ⚠️ Atualização do SharePoint ignorada (drive_id/file_id não disponível)"))
                    else:
                        q.put(("status", "    ⚠️ Falha ao extrair protocolo"))
                
                # Navigate back to order processing page
                q.put(("status", "    -> Retornando ao processamento de pedidos..."))
                page.locator("#iframe-servico").content_frame.get_by_role("button", name="Painel").click()
                page.wait_for_timeout(2000)  # Wait for page to load
                
                # Re-enter the order processing flow
                frame = page.locator("#iframe-servico").content_frame
                frame.get_by_role("button", name="AÇÕES").click()
                page.wait_for_timeout(500)
                
                frame.get_by_role("menuitem", name="CONSUMIR ITENS").click()
                page.wait_for_timeout(2000)  # Wait for page to be ready for next group
                q.put(("status", "    -> ✅ Pronto para o próximo grupo"))
                
                page.wait_for_timeout(1000)
            except TimeoutError:
                q.put(("status", "    -> ℹ️ Nenhum botão de confirmação explícito encontrado (pode enviar automaticamente)"))
            
            # Try to proceed to next step if available
            try:
                data_button = frame.get_by_text("Data sugerida para entrega")
                data_button.wait_for(state="visible", timeout=3000)
                data_button.click()
                q.put(("status", "    -> Prosseguindo para o próximo passo (Data sugerida)"))
            except TimeoutError:
                # q.put(("status", "    -> ℹ️ Botão 'Data sugerida para entrega' não encontrado, prosseguindo..."))
                pass
            
            q.put(("status", "    ✅ Processo de envio de arquivo concluído"))
            
        except TimeoutError as e:
            q.put(("status", f"    ❌ Tempo limite de envio excedido: {e}"))
            print(f"Upload timeout: {e}")
        except Exception as e:
            q.put(("status", f"    ❌ Erro no envio: {e}"))
            print(f"Upload error: {e}")
            import traceback
            traceback.print_exc()

    else:
        q.put(("status", "    ⚠️ Nenhum item correspondente encontrado no Excel. Pulando envio."))

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
        # print(f"Headers found: {headers}")
        q.put(("status", f"    -> Colunas do Excel encontradas: {len(headers)} colunas detectadas"))
        # print(f"Column mapping: {col_mapping}")
        # print(f"Items to match: {len(items)} items")
        
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

                # print(f"  Checking item: chave={item_chave}, produto={item_produto}")
                
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
                    
                    # print(f"    -> Raw Data sugerida: {data_val}")
                    # Convert to datetime and set as date format
                    try:
                        from datetime import datetime, timedelta
                        if isinstance(data_val, (int, float)):
                            
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
                        
                        cell = ws.range(row_num, col_mapping['data_sugerida'])
                        cell.value = result_date.strftime('%m-%d-%Y')  # Store as string in YYYY-MM-DD format
                        
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
                    # print(f"    -> Set Característica da carga: {carga_val}")
                
                # Fill Demanda with chave_pedido_loja-CARRO
                if col_mapping['demanda']:
                    chave_val = matching_item.get('chave_pedido_loja')
                    carro_val = matching_item.get('carro', '')
                    demanda_val = f"{carro_val}" if carro_val else chave_val
                    ws.range(row_num, col_mapping['demanda']).value = demanda_val
                    print(f"    -> Set Demanda: {demanda_val}")
                
                # Fill Observação/ Fornecedor with chave_pedido_loja-CARRO
                if col_mapping['observacao_fornecedor']:
                    chave_val = matching_item.get('chave_pedido_loja')
                    carro_val = matching_item.get('carro', '')
                    observacao_val = f"{chave_val}-{carro_val}" if carro_val else chave_val
                    ws.range(row_num, col_mapping['observacao_fornecedor']).value = observacao_val
                    print(f"    -> Set Observação/Fornecedor: {observacao_val}")
                
                q.put(("status", f"    -> Linha {row_num}: ✅ Todos os campos preenchidos para {code_pedido}"))
                matched_count += 1
            else:
                print(f"  ⚠️ No match found for: {code_pedido} - {code_produto}")
                q.put(("status", f"    -> Linha {row_num}: ⚠️ Nenhuma correspondência para {code_pedido}"))
            
            row_num += 1
        
        print(f"✅ Matched {matched_count} out of {len(items)} items")
        q.put(("status", f"    -> Correspondidos {matched_count} de {len(items)} itens"))
        
        # Delete rows that don't have matches (in reverse to avoid index shifting)
        rows_to_delete = []
        check_row = 4
        while True:
            code_pedido = ws.range(check_row, col_mapping['codigo_pedido']).value if col_mapping['codigo_pedido'] else None
            code_produto = ws.range(check_row, col_mapping['codigo_produto']).value if col_mapping['codigo_produto'] else None
            
            if not code_pedido or not code_produto:
                break
            
            # Check if this row was matched by looking for filled quantidade_entrega
            quantidade_filled = ws.range(check_row, col_mapping['quantidade_entrega']).value if col_mapping['quantidade_entrega'] else None
            
            if quantidade_filled is None:  # Not matched
                rows_to_delete.append(check_row)
            
            check_row += 1
        
        # Delete in reverse order
        if rows_to_delete:
            for row_to_delete in reversed(rows_to_delete):
                ws.range(f"{row_to_delete}:{row_to_delete}").delete()
            
            q.put(("status", f"    -> Removidas {len(rows_to_delete)} linhas não correspondentes"))
        
        # Save the modified workbook
        wb.save()
        wb.close()
        app.quit()
        
        q.put(("status", f"    ✅ Arquivo Excel processado e salvo"))
        print(f"✅ Excel file processed and saved: {file_path}")
        return matched_count
        
    except Exception as e:
        print(f"❌ Error processing Excel file: {e}")
        q.put(("status", f"    ❌ Erro ao processar Excel: {e}"))
        import traceback
        traceback.print_exc()
        return 0



def Extrair_logs_de_upload_e_Atualizar_sharepoint(page: Page, chave: str, carro: str, q):
    """
    Extract protocol from upload logs and navigate back to order processing page.
    Returns dict with {chave, carro, protocol} or None if extraction fails.
    """
    protocol = None
    
    try:
        frame = page.locator("#iframe-servico").content_frame
        
        # Expand the first log row - look for td with aria-label="Expandir"
        q.put(("status", "    -> Expandindo primeira entrada de log..."))
        
        # Find ONLY the first td element with aria-label="Expandir" and click its child div
        expand_td = frame.locator('td[aria-label="Expandir"]').first
        
        # Make sure we're clicking only the first one
        q.put(("status", f"    -> Encontradas {frame.locator('td[aria-label=\"Expandir\"]').count()} linhas expansíveis"))
        
        expand_td.locator("div").first.click()
        page.wait_for_timeout(1500)  # Wait for expansion
        q.put(("status", "    -> Entrada de log expandida"))
        
        parent_row = expand_td.locator('xpath=ancestor::tr').first
        
       
        try:
            # First, check if there's an error message
            # Error would be in a gridcell with title containing "Erro ao gerar"
            error_cells = frame.locator('td[role="gridcell"]').filter(has_text="Erro ao gerar")
            
            if error_cells.count() > 0 and error_cells.first.is_visible():
                error_text = error_cells.first.get_attribute("title") or error_cells.first.inner_text()
                q.put(("status", f"    -> ❌ Status do envio: Erro - {error_text}"))
                # Store the error message as the protocol so we can send it to SharePoint
                protocol = f"ERRO: {error_text}"
            else:
                # No error found, look for "Demanda" message
                demanda_cells = frame.locator('td[role="gridcell"]').filter(has_text="Demanda")
                
                if demanda_cells.count() > 0 and demanda_cells.first.is_visible():
                    demanda_text = demanda_cells.first.get_attribute("title") or demanda_cells.first.inner_text()
                    q.put(("status", f"    -> ✅ Status do envio: Sucesso - {demanda_text}"))
                    
                    # Extract protocol number using regex
                    import re
                    match = re.search(r'Demanda (\d+)', demanda_text)
                    if match:
                        protocol = match.group(1)
                        q.put(("status", f"    -> ✅ Protocolo extraído: {protocol}"))
                    else:
                        q.put(("status", "    -> ⚠️ Não foi possível analisar protocolo da mensagem"))
                        protocol = "ERRO: Could not parse protocol from Demanda message"
                else:
                    q.put(("status", "    -> ⚠️ Nenhuma mensagem de Demanda ou Erro encontrada na linha expandida"))
                    protocol = "ERRO: No status message found"
                    
        except Exception as extract_err:
            q.put(("status", f"    -> ⚠️ Erro durante extração da mensagem: {extract_err}"))
            protocol = f"ERRO: Exception during extraction - {str(extract_err)}"
        
       
        # Return data with protocol (which could be success protocol number or error message)
        return {
            "chave": chave,
            "carro": carro,
            "protocol": protocol
        }
        
        
    except Exception as e:
        q.put(("status", f"    -> ❌ Erro na extração do protocolo: {e}"))
        import traceback
        traceback.print_exc()
        
        # Return error information even if extraction failed completely
        return {
            "chave": chave,
            "carro": carro,
            "protocol": f"ERRO: Critical failure - {str(e)}"
        }
    
    