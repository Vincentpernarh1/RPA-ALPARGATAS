import asyncio
import os
import pandas as pd
from dotenv import load_dotenv
import aiohttp

from azure.identity.aio import ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from msgraph.generated.sites.item.drives.drives_request_builder import DrivesRequestBuilder
from msgraph.generated.drives.item.items.item.workbook.worksheets.item.used_range.used_range_request_builder import UsedRangeRequestBuilder

# Load environment variables
load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SITE_ID = os.getenv("SITE_ID")


def get_graph_client() -> GraphServiceClient:
    print("Creating graph client...")
    credential = ClientSecretCredential(
        tenant_id=TENANT_ID,
        client_id=CLIENT_ID,
        client_secret=CLIENT_SECRET
    )
    return GraphServiceClient(credentials=credential, scopes=['https://graph.microsoft.com/.default'])


async def read_excel_data(graph_client: GraphServiceClient, drive_id: str, file_id: str) -> pd.DataFrame | None:
    print("\nReading data from the first worksheet...")

    try:
        worksheets = await graph_client.drives.by_drive_id(drive_id).items.by_drive_item_id(file_id).workbook.worksheets.get()
        if not worksheets or not worksheets.value:
            print("❌ No worksheets found in the Excel file.")
            return None

        first_worksheet = worksheets.value[0]
        print(f"  - Reading from worksheet: '{first_worksheet.name}'")

        used_range = await graph_client.drives.by_drive_id(drive_id) \
            .items.by_drive_item_id(file_id) \
            .workbook.worksheets.by_workbook_worksheet_id(first_worksheet.id) \
            .used_range.get()

        values = None
        if "values" in used_range.additional_data:
            values = used_range.additional_data["values"]
        elif "text" in used_range.additional_data:
            values = used_range.additional_data["text"]

        if values:
            header = values[0]
            data = values[1:]
            df = pd.DataFrame(data, columns=header)
            df.replace("", None, inplace=True)
            print("✅ Successfully created pandas DataFrame.")
            return df
        else:
            print("  - The worksheet appears to be empty.")
            return None

    except ODataError as e:
        print(f"❌ Error reading Excel data: {e.error.message}")
        return None


async def find_and_read_excel_file(graph_client: GraphServiceClient, site_id: str):
    print("\nAttempting to find and read 'Devolução de Notas.xlsx'...")

    path_segments = [
        "Geral Alpargatas LLP",
        "19. Base RPA"
    ]
    file_name = "CARTEIRA GRUPO ASSAÍ.xlsx"
    document_library_name = "Documents"

    try:
        query_params = DrivesRequestBuilder.DrivesRequestBuilderGetQueryParameters(expand=["root"])
        request_config = DrivesRequestBuilder.DrivesRequestBuilderGetRequestConfiguration(query_parameters=query_params)
        drives = await graph_client.sites.by_site_id(site_id).drives.get(request_configuration=request_config)

        if not drives or not drives.value:
            print("❌ No document libraries found.")
            return None

        target_drive = next((d for d in drives.value if d.name and d.name.lower() == document_library_name.lower()), None)
        if not target_drive or not target_drive.id:
            print("❌ Target drive not found.")
            return None

        drive_id = target_drive.id
        current_folder_id = target_drive.root.id

        for segment in path_segments:
            print(f"  -> Searching folder: {segment}")
            children = await graph_client.drives.by_drive_id(drive_id).items.by_drive_item_id(current_folder_id).children.get()
            found_folder = next((item for item in children.value if item.name.lower() == segment.lower()), None)
            if found_folder:
                current_folder_id = found_folder.id
            else:
                print(f"❌ Folder '{segment}' not found.")
                return None

        print(f"  -> Searching for file '{file_name}'...")
        final_children = await graph_client.drives.by_drive_id(drive_id).items.by_drive_item_id(current_folder_id).children.get()
        excel_file = next((item for item in final_children.value if item.name.lower() == file_name.lower()), None)

        if excel_file:
            print(f"✅ Found file: {excel_file.name}")
            df = await read_excel_data(graph_client, drive_id, excel_file.id)
            return df, drive_id, excel_file.id
        else:
            print("❌ File not found.")
            return None, None, None

    except ODataError as e:
        print(f"❌ API error: {e.error.message}")
        return None, None, None



# ---------------- UPDATE FUNCTION ----------------
async def update_protocol_rows(graph_client: GraphServiceClient, drive_id: str, file_id: str, protocol_data_list: list[dict]):
    """
    Updates rows in Excel with protocol numbers where chave_pedido_loja + CARRO matches.
    
    Args:
        graph_client: GraphServiceClient instance
        drive_id: SharePoint drive ID
        file_id: Excel file ID
        protocol_data_list: List of dicts with {chave, carro, protocol}
    
    Updates column 'PROTOCOLO DA SOLICITAÇÃO' (column BL) where:
        - chave_pedido_loja = Nº Pedido Cliente + '-' + first_part_of(CÓD LOJA)
        - CARRO matches
    """
    print("\n🔄 Starting protocol update process...")
    print(f"  - Protocol data to update: {protocol_data_list}")

    try:
        worksheets = await graph_client.drives.by_drive_id(drive_id) \
            .items.by_drive_item_id(file_id).workbook.worksheets.get()
        if not worksheets or not worksheets.value:
            print("❌ No worksheets found.")
            return

        first_worksheet = worksheets.value[0]
        print(f"  - Target worksheet: {first_worksheet.name}")

        used_range = await graph_client.drives.by_drive_id(drive_id) \
            .items.by_drive_item_id(file_id) \
            .workbook.worksheets.by_workbook_worksheet_id(first_worksheet.id) \
            .used_range.get()

        if "values" not in used_range.additional_data:
            print("❌ No data found to update.")
            return

        if not used_range.address:
            print("❌ Could not determine range address for update.")
            return

        values = used_range.additional_data["values"]
        header, data = values[0], values[1:]

        # Find column indices
        col_indices = {}
        for idx, col_name in enumerate(header):
            col_name_lower = str(col_name).lower().strip() if col_name else ""
            if 'nº pedido cliente' in col_name_lower or 'pedido cliente' in col_name_lower:
                col_indices['pedido_cliente'] = idx
            elif 'cód loja' in col_name_lower or 'cod loja' in col_name_lower:
                col_indices['cod_loja'] = idx
            elif 'carro' in col_name_lower:
                col_indices['carro'] = idx
            elif 'protocolo' in col_name_lower and 'solicitação' in col_name_lower:
                col_indices['protocolo'] = idx
        
        # print(f"  - Found column indices: {col_indices}")
        
        # Verify required columns exist
        required_cols = ['pedido_cliente', 'cod_loja', 'carro', 'protocolo']
        missing_cols = [col for col in required_cols if col not in col_indices]
        if missing_cols:
            print(f"❌ Missing required columns: {missing_cols}")
            return
        
        # Update matching rows
        updated_count = 0
        for row in data:
            if len(row) <= max(col_indices.values()):
                # Extend row if needed
                row.extend([""] * (max(col_indices.values()) + 1 - len(row)))
            
            # Build chave_pedido_loja from row data
            pedido_val = str(row[col_indices['pedido_cliente']]).strip() if row[col_indices['pedido_cliente']] else ""
            loja_val = str(row[col_indices['cod_loja']]).strip() if row[col_indices['cod_loja']] else ""
            carro_val = str(row[col_indices['carro']]).strip() if row[col_indices['carro']] else ""
            
            # Extract first part of loja (before '-')
            loja_first_part = loja_val.split('-')[0] if loja_val else ""
            row_chave = f"{pedido_val}-{loja_first_part}" if pedido_val and loja_first_part else ""
            
            # Check if this row matches any protocol data
            for protocol_item in protocol_data_list:
                item_chave = str(protocol_item.get('chave', '')).strip()
                item_carro = str(protocol_item.get('carro', '')).strip()
                item_protocol = str(protocol_item.get('protocol', '')).strip()
                
                if row_chave == item_chave and carro_val == item_carro:
                    # Update protocol column
                    row[col_indices['protocolo']] = item_protocol
                    updated_count += 1
                    print(f"  - Updated row: {row_chave} / {carro_val} -> Protocol: {item_protocol}")
                    break

        updated_values = [header] + data
        print(f"  - Prepared {updated_count} protocol updates to send...")

        if updated_count == 0:
            print("  - No matching rows found to update.")
            return

        # === Direct REST call using same credentials ===
        credential = ClientSecretCredential(
            tenant_id=os.getenv("TENANT_ID"),
            client_id=os.getenv("CLIENT_ID"),
            client_secret=os.getenv("CLIENT_SECRET")
        )
        token_response = await credential.get_token("https://graph.microsoft.com/.default")
        token = token_response.token
        
        target_address = used_range.address
        print(f"  - Updating range: {target_address}")

        endpoint = (
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}"
            f"/workbook/worksheets/{first_worksheet.id}/range(address='{target_address}')"
        )

        async with aiohttp.ClientSession() as session:
            async with session.patch(
                endpoint,
                headers={
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/json",
                },
                json={"values": updated_values},
            ) as resp:
                if resp.status == 200:
                    print(f"✅ Successfully updated {updated_count} protocol rows in Excel.")
                else:
                    text = await resp.text()
                    print(f"❌ Protocol update failed ({resp.status}): {text}")

    except Exception as ex:
        print(f"❌ Unexpected error during protocol update: {ex}")
        import traceback
        traceback.print_exc()


# Keep old function name for backward compatibility (redirects to new function)
async def update_excel_rows(graph_client: GraphServiceClient, drive_id: str, file_id: str, lookup_values: list[str]):
    """Legacy function - redirects to update_protocol_rows with old behavior"""
    print("⚠️ Using legacy update_excel_rows function")
    # This maintains old behavior if called elsewhere
    protocol_data = [{"chave": val, "carro": "", "protocol": "Não Encontrado"} for val in lookup_values]
    await update_protocol_rows(graph_client, drive_id, file_id, protocol_data)


# ---------------- UPDATE AGENDA COLUMNS (NEW FOR PEGAR RETORNO) ----------------
async def update_agenda_columns(graph_client: GraphServiceClient, drive_id: str, file_id: str, agenda_data_list: list[dict]):
    """
    Updates the 'AGENDA CONFIRMADA', 'PROTOCOLO AGENDA', and 'HORÁRIO' columns in SharePoint Excel with extracted data.
    
    Args:
        graph_client: Microsoft Graph client instance
        drive_id: SharePoint drive ID
        file_id: Excel file ID
        agenda_data_list: List of dicts with structure:
            [{"chave": "12345-67", "protocol": "P001", "agenda_confirmada": "12/08/2025", "protocolo_agenda": "Agendamento12345", "horario": "14:30"}, ...]
    
    Matching logic:
        - Matches rows where chave_pedido_loja == chave AND protocol == protocol
        - Updates the respective columns with the provided values
    """
    try:
        print("\n🔄 Starting agenda columns update process...")
        print(f"  - Processing {len(agenda_data_list)} agenda updates")

        worksheets = await graph_client.drives.by_drive_id(drive_id).items.by_drive_item_id(file_id).workbook.worksheets.get()
        if not worksheets or not worksheets.value:
            print("❌ No worksheets found.")
            return

        first_worksheet = worksheets.value[0]
        print(f"  - Target worksheet: {first_worksheet.name}")

        used_range = await graph_client.drives.by_drive_id(drive_id) \
            .items.by_drive_item_id(file_id) \
            .workbook.worksheets.by_workbook_worksheet_id(first_worksheet.id) \
            .used_range.get()

        if "values" not in used_range.additional_data:
            print("❌ No data found to update.")
            return

        if not used_range.address:
            print("❌ Could not determine range address for update.")
            return

        values = used_range.additional_data["values"]
        header, data = values[0], values[1:]

        # Find column indices
        col_indices = {}
        for idx, col_name in enumerate(header):
            col_name_lower = str(col_name).lower().strip() if col_name else ""
            if 'nº pedido cliente' in col_name_lower or 'pedido cliente' in col_name_lower:
                col_indices['pedido_cliente'] = idx
            elif 'cód loja' in col_name_lower or 'cod loja' in col_name_lower:
                col_indices['cod_loja'] = idx
            elif 'protocolo' in col_name_lower and 'solicitação' in col_name_lower:
                col_indices['protocolo'] = idx
            elif 'agenda confirmada' in col_name_lower:
                col_indices['agenda_confirmada'] = idx
            elif 'protocolo agenda' in col_name_lower:
                col_indices['protocolo_agenda'] = idx
            elif 'horário' in col_name_lower or 'horario' in col_name_lower:
                col_indices['horario'] = idx
        
        # print(f"  - Found column indices: {col_indices}")
        
        # Verify required columns exist
        required_cols = ['pedido_cliente', 'cod_loja', 'protocolo', 'agenda_confirmada', 'protocolo_agenda', 'horario']
        missing_cols = [col for col in required_cols if col not in col_indices]
        if missing_cols:
            print(f"❌ Missing required columns: {missing_cols}")
            return
        
        # Update matching rows
        updated_count = 0
        for row in data:
            if len(row) <= max(col_indices.values()):
                # Extend row if needed
                row.extend([""] * (max(col_indices.values()) + 1 - len(row)))
            
            # Build chave_pedido_loja from row data
            pedido_val = str(row[col_indices['pedido_cliente']]).strip() if row[col_indices['pedido_cliente']] else ""
            loja_val = str(row[col_indices['cod_loja']]).strip() if row[col_indices['cod_loja']] else ""
            protocol_val = str(row[col_indices['protocolo']]).strip() if row[col_indices['protocolo']] else ""
            
            # Extract first part of loja (before '-')
            loja_first_part = loja_val.split('-')[0] if loja_val else ""
            row_chave = f"{pedido_val}-{loja_first_part}" if pedido_val and loja_first_part else ""
            
            # Check if this row matches any agenda data
            for agenda_item in agenda_data_list:
                item_chave = str(agenda_item.get('chave', '')).strip()
                item_protocol = str(agenda_item.get('protocol', '')).strip()
                item_agenda_confirmada = str(agenda_item.get('agenda_confirmada', '')).strip()
                item_protocolo_agenda = str(agenda_item.get('protocolo_agenda', '')).strip()
                item_horario = str(agenda_item.get('horario', '')).strip()
                
                if row_chave == item_chave and protocol_val == item_protocol:
                    # Update the three columns
                    row[col_indices['agenda_confirmada']] = item_agenda_confirmada
                    row[col_indices['protocolo_agenda']] = item_protocolo_agenda
                    row[col_indices['horario']] = item_horario
                    updated_count += 1
                    # print(f"  - Updated row: {row_chave} / {item_protocol} -> Agenda Confirmada: {item_agenda_confirmada}, Protocolo Agenda: {item_protocolo_agenda}, Horário: {item_horario}")
                    break

        updated_values = [header] + data
        print(f"  - Prepared {updated_count} agenda updates to send...")

        if updated_count == 0:
            print("  - No matching rows found to update.")
            return

        # === Direct REST call using same credentials ===
        credential = ClientSecretCredential(
            tenant_id=os.getenv("TENANT_ID"),
            client_id=os.getenv("CLIENT_ID"),
            client_secret=os.getenv("CLIENT_SECRET")
        )
        token_response = await credential.get_token("https://graph.microsoft.com/.default")
        token = token_response.token
        
        target_address = used_range.address
        print(f"  - Updating range: {target_address}")

        endpoint = (
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}"
            f"/workbook/worksheets/{first_worksheet.id}/range(address='{target_address}')"
        )

        async with aiohttp.ClientSession() as session:
            async with session.patch(
                endpoint,
                headers={
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/json",
                },
                json={"values": updated_values},
            ) as resp:
                if resp.status == 200:
                    print(f"✅ Successfully updated {updated_count} agenda columns in Excel.")
                else:
                    text = await resp.text()
                    print(f"❌ Agenda update failed ({resp.status}): {text}")

    except Exception as ex:
        print(f"❌ Unexpected error during agenda update: {ex}")
        import traceback
        traceback.print_exc()



# ---------------- ASYNC WRAPPER FOR PROTOCOL UPDATE ----------------
async def update_protocol_async(drive_id: str, file_id: str, protocol_data_list: list[dict]):
    """
    Async wrapper to update protocols in SharePoint Excel.
    Can be called from a background thread.
    
    Args:
        drive_id: SharePoint drive ID
        file_id: Excel file ID  
        protocol_data_list: List of dicts with {chave, carro, protocol}
    """
    graph_client = get_graph_client()
    try:
        await update_protocol_rows(graph_client, drive_id, file_id, protocol_data_list)
        print("✅ Protocol update completed successfully")
    except Exception as e:
        print(f"❌ Protocol update failed: {e}")
        import traceback
        traceback.print_exc()


# ---------------- ASYNC WRAPPER FOR AGENDA UPDATE ----------------
async def update_agenda_async(drive_id: str, file_id: str, agenda_data_list: list[dict]):
    """
    Async wrapper to update agenda columns in SharePoint Excel.
    Can be called from a background thread.
    
    Args:
        drive_id: SharePoint drive ID
        file_id: Excel file ID
        agenda_data_list: List of dicts with {chave, protocol, agenda_confirmada, protocolo_agenda, horario}
    """
    graph_client = get_graph_client()
    try:
        await update_agenda_columns(graph_client, drive_id, file_id, agenda_data_list)
        print("✅ Agenda update completed successfully")
    except Exception as e:
        print(f"❌ Agenda update failed: {e}")
        import traceback
        traceback.print_exc()


# ---------------- MAIN ----------------
async def main():
    site_id = SITE_ID
    graph_client = get_graph_client()

    try:
        print(f"Connecting to SharePoint site: {site_id}")
        site = await graph_client.sites.by_site_id(site_id).get()
        print(f"✅ Connected to site: {site.name}")

        df, drive_id, file_id = await find_and_read_excel_file(graph_client, site_id)

        if df is not None and not df.empty:
            # print(f"\n--- DataFrame Info ---")
            # print(f"Shape: {df.shape}")
            # print(df.head(2))

            # 🔹 Example usage: Update protocols
            # protocol_data = [{"chave": "12345-67", "carro": "CARRO1", "protocol": "98765"}]
            # await update_protocol_rows(graph_client, drive_id, file_id, protocol_data)

            return df, drive_id, file_id 

        else:
            print("⚠️ Could not read any data.")

    except ODataError as e:
        print(f"❌ Operation failed: {e.error.message}")
    except Exception as ex:
        print(f"❌ Unexpected error: {ex}")


# if __name__ == "__main__":
#     asyncio.run(main())