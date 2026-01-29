import base64
import tempfile
import os
import xlwings as xw
from concurrent.futures import ThreadPoolExecutor
import asyncio
import aiohttp
from urllib.parse import quote
from dotenv import load_dotenv
import openpyxl
from io import BytesIO
import pandas as pd
import requests
from Get_token import get_token

# Load environment variables
load_dotenv()

# SharePoint configuration
SHAREPOINT_API_BASE_URL = os.getenv("SHAREPOINT_API_BASE_URL")
SITE_ID = os.getenv("SITE_ID")


async def get_excel_data():
    """Retrieves Excel data from SharePoint using the REST API with bearer token authentication."""
    print("Accessing SharePoint using Bearer Token authentication...")

    try:
        # Load configuration from environment
        base_api_url = SHAREPOINT_API_BASE_URL
        site_id = SITE_ID

        # Hardcoded values from Azure_Access.py
        path_segments = [
            "Geral Alpargatas LLP",
            "19. Base RPA"
        ]
        file_name = "CARTEIRA GRUPO ASSAÍ.xlsx"
        folder_path = "/".join(path_segments)

        # Step 1: Get the actual drive ID from site ID
        drive_id = await get_drive_id_from_site(base_api_url, site_id)
        if not drive_id:
            print("❌ Could not retrieve drive ID")
            return []


        # Step 2: Find the folder and file
        file_id = await find_file_in_path(base_api_url, drive_id, folder_path, file_name)
        if not file_id:
            print(f"❌ File '{file_name}' not found in folder '{folder_path}'")
            return []

        # print(f"✅ Found file '{file_name}' with ID: {file_id}")

        # Step 3: Download the Excel file
        excel_bytes = await download_file_directly(base_api_url, drive_id, file_id)
        if not excel_bytes:
            print("❌ Failed to download file")
            return []

        # Step 4: Parse the Excel file
        return parse_excel_file(excel_bytes)
    except Exception as ex:
        print(f"❌ Error retrieving Excel data: {ex}")
        return []


async def get_drive_id_from_site(base_api_url: str, site_id: str) -> str | None:
    """Gets the drive ID from a site ID using the custom API"""
    # Refresh token (expires every 5 minutes)
    bearer_token = get_token()
    if not bearer_token:
        print("❌ Failed to get bearer token")
        return None

    headers = {"Authorization": f"Bearer {bearer_token}"}

    async with aiohttp.ClientSession(headers=headers) as session:
        # First, get site info which should contain drive info
        short_site_id = site_id.split(',')[1] if ',' in site_id else site_id
        sites_url = f"{base_api_url}/sites/drives/{short_site_id}"

        try:
            async with session.get(sites_url) as drives_response:
                if not drives_response.ok:
                    drives_error = await drives_response.text()
                    print(f"   Error: {drives_error[:300]}")

                if drives_response.ok:
                    drives_content = await drives_response.json()

                    # Try to parse drives array
                    if isinstance(drives_content, list):
                        drives_array = drives_content
                    elif "value" in drives_content:
                        drives_array = drives_content["value"]
                    elif "id" in drives_content:
                        drive_id = drives_content["id"]
                        return drive_id
                    else:
                        # Try lists fallback
                        return await try_lists_fallback(session, base_api_url, site_id)

                    # Get first drive
                    for drive in drives_array:
                        if "id" in drive:
                            drive_id = drive["id"]
                            return drive_id

        except Exception as e:
            print(f"   Error getting drives: {e}")

        # Fallback: Try getting lists to find Documents library
        return await try_lists_fallback(session, base_api_url, site_id)


async def try_lists_fallback(session, base_api_url: str, site_id: str) -> str | None:
    """Fallback method to get drive ID via lists endpoint"""
    try:
        lists_url = f"{base_api_url}/api/Sharepoint/sites/{site_id}/lists"

        async with session.get(lists_url) as lists_response:
            if not lists_response.ok:
                print("   Could not get lists, will try using root drive")
                return site_id

            lists_content = await lists_response.json()

            # Handle both array and object with "value" property
            if isinstance(lists_content, list):
                lists_array = lists_content
            elif "value" in lists_content:
                lists_array = lists_content["value"]
            else:
                print("   Unexpected lists response format")
                return site_id

            # Look for the "Documents" library by displayName
            for list_item in lists_array:
                if list_item.get("displayName", "").lower() == "documents":
                    return site_id  # Use siteId as driveId parameter

            # Use full site ID as the drive reference
            return "dpdhl.sharepoint.com:/teams/LLPGEHealthCare"

    except Exception as e:
        print(f"   Error in lists fallback: {e}")
        return site_id


async def find_file_in_path(base_api_url: str, drive_id: str, folder_path: str, file_name: str) -> str | None:
    """Find file using listContentFolder with path"""
    # Refresh token (expires every 5 minutes)
    bearer_token = get_token()
    if not bearer_token:
        print("❌ Failed to get bearer token")
        return None

    headers = {"Authorization": f"Bearer {bearer_token}"}

    async with aiohttp.ClientSession(headers=headers) as session:
        # Use listContentFolder endpoint to list files in the folder by path
        files_url = f"{base_api_url}/drives/listContentFolder/{quote(drive_id)}/{quote(folder_path, safe='')}"
        

        try:
            async with session.get(files_url) as files_response:
                if not files_response.ok:
                    print(f"❌ Failed to list files: {files_response.status}")
                    error = await files_response.text()
                    print(f"   Error: {error[:300]}")
                    return None

                response_data = await files_response.json()

                files = response_data if isinstance(response_data, list) else response_data.get("value", [])



                for file_item in files:
                    if file_name.lower() in file_item.get("name", "").lower():
                        return file_item.get("id")

        except Exception as e:
            print(f"   Error listing files: {e}")

    return None


async def download_file_directly(base_api_url: str, drive_id: str, file_id: str) -> bytes | None:
    """Download file directly using downloadFile endpoint"""
    # Refresh token (expires every 5 minutes)
    bearer_token = get_token()
    if not bearer_token:
        print("❌ Failed to get bearer token")
        return None

    headers = {"Authorization": f"Bearer {bearer_token}"}

    async with aiohttp.ClientSession(headers=headers) as session:
        url = f"{base_api_url}/drives/downloadFile/{drive_id}/{file_id}"

        try:
            async with session.get(url) as response:
                if not response.ok:
                    print(f"❌ Failed to download file: {response.status}")
                    error = await response.text()
                    print(f"   Error: {error}")
                    return None

                return await response.read()

        except Exception as e:
            print(f"   Error downloading file: {e}")
            return None


def parse_excel_file(excel_bytes: bytes) -> list[list[str]]:
    """Parses an Excel file from bytes and returns all data as a list of rows."""
    all_rows = []

    stream = BytesIO(excel_bytes)
    workbook = openpyxl.load_workbook(stream, data_only=True)

    worksheet = workbook.active
    if not worksheet:
        print("❌ No worksheets found in the Excel file")
        return all_rows

    print(f"✅ Reading from worksheet: '{worksheet.title}'")

    # Get the used range
    start_row = worksheet.min_row
    end_row = worksheet.max_row
    start_col = worksheet.min_column
    end_col = worksheet.max_column

    # Read all rows
    for row in range(start_row, end_row + 1):
        row_values = []

        for col in range(start_col, end_col + 1):
            cell_value = worksheet.cell(row=row, column=col).value
            cell_str = str(cell_value) if cell_value is not None else ""
            row_values.append(cell_str)

        all_rows.append(row_values)

    return all_rows


async def find_and_read_excel_file(site_id: str):
    """Main function to find and read Excel file - equivalent to Azure_Access.py version"""

    path_segments = [
        "Geral Alpargatas LLP",
        "19. Base RPA"
    ]
    file_name = "CARTEIRA GRUPO ASSAÍ.xlsx"
    folder_path = "/".join(path_segments)

    try:
        # Step 1: Get drive ID from site
        drive_id = await get_drive_id_from_site(SHAREPOINT_API_BASE_URL, site_id)
        if not drive_id:
            print("❌ Could not retrieve drive ID")
            return None, None, None

        # Step 2: Find the file
        file_id = await find_file_in_path(SHAREPOINT_API_BASE_URL, drive_id, folder_path, file_name)
        if not file_id:
            print(f"❌ File '{file_name}' not found in folder '{folder_path}'")
            return None, None, None


        # Step 3: Download the Excel file
        excel_bytes = await download_file_directly(SHAREPOINT_API_BASE_URL, drive_id, file_id)
        if not excel_bytes:
            print("❌ Failed to download Excel file")
            return None, None, None


        # Step 4: Parse the Excel file to DataFrame (matching Azure_Access.py)
        df = parse_excel_to_dataframe(excel_bytes)
        if df is not None:
            return df, drive_id, file_id
        else:
            return None, None, None

    except Exception as e:
        print(f"❌ Error retrieving Excel data: {e}")
        return None, None, None


def parse_excel_to_dataframe(excel_bytes: bytes) -> pd.DataFrame | None:
    """Parse Excel file bytes to pandas DataFrame (matching Azure_Access.py)"""
    try:
        stream = BytesIO(excel_bytes)
        workbook = openpyxl.load_workbook(stream, data_only=True)

        worksheet = workbook.active
        if not worksheet:
            print("❌ No worksheets found in the Excel file")
            return None

        # print(f"✅ Reading from worksheet: '{worksheet.title}'")

        # Get used range
        min_row = worksheet.min_row
        max_row = worksheet.max_row
        min_col = worksheet.min_column
        max_col = worksheet.max_column

        # Read all data
        data = []
        for row in range(min_row, max_row + 1):
            row_data = []
            for col in range(min_col, max_col + 1):
                cell_value = worksheet.cell(row=row, column=col).value
                row_data.append(str(cell_value) if cell_value is not None else "")

            data.append(row_data)

        if not data:
            print("  - The worksheet appears to be empty.")
            return None

        # Create DataFrame
        header = data[0]
        rows = data[1:]
        df = pd.DataFrame(rows, columns=header)
        df.replace("", None, inplace=True)
        # print("✅ Successfully created pandas DataFrame.")
        return df

    except Exception as e:
        print(f"❌ Error parsing Excel file: {e}")
        return None


# ---------------- UPDATE FUNCTIONS (REST API versions) ----------------
def update_excel_local(file_path: str, updates: list[dict]):
    """Update Excel file locally using xlwings"""
    with xw.App(visible=False) as app:
        wb = app.books.open(file_path)
        sheet = wb.sheets['Planilha1']
        for update in updates:
            if 'address' in update and 'values' in update:
                sheet.range(update['address']).value = update['values']
        wb.save()
        wb.close()


async def update_file_content(base_api_url: str, drive_id: str, folder_id: str, file_name: str, local_file_path: str):
    """Update file content using the custom API upload/small endpoint to overwrite"""
    # Refresh token (expires every 5 minutes)
    bearer_token = get_token()
    if not bearer_token:
        print("❌ Failed to get bearer token for update")
        return False

    headers = {"Authorization": f"Bearer {bearer_token}"}

    endpoint = f"{base_api_url}/upload/small"

    async with aiohttp.ClientSession(headers=headers) as session:
        try:
            with open(local_file_path, 'rb') as f:
                data = aiohttp.FormData()
                data.add_field('driveId', drive_id)
                data.add_field('parentItemId', folder_id)
                data.add_field('name', file_name)
                data.add_field('file', f, filename=file_name, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                
                async with session.post(endpoint, data=data) as response:
                    if response.ok:
                        print("✅ File updated successfully on SharePoint")
                        return True
                    else:
                        error = await response.text()
                        print(f"❌ Failed to update file: {response.status} - {error}")
                        return False
        except Exception as e:
            print(f"❌ Error updating file: {e}")
            return False


async def get_folder_id_by_path(base_api_url: str, drive_id: str, folder_path: str) -> str | None:
    """Get folder ID by path"""
    # Refresh token (expires every 5 minutes)
    bearer_token = get_token()
    if not bearer_token:
        print("❌ Failed to get bearer token")
        return None

    headers = {"Authorization": f"Bearer {bearer_token}"}

    async with aiohttp.ClientSession(headers=headers) as session:
        path_segments = folder_path.split('/')
        if not path_segments or path_segments == ['']:
            return drive_id  # Root

        parent_path = "/".join(path_segments[:-1])
        folder_name = path_segments[-1]

        if parent_path:
            files_url = f"{base_api_url}/drives/listContentFolder/{quote(drive_id)}/{quote(parent_path)}"
        else:
            files_url = f"{base_api_url}/drives/listFolder/{drive_id}"

        try:
            async with session.get(files_url) as files_response:
                if not files_response.ok:
                    print(f"❌ Failed to list parent folder: {files_response.status}")
                    error = await files_response.text()
                    print(f"   Error: {error[:300]}")
                    return None

                files_content = await files_response.json()

                files = files_content if isinstance(files_content, list) else files_content.get("value", [])

                for item in files:
                    if item.get("name", "").lower() == folder_name.lower():
                        return item.get("id")

                print(f"❌ Folder '{folder_name}' not found in parent")
                return None

        except Exception as e:
            print(f"   Error getting folder ID: {e}")
            return None


async def update_excel_rows(drive_id: str, file_id: str, lookup_values: list[str]):
    """Update lookup values in the Excel file (assumes Planilha1!A column starting from row 2)"""
    base_api_url = "https://api-storage.connectedcontroltower.com.br/api/Sharepoint"
    
    # Download the file
    excel_bytes = await download_file_directly(base_api_url, drive_id, file_id)
    if not excel_bytes:
        print("❌ Failed to download file for update")
        return False
    
    # Save to temp file
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
        temp_file.write(excel_bytes)
        temp_path = temp_file.name
    
    try:
        # Prepare updates
        updates = [{'address': f'A{i+2}', 'values': value} for i, value in enumerate(lookup_values)]
        
        # Update locally in thread
        loop = asyncio.get_event_loop()
        with ThreadPoolExecutor() as executor:
            await loop.run_in_executor(executor, update_excel_local, temp_path, updates)
        
        # Get folder ID
        folder_path = "Geral Alpargatas LLP/19. Base RPA"
        folder_id = await get_folder_id_by_path(base_api_url, drive_id, folder_path)
        if not folder_id:
            print("❌ Failed to get folder ID")
            return False
        
        # Upload back
        file_name = "CARTEIRA GRUPO ASSAÍ.xlsx"
        success = await update_file_content(base_api_url, drive_id, folder_id, file_name, temp_path)
        return success
    finally:
        os.unlink(temp_path)


async def update_protocol_async(drive_id: str, file_id: str, protocol_data_list: list[dict]):
    """Update protocol values in the Excel file (Planilha1!CK column)"""
    base_api_url = "https://api-storage.connectedcontroltower.com.br/api/Sharepoint"
    
    # Download the file
    excel_bytes = await download_file_directly(base_api_url, drive_id, file_id)
    if not excel_bytes:
        print("❌ Failed to download file for update")
        return False
    
    # Save to temp file
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
        temp_file.write(excel_bytes)
        temp_path = temp_file.name
    
    try:
        # Prepare updates
        updates = [{'address': f'CK{item["row"]}', 'values': item['value']} for item in protocol_data_list if 'row' in item and 'value' in item]
        
        # Update locally in thread
        loop = asyncio.get_event_loop()
        with ThreadPoolExecutor() as executor:
            await loop.run_in_executor(executor, update_excel_local, temp_path, updates)
        
        # Get folder ID
        folder_path = "Geral Alpargatas LLP/19. Base RPA"
        folder_id = await get_folder_id_by_path(base_api_url, drive_id, folder_path)
        if not folder_id:
            print("❌ Failed to get folder ID")
            return False
        
        # Upload back
        file_name = "CARTEIRA GRUPO ASSAÍ.xlsx"
        success = await update_file_content(base_api_url, drive_id, folder_id, file_name, temp_path)
        return success
    finally:
        os.unlink(temp_path)


async def update_agenda_async(drive_id: str, file_id: str, agenda_data_list: list[dict]):
    """Update agenda values in the Excel file (Planilha1!CJ column)"""
    base_api_url = "https://api-storage.connectedcontroltower.com.br/api/Sharepoint"
    
    # Download the file
    excel_bytes = await download_file_directly(base_api_url, drive_id, file_id)
    if not excel_bytes:
        print("❌ Failed to download file for update")
        return False
    
    # Save to temp file
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
        temp_file.write(excel_bytes)
        temp_path = temp_file.name
    
    try:
        # Prepare updates
        updates = [{'address': f'CJ{item["row"]}', 'values': item['value']} for item in agenda_data_list if 'row' in item and 'value' in item]
        
        # Update locally in thread
        loop = asyncio.get_event_loop()
        with ThreadPoolExecutor() as executor:
            await loop.run_in_executor(executor, update_excel_local, temp_path, updates)
        
        # Get folder ID
        folder_path = "Geral Alpargatas LLP/19. Base RPA"
        folder_id = await get_folder_id_by_path(base_api_url, drive_id, folder_path)
        if not folder_id:
            print("❌ Failed to get folder ID")
            return False
        
        # Upload back
        file_name = "CARTEIRA GRUPO ASSAÍ.xlsx"
        success = await update_file_content(base_api_url, drive_id, folder_id, file_name, temp_path)
        return success
    finally:
        os.unlink(temp_path)


# ---------------- MAIN ----------------
async def main():
    try:
        site_id = SITE_ID
        # print(f"Connecting to SharePoint site: {site_id}")

        df, drive_id, file_id = await find_and_read_excel_file(site_id)
        
        # print(df.head(10))
        
        return df, drive_id, file_id

    except Exception as ex:
        print(f"❌ Unexpected error: {ex}")
        return None, None, None


# if __name__ == "__main__":
#     asyncio.run(main())
