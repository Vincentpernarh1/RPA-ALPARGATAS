import os
import json
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
async def update_excel_rows(drive_id: str, file_id: str, lookup_values: list[str]):
    """Legacy function - placeholder for compatibility"""
    print("⚠️ update_excel_rows called but not implemented in REST API version")
    print(f"  - Would update {len(lookup_values)} values")


async def update_protocol_async(drive_id: str, file_id: str, protocol_data_list: list[dict]):
    """Async wrapper for protocol updates - placeholder for compatibility"""
    print("⚠️ update_protocol_async called but not implemented in REST API version")
    print(f"  - Would update {len(protocol_data_list)} protocol entries")


async def update_agenda_async(drive_id: str, file_id: str, agenda_data_list: list[dict]):
    """Async wrapper for agenda updates - placeholder for compatibility"""
    print("⚠️ update_agenda_async called but not implemented in REST API version")
    print(f"  - Would update {len(agenda_data_list)} agenda entries")


# ---------------- MAIN ----------------
async def main():
    try:
        site_id = SITE_ID
        print(f"Connecting to SharePoint site: {site_id}")

        df, drive_id, file_id = await find_and_read_excel_file(site_id)

        if df is not None and not df.empty:
            print(f"✅ Successfully retrieved Excel data with {len(df)} rows")
            print(df.head())
        else:
            print("❌ Failed to retrieve Excel data")

    except Exception as ex:
        print(f"❌ Unexpected error: {ex}")


if __name__ == "__main__":
    asyncio.run(main())
