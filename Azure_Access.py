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
async def update_excel_rows(graph_client: GraphServiceClient, drive_id: str, file_id: str, lookup_values: list[str]):
    """
    Updates rows in Excel where the value in column 2 matches any in lookup_values,
    setting column 3 to 'Não Encontrado'.
    """
    print("\n🔄 Starting update process...")

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

        # NEW: Check that the used_range object has its address property
        if not used_range.address:
            print("❌ Could not determine range address for update.")
            return

        values = used_range.additional_data["values"]
        header, data = values[0], values[1:]

        updated_count = 0
        for row in data:
            if len(row) > 1 and str(row[0]).strip() in lookup_values:
                if len(row) < 3:
                    row.extend([""] * (3 - len(row)))
                row[2] = "Não Encontrado"
                updated_count += 1

        updated_values = [header] + data
        print(f"  - Prepared {updated_count} updates to send...")

        # === Direct REST call using same credentials ===
        
        credential = ClientSecretCredential(
            tenant_id=os.getenv("TENANT_ID"),
            client_id=os.getenv("CLIENT_ID"),
            client_secret=os.getenv("CLIENT_SECRET")
        )
        token_response = await credential.get_token("https://graph.microsoft.com/.default")
        token = token_response.token
        
        # CORRECTED: Get the full address (e.g., "Sheet1!A1:G150") from the range
        target_address = used_range.address
        print(f"  - Updating range: {target_address}")

        endpoint = (
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}"
            # CORRECTED: Use the dynamic target_address instead of hardcoded 'A1'
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
                    print(f"✅ Successfully updated {updated_count} rows in Excel.")
                else:
                    text = await resp.text()
                    print(f"❌ Update failed ({resp.status}): {text}")

    except Exception as ex:
        print(f"❌ Unexpected error: {ex}")





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

            # 🔹 Example usage: Update rows where column 2 has "20751764"
            # await update_excel_rows(graph_client, drive_id, file_id, ["20751764"])

            return df, drive_id, file_id 

        else:
            print("⚠️ Could not read any data.")

    except ODataError as e:
        print(f"❌ Operation failed: {e.error.message}")
    except Exception as ex:
        print(f"❌ Unexpected error: {ex}")


# if __name__ == "__main__":
#     asyncio.run(main())