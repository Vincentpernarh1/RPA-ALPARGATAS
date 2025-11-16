# AI Copilot Instructions for RPA-ALPARGATAS

## Project Overview
**RPA-ALPARGATAS** is an automated Python-based business process automation tool that integrates web automation, SharePoint/Microsoft Graph data retrieval, and Excel processing. The system orchestrates multiple workflows through a tkinter GUI with real-time progress tracking.

## Architecture & Data Flow

### Core Components
1. **`main.py`** - GUI orchestrator
   - Tkinter UI with progress tracking and logging
   - Manages browser automation lifecycle via Playwright
   - Uses queue-based thread communication for GUI updates
   - Loads credentials from `credencial.json`

2. **`Tasks.py`** - Business logic orchestrator
   - `Order_datas_from_sharepoint()`: Fetches order data from Azure/SharePoint (async via thread)
   - `Login_and_Navigation()`: Handles web login and UI interaction with Trizy platform
   - `process_orders()`: Filters and processes orders by `chave_pedido_loja` (Order#-Store#) and product

3. **`Azure_Access.py`** - Microsoft Graph integration
   - `find_and_read_excel_file()`: Navigates SharePoint folders and reads Excel files
   - `read_excel_data()`: Extracts worksheet data into pandas DataFrames
   - `update_excel_rows()`: Updates Excel cells via REST API (async context required)

### Data Flow
```
Credentials (credencial.json)
  ↓
Main GUI → Playwright Browser + Chrome Profile
  ↓
Login_and_Navigation (Trizy platform)
  ↓
Order_datas_from_sharepoint → Azure_Access (SharePoint)
  ↓
Excel DataFrame (chave_pedido_loja grouping)
  ↓
process_orders (loop through groups/products)
  ↓
Web UI filtering + Excel updates
```

## Critical Patterns & Conventions

### Thread & Async Management
- **GUI Updates**: Use `queue.Queue` to push messages from worker threads (never update GUI directly from threads)
- **Azure Operations**: Must run in dedicated thread with `asyncio.new_event_loop()` context (see `azure_main_in_thread()`)
- **Message Types**: `("status", text)` for logging, `("progress", int)` for progress bar, `("done", bool)` to terminate update loop

### Web Automation (Playwright)
- **Profile Copying**: Uses lightweight Chrome profile copies to avoid locking active browser (`safe_copy()` with selective file copying)
- **Human-like Delays**: Implement via `human_like_delay()` to randomize timing between actions
- **Stability Waits**: Add `page.wait_for_timeout()` (1000-1500ms) after form fills before checking results - **critical for race conditions**
- **Timeouts**: Always use explicit timeout values (e.g., `timeout=5000` for Cloudflare checks, `timeout=3000` for data visibility)

### Excel & Data Processing
- **Key Column**: `chave_pedido_loja = 'Nº Pedido Cliente' + '-' + first_part_of('CÓD LOJA')`
- **Grouping**: Use `df.groupby('chave_pedido_loja')` for nested loop structure
- **File Discovery**: Look in `Dados/` folder for files matching pattern `"CARTEIRA GRUPO*.xlsx"` (skip temp files with `~$` prefix)
- **Range Updates**: Use `used_range.address` property (e.g., "Sheet1!A1:G150") to determine Excel range dynamically

### Error Handling
- Wrap async operations and platform interactions in try/except
- Queue error messages as status updates so GUI remains responsive
- Skip individual records on error (don't fail entire batch) - see `process_orders()` exception handling

## Key Files & Examples

| File | Purpose | Key Functions |
|------|---------|---|
| `main.py` | GUI + browser init | `get_playwright_browser_path()`, `run_automation()`, `load_credentials()` |
| `Tasks.py` | Order automation | `process_orders()` (nested loop with filters), `Login_and_Navigation()` (web form + frames) |
| `Azure_Access.py` | Graph API | `find_and_read_excel_file()` (path navigation), `update_excel_rows()` (batch updates) |
| `credencial.json` | Secrets | `url`, `user`, `password` for Trizy platform |
| `Dados/` | Input data | Excel files with order/product data |

## Common Tasks & Code Patterns

### Queue-based Status Update
```python
q.put(("status", "Message"))      # GUI log
q.put(("progress", 25))            # Progress bar
# GUI calls update_gui() in 100ms loop to process queue
```

### Async Excel Reading in Thread
```python
result_queue = queue.Queue()
thread = threading.Thread(target=azure_main_in_thread, args=(result_queue,))
thread.start()
thread.join()
result = result_queue.get_nowait()  # Retrieve after join()
```

### Playwright Frame Navigation
```python
frame_locator = page.locator("#iframe-servico").first.content_frame
frame_locator.locator(".dx-texteditor-input").first.fill(search_value)
page.wait_for_timeout(1500)  # Wait for UI update
frame_locator.get_by_text("Sem dados").is_visible()  # Check results
```

### Excel Range Updates (REST API)
```python
# Requires: token, used_range.address (e.g., "Sheet1!A1:G150")
endpoint = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/workbook/worksheets/{sheet_id}/range(address='{address}')"
async with aiohttp.ClientSession() as session:
    await session.patch(endpoint, headers={...}, json={"values": updated_data})
```

## Development Workflows

### Adding New Steps to Order Processing
1. Locate target action point in `process_orders()` (currently has `# [!!! ACTION NEEDED HERE !!!]` placeholder)
2. Interact with `frame_locator` (iFrame context) using Playwright locators
3. Add status updates to queue: `q.put(("status", "..."))`
4. Include waits after form interactions: `page.wait_for_timeout(1000)`
5. Wrap in try/except and add to `not_found_items` list on failure

### Debugging Playwright Issues
- Use `page.screenshot(path="debug.png")` to capture current state
- Use `page.pause()` to manually inspect page state (removes in production)
- Check timeout values: Increase if race conditions appear, decrease if tests are slow
- Verify frame selectors: `#iframe-servico` may change between environments

### Updating Azure Credentials
- Store in environment variables or `.env` file: `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`, `SITE_ID`
- Update paths in `find_and_read_excel_file()` if SharePoint folder structure changes
- Test `find_and_read_excel_file()` independently before updating production

## Environment & Dependencies
- **Python 3.9+** required for type hints and async features
- **Key Packages**: `pandas`, `openpyxl`, `playwright`, `azure-identity`, `msgraph-core`, `aiohttp`, `xlwings`
- **Playwright Setup**: Chromium browser path auto-detected or uses `sys._MEIPASS` for frozen builds
- **Credentials**: Load from `credencial.json` in execution directory

## Next Steps for Contributors
1. Review `Tasks.py:process_orders()` - understand nested loop structure and frame handling
2. Examine `Azure_Access.py:find_and_read_excel_file()` - know how to navigate SharePoint paths
3. Understand queue/thread lifecycle in `main.py` before modifying GUI or async logic
4. Test with sample Excel file in `Dados/` folder to verify data pipeline
