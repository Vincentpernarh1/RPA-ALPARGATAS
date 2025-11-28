# RPA-ALPARGATAS

**Automated Business Process Automation Tool for Order Management**

[![Python](https://img.shields.io/badge/Python-3.9%2B-blue.svg)](https://www.python.org/)
[![Playwright](https://img.shields.io/badge/Playwright-1.40%2B-green.svg)](https://playwright.dev/)
[![License](https://img.shields.io/badge/License-Proprietary-red.svg)]()

## 📋 Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Architecture](#architecture)
- [Project Structure](#project-structure)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [Workflow Details](#workflow-details)
- [Data Flow](#data-flow)
- [Technical Specifications](#technical-specifications)
- [Development Guide](#development-guide)
- [Troubleshooting](#troubleshooting)
- [Future Enhancements](#future-enhancements)
- [Contributors](#contributors)

---

## 🎯 Overview

**RPA-ALPARGATAS** is a Python-based Robotic Process Automation (RPA) solution designed to streamline order processing workflows between SharePoint/Microsoft Graph and the Trizy platform. The system automates:

- **Data retrieval** from SharePoint Excel files via Microsoft Graph API
- **Web automation** for order entry and processing on the Trizy platform
- **Excel manipulation** with order details, delivery dates, and shipping characteristics
- **File upload** back to the web platform with processed data

The application features a modern **Tkinter GUI** with real-time progress tracking, logging, and DHL/STELLANTIS-themed branding.

---

## ✨ Features

### Core Capabilities

- ✅ **SharePoint Integration**: Secure OAuth authentication with Microsoft Graph API
- ✅ **Intelligent Order Grouping**: Groups orders by `chave_pedido_loja` (Order#-Store#)
- ✅ **Web Automation**: Playwright-based browser automation with human-like behavior
- ✅ **Excel Processing**: Automated reading, writing, and formatting with xlwings
- ✅ **Real-time GUI**: Progress tracking, status updates, and activity logging
- ✅ **Error Handling**: Comprehensive exception handling with detailed logging
- ✅ **Profile Management**: Safe Chrome profile copying for persistent sessions

### Technical Highlights

- **Asynchronous Operations**: Runs Azure/SharePoint operations in separate threads
- **Queue-based Communication**: Thread-safe GUI updates via queue messaging
- **Human-like Delays**: Randomized timing to mimic natural user behavior
- **Cloudflare Detection**: Handles verification challenges automatically
- **Dynamic Excel Mapping**: Flexible column detection for varying file formats

---

## 🏗️ Architecture

### Component Overview

```
┌─────────────────────────────────────────────────────────────┐
│                        main.py                              │
│  ┌────────────────────────────────────────────────────┐    │
│  │  Tkinter GUI (Status, Progress, Logs)              │    │
│  └────────────────────────────────────────────────────┘    │
│           │                                                  │
│           ▼                                                  │
│  ┌────────────────────────────────────────────────────┐    │
│  │  Playwright Browser Context                        │    │
│  │  - Chrome Profile Copy                             │    │
│  │  - Persistent Context                              │    │
│  └────────────────────────────────────────────────────┘    │
└─────────────────┬───────────────────────────────────────────┘
                  │
                  ▼
┌─────────────────────────────────────────────────────────────┐
│                      Tasks.py                               │
│  ┌────────────────────────────────────────────────────┐    │
│  │  Login_and_Navigation()                            │    │
│  │  - Web login with human-like typing                │    │
│  │  - Cloudflare bypass                               │    │
│  │  - Navigate to order management                    │    │
│  └────────────────────────────────────────────────────┘    │
│  ┌────────────────────────────────────────────────────┐    │
│  │  process_orders()                                  │    │
│  │  - Loop through order groups                       │    │
│  │  - Filter by product                               │    │
│  │  - Select checkboxes                               │    │
│  └────────────────────────────────────────────────────┘    │
│  ┌────────────────────────────────────────────────────┐    │
│  │  processar_e_Fazer_upload_Arquivos()               │    │
│  │  - Download template                               │    │
│  │  - Process with xlwings                            │    │
│  │  - Upload completed file                           │    │
│  └────────────────────────────────────────────────────┘    │
└─────────────────┬───────────────────────────────────────────┘
                  │
                  ▼
┌─────────────────────────────────────────────────────────────┐
│                   Azure_Access.py                           │
│  ┌────────────────────────────────────────────────────┐    │
│  │  find_and_read_excel_file()                        │    │
│  │  - Navigate SharePoint folders                     │    │
│  │  - Locate Excel file                               │    │
│  └────────────────────────────────────────────────────┘    │
│  ┌────────────────────────────────────────────────────┐    │
│  │  read_excel_data()                                 │    │
│  │  - Extract worksheet data                          │    │
│  │  - Convert to pandas DataFrame                     │    │
│  └────────────────────────────────────────────────────┘    │
│  ┌────────────────────────────────────────────────────┐    │
│  │  update_excel_rows() [Future Implementation]      │    │
│  │  - Batch update Excel cells via REST API          │    │
│  └────────────────────────────────────────────────────┘    │
└─────────────────────────────────────────────────────────────┘
```

---

## 📁 Project Structure

```
RPA-ALPARGATAS/
│
├── main.py                    # GUI orchestrator & browser initialization
├── Tasks.py                   # Business logic & web automation
├── Azure_Access.py            # Microsoft Graph API integration
│
├── credencial.json            # Trizy platform credentials
├── static_data.json           # Static shipping configuration
│
├── Dados/                     # Input data folder
│   └── CARTEIRA GRUPO*.xlsx   # Order/product Excel files
│
├── Arquivos/                  # Output/download folder
│   └── YYYY-MM-DD.xlsx        # Processed upload files
│
├── __pycache__/              # Python bytecode cache
│
├── .env                       # Azure credentials (not in repo)
│   ├── TENANT_ID
│   ├── CLIENT_ID
│   ├── CLIENT_SECRET
│   └── SITE_ID
│
└── README.md                  # This documentation
```

---

## 🔧 Installation

### Prerequisites

- **Python**: 3.9 or higher
- **Windows OS**: Required for xlwings Excel automation
- **Chrome Browser**: For profile copying and Playwright automation
- **Azure App Registration**: For Microsoft Graph API access

### Step 1: Clone Repository

```powershell
git clone https://github.com/Vincentpernarh1/RPA-ALPARGATAS.git
cd RPA-ALPARGATAS
```

### Step 2: Create Virtual Environment

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
```

### Step 3: Install Dependencies

```powershell
pip install -r requirements.txt
```

**Core Dependencies:**
```
pandas
openpyxl
playwright
azure-identity
msgraph-core
aiohttp
xlwings
python-dotenv
pyxlsb
```

### Step 4: Install Playwright Browsers

```powershell
playwright install chromium
```

### Step 5: Verify Installation

```powershell
python -c "import playwright; print('✅ Playwright installed')"
python -c "import xlwings; print('✅ xlwings installed')"
```

---

## 📦 Building Executable

### Creating Standalone Executable with PyInstaller

To distribute the application as a standalone executable without requiring Python installation:

#### Prerequisites

```powershell
pip install pyinstaller
```

#### Build Command

```powershell
pyinstaller --noconfirm --onefile --windowed --noconsole --name "RPA Solicitar Pedidos" --icon "C:/Users/perna/Desktop/Barrueri/Alpargatas/Pedido-icon.ico" --add-data "C:\Users\perna\AppData\Local\ms-playwright\chromium-1187\chrome-win;ms-playwright\chromium-1187\chrome-win" main.py
```

#### Build Parameters Explained

| Parameter | Purpose |
|-----------|---------|
| `--noconfirm` | Overwrite output directory without confirmation |
| `--onefile` | Bundle everything into a single executable |
| `--windowed` | GUI application (no console window) |
| `--noconsole` | Suppress console window on startup |
| `--name "RPA Solicitar Pedidos"` | Set executable name |
| `--icon "Pedido-icon.ico"` | Application icon (must be .ico format) |
| `--add-data` | Include Playwright Chromium binaries |

#### Output Location

```
dist/
└── RPA Solicitar Pedidos.exe    # Standalone executable
```

#### Important Notes

- **Chromium Path**: Update the `--add-data` path to match your Playwright installation:
  ```powershell
  # Find your Chromium path:
  python -c "from playwright.driver import compute_driver_executable; print(compute_driver_executable())"
  ```

- **File Dependencies**: Ensure these files are in the same directory as the executable:
  - `credencial.json`
  - `static_data.json`
  - `.env` (if using Azure features)

- **First Run**: The executable may take 10-20 seconds to start on first launch

#### Distribution Checklist

Before distributing the executable:

1. ✅ Test on a clean Windows machine without Python
2. ✅ Include `credencial.json` template (remove sensitive data)
3. ✅ Include `static_data.json`
4. ✅ Create `Dados/` and `Arquivos/` folders
5. ✅ Provide user documentation for configuration

---

## ⚙️ Configuration

### 1. Azure App Registration

Create an Azure AD app with the following API permissions:

| API | Permission | Type |
|-----|-----------|------|
| Microsoft Graph | `Sites.Read.All` | Application |
| Microsoft Graph | `Files.Read.All` | Application |
| Microsoft Graph | `Files.ReadWrite.All` | Application |

### 2. Environment Variables (`.env`)

Create a `.env` file in the project root:

```env
TENANT_ID=your-tenant-id-here
CLIENT_ID=your-client-id-here
CLIENT_SECRET=your-client-secret-here
SITE_ID=your-sharepoint-site-id-here
```

**How to find Site ID:**
```powershell
# Navigate to your SharePoint site and run:
https://graph.microsoft.com/v1.0/sites/{your-domain}.sharepoint.com:/sites/{site-name}
```

### 3. Credentials File (`credencial.json`)

```json
{
    "user": "your-email@domain.com",
    "password": "your-password",
    "url": "https://login.trizy.com.br/access/auth/login/"
}
```

⚠️ **Security Warning**: Never commit this file to version control. Add to `.gitignore`.

### 4. Static Data File (`static_data.json`)

```json
{
  "caracteristica": "Paletizada",
  "caracteristica_do_veiculo": "Carreta Simples Toco (3 eixos) 25T"
}
```

This defines shipping characteristics applied to all orders.

### 5. Input Data Folder

Place your order Excel file in `Dados/`:
- File name pattern: `CARTEIRA GRUPO*.xlsx`
- Avoid temporary Excel files (starting with `~$`)

**Required Excel Columns:**
- `Nº Pedido Cliente`
- `CÓD LOJA`
- `PRODUTO INTERNO CLIENTE`
- `Qtd. Faturada`
- `PREVISÃO DE ENTREGA`
- `Descrição` (for product type detection)

---

## 🚀 Usage

### Running the Application

```powershell
python main.py
```

### GUI Workflow

1. **Launch**: Application opens with DHL/STELLANTIS themed interface
2. **Click "▶ Processar"**: Starts automation process
3. **Monitor Progress**:
   - Status label shows current operation
   - Progress bar (0-100%)
   - Activity log with timestamps
4. **Completion**: Process ends at 100% with summary

### Automated Steps

1. ✅ Load credentials from `credencial.json`
2. ✅ Initialize Playwright browser with Chrome profile
3. ✅ Navigate to Trizy login page
4. ✅ Perform human-like login (character-by-character typing)
5. ✅ Handle Cloudflare verification if triggered
6. ✅ Navigate to "Gestão de Pedidos" → "CONSUMIR ITENS"
7. ✅ Fetch order data from SharePoint (async thread)
8. ✅ Group orders by `chave_pedido_loja` (Order#-Store#)
9. ✅ **For each group**:
   - Filter by order key
   - **For each product**:
     - Filter by product code
     - Select matching checkboxes
     - Track found/not found items
10. ✅ Download Excel template from platform
11. ✅ Process Excel with xlwings (fill quantities, dates, characteristics)
12. ✅ Upload completed file to platform
13. ✅ Display completion summary

---

## 🔄 Workflow Details

### Data Retrieval (Azure_Access.py)

**Function**: `find_and_read_excel_file()`

**SharePoint Path Navigation:**
```
Site: {SITE_ID}
  └── Documents (Drive)
      └── Geral Alpargatas LLP
          └── 19. Base RPA
              └── CARTEIRA GRUPO ASSAÍ.xlsx
```

**Process:**
1. Authenticate with `ClientSecretCredential`
2. Navigate folder hierarchy
3. Locate Excel file by name
4. Read worksheet data via Graph API
5. Convert to pandas DataFrame
6. Return `(df, drive_id, file_id)`

**Key Column Creation:**
```python
df['chave_pedido_loja'] = df['Nº Pedido Cliente'].astype(str) + '-' + df['CÓD LOJA'].astype(str).str.split('-').str[0]
# Example: "12345-001" (Order 12345, Store 001)
```

---

### Web Automation (Tasks.py)

#### Login Process

**Function**: `Login_and_Navigation(page, url, q, username, password)`

**Human-like Typing:**
```python
for char in username:
    page.keyboard.insert_text(char)
    time.sleep(random.uniform(0.02, 0.07))  # 20-70ms delay per character
```

**Cloudflare Handling:**
```python
success_locator = page.locator('span#success-text')
success_locator.wait_for(state='visible', timeout=5000)
# Waits up to 5 seconds for verification completion
```

#### Order Processing

**Function**: `process_orders(page, q)`

**Nested Loop Structure:**
```python
grouped_orders = df.groupby('chave_pedido_loja')

# LOOP 1: By Order Group (chave_pedido_loja)
for chave, group_df in grouped_orders:
    # Filter main input by chave
    chave_input.fill(chave)
    wait_for_data_or_sem_dados()
    
    # LOOP 2: By Product (PRODUTO INTERNO CLIENTE)
    for _, row in group_df.iterrows():
        produto = row['PRODUTO INTERNO CLIENTE']
        # Filter product input
        product_input.fill(produto)
        wait_for_data_or_sem_dados()
        
        # Select checkboxes for matching rows
        checkboxes.click()
        
        # Track found items
        found_items.append({...})
```

**Stability Waits:**
```python
page.wait_for_timeout(1500)  # After fill, before checking results
```

Critical for race conditions where UI updates lag behind input events.

---

### Excel Processing

**Function**: `processar_excel_com_dados(file_path, items, q)`

**xlwings Workflow:**
```python
1. Open workbook with visible Excel instance
2. Read headers from row 3
3. Map column names to indices:
   - 'Quantidade entrega'
   - 'Data sugerida de entrega'
   - 'Característica do veículo'
   - 'Característica da carga'
   - 'Demanda'
   - 'Observação/ fornecedor (opcional)'
   
4. Match Excel rows to `items` list by:
   - código_pedido == chave_pedido_loja
   - código_produto == produto_interno_cliente
   
5. Fill matched rows with:
   - Quantidade (from Qtd. Faturada / 24 or 12)
   - Data (formatted as DD/MM/YYYY)
   - Características (from static_data.json)
   
6. Save workbook and close Excel
```

**Date Formatting:**
```python
cell.value = result_date.date()  # Python date object
cell.number_format = "dd/mm/aaaa"  # Excel format
```

---

## 📊 Data Flow

```
┌─────────────────────────────────────────────────────────────┐
│  SharePoint Excel File (Dados/CARTEIRA GRUPO*.xlsx)        │
│  - Columns: Nº Pedido Cliente, CÓD LOJA, PRODUTO, etc.     │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        ▼ (Azure_Access.py)
┌─────────────────────────────────────────────────────────────┐
│  pandas DataFrame                                           │
│  - Create chave_pedido_loja: "Order#-Store#"               │
│  - Group by chave_pedido_loja                               │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        ▼ (Tasks.py)
┌─────────────────────────────────────────────────────────────┐
│  Web Automation (Trizy Platform)                           │
│  - Filter by chave_pedido_loja                              │
│  - Filter by produto_interno_cliente                        │
│  - Select checkboxes for matching items                     │
│  - Build found_items[] list                                 │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        ▼ (Tasks.py)
┌─────────────────────────────────────────────────────────────┐
│  Download Excel Template from Platform                      │
│  - Save to Arquivos/YYYY-MM-DD.xlsx                         │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        ▼ (processar_excel_com_dados)
┌─────────────────────────────────────────────────────────────┐
│  xlwings Excel Processing                                   │
│  - Match rows by código_pedido + código_produto             │
│  - Fill: quantidade, data, características                  │
│  - Save modified file                                        │
└───────────────────────┬─────────────────────────────────────┘
                        │
                        ▼ (Tasks.py)
┌─────────────────────────────────────────────────────────────┐
│  Upload Completed File to Platform                          │
│  - Locate file input element                                │
│  - Set file path                                             │
│  - Click confirm (if available)                             │
└─────────────────────────────────────────────────────────────┘
                        │
                        ▼
┌─────────────────────────────────────────────────────────────┐
│  [Future] Extract Upload Logs                               │
│  [Future] Update SharePoint with Results                    │
└─────────────────────────────────────────────────────────────┘
```

---

## 🔧 Technical Specifications

### Threading Model

**GUI Thread** (main):
- Runs Tkinter event loop
- Checks message queue every 100ms
- Updates UI components (never from worker threads)

**Automation Thread**:
- Runs Playwright browser automation
- Sends messages to queue: `("status", text)`, `("progress", int)`, `("done", bool)`

**Azure Thread**:
- Runs async operations with new event loop
- Uses `asyncio.new_event_loop()` to avoid conflicts
- Returns results via separate queue

### Queue Communication

```python
# From worker thread:
q.put(("status", "Processing item 5/10"))
q.put(("progress", 50))

# In GUI thread (update_gui):
message_type, value = queue_instance.get_nowait()
if message_type == "status":
    status_label.config(text=value)
    log_text.insert(tk.END, f"{timestamp} - {value}\n")
elif message_type == "progress":
    progress_bar['value'] = value
```

### Chrome Profile Management

**Problem**: Chrome locks profile files when running, causing copy failures.

**Solution**: Lightweight selective copying
```python
PROFILE_ITEMS = [
    "Cookies", "Preferences", "Local State",
    "Local Storage", "Web Data", "Login Data"
]

safe_copy(real_profile, temp_profile, PROFILE_ITEMS)
# Skips locked files, copies only essential data
```

**Cleanup**:
```python
finally:
    shutil.rmtree(temp_profile, ignore_errors=True)
```

### Playwright Configuration

```python
context = playwright.chromium.launch_persistent_context(
    user_data_dir=temp_profile,
    headless=False,  # Visible for debugging
    args=[
        "--start-minimized",
        "--disable-blink-features=AutomationControlled",  # Anti-detection
        "--disable-infobars",
        "--no-sandbox",
        "--disable-dev-shm-usage"
    ]
)
```

### Error Handling Patterns

**Frame Navigation**:
```python
try:
    data_locator.wait_for(state="visible", timeout=5000)
    # Process data
except TimeoutError:
    if sem_dados_locator.is_visible():
        # Expected "no data" scenario
        not_found_items.append({...})
        continue
    else:
        # Unexpected error
        q.put(("status", "ERROR: No data or 'Sem dados' found"))
        continue
```

**Group Processing**:
```python
for group_index, (chave, group_df) in enumerate(grouped_orders):
    try:
        # Process entire group
    except Exception as e:
        q.put(("status", f"Error on group {chave}: {e}. Skipping."))
        # Continue to next group (don't fail entire batch)
```

---

## 👨‍💻 Development Guide

### Adding New Processing Steps

**Location**: `Tasks.py → process_orders()`

**Example**: Add special handling for VIP customers

```python
# After checkbox selection:
if row.get('CLIENTE_VIP') == 'Sim':
    q.put(("status", "    -> VIP customer detected. Applying priority."))
    
    priority_button = frame_locator.get_by_role("button", name="Prioridade")
    priority_button.click()
    page.wait_for_timeout(500)
    
    frame_locator.get_by_text("Alta").click()
    page.wait_for_timeout(300)
```

### Debugging Playwright Issues

```python
# 1. Take screenshot
page.screenshot(path="debug_state.png")

# 2. Pause execution (manual inspection)
page.pause()  # Opens Playwright Inspector

# 3. Print frame HTML for analysis
frame = page.locator("#iframe-servico").content_frame
print(frame.content())

# 4. Check element visibility
locator = frame.locator("selector")
print(f"Visible: {locator.is_visible()}")
print(f"Count: {locator.count()}")
```

### Modifying SharePoint Path

**File**: `Azure_Access.py → find_and_read_excel_file()`

```python
path_segments = [
    "Geral Alpargatas LLP",
    "19. Base RPA",
    # "New Folder Name"  # Add more levels
]
file_name = "NEW_FILE_NAME.xlsx"  # Change target file
```

### Adding Excel Columns

**File**: `Tasks.py → processar_excel_com_dados()`

```python
# 1. Add to column mapping
col_mapping = {
    'quantidade_entrega': None,
    # ... existing columns ...
    'new_column': None  # Add new
}

# 2. Add search logic
for header, col_idx in headers.items():
    # ... existing conditions ...
    elif header_lower == 'new column name':
        col_mapping['new_column'] = col_idx

# 3. Add fill logic in matching section
if col_mapping['new_column']:
    new_val = matching_item.get('new_field_name')
    ws.range(row_num, col_mapping['new_column']).value = new_val
```

### Testing Azure Connection

```python
# Standalone test in Azure_Access.py
if __name__ == "__main__":
    asyncio.run(main())
```

```powershell
# Run test
python Azure_Access.py
```

---

## 🐛 Troubleshooting

### Common Issues

| Issue | Cause | Solution |
|-------|-------|----------|
| **Login fails** | Incorrect credentials or Cloudflare timeout | 1. Verify `credencial.json`<br>2. Increase `WAIT_TIMEOUT_MS` to 10000<br>3. Check for CAPTCHA requirement |
| **Profile copy errors** | Chrome is running and locking files | 1. Close Chrome before starting<br>2. `safe_copy()` should skip locked files automatically |
| **"Sem dados" for valid orders** | Network delay or incorrect key format | 1. Increase `wait_for_timeout()` to 3000ms<br>2. Check `chave_pedido_loja` format in Excel<br>3. Verify CÓD LOJA splitting logic |
| **Excel columns not found** | Column names changed in template | 1. Print headers: `print(f"Headers: {headers}")`<br>2. Update `col_mapping` logic<br>3. Use fuzzy matching if needed |
| **Upload fails** | File input selector changed | 1. Inspect page: `page.pause()`<br>2. Update selector: `frame.locator("input[type='file']")`<br>3. Check for modal timing issues |
| **Azure auth fails** | Invalid credentials or expired secret | 1. Verify `.env` values<br>2. Check Azure app permissions<br>3. Regenerate client secret if expired |
| **xlwings error** | Excel not installed or COM issue | 1. Install Microsoft Excel<br>2. Run as Administrator<br>3. Repair Office installation |

### Debug Mode

Enable detailed logging:

```python
# In main.py, after imports:
import logging
logging.basicConfig(level=logging.DEBUG)

# In Tasks.py:
print(f"DEBUG: chave={chave}, produto={produto}")
print(f"DEBUG: Checkbox count: {checkboxes.count()}")
```

### Performance Optimization

```python
# Reduce waits for faster execution (less stable):
page.wait_for_timeout(500)  # Instead of 1500

# Increase waits for stability (slower):
page.wait_for_timeout(3000)  # Instead of 1500
```

---

## 🚧 Future Enhancements

### Planned Features

1. ✅ **Upload Log Extraction** (In Progress)
   - Function: `Extrair_logs_de_upload_e_Atualizar_sharepoint()`
   - Parse upload results (Success/Error rows)
   - Extract error messages from failed items

2. ⏳ **SharePoint Update Patch** (Next Priority)
   - Use `update_excel_rows()` from Azure_Access.py
   - Mark processed items in SharePoint
   - Update status column with results
   - Batch update via REST API

3. 📧 **Email Notifications**
   - Send summary email on completion
   - Attach error report for not_found_items
   - Use Microsoft Graph SendMail API

4. 📊 **Dashboard Reports**
   - Generate Excel report with:
     - Total items processed
     - Success/failure breakdown
     - Processing time per group
   - Save to `Arquivos/Report_{date}.xlsx`

5. 🔐 **Credential Encryption**
   - Encrypt `credencial.json` with `cryptography` library
   - Use keyring for secure storage

6. 🤖 **Retry Logic**
   - Auto-retry failed items (max 3 attempts)
   - Exponential backoff for network errors

7. 🌐 **Multi-site Support**
   - Support multiple Trizy accounts
   - Site selection in GUI dropdown

### Contributing

```bash
# 1. Create feature branch
git checkout -b feature/upload-log-extraction

# 2. Implement changes
# ... code changes ...

# 3. Test thoroughly
python main.py

# 4. Commit with descriptive message
git add .
git commit -m "feat: Add upload log extraction and SharePoint update"

# 5. Push and create PR
git push origin feature/upload-log-extraction
```

---

## 📝 Code Conventions

### Naming Patterns
- **Functions**: `snake_case` (e.g., `process_orders()`)
- **Classes**: `PascalCase` (e.g., `App`)
- **Constants**: `UPPER_CASE` (e.g., `PROFILE_ITEMS`)
- **Variables**: `snake_case` (e.g., `found_items`)

### Queue Message Format
```python
("status", "String message")   # GUI log entry
("progress", 0-100)             # Progress bar update
("done", True)                  # Terminate update loop
```

### Comment Standards
```python
# --- SECTION HEADER ---
# Subsection explanation
q.put(("status", "User-facing message"))  # Internal comment
```

---

## 👥 Contributors

**Developer**: Vincent Pernarh  
**Organization**: DHL → Alpargatas  
**Project**: RPA-ALPARGATAS  
**Year**: 2024-2025

---

## 📄 License

**Proprietary Software** - Internal use only.  
Unauthorized copying, distribution, or modification is prohibited.

---

## 📞 Support

For issues or questions:
1. Check [Troubleshooting](#troubleshooting) section
2. Review error logs in GUI activity window
3. Contact: Vincent Pernarh (Developer)

---

**Last Updated**: November 25, 2025  
**Version**: 1.0 (Pre-SharePoint Update Patch)