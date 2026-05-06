from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
import pdfplumber
import re
import io
import sqlite3
import hashlib
import jwt
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = FastAPI(title="BrokerNote AI", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

SECRET_KEY = "brokernote-ai-secret-2025"
ALGORITHM = "HS256"


# ─── Database ────────────────────────────────────────────────────────────────

def init_db():
    conn = sqlite3.connect('/tmp/brokernote.db')
    conn.execute('''CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        email TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        name TEXT NOT NULL,
        firm TEXT,
        created_at TEXT NOT NULL
    )''')
    conn.commit()
    conn.close()

init_db()


def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode()).hexdigest()


def create_token(email: str) -> str:
    payload = {
        'email': email,
        'exp': datetime.datetime.utcnow() + datetime.timedelta(days=30)
    }
    return jwt.encode(payload, SECRET_KEY, algorithm=ALGORITHM)


def verify_token(token: str) -> str:
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
        return payload['email']
    except Exception:
        raise HTTPException(status_code=401, detail="Invalid or expired token")


# ─── Auth endpoints ───────────────────────────────────────────────────────────

@app.get("/")
async def root():
    return {"status": "ok", "service": "BrokerNote AI", "version": "1.0.0"}


@app.post("/auth/signup")
async def signup(request_data: dict):
    email = request_data.get('email', '').lower().strip()
    password = request_data.get('password', '')
    name = request_data.get('name', '').strip()
    firm = request_data.get('firm', '').strip()

    if not email or not password or not name:
        raise HTTPException(status_code=400, detail="Email, password and name are required")
    if len(password) < 6:
        raise HTTPException(status_code=400, detail="Password must be at least 6 characters")

    conn = sqlite3.connect('/tmp/brokernote.db')
    try:
        conn.execute(
            "INSERT INTO users (email, password, name, firm, created_at) VALUES (?,?,?,?,?)",
            (email, hash_password(password), name, firm, datetime.datetime.now().isoformat())
        )
        conn.commit()
        token = create_token(email)
        return {"token": token, "email": email, "name": name, "firm": firm}
    except sqlite3.IntegrityError:
        raise HTTPException(status_code=400, detail="Email already registered")
    finally:
        conn.close()


@app.post("/auth/login")
async def login(request_data: dict):
    email = request_data.get('email', '').lower().strip()
    password = request_data.get('password', '')

    conn = sqlite3.connect('/tmp/brokernote.db')
    user = conn.execute(
        "SELECT email, name, firm FROM users WHERE email=? AND password=?",
        (email, hash_password(password))
    ).fetchone()
    conn.close()

    if not user:
        raise HTTPException(status_code=401, detail="Invalid email or password")

    token = create_token(user[0])
    return {"token": token, "email": user[0], "name": user[1], "firm": user[2]}


# ─── PDF Extraction ───────────────────────────────────────────────────────────

def safe_float(val) -> float:
    if val is None:
        return 0.0
    try:
        cleaned = str(val).replace(',', '').replace(' ', '').strip()
        cleaned = re.sub(r'\((.+)\)', r'-\1', cleaned)
        return float(cleaned)
    except Exception:
        return 0.0


def extract_header(text: str) -> dict:
    header = {}
    patterns = {
        "client_name":      r"Client Name\s*:\s*(.+?)(?:\n|Trading|Address)",
        "client_code":      r"Client Code\s*\(UCC\)\s*:\s*(\S+)",
        "pan":              r"PAN Number\s*:\s*([A-Z]{5}[0-9]{4}[A-Z])",
        "contract_note_no": r"Contract Note No\s*:\s*(\S+)",
        "trade_date":       r"Trade Date\s*:\s*([\d/]+)",
        "settlement_no":    r"SETTLEMENT NO\s+(\S+)",
        "gst_no":           r"GST NO\s*:\s*([A-Z0-9]+)",
    }
    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
        if match:
            header[key] = match.group(1).strip()
    return header


def extract_broker_info(text: str) -> dict:
    info = {}
    brokers = ['Angel One', 'Zerodha', 'Upstox', 'HDFC Securities', 'ICICI Direct',
               'Sharekhan', 'Kotak Securities', 'Motilal Oswal', '5Paisa', 'Groww']
    for broker in brokers:
        if broker.lower() in text.lower():
            info["broker"] = broker
            break
    gst_match = re.search(r'GST NO\.?:?\s*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][A-Z0-9]Z[A-Z0-9])', text)
    if gst_match:
        info["broker_gst"] = gst_match.group(1)
    pan_match = re.search(r'PAN No\.?:?\s*([A-Z]{5}[0-9]{4}[A-Z])', text)
    if pan_match:
        info["broker_pan"] = pan_match.group(1)
    return info


def extract_trades_from_table(table: list) -> list:
    trades = []
    is_trade_table = False
    for row in table:
        if not row:
            continue
        row_str = ' '.join([str(c) for c in row if c])
        if 'ISIN' in row_str and ('BUY' in row_str or 'Security' in row_str):
            is_trade_table = True
            continue
        if is_trade_table and row[0] and str(row[0]).startswith('INE'):
            try:
                trade = {
                    "isin":                    (row[0] or '').strip(),
                    "security":                (row[1] or '').strip().replace('\n', ' '),
                    "buy_qty":                 safe_float(row[2] if len(row) > 2 else 0),
                    "buy_wap":                 safe_float(row[3] if len(row) > 3 else 0),
                    "buy_brokerage_per_share": safe_float(row[4] if len(row) > 4 else 0),
                    "buy_wap_after_brokerage": safe_float(row[5] if len(row) > 5 else 0),
                    "total_buy_value":         safe_float(row[6] if len(row) > 6 else 0),
                    "sell_qty":                safe_float(row[7] if len(row) > 7 else 0),
                    "sell_wap":                safe_float(row[8] if len(row) > 8 else 0),
                    "sell_brokerage_per_share":safe_float(row[9] if len(row) > 9 else 0),
                    "sell_wap_after_brokerage":safe_float(row[10] if len(row) > 10 else 0),
                    "total_sell_value":        safe_float(row[11] if len(row) > 11 else 0),
                    "net_qty":                 safe_float(row[12] if len(row) > 12 else 0),
                    "net_obligation":          safe_float(row[13] if len(row) > 13 else 0),
                }
                if trade["isin"]:
                    trades.append(trade)
            except Exception:
                pass
    return trades


def extract_charges_from_table(table: list) -> dict:
    for row in table:
        if not row:
            continue
        row_str = ' '.join([str(c) for c in row if c])
        # Look for the TOTAL row which has actual charge numbers
        if 'TOTAL' in row_str and len([c for c in row if c]) >= 8:
            try:
                nums = [safe_float(c) for c in row if c and c != 'TOTAL(NET)' and c != 'NSE-CAPITAL']
                if len(nums) >= 8:
                    return {
                        "pay_obligation":    nums[0] if len(nums) > 0 else 0,
                        "stt":               nums[1] if len(nums) > 1 else 0,
                        "taxable_value":     nums[2] if len(nums) > 2 else 0,
                        "cgst":              nums[3] if len(nums) > 3 else 0,
                        "sgst":              nums[4] if len(nums) > 4 else 0,
                        "exchange_charges":  nums[5] if len(nums) > 5 else 0,
                        "sebi_fees":         nums[6] if len(nums) > 6 else 0,
                        "stamp_duty":        nums[7] if len(nums) > 7 else 0,
                        "ipf_charges":       nums[8] if len(nums) > 8 else 0,
                        "net_amount":        nums[-1] if nums else 0,
                    }
            except Exception:
                pass
    return {}


def extract_orders_from_table(table: list) -> list:
    orders = []
    is_order_table = False
    for row in table:
        if not row:
            continue
        row_str = ' '.join([str(c) for c in row if c])
        if 'Order No' in row_str and 'Trade No' in row_str:
            is_order_table = True
            continue
        if is_order_table and row[0] and re.match(r'\d{10,}', str(row[0]).strip()):
            try:
                orders.append({
                    "order_no":   str(row[0]).strip(),
                    "order_time": str(row[1]).strip() if len(row) > 1 and row[1] else "",
                    "trade_no":   str(row[2]).strip() if len(row) > 2 and row[2] else "",
                    "trade_time": str(row[3]).strip() if len(row) > 3 and row[3] else "",
                    "security":   str(row[4]).strip().replace('\n', ' ') if len(row) > 4 and row[4] else "",
                    "buy_sell":   str(row[5]).strip() if len(row) > 5 and row[5] else "",
                    "qty":        str(row[6]).strip() if len(row) > 6 and row[6] else "",
                    "gross_rate": str(row[11]).strip() if len(row) > 11 and row[11] else "",
                })
            except Exception:
                pass
    return orders


def extract_contract_note(pdf_bytes: bytes) -> dict:
    result = {
        "header": {},
        "broker_info": {},
        "trades": [],
        "charges": {},
        "order_details": [],
    }
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            tables = page.extract_tables()

            if page_num == 0:
                result["header"] = extract_header(text)
                result["broker_info"] = extract_broker_info(text)
                for table in tables:
                    trades = extract_trades_from_table(table)
                    if trades:
                        result["trades"].extend(trades)
                    charges = extract_charges_from_table(table)
                    if charges:
                        result["charges"] = charges

            elif page_num == 1:
                for table in tables:
                    orders = extract_orders_from_table(table)
                    if orders:
                        result["order_details"].extend(orders)

    return result


@app.post("/extract")
async def extract_pdf(file: UploadFile = File(...)):
    if not (file.filename or '').lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="Please upload a PDF file")
    contents = await file.read()
    if len(contents) > 10 * 1024 * 1024:
        raise HTTPException(status_code=400, detail="File too large (max 10MB)")
    try:
        return extract_contract_note(contents)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Extraction failed: {str(e)}")


# ─── Excel Generation ─────────────────────────────────────────────────────────

def make_border():
    side = Side(style='thin', color='D0D7E5')
    return Border(left=side, right=side, top=side, bottom=side)

DARK_BLUE   = "1B2A4A"
MID_BLUE    = "2563EB"
LIGHT_BLUE  = "EBF3FF"
ALT_ROW     = "F5F8FF"
WHITE       = "FFFFFF"
GREEN       = "16A34A"


@app.post("/excel")
async def generate_excel(data: dict):
    wb = openpyxl.Workbook()

    header_info  = data.get('header', {})
    broker_info  = data.get('broker_info', {})
    charges      = data.get('charges', {})
    trades       = data.get('trades', [])
    order_details= data.get('order_details', [])

    # ── Sheet 1: Trade Summary ──────────────────────────────────────────────
    ws = wb.active
    ws.title = "Trade Summary"
    ws.sheet_view.showGridLines = False

    def hdr_cell(cell, text):
        cell.value = text
        cell.font = Font(bold=True, color=WHITE, size=10)
        cell.fill = PatternFill("solid", fgColor=MID_BLUE)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = make_border()

    def title_merge(ws, cell_range, text):
        ws.merge_cells(cell_range)
        c = ws[cell_range.split(':')[0]]
        c.value = text
        c.font = Font(bold=True, color=DARK_BLUE, size=13)
        c.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
        c.alignment = Alignment(horizontal='center', vertical='center')

    # Title row
    title_merge(ws, 'A1:N1', '📊  CONTRACT NOTE — TRADE SUMMARY  |  BrokerNote AI')
    ws.row_dimensions[1].height = 32

    # Client info
    info_pairs = [
        ("Client Name",      header_info.get('client_name', '—')),
        ("Client Code (UCC)",header_info.get('client_code', '—')),
        ("PAN Number",       header_info.get('pan', '—')),
        ("Trade Date",       header_info.get('trade_date', '—')),
        ("Contract Note No.",header_info.get('contract_note_no', '—')),
        ("Settlement No.",   header_info.get('settlement_no', '—')),
        ("Broker",           broker_info.get('broker', '—')),
        ("Broker GST",       broker_info.get('broker_gst', '—')),
    ]
    for i, (label, value) in enumerate(info_pairs):
        row = 3 + i
        lc = ws.cell(row=row, column=1, value=label)
        vc = ws.cell(row=row, column=2, value=value)
        lc.font = Font(bold=True, size=10, color=DARK_BLUE)
        vc.font = Font(size=10)
        lc.border = vc.border = make_border()

    # Trade headers
    trade_cols = [
        ('A', 'ISIN', 14),
        ('B', 'Security Name', 22),
        ('C', 'Buy Qty', 10),
        ('D', 'Buy WAP (₹)', 14),
        ('E', 'Total Buy Value (₹)', 18),
        ('F', 'Sell Qty', 10),
        ('G', 'Sell WAP (₹)', 14),
        ('H', 'Total Sell Value (₹)', 18),
        ('I', 'Net Qty', 10),
        ('J', 'Net Obligation (₹)', 18),
    ]
    HR = 13
    for col_letter, col_name, width in trade_cols:
        hdr_cell(ws[f'{col_letter}{HR}'], col_name)
        ws.column_dimensions[col_letter].width = width

    for i, trade in enumerate(trades):
        r = HR + 1 + i
        fill = PatternFill("solid", fgColor=ALT_ROW) if i % 2 == 0 else None
        vals = [
            trade.get('isin', ''),
            trade.get('security', ''),
            trade.get('buy_qty', 0),
            trade.get('buy_wap', 0),
            trade.get('total_buy_value', 0),
            trade.get('sell_qty', 0),
            trade.get('sell_wap', 0),
            trade.get('total_sell_value', 0),
            trade.get('net_qty', 0),
            trade.get('net_obligation', 0),
        ]
        for col_idx, val in enumerate(vals, 1):
            c = ws.cell(row=r, column=col_idx, value=val)
            c.border = make_border()
            c.font = Font(size=10)
            if fill:
                c.fill = fill
            if col_idx > 2 and isinstance(val, float):
                c.number_format = '#,##0.00'
            if col_idx in [3, 6, 9]:
                c.number_format = '#,##0'

    # Charges section
    cr = HR + len(trades) + 3
    ws.cell(row=cr, column=1).value = "CHARGES & LEVIES"
    ws.cell(row=cr, column=1).font = Font(bold=True, size=11, color=DARK_BLUE)
    cr += 1

    charge_rows = [
        ("Pay / Receive Obligation",          charges.get('pay_obligation', 0)),
        ("Securities Transaction Tax (STT)",  charges.get('stt', 0)),
        ("Taxable Value of Supply",           charges.get('taxable_value', 0)),
        ("CGST @ 9%",                         charges.get('cgst', 0)),
        ("SGST @ 9%",                         charges.get('sgst', 0)),
        ("Exchange Transaction Charges",      charges.get('exchange_charges', 0)),
        ("SEBI Turnover Fees",                charges.get('sebi_fees', 0)),
        ("Stamp Duty",                        charges.get('stamp_duty', 0)),
        ("IPF Charges",                       charges.get('ipf_charges', 0)),
        ("NET AMOUNT PAYABLE / RECEIVABLE",   charges.get('net_amount', 0)),
    ]

    for label, value in charge_rows:
        is_total = label.startswith("NET AMOUNT")
        lc = ws.cell(row=cr, column=1, value=label)
        vc = ws.cell(row=cr, column=2, value=value)
        lc.border = vc.border = make_border()
        vc.number_format = '#,##0.00'
        if is_total:
            for c in [lc, vc]:
                c.fill = PatternFill("solid", fgColor=DARK_BLUE)
                c.font = Font(bold=True, color=WHITE, size=11)
        else:
            lc.font = Font(size=10)
            vc.font = Font(size=10)
        cr += 1

    ws.column_dimensions['A'].width = 36
    ws.column_dimensions['B'].width = 18

    # ── Sheet 2: Tally Journal Entry ────────────────────────────────────────
    ws2 = wb.create_sheet("Tally Journal Entry")
    ws2.sheet_view.showGridLines = False
    title_merge(ws2, 'A1:F1', '📒  TALLY JOURNAL ENTRY FORMAT  |  BrokerNote AI')
    ws2.row_dimensions[1].height = 32

    jcols = [('A','Date',14), ('B','Voucher Type',16), ('C','Narration / Particulars',35),
             ('D','Ledger Name',32), ('E','Debit (₹)',16), ('F','Credit (₹)',16)]
    for col_letter, col_name, width in jcols:
        hdr_cell(ws2[f'{col_letter}3'], col_name)
        ws2.column_dimensions[col_letter].width = width

    trade_date = header_info.get('trade_date', str(datetime.date.today()))

    journal_rows = []
    for trade in trades:
        total_buy = trade.get('total_buy_value', 0)
        security  = trade.get('security', 'Investment')
        if total_buy > 0:
            journal_rows.append([trade_date, 'Journal', f"Purchase: {security} × {int(trade.get('buy_qty',0))} shares @ ₹{trade.get('buy_wap',0):.2f}", f"{security} A/c", total_buy, ''])
            journal_rows.append([trade_date, 'Journal', '', 'Broker Payable A/c', '', total_buy])

    charge_map = [
        ("STT Expense A/c",          'stt'),
        ("CGST Input Credit A/c",    'cgst'),
        ("SGST Input Credit A/c",    'sgst'),
        ("Exchange Charges Exp A/c", 'exchange_charges'),
        ("SEBI Charges Exp A/c",     'sebi_fees'),
        ("Stamp Duty Expense A/c",   'stamp_duty'),
    ]
    broker_credit = 0
    for ledger, key in charge_map:
        amt = charges.get(key, 0)
        if amt and abs(amt) > 0:
            broker_credit += abs(amt)
            journal_rows.append([trade_date, 'Journal', f"{ledger.replace(' A/c','')}", ledger, abs(amt), ''])

    if broker_credit:
        journal_rows.append([trade_date, 'Journal', 'Charges credited to broker account', 'Broker Payable A/c', '', broker_credit])

    net_amt = abs(charges.get('net_amount', 0))
    if net_amt:
        journal_rows.append([trade_date, 'Payment', 'Payment to broker against contract note', 'Broker Payable A/c', net_amt, ''])
        journal_rows.append([trade_date, 'Payment', '', 'Bank / Cash A/c', '', net_amt])

    for i, jr in enumerate(journal_rows):
        r = 4 + i
        fill = PatternFill("solid", fgColor=ALT_ROW) if i % 2 == 0 else None
        for col, val in enumerate(jr, 1):
            c = ws2.cell(row=r, column=col, value=val)
            c.border = make_border()
            c.font = Font(size=10)
            if fill:
                c.fill = fill
            if col in [5, 6] and isinstance(val, (int, float)) and val:
                c.number_format = '#,##0.00'

    # ── Sheet 3: Order Details ───────────────────────────────────────────────
    if order_details:
        ws3 = wb.create_sheet("Order Details")
        ws3.sheet_view.showGridLines = False
        title_merge(ws3, 'A1:H1', '🗂  ORDER & TRADE DETAILS  |  BrokerNote AI')
        ws3.row_dimensions[1].height = 32
        od_cols = [
            ('A','Order No',22), ('B','Order Time',14), ('C','Trade No',22),
            ('D','Trade Time',14), ('E','Security',22), ('F','B/S',8),
            ('G','Qty',10), ('H','Gross Rate (₹)',16)
        ]
        for col_letter, col_name, width in od_cols:
            hdr_cell(ws3[f'{col_letter}3'], col_name)
            ws3.column_dimensions[col_letter].width = width
        for i, order in enumerate(order_details):
            r = 4 + i
            vals = [order.get('order_no'), order.get('order_time'), order.get('trade_no'),
                    order.get('trade_time'), order.get('security'), order.get('buy_sell'),
                    order.get('qty'), order.get('gross_rate')]
            fill = PatternFill("solid", fgColor=ALT_ROW) if i % 2 == 0 else None
            for col, val in enumerate(vals, 1):
                c = ws3.cell(row=r, column=col, value=val)
                c.border = make_border()
                c.font = Font(size=10)
                if fill:
                    c.fill = fill

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    contract_no = header_info.get('contract_note_no', 'ContractNote').replace('/', '-')
    filename = f"BrokerNote_{contract_no}.xlsx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


# ─── Tally XML ────────────────────────────────────────────────────────────────

@app.post("/tally-xml")
async def generate_tally_xml(data: dict):
    header  = data.get('header', {})
    trades  = data.get('trades', [])
    charges = data.get('charges', {})

    trade_date = header.get('trade_date', '')
    try:
        d = datetime.datetime.strptime(trade_date, '%d/%m/%Y')
        tally_date = d.strftime('%Y%m%d')
    except Exception:
        tally_date = datetime.datetime.now().strftime('%Y%m%d')

    narration = (
        f"Contract Note No: {header.get('contract_note_no','')} | "
        f"Client: {header.get('client_name','')} | PAN: {header.get('pan','')} | "
        f"Broker: {data.get('broker_info',{}).get('broker','')}"
    )

    lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<ENVELOPE>',
        '  <HEADER>',
        '    <TALLYREQUEST>Import Data</TALLYREQUEST>',
        '  </HEADER>',
        '  <BODY>',
        '    <IMPORTDATA>',
        '      <REQUESTDESC>',
        '        <REPORTNAME>Vouchers</REPORTNAME>',
        '        <STATICVARIABLES>',
        '          <SVCURRENTCOMPANY>##SVCURRENTCOMPANY##</SVCURRENTCOMPANY>',
        '        </STATICVARIABLES>',
        '      </REQUESTDESC>',
        '      <REQUESTDATA>',
        '        <TALLYMESSAGE xmlns:UDF="TallyUDF">',
        '          <VOUCHER VCHTYPE="Journal" ACTION="Create" OBJVIEW="Accounting Voucher View">',
        f'            <DATE>{tally_date}</DATE>',
        f'            <NARRATION>{narration}</NARRATION>',
        '            <VOUCHERTYPENAME>Journal</VOUCHERTYPENAME>',
    ]

    total_dr = 0.0

    # Debit: each stock purchased
    for trade in trades:
        buy_val = trade.get('total_buy_value', 0)
        if buy_val > 0:
            total_dr += buy_val
            lines += [
                '            <ALLLEDGERENTRIES.LIST>',
                f'              <LEDGERNAME>{trade.get("security", "Investment")} A/c</LEDGERNAME>',
                '              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>',
                f'              <AMOUNT>-{buy_val:.2f}</AMOUNT>',
                '            </ALLLEDGERENTRIES.LIST>',
            ]

    # Debit: charges
    charge_map = [
        ("STT Expense A/c",          'stt'),
        ("CGST Input Credit A/c",    'cgst'),
        ("SGST Input Credit A/c",    'sgst'),
        ("Exchange Charges Exp A/c", 'exchange_charges'),
        ("SEBI Charges Exp A/c",     'sebi_fees'),
        ("Stamp Duty Expense A/c",   'stamp_duty'),
    ]
    for ledger, key in charge_map:
        amt = abs(charges.get(key, 0))
        if amt > 0:
            total_dr += amt
            lines += [
                '            <ALLLEDGERENTRIES.LIST>',
                f'              <LEDGERNAME>{ledger}</LEDGERNAME>',
                '              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>',
                f'              <AMOUNT>-{amt:.2f}</AMOUNT>',
                '            </ALLLEDGERENTRIES.LIST>',
            ]

    # Credit: broker payable
    broker = data.get('broker_info', {}).get('broker', 'Broker')
    lines += [
        '            <ALLLEDGERENTRIES.LIST>',
        f'              <LEDGERNAME>{broker} Payable A/c</LEDGERNAME>',
        '              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>',
        f'              <AMOUNT>{total_dr:.2f}</AMOUNT>',
        '            </ALLLEDGERENTRIES.LIST>',
    ]

    lines += [
        '          </VOUCHER>',
        '        </TALLYMESSAGE>',
        '      </REQUESTDATA>',
        '    </IMPORTDATA>',
        '  </BODY>',
        '</ENVELOPE>',
    ]

    xml_bytes = '\n'.join(lines).encode('utf-8')
    contract_no = header.get('contract_note_no', 'ContractNote').replace('/', '-')
    return StreamingResponse(
        io.BytesIO(xml_bytes),
        media_type="application/xml",
        headers={"Content-Disposition": f"attachment; filename=Tally_{contract_no}.xml"}
    )
