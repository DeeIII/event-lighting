#!/usr/bin/env python3
"""
Excel Workbook Generator for Flash Lighting Company
Creates a comprehensive bookkeeping and accounting workbook with interconnected sheets
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference
from datetime import datetime

def create_workbook():
    """Create and configure the complete workbook"""
    wb = Workbook()
    
    # Remove default sheet
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
    
    # Create all sheets in order
    create_settings_sheet(wb)
    create_invoice_sheet(wb)
    create_revenue_sheet(wb)
    create_expenses_sheet(wb)
    create_inventory_sheet(wb)
    create_trade_debtors_sheet(wb)
    create_cash_position_sheet(wb)
    create_tax_summary_sheet(wb)
    create_dashboard_sheet(wb)
    
    # Hide Settings sheet
    wb['Settings'].sheet_state = 'hidden'
    
    # Set active sheet to Dashboard
    wb.active = wb['Dashboard']
    
    return wb


def apply_header_style(ws, row, start_col, end_col):
    """Apply consistent header styling"""
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )


def create_settings_sheet(wb):
    """Create hidden Settings sheet for configuration"""
    ws = wb.create_sheet("Settings", 0)
    
    # Company Information
    ws['A1'] = "Company Settings"
    ws['A1'].font = Font(bold=True, size=14)
    
    ws['A3'] = "Company Name"
    ws['B3'] = "Flash Lighting Services Ltd"
    
    ws['A4'] = "VAT Rate (%)"
    ws['B4'] = 15
    ws['B4'].number_format = '0.00'
    
    ws['A5'] = "Financial Year Start"
    ws['B5'] = "2025-01-01"
    
    # Opening Balances
    ws['A7'] = "Opening Balances"
    ws['A7'].font = Font(bold=True, size=12)
    
    ws['A8'] = "Cash in Hand (Opening)"
    ws['B8'] = 5000
    ws['B8'].number_format = '#,##0.00'
    
    ws['A9'] = "Cash at Bank (Opening)"
    ws['B9'] = 50000
    ws['B9'].number_format = '#,##0.00'
    
    # Expense Categories
    ws['A11'] = "Expense Categories"
    ws['A11'].font = Font(bold=True, size=12)
    categories = [
        "Equipment Purchase",
        "Equipment Maintenance",
        "Transport/Fuel",
        "Salaries/Wages",
        "Rent",
        "Utilities",
        "Insurance",
        "Marketing",
        "Office Supplies",
        "Professional Fees",
        "Bank Charges",
        "Other"
    ]
    for idx, cat in enumerate(categories, start=12):
        ws[f'A{idx}'] = cat
    
    # Payment Methods
    ws['D11'] = "Payment Methods"
    ws['D11'].font = Font(bold=True, size=12)
    payment_methods = ["Cash", "Bank Transfer", "Cheque", "Mobile Money", "Card"]
    for idx, method in enumerate(payment_methods, start=12):
        ws[f'D{idx}'] = method
    
    # Set column widths
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['D'].width = 20


def create_invoice_sheet(wb):
    """Create Invoice Generator sheet"""
    ws = wb.create_sheet("Invoices")
    
    # Headers
    headers = [
        "Invoice No.", "Date", "Client Name", "Event Date", 
        "Items/Equipment", "Quantity", "Unit Price", "Subtotal",
        "VAT Rate", "VAT Amount", "Total Amount", "Amount Received", "Balance"
    ]
    
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
    
    apply_header_style(ws, 1, 1, len(headers))
    
    # Sample data row 2 with formulas (template)
    ws['A2'] = "INV-2025-001"
    ws['B2'] = datetime(2025, 1, 15)
    ws['B2'].number_format = 'YYYY-MM-DD'
    ws['C2'] = "Sample Client Ltd"
    ws['D2'] = datetime(2025, 1, 20)
    ws['D2'].number_format = 'YYYY-MM-DD'
    ws['E2'] = "LED Stage Lights x10, Sound System"
    ws['F2'] = 1
    ws['G2'] = 15000
    ws['G2'].number_format = '#,##0.00'
    
    # Formulas
    ws['H2'] = "=F2*G2"  # Subtotal
    ws['H2'].number_format = '#,##0.00'
    
    ws['I2'] = "=Settings!$B$4/100"  # VAT Rate from Settings
    ws['I2'].number_format = '0.00%'
    
    ws['J2'] = "=H2*I2"  # VAT Amount
    ws['J2'].number_format = '#,##0.00'
    
    ws['K2'] = "=H2+J2"  # Total Amount
    ws['K2'].number_format = '#,##0.00'
    
    ws['L2'] = 15000  # Amount Received
    ws['L2'].number_format = '#,##0.00'
    
    ws['M2'] = "=K2-L2"  # Balance
    ws['M2'].number_format = '#,##0.00'
    
    # Set column widths
    widths = [15, 12, 25, 12, 35, 10, 12, 12, 10, 12, 12, 15, 12]
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    
    # Freeze header row
    ws.freeze_panes = 'A2'


def create_revenue_sheet(wb):
    """Create Cash Inflow / Revenue sheet"""
    ws = wb.create_sheet("Revenue")
    
    # Headers
    headers = [
        "Date", "Invoice No.", "Client Name", "Service/Equipment Provided",
        "Quantity", "Unit Price", "Total Amount", "Amount Received", "Balance"
    ]
    
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
    
    apply_header_style(ws, 1, 1, len(headers))
    
    # Link to Invoices sheet - Row 2 pulls from Invoices
    ws['A2'] = "=Invoices!B2"
    ws['A2'].number_format = 'YYYY-MM-DD'
    ws['B2'] = "=Invoices!A2"
    ws['C2'] = "=Invoices!C2"
    ws['D2'] = "=Invoices!E2"
    ws['E2'] = "=Invoices!F2"
    ws['F2'] = "=Invoices!G2"
    ws['F2'].number_format = '#,##0.00'
    ws['G2'] = "=Invoices!K2"
    ws['G2'].number_format = '#,##0.00'
    ws['H2'] = "=Invoices!L2"
    ws['H2'].number_format = '#,##0.00'
    ws['I2'] = "=Invoices!M2"
    ws['I2'].number_format = '#,##0.00'
    
    # Summary section
    ws['K1'] = "REVENUE SUMMARY"
    ws['K1'].font = Font(bold=True, size=12)
    
    ws['K3'] = "Today's Revenue"
    ws['L3'] = f'=SUMIFS(H:H,A:A,TODAY())'
    ws['L3'].number_format = '#,##0.00'
    
    ws['K4'] = "This Month's Revenue"
    ws['L4'] = f'=SUMIFS(H:H,A:A,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),A:A,"<"&DATE(YEAR(TODAY()),MONTH(TODAY())+1,1))'
    ws['L4'].number_format = '#,##0.00'
    
    ws['K5'] = "This Year's Revenue"
    ws['L5'] = f'=SUMIFS(H:H,A:A,">="&DATE(YEAR(TODAY()),1,1),A:A,"<"&DATE(YEAR(TODAY())+1,1,1))'
    ws['L5'].number_format = '#,##0.00'
    
    ws['K7'] = "Total Outstanding"
    ws['L7'] = '=SUM(I:I)'
    ws['L7'].number_format = '#,##0.00'
    
    # Set column widths
    widths = [12, 15, 25, 35, 10, 12, 12, 15, 12, 2, 20, 15]
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    
    ws.freeze_panes = 'A2'


def create_expenses_sheet(wb):
    """Create Expenses sheet"""
    ws = wb.create_sheet("Expenses")
    
    # Headers
    headers = [
        "Date", "Expense Category", "Vendor", "Description", 
        "Amount", "Payment Method", "Reference"
    ]
    
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
    
    apply_header_style(ws, 1, 1, len(headers))
    
    # Sample data
    ws['A2'] = datetime(2025, 1, 10)
    ws['A2'].number_format = 'YYYY-MM-DD'
    ws['B2'] = "Equipment Purchase"
    ws['C2'] = "Tech Supplies Ltd"
    ws['D2'] = "10x LED Par Lights"
    ws['E2'] = 25000
    ws['E2'].number_format = '#,##0.00'
    ws['F2'] = "Bank Transfer"
    ws['G2'] = "TXN-001"
    
    # Summary section
    ws['I1'] = "EXPENSE SUMMARY"
    ws['I1'].font = Font(bold=True, size=12)
    
    ws['I3'] = "This Month's Expenses"
    ws['J3'] = f'=SUMIFS(E:E,A:A,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),A:A,"<"&DATE(YEAR(TODAY()),MONTH(TODAY())+1,1))'
    ws['J3'].number_format = '#,##0.00'
    
    ws['I4'] = "This Year's Expenses"
    ws['J4'] = f'=SUMIFS(E:E,A:A,">="&DATE(YEAR(TODAY()),1,1),A:A,"<"&DATE(YEAR(TODAY())+1,1,1))'
    ws['J4'].number_format = '#,##0.00'
    
    ws['I6'] = "By Category (YTD)"
    ws['I6'].font = Font(bold=True)
    
    # Category summaries
    categories = [
        "Equipment Purchase", "Equipment Maintenance", "Transport/Fuel",
        "Salaries/Wages", "Rent", "Utilities", "Insurance", "Marketing",
        "Office Supplies", "Professional Fees", "Bank Charges", "Other"
    ]
    
    for idx, cat in enumerate(categories, start=7):
        ws[f'I{idx}'] = cat
        ws[f'J{idx}'] = f'=SUMIFS(E:E,B:B,I{idx},A:A,">="&DATE(YEAR(TODAY()),1,1))'
        ws[f'J{idx}'].number_format = '#,##0.00'
    
    # Set column widths
    widths = [12, 20, 25, 35, 12, 18, 15, 2, 25, 15]
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    
    ws.freeze_panes = 'A2'


def create_inventory_sheet(wb):
    """Create Stock / Inventory sheet"""
    ws = wb.create_sheet("Inventory")
    
    # Headers
    headers = [
        "Stock ID", "Item Description", "Unit Price", "Quantity in Store",
        "Quantity Rented Out", "Stock in Transit", "Total Quantity", "Total Stock Value"
    ]
    
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
    
    apply_header_style(ws, 1, 1, len(headers))
    
    # Sample data with formulas
    inventory_items = [
        ("STK-001", "LED Par Light", 2500, 20, 5, 0),
        ("STK-002", "Moving Head Light", 8000, 10, 3, 2),
        ("STK-003", "Sound System (Complete)", 45000, 3, 1, 0),
        ("STK-004", "Smoke Machine", 3500, 8, 2, 0),
        ("STK-005", "DMX Controller", 5000, 5, 1, 0),
        ("STK-006", "Lighting Truss (per meter)", 800, 50, 20, 0),
        ("STK-007", "Power Distribution Box", 1200, 15, 5, 0),
        ("STK-008", "Microphone (Wireless)", 1500, 12, 4, 0),
        ("STK-009", "Speaker (500W)", 6000, 8, 3, 1),
        ("STK-010", "Cable Set (Complete)", 300, 30, 10, 5),
    ]
    
    for idx, (stock_id, desc, price, in_store, rented, transit) in enumerate(inventory_items, start=2):
        ws[f'A{idx}'] = stock_id
        ws[f'B{idx}'] = desc
        ws[f'C{idx}'] = price
        ws[f'C{idx}'].number_format = '#,##0.00'
        ws[f'D{idx}'] = in_store
        ws[f'E{idx}'] = rented
        ws[f'F{idx}'] = transit
        
        # Total Quantity formula
        ws[f'G{idx}'] = f'=D{idx}+E{idx}+F{idx}'
        
        # Total Stock Value formula
        ws[f'H{idx}'] = f'=C{idx}*G{idx}'
        ws[f'H{idx}'].number_format = '#,##0.00'
    
    # Summary section
    last_row = len(inventory_items) + 2
    ws[f'A{last_row}'] = "TOTAL STOCK VALUE"
    ws[f'A{last_row}'].font = Font(bold=True, size=11)
    ws[f'H{last_row}'] = f'=SUM(H2:H{last_row-1})'
    ws[f'H{last_row}'].number_format = '#,##0.00'
    ws[f'H{last_row}'].font = Font(bold=True)
    ws[f'H{last_row}'].fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
    
    # Additional summary
    ws[f'J3'] = "Available for Rent"
    ws[f'K3'] = f'=SUM(D:D)'
    
    ws[f'J4'] = "Currently Rented"
    ws[f'K4'] = f'=SUM(E:E)'
    
    ws[f'J5'] = "In Transit"
    ws[f'K5'] = f'=SUM(F:F)'
    
    # Set column widths
    widths = [12, 30, 12, 18, 18, 15, 15, 18, 2, 20, 15]
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    
    ws.freeze_panes = 'A2'


def create_trade_debtors_sheet(wb):
    """Create Trade Debtors (Accounts Receivable) sheet"""
    ws = wb.create_sheet("Trade Debtors")
    
    # Headers
    headers = [
        "Invoice No.", "Date", "Client Name", "Amount Owed", 
        "Due Date", "Days Outstanding", "Status"
    ]
    
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
    
    apply_header_style(ws, 1, 1, len(headers))
    
    # Instructions
    ws['A2'] = "This sheet auto-populates from Invoices where Balance > 0"
    ws['A2'].font = Font(italic=True, color="666666")
    ws.merge_cells('A2:G2')
    
    # Formula to pull unpaid invoices
    # Row 4 onwards will contain the actual data
    ws['A4'] = '=IFERROR(INDEX(Invoices!$A:$A,SMALL(IF(Invoices!$M$2:$M$1000>0,ROW(Invoices!$M$2:$M$1000)),ROW()-3)),"")'
    ws['B4'] = '=IFERROR(INDEX(Invoices!$B:$B,MATCH(A4,Invoices!$A:$A,0)),"")'
    ws['B4'].number_format = 'YYYY-MM-DD'
    ws['C4'] = '=IFERROR(INDEX(Invoices!$C:$C,MATCH(A4,Invoices!$A:$A,0)),"")'
    ws['D4'] = '=IFERROR(INDEX(Invoices!$M:$M,MATCH(A4,Invoices!$A:$A,0)),"")'
    ws['D4'].number_format = '#,##0.00'
    ws['E4'] = '=IFERROR(B4+30,"")'  # Due date = Invoice date + 30 days
    ws['E4'].number_format = 'YYYY-MM-DD'
    ws['F4'] = '=IFERROR(IF(A4<>"",TODAY()-B4,""),"")'
    ws['G4'] = '=IFERROR(IF(F4="","",IF(F4>60,"Overdue",IF(F4>30,"Due Soon","Current"))),"")'
    
    # Summary
    ws['I3'] = "Total Outstanding"
    ws['J3'] = '=SUM(D:D)'
    ws['J3'].number_format = '#,##0.00'
    ws['J3'].font = Font(bold=True)
    
    ws['I5'] = "Current (0-30 days)"
    ws['J5'] = '=SUMIF(G:G,"Current",D:D)'
    ws['J5'].number_format = '#,##0.00'
    
    ws['I6'] = "Due Soon (31-60 days)"
    ws['J6'] = '=SUMIF(G:G,"Due Soon",D:D)'
    ws['J6'].number_format = '#,##0.00'
    
    ws['I7'] = "Overdue (>60 days)"
    ws['J7'] = '=SUMIF(G:G,"Overdue",D:D)'
    ws['J7'].number_format = '#,##0.00'
    
    # Set column widths
    widths = [15, 12, 25, 15, 12, 18, 15, 2, 20, 15]
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    
    ws.freeze_panes = 'A4'


def create_cash_position_sheet(wb):
    """Create Cash Position sheet"""
    ws = wb.create_sheet("Cash Position")
    
    # Title
    ws['A1'] = "CASH POSITION STATEMENT"
    ws['A1'].font = Font(bold=True, size=14)
    ws.merge_cells('A1:D1')
    
    ws['A2'] = f"As at: {datetime.now().strftime('%Y-%m-%d')}"
    ws.merge_cells('A2:D2')
    
    # Cash in Hand section
    ws['A4'] = "CASH IN HAND"
    ws['A4'].font = Font(bold=True, size=12)
    apply_header_style(ws, 4, 1, 2)
    
    ws['A5'] = "Opening Balance"
    ws['B5'] = "=Settings!B8"
    ws['B5'].number_format = '#,##0.00'
    
    ws['A6'] = "Cash Received (YTD)"
    ws['B6'] = '=SUMIFS(Revenue!H:H,Revenue!A:A,">="&Settings!$B$5,Expenses!F:F,"Cash")'
    ws['B6'].number_format = '#,##0.00'
    
    ws['A7'] = "Cash Paid (YTD)"
    ws['B7'] = '=SUMIFS(Expenses!E:E,Expenses!A:A,">="&Settings!$B$5,Expenses!F:F,"Cash")'
    ws['B7'].number_format = '#,##0.00'
    
    ws['A8'] = "Closing Cash in Hand"
    ws['B8'] = "=B5+B6-B7"
    ws['B8'].number_format = '#,##0.00'
    ws['B8'].font = Font(bold=True)
    ws['B8'].fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    
    # Cash at Bank section
    ws['A10'] = "CASH AT BANK"
    ws['A10'].font = Font(bold=True, size=12)
    apply_header_style(ws, 10, 1, 2)
    
    ws['A11'] = "Opening Balance"
    ws['B11'] = "=Settings!B9"
    ws['B11'].number_format = '#,##0.00'
    
    ws['A12'] = "Bank Receipts (YTD)"
    ws['B12'] = '=SUMIFS(Revenue!H:H,Revenue!A:A,">="&Settings!$B$5)-B6'
    ws['B12'].number_format = '#,##0.00'
    
    ws['A13'] = "Bank Payments (YTD)"
    ws['B13'] = '=SUMIFS(Expenses!E:E,Expenses!A:A,">="&Settings!$B$5)-B7'
    ws['B13'].number_format = '#,##0.00'
    
    ws['A14'] = "Closing Cash at Bank"
    ws['B14'] = "=B11+B12-B13"
    ws['B14'].number_format = '#,##0.00'
    ws['B14'].font = Font(bold=True)
    ws['B14'].fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    
    # Total Cash Position
    ws['A16'] = "TOTAL CASH POSITION"
    ws['A16'].font = Font(bold=True, size=13)
    ws['B16'] = "=B8+B14"
    ws['B16'].number_format = '#,##0.00'
    ws['B16'].font = Font(bold=True, size=13)
    ws['B16'].fill = PatternFill(start_color="B4C7E7", end_color="B4C7E7", fill_type="solid")
    
    # Monthly Breakdown
    ws['D4'] = "MONTHLY CASH FLOW"
    ws['D4'].font = Font(bold=True, size=12)
    ws.merge_cells('D4:F4')
    
    ws['D5'] = "Month"
    ws['E5'] = "Inflows"
    ws['F5'] = "Outflows"
    ws['G5'] = "Net"
    apply_header_style(ws, 5, 4, 7)
    
    # Sample months (would need dynamic generation for full year)
    months = ["January", "February", "March", "April", "May", "June", 
              "July", "August", "September", "October", "November", "December"]
    
    for idx, month in enumerate(months, start=6):
        month_num = idx - 5
        ws[f'D{idx}'] = month
        ws[f'E{idx}'] = f'=SUMIFS(Revenue!H:H,Revenue!A:A,">="&DATE(YEAR(TODAY()),{month_num},1),Revenue!A:A,"<"&DATE(YEAR(TODAY()),{month_num}+1,1))'
        ws[f'E{idx}'].number_format = '#,##0.00'
        ws[f'F{idx}'] = f'=SUMIFS(Expenses!E:E,Expenses!A:A,">="&DATE(YEAR(TODAY()),{month_num},1),Expenses!A:A,"<"&DATE(YEAR(TODAY()),{month_num}+1,1))'
        ws[f'F{idx}'].number_format = '#,##0.00'
        ws[f'G{idx}'] = f'=E{idx}-F{idx}'
        ws[f'G{idx}'].number_format = '#,##0.00'
    
    # Set column widths
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 3
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['F'].width = 15
    ws.column_dimensions['G'].width = 15


def create_tax_summary_sheet(wb):
    """Create Tax Summary sheet for end-of-year reporting"""
    ws = wb.create_sheet("Tax Summary")
    
    # Title
    ws['A1'] = "END-OF-YEAR TAX SUMMARY"
    ws['A1'].font = Font(bold=True, size=14)
    ws.merge_cells('A1:D1')
    
    ws['A2'] = "Financial Year: 2025"
    ws['A2'].font = Font(size=11)
    ws.merge_cells('A2:D2')
    
    # Revenue Section
    ws['A4'] = "REVENUE"
    ws['A4'].font = Font(bold=True, size=12, color="FFFFFF")
    ws['A4'].fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    ws.merge_cells('A4:B4')
    
    ws['A5'] = "Total Revenue (Invoiced)"
    ws['B5'] = '=SUMIFS(Revenue!G:G,Revenue!A:A,">="&Settings!$B$5,Revenue!A:A,"<"&DATE(YEAR(Settings!$B$5)+1,1,1))'
    ws['B5'].number_format = '#,##0.00'
    
    ws['A6'] = "Total Revenue (Received)"
    ws['B6'] = '=SUMIFS(Revenue!H:H,Revenue!A:A,">="&Settings!$B$5,Revenue!A:A,"<"&DATE(YEAR(Settings!$B$5)+1,1,1))'
    ws['B6'].number_format = '#,##0.00'
    
    ws['A7'] = "Outstanding Receivables"
    ws['B7'] = "=B5-B6"
    ws['B7'].number_format = '#,##0.00'
    
    # VAT Section
    ws['A9'] = "VALUE ADDED TAX (VAT)"
    ws['A9'].font = Font(bold=True, size=12, color="FFFFFF")
    ws['A9'].fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    ws.merge_cells('A9:B9')
    
    ws['A10'] = "VAT Collected on Sales"
    ws['B10'] = '=SUMIFS(Invoices!J:J,Invoices!B:B,">="&Settings!$B$5,Invoices!B:B,"<"&DATE(YEAR(Settings!$B$5)+1,1,1))'
    ws['B10'].number_format = '#,##0.00'
    
    ws['A11'] = "VAT Paid on Purchases"
    ws['B11'] = '=SUMIFS(Expenses!E:E,Expenses!A:A,">="&Settings!$B$5,Expenses!A:A,"<"&DATE(YEAR(Settings!$B$5)+1,1,1))*(Settings!$B$4/100)/(1+Settings!$B$4/100)'
    ws['B11'].number_format = '#,##0.00'
    
    ws['A12'] = "Net VAT Payable"
    ws['B12'] = "=B10-B11"
    ws['B12'].number_format = '#,##0.00'
    ws['B12'].font = Font(bold=True)
    ws['B12'].fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    
    # Expenses Section
    ws['A14'] = "EXPENSES (Deductible)"
    ws['A14'].font = Font(bold=True, size=12, color="FFFFFF")
    ws['A14'].fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    ws.merge_cells('A14:B14')
    
    categories = [
        "Equipment Purchase", "Equipment Maintenance", "Transport/Fuel",
        "Salaries/Wages", "Rent", "Utilities", "Insurance", "Marketing",
        "Office Supplies", "Professional Fees", "Bank Charges", "Other"
    ]
    
    start_row = 15
    for idx, cat in enumerate(categories, start=start_row):
        ws[f'A{idx}'] = cat
        ws[f'B{idx}'] = f'=SUMIFS(Expenses!E:E,Expenses!B:B,A{idx},Expenses!A:A,">="&Settings!$B$5,Expenses!A:A,"<"&DATE(YEAR(Settings!$B$5)+1,1,1))'
        ws[f'B{idx}'].number_format = '#,##0.00'
    
    total_row = start_row + len(categories)
    ws[f'A{total_row}'] = "Total Expenses"
    ws[f'A{total_row}'].font = Font(bold=True)
    ws[f'B{total_row}'] = f'=SUM(B{start_row}:B{total_row-1})'
    ws[f'B{total_row}'].number_format = '#,##0.00'
    ws[f'B{total_row}'].font = Font(bold=True)
    ws[f'B{total_row}'].fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    
    # Net Profit Section
    profit_row = total_row + 2
    ws[f'A{profit_row}'] = "NET PROFIT BEFORE TAX"
    ws[f'A{profit_row}'].font = Font(bold=True, size=13)
    ws[f'B{profit_row}'] = f'=B6-B{total_row}'
    ws[f'B{profit_row}'].number_format = '#,##0.00'
    ws[f'B{profit_row}'].font = Font(bold=True, size=13)
    ws[f'B{profit_row}'].fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    
    # Profit Margin
    margin_row = profit_row + 1
    ws[f'A{margin_row}'] = "Profit Margin (%)"
    ws[f'B{margin_row}'] = f'=IF(B6=0,0,B{profit_row}/B6)'
    ws[f'B{margin_row}'].number_format = '0.00%'
    ws[f'B{margin_row}'].font = Font(bold=True)
    
    # Set column widths
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 20


def create_dashboard_sheet(wb):
    """Create Business Health Dashboard"""
    ws = wb.create_sheet("Dashboard")
    
    # Title
    ws['A1'] = "BUSINESS HEALTH DASHBOARD"
    ws['A1'].font = Font(bold=True, size=16, color="FFFFFF")
    ws['A1'].fill = PatternFill(start_color="203864", end_color="203864", fill_type="solid")
    ws['A1'].alignment = Alignment(horizontal="center")
    ws.merge_cells('A1:F1')
    
    ws['A2'] = f"Updated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws['A2'].alignment = Alignment(horizontal="center")
    ws.merge_cells('A2:F2')
    
    # Key Metrics Section
    ws['A4'] = "KEY FINANCIAL METRICS"
    ws['A4'].font = Font(bold=True, size=13)
    ws.merge_cells('A4:F4')
    
    # Metric cards
    metrics = [
        ("Monthly Revenue", '=Revenue!L4', 'A6', 'B6'),
        ("Monthly Expenses", '=Expenses!J3', 'A7', 'B7'),
        ("Monthly Net Profit", '=B6-B7', 'A8', 'B8'),
        ("Profit Margin", '=IF(B6=0,0,B8/B6)', 'A9', 'B9'),
        ("", "", "", ""),
        ("YTD Revenue", '=Revenue!L5', 'A11', 'B11'),
        ("YTD Expenses", '=Expenses!J4', 'A12', 'B12'),
        ("YTD Net Profit", '=B11-B12', 'A13', 'B13'),
        ("YTD Profit Margin", '=IF(B11=0,0,B13/B11)', 'A14', 'B14'),
    ]
    
    for metric_name, formula, label_cell, value_cell in metrics:
        if metric_name:
            ws[label_cell] = metric_name
            ws[label_cell].font = Font(bold=True)
            ws[value_cell] = formula
            if "Margin" in metric_name:
                ws[value_cell].number_format = '0.00%'
            else:
                ws[value_cell].number_format = '#,##0.00'
            ws[value_cell].font = Font(bold=True, size=12)
            ws[value_cell].fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
    
    # Cash Position
    ws['D6'] = "Cash in Hand"
    ws['D6'].font = Font(bold=True)
    ws['E6'] = "='Cash Position'!B8"
    ws['E6'].number_format = '#,##0.00'
    ws['E6'].fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
    
    ws['D7'] = "Cash at Bank"
    ws['D7'].font = Font(bold=True)
    ws['E7'] = "='Cash Position'!B14"
    ws['E7'].number_format = '#,##0.00'
    ws['E7'].fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
    
    ws['D8'] = "Total Cash"
    ws['D8'].font = Font(bold=True)
    ws['E8'] = "=E6+E7"
    ws['E8'].number_format = '#,##0.00'
    ws['E8'].font = Font(bold=True, size=12)
    ws['E8'].fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    
    # Receivables
    ws['D10'] = "Outstanding Invoices"
    ws['D10'].font = Font(bold=True)
    ws['E10'] = "='Trade Debtors'!J3"
    ws['E10'].number_format = '#,##0.00'
    ws['E10'].fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    
    ws['D11'] = "Total Stock Value"
    ws['D11'].font = Font(bold=True)
    ws['E11'] = "=Inventory!H12"
    ws['E11'].number_format = '#,##0.00'
    ws['E11'].fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    
    # Revenue vs Expense Trend
    ws['A16'] = "REVENUE VS EXPENSE TREND (Monthly)"
    ws['A16'].font = Font(bold=True, size=12)
    ws.merge_cells('A16:F16')
    
    ws['A17'] = "Month"
    ws['B17'] = "Revenue"
    ws['C17'] = "Expenses"
    ws['D17'] = "Net Profit"
    apply_header_style(ws, 17, 1, 4)
    
    months = ["January", "February", "March", "April", "May", "June", 
              "July", "August", "September", "October", "November", "December"]
    
    for idx, month in enumerate(months, start=18):
        month_num = idx - 17
        ws[f'A{idx}'] = month
        ws[f'B{idx}'] = f'=SUMIFS(Revenue!H:H,Revenue!A:A,">="&DATE(YEAR(TODAY()),{month_num},1),Revenue!A:A,"<"&DATE(YEAR(TODAY()),{month_num}+1,1))'
        ws[f'B{idx}'].number_format = '#,##0.00'
        ws[f'C{idx}'] = f'=SUMIFS(Expenses!E:E,Expenses!A:A,">="&DATE(YEAR(TODAY()),{month_num},1),Expenses!A:A,"<"&DATE(YEAR(TODAY()),{month_num}+1,1))'
        ws[f'C{idx}'].number_format = '#,##0.00'
        ws[f'D{idx}'] = f'=B{idx}-C{idx}'
        ws[f'D{idx}'].number_format = '#,##0.00'
    
    # Create chart
    chart = LineChart()
    chart.title = "Revenue vs Expense Trend"
    chart.style = 12
    chart.y_axis.title = "Amount"
    chart.x_axis.title = "Month"
    chart.height = 10
    chart.width = 20
    
    # Add data
    data = Reference(ws, min_col=2, min_row=17, max_row=29, max_col=3)
    cats = Reference(ws, min_col=1, min_row=18, max_row=29)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    
    ws.add_chart(chart, "F18")
    
    # Set column widths
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 18
    ws.column_dimensions['F'].width = 3


def main():
    """Main function to generate the workbook"""
    print("Generating Event Lighting Company Bookkeeping Workbook...")
    
    wb = create_workbook()
    
    filename = "Event_Lighting_Bookkeeping.xlsx"
    wb.save(filename)
    
    print(f"✓ Workbook created successfully: {filename}")
    print("\nSheets created:")
    for sheet in wb.sheetnames:
        if sheet != "Settings":
            print(f"  - {sheet}")
    print("  - Settings (hidden)")
    
    print("\nWorkbook is ready to use!")
    print("You can customize the Settings sheet for your company details and opening balances.")


if __name__ == "__main__":
    main()
