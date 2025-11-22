# Professional Bookkeeping System - QuickBooks Edition

## Overview

This is a **professional-grade bookkeeping system** for event lighting companies, designed to match the functionality of QuickBooks. Built with Python and Excel, it provides comprehensive financial management, reporting, and operational tools.

---

## 🚀 Quick Start

### Installation

```bash
# Install required package
pip3 install openpyxl

# Generate the workbook
python3 generate_workbook.py
```

The system generates `Event_Lighting_Bookkeeping.xlsx` with 14+ interconnected sheets.

---

## 📊 Professional Features

### Financial Reports (Like QuickBooks)

1. **Dashboard** - Business Health Overview
   - Real-time KPIs (Days Sales Outstanding, Quick Ratio, Profit Margins)
   - Cash position summary
   - Monthly revenue vs expense trend with charts
   - Visual alerts using conditional formatting

2. **Profit & Loss Statement**
   - Standard accounting format
   - Monthly and Year-to-Date columns
   - Revenue, Cost of Services, Operating Expenses breakdown
   - Gross Profit and Net Income calculations
   - Profit margin percentages

3. **Balance Sheet**
   - Assets (Current & Fixed)
   - Liabilities & Equity
   - Auto-balancing check
   - Professional accounting structure

4. **Cash Position Statement**
   - Cash in Hand vs Cash at Bank
   - Monthly breakdown of cash flows
   - Opening and closing balances
   - Year-to-date tracking

5. **Bank Reconciliation**
   - Statement balance reconciliation
   - Outstanding deposits tracking
   - Outstanding checks/payments
   - Automatic difference calculation

6. **Tax Summary Report**
   - VAT collected vs VAT paid
   - Net VAT payable
   - Expense categories for tax purposes
   - Year-end profit calculations

---

## 💼 Operational Management

### Customer Management
- Complete customer database
- Credit limits and payment terms
- Auto-calculated balances from invoices
- Status indicators (Paid, Active, Credit Warning)
- Contact information management

### Vendor/Supplier Management
- Vendor database with contact details
- Purchase history tracking
- Payment terms management
- Total purchased and paid amounts

### Invoice Management
- Professional invoice generation
- Auto-calculated VAT and totals
- Payment status tracking
- **Data validation** - Customer names from dropdown
- Color-coded payment status (Paid/Unpaid/Partially Paid)
- Printable invoice template

### Expense Tracking
- Category-based expense management
- **Data validation** - Categories and vendors from dropdowns
- Payment method tracking
- Monthly and yearly summaries
- Category-wise breakdowns

### Inventory Management
- Stock ID tracking
- Quantity in store/rented out/in transit
- Unit prices and total valuations
- Real-time stock value calculations

### Trade Debtors (Accounts Receivable)
- Auto-populated from unpaid invoices
- Aging analysis (Current, Due Soon, Overdue)
- Days outstanding calculation
- Total outstanding summary

---

## 🔒 Professional Features

### Data Validation
- **Dropdown lists** for consistency:
  - Customer names (from Customer database)
  - Vendor names (from Vendor database)
  - Expense categories (from Settings)
  - Payment methods (Cash, Bank Transfer, Mobile Money, etc.)
  - Invoice status

### Conditional Formatting
- **Visual alerts** throughout:
  - Red highlight for overdue/unpaid amounts
  - Green highlight for paid/positive balances
  - Yellow for warnings (credit limits, reconciliation differences)
  - Profit/loss color indicators

### Audit Trail
- Hidden Audit Log sheet
- Transaction history tracking
- Timestamp logging
- Reference numbers for all transactions

### Cross-Sheet Formulas
- Automatic data synchronization
- Real-time updates across all sheets
- Formula-driven calculations
- No manual data duplication

### Centralized Settings
- Hidden Settings sheet
- Company information
- VAT rate configuration
- Financial year settings
- Opening balances
- Expense categories list
- Payment methods list

---

## 📋 Sheet-by-Sheet Guide

### 1. Dashboard (Landing Page)
**Purpose**: Executive summary of business health

**Key Metrics**:
- Monthly & YTD Revenue, Expenses, Net Profit
- Profit Margins
- Cash positions (Hand & Bank)
- Outstanding invoices & stock value
- VAT payable

**KPIs**:
- Days Sales Outstanding (DSO) - measures how quickly you collect payments
- Quick Ratio - measures liquidity (should be > 1.0)
- Gross Profit Margin
- Total Assets value

**Charts**: Monthly revenue vs expense trend

---

### 2. Customers
**Purpose**: Customer relationship management

**Features**:
- Customer ID, Company Name, Contact Person
- Email, Phone, Address, City
- Payment Terms (Days)
- Credit Limit
- Auto-calculated Total Invoiced, Paid, Balance
- Status indicators with color coding

**How to Use**:
1. Add new customers with unique IDs
2. Set credit limits
3. Balances update automatically from Invoices sheet
4. Status changes to "Credit Warning" when balance exceeds 80% of limit

---

### 3. Vendors
**Purpose**: Supplier/vendor management

**Features**:
- Vendor ID, Contact information
- Payment terms
- Auto-calculated purchase totals

---

### 4. Invoices
**Purpose**: Create and track customer invoices

**Features**:
- Invoice numbering (INV-2025-001)
- Date tracking (Invoice date & Event date)
- **Dropdown validation** for customer selection
- Quantity and pricing
- Auto-calculated subtotal, VAT, and total
- Amount received and balance tracking
- Payment status with color coding

**Formulas**:
- Subtotal = Quantity × Unit Price
- VAT Amount = Subtotal × VAT Rate (from Settings)
- Total = Subtotal + VAT
- Balance = Total - Amount Received
- Status = Auto-determined based on balance

---

### 5. Revenue
**Purpose**: Track all revenue/cash inflows

**Features**:
- **Auto-syncs from Invoices** - no manual entry needed
- Revenue summaries (Today, This Month, This Year)
- Total outstanding calculation

---

### 6. Expenses
**Purpose**: Track all business expenses

**Features**:
- **Dropdown validation** for:
  - Expense Category
  - Vendor Name
  - Payment Method
- Monthly and yearly expense summaries
- Category-wise breakdowns for analysis

**Categories**:
- Equipment Purchase
- Equipment Maintenance
- Transport/Fuel
- Salaries/Wages
- Rent, Utilities, Insurance
- Marketing, Office Supplies
- Professional Fees, Bank Charges
- Other

---

### 7. Inventory
**Purpose**: Manage equipment and stock

**Features**:
- Stock ID tracking
- Item descriptions and unit prices
- Quantities: In Store, Rented Out, In Transit
- Auto-calculated total quantity and stock value
- Summary statistics

**Sample Items**:
- LED Par Lights, Moving Head Lights
- Sound Systems, Smoke Machines
- DMX Controllers, Lighting Truss
- Microphones, Speakers, Cables

---

### 8. Trade Debtors
**Purpose**: Track who owes you money

**Features**:
- **Auto-populated** from unpaid invoices
- Due date calculation (Invoice Date + 30 days)
- Days outstanding
- Status: Current (0-30 days), Due Soon (31-60), Overdue (>60)
- Aging analysis summaries

---

### 9. Cash Position
**Purpose**: Monitor cash flow

**Sections**:
- **Cash in Hand**: Opening + Receipts - Payments
- **Cash at Bank**: Opening + Deposits - Withdrawals
- **Monthly Breakdown**: 12-month cash flow table

---

### 10. Bank Reconciliation
**Purpose**: Match books with bank statements (like QuickBooks)

**Process**:
1. Enter bank statement balance
2. List outstanding deposits (money in transit)
3. List outstanding checks (not yet cleared)
4. System calculates adjusted balance
5. Compares with book balance
6. Highlights differences (should be zero)

---

### 11. Profit & Loss
**Purpose**: Standard P&L statement (like QuickBooks)

**Structure**:
```
REVENUE
- Service Revenue
= Total Revenue

COST OF SERVICES
- Equipment Costs
= Total Cost of Services

GROSS PROFIT

OPERATING EXPENSES
- (11 categories listed separately)
= Total Operating Expenses

NET INCOME (LOSS)
Net Profit Margin %
```

**Columns**: This Month | Year to Date

---

### 12. Balance Sheet
**Purpose**: Financial position snapshot (like QuickBooks)

**Structure**:
```
ASSETS
  Current Assets
    - Cash in Hand
    - Cash at Bank
    - Accounts Receivable
  Fixed Assets
    - Equipment & Inventory
  = TOTAL ASSETS

LIABILITIES & EQUITY
  Current Liabilities
    - Accounts Payable
    - VAT Payable
  Owner's Equity
    - Opening Capital
    - Retained Earnings (YTD)
  = TOTAL LIABILITIES & EQUITY

CHECK: Assets = Liabilities + Equity
```

---

### 13. Tax Summary
**Purpose**: End-of-year tax reporting

**Sections**:
- Revenue (Invoiced vs Received)
- VAT Collected vs VAT Paid
- **Net VAT Payable** (important for tax filing)
- Expense breakdown by category
- Net Profit Before Tax
- Profit Margin %

---

### 14. Invoice Template
**Purpose**: Print-ready invoice for customers

**Features**:
- Professional layout
- Company branding area
- Bill To section
- Line items table
- VAT calculation
- Payment information
- Print area pre-configured

**How to Use**:
1. Copy data from Invoices sheet
2. Customize as needed
3. File → Print or Ctrl+P

---

## 🎯 How to Use This System

### Initial Setup

1. **Unhide Settings Sheet**:
   - Right-click any sheet tab → Unhide → Select "Settings"
   
2. **Configure Settings**:
   - Company Name: Flash Lighting Services Ltd (or your name)
   - VAT Rate: 15% (or your country's rate)
   - Financial Year Start: 2025-01-01
   - Opening Cash Balances

3. **Add Your Data**:
   - Start with **Customers** sheet - add all clients
   - Add **Vendors** sheet - add all suppliers
   - Update **Inventory** - your actual equipment

### Daily Operations

1. **Creating an Invoice**:
   - Go to Invoices sheet
   - New row: Invoice No., Date, Select Customer (dropdown)
   - Enter items, quantity, price
   - VAT and totals calculate automatically
   - Record Amount Received
   - Status updates automatically

2. **Recording an Expense**:
   - Go to Expenses sheet
   - New row: Date, Category (dropdown), Vendor (dropdown)
   - Enter amount and payment method (dropdown)
   - Add reference number

3. **Checking Business Health**:
   - Open Dashboard sheet
   - All metrics update in real-time
   - Review KPIs and trends

### Monthly Tasks

1. **Bank Reconciliation**:
   - Go to Bank Reconciliation sheet
   - Enter bank statement balance
   - List outstanding items
   - Verify difference is zero

2. **Review Reports**:
   - Profit & Loss - check profitability
   - Cash Position - ensure adequate cash
   - Trade Debtors - follow up on overdue invoices

### End-of-Year Tasks

1. **Tax Preparation**:
   - Review Tax Summary sheet
   - Note Net VAT Payable
   - Export expense categories for accountant

2. **Financial Statements**:
   - Print/PDF the Profit & Loss
   - Print/PDF the Balance Sheet
   - Verify Balance Sheet balances (Assets = Liabilities + Equity)

---

## 💡 Pro Tips

### Best Practices

1. **Data Entry**:
   - Always use dropdowns - don't type names manually
   - Use consistent date format (YYYY-MM-DD)
   - Enter transactions same day for accuracy

2. **Invoice Numbering**:
   - Keep sequential: INV-2025-001, INV-2025-002, etc.
   - Include year for easy filing

3. **Customer Credit Management**:
   - Review Customer sheet weekly
   - Watch for "Credit Warning" status
   - Follow up on overdue invoices promptly

4. **Expense Categories**:
   - Use correct categories for tax purposes
   - Keep receipts for all expenses
   - Add reference numbers

5. **Monthly Reconciliation**:
   - Always reconcile bank at month-end
   - Investigate any differences immediately
   - Keep bank statements organized

### Performance Monitoring

**Key Metrics to Watch**:

| Metric | Formula | Target | Action If Below |
|--------|---------|--------|----------------|
| Quick Ratio | Current Assets / Current Liabilities | > 1.0 | Improve cash collection |
| Days Sales Outstanding (DSO) | (Receivables / Revenue) × 365 | < 30 days | Follow up on invoices |
| Gross Profit Margin | (Revenue - COGS) / Revenue × 100 | > 40% | Review pricing/costs |
| Net Profit Margin | Net Income / Revenue × 100 | > 15% | Control expenses |

---

## 🔧 Customization

### Adding New Expense Categories

1. Unhide Settings sheet
2. Add category in column A (row 12 onwards)
3. Category will appear in Expenses dropdown automatically
4. Update Profit & Loss and Tax Summary sheets (add to formulas)

### Adding Payment Methods

1. Unhide Settings sheet
2. Add method in column D (row 12 onwards)
3. Method appears in Expenses dropdown automatically

### Changing VAT Rate

1. Unhide Settings sheet
2. Change cell B4 (currently 15%)
3. All invoices update automatically

### Adding More Inventory Items

1. Go to Inventory sheet
2. Add new row with Stock ID, Description, Unit Price
3. Enter quantities
4. Formulas calculate automatically

---

## 📈 Advanced Features

### Conditional Formatting Rules

**Invoices Sheet**:
- Green background = Fully paid (Balance = 0)
- Red background = Unpaid (Balance > 0)

**Customers Sheet**:
- Green = "Paid" status
- Red = "Credit Warning" status

**Dashboard**:
- Quick Ratio: Green > 1.0, Red < 1.0
- Monthly profit: Green if positive, Red if negative

**Bank Reconciliation**:
- Green = Balanced (Difference = 0)
- Red = Not balanced (Difference ≠ 0)

### Formula Architecture

**Cross-Sheet References**:
```
Revenue → pulls from → Invoices
Dashboard → pulls from → Revenue, Expenses, Trade Debtors, etc.
Balance Sheet → pulls from → Cash Position, Tax Summary, Profit & Loss
```

**Auto-Calculations**:
- All monetary totals use SUM
- Date-based summaries use SUMIFS with DATE functions
- Status fields use IF statements
- Trade Debtors uses INDEX/MATCH for unpaid invoices

---

## 🛡️ Data Protection

### Recommended Protection Steps

1. **Protect Formula Cells**:
   - Select all cells with formulas
   - Format Cells → Protection → Locked
   - Review → Protect Sheet
   - Allow: Insert rows, Delete rows, Format cells

2. **Backup Regularly**:
   - Daily: Save copy to cloud (Google Drive, Dropbox)
   - Weekly: Export to PDF for reports
   - Monthly: Archive copy with date

3. **Access Control**:
   - Keep master file secure
   - Provide view-only copies if sharing
   - Hide Audit Log sheet (already hidden)

---

## 🆚 QuickBooks Comparison

### Features This System Has

| Feature | This System | QuickBooks |
|---------|-------------|------------|
| Customer Database | ✅ | ✅ |
| Vendor Management | ✅ | ✅ |
| Invoice Creation | ✅ | ✅ |
| Expense Tracking | ✅ | ✅ |
| Profit & Loss | ✅ | ✅ |
| Balance Sheet | ✅ | ✅ |
| Bank Reconciliation | ✅ | ✅ |
| Dashboard/KPIs | ✅ | ✅ |
| Tax Reports | ✅ | ✅ |
| Inventory Management | ✅ | ✅ |
| Data Validation | ✅ | ✅ |
| Audit Trail | ✅ | ✅ |
| **Cost** | **FREE** | **$30-100/month** |
| **Customization** | **Fully Customizable** | **Limited** |
| **Offline Use** | **Yes** | **Limited** |
| **Learning Curve** | **Moderate** | **Moderate** |

### When to Use This vs QuickBooks

**Use This System If**:
- You want free, no subscription costs
- You need full control and customization
- You're comfortable with Excel
- Your business is small to medium size
- You want offline access always

**Use QuickBooks If**:
- You need payroll integration
- You want automatic bank feeds
- You need multi-user collaboration
- You want mobile apps
- You need advanced inventory (purchasing, reordering)

---

## 🐛 Troubleshooting

### Common Issues

**Issue**: Dropdown lists not working
**Solution**: Make sure Customers/Vendors/Settings sheets have data in correct columns

**Issue**: Formulas showing #REF! error
**Solution**: Don't delete rows from Customers/Vendors/Settings sheets; hide instead

**Issue**: Trade Debtors not populating
**Solution**: Make sure Invoices sheet has Balance column (column M) with values > 0

**Issue**: Dashboard metrics show 0
**Solution**: Enter data in Invoices and Expenses sheets first

**Issue**: VAT calculations wrong
**Solution**: Check Settings sheet cell B4 for correct VAT rate

---

## 📞 Support & Contribution

### Getting Help

1. Review this README thoroughly
2. Check WARP.md for technical details
3. Review formula errors in Excel's formula bar
4. Test with sample data first

### Contributing

To improve this system:
1. Add new features in separate functions
2. Follow existing naming conventions
3. Update WARP.md with architectural changes
4. Test thoroughly before committing

---

## 📜 Version History

### v2.0 (QuickBooks Edition) - November 2025
- ✅ Added Customer & Vendor management
- ✅ Added Profit & Loss statement
- ✅ Added Balance Sheet
- ✅ Added Bank Reconciliation
- ✅ Added professional Invoice Template
- ✅ Added Audit Log
- ✅ Implemented data validation (dropdowns)
- ✅ Added conditional formatting throughout
- ✅ Enhanced Dashboard with KPIs
- ✅ Professional styling matching QuickBooks

### v1.0 (Basic) - January 2025
- Basic invoice, expense, inventory tracking
- Simple dashboard and tax summary

---

## 📄 License

This is an open-source bookkeeping system. Feel free to use, modify, and distribute.

---

## 🎓 Learning Resources

### Understanding Financial Statements

**Profit & Loss (Income Statement)**:
- Shows profitability over a period
- Formula: Revenue - Expenses = Net Income

**Balance Sheet**:
- Shows financial position at a point in time
- Formula: Assets = Liabilities + Equity

**Cash Flow**:
- Shows cash movement
- Formula: Opening + Inflows - Outflows = Closing

### Key Accounting Principles

1. **Double-Entry**: Every transaction affects at least two accounts
2. **Accrual Basis**: Record when earned/incurred, not when cash changes hands
3. **Matching Principle**: Match expenses to the revenue they help generate
4. **Consistency**: Use same methods period to period

---

**Built with Python + openpyxl | Designed for Event Lighting Companies | QuickBooks-Level Functionality**
