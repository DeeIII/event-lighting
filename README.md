# Event Lighting Company Bookkeeping System

A comprehensive Excel workbook designed for event lighting companies to manage bookkeeping, accounting, and end-of-year tax reporting.

## 📊 Overview

This fully-functional Excel workbook provides an integrated solution for:
- Revenue tracking and invoice management
- Expense monitoring and categorization
- Inventory/stock management
- Cash flow tracking
- Accounts receivable management
- Tax reporting and compliance
- Business health dashboard with visual analytics

## 📁 File Structure

- **Event_Lighting_Bookkeeping.xlsx** - The main Excel workbook
- **generate_workbook.py** - Python script to regenerate the workbook if needed
- **README.md** - This documentation file

## 📋 Workbook Sheets

### 1. **Dashboard** (Business Health)
Your command center showing real-time business metrics:
- Monthly and YTD revenue, expenses, and profit
- Profit margins
- Cash position (cash in hand and at bank)
- Outstanding invoices
- Total stock value
- Revenue vs Expense trend chart

**How to use:** This sheet auto-updates based on data entered in other sheets. Use it for quick business health checks.

### 2. **Invoices** (Invoice Generator)
Central invoice management system with automatic calculations.

**Columns:**
- Invoice No., Date, Client Name, Event Date
- Items/Equipment, Quantity, Unit Price
- Subtotal, VAT Rate, VAT Amount (auto-calculated)
- Total Amount, Amount Received, Balance (auto-calculated)

**How to use:**
1. Enter invoice number, dates, client name
2. Describe items/equipment provided
3. Enter quantity and unit price
4. VAT and totals calculate automatically
5. Enter amount received - balance auto-calculates

### 3. **Revenue** (Cash Inflow)
Automatically pulls invoice data and tracks revenue.

**Features:**
- Auto-syncs with Invoices sheet
- Daily, monthly, and yearly revenue summaries
- Outstanding balance tracking

**How to use:** This sheet auto-populates from the Invoices sheet. No manual entry needed.

### 4. **Expenses**
Track all business expenses with categorization.

**Columns:**
- Date, Expense Category, Vendor, Description
- Amount, Payment Method, Reference

**Features:**
- Automatic summaries by category
- Monthly and yearly expense totals
- Category-wise breakdowns (YTD)

**How to use:**
1. Enter date and select category from predefined list
2. Enter vendor name and description
3. Enter amount and payment method
4. Add reference number for tracking

**Expense Categories:**
- Equipment Purchase
- Equipment Maintenance
- Transport/Fuel
- Salaries/Wages
- Rent
- Utilities
- Insurance
- Marketing
- Office Supplies
- Professional Fees
- Bank Charges
- Other

### 5. **Inventory** (Stock Management)
Complete inventory tracking with automatic valuations.

**Columns:**
- Stock ID, Item Description, Unit Price
- Quantity in Store, Quantity Rented Out, Stock in Transit
- Total Quantity (auto-calculated)
- Total Stock Value (auto-calculated)

**Features:**
- Real-time stock availability
- Automatic stock valuation
- Summary of available, rented, and in-transit items

**How to use:**
1. Each item has a unique Stock ID
2. Update quantities as items are rented or returned
3. Total value calculates automatically

### 6. **Trade Debtors** (Accounts Receivable)
Automatically tracks unpaid invoices and aging.

**Features:**
- Auto-populates from Invoices where Balance > 0
- Calculates days outstanding
- Status classification (Current, Due Soon, Overdue)
- Summary by aging category

**How to use:** This sheet auto-updates when invoices have outstanding balances. No manual entry needed.

**Status Categories:**
- **Current**: 0-30 days
- **Due Soon**: 31-60 days
- **Overdue**: >60 days

### 7. **Cash Position**
Real-time cash flow statement.

**Sections:**
- **Cash in Hand**: Opening balance + cash receipts - cash payments
- **Cash at Bank**: Opening balance + bank receipts - bank payments
- **Total Cash Position**: Combined cash summary
- **Monthly Cash Flow**: Month-by-month inflow/outflow breakdown

**How to use:** This sheet auto-calculates based on Revenue and Expenses. Review regularly to monitor liquidity.

### 8. **Tax Summary**
End-of-year tax reporting summary.

**Includes:**
- Total revenue (invoiced and received)
- Outstanding receivables
- VAT collected on sales
- VAT paid on purchases
- Net VAT payable
- Expenses by category (deductible)
- Net profit before tax
- Profit margin percentage

**How to use:** Use this sheet for annual tax filing and audit preparation. All values auto-populate from other sheets.

### 9. **Settings** (Hidden Sheet)
Configuration and reference data.

**Contains:**
- Company name: Flash Lighting Services Ltd
- VAT rate: 15%
- Financial year start date
- Opening balances (cash in hand and at bank)
- Expense categories list
- Payment methods list

**How to customize:**
1. Unhide the Settings sheet: Right-click any sheet tab → Unhide → Settings
2. Update company name, VAT rate, and opening balances as needed
3. Hide the sheet again when done

## 🚀 Getting Started

### Initial Setup

1. **Unhide and Configure Settings Sheet**
   - Right-click any sheet tab → Unhide → Settings
   - Update company name (Cell B3)
   - Verify/update VAT rate (Cell B4) - currently set to 15%
   - Update financial year start date (Cell B5)
   - Enter your opening balances:
     - Cash in Hand Opening (Cell B8)
     - Cash at Bank Opening (Cell B9)
   - Hide the sheet again: Right-click Settings tab → Hide

2. **Start Entering Data**
   - Begin with the **Invoices** sheet to record your first invoice
   - Add expenses in the **Expenses** sheet
   - Update inventory quantities in the **Inventory** sheet

3. **Monitor Your Business**
   - Check the **Dashboard** regularly for business health
   - Review **Cash Position** weekly for liquidity
   - Monitor **Trade Debtors** for collections

## 💡 Tips for Best Use

### Daily Operations
- Enter all invoices immediately in the Invoices sheet
- Record expenses as they occur
- Update inventory after each event/rental

### Weekly Tasks
- Review Trade Debtors for overdue invoices
- Follow up on outstanding payments
- Check Cash Position to ensure adequate liquidity

### Monthly Tasks
- Review Dashboard metrics
- Analyze Revenue vs Expense trends
- Reconcile bank statements
- Update inventory counts

### Quarterly/Annual Tasks
- Review Tax Summary for compliance
- Analyze profit margins by quarter
- Plan for tax payments
- Conduct full inventory audit

## 🔧 Maintenance

### Backing Up Data
- Save regular backups of the workbook
- Use version control (e.g., Event_Lighting_Bookkeeping_2025-01.xlsx)
- Consider cloud storage for redundancy

### Adding More Data Rows
When you run out of rows for Invoices, Expenses, or other sheets:
1. Insert new rows between the header and existing data
2. Copy formulas from the row above
3. Ensure cell references in formulas still work correctly

### Year-End Procedures
1. Review Tax Summary sheet for completeness
2. Export or print all sheets for records
3. Create a new workbook for the new financial year
4. Update Settings sheet with new opening balances
5. Archive the previous year's workbook

## 📊 Formula Features

The workbook uses advanced Excel formulas including:
- **SUMIFS** - Conditional summation for date and category filtering
- **VLOOKUP/INDEX-MATCH** - Cross-sheet data linking
- **IFERROR** - Clean error handling
- **DATE functions** - Automatic date-based calculations
- **Conditional formulas** - Dynamic status and aging calculations

All sheets are interconnected through formulas - **do not manually enter data that should auto-populate**.

## ⚠️ Important Notes

### Do's ✅
- ✅ Always enter dates in YYYY-MM-DD format
- ✅ Use consistent invoice numbering (e.g., INV-2025-001)
- ✅ Select expense categories from the predefined list
- ✅ Update inventory quantities regularly
- ✅ Save frequently while working

### Don'ts ❌
- ❌ Don't manually edit the Revenue sheet (auto-populated from Invoices)
- ❌ Don't manually edit Trade Debtors (auto-populated based on balances)
- ❌ Don't delete the Settings sheet
- ❌ Don't modify formulas unless you understand Excel well
- ❌ Don't use special characters in client names or descriptions

## 🔄 Regenerating the Workbook

If you need to regenerate the workbook from scratch:

```bash
python3 generate_workbook.py
```

**Prerequisites:**
```bash
pip3 install openpyxl
```

## 📞 Support & Customization

To customize this workbook for your specific needs:
1. Make a backup copy first
2. Unhide the Settings sheet to modify categories or configurations
3. Adjust formulas carefully if needed
4. Test with sample data before using with real data

## 📄 License

This workbook is provided as-is for internal business use.

## 🎯 Audit Readiness

This workbook is designed for audit readiness with:
- Complete audit trail through linked sheets
- Automatic calculations reduce manual errors
- VAT tracking for compliance
- Comprehensive expense categorization
- Clear financial summaries
- Historical data preservation

---

**Version:** 1.0  
**Created:** 2025  
**Last Updated:** 2025-01-15

For questions or issues, refer to this documentation or consult with your accountant for business-specific requirements.
