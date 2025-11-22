# QuickStart Guide - Professional Bookkeeping System

## 🚀 Get Started in 3 Minutes

### What You Have

A **FREE QuickBooks-level bookkeeping system** worth $1,200/year in subscription costs!

✅ **14 Professional Sheets**  
✅ **Customer & Vendor Management**  
✅ **Financial Statements (P&L, Balance Sheet)**  
✅ **Bank Reconciliation**  
✅ **KPI Dashboard**  
✅ **Auto-Calculations & Validations**

---

## Step 1: Generate Your Workbook (30 seconds)

```bash
# Install dependency (one-time only)
pip3 install openpyxl

# Generate the workbook
python3 generate_workbook.py
```

**Output**: `Event_Lighting_Bookkeeping.xlsx` (28 KB)

---

## Step 2: Initial Setup (2 minutes)

### Open the File
- Double-click `Event_Lighting_Bookkeeping.xlsx`
- Opens directly to **Dashboard** sheet

### Customize Settings
1. **Unhide Settings Sheet**:
   - Right-click any sheet tab → Unhide → Select "Settings" → OK

2. **Update Your Details**:
   - Cell B3: Your Company Name
   - Cell B4: Your VAT Rate (currently 15%)
   - Cell B5: Financial Year Start Date
   - Cells B8-B9: Your Opening Cash Balances

3. **Hide Settings Again** (optional):
   - Right-click Settings tab → Hide

---

## Step 3: Start Using (Immediately!)

### Add Your First Customer
1. Go to **Customers** sheet
2. Row 5: Enter new customer details
3. System auto-calculates their balances

### Create Your First Invoice
1. Go to **Invoices** sheet
2. Row 3: 
   - Invoice No: INV-2025-002
   - Date: Today's date
   - Client Name: **Use dropdown** (from Customers)
   - Enter items, quantity, price
   - VAT and totals calculate automatically!

### Record Your First Expense
1. Go to **Expenses** sheet
2. Row 3:
   - Date: Today
   - Category: **Use dropdown** (Equipment, Fuel, etc.)
   - Vendor: **Use dropdown** (from Vendors)
   - Amount: Enter amount
   - Payment Method: **Use dropdown** (Cash, Bank, etc.)

### Check Your Dashboard
1. Go to **Dashboard** sheet
2. All metrics update **automatically in real-time**!

---

## What Makes This Professional?

### Like QuickBooks:
✅ Customer relationship management  
✅ Vendor tracking  
✅ Invoice generation with auto-calculations  
✅ Profit & Loss statement  
✅ Balance Sheet  
✅ Bank reconciliation  
✅ Tax reporting  

### Better Than QuickBooks:
✅ **FREE** (no $50-100/month subscription)  
✅ **Fully customizable** (change anything)  
✅ **Works offline** (always)  
✅ **Your data stays with you** (privacy)  
✅ **No learning curve** (it's Excel)

---

## Your Sheets Explained (1-Line Each)

| Sheet | What It Does |
|-------|--------------|
| **Dashboard** | Business health overview with KPIs and charts |
| **Customers** | Customer database with credit tracking |
| **Vendors** | Supplier database with purchase history |
| **Invoices** | Create invoices with auto-calculations |
| **Revenue** | Auto-syncs from invoices (no manual entry!) |
| **Expenses** | Track all expenses by category |
| **Inventory** | Manage equipment with stock valuation |
| **Trade Debtors** | See who owes you money (auto-updates) |
| **Cash Position** | Cash flow tracking (hand + bank) |
| **Bank Reconciliation** | Match your books with bank statements |
| **Profit & Loss** | Professional P&L statement (monthly + YTD) |
| **Balance Sheet** | Assets = Liabilities + Equity |
| **Tax Summary** | Year-end tax report with VAT calculations |
| **Invoice Template** | Print-ready invoice for customers |

---

## Pro Tips for Immediate Success

### 1. Always Use Dropdowns ⚠️
- Don't type customer/vendor names manually
- Click the dropdown arrow in cells with validation
- This prevents typos and ensures data consistency

### 2. Watch the Colors 🎨
- **Green** = Good (paid, healthy, balanced)
- **Red** = Alert (unpaid, overdue, problem)
- **Yellow** = Warning (approaching limit)

### 3. Data Flows Automatically 🔄
```
Invoices → Revenue → Dashboard → P&L → Balance Sheet
Expenses → Dashboard → P&L → Tax Summary
```
You only enter data in **Invoices** and **Expenses**. Everything else updates automatically!

### 4. Check Dashboard Daily 📊
- Open Dashboard sheet
- Review key metrics:
  - Are you profitable? (Net Profit should be green)
  - Do you have enough cash? (Total Cash)
  - Who owes you money? (Outstanding Invoices)
  - Is your Quick Ratio > 1.0? (Green = healthy business)

---

## Common First-Time Questions

**Q: Where do I enter revenue?**  
A: You don't! It auto-syncs from Invoices sheet. Just create invoices.

**Q: How do I print an invoice?**  
A: Go to Invoice Template sheet → Customize → File → Print (or Ctrl+P)

**Q: Can I change the VAT rate?**  
A: Yes! Unhide Settings → Change cell B4 → All invoices update automatically

**Q: What if I make a mistake?**  
A: Just edit the cell. All formulas update automatically. Use Ctrl+Z to undo.

**Q: How do I back up my data?**  
A: Copy the .xlsx file to Google Drive, Dropbox, or USB drive daily.

---

## Key Performance Indicators (KPIs)

Your dashboard tracks these automatically:

| KPI | What It Means | Target |
|-----|---------------|--------|
| **Days Sales Outstanding (DSO)** | How fast you collect payments | < 30 days |
| **Quick Ratio** | Can you pay bills? | > 1.0 |
| **Gross Profit Margin** | Profit per sale | > 40% |
| **Net Profit Margin** | Overall profitability | > 15% |

All color-coded: Green = Good, Red = Needs attention

---

## Monthly Routine (10 minutes)

### Last Day of Month:
1. **Bank Reconciliation**:
   - Go to Bank Reconciliation sheet
   - Enter bank statement balance
   - List outstanding items
   - Verify difference = 0 (should be green)

2. **Review Reports**:
   - Profit & Loss → Check profitability
   - Trade Debtors → Follow up on overdue invoices
   - Cash Position → Ensure adequate cash

3. **Customer Credit Review**:
   - Check Customers sheet for "Credit Warning" status
   - Follow up on high balances

---

## Troubleshooting

**Problem**: Dropdown lists don't work  
**Solution**: Make sure you have data in Customers/Vendors sheets first

**Problem**: Dashboard shows all zeros  
**Solution**: Enter some invoices and expenses first. Data needs to exist!

**Problem**: Formulas show #REF! error  
**Solution**: Don't delete rows from reference sheets. Hide them instead.

**Problem**: VAT calculations seem wrong  
**Solution**: Check Settings sheet cell B4 for correct VAT rate percentage

---

## Next Steps

### Now That You're Running:
1. ✅ Read `README_PROFESSIONAL.md` for complete documentation
2. ✅ Add all your customers and vendors
3. ✅ Enter historical invoices (if needed)
4. ✅ Set up monthly bank reconciliation routine
5. ✅ Bookmark the Dashboard sheet

### Want to Customize?
- See `README_PROFESSIONAL.md` → "Customization" section
- Change colors, add categories, modify layouts
- It's your system—make it perfect for your business!

---

## Support

**Documentation**:
- `README_PROFESSIONAL.md` - Complete 700+ line guide
- `README.md` - Original documentation
- `READ.md` - Quick reference

**Learning**:
- Play with sample data first
- Can't break anything—just regenerate the file!
- Excel functions are standard—Google them if stuck

---

## Comparison: This vs Paid Software

| Feature | This System | QuickBooks Online | Xero |
|---------|-------------|-------------------|------|
| Monthly Cost | **$0** | $50-100 | $40-70 |
| Yearly Cost | **$0** | $600-1,200 | $480-840 |
| Customer Management | ✅ | ✅ | ✅ |
| Invoicing | ✅ | ✅ | ✅ |
| Financial Reports | ✅ | ✅ | ✅ |
| Bank Reconciliation | ✅ | ✅ | ✅ |
| Customizable | **100%** | Limited | Limited |
| Works Offline | **Always** | No | No |
| Your Data | **Yours** | Their servers | Their servers |
| Learning Curve | Low | Medium | Medium |

**You save $600-1,200 per year!** 💰

---

## You're All Set! 🎉

**Your professional bookkeeping system is ready.**

1. Generate the workbook (done ✅)
2. Open and customize (2 minutes)
3. Start entering data (immediately)
4. Watch everything calculate automatically (magic ✨)

Questions? Check `README_PROFESSIONAL.md` for the complete guide.

---

**Built with ❤️ for Event Lighting Companies | Free & Open Source**
