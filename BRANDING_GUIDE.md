# Company Branding Guide

## Your Workbook is Now Branded! 🎨

Your bookkeeping system now includes:
- ✅ **Company Name** throughout all sheets
- ✅ **Company Logo** on Dashboard and Invoice Template
- ✅ **Contact Information** displayed professionally
- ✅ **Brand Colors** used consistently

---

## Current Branding Configuration

Your system is configured with:

```python
COMPANY_NAME = "Flash Illumination Services Ltd"
COMPANY_TAGLINE = "Professional Event Lighting Solutions"
COMPANY_EMAIL = "flashillumination.com"
COMPANY_PHONE = "+234 7032728785"
COMPANY_ADDRESS = "Lekki Gardens PhaseII, Lagos, Nigeria"
COMPANY_WEBSITE = "www.flashillumination.com"
COMPANY_IG = "@flashillumination"
LOGO_FILENAME = "flash21-1.jpg"
```

Logo Size: 140 x 80 pixels

---

## How to Customize Your Branding

### 1. Change Company Information

Edit `generate_workbook.py` at lines 24-30:

```python
COMPANY_NAME = "Your Company Name Here"
COMPANY_TAGLINE = "Your Tagline"
COMPANY_EMAIL = "your@email.com"
COMPANY_PHONE = "+234 123 456 7890"
COMPANY_ADDRESS = "Your Address"
COMPANY_WEBSITE = "www.yourwebsite.com"
COMPANY_IG = "@yourinstagram"
```

### 2. Change/Add Your Logo

**Step 1**: Prepare your logo file
- Supported formats: PNG, JPG, or GIF
- Recommended size: 120-150 pixels wide, 60-100 pixels tall
- Place the file in the same folder as `generate_workbook.py`

**Step 2**: Update logo settings (lines 35-37):
```python
LOGO_FILENAME = "your_logo.png"  # Change to your filename
LOGO_WIDTH = 140  # Adjust width in pixels
LOGO_HEIGHT = 80  # Adjust height in pixels
```

**Step 3**: Regenerate
```bash
python3 generate_workbook.py
```

### 3. Change Brand Colors

Edit lines 40-42 in `generate_workbook.py`:

```python
BRAND_COLOR_PRIMARY = "203864"    # Dark blue (hex without #)
BRAND_COLOR_SECONDARY = "366092"  # Medium blue
BRAND_COLOR_ACCENT = "4472C4"     # Light blue
```

**Popular Color Schemes**:
- **Corporate Blue**: Primary="1F4E78", Secondary="2E5C8A", Accent="3D6FA6"
- **Professional Green**: Primary="2C5F2D", Secondary="3F7F40", Accent="52A553"
- **Modern Purple**: Primary="5B2C6F", Secondary="7D3C98", Accent="A04CB8"
- **Elegant Gray**: Primary="2C3E50", Secondary="34495E", Accent="546E7A"

To find hex colors: Google "color picker" or use any design tool

---

## Where Your Branding Appears

### Dashboard
- 🎯 **Logo** in top-left corner (if provided)
- 🎯 **Company Name** as main title
- 🎯 **Brand colors** in headers

### Invoice Template
- 🎯 **Logo** prominently displayed
- 🎯 **Company Name** and tagline
- 🎯 **Full contact information**
- 🎯 **Professional layout** for printing

### Settings Sheet (Hidden)
- 🎯 All company details stored here
- 🎯 Other sheets reference this data
- 🎯 Change once, updates everywhere!

### All Report Sheets
- 🎯 **Brand colors** in all headers
- 🎯 **Consistent styling** throughout

---

## Pro Tips for Professional Branding

### Logo Best Practices

1. **File Format**:
   - PNG with transparent background (best for professional look)
   - JPG works fine too
   - Avoid very large files (keep under 500KB)

2. **Size Guidelines**:
   - **Width**: 100-150 pixels (larger looks unprofessional)
   - **Height**: 50-100 pixels
   - **Aspect Ratio**: Keep your logo's original proportions

3. **Quality**:
   - Use high-resolution logo (300 DPI if possible)
   - Avoid pixelated or blurry images
   - Test how it looks when printed

### Color Scheme Tips

1. **Professional Colors**:
   - Blues → Trust, reliability (great for finance)
   - Greens → Growth, stability
   - Grays → Professional, neutral
   - Purples → Creative, innovative

2. **Contrast**:
   - Ensure text is readable on colored backgrounds
   - Dark colors work well for headers
   - Light colors for accents

3. **Consistency**:
   - Use same colors as your business cards
   - Match your website color scheme
   - 2-3 colors maximum for clean look

---

## Troubleshooting

### Logo Not Showing

**Problem**: "Logo not found" warning
**Solution**:
1. Check logo file is in same folder as `generate_workbook.py`
2. Verify filename exactly matches `LOGO_FILENAME` (case-sensitive!)
3. Check file extension (.png, .jpg, .gif)

**Problem**: "Could not add logo" error
**Solution**:
```bash
pip3 install Pillow
```

**Problem**: Logo appears distorted
**Solution**:
- Adjust `LOGO_WIDTH` and `LOGO_HEIGHT` to maintain original aspect ratio
- If original is 300x150, use WIDTH=140, HEIGHT=70

### Colors Not Updating

**Problem**: New colors don't show
**Solution**:
1. Make sure to save `generate_workbook.py` after editing
2. Regenerate the workbook: `python3 generate_workbook.py`
3. Close old workbook before opening new one

---

## Quick Rebranding Workflow

Want to rebrand for a different company? Here's the 2-minute process:

```bash
# 1. Edit company info (lines 24-30)
# 2. Update logo filename (line 35)
# 3. Adjust brand colors (lines 40-42)

# 4. Regenerate
python3 generate_workbook.py

# 5. Done! Your branded workbook is ready
```

---

## Examples of Branding Customization

### Example 1: Tech Startup

```python
COMPANY_NAME = "TechVibe Solutions"
COMPANY_TAGLINE = "Innovation at Speed"
LOGO_FILENAME = "techvibe_logo.png"
BRAND_COLOR_PRIMARY = "1A237E"  # Deep Blue
BRAND_COLOR_SECONDARY = "3949AB"  # Medium Blue
BRAND_COLOR_ACCENT = "5C6BC0"  # Light Blue
```

### Example 2: Eco-Friendly Business

```python
COMPANY_NAME = "GreenLeaf Enterprises"
COMPANY_TAGLINE = "Sustainable Solutions"
LOGO_FILENAME = "greenleaf.png"
BRAND_COLOR_PRIMARY = "1B5E20"  # Dark Green
BRAND_COLOR_SECONDARY = "2E7D32"  # Medium Green
BRAND_COLOR_ACCENT = "43A047"  # Bright Green
```

### Example 3: Creative Agency

```python
COMPANY_NAME = "Pixel Perfect Studio"
COMPANY_TAGLINE = "Design. Create. Inspire."
LOGO_FILENAME = "pixel_logo.png"
BRAND_COLOR_PRIMARY = "4A148C"  # Deep Purple
BRAND_COLOR_SECONDARY = "6A1B9A"  # Medium Purple
BRAND_COLOR_ACCENT = "8E24AA"  # Light Purple
```

---

## Advanced Customization

### Adding More Contact Fields

Want to add more fields to Settings sheet? Edit `create_settings_sheet()` function:

```python
ws['A10'] = "Fax"
ws['A10'].font = Font(bold=True)
ws['B10'] = "+234 800 000 0001"
```

### Changing Logo Position

Logos currently appear at:
- Dashboard: Cell A1
- Invoice Template: Cell A1

To change position, edit the sheet creation functions and update:
```python
add_logo_to_sheet(ws, 'B1')  # Change from A1 to B1
```

---

## FAQ

**Q: Can I use different logos on different sheets?**
A: Yes! When calling `add_logo_to_sheet()`, pass a different filename:
```python
add_logo_to_sheet(ws, 'A1', 'alternative_logo.png')
```

**Q: How do I remove the logo?**
A: Either:
1. Delete/rename the logo file, or
2. Set `include_logo=False` when calling `add_company_header()`

**Q: Can I change logo size after generating?**
A: No, you must regenerate the workbook. Just change `LOGO_WIDTH` and `LOGO_HEIGHT`, then run the script again.

**Q: Will branding affect formulas?**
A: No! All formulas still work. Branding is visual only.

---

## Need Help?

Your company branding is now fully configured and working! 

**Current Status**:
✅ Company: Flash Illumination Services Ltd
✅ Logo: flash21-1.jpg (140x80 pixels)
✅ All branding applied successfully

To make changes:
1. Edit `generate_workbook.py` (lines 24-42)
2. Run `python3 generate_workbook.py`
3. Your new branded workbook is ready!

---

**Professional Bookkeeping System with Full Company Branding** 🎨
