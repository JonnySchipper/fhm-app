# Process Orders Script - User Guide

## Overview

`process_orders.py` is a standalone script that processes Disney character magnet orders from start to finish. No UI, no server, no API calls required - just run it and get your personalized PDFs!

## What It Does

1. **Reads a CSV file** with character names and personalization names
2. **Finds the matching images** in the FHM_Images folder
3. **Generates personalized images** by adding names to the character images
4. **Creates PDF files** with pairs of images ready for printing

## Requirements

### Files Needed

- **CSV file**: Your order list (see format below)
- **FHM_Images folder**: Must be in the parent directory (e.g., `c:\dev\FHM_Images\`)
- **format.pdf**: Template PDF file in the current directory
- **font folder**: Contains the required fonts (waltographUI.ttf, blueberry.ttf)

### Python Packages

- PIL (Pillow)
- PyPDF2

Install with:
```bash
pip install Pillow PyPDF2
```

## CSV File Format

Create a CSV file with two columns:

1. **Character column**: Character name WITHOUT the .png extension
2. **Name column**: Personalization name (can be empty for no text)

### Example CSV:

```csv
character,name
mickey-captain,Johnny
minnie-captain,Sarah
donald-normal,Michael
daisy-normal,
goofy-pumpkin,Oliver
stitch-normal,Liam
```

### Important Notes:

- **Character names must match exactly** the image filenames (without .png)
  - If image is `mickey-captain.png`, use `mickey-captain`
  - If image is `stitch-normal.png`, use `stitch-normal`
  
- **Leave name column empty** for images without personalization
  
- **Character name format**: Usually `character-variant`
  - Examples: `mickey-normal`, `mickey-captain`, `mickey-halloween`, `mickey-pumpkin`
  - For dog/rubber variants: `dog-30`, `rubber-duck`, etc.

## Usage

### Basic Usage

```bash
python process_orders.py your_order_file.csv
```

### Step-by-Step Example

1. **Create your CSV file** (e.g., `my_order.csv`):
   ```csv
   character,name
   mickey-captain,John Smith
   minnie-captain,Mary Jones
   donald-normal,Bob
   ```

2. **Make sure your folder structure is correct**:
   ```
   c:\dev\
   ├── FHM_Images\           ← Images must be here
   │   ├── mickey-captain.png
   │   ├── minnie-captain.png
   │   └── donald-normal.png
   └── canva\                ← Your working directory
       ├── process_orders.py
       ├── format.pdf
       ├── my_order.csv
       └── font\
           ├── waltographUI.ttf
           └── blueberry.ttf
   ```

3. **Run the script**:
   ```bash
   cd c:\dev\canva
   python process_orders.py my_order.csv
   ```

4. **Check the output**:
   - Personalized images: `outputs/1.png`, `outputs/2.png`, etc.
   - PDFs: `order_output_20231215_143022_1.pdf`, etc.

## Output Files

### Personalized Images

- Saved to `outputs/` folder
- Named sequentially: `1.png`, `2.png`, `3.png`, etc.
- Each image has the personalization name curved along the bottom

### PDF Files

- Saved to current directory
- Named: `order_output_YYYYMMDD_HHMMSS_N.pdf`
- Each PDF contains 2 images (top and bottom positions)
- If you have an odd number of images, the last one won't have a PDF

### Example Output:

```
Order with 6 images:
├── outputs/
│   ├── 1.png
│   ├── 2.png
│   ├── 3.png
│   ├── 4.png
│   ├── 5.png
│   └── 6.png
└── Current directory:
    ├── order_output_20231215_143022_1.pdf  (images 1 & 2)
    ├── order_output_20231215_143022_2.pdf  (images 3 & 4)
    └── order_output_20231215_143022_3.pdf  (images 5 & 6)
```

## Troubleshooting

### "Images folder not found"

**Problem**: Script can't find FHM_Images folder

**Solution**: Make sure FHM_Images is in the parent directory:
```
c:\dev\FHM_Images\     ← Must be here
c:\dev\canva\          ← Working directory
```

### "Image not found for character"

**Problem**: Character name in CSV doesn't match image filename

**Solution**: 
1. Check the exact filename in FHM_Images folder
2. Use the EXACT name (without .png) in your CSV
3. Check for typos and case sensitivity

### "Template PDF not found"

**Problem**: format.pdf is missing

**Solution**: Make sure `format.pdf` is in the same directory as the script

### "Font not found" or "Text looks wrong"

**Problem**: Font files are missing or in wrong location

**Solution**: Ensure you have a `font/` folder with:
- waltographUI.ttf
- blueberry.ttf (for dog/rubber variants)

## Advanced Tips

### Processing Large Orders

The script handles any number of orders. For very large batches:
- Process in groups (e.g., 50 at a time)
- Check outputs/ folder regularly
- PDFs are timestamped so they won't overwrite each other

### No Personalization Orders

Leave the name column empty for items without personalization:

```csv
character,name
mickey-normal,
donald-normal,
```

### Duplicate Names

The script handles duplicate names fine. Each will get its own image:

```csv
character,name
mickey-captain,Johnny
minnie-captain,Johnny
goofy-captain,Johnny
```

### Mixed Themes in One Order

You can mix different character variants in the same CSV:

```csv
character,name
mickey-captain,John
mickey-halloween,Sarah
mickey-pumpkin,Mike
mickey-normal,Lisa
```

## Comparison with Web Interface

| Feature | process_orders.py | Web Interface |
|---------|------------------|---------------|
| Setup | None - just run it | Need to start server |
| Input | CSV file | Paste order text |
| AI Matching | No (manual CSV) | Yes (Grok API) |
| Speed | Very fast | Depends on API |
| Control | Full control | AI decides matches |
| Best for | Known orders, bulk processing | New orders, complex parsing |

## When to Use This Script

✅ **Use process_orders.py when:**
- You know exactly which characters you need
- You're reprocessing previous orders
- You want maximum control
- You want fastest processing
- You don't need AI text parsing

❌ **Use the web interface when:**
- You have raw customer order text
- Character names need parsing/interpretation
- You need to handle typos/misspellings
- Orders have complex formatting

## Tips for Creating CSV Files

### Excel/Google Sheets Method

1. Create a spreadsheet:
   | character | name |
   |-----------|------|
   | mickey-captain | Johnny |
   | minnie-captain | Sarah |

2. Save/Export as CSV

### Quick Text Editor Method

Just create a plain text file with `.csv` extension:
```
character,name
mickey-captain,Johnny
minnie-captain,Sarah
```

### From Previous Orders

If you have previous order records, you can:
1. Extract character and name data
2. Format as CSV
3. Reprocess anytime!

## Support

If you encounter issues:

1. Check the error message - it usually tells you exactly what's wrong
2. Verify your folder structure matches the requirements
3. Check CSV format (commas, no extra spaces)
4. Make sure character names match image filenames exactly
5. Ensure all required files (fonts, template PDF) are present

## Example Workflow

Here's a complete workflow from start to finish:

```bash
# 1. Navigate to your working directory
cd c:\dev\canva

# 2. Create your CSV file (using Excel, Notepad, or any editor)
# Save as: my_orders.csv

# 3. Run the script
python process_orders.py my_orders.csv

# 4. Check the output
# - Look in outputs/ for personalized images
# - Look in current directory for PDFs
# - Print or send PDFs as needed

# 5. All done! 🎉
```

## License

This script is part of the Disney magnet order processing system.

