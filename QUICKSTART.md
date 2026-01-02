# Quick Start Guide - Process Orders Script

## 🚀 Fast Setup (2 minutes)

### 1. Install Dependencies
```bash
pip install Pillow PyPDF2
```

### 2. Check Your Folder Structure
```
c:\dev\
├── FHM_Images\              ← Your character images go here
│   ├── mickey-captain.png
│   ├── minnie-captain.png
│   └── ...
└── canva\                   ← Your working directory
    ├── process_orders.py    ← The script
    ├── format.pdf           ← Template
    └── font\                ← Fonts folder
```

### 3. Create Your CSV
Create a file called `orders.csv`:
```csv
character,name
mickey-captain,Johnny
minnie-captain,Sarah
donald-normal,Mike
```

**Important**: Character names must match the image filenames exactly (without .png)

### 4. Run It!
```bash
python process_orders.py orders.csv
```

### 5. Get Your Results
- **Personalized images**: `outputs/1.png`, `2.png`, etc.
- **PDFs**: `order_output_*.pdf` files in current directory

## 📋 CSV Format Cheat Sheet

```csv
character,name
mickey-captain,John        ← Full name with character variant
minnie-normal,Sarah        ← Normal variant
donald-pumpkin,Mike        ← Halloween variant
daisy-normal,              ← Empty name = no personalization
goofy-captain,Lisa
```

### Character Name Format:
- Format: `charactername-variant`
- Examples:
  - `mickey-captain`
  - `mickey-normal`
  - `mickey-pumpkin`
  - `mickey-halloween`
  - `stitch-normal`
  - `belle-normal`

**Tip**: Look at your image filenames in FHM_Images folder and use those names (without .png)

## ⚠️ Common Issues

### "Images folder not found"
→ Make sure `FHM_Images` is in `c:\dev\` (parent directory)

### "Image not found for character"
→ Check image filename matches exactly what's in your CSV

### "Template PDF not found"
→ Make sure `format.pdf` is in the same folder as the script

## 💡 Pro Tips

- **Reprocess old orders**: Just create a CSV from your records and run again
- **Bulk processing**: Put all orders in one CSV file
- **No name needed**: Leave name column empty for non-personalized items
- **Preview before printing**: Check images in `outputs/` folder first

## 🎯 Full Example

```bash
# Step 1: Create orders.csv
echo character,name > orders.csv
echo mickey-captain,Johnny >> orders.csv
echo minnie-captain,Sarah >> orders.csv

# Step 2: Run
python process_orders.py orders.csv

# Step 3: Done! Check your PDFs
```

## 📖 Need More Help?

See `PROCESS_ORDERS_README.md` for detailed documentation.

