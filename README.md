# Disney Magnet Order Processor

## 🎨 Two Ways to Use

### Option 1: GUI Application (Recommended) ⭐
**Easy, visual, drag-and-drop interface**
- Double-click `launch_gui.bat` (Windows) or `launch_gui.sh` (Mac/Linux)
- Drag & drop your CSV file
- Watch real-time progress
- View results instantly

### Option 2: Command Line
**Fast, scriptable, automation-friendly**
- Run: `python process_orders.py your_orders.csv`
- Perfect for batch processing
- Can be automated

---

## 📦 What's Included

This package contains everything you need to process Disney character magnet orders from CSV files to final PDFs.

### Core Files:
- **gui_app.py** - 🎨 GUI Application (NEW!)
- **launch_gui.bat/sh** - Easy GUI launcher
- **process_orders.py** - Command-line script
- **format.pdf** - PDF template for output
- **requirements.txt** - Python dependencies
- **font/** - Required fonts (Waltograph, Blueberry)

### Documentation:
- **sample_order.csv** - Example order file
- **template_order.csv** - Blank template to copy
- **QUICKSTART.md** - Quick 2-minute setup guide
- **PROCESS_ORDERS_README.md** - Complete documentation
- **SOLUTION_SUMMARY.md** - Technical details

## 🚀 Quick Setup

### 1. Extract this folder
Extract to any location on your computer (e.g., `c:\orders\`)

### 2. Install Python dependencies
```bash
pip install Pillow PyPDF2 requests
```
Or use the requirements file:
```bash
pip install -r requirements.txt
```

### 3. Set up API Key (for AI parsing feature)
**Optional but recommended for AI parsing**

Create `C:\source\grok_config.txt` with your Grok API key.

See [API_KEY_SETUP.md](API_KEY_SETUP.md) for detailed instructions.

**Note:** You can still use the app without this - just no AI parsing!

### 3. Set up FHM_Images folder
**IMPORTANT**: Place your `FHM_Images` folder in the **parent directory**

Example folder structure:
```
c:\orders\                    ← Parent directory
├── FHM_Images\              ← Your character images go HERE
│   ├── mickey-captain.png
│   ├── minnie-captain.png
│   └── ...
└── canva\                   ← This extracted folder (rename if needed)
    ├── process_orders.py
    ├── format.pdf
    ├── font\
    └── ...
```

### 4. Create your orders CSV
Use Excel, Google Sheets, or any text editor:

```csv
character,name
mickey-captain,Johnny
minnie-captain,Sarah
donald-normal,Mike
```

**Important**: Character names must match image filenames (without .png)

### 5. Run the application

**OPTION A: GUI (Recommended)**
```bash
# Windows
Double-click launch_gui.bat

# Mac/Linux
chmod +x launch_gui.sh
./launch_gui.sh
```

**OPTION B: Command Line**
```bash
python process_orders.py your_orders.csv
```

### 6. Get your results
- **Images**: Check `outputs/` folder for personalized images
- **PDFs**: Check current directory for `order_output_*.pdf` files

## 📋 CSV Format

Your CSV file needs two columns:

1. **character** - The image filename without .png
2. **name** - Personalization name (or empty for no text)

### Example:
```csv
character,name
mickey-captain,John Smith
minnie-normal,Mary
donald-pumpkin,
goofy-captain,Bob Jones
```

### Character Name Format:
Character names typically follow the pattern: `charactername-variant`

Common variants:
- `-normal` (standard version)
- `-captain` (Captain theme)
- `-pumpkin`, `-witch`, `-vampire`, `-mummy` (Halloween)
- `-halloween` (generic Halloween)

**Tip**: Look at your image filenames in the FHM_Images folder and use those exact names!

## 📖 Documentation

Three levels of help:

1. **QUICKSTART.md** - Fast reference (2 minutes)
2. **PROCESS_ORDERS_README.md** - Complete guide with examples
3. **SOLUTION_SUMMARY.md** - Technical details

## 🧪 Test It

Try the included sample file:
```bash
python process_orders.py sample_order.csv
```

This will process 10 sample orders (if you have the matching images).

## ⚠️ Troubleshooting

### "Images folder not found"
→ Make sure `FHM_Images` is in the **parent directory**, not inside this folder

### "Image not found for character 'mickey-captain'"
→ Check that `mickey-captain.png` exists in FHM_Images folder
→ Character names are case-sensitive

### "Template PDF not found"
→ Make sure `format.pdf` is in the same folder as the script
→ Run the script from this folder

### "Font not found"
→ Make sure the `font/` folder is present with all font files

## 💡 Tips

- **Test first**: Run with a small CSV file to verify everything works
- **Check images**: Verify character names match your image files exactly
- **Batch processing**: Put all orders in one CSV file for efficient processing
- **Reprocess orders**: Keep your CSV files to easily regenerate orders later

## 🆘 Need Help?

Check the detailed documentation files:
- **QUICKSTART.md** for quick answers
- **PROCESS_ORDERS_README.md** for step-by-step guides
- **SOLUTION_SUMMARY.md** for technical information

## ⚙️ System Requirements

- Python 3.7+
- Windows, Mac, or Linux
- Pillow (PIL) library
- PyPDF2 library

## 📝 License

This tool is provided for processing Disney-themed magnet orders.

---

**Happy processing! 🎉**

For best results, start with QUICKSTART.md and refer to PROCESS_ORDERS_README.md for detailed information.


