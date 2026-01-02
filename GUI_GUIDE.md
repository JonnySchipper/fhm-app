# 🎨 GUI Application Guide

## Quick Start

### Windows
1. Double-click `launch_gui.bat`
2. Wait for the application to start
3. Start processing!

### Mac/Linux
1. Open terminal in this folder
2. Run: `chmod +x launch_gui.sh`
3. Run: `./launch_gui.sh`
4. Start processing!

---

## 🖥️ Interface Overview

### 📁 File Selection Section
- **Drag & Drop**: Simply drag your CSV file onto the drop zone
- **Browse Button**: Click to select a file using file picker
- **File Path**: Shows the currently loaded file

### 📋 Order Preview Section
- **Order Count**: Shows how many orders are loaded
- **Preview Table**: Displays all character-name pairs
- Scroll to review all orders before processing

### ⚙️ Processing Section
- **Progress Bar**: Visual progress indicator
- **Processing Log**: Real-time output and status messages
- Shows detailed progress for each step

### 🎮 Control Buttons

#### ▶ Process Orders
- Main action button (green)
- Starts processing the loaded orders
- Disabled until a CSV file is loaded
- Shows real-time progress

#### 📁 View Outputs
- Opens the outputs folder
- Only enabled after successful processing
- Quick access to generated images

#### 🗑️ Clear
- Clears all fields and resets the interface
- Start fresh with a new order

#### ❓ Help
- Opens quick help dialog
- Shows CSV format and usage tips

---

## 📋 How to Use - Step by Step

### Step 1: Prepare Your CSV
Create a CSV file with two columns:
```csv
character,name
mickey-captain,Johnny
minnie-captain,Sarah
donald-normal,Mike
```

### Step 2: Load the CSV
**Option A - Drag & Drop**:
- Drag your CSV file from file explorer
- Drop it onto the "📂 Drag & Drop" area

**Option B - Browse**:
- Click the "Browse..." button
- Select your CSV file
- Click "Open"

### Step 3: Review Orders
- Check the Order Preview section
- Verify all characters and names are correct
- Order count should match your expectations

### Step 4: Process
- Click the "▶ Process Orders" button
- Watch the progress bar and log output
- Processing typically takes a few seconds

### Step 5: View Results
- Click "📁 View Outputs" button to open the outputs folder
- Check the personalized images (1.png, 2.png, etc.)
- PDFs will be in the main folder

---

## 🎯 Features

### ✨ Drag & Drop Support
- Drop CSV files directly onto the application
- Automatic file validation
- Visual feedback

### 📊 Real-time Progress
- Live progress bar updates
- Detailed processing logs
- Step-by-step status messages

### 🎨 User-Friendly Design
- Clean, modern interface
- Color-coded status messages
- Intuitive button layout
- Helpful tooltips and labels

### ⚡ Fast & Responsive
- Processing runs in background
- UI stays responsive during processing
- No freezing or hanging

### 🔍 Order Preview
- See all orders before processing
- Verify character names
- Check personalization names
- Catch errors early

### 📝 Comprehensive Logging
- Timestamped log entries
- Color-coded messages:
  - ℹ Blue: Info
  - ✓ Green: Success
  - ⚠ Yellow: Warning
  - ❌ Red: Error

---

## 💡 Pro Tips

### Tip 1: Verify Before Processing
Always review the Order Preview before clicking Process. This helps catch:
- Typos in character names
- Missing personalization names
- Wrong file loaded

### Tip 2: Use the Log
The processing log shows detailed information:
- Which images are being processed
- Any errors or warnings
- Success confirmations
- Look for ✓ marks for successful operations

### Tip 3: Check File Paths
If processing fails:
1. Check that FHM_Images folder is in parent directory
2. Verify format.pdf exists
3. Make sure font folder is present
4. Read error messages in the log

### Tip 4: Batch Processing
For large orders:
- Process in groups if needed
- Clear between batches
- Keep logs for reference

### Tip 5: Keep CSV Files
Save your CSV files for:
- Reprocessing orders
- Future reference
- Quick reprints

---

## 🔧 Troubleshooting

### Application Won't Start
**Problem**: Double-clicking launcher does nothing

**Solutions**:
- Make sure Python is installed
- Run `pip install -r requirements.txt`
- Try running: `python gui_app.py` directly

### Can't Drop Files
**Problem**: Drag & drop doesn't work

**Solution**:
- Use the "Browse..." button instead
- This is normal if tkinterdnd2 isn't installed
- Functionality is the same

### "No Module Named" Error
**Problem**: Missing Python packages

**Solution**:
```bash
pip install -r requirements.txt
```

### Processing Fails
**Problem**: Errors during processing

**Check**:
1. FHM_Images folder in parent directory?
2. Character names match image files exactly?
3. format.pdf present in current folder?
4. Read the error in the processing log

### Images Look Wrong
**Problem**: Output images don't look right

**Check**:
1. Are the correct source images being used?
2. Check the processing log for warnings
3. Verify character names in CSV
4. Look in outputs/ folder for the actual files

---

## ⌨️ Keyboard Shortcuts

Currently the GUI uses mouse/click interface. Future versions may include:
- Ctrl+O: Open file
- Ctrl+P: Process orders
- Ctrl+L: Clear all
- F1: Help

---

## 🎨 Interface Colors

The interface uses color coding for clarity:

- **Blue (#4a90e2)**: Primary actions, headers
- **Green (#5cb85c)**: Success, process button
- **Red (#d9534f)**: Errors, warnings
- **Gray (#6c757d)**: Secondary actions
- **Dark (#1e1e1e)**: Log background (console-like)

---

## 📸 What You Should See

### Normal Flow:
1. **Empty State**: Drop zone with instructions
2. **File Loaded**: Green checkmark, preview populated
3. **Processing**: Progress bar moving, logs updating
4. **Complete**: Success message, View Outputs enabled

### Status Messages:
- "Ready to process orders" → Waiting for file
- "Loaded X orders from file.csv" → File loaded
- "Processing orders..." → Working
- "✓ Orders processed successfully!" → Done!

---

## 🆘 Getting Help

### In-App Help
Click the "❓ Help" button for quick reference

### Documentation
- **README.md** - Main instructions
- **QUICKSTART.md** - Fast setup
- **PROCESS_ORDERS_README.md** - Detailed guide
- **GUI_GUIDE.md** - This file!

### Common Questions

**Q: Can I process multiple CSV files at once?**
A: Process one at a time, or combine them into a single CSV.

**Q: Can I preview the images before generating PDFs?**
A: Yes! Images are saved to outputs/ folder first. Check them before PDFs are made.

**Q: What if I have an odd number of orders?**
A: The last order will generate an image but no PDF (PDFs need 2 images).

**Q: Can I change the fonts or styles?**
A: Not through the GUI. You'd need to modify the process_orders.py code.

---

## 🎯 Example Workflow

**Sarah's typical workflow:**

1. Opens `launch_gui.bat` → App starts in 2 seconds
2. Drags `todays_orders.csv` onto drop zone
3. Sees "15 orders" in preview
4. Quickly scans the preview table
5. Clicks "▶ Process Orders"
6. Watches progress bar fill (30 seconds)
7. Sees "✓ Orders processed successfully!"
8. Clicks "📁 View Outputs" to check images
9. Sends PDFs to printer
10. Done in under 1 minute! 🎉

---

## 🔄 Version History

### v1.0 (Current)
- Initial GUI release
- Drag & drop support
- Real-time progress tracking
- Order preview
- Integrated logging
- Cross-platform support

---

**Enjoy the streamlined workflow! 🚀**

For more help, check the other documentation files or open the in-app help.

