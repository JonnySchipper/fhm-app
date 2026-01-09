# Quick Start: Order Management System

## 🚀 Get Started in 3 Minutes

### Step 1: Setup Gmail Access (One-time)

Create `c:\dev\canva\gmailconfig.txt`:
```
GMAIL_USER=your.email@gmail.com
GMAIL_APP_PASSWORD=your-16-char-app-password
```

**Get App Password**: https://myaccount.google.com/apppasswords

---

### Step 2: Launch the GUI

Double-click `launch_gui.bat` or run:
```bash
python gui_app.py
```

---

### Step 3: Pull Your Orders

1. Look for the **📦 Order Management** section at the top
2. Click **"📥 Pull New Orders"**
3. Wait a few seconds
4. Orders appear in the list below!

---

### Step 4: Process Orders (2 Ways)

#### Option A: Automatic (Recommended) ⚡
1. Orders are **auto-selected** (pending orders checked)
2. Click **"▶ Begin Selected Orders"**
3. Click **"Yes"** to let AI parse automatically
4. Click **"✅ Confirm & Process Orders"** in preview window
5. ✅ Done! Orders marked as completed

#### Option B: Manual Review 📝
1. Select/deselect orders using checkboxes
2. Click **"▶ Begin Selected Orders"**
3. Click **"No"** to review manually
4. Edit the AI input field if needed
5. Click **"✨ Parse with AI"**
6. Review and process

---

### Step 5: Validate Before Shipping ✓

1. Click **"✓ Open Validation Page"**
2. Review each order:
   - ✅ Order number correct?
   - ✅ Address complete?
   - ✅ All images present?
3. Print and ship!

---

## 📊 Understanding the Order List

### Order Status:
- **⏳ Pending** (Yellow/Gray) = Not processed yet
- **✓ Completed** (Green) = Already processed

### Filter Options:
- **All**: Show everything
- **Pending**: Only unprocessed orders (default)
- **Completed**: Order history

### Each Order Shows:
```
[✓] Order #3942274249     ⏳ Pending     Miriam D Kestner • MELROSE, FL     (2 items)
    • Character Joy and name Miriam
    • Character Anger and name Chip
```

---

## 🎯 Typical Daily Workflow

**Morning:**
1. Launch GUI
2. Pull new orders
3. Review the list
4. Begin selected orders
5. Let AI parse
6. Click through preview → Process

**Before Shipping:**
1. Open Validation Page
2. Verify each order
3. Print PDFs
4. Pack and ship

**That's it!** ✨

---

## 💡 Pro Tips

### Efficiency Hacks:
- 🔄 **Pull orders multiple times** - It only fetches new ones
- ⌨️ **Use Select/Deselect All** - For bulk operations
- 🎨 **Filter by Pending** - Focus on what needs to be done
- ✓ **Always validate** - Catch errors before shipping
- 📦 **Check completed view** - See order history anytime

### Keyboard & Mouse:
- **Mouse wheel**: Scroll through long order lists
- **Checkboxes**: Quick select/deselect individual orders
- **Tab key**: Navigate between fields

### Best Practices:
1. ✅ Pull orders daily (or when Etsy emails arrive)
2. ✅ Process in batches (select multiple orders at once)
3. ✅ Let AI do the parsing (it's good at it!)
4. ✅ Always use validation page before printing
5. ✅ Keep the GUI open while working

---

## ❓ Quick Troubleshooting

### "No orders showing up"
→ Click **"Pull New Orders"** first
→ Check `etsy_sales_emails.txt` exists

### "Can't pull orders"
→ Check `gmailconfig.txt` is set up
→ Verify internet connection
→ Try generating a new App Password

### "AI parsing failed"
→ Check `grok_config.txt` has your API key
→ Try **"⚡ Quick Parse"** instead
→ Or edit manually in the input field

### "Orders not marked complete"
→ Use the proper workflow: Begin Orders → AI Parse → Process
→ Check that processing succeeded
→ Look for ✓ in the log

---

## 📁 Files You Should Know

### Created by System:
- `order_state.json` - Tracks completed orders (auto-managed)
- `etsy_sales_emails.txt` - Raw order data from email
- `processed_orders.txt` - Tracks pulled orders (prevents duplicates)

### You Manage:
- `c:\dev\canva\gmailconfig.txt` - Gmail credentials
- `grok_config.txt` - AI API key (optional but recommended)

---

## 🎓 Learn More

- **Full Guide**: Read `ORDER_MANAGEMENT_GUIDE.md`
- **Technical Details**: See `CHANGES_SUMMARY.md`
- **General GUI Help**: Check `GUI_GUIDE.md`
- **Setup Help**: Read `QUICKSTART.md`

---

## 🎉 You're Ready!

The order management system makes your workflow **10x faster**:
- ❌ **Before**: Manually copy/paste from emails, track on paper
- ✅ **After**: Click "Pull Orders", click "Begin Orders", done!

**Happy processing! 🎨✨**

---

### Need Help?
1. Check this guide
2. Review the full documentation
3. Look at terminal output for errors
4. Make sure all dependencies are installed

### Feedback?
Found a bug or want a feature? Document it and reach out!

