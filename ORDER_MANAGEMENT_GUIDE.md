# Order Management System Guide

## Overview

The Disney Magnet Order Processor now includes a comprehensive order management system that allows you to:
- Pull new orders directly from Gmail/Etsy
- Track which orders have been completed
- Automatically select pending orders for processing
- View order history (current and past orders)
- Validate processed orders with images before shipping

## New Features

### 1. **Order Management Section** 📦

Located at the top of the GUI, this section displays all your orders with their status.

#### Buttons:
- **📥 Pull New Orders**: Fetches new orders from your Gmail account (Etsy sales)
- **🔄 Refresh List**: Reloads the order list from saved files
- **Show Filter**: Toggle between All/Pending/Completed orders
- **✓ Select All / ✗ Deselect All**: Bulk selection controls
- **▶ Begin Selected Orders**: Process the selected orders
- **✓ Open Validation Page**: View completed orders with their images

### 2. **Order Display**

Each order shows:
- ✓ Completed or ⏳ Pending status
- Order number
- Customer name and location
- Number of items
- List of personalization details

Orders are displayed with checkboxes for easy selection.

### 3. **Automatic Workflow**

#### Step-by-Step Process:

1. **Pull New Orders**
   - Click "📥 Pull New Orders"
   - System fetches new Etsy sales from Gmail
   - Orders appear in the list automatically
   - Pending orders are auto-selected

2. **Select Orders to Process**
   - Pending orders are automatically selected
   - Click checkboxes to select/deselect specific orders
   - Use "Select All" or "Deselect All" for bulk actions

3. **Begin Processing**
   - Click "▶ Begin Selected Orders"
   - System extracts Item and Personalization info
   - Data is loaded into the AI parsing field
   - Choose to:
     - Auto-parse with AI (recommended)
     - Manually review/edit first

4. **AI Processing**
   - AI parses the order text
   - Matches characters to available images
   - Creates CSV format for processing

5. **Generate Images & PDFs**
   - Click "Process Orders" (or use the preview system)
   - System generates personalized images
   - Creates PDFs for printing
   - Marks orders as completed automatically

6. **Validate Before Shipping**
   - Click "✓ Open Validation Page"
   - Review each order with:
     - Order number
     - Shipping address
     - All images for that order
   - Verify everything is correct before shipping

### 4. **Order State Tracking**

The system maintains order status in `order_state.json`:
- Tracks which orders are completed
- Stores completion dates
- Persists across sessions

**Order Status:**
- **Pending**: New orders that haven't been processed
- **Completed**: Orders that have been processed and PDFs generated

### 5. **Validation Page** ✓

The validation page provides a final review before shipping:

- **Clean Organization**: Each order displayed separately
- **Complete Address**: Customer name, city, and state
- **Image Previews**: Actual thumbnails of generated images
- **Order Details**: All personalization information
- **Easy Verification**: Scroll through all orders at once

## Files Created

### New Files:
1. **`order_state.py`**: Order state management module
2. **`order_state.json`**: Tracks order completion status (auto-generated)
3. **`ORDER_MANAGEMENT_GUIDE.md`**: This guide

### Modified Files:
1. **`gui_app.py`**: Enhanced with order management UI
2. **`pullorders.py`**: Integrated for pulling new orders

## Configuration

### Gmail Setup (for pulling orders):
Create `gmailconfig.txt` in the parent directory (`c:\dev\canva\`) with:

```
GMAIL_USER=your.email@gmail.com
GMAIL_APP_PASSWORD=your-16-char-app-password
```

**Note**: You need a Gmail App Password (not your regular password). Generate one at:
https://myaccount.google.com/apppasswords

### How It Works:
1. Pulls Etsy sale emails from Gmail
2. Extracts order details (order #, name, address, items)
3. Saves to `etsy_sales_emails.txt`
4. Tracks processed orders in `processed_orders.txt`
5. Only pulls NEW orders (stops when it hits a previously processed one)

## Usage Tips

### Best Practices:

1. **Pull orders regularly**: 
   - Check for new orders daily or when you receive Etsy notifications
   - System only pulls new orders, so it's safe to run multiple times

2. **Review before processing**:
   - Check the order list before clicking "Begin Orders"
   - Verify customer names and personalizations
   - Deselect any orders you want to skip

3. **Use AI parsing**:
   - Let AI automatically parse orders (saves time)
   - It handles variations and typos
   - Review the parsed output before final processing

4. **Always validate**:
   - Use the validation page before printing/shipping
   - Verify each image matches the correct order
   - Check addresses are complete

5. **Track completion**:
   - Orders are automatically marked complete after processing
   - Use the "Completed" filter to review past orders
   - Completed orders won't be auto-selected again

### Keyboard Shortcuts & Tips:

- **Mouse wheel**: Scroll through long order lists
- **Filter views**: Switch between All/Pending/Completed to focus
- **Bulk operations**: Use Select/Deselect All for efficiency
- **Order status**: Green ✓ = Completed, Yellow ⏳ = Pending

## Troubleshooting

### Issue: No orders showing up
**Solution**: 
1. Click "Pull New Orders" first
2. Check that `etsy_sales_emails.txt` exists and has content
3. Try "Refresh List" button

### Issue: Can't pull new orders
**Solution**:
1. Verify `gmailconfig.txt` is set up correctly
2. Check internet connection
3. Verify Gmail App Password is valid
4. Check terminal output for error messages

### Issue: Orders not being marked complete
**Solution**:
1. Make sure you're using "Begin Selected Orders" → "Process Orders" workflow
2. Check that `order_state.json` file is writable
3. Verify processing completed successfully

### Issue: Validation page shows wrong images
**Solution**:
1. The current version shows all generated images
2. Make sure to process one batch of orders at a time
3. Check the `outputs/` folder to verify image generation

### Issue: AI parsing fails
**Solution**:
1. Verify Grok API key is set up in `grok_config.txt`
2. Check internet connection
3. Try "Quick Parse" instead of full reasoning mode
4. Manually edit the input if AI parsing doesn't work

## Advanced Features

### Manual Order Entry:
You can still manually enter orders in the AI parsing or order input sections:
1. Paste raw order text into "AI Order Parser"
2. Or type directly in "Add Orders" section
3. Use existing workflows as before

### CSV Import:
The "Load from CSV" button still works for bulk order imports.

### Archive System:
Old PDFs are automatically archived in `pdf_archive/` folder.

## Order State File Format

`order_state.json` structure:
```json
{
  "3942274249": {
    "completed": true,
    "completed_date": "2026-01-07T12:30:45"
  }
}
```

You can manually edit this file if needed (with the GUI closed).

## Future Enhancements

Potential additions:
- Image-to-order mapping in validation page
- Batch printing interface
- Order search and filtering
- Export completed orders to Excel
- Email notifications when new orders arrive
- Automatic daily order pulling (scheduled task)

## Support

For issues or questions:
1. Check this guide first
2. Review `GUI_GUIDE.md` for general GUI help
3. Check `QUICKSTART.md` for basic setup
4. Look at terminal output for error messages

---

**Enjoy streamlined order processing! 🎨✨**

