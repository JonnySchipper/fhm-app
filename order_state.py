"""
Order State Management System
Tracks order completion status and manages order data
"""

import json
import os
import re
from datetime import datetime
from pathlib import Path

# State file location
STATE_FILE = "order_state.json"
ETSY_EMAILS_FILE = "etsy_sales_emails.txt"

def parse_orders_from_file(filepath=ETSY_EMAILS_FILE):
    """
    Parse orders from etsy_sales_emails.txt file.
    
    Returns list of order dictionaries:
    {
        'order_number': str,
        'name': str,
        'city': str,
        'state': str,
        'items': [{'item': str, 'personalization': str}, ...]
    }
    """
    if not os.path.exists(filepath):
        return []
    
    orders = []
    current_order = None
    
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Split by separator lines
    sections = content.split('------------------------------------------------------------')
    
    for section in sections:
        section = section.strip()
        if not section:
            continue
        
        lines = section.split('\n')
        order_data = {
            'order_number': '',
            'name': '',
            'city': '',
            'state': '',
            'items': []
        }
        
        current_item = None
        current_personalization = None
        
        for line in lines:
            line = line.strip()
            if not line:
                # If we have a pending item/personalization pair, add it
                if current_item:
                    order_data['items'].append({
                        'item': current_item,
                        'personalization': current_personalization or ''
                    })
                    current_item = None
                    current_personalization = None
                continue
            
            # Parse order fields
            if line.startswith('Order:'):
                order_data['order_number'] = line.replace('Order:', '').strip()
            elif line.startswith('Name:'):
                order_data['name'] = line.replace('Name:', '').strip()
                # Remove [Not found] placeholder
                if order_data['name'] == '[Not found]':
                    order_data['name'] = ''
            elif line.startswith('City:'):
                order_data['city'] = line.replace('City:', '').strip()
                if order_data['city'] == '[Not found]':
                    order_data['city'] = ''
            elif line.startswith('State:'):
                order_data['state'] = line.replace('State:', '').strip()
                if order_data['state'] == '[Not found]':
                    order_data['state'] = ''
            elif line.startswith('Transaction ID:'):
                # Skip Transaction ID lines - we just care about items/personalization
                continue
            elif line.startswith('Item:'):
                # Save previous item if exists
                if current_item:
                    order_data['items'].append({
                        'item': current_item,
                        'personalization': current_personalization or ''
                    })
                current_item = line.replace('Item:', '').strip()
                current_personalization = None
            elif line.startswith('Personalization:'):
                # Each personalization line creates a separate item (for image matching)
                new_personalization = line.replace('Personalization:', '').strip()
                
                if current_personalization:
                    # If we already have a personalization, save the previous item first
                    if current_item:
                        order_data['items'].append({
                            'item': current_item,
                            'personalization': current_personalization
                        })
                    # Start a new item with same item description
                    current_personalization = new_personalization
                else:
                    # First personalization for this item
                    current_personalization = new_personalization
        
        # Add last pending item
        if current_item:
            order_data['items'].append({
                'item': current_item,
                'personalization': current_personalization or ''
            })
        
        # IMPORTANT: Split comma-separated character-name pairs into separate items
        # This handles cases like "Character Stitch and Name Rhett, Character Star Lord and Name Dad"
        expanded_items = []
        for item in order_data['items']:
            pers = item['personalization']
            
            # Check if personalization has multiple character-name pairs (comma separated)
            # Pattern: "Character X ... Name Y, Character A ... Name B"
            if ', Character' in pers and 'Name' in pers:
                # Split by ", Character" and restore "Character" prefix
                parts = pers.split(', Character')
                for i, part in enumerate(parts):
                    if i == 0:
                        # First part already has "Character"
                        expanded_items.append({
                            'item': item['item'],
                            'personalization': part.strip()
                        })
                    else:
                        # Restore "Character" prefix with space
                        expanded_items.append({
                            'item': item['item'],
                            'personalization': 'Character ' + part.strip()
                        })
            else:
                # No comma-separated characters, keep as-is
                expanded_items.append(item)
        
        order_data['items'] = expanded_items
        
        # Only add order if it has an order number
        if order_data['order_number']:
            orders.append(order_data)
    
    return orders


def load_order_state():
    """
    Load order state from JSON file.
    
    Returns dict mapping order_number to completion status:
    {
        'order_number': {
            'completed': bool,
            'completed_date': str (ISO format),
            'items_processed': [...]
        }
    }
    """
    if not os.path.exists(STATE_FILE):
        return {}
    
    try:
        with open(STATE_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"Error loading order state: {e}")
        return {}


def save_order_state(state):
    """Save order state to JSON file."""
    try:
        with open(STATE_FILE, 'w', encoding='utf-8') as f:
            json.dump(state, f, indent=2)
        return True
    except Exception as e:
        print(f"Error saving order state: {e}")
        return False


def mark_orders_completed(order_numbers):
    """
    Mark specified orders as completed.
    
    Args:
        order_numbers: List of order numbers to mark as completed
    """
    state = load_order_state()
    
    for order_num in order_numbers:
        if order_num not in state:
            state[order_num] = {}
        
        state[order_num]['completed'] = True
        state[order_num]['completed_date'] = datetime.now().isoformat()
    
    save_order_state(state)


def get_order_status(order_number):
    """
    Get completion status for an order.
    
    Returns dict with completion info or None if not tracked.
    """
    state = load_order_state()
    return state.get(order_number)


def get_all_orders_with_status():
    """
    Get all orders from etsy_sales_emails.txt with their completion status.
    
    Returns list of order dicts with 'completed' field added.
    """
    orders = parse_orders_from_file()
    state = load_order_state()
    
    for order in orders:
        order_num = order['order_number']
        order_state = state.get(order_num, {})
        order['completed'] = order_state.get('completed', False)
        order['completed_date'] = order_state.get('completed_date', '')
    
    return orders


def get_pending_orders():
    """Get all orders that haven't been completed."""
    all_orders = get_all_orders_with_status()
    return [o for o in all_orders if not o.get('completed', False)]


def get_completed_orders():
    """Get all orders that have been completed."""
    all_orders = get_all_orders_with_status()
    return [o for o in all_orders if o.get('completed', False)]


def extract_order_text_for_ai(orders):
    """
    Extract Item and Personalization text from orders for AI processing.
    
    Args:
        orders: List of order dicts
        
    Returns:
        String formatted for AI input
    """
    lines = []
    
    for order in orders:
        order_num = order.get('order_number', 'Unknown')
        lines.append(f"\nOrder #{order_num}:")
        
        for item_data in order.get('items', []):
            item = item_data.get('item', '')
            personalization = item_data.get('personalization', '')
            
            if item:
                lines.append(f"Item: {item}")
            if personalization:
                lines.append(f"Personalization: {personalization}")
    
    return '\n'.join(lines)


if __name__ == "__main__":
    # Test the parser
    orders = parse_orders_from_file()
    print(f"Found {len(orders)} orders:")
    for order in orders:
        print(f"\nOrder {order['order_number']}")
        print(f"  Name: {order['name']}")
        print(f"  Location: {order['city']}, {order['state']}")
        print(f"  Items: {len(order['items'])}")
        for item in order['items']:
            print(f"    - {item['personalization']}")

