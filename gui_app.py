"""
Disney Magnet Order Processor - GUI Application
Interactive interface for processing orders with real-time feedback
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import tkinter.font as tkfont
import fitz

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    TkinterDnD = None
import os
import sys
import csv
import tempfile
import threading
import requests
import json
import shutil
import subprocess
import platform
import time
import re
from datetime import datetime
from pathlib import Path

# Import the processing functions
import process_orders

# Grok API configuration
GROK_API_URL = "https://api.x.ai/v1/chat/completions"

def open_file_or_folder(path):
    """Cross-platform way to open files or folders"""
    try:
        system = platform.system()
        if system == 'Darwin':  # macOS
            subprocess.run(['open', path], check=True)
        elif system == 'Windows':
            os.startfile(path)
        else:  # Linux
            subprocess.run(['xdg-open', path], check=True)
    except Exception as e:
        raise Exception(f"Failed to open {path}: {e}")

def load_api_key():
    """Load API key from config file"""
    try:
        # Try parent directory first (canva folder)
        parent_dir = os.path.dirname(os.path.abspath(os.getcwd()))
        parent_config = os.path.join(parent_dir, 'grok_config.txt')
        if os.path.exists(parent_config):
            with open(parent_config, 'r') as f:
                key = f.read().strip()
                if key:
                    return key
        
        # Fallback: check local config
        if os.path.exists("grok_config.txt"):
            with open("grok_config.txt", 'r') as f:
                key = f.read().strip()
                if key:
                    return key
    except Exception as e:
        print(f"Warning: Could not load API key: {e}")
    return None

GROK_API_KEY = load_api_key()


def extract_number(filename):
    """Extract the numeric suffix from order/boat PDF filenames for natural sorting.
    
    Examples:
        order_output_20260214_132805_1.pdf -> 1
        order_output_20260214_132805_10.pdf -> 10
        boat_output_20260214_132805_1.pdf -> 1
    """
    match = re.search(r'_(\d+)\.pdf$', filename)
    return int(match.group(1)) if match else 0


class OrderProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Disney Magnet Order Processor 🤖 AI-Powered")
        self.root.geometry("1000x750")
        self.root.minsize(900, 650)
        
        # Configure colors
        self.bg_color = "#f0f0f0"
        self.accent_color = "#4a90e2"
        self.success_color = "#5cb85c"
        self.error_color = "#d9534f"
        self.ai_color = "#9b59b6"
        
        self.root.configure(bg=self.bg_color)
        
        # Variables
        self.csv_path = tk.StringVar()
        self.status_text = tk.StringVar(value="Ready to process orders")
        self.orders_count = tk.StringVar(value="0 orders")
        self.available_images = tk.StringVar(value="Loading images...")
        self.processing = False
        self.ai_processing = False
        self.master_pdf_path = None
        self.zoom_level = 1.0  # Default zoom level
        
        # Get available images
        self.image_list = self.get_available_images()
        self.available_images.set(f"{len(self.image_list)} character images available")
        
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the user interface"""
        
        # Title Section
        title_frame = tk.Frame(self.root, bg=self.accent_color, height=90)
        title_frame.pack(fill=tk.X, padx=0, pady=0)
        title_frame.pack_propagate(False)
        
        title_label = tk.Label(
            title_frame,
            text="🎨 Disney Magnet Order Processor",
            font=("Segoe UI", 20, "bold"),
            bg=self.accent_color,
            fg="white"
        )
        title_label.pack(pady=(15, 5))
        
        subtitle_label = tk.Label(
            title_frame,
            text="✨ AI-Powered • Just Paste & Process",
            font=("Segoe UI", 10),
            bg=self.accent_color,
            fg="#e0e0e0"
        )
        subtitle_label.pack(pady=(0, 10))
        
        # Zoom controls (top right)
        zoom_frame = tk.Frame(title_frame, bg=self.accent_color)
        zoom_frame.place(relx=1.0, rely=0.5, anchor=tk.E, x=-10)
        
        zoom_label = tk.Label(
            zoom_frame,
            text="Zoom:",
            font=("Segoe UI", 9),
            bg=self.accent_color,
            fg="white"
        )
        zoom_label.pack(side=tk.LEFT, padx=(0, 5))
        
        zoom_out_btn = tk.Button(
            zoom_frame,
            text="−",
            command=self.zoom_out,
            font=("Segoe UI", 12, "bold"),
            bg="white",
            fg="black",
            relief=tk.FLAT,
            width=2,
            cursor="hand2"
        )
        zoom_out_btn.pack(side=tk.LEFT, padx=2)
        
        self.zoom_display = tk.Label(
            zoom_frame,
            text="100%",
            font=("Segoe UI", 9),
            bg=self.accent_color,
            fg="white",
            width=5
        )
        self.zoom_display.pack(side=tk.LEFT, padx=2)
        
        zoom_in_btn = tk.Button(
            zoom_frame,
            text="+",
            command=self.zoom_in,
            font=("Segoe UI", 12, "bold"),
            bg="white",
            fg="black",
            relief=tk.FLAT,
            width=2,
            cursor="hand2"
        )
        zoom_in_btn.pack(side=tk.LEFT, padx=2)
        
        # Main Content Frame
        content_frame = tk.Frame(self.root, bg=self.bg_color)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # AI Parsing Section
        self.create_ai_section(content_frame)
        
        # File Selection Section
        self.create_file_section(content_frame)
        
        # Order Preview Section
        self.create_preview_section(content_frame)
        
        # Progress Section
        self.create_progress_section(content_frame)
        
        # Control Buttons
        self.create_control_buttons(content_frame)
        
        # Status Bar
        self.create_status_bar()
        
    def create_ai_section(self, parent):
        """Create AI parsing section"""
        ai_frame = tk.LabelFrame(
            parent,
            text="🤖 AI Order Parser (Paste Raw Order Text)",
            font=("Segoe UI", 11, "bold"),
            bg=self.bg_color,
            fg=self.ai_color
        )
        ai_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Info row
        info_frame = tk.Frame(ai_frame, bg=self.bg_color)
        info_frame.pack(fill=tk.X, padx=10, pady=(5, 5))
        
        info_label = tk.Label(
            info_frame,
            text="Paste raw order text from emails/messages and let AI parse it automatically!",
            font=("Segoe UI", 9, "italic"),
            bg=self.bg_color,
            fg="#666"
        )
        info_label.pack(side=tk.LEFT)
        
        images_label = tk.Label(
            info_frame,
            textvariable=self.available_images,
            font=("Segoe UI", 8),
            bg=self.bg_color,
            fg=self.ai_color
        )
        images_label.pack(side=tk.RIGHT)
        
        # Raw text input
        raw_frame = tk.Frame(ai_frame, bg="white", relief=tk.RIDGE, bd=2)
        raw_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        self.raw_text = scrolledtext.ScrolledText(
            raw_frame,
            font=("Consolas", 9),
            height=6,
            bg="white",
            fg="#333",
            wrap=tk.WORD,
            relief=tk.FLAT
        )
        self.raw_text.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        # Placeholder
        placeholder = """Paste order details here, like:

Order #12345 - Mickey Captain themed, names: Johnny, Sarah, Michael
or
Disney Cruise Door Magnet - 3 magnets: Minnie, Donald, Goofy (all captain theme)"""
        self.raw_text.insert(1.0, placeholder)
        self.raw_text.config(fg="#999")
        
        self.raw_text.bind("<FocusIn>", self.clear_raw_placeholder)
        self.raw_text.bind("<FocusOut>", self.restore_raw_placeholder)
        
        # AI Button
        ai_btn_frame = tk.Frame(ai_frame, bg=self.bg_color)
        ai_btn_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        self.ai_parse_btn = tk.Button(
            ai_btn_frame,
            text="✨ Parse with AI (Grok)",
            command=self.parse_with_ai,
            font=("Segoe UI", 10, "bold"),
            bg=self.ai_color,
            fg="black",
            relief=tk.FLAT,
            padx=25,
            pady=8,
            cursor="hand2"
        )
        self.ai_parse_btn.pack(side=tk.LEFT)
        
        # Quick Parse button (non-reasoning)
        self.quick_parse_btn = tk.Button(
            ai_btn_frame,
            text="⚡ Quick Parse",
            command=self.quick_parse_with_ai,
            font=("Segoe UI", 9),
            bg="#17a2b8",
            fg="black",
            relief=tk.FLAT,
            padx=20,
            pady=8,
            cursor="hand2"
        )
        self.quick_parse_btn.pack(side=tk.LEFT, padx=(5, 0))
        
        # Clear raw button
        clear_raw_btn = tk.Button(
            ai_btn_frame,
            text="🗑️ Clear",
            command=self.clear_raw_text,
            font=("Segoe UI", 9),
            bg="#6c757d",
            fg="black",
            relief=tk.FLAT,
            padx=15,
            pady=8,
            cursor="hand2"
        )
        clear_raw_btn.pack(side=tk.LEFT, padx=(10, 0))
        
        # Sample button
        sample_raw_btn = tk.Button(
            ai_btn_frame,
            text="📋 Load Example",
            command=self.load_raw_sample,
            font=("Segoe UI", 9),
            bg="#f0ad4e",
            fg="black",
            relief=tk.FLAT,
            padx=15,
            pady=8,
            cursor="hand2"
        )
        sample_raw_btn.pack(side=tk.LEFT, padx=(5, 0))
        
    def create_file_section(self, parent):
        """Create order input section"""
        file_frame = tk.LabelFrame(
            parent,
            text="📝 Add Orders (Type or Paste Here)",
            font=("Segoe UI", 11, "bold"),
            bg=self.bg_color,
            fg="#333"
        )
        file_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Instructions
        instructions = tk.Label(
            file_frame,
            text="Type one order per line:  character-name,PersonName  (or paste from Excel/CSV)",
            font=("Segoe UI", 9, "italic"),
            bg=self.bg_color,
            fg="#666"
        )
        instructions.pack(pady=(5, 0), padx=10, anchor=tk.W)
        
        # Text input area
        input_frame = tk.Frame(file_frame, bg="white", relief=tk.RIDGE, bd=2)
        input_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.order_input = scrolledtext.ScrolledText(
            input_frame,
            font=("Consolas", 10),
            height=8,
            bg="white",
            fg="#333",
            wrap=tk.WORD,
            relief=tk.FLAT
        )
        self.order_input.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        # Placeholder text
        placeholder = "mickey-captain,Johnny\nminnie-captain,Sarah\nstitch-captain,Michael\nmoana-captain,Emma"
        self.order_input.insert(1.0, placeholder)
        self.order_input.config(fg="#999")
        
        # Bind events for placeholder
        self.order_input.bind("<FocusIn>", self.clear_placeholder)
        self.order_input.bind("<FocusOut>", self.restore_placeholder)
        self.order_input.bind("<KeyRelease>", self.update_count)
        
        # Button row
        btn_frame = tk.Frame(file_frame, bg=self.bg_color)
        btn_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        # Load from CSV button (optional)
        load_btn = tk.Button(
            btn_frame,
            text="📁 Load from CSV",
            command=self.browse_file,
            font=("Segoe UI", 9),
            bg="#6c757d",
            fg="black",
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor="hand2"
        )
        load_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # Clear button
        clear_input_btn = tk.Button(
            btn_frame,
            text="🗑️ Clear Input",
            command=self.clear_input,
            font=("Segoe UI", 9),
            bg="#6c757d",
            fg="black",
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor="hand2"
        )
        clear_input_btn.pack(side=tk.LEFT)
        
        # Add sample button
        sample_btn = tk.Button(
            btn_frame,
            text="📋 Paste Sample",
            command=self.load_sample,
            font=("Segoe UI", 9),
            bg="#f0ad4e",
            fg="black",
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor="hand2"
        )
        sample_btn.pack(side=tk.LEFT, padx=(5, 0))
        
    def create_preview_section(self, parent):
        """Create order preview section"""
        preview_frame = tk.LabelFrame(
            parent,
            text="📋 Order Preview",
            font=("Segoe UI", 11, "bold"),
            bg=self.bg_color,
            fg="#333"
        )
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Orders count and preview button
        top_frame = tk.Frame(preview_frame, bg=self.bg_color)
        top_frame.pack(fill=tk.X, padx=10, pady=(5, 0))
        
        count_label = tk.Label(
            top_frame,
            textvariable=self.orders_count,
            font=("Segoe UI", 9),
            bg=self.bg_color,
            fg="#666"
        )
        count_label.pack(side=tk.LEFT)
        
        preview_btn = tk.Button(
            top_frame,
            text="👁️ Preview Orders",
            command=self.preview_orders,
            font=("Segoe UI", 9),
            bg=self.accent_color,
            fg="black",
            relief=tk.FLAT,
            padx=15,
            pady=3,
            cursor="hand2"
        )
        preview_btn.pack(side=tk.RIGHT)
        
        # Scrolled text for preview
        self.preview_text = scrolledtext.ScrolledText(
            preview_frame,
            font=("Consolas", 9),
            height=10,
            bg="#f8f9fa",
            relief=tk.FLAT,
            wrap=tk.WORD
        )
        self.preview_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
    def create_progress_section(self, parent):
        """Create progress tracking section"""
        progress_frame = tk.LabelFrame(
            parent,
            text="⚙️ Processing",
            font=("Segoe UI", 11, "bold"),
            bg=self.bg_color,
            fg="#333"
        )
        progress_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            mode='determinate'
        )
        self.progress_bar.pack(fill=tk.X, padx=10, pady=10)
        
        # Log output
        self.log_text = scrolledtext.ScrolledText(
            progress_frame,
            font=("Consolas", 8),
            height=8,
            bg="#1e1e1e",
            fg="#d4d4d4",
            relief=tk.FLAT
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
    def create_control_buttons(self, parent):
        """Create control buttons"""
        btn_frame = tk.Frame(parent, bg=self.bg_color)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Process button
        self.process_btn = tk.Button(
            btn_frame,
            text="▶ Process Orders",
            command=self.process_orders,
            font=("Segoe UI", 11, "bold"),
            bg=self.success_color,
            fg="black",
            relief=tk.FLAT,
            padx=30,
            pady=10,
            cursor="hand2"
        )
        self.process_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Process button hover
        self.process_btn.bind("<Enter>", lambda e: self.process_btn.config(bg="#4cae4c"))
        self.process_btn.bind("<Leave>", lambda e: self.process_btn.config(bg=self.success_color))
        
        # View outputs button
        self.view_btn = tk.Button(
            btn_frame,
            text="📁 View Outputs",
            command=self.view_outputs,
            font=("Segoe UI", 9),
            bg="#6c757d",
            fg="black",
            relief=tk.FLAT,
            padx=20,
            pady=10,
            cursor="hand2",
            state=tk.DISABLED
        )
        self.view_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # View master PDF button
        self.master_pdf_btn = tk.Button(
            btn_frame,
            text="📄 Open Master PDF",
            command=self.open_master_pdf,
            font=("Segoe UI", 9, "bold"),
            bg="#17a2b8",
            fg="black",
            relief=tk.FLAT,
            padx=20,
            pady=10,
            cursor="hand2",
            state=tk.DISABLED
        )
        self.master_pdf_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Add hover effect
        def on_master_enter(e):
            if self.master_pdf_btn['state'] == tk.NORMAL:
                self.master_pdf_btn.config(bg="#138496")
        def on_master_leave(e):
            if self.master_pdf_btn['state'] == tk.NORMAL:
                self.master_pdf_btn.config(bg="#17a2b8")
        
        self.master_pdf_btn.bind("<Enter>", on_master_enter)
        self.master_pdf_btn.bind("<Leave>", on_master_leave)
        
        # Clear button
        clear_btn = tk.Button(
            btn_frame,
            text="🗑️ Clear",
            command=self.clear_all,
            font=("Segoe UI", 9),
            bg="#6c757d",
            fg="black",
            relief=tk.FLAT,
            padx=20,
            pady=10,
            cursor="hand2"
        )
        clear_btn.pack(side=tk.LEFT)
        
        # Help button (right side)
        help_btn = tk.Button(
            btn_frame,
            text="❓ Help",
            command=self.show_help,
            font=("Segoe UI", 9),
            bg="#f0ad4e",
            fg="black",
            relief=tk.FLAT,
            padx=20,
            pady=10,
            cursor="hand2"
        )
        help_btn.pack(side=tk.RIGHT)
        
        # Archive button
        archive_btn = tk.Button(
            btn_frame,
            text="📦 View Archive",
            command=self.view_archive,
            font=("Segoe UI", 9),
            bg="#6c757d",
            fg="black",
            relief=tk.FLAT,
            padx=15,
            pady=10,
            cursor="hand2"
        )
        archive_btn.pack(side=tk.RIGHT, padx=(0, 10))
        
    def create_status_bar(self):
        """Create status bar at bottom"""
        status_frame = tk.Frame(self.root, bg="#2c3e50", height=30)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        status_frame.pack_propagate(False)
        
        status_label = tk.Label(
            status_frame,
            textvariable=self.status_text,
            font=("Segoe UI", 9),
            bg="#2c3e50",
            fg="white",
            anchor=tk.W
        )
        status_label.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        
    def clear_placeholder(self, event):
        """Clear placeholder text on focus"""
        if self.order_input.get(1.0, tk.END).strip() in ["mickey-captain,Johnny\nminnie-captain,Sarah\nstitch-captain,Michael\nmoana-captain,Emma", ""]:
            self.order_input.delete(1.0, tk.END)
            self.order_input.config(fg="#333")
            
    def restore_placeholder(self, event):
        """Restore placeholder if empty"""
        if not self.order_input.get(1.0, tk.END).strip():
            placeholder = "mickey-captain,Johnny\nminnie-captain,Sarah\nstitch-captain,Michael\nmoana-captain,Emma"
            self.order_input.insert(1.0, placeholder)
            self.order_input.config(fg="#999")
            
    def update_count(self, event=None):
        """Update order count as user types"""
        text = self.order_input.get(1.0, tk.END).strip()
        if text and self.order_input.cget("fg") == "#333":  # Not placeholder
            lines = [l.strip() for l in text.split('\n') if l.strip() and ',' in l]
            self.orders_count.set(f"{len(lines)} orders")
        else:
            self.orders_count.set("0 orders")
            
    def clear_input(self):
        """Clear the input area"""
        self.order_input.delete(1.0, tk.END)
        self.order_input.config(fg="#333")
        self.preview_text.delete(1.0, tk.END)
        self.orders_count.set("0 orders")
        self.status_text.set("Ready to process orders")
        
    def load_sample(self):
        """Load sample orders"""
        sample = """mickey-captain,Johnny
minnie-captain,Sarah
donald-captain,Michael
daisy-captain,Emma
goofy-captain,Oliver
pluto-captain,Sophia"""
        self.order_input.delete(1.0, tk.END)
        self.order_input.insert(1.0, sample)
        self.order_input.config(fg="#333")
        self.update_count()
        self.status_text.set("Sample orders loaded")
        self.preview_orders()
                
    def browse_file(self):
        """Browse for CSV file and load into input"""
        filename = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if filename:
            self.load_csv(filename)
            
    def load_csv(self, filepath):
        """Load CSV file into input area"""
        try:
            # Read CSV
            orders = []
            with open(filepath, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                header = next(reader, None)
                
                for row in reader:
                    if row and row[0].strip():
                        character = row[0].strip()
                        name = row[1].strip() if len(row) > 1 else ""
                        orders.append(f"{character},{name}")
            
            # Load into input area
            self.order_input.delete(1.0, tk.END)
            self.order_input.insert(1.0, '\n'.join(orders))
            self.order_input.config(fg="#333")
            
            self.update_count()
            self.status_text.set(f"Loaded {len(orders)} orders from {os.path.basename(filepath)}")
            self.preview_orders()
            
        except Exception as e:
            error_msg = str(e)
            messagebox.showerror("Error", f"Failed to load CSV:\n{error_msg}")
            self.log(f"ERROR: {error_msg}", "error")
            
    def preview_orders(self):
        """Preview the orders before processing"""
        text = self.order_input.get(1.0, tk.END).strip()
        
        if not text or self.order_input.cget("fg") == "#999":
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(tk.END, "No orders to preview. Please add orders above.")
            return
        
        # Parse orders
        orders = []
        lines = text.split('\n')
        for line in lines:
            line = line.strip()
            if not line:
                continue
            if ',' in line:
                parts = line.split(',', 1)
                character = parts[0].strip()
                name = parts[1].strip() if len(parts) > 1 else ""
                orders.append((character, name))  # Keep empty string for no personalization
        
        # Update preview
        self.preview_text.delete(1.0, tk.END)
        self.preview_text.insert(tk.END, f"{'Character':<30} {'Name':<25}\n")
        self.preview_text.insert(tk.END, "-" * 55 + "\n")
        
        for character, name in orders:
            display_name = name if name else "(no name)"
            self.preview_text.insert(tk.END, f"{character:<30} {display_name:<25}\n")
        
        self.orders_count.set(f"{len(orders)} orders")
        self.status_text.set(f"Preview ready: {len(orders)} orders")
        
        # Show image preview window
        self.show_image_preview(orders)
    
    def show_image_preview(self, orders):
        """Show an interactive window to edit orders with image previews"""
        try:
            from PIL import Image, ImageTk
        except ImportError:
            messagebox.showinfo("Preview Unavailable", "Image preview requires the Pillow library.\nInstall it with: pip install Pillow")
            return
        
        # Get the images directory
        parent_dir = os.path.dirname(os.path.abspath(os.getcwd()))
        images_dir = os.path.join(parent_dir, 'fhm_images')
        
        if not os.path.exists(images_dir):
            # Try current directory's fhm_images
            images_dir = os.path.join(os.getcwd(), 'fhm_images')
            if not os.path.exists(images_dir):
                messagebox.showwarning("Images Not Found", "Cannot find fhm_images folder.")
                return
        
        # Get boats directory
        boats_dir = "boats"
        
        # Get all available images (characters + boats)
        available_images = []
        for file in sorted(os.listdir(images_dir)):
            if file.lower().endswith('.png'):
                available_images.append(file.replace('.png', ''))
        
        # Add boat images if boats folder exists
        if os.path.exists(boats_dir):
            for file in sorted(os.listdir(boats_dir)):
                if file.lower().endswith('.png'):
                    available_images.append(file.replace('.png', ''))
        
        # Create preview window
        preview_window = tk.Toplevel(self.root)
        preview_window.title("✏️ Edit & Confirm Orders")
        preview_window.geometry("950x700")
        preview_window.configure(bg="white")
        
        # Make it modal
        preview_window.transient(self.root)
        preview_window.grab_set()
        
        # Title
        title_frame = tk.Frame(preview_window, bg=self.accent_color, height=70)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)
        
        title_label = tk.Label(
            title_frame,
            text="✏️ Edit & Confirm Orders",
            font=("Segoe UI", 16, "bold"),
            bg=self.accent_color,
            fg="white"
        )
        title_label.pack(pady=(10, 0))
        
        subtitle_label = tk.Label(
            title_frame,
            text="Select images, edit names, then click 'Confirm & Process'",
            font=("Segoe UI", 9),
            bg=self.accent_color,
            fg="#e0e0e0"
        )
        subtitle_label.pack(pady=(0, 10))
        
        # Create scrollable frame
        canvas = tk.Canvas(preview_window, bg="white")
        scrollbar = tk.Scrollbar(preview_window, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="white")
        
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Update scroll region when content changes
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        # Make scrollable frame expand to canvas width
        def configure_canvas_width(event):
            canvas.itemconfig(canvas_window, width=event.width)
        
        canvas.bind('<Configure>', configure_canvas_width)
        
        # Enable mouse wheel scrolling (cross-platform)
        def on_mousewheel(event):
            # Windows and Linux
            if event.delta:
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def on_mac_mousewheel(event):
            # Mac uses different delta values
            canvas.yview_scroll(-1 * event.delta, "units")
        
        def on_button_4(event):
            # X11 scroll up (Linux, some Mac configs)
            canvas.yview_scroll(-1, "units")
        
        def on_button_5(event):
            # X11 scroll down (Linux, some Mac configs)
            canvas.yview_scroll(1, "units")
        
        # Bind mouse wheel events based on platform
        system = platform.system()
        if system == 'Darwin':  # macOS
            canvas.bind_all("<MouseWheel>", on_mac_mousewheel)
            canvas.bind_all("<Button-4>", on_button_4)
            canvas.bind_all("<Button-5>", on_button_5)
        elif system == 'Linux':
            canvas.bind_all("<Button-4>", on_button_4)
            canvas.bind_all("<Button-5>", on_button_5)
        else:  # Windows
            canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        # Store order data (will be updated as user edits)
        order_data = []
        order_widgets = []  # Store references to widgets for updates
        
        # Function to create an order row
        def create_order_row(character="", name="", year=""):
            """Create a single editable order row"""
            idx = len(order_data)
            order_data.append({'character': character, 'name': name, 'year': year})
            
            # Create frame for this order
            order_frame = tk.Frame(scrollable_frame, bg="white", relief=tk.RIDGE, bd=2)
            order_frame.pack(fill=tk.X, padx=10, pady=8)
            
            # Image preview (left side) - fixed size container
            image_container = tk.Frame(order_frame, bg="white", width=100, height=100)
            image_container.pack(side=tk.LEFT, padx=10, pady=10)
            image_container.pack_propagate(False)  # Prevent resizing
            
            # Load initial image - check both fhm_images and boats folders
            image_filename = f"{character}.png"
            image_path = os.path.join(images_dir, image_filename)
            
            # If not found in images_dir, check boats folder
            if not os.path.exists(image_path) and os.path.exists(boats_dir):
                boat_image_path = os.path.join(boats_dir, image_filename)
                if os.path.exists(boat_image_path):
                    image_path = boat_image_path
            
            photo = None
            if os.path.exists(image_path):
                try:
                    img = Image.open(image_path)
                    img.thumbnail((100, 100), Image.Resampling.LANCZOS)
                    photo = ImageTk.PhotoImage(img)
                except:
                    photo = None
            
            # Image label - fills the container
            img_label = tk.Label(image_container, bg="#f0f0f0", relief=tk.SUNKEN, bd=1)
            if photo:
                img_label.config(image=photo)
                if not hasattr(self, '_preview_images'):
                    self._preview_images = []
                self._preview_images.append(photo)
            else:
                img_label.config(text="No\nImage", font=("Segoe UI", 10), fg="#999")
            img_label.pack(fill=tk.BOTH, expand=True)
            
            # Edit controls (right side)
            edit_frame = tk.Frame(order_frame, bg="white")
            edit_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=10, padx=10)
            
            # Row 1: Character selection
            char_row = tk.Frame(edit_frame, bg="white")
            char_row.pack(fill=tk.X, pady=(0, 5))
            
            tk.Label(
                char_row,
                text="Character:",
                font=("Segoe UI", 9, "bold"),
                bg="white",
                width=10,
                anchor=tk.W
            ).pack(side=tk.LEFT)
            
            # Character text display (no dropdown)
            char_var = tk.StringVar(value=character)
            
            char_label = tk.Label(
                char_row,
                textvariable=char_var,
                font=("Consolas", 10),
                bg="#f8f9fa",
                fg="black",
                anchor=tk.W,
                relief=tk.SUNKEN,
                bd=1,
                padx=5,
                pady=3
            )
            char_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
            
            # Row 2: Name input
            name_row = tk.Frame(edit_frame, bg="white")
            name_row.pack(fill=tk.X, pady=(0, 5))
            
            tk.Label(
                name_row,
                text="Name:",
                font=("Segoe UI", 9, "bold"),
                bg="white",
                width=10,
                anchor=tk.W
            ).pack(side=tk.LEFT)
            
            name_var = tk.StringVar(value=name)
            name_entry = tk.Entry(
                name_row,
                textvariable=name_var,
                font=("Segoe UI", 10),
                bg="#f8f9fa",
                fg="black"
            )
            name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            # Row 2b: Year input (only for boats)
            year_var = tk.StringVar(value=year)
            year_row = tk.Frame(edit_frame, bg="white")
            
            # Check if this is a boat order
            is_boat = character.lower().startswith('boat_')
            
            if is_boat:
                year_row.pack(fill=tk.X, pady=(0, 5))
                
                tk.Label(
                    year_row,
                    text="Year:",
                    font=("Segoe UI", 9, "bold"),
                    bg="white",
                    width=10,
                    anchor=tk.W
                ).pack(side=tk.LEFT)
                
                year_entry = tk.Entry(
                    year_row,
                    textvariable=year_var,
                    font=("Segoe UI", 10),
                    bg="#fff8e7",  # Light gold background to indicate boat-specific
                    fg="black",
                    width=10
                )
                year_entry.pack(side=tk.LEFT)
                
                tk.Label(
                    year_row,
                    text="(optional - for boat orders only)",
                    font=("Segoe UI", 8, "italic"),
                    bg="white",
                    fg="#999"
                ).pack(side=tk.LEFT, padx=(10, 0))
            
            # Row 3: Status and actions
            action_row = tk.Frame(edit_frame, bg="white")
            action_row.pack(fill=tk.X, pady=(5, 0))
            
            status_label = tk.Label(
                action_row,
                text="✓ Ready" if os.path.exists(image_path) else "⚠ Not found",
                font=("Segoe UI", 8),
                bg="white",
                fg="#5cb85c" if os.path.exists(image_path) else "#f0ad4e"
            )
            status_label.pack(side=tk.LEFT)
            
            # Delete button
            def make_delete_handler(idx):
                return lambda: delete_order(idx)
            
            delete_btn = tk.Button(
                action_row,
                text="🗑️ Delete",
                command=make_delete_handler(idx),
                font=("Segoe UI", 8),
                bg="#d9534f",
                fg="white",
                relief=tk.FLAT,
                padx=8,
                pady=2,
                cursor="hand2"
            )
            delete_btn.pack(side=tk.RIGHT, padx=(5, 0))
            
            # Function to update image when character changes
            def make_update_handler(idx, img_lbl, char_v, name_v, year_v, status_lbl):
                def update_image(*args):
                    new_char = char_v.get()
                    order_data[idx]['character'] = new_char
                    # Note: name and year are updated by their own trace handlers
                    # so we don't overwrite them here to avoid sync issues
                    
                    # Update image - check both fhm_images and boats folders
                    new_image_path = os.path.join(images_dir, f"{new_char}.png")
                    
                    # If not found in images_dir, check boats folder
                    if not os.path.exists(new_image_path) and os.path.exists(boats_dir):
                        boat_path = os.path.join(boats_dir, f"{new_char}.png")
                        if os.path.exists(boat_path):
                            new_image_path = boat_path
                    
                    if os.path.exists(new_image_path):
                        try:
                            img = Image.open(new_image_path)
                            img.thumbnail((100, 100), Image.Resampling.LANCZOS)
                            new_photo = ImageTk.PhotoImage(img)
                            img_lbl.config(image=new_photo, text="", bg="#f0f0f0")
                            self._preview_images.append(new_photo)
                            status_lbl.config(text="✓ Ready", fg="#5cb85c")
                        except:
                            img_lbl.config(image="", text="Error\nLoading", bg="#fff0f0")
                            status_lbl.config(text="❌ Error", fg="#d9534f")
                    else:
                        img_lbl.config(image="", text="No\nImage", bg="#f0f0f0")
                        status_lbl.config(text="⚠ Not found", fg="#f0ad4e")
                return update_image
            
            update_handler = make_update_handler(idx, img_label, char_var, name_var, year_var, status_label)
            # Update data when fields change
            char_var.trace_add('write', lambda *args, h=update_handler: h())
            name_var.trace_add('write', lambda *args, i=idx, n=name_var: order_data[i].update({'name': n.get()}))
            year_var.trace_add('write', lambda *args, i=idx, y=year_var: order_data[i].update({'year': y.get()}))
            
            # Search button (after update_handler exists so we can pass it as callback)
            def make_search_handler(char_v, update_h):
                def open_search():
                    # Process any pending UI updates (e.g. focus loss from name entry)
                    # This ensures the name value is synced before opening the modal dialog
                    preview_window.update_idletasks()
                    open_image_search(char_v, on_select_callback=update_h)
                return open_search
            
            search_btn = tk.Button(
                char_row,
                text="🔍 Search",
                command=make_search_handler(char_var, update_handler),
                font=("Segoe UI", 9),
                bg="#17a2b8",
                fg="white",
                relief=tk.FLAT,
                padx=12,
                pady=5,
                cursor="hand2"
            )
            search_btn.pack(side=tk.LEFT)
            
            order_widgets.append({
                'frame': order_frame,
                'char_var': char_var,
                'name_var': name_var,
                'year_var': year_var,
                'img_label': img_label,
                'status_label': status_label
            })
            
            update_summary()
            return idx
        
        # Image search dialog
        def open_image_search(target_var, on_select_callback=None):
            """Open a fast searchable list to select character images"""
            search_window = tk.Toplevel(preview_window)
            search_window.title("🔍 Search Characters")
            search_window.geometry("500x600")
            search_window.configure(bg="white")
            search_window.transient(preview_window)
            search_window.grab_set()  # Make modal
            
            # Title
            tk.Label(
                search_window,
                text="🔍 Search Characters",
                font=("Segoe UI", 14, "bold"),
                bg="white"
            ).pack(pady=10)
            
            # Search input
            search_frame = tk.Frame(search_window, bg="white")
            search_frame.pack(fill=tk.X, padx=15, pady=(0, 10))
            
            tk.Label(
                search_frame,
                text="Type to filter:",
                font=("Segoe UI", 11),
                bg="white"
            ).pack(anchor=tk.W, pady=(0, 5))
            
            search_var = tk.StringVar()
            search_entry = tk.Entry(
                search_frame,
                textvariable=search_var,
                font=("Segoe UI", 12),
                bg="white",
                fg="black",
                insertbackground="black"
            )
            search_entry.pack(fill=tk.X)
            search_entry.focus()
            
            # Results count
            count_label = tk.Label(
                search_window,
                text=f"{len(available_images)} characters available",
                font=("Segoe UI", 9),
                bg="white",
                fg="#666"
            )
            count_label.pack(pady=(0, 5))
            
            # Listbox with scrollbar
            list_frame = tk.Frame(search_window, bg="white")
            list_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 10))
            
            scrollbar = tk.Scrollbar(list_frame)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            listbox = tk.Listbox(
                list_frame,
                font=("Consolas", 11),
                yscrollcommand=scrollbar.set,
                selectmode=tk.SINGLE,
                activestyle='dotbox',
                bg="white",
                fg="black",
                selectbackground=self.accent_color,
                selectforeground="white",
                relief=tk.SOLID,
                bd=1
            )
            listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.config(command=listbox.yview)
            
            # Populate initial list
            def update_list(query=""):
                listbox.delete(0, tk.END)
                
                if query:
                    filtered = [img for img in available_images if query.lower() in img.lower()]
                else:
                    filtered = available_images
                
                for item in filtered:
                    listbox.insert(tk.END, item)
                
                count_label.config(text=f"{len(filtered)} matches")
                
                # Auto-select first item if any
                if filtered:
                    listbox.selection_set(0)
                    listbox.see(0)
            
            # Initial population
            update_list()
            
            # Bind search
            def on_search_change(*args):
                update_list(search_var.get())
            
            search_var.trace_add('write', on_search_change)
            
            # Select on double-click
            def on_double_click(event):
                selection = listbox.curselection()
                if selection:
                    selected = listbox.get(selection[0])
                    target_var.set(selected)
                    if on_select_callback:
                        on_select_callback()
                    search_window.destroy()
            
            listbox.bind('<Double-Button-1>', on_double_click)
            
            # Select on Enter key
            def on_enter(event):
                selection = listbox.curselection()
                if selection:
                    selected = listbox.get(selection[0])
                    target_var.set(selected)
                    if on_select_callback:
                        on_select_callback()
                    search_window.destroy()
            
            listbox.bind('<Return>', on_enter)
            search_entry.bind('<Return>', on_enter)
            
            # Enable mouse wheel scrolling for listbox (cross-platform)
            def on_listbox_mousewheel(event):
                if event.delta:
                    listbox.yview_scroll(int(-1*(event.delta/120)), "units")
            
            def on_listbox_mac_mousewheel(event):
                listbox.yview_scroll(-1 * event.delta, "units")
            
            def on_listbox_button_4(event):
                listbox.yview_scroll(-1, "units")
            
            def on_listbox_button_5(event):
                listbox.yview_scroll(1, "units")
            
            # Bind based on platform
            system = platform.system()
            if system == 'Darwin':  # macOS
                listbox.bind("<MouseWheel>", on_listbox_mac_mousewheel)
                listbox.bind("<Button-4>", on_listbox_button_4)
                listbox.bind("<Button-5>", on_listbox_button_5)
            elif system == 'Linux':
                listbox.bind("<Button-4>", on_listbox_button_4)
                listbox.bind("<Button-5>", on_listbox_button_5)
            else:  # Windows
                listbox.bind("<MouseWheel>", on_listbox_mousewheel)
            
            # Button frame
            btn_frame = tk.Frame(search_window, bg="white")
            btn_frame.pack(fill=tk.X, padx=15, pady=(0, 15))
            
            # Select button
            def on_select():
                selection = listbox.curselection()
                if selection:
                    selected = listbox.get(selection[0])
                    target_var.set(selected)
                    if on_select_callback:
                        on_select_callback()
                    search_window.destroy()
                else:
                    messagebox.showwarning("No Selection", "Please select a character from the list.")
            
            select_btn = tk.Button(
                btn_frame,
                text="✓ Select Character",
                command=on_select,
                font=("Segoe UI", 11, "bold"),
                bg=self.success_color,
                fg="white",
                relief=tk.FLAT,
                cursor="hand2",
                padx=20,
                pady=8
            )
            select_btn.pack(side=tk.LEFT, padx=(0, 10))
            
            # Cancel button
            cancel_btn = tk.Button(
                btn_frame,
                text="✕ Cancel",
                command=search_window.destroy,
                font=("Segoe UI", 10),
                bg="#6c757d",
                fg="white",
                relief=tk.FLAT,
                cursor="hand2",
                padx=20,
                pady=8
            )
            cancel_btn.pack(side=tk.LEFT)
            
            # Instructions
            tk.Label(
                search_window,
                text="💡 Tip: Type to filter • Double-click or press Enter to select",
                font=("Segoe UI", 8, "italic"),
                bg="white",
                fg="#999"
            ).pack(pady=(0, 10))
        
        # Delete order function
        def delete_order(idx):
            if messagebox.askyesno("Delete Order", f"Delete order #{idx + 1}?"):
                order_widgets[idx]['frame'].destroy()
                order_data[idx] = None  # Mark as deleted
                update_summary()
        
        # Function to update summary
        def update_summary():
            active_orders = [o for o in order_data if o is not None]
            
            # Separate boats from character orders
            boats = [o for o in active_orders if o['character'].lower().startswith('boat_')]
            portholes = [o for o in active_orders if not o['character'].lower().startswith('boat_')]
            
            # Check both fhm_images and boats folders for image existence
            def image_exists(char):
                if os.path.exists(os.path.join(images_dir, f"{char}.png")):
                    return True
                if os.path.exists(boats_dir) and os.path.exists(os.path.join(boats_dir, f"{char}.png")):
                    return True
                return False
            
            found = sum(1 for o in active_orders if image_exists(o['character']))
            missing = len(active_orders) - found
            
            # Build summary with warning for odd portholes
            summary_text = f"Boats: {len(boats)}  |  Portholes: {len(portholes)}  |  ✓ Ready: {found}  |  ⚠ Issues: {missing}"
            
            # Change color if portholes are odd
            if len(portholes) % 2 == 1:
                summary_label.config(text=summary_text, fg="#d9534f")  # Red warning
            else:
                summary_label.config(text=summary_text, fg="#333")  # Normal
        
        # Add Order button (above canvas)
        add_order_frame = tk.Frame(preview_window, bg="white")
        add_order_frame.pack(fill=tk.X, padx=10, pady=(10, 5))
        
        def add_new_order():
            create_order_row("", "")
            # Scroll to bottom
            canvas.update_idletasks()
            canvas.yview_moveto(1.0)
        
        add_btn = tk.Button(
            add_order_frame,
            text="➕ Add New Order",
            command=add_new_order,
            font=("Segoe UI", 10, "bold"),
            bg="#28a745",
            fg="black",
            relief=tk.FLAT,
            padx=20,
            pady=8,
            cursor="hand2"
        )
        add_btn.pack()
        
        # Bottom control panel (pack BEFORE canvas so it reserves space at bottom)
        bottom_frame = tk.Frame(preview_window, bg="#f0f0f0", relief=tk.RAISED, bd=2)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Summary
        summary_label = tk.Label(
            bottom_frame,
            text="Loading orders...",
            font=("Segoe UI", 10, "bold"),
            bg="#f0f0f0",
            fg="#333"
        )
        summary_label.pack(pady=10)
        
        # Pack scrollbar and canvas (pack AFTER bottom frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Initialize existing orders (AFTER canvas is packed)
        for character, name in orders:
            create_order_row(character, name)
        
        # Force canvas to update its scroll region
        canvas.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))
        
        update_summary()
        
        # Buttons
        button_frame = tk.Frame(bottom_frame, bg="#f0f0f0")
        button_frame.pack(pady=(0, 15))
        
        # Confirm and Process button
        def confirm_and_process():
            # Collect active orders
            active_orders = [o for o in order_data if o is not None]
            
            if not active_orders:
                messagebox.showwarning("No Orders", "No orders to process!")
                return
            
            # Separate boats from character orders
            portholes = [o for o in active_orders if not o['character'].lower().startswith('boat_')]
            
            # Check for odd number of character orders
            if len(portholes) % 2 == 1:
                messagebox.showwarning(
                    "Odd Number of Portholes",
                    f"You have {len(portholes)} character order(s) (portholes).\n\n"
                    "Character magnets are printed 2 per page.\n"
                    "With an odd number, the last order will be skipped!\n\n"
                    "Please add or remove one order to make it even.",
                    icon='warning'
                )
                return
            
            # Check for missing images (check both fhm_images and boats folders)
            missing = []
            for o in active_orders:
                char_image_path = os.path.join(images_dir, f"{o['character']}.png")
                boat_image_path = os.path.join(boats_dir, f"{o['character']}.png") if os.path.exists(boats_dir) else None
                if not os.path.exists(char_image_path) and (boat_image_path is None or not os.path.exists(boat_image_path)):
                    display_name = o['name'] if o['name'] else "(no name)"
                    missing.append(f"{o['character']} (for {display_name})")
            
            if missing:
                response = messagebox.askyesno(
                    "Missing Images",
                    f"{len(missing)} order(s) have missing images:\n\n" +
                    "\n".join(missing[:5]) +
                    ("\n..." if len(missing) > 5 else "") +
                    "\n\nContinue anyway?",
                    icon='warning'
                )
                if not response:
                    return
            
            # Update main input with edited orders (include year for boats if present)
            def format_order(o):
                year = o.get('year', '')
                if year:
                    return f"{o['character']},{o['name']},{year}"
                return f"{o['character']},{o['name']}"
            
            orders_text = '\n'.join([format_order(o) for o in active_orders])
            self.order_input.delete(1.0, tk.END)
            self.order_input.insert(1.0, orders_text)
            self.order_input.config(fg="#333")
            self.update_count()
            
            # Close preview and start processing
            preview_window.destroy()
            self.process_orders()
        
        confirm_btn = tk.Button(
            button_frame,
            text="✅ Confirm & Process Orders",
            command=confirm_and_process,
            font=("Segoe UI", 11, "bold"),
            bg=self.success_color,
            fg="black",
            relief=tk.FLAT,
            padx=30,
            pady=10,
            cursor="hand2"
        )
        confirm_btn.pack(side=tk.LEFT, padx=5)
        
        # Update only button
        def update_only():
            active_orders = [o for o in order_data if o is not None]
            if not active_orders:
                messagebox.showwarning("No Orders", "No orders to update!")
                return
            
            # Include year for boats if present
            def format_order(o):
                year = o.get('year', '')
                if year:
                    return f"{o['character']},{o['name']},{year}"
                return f"{o['character']},{o['name']}"
            
            orders_text = '\n'.join([format_order(o) for o in active_orders])
            self.order_input.delete(1.0, tk.END)
            self.order_input.insert(1.0, orders_text)
            self.order_input.config(fg="#333")
            self.update_count()
            
            messagebox.showinfo("Updated", "Orders updated in the main window!")
            preview_window.destroy()
        
        update_btn = tk.Button(
            button_frame,
            text="💾 Save Changes",
            command=update_only,
            font=("Segoe UI", 10),
            bg="#17a2b8",
            fg="black",
            relief=tk.FLAT,
            padx=20,
            pady=10,
            cursor="hand2"
        )
        update_btn.pack(side=tk.LEFT, padx=5)
        
        # Cancel button
        def cancel():
            if messagebox.askyesno("Cancel", "Discard all changes?"):
                preview_window.destroy()
        
        cancel_btn = tk.Button(
            button_frame,
            text="❌ Cancel",
            command=cancel,
            font=("Segoe UI", 10),
            bg="#6c757d",
            fg="black",
            relief=tk.FLAT,
            padx=20,
            pady=10,
            cursor="hand2"
        )
        cancel_btn.pack(side=tk.LEFT, padx=5)
        
        # Cleanup on close
        def on_close():
            # Unbind all mouse wheel events
            preview_window.unbind_all("<MouseWheel>")
            preview_window.unbind_all("<Button-4>")
            preview_window.unbind_all("<Button-5>")
            if hasattr(self, '_preview_images'):
                self._preview_images.clear()
            preview_window.destroy()
        
        preview_window.protocol("WM_DELETE_WINDOW", lambda: cancel())
            
    def log(self, message, level="info"):
        """Add message to log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        if level == "error":
            prefix = "❌"
            color = "#ff6b6b"
        elif level == "success":
            prefix = "✓"
            color = "#51cf66"
        elif level == "warning":
            prefix = "⚠"
            color = "#ffd43b"
        else:
            prefix = "ℹ"
            color = "#74c0fc"
        
        self.log_text.insert(tk.END, f"[{timestamp}] {prefix} {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def process_orders(self):
        """Process the orders"""
        text = self.order_input.get(1.0, tk.END).strip()
        
        if not text or self.order_input.cget("fg") == "#999":
            messagebox.showwarning("No Orders", "Please add orders first.")
            return
        
        if self.processing:
            messagebox.showinfo("Processing", "Already processing orders...")
            return
        
        # Parse and validate orders
        orders = []
        lines = text.split('\n')
        for line in lines:
            line = line.strip()
            if not line:
                continue
            if ',' in line:
                parts = line.split(',')
                character = parts[0].strip()
                name = parts[1].strip() if len(parts) > 1 else ""
                year = parts[2].strip() if len(parts) > 2 else ""
                orders.append((character, name, year))
        
        if not orders:
            messagebox.showwarning("No Valid Orders", "Please add at least one valid order.")
            return
        
        # Save to temp CSV for processing
        import tempfile
        temp_csv = os.path.join(tempfile.gettempdir(), "temp_orders.csv")
        try:
            with open(temp_csv, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['character', 'name', 'year'])
                for character, name, year in orders:
                    writer.writerow([character, name, year])
            
            self.csv_path.set(temp_csv)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to prepare orders:\n{str(e)}")
            return
        
        # Run in thread to keep UI responsive
        thread = threading.Thread(target=self.process_orders_thread)
        thread.daemon = True
        thread.start()
        
    def cleanup_old_pdfs(self):
        """Archive old PDFs and keep only last 10"""
        try:
            # Create archive folder
            archive_dir = "pdf_archive"
            os.makedirs(archive_dir, exist_ok=True)
            
            # Find all order PDFs (magnets, boats, and master)
            pdf_files = []
            for file in os.listdir('.'):
                if (file.startswith('order_output_') or file.startswith('boat_output_') or file.startswith('MASTER_ORDER_')) and file.endswith('.pdf'):
                    pdf_files.append(file)
            
            if not pdf_files:
                return
            
            self.log(f"Archiving {len(pdf_files)} old PDF(s)...", "info")
            
            # Move PDFs to archive
            for pdf_file in pdf_files:
                try:
                    dest = os.path.join(archive_dir, pdf_file)
                    if os.path.exists(dest):
                        os.remove(dest)  # Remove if exists
                    shutil.move(pdf_file, dest)
                except Exception as e:
                    self.log(f"Could not archive {pdf_file}: {str(e)}", "warning")
            
            # Clean up old archives (keep last 10 master PDFs)
            archive_masters = []
            for file in os.listdir(archive_dir):
                if file.startswith('MASTER_ORDER_') and file.endswith('.pdf'):
                    file_path = os.path.join(archive_dir, file)
                    mtime = os.path.getmtime(file_path)
                    archive_masters.append((mtime, file_path))
            
            # Sort by modification time (newest first)
            archive_masters.sort(reverse=True)
            
            # Delete old ones beyond the 10 most recent
            if len(archive_masters) > 10:
                for _, file_path in archive_masters[10:]:
                    try:
                        os.remove(file_path)
                        self.log(f"Deleted old archive: {os.path.basename(file_path)}", "info")
                    except Exception as e:
                        self.log(f"Could not delete {file_path}: {str(e)}", "warning")
            
            self.log(f"✓ Archived old PDFs (keeping last 10 masters)", "success")
            
        except Exception as e:
            self.log(f"Warning: Cleanup failed: {str(e)}", "warning")
    
    def process_orders_thread(self):
        """Process orders in background thread"""
        try:
            self.processing = True
            self.process_btn.config(state=tk.DISABLED)
            self.status_text.set("Processing orders...")
            self.log_text.delete(1.0, tk.END)
            self.progress_var.set(0)
            
            # Clean up old PDFs first
            self.cleanup_old_pdfs()
            
            # Redirect stdout to log
            original_stdout = sys.stdout
            
            class LogRedirector:
                def __init__(self, log_func):
                    self.log_func = log_func
                    
                def write(self, message):
                    if message.strip():
                        self.log_func(message.strip())
                        
                def flush(self):
                    pass
            
            sys.stdout = LogRedirector(lambda msg: self.root.after(0, self.log, msg))
            
            self.log("Starting order processing...", "info")
            self.progress_var.set(10)
            
            # Call the processing function
            success = process_orders.process_all_orders(self.csv_path.get())
            
            # Restore stdout
            sys.stdout = original_stdout
            
            self.progress_var.set(100)
            
            if success:
                self.log("✓ Processing complete!", "success")
                self.view_btn.config(state=tk.NORMAL)
               
                # Find all individual magnet PDFs
                magnet_pdf_files = []
                for file in os.listdir('.'):
                    if file.startswith('order_output_') and file.endswith('.pdf'):
                        magnet_pdf_files.append(file)
                magnet_pdf_files.sort(key=extract_number)
                
                # Find all boat PDFs
                boat_pdf_files = []
                for file in os.listdir('.'):
                    if file.startswith('boat_output_') and file.endswith('.pdf'):
                        boat_pdf_files.append(file)
                boat_pdf_files.sort(key=extract_number)
                
                # Combine: magnets first, then boats at the end
                all_pdf_files = magnet_pdf_files + boat_pdf_files
               
                if all_pdf_files:
                    # === Flatten each individual PDF in place ===
                    self.log("Flattening individual PDFs (removing hidden layers/data)...", "info")
                    for pdf_file in all_pdf_files:
                        if os.path.exists(pdf_file):
                            self.flatten_pdf_in_place(pdf_file, dpi=300)
                    # ===========================================

                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    master_pdf_name = f"MASTER_ORDER_{timestamp}.pdf"
                   
                    if self.merge_pdfs(all_pdf_files, master_pdf_name):
                        # === Flatten the master PDF ===
                        self.log("Flattening master PDF...", "info")
                        self.flatten_pdf_in_place(master_pdf_name, dpi=300)
                        # ==============================

                        self.master_pdf_path = master_pdf_name
                        self.master_pdf_btn.config(state=tk.NORMAL)
                        self.status_text.set(f"✓ Complete! Master PDF: {master_pdf_name}")
                       
                        # Automatically open the (now flattened) master PDF
                        try:
                            self.log(f"Opening master PDF...", "info")
                            open_file_or_folder(master_pdf_name)
                            time.sleep(0.5)  # Give system time to open the file
                        except Exception as e:
                            self.log(f"Could not auto-open PDF: {e}", "warning")
                       
                        magnet_count = len(magnet_pdf_files)
                        boat_count = len(boat_pdf_files)
                        messagebox.showinfo(
                            "Success",
                            f"Orders processed and FLATTENED!\n\n"
                            f"✓ {magnet_count} magnet PDF(s) created\n"
                            f"✓ {boat_count} boat PDF(s) created\n"
                            f"✓ Master PDF created: {master_pdf_name}\n"
                            f"✓ No editable text or hidden template data remains\n\n"
                            f"Master PDF opened automatically!\n\n"
                            f"Check your PDF viewer!"
                        )
                    else:
                        self.status_text.set("✓ Orders processed (individual PDFs flattened)")
                        messagebox.showinfo(
                            "Success",
                            "Orders processed successfully!\n\n"
                            "Individual PDFs have been flattened.\n"
                            "Check current directory for PDFs and outputs/ folder for images."
                        )
                else:
                    self.status_text.set("✓ Images created (no PDFs to merge)")
                    messagebox.showinfo(
                        "Success",
                        "Images created successfully!\n\n"
                        "Check outputs/ folder for images"
                    )
            else:
                self.log("Processing completed with errors", "warning")
                self.status_text.set("Processing completed with errors")
                
        except Exception as e:
            self.log(f"ERROR: {str(e)}", "error")
            self.status_text.set(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Processing failed:\n{str(e)}")
            
        finally:
            self.processing = False
            self.process_btn.config(state=tk.NORMAL)
            
    def view_outputs(self):
        """Open outputs folder"""
        outputs_dir = "outputs"
        if os.path.exists(outputs_dir):
            try:
                open_file_or_folder(outputs_dir)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open folder:\n{str(e)}")
        else:
            messagebox.showinfo("No Outputs", "No outputs folder found yet.")
            
    def view_archive(self):
        """Open PDF archive folder"""
        archive_dir = "pdf_archive"
        if os.path.exists(archive_dir):
            try:
                open_file_or_folder(archive_dir)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open folder:\n{str(e)}")
        else:
            messagebox.showinfo("No Archive", "No archived PDFs yet. Process some orders first!")
            
    def open_master_pdf(self):
        """Open the master PDF"""
        if self.master_pdf_path and os.path.exists(self.master_pdf_path):
            try:
                open_file_or_folder(self.master_pdf_path)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open PDF:\n{str(e)}")
        else:
            messagebox.showinfo("No Master PDF", "Master PDF not found. Process orders first.")

    def flatten_pdf_in_place(self, pdf_path, dpi=300):
        """Rasterize all pages of a PDF and overwrite it with a flattened version."""
        try:
            doc = fitz.open(pdf_path)
            if doc.page_count == 0:
                doc.close()
                return

            # Create a temporary output path
            temp_path = pdf_path + ".flattened_tmp.pdf"

            writer = fitz.open()  # New empty PDF

            for page in doc:
                # Create a new blank page with the exact same dimensions
                new_page = writer.new_page(width=page.rect.width, height=page.rect.height)

                # Render the original page to a high-resolution pixmap
                zoom = dpi / 72.0
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat, alpha=False)

                # Insert the rendered image (covers the entire page)
                new_page.insert_image(new_page.rect, pixmap=pix)

            # Save to temp file then replace original
            writer.save(temp_path, garbage=4, deflate=True, clean=True)
            writer.close()
            doc.close()

            # Overwrite the original file
            os.replace(temp_path, pdf_path)

            self.log(f"Flattened: {os.path.basename(pdf_path)}", "success")
        except Exception as e:
            self.log(f"Failed to flatten {os.path.basename(pdf_path)}: {str(e)}", "error")
            # If flattening fails, keep the original (non-flattened) file                
    def merge_pdfs(self, pdf_files, output_path):
        """Merge multiple PDF files into one master PDF"""
        try:
            import fitz
            
            self.log(f"Merging {len(pdf_files)} PDFs into master PDF...", "info")
            
            writer = fitz.open()
            
            for pdf_file in pdf_files:
                if os.path.exists(pdf_file):
                    doc = fitz.open(pdf_file)
                    writer.insert_pdf(doc)
                    doc.close()
                    self.log(f"  Added {pdf_file}", "info")
            
            writer.save(output_path, garbage=4, deflate=True, clean=True)
            writer.close()
            
            self.log(f"✓ Master PDF created: {output_path}", "success")
            return True
            
        except Exception as e:
            self.log(f"Error merging PDFs: {str(e)}", "error")
            return False
            
    def clear_all(self):
        """Clear all fields"""
        self.csv_path.set("")
        self.order_input.delete(1.0, tk.END)
        placeholder = "mickey-captain,Johnny\nminnie-captain,Sarah\nstitch-captain,Michael\nmoana-captain,Emma"
        self.order_input.insert(1.0, placeholder)
        self.order_input.config(fg="#999")
        self.preview_text.delete(1.0, tk.END)
        self.log_text.delete(1.0, tk.END)
        self.orders_count.set("0 orders")
        self.progress_var.set(0)
        self.status_text.set("Ready to process orders")
        self.view_btn.config(state=tk.DISABLED)
        self.master_pdf_btn.config(state=tk.DISABLED)
        self.master_pdf_path = None
        
    def get_available_images(self):
        """Get list of available images from FHM_Images folder AND boats folder"""
        try:
            image_files = []
            
            # Get character images from FHM_Images (parent directory)
            parent_dir = os.path.dirname(os.path.abspath(os.getcwd()))
            images_dir = os.path.join(parent_dir, 'FHM_Images')
            
            if os.path.exists(images_dir):
                for file in os.listdir(images_dir):
                    if file.lower().endswith('.png'):
                        image_files.append(file)
            
            # Get boat images from boats folder (same directory as this script)
            boats_dir = "boats"
            if os.path.exists(boats_dir):
                for file in os.listdir(boats_dir):
                    if file.lower().endswith('.png'):
                        image_files.append(file)
            
            image_files.sort()
            return image_files
        except Exception as e:
            print(f"Error getting images: {e}")
            return []
    
    def clear_raw_placeholder(self, event):
        """Clear raw text placeholder"""
        current = self.raw_text.get(1.0, tk.END).strip()
        if "Paste order details here" in current:
            self.raw_text.delete(1.0, tk.END)
            self.raw_text.config(fg="#333")
            
    def restore_raw_placeholder(self, event):
        """Restore raw text placeholder if empty"""
        if not self.raw_text.get(1.0, tk.END).strip():
            placeholder = """Paste order details here, like:

Order #12345 - Mickey Captain themed, names: Johnny, Sarah, Michael
or
Disney Cruise Door Magnet - 3 magnets: Minnie, Donald, Goofy (all captain theme)"""
            self.raw_text.insert(1.0, placeholder)
            self.raw_text.config(fg="#999")
            
    def clear_raw_text(self):
        """Clear raw text area"""
        self.raw_text.delete(1.0, tk.END)
        self.raw_text.config(fg="#333")
        
    def load_raw_sample(self):
        """Load sample raw order text"""
        sample = """Order #3769495301 - $46.88 FLASHSALE
Disney Cruise Door Magnet Captain Themed Porthole Personalized

Customer: Sarah Johnson
Order Date: Jan 1, 2026

Personalization:
Mickey - Johnny
Minnie - Sarah  
Donald - Michael

Please also add:
Goofy for Emma
Pluto for Oliver
Stitch for Sophia"""
        self.raw_text.delete(1.0, tk.END)
        self.raw_text.insert(1.0, sample)
        self.raw_text.config(fg="#333")
        self.status_text.set("Sample order loaded - click 'Parse with AI' to process")
        
    def parse_with_ai(self):
        """Parse raw order text with Grok AI (reasoning model)"""
        raw_text = self.raw_text.get(1.0, tk.END).strip()
        
        if not raw_text or self.raw_text.cget("fg") == "#999":
            messagebox.showwarning("No Text", "Please paste order text first.")
            return
        
        if self.ai_processing:
            messagebox.showinfo("Processing", "AI is already processing...")
            return
        
        if not GROK_API_KEY:
            parent_dir = os.path.dirname(os.path.abspath(os.getcwd()))
            messagebox.showerror(
                "API Key Missing", 
                "Grok API key not found!\n\n"
                "Please create a file with your API key at:\n"
                f"• {parent_dir}\\grok_config.txt (recommended)\n"
                "OR\n"
                "• grok_config.txt (in this folder)\n\n"
                "See grok_config.txt.sample for instructions."
            )
            return
        
        if not self.image_list:
            messagebox.showerror("No Images", "No images found in FHM_Images folder.\nMake sure it's in the parent directory.")
            return
        
        # Run in thread with reasoning model
        thread = threading.Thread(target=self.parse_with_ai_thread, args=(raw_text, True))
        thread.daemon = True
        thread.start()
    
    def quick_parse_with_ai(self):
        """Quick parse with Grok AI (non-reasoning model for faster results)"""
        raw_text = self.raw_text.get(1.0, tk.END).strip()
        
        if not raw_text or self.raw_text.cget("fg") == "#999":
            messagebox.showwarning("No Text", "Please paste order text first.")
            return
        
        if self.ai_processing:
            messagebox.showinfo("Processing", "AI is already processing...")
            return
        
        if not GROK_API_KEY:
            parent_dir = os.path.dirname(os.path.abspath(os.getcwd()))
            messagebox.showerror(
                "API Key Missing", 
                "Grok API key not found!\n\n"
                "Please create a file with your API key at:\n"
                f"• {parent_dir}\\grok_config.txt (recommended)\n"
                "OR\n"
                "• grok_config.txt (in this folder)\n\n"
                "See grok_config.txt.sample for instructions."
            )
            return
        
        if not self.image_list:
            messagebox.showerror("No Images", "No images found in FHM_Images folder.\nMake sure it's in the parent directory.")
            return
        
        # Run in thread with non-reasoning model (faster)
        thread = threading.Thread(target=self.parse_with_ai_thread, args=(raw_text, False))
        thread.daemon = True
        thread.start()
    
    def format_with_ai_stage1(self, raw_text, use_reasoning=True):
        """STAGE 1: Format raw order text into simple character-name pairs"""
        model = "grok-4-1-fast-reasoning" if use_reasoning else "grok-4-1-fast-non-reasoning"
        
        prompt = f"""You are extracting character and name pairs from Disney magnet order text.

Your job is to format raw order data into a simple, clean list.

IMPORTANT RULES:
1. Extract ONLY character-name pairs (one per line)
2. Format EXACTLY as: "Character description - name: PersonName"
3. For boat orders, extract the boat with its ship name. Format as: "boat ShipName - name: FamilyName"
4. If a line has BOTH a boat AND regular character orders, extract BOTH the boat AND the regular characters
5. A single order line can have 1-5 character-name pairs (plus possibly a boat)
6. Keep character descriptions simple and natural (e.g., "Luke Skywalker", "Stitch captain", "Minnie Spiderman")
7. Preserve exact name spellings from the order
8. If no name is specified for a character, use "no name" for the name
9. Duck and dog orders do not need captain, pirate, etc. Just duck/dog and the ID number with them.
10. Do not omit any items from the order including boats.
11. The header for an order determines the theme. For example, if the header is pirate, every item in that order is pirate. Same with captain, christmas, etc. If the order is boat, assume captain theme for the character magnets.

=== BOAT ORDER DETECTION ===
If the order mentions ANY of these, it includes a BOAT order:
- "boat" or "ship"
- Disney cruise ship names: Fantasy, Magic, Wonder, Wish, Dream, Treasure, Destiny
- "cruise ship door decoration" or similar

BOAT SHIP NAME MAPPING:
- Disney Fantasy → boat Fantasy
- Disney Magic → boat Magic
- Disney Wonder → boat Wonder
- Disney Wish → boat Wish
- Disney Dream → boat Dream
- Disney Treasure → boat Treasure
- Disney Destiny → boat Destiny
- No name mentioned - boat No Name

EXAMPLES:

Input: "Item: Captain Mickey, Personalization: Johnny"
Output: Mickey captain - name: Johnny

Input: "Item: Christmas Elsa\\nPersonalization: Sarah"
Output: Elsa christmas - name: Sarah

Input: "captain Order has boat +  Minnie for 'Katie' and  Woody for 'Sean' and dog 16 for 'Joni'"
Output: 
boat Fantasy - name: Katie
Minnie captain - name: Katie
Woody captain - name: Sean
dog 16 - name: Joni

Input: "Disney Fantasy boat for The Smith Family"
Output:
boat Fantasy - name: The Smith Family

Input: "Disney Magic ship for Johnson Crew, also Mickey captain for Johnny"
Output:
boat Magic - name: Johnson Crew
Mickey captain - name: Johnny

Input: "1. Dory Captain - Joni  2. Ariel pirate - Gracie  3. Rapunzel Captain - Lila"
Output:
Dory captain - name: Joni
Ariel pirate - name: Gracie
Rapunzel captain - name: Lila

Now process this order text:
{raw_text}

OUTPUT FORMAT: Return ONLY the formatted lines, one per line, no explanations, no markdown, just the text:"""
        
        headers = {
            "Authorization": f"Bearer {GROK_API_KEY}",
            "Content-Type": "application/json"
        }
        
        data = {
            "model": model,
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.1,
            "max_tokens": 2000
        }
        
        try:
            self.root.after(0, lambda: self.log("Stage 1: Sending formatting request to Grok AI...", "info"))
            response = requests.post(GROK_API_URL, headers=headers, json=data, timeout=60)
            response.raise_for_status()
            
            result = response.json()
            content = result['choices'][0]['message']['content'].strip()
            
            # Clean up the response - remove markdown code blocks if present
            if content.startswith('```'):
                lines = content.split('\n')
                content = '\n'.join(line for line in lines if not line.startswith('```'))
                content = content.strip()
            
            self.root.after(0, lambda: self.log(f"Stage 1: Received formatted text ({len(content.split(chr(10)))} lines)", "info"))
            return content
            
        except Exception as e:
            error_msg = str(e)
            self.root.after(0, lambda msg=error_msg: self.log(f"Stage 1 Error: {msg}", "error"))
            return None
        
    def parse_with_ai_thread(self, raw_text, use_reasoning=True):
        """Parse with AI in background thread - 2-STAGE SYSTEM"""
        try:
            self.ai_processing = True
            # Disable both AI buttons
            self.root.after(0, lambda: self.ai_parse_btn.config(state=tk.DISABLED, text="🤖 Processing..."))
            self.root.after(0, lambda: self.quick_parse_btn.config(state=tk.DISABLED, text="⚡ Processing..."))
            
            model_type = "reasoning" if use_reasoning else "quick"
            self.root.after(0, lambda: self.status_text.set(f"AI is parsing your order ({model_type})..."))
            
            # === STAGE 1: Format the raw text into simple character-name format ===
            self.root.after(0, lambda: self.log(f"STAGE 1: Formatting raw order text...", "info"))
            formatted_text = self.format_with_ai_stage1(raw_text, use_reasoning)
            
            if not formatted_text:
                self.root.after(0, lambda: messagebox.showerror("AI Error", "Stage 1: Failed to format orders. Check the log for details."))
                self.root.after(0, lambda: self.log("Stage 1 formatting failed", "error"))
                return
            
            self.root.after(0, lambda: self.log(f"STAGE 1 Complete: Formatted text ready", "success"))
            self.root.after(0, lambda ft=formatted_text: self.log(f"Formatted output:\n{ft}", "info"))
            
            # Update the AI input field with Stage 1 formatted text
            self.root.after(0, lambda: self.raw_text.delete(1.0, tk.END))
            self.root.after(0, lambda ft=formatted_text: self.raw_text.insert(1.0, ft))
            self.root.after(0, lambda: self.raw_text.config(fg="black"))
            
            # === STAGE 2: Parse the formatted text to match images ===
            self.root.after(0, lambda: self.log(f"STAGE 2: Matching to images with {len(self.image_list)} available images...", "info"))
            result = self.call_grok_api(self.image_list, formatted_text, use_reasoning)
            
            if not result:
                self.root.after(0, lambda: messagebox.showerror("AI Error", "Stage 2: Failed to match images. Check the log for details."))
                self.root.after(0, lambda: self.log("Stage 2 image matching failed", "error"))
                return
            
            self.root.after(0, lambda: self.log(f"STAGE 2 Complete: Matched images", "success"))
            
            # Convert result to simple format - PRESERVE UNMATCHED ITEMS
            orders = []
            unmatched_count = 0
            
            # Handle both old dictionary format and new list format
            if isinstance(result, list):
                # New list format - supports duplicate names
                for item in result:
                    if isinstance(item, dict) and 'name' in item and 'image' in item:
                        name = item['name']
                        image_file = item['image']
                        original_item = item.get('item', '')  # Get original item description if available
                        
                        # Check for N/A or unmatched items
                        if image_file.lower() in ['n/a', 'n/a.png', 'unknown', 'unknown.png', 'not_found', 'not_found.png']:
                            # Include original item info for unmatched items
                            if original_item:
                                orders.append(f"IMAGE-NOT-FOUND [{original_item}],{name}")
                            else:
                                orders.append(f"IMAGE-NOT-FOUND,{name}")
                            unmatched_count += 1
                        else:
                            # Remove .png extension
                            character = image_file.replace('.png', '')
                            orders.append(f"{character},{name}")
            elif isinstance(result, dict):
                # Old dictionary format - kept for backwards compatibility
                for name, image_file in result.items():
                    if name not in ['_order', 'unmatched']:
                        if not image_file or image_file.lower() in ['n/a', 'unknown', 'not_found']:
                            orders.append(f"IMAGE-NOT-FOUND,{name}")
                            unmatched_count += 1
                        else:
                            # Remove .png extension
                            character = image_file.replace('.png', '')
                            orders.append(f"{character},{name}")
            
            if not orders:
                self.root.after(0, lambda: messagebox.showwarning("No Matches", "AI couldn't find any matching orders.\nTry being more specific or use manual entry."))
                self.root.after(0, lambda: self.log("No orders matched by AI", "warning"))
                return
            
            # Fill input area
            orders_text = '\n'.join(orders)
            self.root.after(0, lambda: self.order_input.delete(1.0, tk.END))
            self.root.after(0, lambda: self.order_input.insert(1.0, orders_text))
            self.root.after(0, lambda: self.order_input.config(fg="black"))
            self.root.after(0, lambda: self.update_count())
            self.root.after(0, lambda: self.preview_orders())
            
            # Show appropriate message based on matches
            if unmatched_count > 0:
                self.root.after(0, lambda: self.log(f"✓ 2-Stage AI Complete: {len(orders)} orders ({unmatched_count} need image selection)", "warning"))
                self.root.after(0, lambda: self.status_text.set(f"✓ AI found {len(orders)} orders - {unmatched_count} need image selection"))
                self.root.after(0, lambda um=unmatched_count, tot=len(orders): messagebox.showwarning(
                    "Partial Match",
                    f"2-Stage AI Processing Complete!\n\n"
                    f"✓ Stage 1: Formatted {tot} orders\n"
                    f"✓ Stage 2: Matched images\n\n"
                    f"⚠️ {um} item{'s' if um != 1 else ''} couldn't be matched to images.\n"
                    f"They are marked as 'IMAGE-NOT-FOUND'.\n\n"
                    f"In the preview window, you can search and select\n"
                    f"the correct images for these items.\n\n"
                    f"Click 'Preview Orders' to review and fix."
                ))
            else:
                self.root.after(0, lambda: self.log(f"✓ 2-Stage AI Complete: {len(orders)} orders successfully parsed!", "success"))
                self.root.after(0, lambda: self.status_text.set(f"✓ AI found {len(orders)} orders! Review and click Process."))
                self.root.after(0, lambda tot=len(orders): messagebox.showinfo("Success!", f"2-Stage AI Processing Complete!\n\n✓ Stage 1: Formatted {tot} orders\n✓ Stage 2: All images matched\n\nReview and click 'Process Orders' when ready."))
            
        except Exception as e:
            error_msg = str(e)
            self.root.after(0, lambda msg=error_msg: self.log(f"2-Stage AI Error: {msg}", "error"))
            self.root.after(0, lambda msg=error_msg: messagebox.showerror("Error", f"2-Stage AI parsing failed:\n{msg}"))
            
        finally:
            self.ai_processing = False
            # Re-enable both AI buttons
            self.root.after(0, lambda: self.ai_parse_btn.config(state=tk.NORMAL, text="✨ Parse with AI (Grok)"))
            self.root.after(0, lambda: self.quick_parse_btn.config(state=tk.NORMAL, text="⚡ Quick Parse"))
            
    def call_grok_api(self, image_list, order_text, use_reasoning=True):
        """STAGE 2: Call Grok API to match formatted orders to images"""
        # Format the complete list of available images
        # Send ALL images so AI can choose from exact matches
        images_formatted = '\n'.join(f"  - {img}" for img in image_list)
        
        # Choose model based on use_reasoning
        model = "grok-4-1-fast-reasoning" if use_reasoning else "grok-4-1-fast-non-reasoning"
        
        prompt = f"""You are matching Disney character descriptions to exact image filenames.

You will receive pre-formatted order text in this format:
"Character description - name: PersonName"

Your job is to match each character description to an EXACT image filename from the available list.

COMPLETE LIST OF AVAILABLE IMAGES ({len(image_list)} total):
{images_formatted}

Pre-formatted order text:
{order_text}

MATCHING RULES:
1. For each line, extract the character description and the name
2. Match the character description to an EXACT filename from the list above
3. You MUST use filenames EXACTLY as shown (including .png extension)
4. **CRITICAL: If you cannot find a match, use "N/A.png" - DO NOT skip the item!**
4.5 if there is a request for a captain, but there is only a normal image, use the normal image.
5. Common patterns to match:
   - "Mickey captain" → "mickey-captain.png"
   - "Stitch captain" → "stitch-captain.png"
   - "Elsa christmas" → "elsa-christmas.png"
   - "Minnie Spiderman" → "minnie-spiderman.png"
   - "Minnie pirate" → "minnie-pirate.png"
   - "Donald Hulk" → "donald-hulk.png"
   - "dog 16" → "dog-16.png"
   - "duck 23" → "duck-23.png"
6. Character names are case-insensitive for matching
7. ALWAYS include ALL items, even if no match found (use "N/A.png")
8. for magnets with no name, leave it blank. (example "name": "")

=== BOAT IMAGE MATCHING ===
Boat orders appear as "boat ShipName - name: FamilyName"
Match boat orders to these EXACT filenames:
- "boat Fantasy" → "boat_fantasy.png"
- "boat Magic" → "boat_magic.png"
- "boat Wonder" → "boat_wonder.png"
- "boat Wish" → "boat_wish.png"
- "boat Dream" → "boat_dream.png"
- "boat Treasure" → "boat_treasure.png"
- "boat Destiny" → "boat_destiny.png"
- If ship name is not present, default to "boat_noname.png"

OUTPUT FORMAT - Return ONLY a Python LIST of dictionaries, no other text:
[
  {{"name": "PersonName1", "image": "exact-filename.png", "item": "original character description"}},
  {{"name": "PersonName2", "image": "exact-filename.png", "item": "original character description"}},
  {{"name": "", "image": "N/A.png", "item": "original character description"}}
]

EXAMPLES:

Input: "Mickey captain - name: Johnny"
Output: [{{"name": "Johnny", "image": "mickey-captain.png", "item": "Mickey captain"}}]

Input: "Stitch pirate - name: Michael\\nMinnie Spiderman - name: Cecile"
Output: [
  {{"name": "Michael", "image": "stitch-pirate.png", "item": "Stitch pirate"}},
  {{"name": "Cecile", "image": "minnie-spiderman.png", "item": "Minnie Spiderman"}}
]

Input: "boat Fantasy - name: The Smith Family"
Output: [{{"name": "The Smith Family", "image": "boat_fantasy.png", "item": "boat Fantasy"}}]

Input: "Mickey captain - name: Johnny\\nboat Magic - name: Johnson Crew\\nMinnie captain - name: Sarah"
Output: [
  {{"name": "Johnny", "image": "mickey-captain.png", "item": "Mickey captain"}},
  {{"name": "Johnson Crew", "image": "boat_magic.png", "item": "boat Magic"}},
  {{"name": "Sarah", "image": "minnie-captain.png", "item": "Minnie captain"}}
]

Input: "RareCharacter - name: Test"
Output: [{{"name": "Test", "image": "N/A.png", "item": "RareCharacter"}}]

CRITICAL: 
- Only use filenames that EXACTLY match the list above
- If no match found, use "N/A.png" - DO NOT OMIT THE ITEM
- Return a LIST, preserving order
- Include ALL items from the input
- Boat orders use boat_*.png files, magnet orders use character-*.png files

Return the list now:"""
        
        headers = {
            "Authorization": f"Bearer {GROK_API_KEY}",
            "Content-Type": "application/json"
        }
        
        data = {
            "model": model,
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.1,
            "max_tokens": 2000  # Sufficient for list output format
        }
        
        try:
            self.root.after(0, lambda: self.log("Sending request to Grok AI with complete image list...", "info"))
            response = requests.post(GROK_API_URL, headers=headers, json=data, timeout=60)  # Increased timeout for larger prompt
            response.raise_for_status()
            
            result = response.json()
            content = result['choices'][0]['message']['content']
            
            self.root.after(0, lambda: self.log(f"Received response from AI", "info"))
            
            # Try to extract list format first (new format)
            list_start = content.find('[')
            list_end = content.rfind(']') + 1
            
            if list_start != -1 and list_end != 0:
                try:
                    list_str = content[list_start:list_end]
                    matches = json.loads(list_str)
                    if isinstance(matches, list):
                        self.root.after(0, lambda: self.log(f"Parsed {len(matches)} orders from AI (list format)", "info"))
                        return matches
                except json.JSONDecodeError:
                    pass
            
            # Fallback to dictionary format (old format)
            dict_start = content.find('{')
            dict_end = content.rfind('}') + 1
            
            if dict_start != -1 and dict_end != 0:
                try:
                    dict_str = content[dict_start:dict_end]
                    matches = json.loads(dict_str)
                    self.root.after(0, lambda: self.log(f"Parsed {len(matches)} orders from AI (dict format)", "info"))
                    return matches
                except json.JSONDecodeError:
                    pass
            
            self.root.after(0, lambda: self.log("Could not parse AI response", "error"))
            return {}
                
        except Exception as e:
            error_msg = str(e)
            self.root.after(0, lambda msg=error_msg: self.log(f"API Error: {msg}", "error"))
            return {}
    
    def zoom_in(self):
        """Increase app size"""
        if self.zoom_level < 1.5:  # Max 150%
            self.zoom_level += 0.1
            self.apply_zoom()
    
    def zoom_out(self):
        """Decrease app size"""
        if self.zoom_level > 0.7:  # Min 70%
            self.zoom_level -= 0.1
            self.apply_zoom()
    
    def apply_zoom(self):
        """Apply zoom level to window"""
        # Update zoom display
        zoom_percent = int(self.zoom_level * 100)
        self.zoom_display.config(text=f"{zoom_percent}%")
        
        # Calculate new size
        base_width = 1000
        base_height = 750
        new_width = int(base_width * self.zoom_level)
        new_height = int(base_height * self.zoom_level)
        
        # Get current position
        x = self.root.winfo_x()
        y = self.root.winfo_y()
        
        # Apply new size (maintain position)
        self.root.geometry(f'{new_width}x{new_height}+{x}+{y}')
        self.root.update_idletasks()
    
    def show_help(self):
        """Show help dialog"""
        help_text = """
Disney Magnet Order Processor - Quick Guide

🤖 AI PARSING (EASIEST!):
1. Paste raw order text (emails, messages, anything!)
2. Click "Parse with AI"
3. AI automatically extracts characters and names
4. Review and click "Process Orders"

Example raw text:
"Order #123 - Captain themed, 3 magnets
Names: Johnny (Mickey), Sarah (Minnie), Mike (Donald)"

AI will parse it to:
mickey-captain,Johnny
minnie-captain,Sarah
donald-captain,Mike

📝 MANUAL ENTRY:
Type or paste directly, one order per line:
• Format: character-name,PersonName
• Example: mickey-captain,Johnny

You can also:
• Click "Load from CSV" to import from a file
• Click "Paste Sample" to see examples
• Copy from Excel and paste directly

🎯 How to Use:
OPTION 1 (AI): Paste raw text → Parse with AI → Process
OPTION 2 (Manual): Type orders → Preview → Process

After processing:
• Click "Open Master PDF" to view all orders in one file
• Old PDFs automatically archived (keeps last 10)
• Click "View Archive" to see previous orders
• Images in outputs/ folder

📁 Required Files:
• FHM_Images folder in parent directory
• format.pdf template in current directory
• font/ folder with required fonts

💡 Pro Tips:
• AI is smart about typos and variations
• AI understands themes (captain, halloween, normal)
• Character names: mickey, minnie, donald, stitch, elsa, moana, etc.
• Variants: -captain, -normal, -pumpkin, -witch, -halloween
• You can paste from Excel/CSV directly!

For detailed help, see README.md or QUICKSTART.md
        """
        
        help_window = tk.Toplevel(self.root)
        help_window.title("Help")
        help_window.geometry("600x500")
        help_window.configure(bg="white")
        
        help_label = tk.Label(
            help_window,
            text="❓ Help & Instructions",
            font=("Segoe UI", 14, "bold"),
            bg="white"
        )
        help_label.pack(pady=10)
        
        help_scroll = scrolledtext.ScrolledText(
            help_window,
            font=("Segoe UI", 10),
            bg="white",
            wrap=tk.WORD
        )
        help_scroll.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))
        help_scroll.insert(1.0, help_text)
        help_scroll.config(state=tk.DISABLED)


def main():
    """Main application entry point"""
    if TkinterDnD is not None:
        try:
            # Try to use TkinterDnD for drag and drop
            root = TkinterDnD.Tk()
        except:
            # Fallback to regular Tk if TkinterDnD fails
            root = tk.Tk()
    else:
        # Use regular Tk if TkinterDnD not available
        root = tk.Tk()
    
    app = OrderProcessorGUI(root)
    
    # Center window
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    root.mainloop()


if __name__ == "__main__":
    main()

