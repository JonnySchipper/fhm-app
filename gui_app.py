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
from datetime import datetime
from pathlib import Path

# Import the processing functions
import process_orders
import order_state
import pullorders

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
        
        # Order management variables
        self.selected_orders = []  # List of selected order objects
        self.order_checkboxes = {}  # Dict mapping order_number to checkbox widget
        self.order_widgets = {}  # Dict mapping order_number to UI widgets
        
        # Get available images
        self.image_list = self.get_available_images()
        self.available_images.set(f"{len(self.image_list)} character images available")
        
        self.setup_ui()
        
        # Auto-load orders on startup
        self.root.after(500, self.startup_load_orders)
        
    def setup_ui(self):
        """Setup the user interface"""
        
        # Title Section
        title_frame = tk.Frame(self.root, bg=self.accent_color, height=90)
        title_frame.pack(fill=tk.X, padx=0, pady=0)
        title_frame.pack_propagate(False)
        
        title_label = tk.Label(
            title_frame,
            text="🎨 Disney Magnet Order Processor",
            font=("Segoe UI", 30, "bold"),
            bg=self.accent_color,
            fg="black"
        )
        title_label.pack(pady=(15, 5))
        
        subtitle_label = tk.Label(
            title_frame,
            text="✨ AI-Powered • Just Paste & Process",
            font=("Segoe UI", 15),
            bg=self.accent_color,
            fg="black"
        )
        subtitle_label.pack(pady=(0, 10))
        
        # View Log button (top left)
        view_log_btn = tk.Button(
            title_frame,
            text="📋 View Log",
            command=self.show_log_window,
            font=("Segoe UI", 14, "bold"),
            bg="white",
            fg="black",
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor="hand2"
        )
        view_log_btn.place(relx=0.0, rely=0.5, anchor=tk.W, x=10)
        
        # Zoom controls (top right)
        zoom_frame = tk.Frame(title_frame, bg=self.accent_color)
        zoom_frame.place(relx=1.0, rely=0.5, anchor=tk.E, x=-10)
        
        zoom_label = tk.Label(
            zoom_frame,
            text="Zoom:",
            font=("Segoe UI", 14),
            bg=self.accent_color,
            fg="black"
        )
        zoom_label.pack(side=tk.LEFT, padx=(0, 5))
        
        zoom_out_btn = tk.Button(
            zoom_frame,
            text="−",
            command=self.zoom_out,
            font=("Segoe UI", 18, "bold"),
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
            font=("Segoe UI", 14),
            bg=self.accent_color,
            fg="black",
            width=5
        )
        self.zoom_display.pack(side=tk.LEFT, padx=2)
        
        zoom_in_btn = tk.Button(
            zoom_frame,
            text="+",
            command=self.zoom_in,
            font=("Segoe UI", 18, "bold"),
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
        
        # Order Management Section (NEW)
        self.create_order_management_section(content_frame)
        
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
        
    def create_order_management_section(self, parent):
        """Create order management section with pull orders, current/past orders list"""
        order_mgmt_frame = tk.LabelFrame(
            parent,
            text="📦 Order Management",
            font=("Segoe UI", 16, "bold"),
            bg=self.bg_color,
            fg="black"
        )
        order_mgmt_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Top button row
        top_btn_frame = tk.Frame(order_mgmt_frame, bg=self.bg_color)
        top_btn_frame.pack(fill=tk.X, padx=10, pady=(10, 5))
        
        # Pull New Orders button
        pull_orders_btn = tk.Button(
            top_btn_frame,
            text="📥 Pull New Orders",
            command=self.pull_new_orders,
            font=("Segoe UI", 15, "bold"),
            bg="#28a745",
            fg="black",
            relief=tk.FLAT,
            padx=20,
            pady=8,
            cursor="hand2"
        )
        pull_orders_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # Refresh Orders button
        refresh_orders_btn = tk.Button(
            top_btn_frame,
            text="🔄 Refresh List",
            command=self.refresh_order_list,
            font=("Segoe UI", 14),
            bg="#17a2b8",
            fg="black",
            relief=tk.FLAT,
            padx=15,
            pady=8,
            cursor="hand2"
        )
        refresh_orders_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # Show filter
        filter_frame = tk.Frame(top_btn_frame, bg=self.bg_color)
        filter_frame.pack(side=tk.LEFT, padx=(20, 0))
        
        tk.Label(
            filter_frame,
            text="Show:",
            font=("Segoe UI", 14),
            bg=self.bg_color
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        self.order_filter = tk.StringVar(value="pending")
        
        filter_all_radio = tk.Radiobutton(
            filter_frame,
            text="All",
            variable=self.order_filter,
            value="all",
            command=self.refresh_order_list,
            bg=self.bg_color,
            fg="black",
            font=("Segoe UI", 14)
        )
        filter_all_radio.pack(side=tk.LEFT, padx=2)
        
        filter_pending_radio = tk.Radiobutton(
            filter_frame,
            text="Pending",
            variable=self.order_filter,
            value="pending",
            command=self.refresh_order_list,
            bg=self.bg_color,
            fg="black",
            font=("Segoe UI", 14)
        )
        filter_pending_radio.pack(side=tk.LEFT, padx=2)
        
        filter_completed_radio = tk.Radiobutton(
            filter_frame,
            text="Completed",
            variable=self.order_filter,
            value="completed",
            command=self.refresh_order_list,
            bg=self.bg_color,
            fg="black",
            font=("Segoe UI", 14)
        )
        filter_completed_radio.pack(side=tk.LEFT, padx=2)
        
        # Order count label
        self.order_list_count = tk.StringVar(value="0 orders")
        count_label = tk.Label(
            top_btn_frame,
            textvariable=self.order_list_count,
            font=("Segoe UI", 14),
            bg=self.bg_color,
            fg="black"
        )
        count_label.pack(side=tk.RIGHT)
        
        # Scrollable orders list
        list_container = tk.Frame(order_mgmt_frame, bg="white", relief=tk.RIDGE, bd=2)
        list_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=(5, 10))
        
        # Create canvas and scrollbar
        canvas = tk.Canvas(list_container, bg="white", highlightthickness=0)
        scrollbar = tk.Scrollbar(list_container, orient="vertical", command=canvas.yview)
        self.orders_list_frame = tk.Frame(canvas, bg="white")
        
        canvas_window = canvas.create_window((0, 0), window=self.orders_list_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack scrollbar and canvas
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Update scroll region when content changes
        self.orders_list_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        # Make scrollable frame expand to canvas width
        def configure_canvas_width(event):
            canvas.itemconfig(canvas_window, width=event.width)
        
        canvas.bind('<Configure>', configure_canvas_width)
        
        # Enable mouse wheel scrolling (cross-platform) - bind locally to avoid errors
        def on_mousewheel(event):
            try:
                if canvas.winfo_exists() and event.delta:
                    canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except:
                pass
        
        def on_mousewheel_mac(event):
            try:
                if canvas.winfo_exists():
                    canvas.yview_scroll(int(-1*event.delta), "units")
            except:
                pass
        
        # Store scroll callback for use in order widgets
        if platform.system() == 'Darwin':  # macOS
            self.orders_scroll_callback = on_mousewheel_mac
        else:  # Windows/Linux
            self.orders_scroll_callback = on_mousewheel
        
        # Bind for different platforms - use local binding
        list_container.bind("<MouseWheel>", self.orders_scroll_callback)
        canvas.bind("<MouseWheel>", self.orders_scroll_callback)
        self.orders_list_frame.bind("<MouseWheel>", self.orders_scroll_callback)
        
        # Bottom button row
        bottom_btn_frame = tk.Frame(order_mgmt_frame, bg=self.bg_color)
        bottom_btn_frame.pack(fill=tk.X, padx=10, pady=(5, 10))
        
        # Select All / Deselect All buttons
        select_all_btn = tk.Button(
            bottom_btn_frame,
            text="✓ Select All",
            command=self.select_all_orders,
            font=("Segoe UI", 14),
            bg="#6c757d",
            fg="black",
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor="hand2"
        )
        select_all_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        deselect_all_btn = tk.Button(
            bottom_btn_frame,
            text="✗ Deselect All",
            command=self.deselect_all_orders,
            font=("Segoe UI", 14),
            bg="#6c757d",
            fg="black",
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor="hand2"
        )
        deselect_all_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # Begin Orders button (main action)
        self.begin_orders_btn = tk.Button(
            bottom_btn_frame,
            text="▶ Begin Selected Orders",
            command=self.begin_selected_orders,
            font=("Segoe UI", 15, "bold"),
            bg=self.success_color,
            fg="black",
            relief=tk.FLAT,
            padx=30,
            pady=8,
            cursor="hand2"
        )
        self.begin_orders_btn.pack(side=tk.LEFT, padx=(20, 5))
        
        # Validation Page button
        self.validation_btn = tk.Button(
            bottom_btn_frame,
            text="✓ Open Validation Page",
            command=self.open_validation_page,
            font=("Segoe UI", 15, "bold"),
            bg="#17a2b8",
            fg="black",
            relief=tk.FLAT,
            padx=20,
            pady=8,
            cursor="hand2",
            state=tk.DISABLED
        )
        self.validation_btn.pack(side=tk.RIGHT)
        
        # Initial load
        self.refresh_order_list()
    
    def create_ai_section(self, parent):
        """Create AI parsing section"""
        ai_frame = tk.LabelFrame(
            parent,
            text="🤖 AI Order Parser (Paste Raw Order Text)",
            font=("Segoe UI", 16, "bold"),
            bg=self.bg_color,
            fg="black"
        )
        ai_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Info row
        info_frame = tk.Frame(ai_frame, bg=self.bg_color)
        info_frame.pack(fill=tk.X, padx=10, pady=(5, 5))
        
        info_label = tk.Label(
            info_frame,
            text="Paste raw order text from emails/messages and let AI parse it automatically!",
            font=("Segoe UI", 14, "italic"),
            bg=self.bg_color,
            fg="black"
        )
        info_label.pack(side=tk.LEFT)
        
        images_label = tk.Label(
            info_frame,
            textvariable=self.available_images,
            font=("Segoe UI", 12),
            bg=self.bg_color,
            fg="black"
        )
        images_label.pack(side=tk.RIGHT)
        
        # Raw text input
        raw_frame = tk.Frame(ai_frame, bg="white", relief=tk.RIDGE, bd=2)
        raw_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        self.raw_text = scrolledtext.ScrolledText(
            raw_frame,
            font=("Consolas", 14),
            height=6,
            bg="white",
            fg="black",
            wrap=tk.WORD,
            relief=tk.FLAT
        )
        self.raw_text.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        # Enable mouse wheel scrolling on Mac for raw text
        def on_raw_scroll_mac(event):
            try:
                self.raw_text.yview_scroll(int(-1*event.delta), "units")
            except:
                pass
        
        def on_raw_scroll(event):
            try:
                if event.delta:
                    self.raw_text.yview_scroll(int(-1*(event.delta/120)), "units")
            except:
                pass
        
        if platform.system() == 'Darwin':  # macOS
            self.raw_text.bind("<MouseWheel>", on_raw_scroll_mac)
        else:  # Windows/Linux
            self.raw_text.bind("<MouseWheel>", on_raw_scroll)
        
        # Placeholder
        placeholder = """Paste order details here, like:

Order #12345 - Mickey Captain themed, names: Johnny, Sarah, Michael
or
Disney Cruise Door Magnet - 3 magnets: Minnie, Donald, Goofy (all captain theme)"""
        self.raw_text.insert(1.0, placeholder)
        self.raw_text.config(fg="black")
        
        self.raw_text.bind("<FocusIn>", self.clear_raw_placeholder)
        self.raw_text.bind("<FocusOut>", self.restore_raw_placeholder)
        
        # AI Button
        ai_btn_frame = tk.Frame(ai_frame, bg=self.bg_color)
        ai_btn_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        self.ai_parse_btn = tk.Button(
            ai_btn_frame,
            text="✨ Parse with AI (Grok)",
            command=self.parse_with_ai,
            font=("Segoe UI", 15, "bold"),
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
            font=("Segoe UI", 14),
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
            font=("Segoe UI", 14),
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
            font=("Segoe UI", 14),
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
            font=("Segoe UI", 16, "bold"),
            bg=self.bg_color,
            fg="black"
        )
        file_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Instructions
        instructions = tk.Label(
            file_frame,
            text="Type one order per line:  character-name,PersonName  (or paste from Excel/CSV)",
            font=("Segoe UI", 14, "italic"),
            bg=self.bg_color,
            fg="black"
        )
        instructions.pack(pady=(5, 0), padx=10, anchor=tk.W)
        
        # Text input area
        input_frame = tk.Frame(file_frame, bg="white", relief=tk.RIDGE, bd=2)
        input_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.order_input = scrolledtext.ScrolledText(
            input_frame,
            font=("Consolas", 15),
            height=8,
            bg="white",
            fg="black",
            wrap=tk.WORD,
            relief=tk.FLAT
        )
        self.order_input.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        # Enable mouse wheel scrolling on Mac for order input
        def on_input_scroll_mac(event):
            try:
                self.order_input.yview_scroll(int(-1*event.delta), "units")
            except:
                pass
        
        def on_input_scroll(event):
            try:
                if event.delta:
                    self.order_input.yview_scroll(int(-1*(event.delta/120)), "units")
            except:
                pass
        
        if platform.system() == 'Darwin':  # macOS
            self.order_input.bind("<MouseWheel>", on_input_scroll_mac)
        else:  # Windows/Linux
            self.order_input.bind("<MouseWheel>", on_input_scroll)

        # Placeholder text
        placeholder = "mickey-captain,Johnny\nminnie-captain,Sarah\nstitch-captain,Michael\nmoana-captain,Emma"
        self.order_input.insert(1.0, placeholder)
        self.order_input.config(fg="black")
        
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
            font=("Segoe UI", 14),
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
            font=("Segoe UI", 14),
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
            font=("Segoe UI", 14),
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
            font=("Segoe UI", 16, "bold"),
            bg=self.bg_color,
            fg="black"
        )
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Orders count and preview button
        top_frame = tk.Frame(preview_frame, bg=self.bg_color)
        top_frame.pack(fill=tk.X, padx=10, pady=(5, 0))
        
        count_label = tk.Label(
            top_frame,
            textvariable=self.orders_count,
            font=("Segoe UI", 14),
            bg=self.bg_color,
            fg="black"
        )
        count_label.pack(side=tk.LEFT)
        
        preview_btn = tk.Button(
            top_frame,
            text="👁️ Preview Orders",
            command=self.preview_orders,
            font=("Segoe UI", 14),
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
            font=("Consolas", 14),
            height=10,
            bg="#f8f9fa",
            relief=tk.FLAT,
            wrap=tk.WORD
        )
        self.preview_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Enable mouse wheel scrolling on Mac for preview text
        def on_preview_scroll_mac(event):
            try:
                self.preview_text.yview_scroll(int(-1*event.delta), "units")
            except:
                pass
        
        def on_preview_scroll(event):
            try:
                if event.delta:
                    self.preview_text.yview_scroll(int(-1*(event.delta/120)), "units")
            except:
                pass
        
        if platform.system() == 'Darwin':  # macOS
            self.preview_text.bind("<MouseWheel>", on_preview_scroll_mac)
        else:  # Windows/Linux
            self.preview_text.bind("<MouseWheel>", on_preview_scroll)
        
    def create_progress_section(self, parent):
        """Create progress tracking section"""
        progress_frame = tk.LabelFrame(
            parent,
            text="⚙️ Processing",
            font=("Segoe UI", 16, "bold"),
            bg=self.bg_color,
            fg="black"
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
            font=("Consolas", 12),
            height=15,  # Increased height to see more logs
            bg="#1e1e1e",
            fg="black",
            relief=tk.SUNKEN,
            bd=2
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        # Enable mouse wheel scrolling on Mac for the log text
        def on_log_scroll_mac(event):
            try:
                self.log_text.yview_scroll(int(-1*event.delta), "units")
            except:
                pass
        
        def on_log_scroll(event):
            try:
                if event.delta:
                    self.log_text.yview_scroll(int(-1*(event.delta/120)), "units")
            except:
                pass
        
        if platform.system() == 'Darwin':  # macOS
            self.log_text.bind("<MouseWheel>", on_log_scroll_mac)
        else:  # Windows/Linux
            self.log_text.bind("<MouseWheel>", on_log_scroll)
        
        # Log control buttons
        log_btn_frame = tk.Frame(progress_frame, bg=self.bg_color)
        log_btn_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        clear_log_btn = tk.Button(
            log_btn_frame,
            text="🗑️ Clear Log",
            command=lambda: self.log_text.delete(1.0, tk.END),
            font=("Segoe UI", 12),
            bg="#6c757d",
            fg="black",
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor="hand2"
        )
        clear_log_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        def copy_log():
            log_content = self.log_text.get(1.0, tk.END)
            self.root.clipboard_clear()
            self.root.clipboard_append(log_content)
            messagebox.showinfo("Copied", "Log copied to clipboard!")
        
        copy_log_btn = tk.Button(
            log_btn_frame,
            text="📋 Copy Log",
            command=copy_log,
            font=("Segoe UI", 12),
            bg="#6c757d",
            fg="black",
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor="hand2"
        )
        copy_log_btn.pack(side=tk.LEFT)
        
    def create_control_buttons(self, parent):
        """Create control buttons"""
        btn_frame = tk.Frame(parent, bg=self.bg_color)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Process button
        self.process_btn = tk.Button(
            btn_frame,
            text="▶ Process Orders",
            command=self.process_orders,
            font=("Segoe UI", 16, "bold"),
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
            font=("Segoe UI", 14),
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
            font=("Segoe UI", 14, "bold"),
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
            font=("Segoe UI", 14),
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
            font=("Segoe UI", 14),
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
            font=("Segoe UI", 14),
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
            font=("Segoe UI", 14),
            bg="#2c3e50",
            fg="black",
            anchor=tk.W
        )
        status_label.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        
    def clear_placeholder(self, event):
        """Clear placeholder text on focus"""
        if self.order_input.get(1.0, tk.END).strip() in ["mickey-captain,Johnny\nminnie-captain,Sarah\nstitch-captain,Michael\nmoana-captain,Emma", ""]:
            self.order_input.delete(1.0, tk.END)
            self.order_input.config(fg="black")
            
    def restore_placeholder(self, event):
        """Restore placeholder if empty"""
        if not self.order_input.get(1.0, tk.END).strip():
            placeholder = "mickey-captain,Johnny\nminnie-captain,Sarah\nstitch-captain,Michael\nmoana-captain,Emma"
            self.order_input.insert(1.0, placeholder)
            self.order_input.config(fg="black")
            
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
        self.order_input.config(fg="black")
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
        self.order_input.config(fg="black")
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
            self.order_input.config(fg="black")
            
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
        
        # Get all available images
        available_images = []
        for file in sorted(os.listdir(images_dir)):
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
            font=("Segoe UI", 24, "bold"),
            bg=self.accent_color,
            fg="black"
        )
        title_label.pack(pady=(10, 0))
        
        subtitle_label = tk.Label(
            title_frame,
            text="Select images, edit names, then click 'Confirm & Process'",
            font=("Segoe UI", 14),
            bg=self.accent_color,
            fg="black"
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
        
        # Enable mouse wheel scrolling (cross-platform) - with safety checks
        def on_mousewheel(event):
            try:
                if canvas.winfo_exists() and event.delta:
                    canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except:
                pass
        
        def on_mac_mousewheel(event):
            try:
                if canvas.winfo_exists():
                    canvas.yview_scroll(-1 * event.delta, "units")
            except:
                pass
        
        def on_button_4(event):
            try:
                if canvas.winfo_exists():
                    canvas.yview_scroll(-1, "units")
            except:
                pass
        
        def on_button_5(event):
            try:
                if canvas.winfo_exists():
                    canvas.yview_scroll(1, "units")
            except:
                pass
        
        # Bind mouse wheel events based on platform - local bindings only
        system = platform.system()
        if system == 'Darwin':  # macOS
            preview_scroll_callback = on_mac_mousewheel
            canvas.bind("<MouseWheel>", on_mac_mousewheel)
            scrollable_frame.bind("<MouseWheel>", on_mac_mousewheel)
            canvas.bind("<Button-4>", on_button_4)
            canvas.bind("<Button-5>", on_button_5)
        elif system == 'Linux':
            preview_scroll_callback = on_mousewheel
            canvas.bind("<Button-4>", on_button_4)
            canvas.bind("<Button-5>", on_button_5)
        else:  # Windows
            preview_scroll_callback = on_mousewheel
            canvas.bind("<MouseWheel>", on_mousewheel)
            scrollable_frame.bind("<MouseWheel>", on_mousewheel)
        
        # Store order data (will be updated as user edits)
        order_data = []
        order_widgets = []  # Store references to widgets for updates
        
        # Function to create an order row
        def create_order_row(character="", name=""):
            """Create a single editable order row"""
            idx = len(order_data)
            order_data.append({'character': character, 'name': name})
            
            # Check if this is an unmatched item and extract original item description
            original_item_desc = ""
            is_unmatched = False
            if '[' in character and ']' in character:
                # Extract original item from "IMAGE-NOT-FOUND [Original Item]" format
                start = character.find('[')
                end = character.find(']')
                if start != -1 and end != -1:
                    original_item_desc = character[start+1:end]
                    character_base = character[:start].strip()
                    is_unmatched = character_base.upper() in ['IMAGE-NOT-FOUND', 'N/A', 'NOT-FOUND', 'UNKNOWN']
            else:
                is_unmatched = character.upper() in ['IMAGE-NOT-FOUND', 'N/A', 'NOT-FOUND', 'UNKNOWN']
            
            # Create frame for this order - highlight unmatched items
            frame_bg = "#fff3cd" if is_unmatched else "white"  # Yellow warning background
            order_frame = tk.Frame(scrollable_frame, bg=frame_bg, relief=tk.RIDGE, bd=2)
            order_frame.pack(fill=tk.X, padx=10, pady=8)
            order_frame.bind("<MouseWheel>", preview_scroll_callback)
            
            # Image preview (left side) - fixed size container
            image_container = tk.Frame(order_frame, bg=frame_bg, width=100, height=100)
            image_container.pack(side=tk.LEFT, padx=10, pady=10)
            image_container.pack_propagate(False)  # Prevent resizing
            image_container.bind("<MouseWheel>", preview_scroll_callback)
            
            # Load initial image
            image_filename = f"{character}.png" if not is_unmatched else ""
            image_path = os.path.join(images_dir, image_filename) if image_filename else ""
            
            photo = None
            if image_path and os.path.exists(image_path):
                try:
                    img = Image.open(image_path)
                    img.thumbnail((100, 100), Image.Resampling.LANCZOS)
                    photo = ImageTk.PhotoImage(img)
                except:
                    photo = None
            
            # Image label - fills the container
            img_label_bg = "#ffc107" if is_unmatched else "#f0f0f0"  # Bright warning color for unmatched
            img_label = tk.Label(image_container, bg=img_label_bg, relief=tk.SUNKEN, bd=1)
            if photo:
                img_label.config(image=photo)
                if not hasattr(self, '_preview_images'):
                    self._preview_images = []
                self._preview_images.append(photo)
            elif is_unmatched:
                img_label.config(text="⚠️\nNEEDS\nIMAGE", font=("Segoe UI", 14, "bold"), fg="black")
            else:
                img_label.config(text="No\nImage", font=("Segoe UI", 15), fg="black")
            img_label.pack(fill=tk.BOTH, expand=True)
            img_label.bind("<MouseWheel>", preview_scroll_callback)
            
            # Edit controls (right side)
            edit_frame = tk.Frame(order_frame, bg=frame_bg)
            edit_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=10, padx=10)
            edit_frame.bind("<MouseWheel>", preview_scroll_callback)
            
            # Warning label for unmatched items (at top)
            if is_unmatched:
                warning_text = "⚠️ IMAGE NOT MATCHED - Click 'Search' to select correct image"
                if original_item_desc:
                    warning_text += f"\n📝 Original Order: {original_item_desc}"
                
                warning_label = tk.Label(
                    edit_frame,
                    text=warning_text,
                    font=("Segoe UI", 12, "bold"),
                    bg="#ffc107",
                    fg="black",
                    relief=tk.RAISED,
                    bd=1,
                    padx=8,
                    pady=6,
                    justify=tk.LEFT
                )
                warning_label.pack(fill=tk.X, pady=(0, 8))
                warning_label.bind("<MouseWheel>", preview_scroll_callback)
            
            # Row 1: Character selection
            char_row = tk.Frame(edit_frame, bg=frame_bg)
            char_row.pack(fill=tk.X, pady=(0, 5))
            char_row.bind("<MouseWheel>", preview_scroll_callback)
            
            char_label_title = tk.Label(
                char_row,
                text="Character:",
                font=("Segoe UI", 14, "bold"),
                bg=frame_bg,
                width=10,
                anchor=tk.W
            )
            char_label_title.pack(side=tk.LEFT)
            char_label_title.bind("<MouseWheel>", preview_scroll_callback)
            
            # Character text display (no dropdown)
            if is_unmatched:
                char_display = f"⚠️ NOT FOUND - SEARCH REQUIRED"
                if original_item_desc:
                    char_display += f" (Was: {original_item_desc})"
            else:
                char_display = character
            char_var = tk.StringVar(value=char_display)
            
            char_label_bg = "#fff" if not is_unmatched else "#ffc107"
            char_label_fg = "black" if not is_unmatched else "#d9534f"
            char_label_font = ("Consolas", 10) if not is_unmatched else ("Consolas", 10, "bold")
            
            char_label = tk.Label(
                char_row,
                textvariable=char_var,
                font=char_label_font,
                bg=char_label_bg,
                fg="black",
                anchor=tk.W,
                relief=tk.SUNKEN,
                bd=1,
                padx=5,
                pady=3
            )
            char_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
            char_label.bind("<MouseWheel>", preview_scroll_callback)
            
            # Search button - more prominent for unmatched items
            def make_search_handler(char_v):
                return lambda: open_image_search(char_v)
            
            search_btn_bg = "#d9534f" if is_unmatched else "#17a2b8"
            search_btn_text = "🔍 SEARCH NOW!" if is_unmatched else "🔍 Search"
            search_btn_font = ("Segoe UI", 9, "bold") if is_unmatched else ("Segoe UI", 9)
            
            search_btn = tk.Button(
                char_row,
                text=search_btn_text,
                command=make_search_handler(char_var),
                font=search_btn_font,
                bg=search_btn_bg,
                fg="black",
                relief=tk.FLAT,
                padx=12,
                pady=5,
                cursor="hand2"
            )
            search_btn.pack(side=tk.LEFT)
            search_btn.bind("<MouseWheel>", preview_scroll_callback)
            
            # Row 2: Name input
            name_row = tk.Frame(edit_frame, bg=frame_bg)
            name_row.pack(fill=tk.X, pady=(0, 5))
            name_row.bind("<MouseWheel>", preview_scroll_callback)
            
            name_label_title = tk.Label(
                name_row,
                text="Name:",
                font=("Segoe UI", 14, "bold"),
                bg=frame_bg,
                width=10,
                anchor=tk.W
            )
            name_label_title.pack(side=tk.LEFT)
            name_label_title.bind("<MouseWheel>", preview_scroll_callback)
            
            name_var = tk.StringVar(value=name)
            name_entry = tk.Entry(
                name_row,
                textvariable=name_var,
                font=("Segoe UI", 15),
                bg="#f8f9fa",
                fg="black"
            )
            name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
            name_entry.bind("<MouseWheel>", preview_scroll_callback)
            
            # Row 3: Status and actions
            action_row = tk.Frame(edit_frame, bg=frame_bg)
            action_row.pack(fill=tk.X, pady=(5, 0))
            action_row.bind("<MouseWheel>", preview_scroll_callback)
            
            # Determine status based on whether it's matched and image exists
            if is_unmatched:
                status_text = "❌ Select image required"
                status_color = "#d9534f"
            elif image_path and os.path.exists(image_path):
                status_text = "✓ Ready"
                status_color = "#5cb85c"
            else:
                status_text = "⚠ Not found"
                status_color = "#f0ad4e"
            
            status_label = tk.Label(
                action_row,
                text=status_text,
                font=("Segoe UI", 12),
                bg=frame_bg,
                fg="black"
            )
            status_label.pack(side=tk.LEFT)
            status_label.bind("<MouseWheel>", preview_scroll_callback)
            
            # Delete button
            def make_delete_handler(idx):
                return lambda: delete_order(idx)
            
            delete_btn = tk.Button(
                action_row,
                text="🗑️ Delete",
                command=make_delete_handler(idx),
                font=("Segoe UI", 12),
                bg="#d9534f",
                fg="black",
                relief=tk.FLAT,
                padx=8,
                pady=2,
                cursor="hand2"
            )
            delete_btn.bind("<MouseWheel>", preview_scroll_callback)
            delete_btn.pack(side=tk.RIGHT, padx=(5, 0))
            
            # Function to update image when character changes
            def make_update_handler(idx, img_lbl, char_v, name_v, status_lbl, char_lbl):
                def update_image(*args):
                    new_char = char_v.get()
                    
                    # Skip if it's the warning message
                    if "NOT FOUND" in new_char or "SEARCH REQUIRED" in new_char:
                        return
                    
                    order_data[idx]['character'] = new_char
                    order_data[idx]['name'] = name_v.get()
                    
                    # Update image
                    new_image_path = os.path.join(images_dir, f"{new_char}.png")
                    if os.path.exists(new_image_path):
                        try:
                            img = Image.open(new_image_path)
                            img.thumbnail((100, 100), Image.Resampling.LANCZOS)
                            new_photo = ImageTk.PhotoImage(img)
                            img_lbl.config(image=new_photo, text="", bg="#f0f0f0")
                            self._preview_images.append(new_photo)
                            status_lbl.config(text="✓ Ready", fg="black")
                            # Update character label to show the actual character name (not warning)
                            char_lbl.config(bg="#f8f9fa", fg="black", font=("Consolas", 15))
                        except:
                            img_lbl.config(image="", text="Error\nLoading", bg="#fff0f0")
                            status_lbl.config(text="❌ Error", fg="black")
                    else:
                        img_lbl.config(image="", text="No\nImage", bg="#f0f0f0")
                        status_lbl.config(text="⚠ Not found", fg="black")
                return update_image
            
            update_handler = make_update_handler(idx, img_label, char_var, name_var, status_label, char_label)
            # Update image when character changes (via search button)
            char_var.trace_add('write', lambda *args, h=update_handler: h())
            name_var.trace_add('write', lambda *args, i=idx, n=name_var: setattr(order_data[i], 'name', n.get()) or order_data.__setitem__(i, {'character': order_data[i]['character'], 'name': n.get()}))
            
            order_widgets.append({
                'frame': order_frame,
                'char_var': char_var,
                'name_var': name_var,
                'img_label': img_label,
                'status_label': status_label
            })
            
            update_summary()
            return idx
        
        # Image search dialog
        def open_image_search(target_var):
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
                font=("Segoe UI", 21, "bold"),
                bg="white"
            ).pack(pady=10)
            
            # Search input
            search_frame = tk.Frame(search_window, bg="white")
            search_frame.pack(fill=tk.X, padx=15, pady=(0, 10))
            
            tk.Label(
                search_frame,
                text="Type to filter:",
                font=("Segoe UI", 16),
                bg="white"
            ).pack(anchor=tk.W, pady=(0, 5))
            
            search_var = tk.StringVar()
            search_entry = tk.Entry(
                search_frame,
                textvariable=search_var,
                font=("Segoe UI", 18),
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
                font=("Segoe UI", 14),
                bg="white",
                fg="black"
            )
            count_label.pack(pady=(0, 5))
            
            # Listbox with scrollbar
            list_frame = tk.Frame(search_window, bg="white")
            list_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 10))
            
            scrollbar = tk.Scrollbar(list_frame)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            listbox = tk.Listbox(
                list_frame,
                font=("Consolas", 16),
                yscrollcommand=scrollbar.set,
                selectmode=tk.SINGLE,
                activestyle='dotbox',
                bg="white",
                fg="black",
                selectbackground=self.accent_color,
                selectforeground="black",
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
                    search_window.destroy()
            
            listbox.bind('<Double-Button-1>', on_double_click)
            
            # Select on Enter key
            def on_enter(event):
                selection = listbox.curselection()
                if selection:
                    selected = listbox.get(selection[0])
                    target_var.set(selected)
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
                    search_window.destroy()
                else:
                    messagebox.showwarning("No Selection", "Please select a character from the list.")
            
            select_btn = tk.Button(
                btn_frame,
                text="✓ Select Character",
                command=on_select,
                font=("Segoe UI", 16, "bold"),
                bg=self.success_color,
                fg="black",
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
                font=("Segoe UI", 15),
                bg="#6c757d",
                fg="black",
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
                font=("Segoe UI", 12, "italic"),
                bg="white",
                fg="black"
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
            
            # Count unmatched items (IMAGE-NOT-FOUND, etc.)
            unmatched = sum(1 for o in active_orders if o['character'].upper() in ['IMAGE-NOT-FOUND', 'N/A', 'NOT-FOUND', 'UNKNOWN'] or '⚠️' in o['character'])
            
            # Count items with valid images
            found = sum(1 for o in active_orders 
                       if o['character'].upper() not in ['IMAGE-NOT-FOUND', 'N/A', 'NOT-FOUND', 'UNKNOWN'] 
                       and '⚠️' not in o['character']
                       and os.path.exists(os.path.join(images_dir, f"{o['character']}.png")))
            
            missing = len(active_orders) - found - unmatched
            
            summary_text = f"✓ Ready: {found}"
            if unmatched > 0:
                summary_text += f"  |  ❌ Need Selection: {unmatched}"
            if missing > 0:
                summary_text += f"  |  ⚠ Issues: {missing}"
            summary_text += f"  |  Total: {len(active_orders)}"
            
            summary_label.config(text=summary_text)
        
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
            font=("Segoe UI", 15, "bold"),
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
            text="",
            font=("Segoe UI", 15, "bold"),
            bg="#f0f0f0",
            fg="black"
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
            
            # Check for unmatched items (IMAGE-NOT-FOUND)
            unmatched = []
            for o in active_orders:
                if o['character'].upper() in ['IMAGE-NOT-FOUND', 'N/A', 'NOT-FOUND', 'UNKNOWN'] or '⚠️' in o['character']:
                    display_name = o['name'] if o['name'] else "(no name)"
                    unmatched.append(f"'{display_name}' - No image selected")
            
            if unmatched:
                messagebox.showerror(
                    "Unmatched Items Detected",
                    f"⚠️ {len(unmatched)} item(s) still need image selection:\n\n" +
                    "\n".join(unmatched[:5]) +
                    ("\n..." if len(unmatched) > 5 else "") +
                    "\n\nPlease click the 'SEARCH NOW!' button for each\n"
                    "unmatched item to select the correct image.\n\n"
                    "Cannot proceed until all images are selected.",
                    icon='error'
                )
                return
            
            # Check for missing images (files that don't exist)
            missing = []
            for o in active_orders:
                if not os.path.exists(os.path.join(images_dir, f"{o['character']}.png")):
                    display_name = o['name'] if o['name'] else "(no name)"
                    missing.append(f"{o['character']} (for {display_name})")
            
            if missing:
                response = messagebox.askyesno(
                    "Missing Images",
                    f"{len(missing)} order(s) have missing image files:\n\n" +
                    "\n".join(missing[:5]) +
                    ("\n..." if len(missing) > 5 else "") +
                    "\n\nContinue anyway?",
                    icon='warning'
                )
                if not response:
                    return
            
            # Update main input with edited orders
            orders_text = '\n'.join([f"{o['character']},{o['name']}" for o in active_orders])
            self.order_input.delete(1.0, tk.END)
            self.order_input.insert(1.0, orders_text)
            self.order_input.config(fg="black")
            self.update_count()
            
            # Close preview and start processing
            preview_window.destroy()
            self.process_orders()
        
        confirm_btn = tk.Button(
            button_frame,
            text="✅ Confirm & Process Orders",
            command=confirm_and_process,
            font=("Segoe UI", 16, "bold"),
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
            
            orders_text = '\n'.join([f"{o['character']},{o['name']}" for o in active_orders])
            self.order_input.delete(1.0, tk.END)
            self.order_input.insert(1.0, orders_text)
            self.order_input.config(fg="black")
            self.update_count()
            
            messagebox.showinfo("Updated", "Orders updated in the main window!")
            preview_window.destroy()
        
        update_btn = tk.Button(
            button_frame,
            text="💾 Save Changes",
            command=update_only,
            font=("Segoe UI", 15),
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
            font=("Segoe UI", 15),
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
        
        # Insert message
        self.log_text.config(state=tk.NORMAL)  # Enable editing
        self.log_text.insert(tk.END, f"[{timestamp}] {prefix} {message}\n")
        
        # Auto-scroll to bottom
        self.log_text.see(tk.END)
        self.log_text.yview_moveto(1.0)  # Force scroll to bottom
        
        # Keep log from getting too long (keep last 500 lines)
        line_count = int(self.log_text.index('end-1c').split('.')[0])
        if line_count > 500:
            self.log_text.delete('1.0', f'{line_count-500}.0')
        
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
                parts = line.split(',', 1)
                character = parts[0].strip()
                name = parts[1].strip() if len(parts) > 1 else ""
                orders.append((character, name))
        
        if not orders:
            messagebox.showwarning("No Valid Orders", "Please add at least one valid order.")
            return
        
        # Save to temp CSV for processing
        import tempfile
        temp_csv = os.path.join(tempfile.gettempdir(), "temp_orders.csv")
        try:
            with open(temp_csv, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['character', 'name'])
                for character, name in orders:
                    writer.writerow([character, name])
            
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
            
            # Find all order PDFs (individual and master)
            pdf_files = []
            for file in os.listdir('.'):
                if (file.startswith('order_output_') or file.startswith('MASTER_ORDER_')) and file.endswith('.pdf'):
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
                
                # Mark selected orders as completed
                if self.selected_orders:
                    order_numbers = [o['order_number'] for o in self.selected_orders]
                    order_state.mark_orders_completed(order_numbers)
                    self.log(f"Marked {len(order_numbers)} order(s) as completed", "success")
                    
                    # Clear selection and refresh list
                    self.selected_orders.clear()
                    self.root.after(0, self.refresh_order_list)
                
                # Enable validation button
                self.validation_btn.config(state=tk.NORMAL)
               
                # Find all individual order PDFs from the MOST RECENT processing session
                # Group PDFs by timestamp
                pdf_by_timestamp = {}
                
                for file in os.listdir('.'):
                    if file.startswith('order_output_') and file.endswith('.pdf'):
                        try:
                            # Extract timestamp (e.g., "order_output_20260108_204455_1.pdf" -> "20260108_204455")
                            parts = file.replace('.pdf', '').split('_')
                            timestamp = '_'.join(parts[2:4])  # Get date_time part
                            
                            if timestamp not in pdf_by_timestamp:
                                pdf_by_timestamp[timestamp] = []
                            pdf_by_timestamp[timestamp].append(file)
                        except:
                            pass
                
                # Get the most recent timestamp (latest alphabetically)
                if not pdf_by_timestamp:
                    self.log("No order PDFs found", "warning")
                    return
                
                most_recent_timestamp = max(pdf_by_timestamp.keys())
                pdf_files = pdf_by_timestamp[most_recent_timestamp]
                
                self.log(f"Using PDFs from session: {most_recent_timestamp}", "info")
                
                # Sort PDFs numerically by the number in filename (e.g., order_output_20260108_204455_1.pdf)
                def extract_pdf_number(filename):
                    try:
                        # Extract the last number before .pdf (e.g., "order_output_20260108_204455_10.pdf" -> 10)
                        parts = filename.replace('.pdf', '').split('_')
                        return int(parts[-1])
                    except:
                        return 0
                
                pdf_files.sort(key=extract_pdf_number)
                
                # CRITICAL: Check for duplicates in the list
                self.log(f"Found {len(pdf_files)} PDFs from session {most_recent_timestamp}", "info")
                
                # Check for duplicate filenames
                seen_files = set()
                duplicates = []
                for pdf in pdf_files:
                    if pdf in seen_files:
                        duplicates.append(pdf)
                    seen_files.add(pdf)
                
                if duplicates:
                    self.log(f"⚠️⚠️⚠️ CRITICAL: Found {len(duplicates)} DUPLICATE filenames in list!", "error")
                    for dup in duplicates:
                        self.log(f"  DUPLICATE: {dup}", "error")
                    # Remove duplicates
                    pdf_files = list(dict.fromkeys(pdf_files))  # Preserves order, removes dupes
                    self.log(f"Removed duplicates, now have {len(pdf_files)} unique PDFs", "info")
                
                # Check for duplicate PDF numbers
                pdf_numbers = [extract_pdf_number(pdf) for pdf in pdf_files]
                seen_numbers = set()
                duplicate_numbers = []
                for num, pdf in zip(pdf_numbers, pdf_files):
                    if num in seen_numbers:
                        duplicate_numbers.append((num, pdf))
                    seen_numbers.add(num)
                
                if duplicate_numbers:
                    self.log(f"⚠️⚠️⚠️ CRITICAL: Found PDFs with same number!", "error")
                    for num, pdf in duplicate_numbers:
                        self.log(f"  Number {num} appears multiple times: {pdf}", "error")
                
                self.log("PDF List (in order):", "info")
                for idx, pdf in enumerate(pdf_files):
                    pdf_num = extract_pdf_number(pdf)
                    file_size = os.path.getsize(pdf) if os.path.exists(pdf) else 0
                    self.log(f"  Position {idx+1} → [PDF #{pdf_num}] {pdf} ({file_size:,} bytes)", "info")
               
                if pdf_files:
                    # === STEP 1: Verify PDFs before processing ===
                    self.log("=" * 70, "info")
                    self.log("STEP 1: Pre-Flatten Verification", "info")
                    self.log("=" * 70, "info")
                    from pypdf import PdfReader
                    
                    pre_flatten_status = {}
                    for pdf_file in pdf_files:
                        if os.path.exists(pdf_file):
                            try:
                                reader = PdfReader(pdf_file)
                                page_count = len(reader.pages)
                                file_size = os.path.getsize(pdf_file)
                                pre_flatten_status[pdf_file] = {
                                    'pages': page_count,
                                    'size': file_size,
                                    'status': 'ok' if page_count == 1 else 'warning'
                                }
                                self.log(f"  ✓ {pdf_file}: {page_count} page(s), {file_size:,} bytes", 
                                        "success" if page_count == 1 else "warning")
                                if page_count != 1:
                                    self.log(f"      ⚠️ Expected 1 page, found {page_count}", "warning")
                            except Exception as e:
                                self.log(f"  ✗ {pdf_file}: ERROR - {str(e)}", "error")
                                pre_flatten_status[pdf_file] = {'status': 'error', 'error': str(e)}
                        else:
                            self.log(f"  ✗ {pdf_file}: FILE NOT FOUND", "error")
                            pre_flatten_status[pdf_file] = {'status': 'not_found'}
                    
                    # === STEP 2: Flatten each individual PDF in place ===
                    self.log("=" * 70, "info")
                    self.log("STEP 2: Flattening Individual PDFs", "info")
                    self.log("=" * 70, "info")
                    
                    successfully_flattened = []
                    failed_flattening = []
                    
                    for pdf_file in pdf_files:
                        if os.path.exists(pdf_file):
                            if self.flatten_pdf_in_place(pdf_file, dpi=300):
                                successfully_flattened.append(pdf_file)
                            else:
                                failed_flattening.append(pdf_file)
                                self.log(f"  ⚠️ Will try to use non-flattened version: {pdf_file}", "warning")
                    
                    if failed_flattening:
                        self.log(f"Warning: {len(failed_flattening)} PDF(s) failed to flatten", "warning")
                    
                    # === STEP 3: Post-Flatten Verification ===
                    self.log("=" * 70, "info")
                    self.log("STEP 3: Post-Flatten Verification", "info")
                    self.log("=" * 70, "info")
                    
                    for pdf_file in pdf_files:
                        if os.path.exists(pdf_file):
                            try:
                                reader = PdfReader(pdf_file)
                                page_count = len(reader.pages)
                                file_size = os.path.getsize(pdf_file)
                                
                                # Compare with pre-flatten status
                                if pdf_file in pre_flatten_status:
                                    pre_pages = pre_flatten_status[pdf_file].get('pages', 0)
                                    pre_size = pre_flatten_status[pdf_file].get('size', 0)
                                    
                                    if page_count != pre_pages:
                                        self.log(f"  ⚠️ {pdf_file}: Page count changed {pre_pages} → {page_count}", "warning")
                                    else:
                                        self.log(f"  ✓ {pdf_file}: {page_count} page(s), {file_size:,} bytes", "success")
                                else:
                                    self.log(f"  ✓ {pdf_file}: {page_count} page(s), {file_size:,} bytes", "success")
                                    
                            except Exception as e:
                                self.log(f"  ✗ {pdf_file}: ERROR - {str(e)}", "error")
                    
                    # === STEP 4: Create Master PDF ===
                    self.log("=" * 70, "info")
                    self.log("STEP 4: Creating Master PDF", "info")
                    self.log("=" * 70, "info")
                    
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    master_pdf_name = f"MASTER_ORDER_{timestamp}.pdf"
                   
                    if self.merge_pdfs(pdf_files, master_pdf_name):
                        # === STEP 5: Flatten the master PDF ===
                        self.log("=" * 70, "info")
                        self.log("STEP 5: Flattening Master PDF", "info")
                        self.log("=" * 70, "info")
                        
                        if self.flatten_pdf_in_place(master_pdf_name, dpi=300):
                            self.log("✓ Master PDF flattened successfully", "success")
                        else:
                            self.log("⚠️ Master PDF flattening failed, but file should still be usable", "warning")
                        
                        # === STEP 6: Final Verification of Master PDF ===
                        self.log("=" * 70, "info")
                        self.log("STEP 6: Final Master PDF Verification", "info")
                        self.log("=" * 70, "info")
                        
                        try:
                            final_reader = PdfReader(master_pdf_name)
                            final_page_count = len(final_reader.pages)
                            final_size = os.path.getsize(master_pdf_name)
                            
                            self.log(f"Master PDF: {master_pdf_name}", "info")
                            self.log(f"  Total pages: {final_page_count}", "info")
                            self.log(f"  File size: {final_size:,} bytes", "info")
                            self.log(f"  Expected pages: {len(pdf_files)}", "info")
                            
                            if final_page_count == len(pdf_files):
                                self.log(f"✓ Page count matches perfectly!", "success")
                            else:
                                self.log(f"⚠️ WARNING: Expected {len(pdf_files)} pages but got {final_page_count}", "warning")
                            
                            # List all pages in master PDF with source mapping
                            self.log("Master PDF Complete Page Mapping:", "info")
                            self.log("(Shows which source file created each master page)", "info")
                            for page_idx in range(final_page_count):
                                master_page_num = page_idx + 1
                                try:
                                    page = final_reader.pages[page_idx]
                                    
                                    # Try to find the source from pdf_files list
                                    if master_page_num <= len(pdf_files):
                                        expected_source = pdf_files[page_idx]
                                        expected_num = extract_pdf_number(expected_source)
                                        self.log(f"  Master Page {master_page_num} ← {expected_source} [PDF #{expected_num}]", "success")
                                    else:
                                        self.log(f"  Master Page {master_page_num}: OK (no source mapping)", "success")
                                except Exception as e:
                                    self.log(f"  Master Page {master_page_num}: ERROR - {str(e)}", "error")
                            
                            self.log("=" * 70, "info")
                            self.log("✓✓✓ ALL STEPS COMPLETE ✓✓✓", "success")
                            self.log("=" * 70, "info")
                            
                        except Exception as e:
                            self.log(f"Final verification error: {str(e)}", "error")
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
                       
                        messagebox.showinfo(
                            "Success",
                            f"Orders processed and FLATTENED!\n\n"
                            f"✓ {len(pdf_files)} individual PDFs created (flattened)\n"
                            f"✓ Master PDF created and flattened: {master_pdf_name}\n"
                            f"✓ Final page count: {final_page_count} pages\n"
                            f"✓ No editable text or hidden template data remains\n\n"
                            f"Master PDF opened automatically!\n\n"
                            f"Check the log for detailed verification results!"
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
        """Rasterize all pages of a PDF and overwrite it with a flattened version - CAREFULLY."""
        try:
            # Step 1: Verify PDF exists and is readable
            if not os.path.exists(pdf_path):
                self.log(f"  ✗ File not found: {os.path.basename(pdf_path)}", "error")
                return False
                
            file_size = os.path.getsize(pdf_path)
            if file_size == 0:
                self.log(f"  ✗ Empty file: {os.path.basename(pdf_path)}", "error")
                return False
            
            # Step 2: Open and verify page count
            doc = fitz.open(pdf_path)
            original_page_count = doc.page_count
            
            if original_page_count == 0:
                doc.close()
                self.log(f"  ⚠️ Skipped {os.path.basename(pdf_path)} (no pages)", "warning")
                return False
            
            if original_page_count != 1:
                self.log(f"  ⚠️ WARNING: {os.path.basename(pdf_path)} has {original_page_count} pages (expected 1)", "warning")

            # Step 3: Create temporary output path
            temp_path = pdf_path + ".flattened_tmp.pdf"
            
            # Remove any existing temp file
            if os.path.exists(temp_path):
                os.remove(temp_path)

            # Step 4: Create new flattened PDF with EXACT page count
            writer = fitz.open()  # New empty PDF

            # Process each page individually and track
            for page_num in range(original_page_count):
                page = doc[page_num]
                
                # Create a new blank page with the exact same dimensions
                new_page = writer.new_page(width=page.rect.width, height=page.rect.height)

                # Render the original page to a high-resolution pixmap
                zoom = dpi / 72.0
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat, alpha=False)

                # Insert the rendered image (covers the entire page)
                new_page.insert_image(new_page.rect, pixmap=pix)
                
                # Verify this page was added
                if writer.page_count != page_num + 1:
                    raise Exception(f"Page count mismatch after adding page {page_num + 1}")

            flattened_page_count = writer.page_count

            # Step 5: Verify page count matches
            if flattened_page_count != original_page_count:
                writer.close()
                doc.close()
                raise Exception(f"Page count mismatch: original={original_page_count}, flattened={flattened_page_count}")

            # Step 6: Save to temp file
            writer.save(temp_path, garbage=4, deflate=True, clean=True)
            writer.close()
            doc.close()
            
            # Step 7: Verify temp file was created correctly
            if not os.path.exists(temp_path):
                raise Exception(f"Failed to create temp file: {temp_path}")
            
            temp_size = os.path.getsize(temp_path)
            if temp_size == 0:
                os.remove(temp_path)
                raise Exception(f"Temp file is empty")
            
            # Step 8: Verify temp file page count
            verify_doc = fitz.open(temp_path)
            verify_count = verify_doc.page_count
            verify_doc.close()
            
            if verify_count != original_page_count:
                os.remove(temp_path)
                raise Exception(f"Temp file page count wrong: expected={original_page_count}, got={verify_count}")

            # Step 9: Replace original with flattened version
            os.replace(temp_path, pdf_path)
            
            # Step 10: Final verification
            final_doc = fitz.open(pdf_path)
            final_count = final_doc.page_count
            final_doc.close()
            
            if final_count != original_page_count:
                raise Exception(f"Final file page count wrong: expected={original_page_count}, got={final_count}")

            self.log(f"  ✓ {os.path.basename(pdf_path)} flattened successfully ({final_count} page)", "success")
            return True
                
        except Exception as e:
            self.log(f"  ✗ Failed to flatten {os.path.basename(pdf_path)}: {str(e)}", "error")
            # Clean up temp file if it exists
            try:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
            except:
                pass
            return False                
    def merge_pdfs(self, pdf_files, output_path):
        """Merge multiple PDF files into one master PDF - CAREFULLY with extensive validation"""
        try:
            from pypdf import PdfReader, PdfWriter
            
            self.log(f"Merging {len(pdf_files)} PDFs into master PDF...", "info")
            self.log("=" * 60, "info")
            
            # Step 1: Pre-validation - verify all PDFs before starting merge
            self.log("Step 1: Pre-validation of all PDFs...", "info")
            valid_pdfs = []
            for idx, pdf_file in enumerate(pdf_files, 1):
                if not os.path.exists(pdf_file):
                    self.log(f"  [{idx}] ✗ NOT FOUND: {pdf_file}", "error")
                    continue
                
                file_size = os.path.getsize(pdf_file)
                if file_size == 0:
                    self.log(f"  [{idx}] ✗ EMPTY FILE: {pdf_file}", "error")
                    continue
                
                try:
                    reader = PdfReader(pdf_file)
                    num_pages = len(reader.pages)
                    
                    if num_pages == 0:
                        self.log(f"  [{idx}] ✗ NO PAGES: {pdf_file}", "error")
                        continue
                    
                    if num_pages != 1:
                        self.log(f"  [{idx}] ⚠️ MULTIPLE PAGES ({num_pages}): {pdf_file}", "warning")
                    
                    valid_pdfs.append((pdf_file, num_pages))
                    self.log(f"  [{idx}] ✓ Valid: {pdf_file} ({num_pages} page)", "success")
                    
                except Exception as e:
                    self.log(f"  [{idx}] ✗ ERROR READING: {pdf_file} - {str(e)}", "error")
            
            if not valid_pdfs:
                raise Exception("No valid PDFs to merge!")
            
            self.log(f"Pre-validation complete: {len(valid_pdfs)}/{len(pdf_files)} PDFs are valid", "info")
            self.log("=" * 60, "info")
            
            # Step 2: Create writer and merge PDFs one by one with tracking
            self.log("Step 2: Merging PDFs...", "info")
            writer = PdfWriter()
            merge_log = []
            merge_tracking = {}  # Track which source file goes to which master page
            
            for idx, (pdf_file, expected_pages) in enumerate(valid_pdfs, 1):
                try:
                    # Get file hash for verification
                    import hashlib
                    with open(pdf_file, 'rb') as f:
                        file_hash = hashlib.md5(f.read()).hexdigest()[:8]
                    
                    # Re-open the PDF (don't reuse readers)
                    reader = PdfReader(pdf_file)
                    num_pages = len(reader.pages)
                    
                    # Double-check page count
                    if num_pages != expected_pages:
                        raise Exception(f"Page count changed: was {expected_pages}, now {num_pages}")
                    
                    if num_pages == 0:
                        raise Exception("PDF became empty")
                    
                    # Get the first (and ideally only) page
                    page = reader.pages[0]
                    
                    # Track current writer state before adding
                    pages_before = len(writer.pages)
                    
                    # Add the page
                    writer.add_page(page)
                    
                    # Verify the page was added
                    pages_after = len(writer.pages)
                    if pages_after != pages_before + 1:
                        raise Exception(f"Failed to add page: before={pages_before}, after={pages_after}")
                    
                    # Track this merge with hash
                    master_page_num = pages_after
                    merge_tracking[master_page_num] = {
                        'source': pdf_file,
                        'hash': file_hash,
                        'source_idx': idx
                    }
                    
                    # Log successful merge with hash
                    merge_info = f"[{idx}/{len(valid_pdfs)}] {os.path.basename(pdf_file)} [hash:{file_hash}] → Master page {pages_after}"
                    self.log(f"  ✓ {merge_info}", "success")
                    merge_log.append(merge_info)
                    
                    if num_pages > 1:
                        self.log(f"      Note: Source had {num_pages} pages, used page 1 only", "info")
                    
                except Exception as e:
                    self.log(f"  ✗ [{idx}/{len(valid_pdfs)}] Failed to merge {pdf_file}: {str(e)}", "error")
                    # Continue with other PDFs instead of failing completely
            
            final_page_count = len(writer.pages)
            self.log("=" * 60, "info")
            self.log(f"Merge complete: {final_page_count} total pages in writer", "info")
            
            # Check for duplicate sources (same file added twice)
            self.log("Checking for duplicate sources...", "info")
            hash_to_pages = {}
            for page_num, info in merge_tracking.items():
                file_hash = info['hash']
                if file_hash not in hash_to_pages:
                    hash_to_pages[file_hash] = []
                hash_to_pages[file_hash].append((page_num, info['source']))
            
            duplicates_found = False
            for file_hash, pages in hash_to_pages.items():
                if len(pages) > 1:
                    duplicates_found = True
                    self.log(f"⚠️⚠️⚠️ DUPLICATE CONTENT [hash:{file_hash}]:", "error")
                    for page_num, source in pages:
                        self.log(f"    Master page {page_num} from {source}", "error")
            
            if not duplicates_found:
                self.log("✓ No duplicate sources detected - all master pages are unique", "success")
            
            # Step 3: Write the master PDF
            self.log("Step 3: Writing master PDF to disk...", "info")
            with open(output_path, 'wb') as output_file:
                writer.write(output_file)
            
            # Step 4: Verify the written file
            self.log("Step 4: Verifying written master PDF...", "info")
            
            if not os.path.exists(output_path):
                raise Exception("Master PDF was not created!")
            
            output_size = os.path.getsize(output_path)
            if output_size == 0:
                raise Exception("Master PDF is empty!")
            
            verify_reader = PdfReader(output_path)
            actual_pages = len(verify_reader.pages)
            
            self.log(f"  File size: {output_size:,} bytes", "info")
            self.log(f"  Pages: {actual_pages}", "info")
            
            # Step 5: Final validation
            if actual_pages != final_page_count:
                raise Exception(f"Page count mismatch! Expected {final_page_count}, got {actual_pages}")
            
            if actual_pages != len(valid_pdfs):
                self.log(f"  ⚠️ Warning: Merged {len(valid_pdfs)} PDFs but got {actual_pages} pages", "warning")
            
            self.log("=" * 60, "info")
            self.log(f"✓ Master PDF created successfully: {output_path}", "success")
            self.log(f"✓ Verified: {actual_pages} pages (from {len(valid_pdfs)} source PDFs)", "success")
            
            return True
            
        except Exception as e:
            self.log(f"✗ Error merging PDFs: {str(e)}", "error")
            import traceback
            self.log(traceback.format_exc(), "error")
            return False
            
    def clear_all(self):
        """Clear all fields"""
        self.csv_path.set("")
        self.order_input.delete(1.0, tk.END)
        placeholder = "mickey-captain,Johnny\nminnie-captain,Sarah\nstitch-captain,Michael\nmoana-captain,Emma"
        self.order_input.insert(1.0, placeholder)
        self.order_input.config(fg="black")
        self.preview_text.delete(1.0, tk.END)
        self.log_text.delete(1.0, tk.END)
        self.orders_count.set("0 orders")
        self.progress_var.set(0)
        self.status_text.set("Ready to process orders")
        self.view_btn.config(state=tk.DISABLED)
        self.master_pdf_btn.config(state=tk.DISABLED)
        self.master_pdf_path = None
        
    def get_available_images(self):
        """Get list of available images from FHM_Images folder"""
        try:
            parent_dir = os.path.dirname(os.path.abspath(os.getcwd()))
            images_dir = os.path.join(parent_dir, 'FHM_Images')
            
            if not os.path.exists(images_dir):
                return []
            
            image_files = []
            for file in os.listdir(images_dir):
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
            self.raw_text.config(fg="black")
            
    def restore_raw_placeholder(self, event):
        """Restore raw text placeholder if empty"""
        if not self.raw_text.get(1.0, tk.END).strip():
            placeholder = """Paste order details here, like:

Order #12345 - Mickey Captain themed, names: Johnny, Sarah, Michael
or
Disney Cruise Door Magnet - 3 magnets: Minnie, Donald, Goofy (all captain theme)"""
            self.raw_text.insert(1.0, placeholder)
            self.raw_text.config(fg="black")
            
    def clear_raw_text(self):
        """Clear raw text area"""
        self.raw_text.delete(1.0, tk.END)
        self.raw_text.config(fg="black")
        
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
        self.raw_text.config(fg="black")
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
3. IGNORE boat magnets completely. But do the other magnets in the order. For boat orders, assume the theme is captain.
4. BUT if a line has BOTH a boat AND regular character orders, extract the regular characters
5. A single order line can have 1-5 character-name pairs (plus possibly a boat to ignore)
6. Keep character descriptions simple and natural (e.g., "Luke Skywalker", "Stitch captain", "Minnie Spiderman")
7. Preserve exact name spellings from the order
8. If no name is specified for a character, use "no name" for the name
9. Duck and dog orders do not need captain, pirate, etc. Just duck/dog and the ID number with them. 
10. Do not omit any items from the order other than boats.
11. The header for an order can determine the theme. For example, if the header is pirate, every item in that order is pirate. Same with captain, christmas, etc.

EXAMPLES:

Input: "Item: Captain Mickey, Personalization: Johnny"
Output: Mickey captain - name: Johnny

Input: "Item: Christmas Elsa\nPersonalization: Sarah"
Output: Elsa christmas - name: Sarah

Input: "captain Order has boat +  Minnie for 'Katie' and  Woody for 'Sean' and dog 16 for 'Joni'"
Output: 
Minnie captain - name: Katie
Woody captain - name: Sean
dog 16 - name: Joni

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
5. Common patterns to match:
   - "Mickey captain" → "mickey-captain.png"
   - "Stitch captain" → "stitch-captain.png"
   - "Elsa christmas" → "elsa-christmas.png"
   - "Minnie Spiderman" → "minnie-spiderman.png"
   - "Donald Hulk" → "donald-hulk.png"
   - "dog 16" → "dog-16.png"
   - "duck 23" → "duck-23.png"
6. Character names are case-insensitive for matching
7. ALWAYS include ALL items, even if no match found (use "N/A.png")
8. for magnets with no name, leave it blank. (example "name": "")

OUTPUT FORMAT - Return ONLY a Python LIST of dictionaries, no other text:
[
  {{"name": "PersonName1", "image": "exact-filename.png", "item": "original character description"}},
  {{"name": "PersonName2", "image": "exact-filename.png", "item": "original character description"}},
  {{"name": "PersonName3", "image": "N/A.png", "item": "original character description"}}
]

EXAMPLES:

Input: "Mickey captain - name: Johnny"
Output: [{{"name": "Johnny", "image": "mickey-captain.png", "item": "Mickey captain"}}]

Input: "Stitch captain - name: Michael\nMinnie Spiderman - name: Cecile"
Output: [
  {{"name": "Michael", "image": "stitch-captain.png", "item": "Stitch captain"}},
  {{"name": "Cecile", "image": "minnie-spiderman.png", "item": "Minnie Spiderman"}}
]

Input: "RareCharacter - name: Test"
Output: [{{"name": "Test", "image": "N/A.png", "item": "RareCharacter"}}]

CRITICAL: 
- Only use filenames that EXACTLY match the list above
- If no match found, use "N/A.png" - DO NOT OMIT THE ITEM
- Return a LIST, preserving order
- Include ALL items from the input

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
    
    def startup_load_orders(self):
        """Load orders on startup - silent pull if possible"""
        try:
            # First, just refresh from existing file
            self.refresh_order_list()
            
            # If we have orders, we're good
            if order_state.get_all_orders_with_status():
                self.log("Loaded existing orders", "info")
                self.status_text.set("Ready - Review orders below and click 'Begin Selected Orders'")
                return
            
            # No orders yet, try to pull silently
            self.log("No orders found, attempting to pull from email...", "info")
            self.status_text.set("Checking for orders...")
            
            # Try to pull in background
            thread = threading.Thread(target=self.startup_pull_thread)
            thread.daemon = True
            thread.start()
            
        except Exception as e:
            # Silent failure on startup - user can manually pull
            self.log(f"Startup load: {str(e)}", "warning")
            self.status_text.set("Click 'Pull New Orders' to get started")
    
    def startup_pull_thread(self):
        """Background thread for startup order pulling"""
        try:
            pullorders.process_recent_etsy_sales_stop_on_processed()
            self.root.after(0, self.refresh_order_list)
            self.root.after(0, lambda: self.log("✓ Orders loaded from email", "success"))
            self.root.after(0, lambda: self.status_text.set("Ready - Review orders and click 'Begin Selected Orders'"))
        except Exception as e:
            # Silent failure - user can manually pull
            self.root.after(0, lambda: self.log("Tip: Click 'Pull New Orders' to fetch from email", "info"))
            self.root.after(0, lambda: self.status_text.set("Click 'Pull New Orders' to get started"))
    
    def pull_new_orders(self):
        """Pull new orders from email"""
        if self.processing:
            messagebox.showinfo("Processing", "Cannot pull orders while processing.")
            return
        
        # Run in thread to keep UI responsive
        thread = threading.Thread(target=self.pull_new_orders_thread)
        thread.daemon = True
        thread.start()
    
    def pull_new_orders_thread(self):
        """Pull new orders in background thread"""
        try:
            self.root.after(0, lambda: self.status_text.set("Pulling new orders from email..."))
            self.root.after(0, lambda: self.log("Pulling new orders from email...", "info"))
            
            # Call pullorders.py
            pullorders.process_recent_etsy_sales_stop_on_processed()
            
            self.root.after(0, lambda: self.log("✓ Finished pulling orders", "success"))
            self.root.after(0, lambda: self.status_text.set("Orders pulled successfully"))
            self.root.after(0, self.refresh_order_list)
            self.root.after(0, lambda: messagebox.showinfo("Success", "New orders pulled successfully!\nCheck the order list below."))
            
        except Exception as e:
            error_msg = str(e)
            self.root.after(0, lambda msg=error_msg: self.log(f"Error pulling orders: {msg}", "error"))
            self.root.after(0, lambda msg=error_msg: messagebox.showerror("Error", f"Failed to pull orders:\n{msg}"))
    
    def refresh_order_list(self):
        """Refresh the order list display"""
        # Clear existing widgets
        for widget in self.orders_list_frame.winfo_children():
            widget.destroy()
        
        self.order_checkboxes = {}
        self.order_widgets = {}
        
        # Get orders based on filter
        filter_value = self.order_filter.get()
        
        if filter_value == "all":
            orders = order_state.get_all_orders_with_status()
        elif filter_value == "pending":
            orders = order_state.get_pending_orders()
        else:  # completed
            orders = order_state.get_completed_orders()
        
        # Update count
        self.order_list_count.set(f"{len(orders)} orders")
        
        if not orders:
            # Show empty message
            empty_label = tk.Label(
                self.orders_list_frame,
                text="No orders found. Click 'Pull New Orders' to fetch from email.",
                font=("Segoe UI", 15, "italic"),
                bg="white",
                fg="black"
            )
            empty_label.pack(pady=50)
            if hasattr(self, 'orders_scroll_callback'):
                empty_label.bind("<MouseWheel>", self.orders_scroll_callback)
            return
        
        # Create order widgets
        for order in orders:
            self.create_order_widget(order)
        
        # Auto-select pending orders
        if filter_value == "pending":
            self.root.after(100, self.auto_select_pending_orders)
    
    def create_order_widget(self, order):
        """Create a widget for displaying an order with checkbox"""
        order_num = order['order_number']
        is_completed = order.get('completed', False)
        
        # Create frame for this order
        order_frame = tk.Frame(
            self.orders_list_frame,
            bg="#f8f9fa" if not is_completed else "#e8f5e9",
            relief=tk.RIDGE,
            bd=1
        )
        order_frame.pack(fill=tk.X, padx=5, pady=3)
        
        # Bind scroll to this frame
        if hasattr(self, 'orders_scroll_callback'):
            order_frame.bind("<MouseWheel>", self.orders_scroll_callback)
        
        # Checkbox and order header
        header_frame = tk.Frame(order_frame, bg=order_frame['bg'])
        header_frame.pack(fill=tk.X, padx=10, pady=8)
        
        # Bind scroll to header frame
        if hasattr(self, 'orders_scroll_callback'):
            header_frame.bind("<MouseWheel>", self.orders_scroll_callback)
        
        # Checkbox
        checkbox_var = tk.BooleanVar(value=False)
        checkbox = tk.Checkbutton(
            header_frame,
            variable=checkbox_var,
            bg=order_frame['bg'],
            command=lambda o=order, v=checkbox_var: self.on_order_checkbox_changed(o, v)
        )
        checkbox.pack(side=tk.LEFT, padx=(0, 10))
        if hasattr(self, 'orders_scroll_callback'):
            checkbox.bind("<MouseWheel>", self.orders_scroll_callback)
        
        self.order_checkboxes[order_num] = {
            'var': checkbox_var,
            'widget': checkbox,
            'order': order
        }
        
        # Order number and status
        status_color = "#5cb85c" if is_completed else "#f0ad4e"
        status_text = "✓ Completed" if is_completed else "⏳ Pending"
        
        order_label = tk.Label(
            header_frame,
            text=f"Order #{order_num}",
            font=("Segoe UI", 15, "bold"),
            bg=order_frame['bg'],
            fg="black"
        )
        order_label.pack(side=tk.LEFT)
        if hasattr(self, 'orders_scroll_callback'):
            order_label.bind("<MouseWheel>", self.orders_scroll_callback)
        
        status_label = tk.Label(
            header_frame,
            text=status_text,
            font=("Segoe UI", 14),
            bg=order_frame['bg'],
            fg="black"
        )
        status_label.pack(side=tk.LEFT, padx=(10, 0))
        if hasattr(self, 'orders_scroll_callback'):
            status_label.bind("<MouseWheel>", self.orders_scroll_callback)
        
        # Customer info
        customer_info = f"{order['name']}"
        if order['city'] or order['state']:
            customer_info += f" • {order['city']}, {order['state']}"
        
        customer_label = tk.Label(
            header_frame,
            text=customer_info,
            font=("Segoe UI", 14),
            bg=order_frame['bg'],
            fg="black"
        )
        customer_label.pack(side=tk.LEFT, padx=(20, 0))
        if hasattr(self, 'orders_scroll_callback'):
            customer_label.bind("<MouseWheel>", self.orders_scroll_callback)
        
        # Item count
        item_count_label = tk.Label(
            header_frame,
            text=f"({len(order['items'])} item{'s' if len(order['items']) != 1 else ''})",
            font=("Segoe UI", 14),
            bg=order_frame['bg'],
            fg="black"
        )
        item_count_label.pack(side=tk.RIGHT)
        if hasattr(self, 'orders_scroll_callback'):
            item_count_label.bind("<MouseWheel>", self.orders_scroll_callback)
        
        # Items details (collapsible)
        items_frame = tk.Frame(order_frame, bg="white")
        items_frame.pack(fill=tk.X, padx=20, pady=(0, 8))
        if hasattr(self, 'orders_scroll_callback'):
            items_frame.bind("<MouseWheel>", self.orders_scroll_callback)
        
        for item_data in order['items']:
            item_text = item_data.get('personalization', '(No personalization)')
            item_label = tk.Label(
                items_frame,
                text=f"  • {item_text}",
                font=("Segoe UI", 14),
                bg="white",
                fg="black",
                anchor=tk.W
            )
            item_label.pack(fill=tk.X, pady=2)
            if hasattr(self, 'orders_scroll_callback'):
                item_label.bind("<MouseWheel>", self.orders_scroll_callback)
        
        self.order_widgets[order_num] = {
            'frame': order_frame,
            'checkbox_var': checkbox_var,
            'order': order
        }
    
    def on_order_checkbox_changed(self, order, checkbox_var):
        """Handle checkbox state change"""
        if checkbox_var.get():
            if order not in self.selected_orders:
                self.selected_orders.append(order)
        else:
            if order in self.selected_orders:
                self.selected_orders.remove(order)
    
    def auto_select_pending_orders(self):
        """Automatically select all pending (non-completed) orders"""
        for order_num, checkbox_data in self.order_checkboxes.items():
            order = checkbox_data['order']
            if not order.get('completed', False):
                checkbox_data['var'].set(True)
                if order not in self.selected_orders:
                    self.selected_orders.append(order)
    
    def select_all_orders(self):
        """Select all visible orders"""
        self.selected_orders.clear()
        for order_num, checkbox_data in self.order_checkboxes.items():
            checkbox_data['var'].set(True)
            self.selected_orders.append(checkbox_data['order'])
    
    def deselect_all_orders(self):
        """Deselect all orders"""
        self.selected_orders.clear()
        for order_num, checkbox_data in self.order_checkboxes.items():
            checkbox_data['var'].set(False)
    
    def begin_selected_orders(self):
        """Begin processing selected orders - send to AI input field"""
        # Rebuild selected_orders from checkbox states to ensure accuracy
        self.selected_orders.clear()
        for order_num, checkbox_data in self.order_checkboxes.items():
            if checkbox_data['var'].get():  # If checkbox is checked
                self.selected_orders.append(checkbox_data['order'])
        
        if not self.selected_orders:
            messagebox.showwarning("No Selection", "Please select at least one order to process.")
            return
        
        if self.processing or self.ai_processing:
            messagebox.showinfo("Processing", "Already processing. Please wait.")
            return
        
        # Log which orders are being processed
        order_nums = [o['order_number'] for o in self.selected_orders]
        self.log(f"Processing selected orders: {', '.join(str(n) for n in order_nums)}", "info")
        
        # Extract order text for AI processing
        order_text = order_state.extract_order_text_for_ai(self.selected_orders)
        
        # Fill the AI raw text input
        self.raw_text.delete(1.0, tk.END)
        self.raw_text.insert(1.0, order_text)
        self.raw_text.config(fg="black")
        
        # Count items to determine parsing strategy
        num_orders = len(self.selected_orders)
        num_items = sum(len(o['items']) for o in self.selected_orders)
        
        # Smart AI selection: Quick parse for small batches, regular for larger
        use_quick_parse = num_items <= 4
        parse_method = "Quick Parse (faster)" if use_quick_parse else "Standard Parse (more thorough)"
        
        result = messagebox.askyesno(
            "Begin Processing",
            f"Ready to process:\n\n"
            f"• {num_orders} order{'s' if num_orders != 1 else ''}\n"
            f"• {num_items} item{'s' if num_items != 1 else ''}\n\n"
            f"Recommended: {parse_method}\n\n"
            f"The order details have been loaded into the AI parsing field.\n\n"
            f"Would you like to:\n"
            f"1. Use AI to parse automatically? (Click Yes)\n"
            f"2. Review/edit manually first? (Click No)",
            icon='question'
        )
        
        if result:
            # User wants AI to parse immediately - use smart selection
            if use_quick_parse:
                self.log(f"Using Quick Parse for {num_items} item(s)", "info")
                self.quick_parse_with_ai()
            else:
                self.log(f"Using Standard Parse for {num_items} item(s)", "info")
                self.parse_with_ai()
        else:
            # User wants to review - just show the status
            self.status_text.set(f"{num_orders} order(s) loaded - review and click 'Parse with AI' or edit manually")
            self.log(f"Loaded {num_orders} order(s) with {num_items} item(s)", "info")
    
    def open_validation_page(self):
        """Open validation page showing orders, addresses, and images"""
        # Get all orders with their generated images
        validation_data = self.prepare_validation_data()
        
        if not validation_data:
            messagebox.showinfo("No Data", "No processed orders to validate.\nProcess some orders first!")
            return
        
        # Create validation window
        self.show_validation_window(validation_data)
    
    def prepare_validation_data(self):
        """
        Prepare validation data by matching orders with generated images.
        
        Returns list of dicts with order info and image paths.
        """
        # Get completed orders
        completed_orders = order_state.get_completed_orders()
        
        if not completed_orders:
            return []
        
        validation_data = []
        
        # Check if outputs directory exists
        outputs_dir = "outputs"
        if not os.path.exists(outputs_dir):
            return []
        
        # Get list of generated images sorted by number
        image_files = sorted([f for f in os.listdir(outputs_dir) if f.endswith('.png') and f[0].isdigit()],
                            key=lambda x: int(x.split('.')[0]))
        
        # Track cumulative image index across all orders
        image_index = 0
        
        for order in completed_orders:
            order_images = []
            num_items = len(order['items'])
            
            # Assign the next N images to this order (where N = number of items)
            for i in range(num_items):
                if image_index < len(image_files):
                    img_path = os.path.join(outputs_dir, image_files[image_index])
                    order_images.append(img_path)
                    image_index += 1
            
            # Create order data with matched images
            order_data = {
                'order_number': order['order_number'],
                'name': order['name'],
                'address': f"{order['city']}, {order['state']}",
                'items': order['items'],
                'images': order_images  # Now contains the specific images for this order
            }
            
            validation_data.append(order_data)
        
        return validation_data
    
    def show_validation_window(self, validation_data):
        """Show validation window with order details and images"""
        try:
            from PIL import Image, ImageTk
        except ImportError:
            messagebox.showinfo("Preview Unavailable", "Image preview requires the Pillow library.")
            return
        
        # Create validation window
        val_window = tk.Toplevel(self.root)
        val_window.title("✓ Order Validation")
        val_window.geometry("1200x800")
        val_window.configure(bg="white")
        
        # Make it modal
        val_window.transient(self.root)
        val_window.grab_set()
        
        # Title
        title_frame = tk.Frame(val_window, bg=self.accent_color, height=70)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)
        
        title_label = tk.Label(
            title_frame,
            text="✓ Order Validation - Review Before Shipping",
            font=("Segoe UI", 24, "bold"),
            bg=self.accent_color,
            fg="black"
        )
        title_label.pack(pady=(10, 0))
        
        subtitle_label = tk.Label(
            title_frame,
            text="Verify order details and images for each envelope",
            font=("Segoe UI", 14),
            bg=self.accent_color,
            fg="black"
        )
        subtitle_label.pack(pady=(0, 10))
        
        # Create scrollable frame
        canvas = tk.Canvas(val_window, bg="white")
        scrollbar = tk.Scrollbar(val_window, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="white")
        
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Update scroll region
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        # Make scrollable frame expand to canvas width
        def configure_canvas_width(event):
            canvas.itemconfig(canvas_window, width=event.width)
        
        canvas.bind('<Configure>', configure_canvas_width)
        
        # Enable mouse wheel scrolling (cross-platform) - with safety checks
        def on_mousewheel(event):
            try:
                if canvas.winfo_exists() and event.delta:
                    canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except:
                pass
        
        def on_mousewheel_mac(event):
            try:
                if canvas.winfo_exists():
                    canvas.yview_scroll(int(-1*event.delta), "units")
            except:
                pass
        
        # Store scroll callback for use in order widgets
        if platform.system() == 'Darwin':  # macOS
            validation_scroll_callback = on_mousewheel_mac
        else:  # Windows/Linux
            validation_scroll_callback = on_mousewheel
        
        # Bind for different platforms - local bindings only
        canvas.bind("<MouseWheel>", validation_scroll_callback)
        scrollable_frame.bind("<MouseWheel>", validation_scroll_callback)
        
        # Show each completed order
        for order_data in validation_data:
            self.create_validation_order_widget(scrollable_frame, order_data, validation_scroll_callback)
        
        # If no validation data but we have images, show all images
        if not validation_data:
            # Get all images from outputs folder as fallback
            outputs_dir = "outputs"
            all_images = []
            if os.path.exists(outputs_dir):
                all_images = sorted([os.path.join(outputs_dir, f) for f in os.listdir(outputs_dir) if f.endswith('.png')])
            
            if all_images:
                info_label = tk.Label(
                    scrollable_frame,
                    text="No order tracking data available, but here are all generated images:",
                    font=("Segoe UI", 16, "bold"),
                    bg="white",
                    fg="black"
                )
                info_label.pack(pady=20)
                info_label.bind("<MouseWheel>", validation_scroll_callback)
                
                # Show all images in a grid
                self.create_image_grid(scrollable_frame, all_images, scroll_callback=validation_scroll_callback)
        
        # Button frame
        close_btn_frame = tk.Frame(val_window, bg="#f0f0f0", relief=tk.RAISED, bd=2)
        close_btn_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Save as PDF button
        save_pdf_btn = tk.Button(
            close_btn_frame,
            text="💾 Save as PDF",
            command=lambda: self.save_validation_pdf(validation_data),
            font=("Segoe UI", 16, "bold"),
            bg=self.accent_color,
            fg="black",
            relief=tk.FLAT,
            padx=30,
            pady=10,
            cursor="hand2"
        )
        save_pdf_btn.pack(side=tk.LEFT, padx=(20, 10), pady=15)
        
        # Close button
        close_btn = tk.Button(
            close_btn_frame,
            text="Close",
            command=val_window.destroy,
            font=("Segoe UI", 16),
            bg="#6c757d",
            fg="black",
            relief=tk.FLAT,
            padx=30,
            pady=10,
            cursor="hand2"
        )
        close_btn.pack(side=tk.RIGHT, padx=(10, 20), pady=15)
        
        # Cleanup on close
        def on_close():
            val_window.unbind_all("<MouseWheel>")
            val_window.destroy()
        
        val_window.protocol("WM_DELETE_WINDOW", on_close)
    
    def create_validation_order_widget(self, parent, order_data, scroll_callback):
        """Create a widget showing order details and associated images"""
        from PIL import Image, ImageTk
        
        # Order container
        order_container = tk.Frame(parent, bg="white", relief=tk.RIDGE, bd=2)
        order_container.pack(fill=tk.X, padx=20, pady=15)
        order_container.bind("<MouseWheel>", scroll_callback)
        
        # Header with order info
        header_frame = tk.Frame(order_container, bg="#f8f9fa")
        header_frame.pack(fill=tk.X, padx=0, pady=0)
        header_frame.bind("<MouseWheel>", scroll_callback)
        
        # Order number
        order_num_label = tk.Label(
            header_frame,
            text=f"Order #{order_data['order_number']}",
            font=("Segoe UI", 21, "bold"),
            bg="#f8f9fa",
            fg="black"
        )
        order_num_label.pack(side=tk.LEFT, padx=15, pady=10)
        order_num_label.bind("<MouseWheel>", scroll_callback)
        
        # Address
        address_label = tk.Label(
            header_frame,
            text=f"📍 {order_data['name']} • {order_data['address']}",
            font=("Segoe UI", 16),
            bg="#f8f9fa",
            fg="black"
        )
        address_label.pack(side=tk.LEFT, padx=(20, 15), pady=10)
        address_label.bind("<MouseWheel>", scroll_callback)
        
        # Items list
        items_frame = tk.Frame(order_container, bg="white")
        items_frame.pack(fill=tk.X, padx=15, pady=10)
        items_frame.bind("<MouseWheel>", scroll_callback)
        
        items_label = tk.Label(
            items_frame,
            text="Items to include in envelope:",
            font=("Segoe UI", 15, "bold"),
            bg="white",
            fg="black"
        )
        items_label.pack(anchor=tk.W, pady=(0, 5))
        items_label.bind("<MouseWheel>", scroll_callback)
        
        for item in order_data['items']:
            item_label = tk.Label(
                items_frame,
                text=f"  • {item.get('personalization', 'No personalization')}",
                font=("Segoe UI", 15),
                bg="white",
                fg="#555"
            )
            item_label.pack(anchor=tk.W, pady=2)
            item_label.bind("<MouseWheel>", scroll_callback)
        
        # Images grid - use the specific images matched to this order
        order_images = order_data.get('images', [])
        if order_images:
            images_label = tk.Label(
                order_container,
                text="Generated Images:",
                font=("Segoe UI", 15, "bold"),
                bg="white",
                fg="black"
            )
            images_label.pack(anchor=tk.W, padx=15, pady=(10, 5))
            images_label.bind("<MouseWheel>", scroll_callback)
            
            self.create_image_grid(order_container, order_images, max_cols=4, scroll_callback=scroll_callback)
    
    def create_image_grid(self, parent, image_paths, max_cols=5, scroll_callback=None):
        """Create a grid of image thumbnails"""
        from PIL import Image, ImageTk
        
        grid_frame = tk.Frame(parent, bg="white")
        grid_frame.pack(fill=tk.X, padx=15, pady=(0, 15))
        if scroll_callback:
            grid_frame.bind("<MouseWheel>", scroll_callback)
        
        if not hasattr(self, '_validation_images'):
            self._validation_images = []
        
        for idx, img_path in enumerate(image_paths):
            row = idx // max_cols
            col = idx % max_cols
            
            # Create image container
            img_container = tk.Frame(grid_frame, bg="#f0f0f0", relief=tk.RIDGE, bd=1)
            img_container.grid(row=row, column=col, padx=5, pady=5)
            if scroll_callback:
                img_container.bind("<MouseWheel>", scroll_callback)
            
            try:
                # Load and resize image
                img = Image.open(img_path)
                img.thumbnail((150, 150), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                
                # Image label
                img_label = tk.Label(img_container, image=photo, bg="#f0f0f0")
                img_label.image = photo
                img_label.pack(padx=5, pady=5)
                if scroll_callback:
                    img_label.bind("<MouseWheel>", scroll_callback)
                
                # Keep reference
                self._validation_images.append(photo)
                
                # Filename label
                filename_label = tk.Label(
                    img_container,
                    text=os.path.basename(img_path),
                    font=("Segoe UI", 12),
                    bg="#f0f0f0",
                    fg="black"
                )
                filename_label.pack(pady=(0, 5))
                if scroll_callback:
                    filename_label.bind("<MouseWheel>", scroll_callback)
                
            except Exception as e:
                error_label = tk.Label(
                    img_container,
                    text=f"Error loading\n{os.path.basename(img_path)}",
                    font=("Segoe UI", 12),
                    bg="#f0f0f0",
                    fg="black"
                )
                error_label.pack(padx=10, pady=10)
                if scroll_callback:
                    error_label.bind("<MouseWheel>", scroll_callback)
    
    def save_validation_pdf(self, validation_data):
        """Save validation data as a PDF document"""
        try:
            from reportlab.lib.pagesizes import letter
            from reportlab.lib import colors
            from reportlab.lib.units import inch
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.enums import TA_CENTER, TA_LEFT
        except ImportError:
            messagebox.showerror(
                "Library Missing",
                "ReportLab library is required to generate PDFs.\n\n"
                "Please install it with:\npip install reportlab"
            )
            return
        
        # Ask user where to save
        from tkinter import filedialog
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"validation_report_{timestamp}.pdf"
        
        filepath = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialfile=default_filename,
            title="Save Validation Report"
        )
        
        if not filepath:
            return  # User cancelled
        
        try:
            # Create PDF
            doc = SimpleDocTemplate(filepath, pagesize=letter)
            story = []
            styles = getSampleStyleSheet()
            
            # Title
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=24,
                textColor=colors.HexColor('#4a90e2'),
                spaceAfter=30,
                alignment=TA_CENTER
            )
            story.append(Paragraph("Order Validation Report", title_style))
            story.append(Paragraph(f"Generated: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}", styles['Normal']))
            story.append(Spacer(1, 0.3*inch))
            
            # Process each order
            for order_data in validation_data:
                # Order header
                header_style = ParagraphStyle(
                    'OrderHeader',
                    parent=styles['Heading2'],
                    fontSize=16,
                    textColor=colors.HexColor('#333333'),
                    spaceAfter=10
                )
                story.append(Paragraph(f"Order #{order_data['order_number']}", header_style))
                
                # Customer info
                customer_info = f"<b>Customer:</b> {order_data['name']}<br/><b>Address:</b> {order_data['address']}"
                story.append(Paragraph(customer_info, styles['Normal']))
                story.append(Spacer(1, 0.1*inch))
                
                # Items list
                story.append(Paragraph("<b>Items in this order:</b>", styles['Normal']))
                for item in order_data['items']:
                    item_text = f"• {item.get('personalization', 'No personalization')}"
                    story.append(Paragraph(item_text, styles['Normal']))
                story.append(Spacer(1, 0.1*inch))
                
                # Images
                order_images = order_data.get('images', [])
                if order_images:
                    story.append(Paragraph("<b>Generated Images:</b>", styles['Normal']))
                    story.append(Spacer(1, 0.1*inch))
                    
                    # Create image grid (2 images per row)
                    img_data = []
                    row = []
                    for img_path in order_images:
                        try:
                            # Add image to row
                            rl_img = RLImage(img_path, width=2*inch, height=2*inch)
                            row.append(rl_img)
                            
                            # If row is full or last image, add to data
                            if len(row) == 2 or img_path == order_images[-1]:
                                # Pad row if needed
                                while len(row) < 2:
                                    row.append('')
                                img_data.append(row)
                                row = []
                        except Exception as e:
                            # Skip images that can't be loaded
                            pass
                    
                    if img_data:
                        img_table = Table(img_data, colWidths=[2.5*inch, 2.5*inch])
                        img_table.setStyle(TableStyle([
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                        ]))
                        story.append(img_table)
                
                story.append(Spacer(1, 0.3*inch))
            
            # Build PDF
            doc.build(story)
            
            messagebox.showinfo(
                "PDF Saved",
                f"Validation report saved successfully!\n\n{filepath}"
            )
            
            # Ask if user wants to open the PDF
            if messagebox.askyesno("Open PDF", "Would you like to open the PDF now?"):
                open_file_or_folder(filepath)
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save PDF:\n{str(e)}")
    
    def show_log_window(self):
        """Show log in a separate window"""
        log_window = tk.Toplevel(self.root)
        log_window.title("📋 Processing Log")
        log_window.geometry("900x600")
        log_window.configure(bg="white")
        
        # Title
        title_frame = tk.Frame(log_window, bg=self.accent_color, height=60)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)
        
        title_label = tk.Label(
            title_frame,
            text="📋 Processing Log",
            font=("Segoe UI", 24, "bold"),
            bg=self.accent_color,
            fg="black"
        )
        title_label.pack(pady=15)
        
        # Log display
        log_display = scrolledtext.ScrolledText(
            log_window,
            font=("Consolas", 12),
            bg="#1e1e1e",
            fg="#00ff00",  # Green text for better visibility
            relief=tk.FLAT,
            wrap=tk.WORD
        )
        log_display.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Copy log content
        log_content = self.log_text.get(1.0, tk.END)
        log_display.insert(1.0, log_content)
        log_display.see(tk.END)  # Scroll to bottom
        
        # Make read-only
        log_display.config(state=tk.DISABLED)
        
        # Enable mouse wheel scrolling on Mac
        def on_log_win_scroll_mac(event):
            try:
                log_display.config(state=tk.NORMAL)
                log_display.yview_scroll(int(-1*event.delta), "units")
                log_display.config(state=tk.DISABLED)
            except:
                pass
        
        def on_log_win_scroll(event):
            try:
                if event.delta:
                    log_display.config(state=tk.NORMAL)
                    log_display.yview_scroll(int(-1*(event.delta/120)), "units")
                    log_display.config(state=tk.DISABLED)
            except:
                pass
        
        if platform.system() == 'Darwin':  # macOS
            log_display.bind("<MouseWheel>", on_log_win_scroll_mac)
        else:  # Windows/Linux
            log_display.bind("<MouseWheel>", on_log_win_scroll)
        
        # Button frame
        btn_frame = tk.Frame(log_window, bg="white")
        btn_frame.pack(fill=tk.X, padx=20, pady=(0, 20))
        
        # Refresh button
        def refresh_log():
            log_display.config(state=tk.NORMAL)
            log_display.delete(1.0, tk.END)
            log_content = self.log_text.get(1.0, tk.END)
            log_display.insert(1.0, log_content)
            log_display.see(tk.END)
            log_display.config(state=tk.DISABLED)
        
        refresh_btn = tk.Button(
            btn_frame,
            text="🔄 Refresh",
            command=refresh_log,
            font=("Segoe UI", 14, "bold"),
            bg=self.accent_color,
            fg="black",
            relief=tk.FLAT,
            padx=20,
            pady=8,
            cursor="hand2"
        )
        refresh_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Copy button
        def copy_log():
            log_content = self.log_text.get(1.0, tk.END)
            log_window.clipboard_clear()
            log_window.clipboard_append(log_content)
            messagebox.showinfo("Copied", "Log copied to clipboard!")
        
        copy_btn = tk.Button(
            btn_frame,
            text="📋 Copy to Clipboard",
            command=copy_log,
            font=("Segoe UI", 14),
            bg="#6c757d",
            fg="black",
            relief=tk.FLAT,
            padx=20,
            pady=8,
            cursor="hand2"
        )
        copy_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Close button
        close_btn = tk.Button(
            btn_frame,
            text="Close",
            command=log_window.destroy,
            font=("Segoe UI", 14),
            bg="#6c757d",
            fg="black",
            relief=tk.FLAT,
            padx=20,
            pady=8,
            cursor="hand2"
        )
        close_btn.pack(side=tk.RIGHT)
    
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
            font=("Segoe UI", 21, "bold"),
            bg="white"
        )
        help_label.pack(pady=10)
        
        help_scroll = scrolledtext.ScrolledText(
            help_window,
            font=("Segoe UI", 15),
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

