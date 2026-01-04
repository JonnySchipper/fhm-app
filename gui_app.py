"""
Disney Magnet Order Processor - GUI Application
Interactive interface for processing orders with real-time feedback
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import tkinter.font as tkfont
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
                orders.append((character, name if name else "(no personalization)"))
        
        # Update preview
        self.preview_text.delete(1.0, tk.END)
        self.preview_text.insert(tk.END, f"{'Character':<30} {'Name':<25}\n")
        self.preview_text.insert(tk.END, "-" * 55 + "\n")
        
        for character, name in orders:
            self.preview_text.insert(tk.END, f"{character:<30} {name:<25}\n")
        
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
        
        # Store order data (will be updated as user edits)
        order_data = []
        order_widgets = []  # Store references to widgets for updates
        
        # Function to create an order row
        def create_order_row(character="", name=""):
            """Create a single editable order row"""
            idx = len(order_data)
            order_data.append({'character': character, 'name': name})
            
            # Create frame for this order
            order_frame = tk.Frame(scrollable_frame, bg="white", relief=tk.RIDGE, bd=2)
            order_frame.pack(fill=tk.X, padx=10, pady=8)
            
            # Image preview (left side) - fixed size container
            image_container = tk.Frame(order_frame, bg="white", width=100, height=100)
            image_container.pack(side=tk.LEFT, padx=10, pady=10)
            image_container.pack_propagate(False)  # Prevent resizing
            
            # Load initial image
            image_filename = f"{character}.png"
            image_path = os.path.join(images_dir, image_filename)
            
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
            
            # Character dropdown with search
            char_var = tk.StringVar(value=character)
            char_combo = ttk.Combobox(
                char_row,
                textvariable=char_var,
                values=available_images,
                font=("Segoe UI", 9),
                width=28,
                state="normal"
            )
            char_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
            
            # Search button
            def make_search_handler(char_v):
                return lambda: open_image_search(char_v)
            
            search_btn = tk.Button(
                char_row,
                text="🔍 Search",
                command=make_search_handler(char_var),
                font=("Segoe UI", 8),
                bg="#17a2b8",
                fg="black",
                relief=tk.FLAT,
                padx=8,
                pady=2,
                cursor="hand2"
            )
            search_btn.pack(side=tk.LEFT)
            
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
            def make_update_handler(idx, img_lbl, char_v, name_v, status_lbl):
                def update_image(*args):
                    new_char = char_v.get()
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
                            status_lbl.config(text="✓ Ready", fg="#5cb85c")
                        except:
                            img_lbl.config(image="", text="Error\nLoading", bg="#fff0f0")
                            status_lbl.config(text="❌ Error", fg="#d9534f")
                    else:
                        img_lbl.config(image="", text="No\nImage", bg="#f0f0f0")
                        status_lbl.config(text="⚠ Not found", fg="#f0ad4e")
                return update_image
            
            update_handler = make_update_handler(idx, img_label, char_var, name_var, status_label)
            char_combo.bind('<<ComboboxSelected>>', update_handler)
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
            found = sum(1 for o in active_orders if os.path.exists(os.path.join(images_dir, f"{o['character']}.png")))
            missing = len(active_orders) - found
            summary_label.config(text=f"✓ Ready: {found}  |  ⚠ Issues: {missing}  |  Total: {len(active_orders)}")
        
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
            text="",
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
            
            # Check for missing images
            missing = []
            for o in active_orders:
                if not os.path.exists(os.path.join(images_dir, f"{o['character']}.png")):
                    missing.append(f"{o['character']} (for {o['name']})")
            
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
            
            # Update main input with edited orders
            orders_text = '\n'.join([f"{o['character']},{o['name']}" for o in active_orders])
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
            
            orders_text = '\n'.join([f"{o['character']},{o['name']}" for o in active_orders])
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
        
        # Enable mousewheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Cleanup on close
        def on_close():
            canvas.unbind_all("<MouseWheel>")
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
                
                # Find and merge all PDFs
                self.log("Creating master PDF...", "info")
                pdf_files = []
                for file in sorted(os.listdir('.')):
                    if file.startswith('order_output_') and file.endswith('.pdf'):
                        pdf_files.append(file)
                
                if pdf_files:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    master_pdf_name = f"MASTER_ORDER_{timestamp}.pdf"
                    
                    if self.merge_pdfs(pdf_files, master_pdf_name):
                        self.master_pdf_path = master_pdf_name
                        self.master_pdf_btn.config(state=tk.NORMAL)
                        self.status_text.set(f"✓ Complete! Master PDF: {master_pdf_name}")
                        messagebox.showinfo(
                            "Success",
                            f"Orders processed successfully!\n\n"
                            f"✓ {len(pdf_files)} individual PDFs created\n"
                            f"✓ Master PDF created: {master_pdf_name}\n\n"
                            f"Click 'Open Master PDF' to view!"
                        )
                    else:
                        self.status_text.set("✓ Orders processed successfully!")
                        messagebox.showinfo(
                            "Success",
                            "Orders processed successfully!\n\n"
                            "Check:\n"
                            "• outputs/ folder for images\n"
                            "• Current directory for PDFs"
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
            
    def merge_pdfs(self, pdf_files, output_path):
        """Merge multiple PDF files into one master PDF"""
        try:
            from PyPDF2 import PdfReader, PdfWriter
            
            self.log(f"Merging {len(pdf_files)} PDFs into master PDF...", "info")
            
            writer = PdfWriter()
            
            for pdf_file in pdf_files:
                if os.path.exists(pdf_file):
                    reader = PdfReader(pdf_file)
                    for page in reader.pages:
                        writer.add_page(page)
                    self.log(f"  Added {pdf_file}", "info")
            
            with open(output_path, 'wb') as output_file:
                writer.write(output_file)
            
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
        
    def parse_with_ai_thread(self, raw_text, use_reasoning=True):
        """Parse with AI in background thread"""
        try:
            self.ai_processing = True
            # Disable both AI buttons
            self.root.after(0, lambda: self.ai_parse_btn.config(state=tk.DISABLED, text="🤖 Processing..."))
            self.root.after(0, lambda: self.quick_parse_btn.config(state=tk.DISABLED, text="⚡ Processing..."))
            
            model_type = "reasoning" if use_reasoning else "quick"
            self.root.after(0, lambda: self.status_text.set(f"AI is parsing your order ({model_type})..."))
            self.root.after(0, lambda: self.log(f"Calling Grok AI ({model_type} model) to parse order text...", "info"))
            
            # Call Grok API with reasoning parameter
            result = self.call_grok_api(self.image_list, raw_text, use_reasoning)
            
            if not result:
                self.root.after(0, lambda: messagebox.showerror("AI Error", "Failed to parse orders. Check the log for details."))
                self.root.after(0, lambda: self.log("AI parsing failed", "error"))
                return
            
            # Convert result to simple format
            orders = []
            for name, image_file in result.items():
                if name not in ['_order', 'unmatched'] and image_file:
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
            self.root.after(0, lambda: self.order_input.config(fg="#333"))
            self.root.after(0, lambda: self.update_count())
            self.root.after(0, lambda: self.preview_orders())
            
            self.root.after(0, lambda: self.log(f"✓ AI parsed {len(orders)} orders successfully!", "success"))
            self.root.after(0, lambda: self.status_text.set(f"✓ AI found {len(orders)} orders! Review and click Process."))
            self.root.after(0, lambda: messagebox.showinfo("Success!", f"AI parsed {len(orders)} orders!\n\nReview them below and click 'Process Orders' when ready."))
            
        except Exception as e:
            error_msg = str(e)
            self.root.after(0, lambda msg=error_msg: self.log(f"AI Error: {msg}", "error"))
            self.root.after(0, lambda msg=error_msg: messagebox.showerror("Error", f"AI parsing failed:\n{msg}"))
            
        finally:
            self.ai_processing = False
            # Re-enable both AI buttons
            self.root.after(0, lambda: self.ai_parse_btn.config(state=tk.NORMAL, text="✨ Parse with AI (Grok)"))
            self.root.after(0, lambda: self.quick_parse_btn.config(state=tk.NORMAL, text="⚡ Quick Parse"))
            
    def call_grok_api(self, image_list, order_text, use_reasoning=True):
        """Call Grok API to parse order text"""
        images_str = ', '.join(image_list)
        
        # Choose model based on use_reasoning
        model = "grok-4-1-fast-reasoning" if use_reasoning else "grok-4-1-fast-non-reasoning"
        
        prompt = f"""You are parsing Disney-themed magnet orders. Convert the order text into a simple character-name format.

Available images: {images_str[:500]}{'...' if len(images_str) > 500 else ''}

Order text:
{order_text}

Instructions:
1. Find all personalization names in the order
2. Match each name to a character based on context
3. Determine the theme/variant (captain, pumpkin, witch, halloween, pirate, christmas, etc.)
4. Match to available images using format: charactername-variant (e.g., mickey-captain, stitch-captain)
5. If theme unclear, use "captain" (NOT normal)
6. Only return matches you're confident about

Return ONLY a Python dictionary, no extra text:
{{
  "PersonName1": "character-variant.png",
  "PersonName2": "character-variant.png"
}}

Example input: "Captain Mickey for Johnny, Captain Stitch for Sarah"
Example output: {{"Johnny": "mickey-captain.png", "Sarah": "stitch-captain.png"}}

Be smart about:
- Typos (Micky -> mickey, Stitsh -> stitch)
- Variations (Mike Wazowski -> mikewazowski, Buzz Lightyear -> buzzlightyear)
- Character context from theme
- Common Disney character names

Return the dictionary now:"""
        
        headers = {
            "Authorization": f"Bearer {GROK_API_KEY}",
            "Content-Type": "application/json"
        }
        
        data = {
            "model": model,
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.1,
            "max_tokens": 16000
        }
        
        try:
            self.root.after(0, lambda: self.log("Sending request to Grok AI...", "info"))
            response = requests.post(GROK_API_URL, headers=headers, json=data, timeout=30)
            response.raise_for_status()
            
            result = response.json()
            content = result['choices'][0]['message']['content']
            
            self.root.after(0, lambda: self.log(f"Received response from AI", "info"))
            
            # Extract dictionary
            start = content.find('{')
            end = content.rfind('}') + 1
            
            if start != -1 and end != 0:
                dict_str = content[start:end]
                matches = json.loads(dict_str)
                return matches
            else:
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

