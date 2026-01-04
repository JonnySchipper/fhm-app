"""
Single-script order processor for Disney character magnets
Reads CSV with character-name pairs, generates personalized images, and creates PDFs
No UI, server, or API required - just run and go!
"""

import os
import sys
import csv
import math
import shutil
import time
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont
from PyPDF2 import PdfReader, PdfWriter


# ============================================================================
# CONFIGURATION
# ============================================================================

# Images folder location (FHM_Images in parent directory)
PARENT_DIR = os.path.dirname(os.path.abspath(os.getcwd()))
IMAGES_DIR = os.path.join(PARENT_DIR, 'FHM_Images')

# Template PDF for output
TEMPLATE_PDF = "format.pdf"

# Output directories
OUTPUTS_DIR = "outputs"
TEMP_DIR = "temp"

# Font paths
FONT_WALTOGRAPH = "font/waltographUI.ttf"
FONT_BLUEBERRY = "font/blueberry.ttf"
FONT_FALLBACK = "font/waltograph42.otf"


# ============================================================================
# TEXT RENDERING FUNCTIONS (from add_names.py)
# ============================================================================

def _glyph_advance(font: ImageFont.FreeTypeFont, ch: str) -> float:
    """Width advance for a glyph, with a safe fallback."""
    try:
        return font.getlength(ch)
    except Exception:
        bbox = font.getbbox(ch)
        return max(1, bbox[2] - bbox[0])


def _render_glyph_rgba(font, ch, fill, stroke_width=0, stroke_fill=None):
    """Render a single glyph to a tight RGBA image (no clipping)."""
    bbox = font.getbbox(ch, stroke_width=stroke_width)
    w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
    pad = max(2, stroke_width + 1)
    img = Image.new("RGBA", (w + 2*pad, h + 2*pad), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.text(
        (pad - bbox[0], pad - bbox[1]),
        ch,
        font=font,
        fill=fill,
        stroke_width=stroke_width,
        stroke_fill=stroke_fill if stroke_fill else fill
    )
    return img


def draw_text_on_arc(
    base_img: Image.Image,
    text: str,
    *,
    center: tuple,
    radius: float,
    font_path: str,
    font_size: int = 80,
    fill=(255, 255, 255, 255),
    angle_deg: float = 180.0,
    outward: bool = True,
    kerning: float = 0.0,
    stroke_width: int = 0,
    stroke_fill=None,
):
    """Render text along a circular arc."""
    
    # Load font robustly
    try:
        font = ImageFont.truetype(font_path, font_size)
    except OSError:
        for fallback in [FONT_WALTOGRAPH, FONT_FALLBACK]:
            try:
                font = ImageFont.truetype(fallback, font_size)
                break
            except OSError:
                font = None
        if font is None:
            font = ImageFont.load_default()

    cx, cy = center

    # Measure total advance
    advances = [_glyph_advance(font, ch) + (kerning if i < len(text)-1 else 0)
                for i, ch in enumerate(text)]
    text_width = sum(advances)
    if radius <= 0:
        raise ValueError("radius must be > 0")

    # Adjust radius and position for longer names
    current_radius = radius
    y_offset = 0
    x_offset = 0
    angle_adjustment = 0
    if len(text) > 6:
        current_radius = radius + (len(text) - 6) * -6
        if len(text) > 10:
            current_radius = radius + (len(text) - 10) * -11
        y_offset = (len(text) - 6) * 2
        x_offset = (len(text) - 6) * .8
        angle_adjustment = (len(text) - 5) * .5
        if len(text) > 10:
            angle_adjustment = (len(text) - 10) * 1
    
    # Angular span
    theta_total = text_width / float(current_radius)
    adjusted_angle = angle_deg + angle_adjustment
    theta_center = math.radians(adjusted_angle)
    theta_start = theta_center - theta_total / 2.0

    # Place each glyph
    s_cum = 0.0
    for i, ch in enumerate(text):
        adv = advances[i]
        s_mid = s_cum + adv / 2.0
        theta = theta_start + (s_mid / radius)
        
        current_radius = radius
        if i >= 6:
            current_radius = radius + (i - 5)

        x = cx + current_radius * math.cos(theta) - x_offset
        y = cy - current_radius * math.sin(theta) - y_offset

        rot_deg = math.degrees(theta) - (90 if outward else -90)

        glyph_img = _render_glyph_rgba(font, ch, fill, stroke_width, stroke_fill)
        glyph_rot = glyph_img.rotate(rot_deg, resample=Image.BICUBIC, expand=True)

        gw, gh = glyph_rot.size
        paste_xy = (int(x - gw / 2), int(y - gh / 2))
        base_img.alpha_composite(glyph_rot, dest=paste_xy)

        s_cum += adv

    return base_img


def create_personalized_image(name, image_path, output_path):
    """
    Create a personalized image by adding curved text to the base image.
    
    Args:
        name: Personalization name (can be empty string for no text)
        image_path: Path to the source character image
        output_path: Where to save the personalized image
    """
    
    # If no name, just copy the original image
    if not name or name.strip() == "":
        print(f"  Creating image without text for {os.path.basename(image_path)}")
        shutil.copy2(image_path, output_path)
        
        # Small delay to ensure file is copied
        time.sleep(0.05)
        
        # Verify copy
        if not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
            raise RuntimeError(f"Failed to copy image: {output_path}")
        
        return output_path
    
    print(f"  Adding text '{name}' to {os.path.basename(image_path)}")
    
    # Determine font based on image filename
    font_path = FONT_WALTOGRAPH
    font_size = 100
    if "dog-" in image_path or "rubberduck-" in image_path:
        font_path = FONT_BLUEBERRY
        font_size = 90  # 10% smaller for blueberry font
    
    # Open and convert image
    img = Image.open(image_path).convert("RGBA")
    
    # Add curved text
    draw_text_on_arc(
        img,
        name,
        center=(780, 690),
        radius=600,
        font_path=font_path,
        font_size=font_size,
        fill=(0, 0, 0, 255),  # Black text
        angle_deg=270,  # Bottom of circle (upward curve)
        outward=False,  # Characters face toward center
        kerning=1.2,
        stroke_width=0,
        stroke_fill=(0, 0, 0, 255),
    )
    
    # Save with proper file handling
    output_img = img.convert("RGBA")
    output_img.save(output_path)
    output_img.close()  # Explicitly close
    img.close()  # Close original image too
    
    # Small delay to ensure file is written
    time.sleep(0.05)
    
    # Verify file was created and has content
    if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
        print(f"  ✓ Saved to {output_path}")
    else:
        print(f"  ✗ ERROR: Failed to save {output_path}")
        raise RuntimeError(f"Image file not properly created: {output_path}")
    
    return output_path


# ============================================================================
# PDF GENERATION FUNCTIONS (from pdf.py)
# ============================================================================

def png_to_pdf(png_path, pdf_path):
    """Convert PNG to PDF with proper file handling."""
    try:
        img = Image.open(png_path)
        img.save(pdf_path, "PDF", resolution=100.0)
        img.close()  # Explicitly close the image
        
        # Small delay to ensure file is written and closed properly
        time.sleep(0.1)
        
        # Verify the PDF was created
        if not os.path.exists(pdf_path):
            raise RuntimeError(f"PDF file was not created: {pdf_path}")
        
        # Wait for file to be fully written (check file size stabilizes)
        max_wait = 2.0  # Maximum 2 seconds
        start_time = time.time()
        last_size = 0
        while time.time() - start_time < max_wait:
            current_size = os.path.getsize(pdf_path)
            if current_size > 0 and current_size == last_size:
                break  # File size stable
            last_size = current_size
            time.sleep(0.05)
            
    except Exception as e:
        print(f"Error converting {png_path} to PDF: {e}")
        raise


def create_pdf_with_images(input_pdf, image1, image2, output_pdf):
    """
    Create a PDF by compositing images onto PDF template as a single image.
    Much more reliable than PDF merging.
    
    Args:
        input_pdf: Template PDF file path
        image1: Path to first image (top position)
        image2: Path to second image (bottom position)
        output_pdf: Output PDF file path
    """
    
    # Validation
    if not os.path.exists(input_pdf):
        raise FileNotFoundError(f"Template PDF not found: {input_pdf}")
    if not os.path.exists(image1):
        raise FileNotFoundError(f"Image 1 not found: {image1}")
    if not os.path.exists(image2):
        raise FileNotFoundError(f"Image 2 not found: {image2}")
    
    print(f"  Creating composite with {os.path.basename(image1)} and {os.path.basename(image2)}")
    
    # Create temp directory
    os.makedirs(TEMP_DIR, exist_ok=True)
    temp_composite = os.path.join(TEMP_DIR, "composite.png")
    
    try:
        # Step 1: Convert PDF template to image
        print(f"  Converting PDF template to image...")
        reader = PdfReader(input_pdf)
        num_pages = len(reader.pages)
        target_page = num_pages // 2  # Middle page
        
        # Get the page we want to use as template
        page = reader.pages[target_page]
        
        # Get page dimensions (in points)
        page_box = page.mediabox
        page_width = float(page_box.width)
        page_height = float(page_box.height)
        
        # Standard PDF is 612x792 points (8.5x11 inches at 72 DPI)
        # We'll work at higher resolution for quality
        dpi = 150
        img_width = int(page_width * dpi / 72)
        img_height = int(page_height * dpi / 72)
        
        # Convert PDF page to image using pdf2image if available, otherwise create blank
        try:
            from pdf2image import convert_from_path
            images = convert_from_path(input_pdf, dpi=dpi, first_page=target_page+1, last_page=target_page+1)
            base_image = images[0].convert('RGB')
        except ImportError:
            # pdf2image not available, create a white background
            print(f"  Note: pdf2image not available, using white background")
            base_image = Image.new('RGB', (img_width, img_height), 'white')
        
        print(f"  Base image size: {base_image.size}")
        
        # Step 2: Load and resize the character images
        char_img1 = Image.open(image1).convert('RGBA')
        char_img2 = Image.open(image2).convert('RGBA')
        
        # Calculate positions and sizes
        # Original coordinates from PDF (in points):
        # Top: x=121, y=396, size=500x500
        # Bottom: x=121, y=36, size=500x500
        
        # Scale to our DPI
        scale = dpi / 72
        target_size = int(500 * scale)
        x_pos = int(121 * scale)
        
        # PDF coordinates are from bottom-left, PIL uses top-left
        # y_top in PIL = page_height - (y_pdf + image_height)
        y_top = int(page_height * scale - (396 * scale + 500 * scale))
        y_bottom = int(page_height * scale - (36 * scale + 500 * scale))
        
        # Resize character images
        char_img1 = char_img1.resize((target_size, target_size), Image.Resampling.LANCZOS)
        char_img2 = char_img2.resize((target_size, target_size), Image.Resampling.LANCZOS)
        
        print(f"  Positioning images - Top: ({x_pos}, {y_top}), Bottom: ({x_pos}, {y_bottom})")
        
        # Step 3: Composite images onto base
        # Convert base to RGBA for alpha compositing
        base_rgba = base_image.convert('RGBA')
        
        # Paste images (bottom first, then top)
        base_rgba.alpha_composite(char_img2, (x_pos, y_bottom))
        base_rgba.alpha_composite(char_img1, (x_pos, y_top))
        
        # Convert back to RGB
        final_image = base_rgba.convert('RGB')
        
        # Save as temporary PNG
        final_image.save(temp_composite, 'PNG', dpi=(dpi, dpi))
        print(f"  Composite image saved")
        
        # Close images
        char_img1.close()
        char_img2.close()
        base_image.close()
        base_rgba.close()
        
        # Small delay
        time.sleep(0.1)
        
        # Step 4: Convert composite PNG to PDF
        print(f"  Converting to PDF...")
        final_image.save(output_pdf, 'PDF', resolution=dpi)
        final_image.close()
        
        # Wait for file to be written
        time.sleep(0.1)
        
        # Verify output
        if not os.path.exists(output_pdf):
            raise RuntimeError(f"Failed to create output PDF: {output_pdf}")
        
        output_size = os.path.getsize(output_pdf)
        if output_size == 0:
            raise RuntimeError(f"Output PDF is empty: {output_pdf}")
        
        print(f"  PDF created successfully: {output_size} bytes")
        
        # Clean up temp files
        try:
            if os.path.exists(temp_composite):
                os.remove(temp_composite)
        except Exception as e:
            print(f"  Warning: Could not clean up temp file: {e}")
        
        return output_pdf
        
    except Exception as e:
        print(f"  Error in create_pdf_with_images: {e}")
        # Clean up on error
        try:
            if os.path.exists(temp_composite):
                os.remove(temp_composite)
        except:
            pass
        raise


# ============================================================================
# CSV PROCESSING AND MAIN WORKFLOW
# ============================================================================

def read_csv_orders(csv_path):
    """
    Read CSV file with character-name pairs.
    
    Expected CSV format:
    - Column 1: Character name (without .png extension)
    - Column 2: Personalization name (can be empty for no text)
    
    Returns list of (character, name) tuples in order.
    """
    orders = []
    
    with open(csv_path, 'r', encoding='utf-8') as f:
        reader = csv.reader(f)
        
        # Skip header if it exists (detect common header patterns)
        first_row = next(reader, None)
        if first_row:
            # If first row looks like a header, skip it
            if first_row[0].lower() in ['character', 'characters', 'image', 'file', 'name']:
                pass  # Skip header
            else:
                # First row is data, process it
                if len(first_row) >= 2:
                    character = first_row[0].strip()
                    name = first_row[1].strip() if len(first_row) > 1 else ""
                    orders.append((character, name))
                elif len(first_row) == 1:
                    # Only character, no name
                    character = first_row[0].strip()
                    orders.append((character, ""))
        
        # Process remaining rows
        for row in reader:
            if not row or not row[0].strip():
                continue  # Skip empty rows
            
            character = row[0].strip()
            name = row[1].strip() if len(row) > 1 else ""
            orders.append((character, name))
    
    return orders


def find_image_file(character_name):
    """
    Find the image file for a character in the FHM_Images folder.
    
    Args:
        character_name: Character name without .png extension
        
    Returns:
        Full path to the image file, or None if not found
    """
    # Try exact match first
    image_filename = f"{character_name}.png"
    image_path = os.path.join(IMAGES_DIR, image_filename)
    
    if os.path.exists(image_path):
        return image_path
    
    # Try case-insensitive match
    if os.path.exists(IMAGES_DIR):
        for file in os.listdir(IMAGES_DIR):
            if file.lower() == image_filename.lower():
                return os.path.join(IMAGES_DIR, file)
    
    return None


def process_all_orders(csv_path):
    """
    Main processing function that reads CSV and generates all outputs.
    
    Steps:
    1. Read CSV with character-name pairs
    2. Generate personalized images
    3. Create PDFs with pairs of images
    """
    
    print("\n" + "="*70)
    print("DISNEY MAGNET ORDER PROCESSOR")
    print("="*70)
    
    # Validate inputs
    print("\n[1/6] Validating inputs...")
    
    if not os.path.exists(csv_path):
        print(f"✗ ERROR: CSV file not found: {csv_path}")
        return False
    
    if not os.path.exists(IMAGES_DIR):
        print(f"✗ ERROR: Images folder not found: {IMAGES_DIR}")
        print(f"   Please ensure FHM_Images folder exists in parent directory")
        return False
    
    if not os.path.exists(TEMPLATE_PDF):
        print(f"✗ ERROR: Template PDF not found: {TEMPLATE_PDF}")
        return False
    
    print(f"✓ CSV file: {csv_path}")
    print(f"✓ Images folder: {IMAGES_DIR}")
    print(f"✓ Template PDF: {TEMPLATE_PDF}")
    
    # Read orders
    print("\n[2/6] Reading orders from CSV...")
    orders = read_csv_orders(csv_path)
    
    if not orders:
        print("✗ ERROR: No orders found in CSV file")
        return False
    
    print(f"✓ Found {len(orders)} orders")
    for i, (char, name) in enumerate(orders, 1):
        name_display = name if name else "(no personalization)"
        print(f"  {i}. {char} → {name_display}")
    
    # Create output directory
    print("\n[3/6] Preparing output directories...")
    os.makedirs(OUTPUTS_DIR, exist_ok=True)
    
    # Clear previous outputs
    for file in os.listdir(OUTPUTS_DIR):
        if file.endswith('.png'):
            os.remove(os.path.join(OUTPUTS_DIR, file))
    
    print(f"✓ Output directory ready: {OUTPUTS_DIR}")
    
    # Generate personalized images
    print("\n[4/6] Generating personalized images...")
    
    generated_images = []
    
    for i, (character, name) in enumerate(orders, 1):
        print(f"\nProcessing order {i}/{len(orders)}: {character}")
        
        # Find source image
        source_image = find_image_file(character)
        if not source_image:
            print(f"  ✗ ERROR: Image not found for character '{character}'")
            print(f"     Looking for: {character}.png in {IMAGES_DIR}")
            continue
        
        # Generate output filename
        output_filename = f"{i}.png"
        output_path = os.path.join(OUTPUTS_DIR, output_filename)
        
        # Create personalized image
        try:
            create_personalized_image(name, source_image, output_path)
            generated_images.append(output_path)
        except Exception as e:
            print(f"  ✗ ERROR creating image: {e}")
            import traceback
            traceback.print_exc()
    
    if not generated_images:
        print("\n✗ ERROR: No images were generated successfully")
        return False
    
    print(f"\n✓ Generated {len(generated_images)} personalized images")
    
    # Create PDFs
    print("\n[5/6] Creating PDF outputs...")
    
    pdf_count = 0
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Process images in pairs
    for i in range(0, len(generated_images), 2):
        if i + 1 < len(generated_images):
            # We have a pair
            image1 = generated_images[i]
            image2 = generated_images[i + 1]
            
            pdf_num = (i // 2) + 1
            output_pdf = f"order_output_{timestamp}_{pdf_num}.pdf"
            
            print(f"\nCreating PDF {pdf_num}:")
            print(f"  Top: {os.path.basename(image1)}")
            print(f"  Bottom: {os.path.basename(image2)}")
            
            try:
                create_pdf_with_images(TEMPLATE_PDF, image1, image2, output_pdf)
                print(f"  ✓ Saved to {output_pdf}")
                pdf_count += 1
                
                # Small delay between PDFs to prevent file system issues
                time.sleep(0.15)
                
            except Exception as e:
                print(f"  ✗ ERROR creating PDF: {e}")
                import traceback
                traceback.print_exc()
        else:
            # Odd number of images - last one unpaired
            print(f"\nNote: Image {os.path.basename(generated_images[i])} has no pair")
            print(f"      You can manually create a PDF using pdf.py if needed")
    
    # Summary
    print("\n[6/6] Processing complete!")
    print("\n" + "="*70)
    print("SUMMARY")
    print("="*70)
    print(f"Orders processed: {len(orders)}")
    print(f"Images generated: {len(generated_images)}")
    print(f"PDFs created: {pdf_count}")
    print(f"\nPersonalized images saved to: {OUTPUTS_DIR}/")
    if pdf_count > 0:
        print(f"PDF files saved to current directory: order_output_{timestamp}_*.pdf")
    print("="*70 + "\n")
    
    return True


# ============================================================================
# COMMAND LINE INTERFACE
# ============================================================================

def print_usage():
    """Print usage instructions."""
    print("\n" + "="*70)
    print("DISNEY MAGNET ORDER PROCESSOR")
    print("="*70)
    print("\nUsage:")
    print(f"  python {os.path.basename(__file__)} <csv_file>")
    print("\nCSV Format:")
    print("  Column 1: Character name (without .png extension)")
    print("  Column 2: Personalization name (optional)")
    print("\nExample CSV:")
    print("  character,name")
    print("  mickey-captain,Johnny")
    print("  minnie-normal,Sarah")
    print("  donald-pumpkin,")
    print("  stitch-normal,Michael")
    print("\nNote:")
    print("  - Images must be in FHM_Images folder in parent directory")
    print("  - format.pdf template must be in current directory")
    print("  - Output images saved to outputs/ folder")
    print("  - Output PDFs saved to current directory")
    print("="*70 + "\n")


def main():
    """Main entry point."""
    
    # Check arguments
    if len(sys.argv) < 2:
        print_usage()
        
        # Look for CSV files in current directory
        csv_files = [f for f in os.listdir('.') if f.endswith('.csv')]
        if csv_files:
            print(f"Found {len(csv_files)} CSV file(s) in current directory:")
            for f in csv_files:
                print(f"  - {f}")
            print("\nTo process, run:")
            print(f"  python {os.path.basename(__file__)} {csv_files[0]}")
        
        sys.exit(1)
    
    csv_path = sys.argv[1]
    
    # Process orders
    success = process_all_orders(csv_path)
    
    if success:
        print("✓ All done! Your personalized magnets are ready.")
        sys.exit(0)
    else:
        print("✗ Processing failed. Please check the errors above.")
        sys.exit(1)


if __name__ == "__main__":
    main()

