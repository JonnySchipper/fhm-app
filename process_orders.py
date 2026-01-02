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
    
    # Save
    img.convert("RGBA").save(output_path)
    
    if os.path.exists(output_path):
        print(f"  ✓ Saved to {output_path}")
    else:
        print(f"  ✗ ERROR: Failed to save {output_path}")
    
    return output_path


# ============================================================================
# PDF GENERATION FUNCTIONS (from pdf.py)
# ============================================================================

def png_to_pdf(png_path, pdf_path):
    """Convert PNG to PDF."""
    try:
        img = Image.open(png_path)
        img.save(pdf_path, "PDF", resolution=100.0)
    except Exception as e:
        print(f"Error converting {png_path} to PDF: {e}")
        raise


def create_pdf_with_images(input_pdf, image1, image2, output_pdf):
    """
    Overlay two images on a PDF template.
    
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
    
    # Validate images
    try:
        with Image.open(image1) as img1_test:
            if img1_test.size[0] == 0 or img1_test.size[1] == 0:
                raise ValueError(f"Image 1 has invalid dimensions: {image1}")
        with Image.open(image2) as img2_test:
            if img2_test.size[0] == 0 or img2_test.size[1] == 0:
                raise ValueError(f"Image 2 has invalid dimensions: {image2}")
    except Exception as e:
        raise ValueError(f"Invalid image file(s): {e}")
    
    # Convert images to temporary PDFs
    os.makedirs(TEMP_DIR, exist_ok=True)
    temp_pdf1 = os.path.join(TEMP_DIR, "temp_1.pdf")
    temp_pdf2 = os.path.join(TEMP_DIR, "temp_2.pdf")
    
    # Clean up any existing temp files
    for temp_file in [temp_pdf1, temp_pdf2]:
        if os.path.exists(temp_file):
            os.remove(temp_file)
    
    png_to_pdf(image1, temp_pdf1)
    png_to_pdf(image2, temp_pdf2)
    
    # Verify temp PDFs
    if not os.path.exists(temp_pdf1) or os.path.getsize(temp_pdf1) == 0:
        raise RuntimeError(f"Failed to create valid temp PDF for image 1")
    if not os.path.exists(temp_pdf2) or os.path.getsize(temp_pdf2) == 0:
        raise RuntimeError(f"Failed to create valid temp PDF for image 2")
    
    # Read PDFs
    reader = PdfReader(input_pdf)
    num_pages = len(reader.pages)
    target = num_pages // 2
    
    reader_img1 = PdfReader(temp_pdf1)
    reader_img2 = PdfReader(temp_pdf2)
    
    img_page1 = reader_img1.pages[0]
    img_page2 = reader_img2.pages[0]
    
    # Size and positioning
    img_width = 500
    img_height = 500
    scale_factor = 1.027
    common_x = 121
    
    # Top image position
    tx1 = common_x
    ty1 = 396
    
    # Bottom image position
    tx2 = common_x
    ty2 = 36
    
    # Apply transformations
    target_scale = (img_width / 1500) * scale_factor
    img_page1.add_transformation((target_scale, 0, 0, target_scale, tx1, ty1))
    img_page2.add_transformation((target_scale, 0, 0, target_scale, tx2, ty2))
    
    # Overlay on target page
    original_page = reader.pages[target]
    original_page.merge_page(img_page2)
    original_page.merge_page(img_page1)
    
    # Create output
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    
    with open(output_pdf, "wb") as out_file:
        writer.write(out_file)
    
    # Verify output
    if not os.path.exists(output_pdf):
        raise RuntimeError(f"Failed to create output PDF: {output_pdf}")
    
    output_size = os.path.getsize(output_pdf)
    if output_size == 0:
        raise RuntimeError(f"Output PDF is empty: {output_pdf}")
    
    # Verify readability
    try:
        verification_reader = PdfReader(output_pdf)
        if len(verification_reader.pages) == 0:
            raise RuntimeError(f"Output PDF has no pages")
    except Exception as e:
        raise RuntimeError(f"Output PDF validation failed: {e}")
    
    # Clean up temp files
    try:
        os.remove(temp_pdf1)
        os.remove(temp_pdf2)
    except Exception as e:
        print(f"Warning: Could not clean up temp files: {e}")
    
    return output_pdf


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

