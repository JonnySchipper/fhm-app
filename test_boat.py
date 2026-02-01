"""
Quick Boat Test Script - For fine-tuning text placement and PDF positioning
Run this over and over to test your adjustments!
"""

import os
import sys
import math
import shutil
from PIL import Image, ImageDraw, ImageFont
from PyPDF2 import PdfReader, PdfWriter

# ============================================================================
# BOAT CONFIGURATION - EDIT THESE TO FINE-TUNE!
# ============================================================================

# Text positioning for boats
BOAT_TEXT_CENTER = (950, 4220)    # (X, Y) center BELOW text = frown curve
BOAT_TEXT_RADIUS = 2700            # Arc radius in pixels - larger = flatter curve
BOAT_TEXT_ANGLE = 90               # Angle in degrees (90 = centered at top of arc)
BOAT_TEXT_OUTWARD = True           # Flipped for center at Y=2800
BOAT_TEXT_KERNING = 1.2            # Letter spacing

# === NAME LENGTH SCALING - Adjust these to handle different name lengths ===
# Reference name length (what the base settings are tuned for)
BOAT_TEXT_REFERENCE_LENGTH = 16    # "The Smith Family" = 16 chars (including spaces)

# Base font size (for reference length name)
BOAT_TEXT_BASE_FONT_SIZE = 120

# Font scaling: how much to adjust font size per character difference
# Positive = larger font for shorter names, smaller for longer names
BOAT_TEXT_FONT_SCALE_PER_CHAR = 4  # Decrease font by 4px per extra char

# Minimum and maximum font sizes (safety limits)
BOAT_TEXT_MIN_FONT_SIZE = 60
BOAT_TEXT_MAX_FONT_SIZE = 180

# Radius scaling: adjust curve radius based on name length
BOAT_TEXT_RADIUS_SCALE_PER_CHAR = 50  # Increase radius by 50px per extra char

# Y position scaling: move text up/down based on name length  
BOAT_TEXT_Y_SCALE_PER_CHAR = 50     # Adjust Y position (0 = no adjustment)

# Boat PDF positioning (single centered image on boat_format.pdf)
BOAT_PDF_X = -9                   # X position on PDF (left edge of image)
BOAT_PDF_Y = 80                   # Y position on PDF (bottom edge of image)
BOAT_PDF_SCALE = 1.41             # Scale factor - INCREASED 35% (was 1.027)

# ============================================================================
# TEST CONFIGURATION - CHANGE THESE FOR YOUR TEST
# ============================================================================

TEST_BOAT_IMAGE = "boats/boat_fantasy.png"  # Which boat to test
TEST_OUTPUT_IMAGE = "test_boat_output.png"   # Output image file
TEST_OUTPUT_PDF = "test_boat_output.pdf"     # Output PDF file

# === TEST NAMES - Uncomment one to test different lengths ===
# TEST_NAME = "The Smith Family"               # 16 chars - REFERENCE (what settings are tuned for)
# TEST_NAME = "Bob"                          # 3 chars - very short
# TEST_NAME = "Johnson"                      # 7 chars - short
# TEST_NAME = "The Johnsons"                 # 12 chars - medium
TEST_NAME = "The Christopher Family"       # 23 chars - long
# TEST_NAME = "The Bartholomew-Henderson Family"  # 35 chars - very long

# ============================================================================
# CODE - Don't edit below unless you know what you're doing
# ============================================================================

FONT_WALTOGRAPH = "font/waltographUI.ttf"
FONT_FALLBACK = "font/waltograph42.otf"
TEMP_DIR = "temp"


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
    """Render text along a circular arc - SIMPLIFIED for boats (symmetric curve)."""
    
    # Load font robustly
    try:
        font = ImageFont.truetype(font_path, font_size)
    except OSError:
        try:
            font = ImageFont.truetype(FONT_FALLBACK, font_size)
        except OSError:
            font = ImageFont.load_default()

    cx, cy = center

    # Measure total advance
    advances = [_glyph_advance(font, ch) + (kerning if i < len(text)-1 else 0)
                for i, ch in enumerate(text)]
    text_width = sum(advances)
    if radius <= 0:
        raise ValueError("radius must be > 0")

    # Angular span - SIMPLIFIED: no adjustments, perfectly symmetric
    theta_total = text_width / float(radius)
    theta_center = math.radians(angle_deg)
    theta_start = theta_center - theta_total / 2.0

    # Place each glyph - SIMPLIFIED: constant radius for all letters
    s_cum = 0.0
    for i, ch in enumerate(text):
        adv = advances[i]
        s_mid = s_cum + adv / 2.0
        theta = theta_start + (s_mid / radius)

        x = cx + radius * math.cos(theta)
        y = cy - radius * math.sin(theta)

        rot_deg = math.degrees(theta) - (90 if outward else -90)

        glyph_img = _render_glyph_rgba(font, ch, fill, stroke_width, stroke_fill)
        glyph_rot = glyph_img.rotate(rot_deg, resample=Image.BICUBIC, expand=True)

        gw, gh = glyph_rot.size
        paste_xy = (int(x - gw / 2), int(y - gh / 2))
        base_img.alpha_composite(glyph_rot, dest=paste_xy)

        s_cum += adv

    return base_img


def calculate_scaled_settings(name):
    """Calculate font size, radius, and position based on name length."""
    name_length = len(name)
    length_diff = name_length - BOAT_TEXT_REFERENCE_LENGTH
    
    # Calculate scaled font size
    font_size = BOAT_TEXT_BASE_FONT_SIZE - (length_diff * BOAT_TEXT_FONT_SCALE_PER_CHAR)
    font_size = max(BOAT_TEXT_MIN_FONT_SIZE, min(BOAT_TEXT_MAX_FONT_SIZE, font_size))
    
    # Calculate scaled radius
    radius = BOAT_TEXT_RADIUS + (length_diff * BOAT_TEXT_RADIUS_SCALE_PER_CHAR)
    
    # Calculate scaled Y position
    center_x, center_y = BOAT_TEXT_CENTER
    center_y = center_y + (length_diff * BOAT_TEXT_Y_SCALE_PER_CHAR)
    
    return font_size, radius, (center_x, center_y)


def create_test_boat_image(name, boat_image_path, output_path):
    """Create a test boat image with curved text."""
    print(f"\n🚢 Creating boat image with text: '{name}'")
    print(f"   Source: {boat_image_path}")
    print(f"   Name length: {len(name)} chars (reference: {BOAT_TEXT_REFERENCE_LENGTH})")
    
    if not os.path.exists(boat_image_path):
        print(f"❌ ERROR: Boat image not found: {boat_image_path}")
        return None
    
    # Calculate scaled settings based on name length
    font_size, radius, center = calculate_scaled_settings(name)
    print(f"   Scaled font size: {font_size} (base: {BOAT_TEXT_BASE_FONT_SIZE})")
    print(f"   Scaled radius: {radius} (base: {BOAT_TEXT_RADIUS})")
    print(f"   Scaled center: {center}")
    
    # Open and convert image
    img = Image.open(boat_image_path).convert("RGBA")
    
    # Add curved text using boat-specific configuration
    # Reverse the text so it reads correctly on the arc
    draw_text_on_arc(
        img,
        name[::-1],  # Reverse the string
        center=center,
        radius=radius,
        font_path=FONT_WALTOGRAPH,
        font_size=font_size,
        fill=(255, 255, 255, 255),  # White text
        angle_deg=BOAT_TEXT_ANGLE,
        outward=BOAT_TEXT_OUTWARD,
        kerning=BOAT_TEXT_KERNING,
        stroke_width=0,
        stroke_fill=(255, 255, 255, 255),
    )
    
    # Save
    img.save(output_path)
    img.close()
    
    print(f"✅ Saved boat image to: {output_path}")
    return output_path


def create_test_boat_pdf(boat_image_path, output_pdf):
    """Create a test boat PDF."""
    print(f"\n📄 Creating boat PDF...")
    
    boat_template = "boat_format.pdf"
    
    if not os.path.exists(boat_template):
        print(f"❌ ERROR: Boat template not found: {boat_template}")
        return False
    
    if not os.path.exists(boat_image_path):
        print(f"❌ ERROR: Boat image not found: {boat_image_path}")
        return False
    
    # Convert image to temporary PDF
    os.makedirs(TEMP_DIR, exist_ok=True)
    temp_pdf = os.path.join(TEMP_DIR, "temp_test_boat.pdf")
    
    if os.path.exists(temp_pdf):
        os.remove(temp_pdf)
    
    # PNG to PDF
    img = Image.open(boat_image_path)
    img.save(temp_pdf, "PDF", resolution=100.0)
    
    # Read the template
    reader = PdfReader(boat_template)
    num_pages = len(reader.pages)
    target = num_pages // 2
    
    # Read the image PDF
    reader_img = PdfReader(temp_pdf)
    img_page = reader_img.pages[0]
    
    # Apply transformation
    img_width = 500
    target_scale = (img_width / 1500) * BOAT_PDF_SCALE
    
    # Transform: (a, b, c, d, e, f) where a=d=scale, e=x, f=y
    img_page.add_transformation((target_scale, 0, 0, target_scale, BOAT_PDF_X, BOAT_PDF_Y))
    
    # Merge
    original_page = reader.pages[target]
    original_page.merge_page(img_page)
    
    # Write
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    
    with open(output_pdf, "wb") as out_file:
        writer.write(out_file)
    
    # Cleanup
    if os.path.exists(temp_pdf):
        os.remove(temp_pdf)
    
    print(f"✅ Saved boat PDF to: {output_pdf}")
    return True


def main():
    """Run the test."""
    print("\n" + "="*70)
    print("🚢 BOAT TEST SCRIPT - Fine-tune your boat text & PDF placement!")
    print("="*70)
    
    print("\n📝 Current Configuration:")
    print(f"   Text Center: {BOAT_TEXT_CENTER}")
    print(f"   Text Radius: {BOAT_TEXT_RADIUS}")
    print(f"   Text Angle: {BOAT_TEXT_ANGLE}° (90=upward, 270=downward)")
    print(f"   Base Font Size: {BOAT_TEXT_BASE_FONT_SIZE} (scales with name length)")
    print(f"   Text Outward: {BOAT_TEXT_OUTWARD}")
    print(f"   PDF Position: ({BOAT_PDF_X}, {BOAT_PDF_Y})")
    print(f"   PDF Scale: {BOAT_PDF_SCALE}")
    
    print(f"\n🎯 Test Settings:")
    print(f"   Boat: {TEST_BOAT_IMAGE}")
    print(f"   Name: '{TEST_NAME}'")
    
    # Create test image
    result_image = create_test_boat_image(TEST_NAME, TEST_BOAT_IMAGE, TEST_OUTPUT_IMAGE)
    
    if not result_image:
        print("\n❌ Test failed - could not create boat image")
        return
    
    # Create test PDF
    result_pdf = create_test_boat_pdf(TEST_OUTPUT_IMAGE, TEST_OUTPUT_PDF)
    
    if not result_pdf:
        print("\n❌ Test failed - could not create boat PDF")
        return
    
    print("\n" + "="*70)
    print("✅ TEST COMPLETE!")
    print("="*70)
    print(f"\n📁 Output files created:")
    print(f"   • {TEST_OUTPUT_IMAGE} - Boat with text")
    print(f"   • {TEST_OUTPUT_PDF} - Final PDF")
    
    print(f"\n💡 To fine-tune:")
    print(f"   1. Edit the configuration at the top of this script")
    print(f"   2. Run: python test_boat.py")
    print(f"   3. Check the output files")
    print(f"   4. Repeat until perfect!")
    
    print(f"\n🔧 Common adjustments:")
    print(f"   • Move text up/down: Change BOAT_TEXT_CENTER Y value")
    print(f"   • Move text left/right: Change BOAT_TEXT_CENTER X value")
    print(f"   • Flatten/steepen curve: Change BOAT_TEXT_RADIUS")
    print(f"   • Flip curve: Change BOAT_TEXT_ANGLE (90 vs 270)")
    print(f"   • Base text size: Change BOAT_TEXT_BASE_FONT_SIZE")
    print(f"   • Scaling per char: Change BOAT_TEXT_FONT_SCALE_PER_CHAR")
    print(f"   • Image position on PDF: Change BOAT_PDF_X, BOAT_PDF_Y")
    print("="*70 + "\n")


if __name__ == "__main__":
    main()

