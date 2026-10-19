# generate_ppt.py
# Generates an editable 5-slide PPTX: sai_srikar_Visualstroyboard_ModernBoho.pptx
# Requirements: pip install python-pptx Pillow requests

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from PIL import Image
import requests, os, io

OUTFILE = "sai_srikar_Visualstroyboard_ModernBoho.pptx"
TMP_IMG_DIR = "ppt_images"
os.makedirs(TMP_IMG_DIR, exist_ok=True)

# Image queries (Unsplash Source — returns a matching image)
image_queries = {
    "crop_top_1": "men crop top fashion",
    "crop_top_2": "gender fluid mens fashion crop top",
    "suit_feminine_1": "mens suit details lapel feminine tailoring",
    "suit_feminine_2": "gender neutral tailored suit",
    "lace_texture": "lace fabric texture",
    "silk_texture": "silk fabric texture",
    "boho_print": "boho textile pattern",
    "pastel_interior": "boho interior pastel",
    "street_style": "street style mens fashion feminine",
    "runway": "runway gender fluid fashion"
}

def download_images():
    local_paths = {}
    for key, q in image_queries.items():
        url = f"https://source.unsplash.com/1200x800/?{q.replace(' ',',')}"
        try:
            r = requests.get(url, timeout=30)
            r.raise_for_status()
            img = Image.open(io.BytesIO(r.content)).convert("RGB")
            path = os.path.join(TMP_IMG_DIR, f"{key}.jpg")
            img.save(path, format="JPEG", quality=85)
            local_paths[key] = path
            print("Downloaded:", key)
        except Exception as e:
            print("Failed to download", q, e)
    return local_paths

def hex_to_rgb(hexstr):
    h = hexstr.lstrip('#')
    return tuple(int(h[i:i+2], 16) for i in (0,2,4))

def add_title_slide(prs, title, subtitle, bg_hex):
    slide_layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(slide_layout)
    # background
    rgb = hex_to_rgb(bg_hex)
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor(*rgb)
    # Title box
    tx = slide.shapes.add_textbox(Inches(0.5), Inches(1.0), Inches(9), Inches(1.2)).text_frame
    p = tx.paragraphs[0]
    p.text = title
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255,255,255)
    p.alignment = PP_ALIGN.LEFT
    # subtitle
    st = slide.shapes.add_textbox(Inches(0.5), Inches(2.2), Inches(9), Inches(1)).text_frame
    p = st.paragraphs[0]
    p.text = subtitle
    p.font.size = Pt(18)
    p.font.color.rgb = RGBColor(255,255,255)
    return slide

def add_palette_slide(prs, palettes):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    left = Inches(0.5)
    top = Inches(0.6)
    # Title
    t = slide.shapes.add_textbox(left, Inches(0.3), Inches(9), Inches(0.5)).text_frame
    t.paragraphs[0].text = "Mood & Colour Palette — Complementary pairs"
    t.paragraphs[0].font.size = Pt(22)
    # Draw swatches
    sw_w = Inches(1.2)
    gap = Inches(0.3)
    x = left
    y = top
    for i, pair in enumerate(palettes):
        a_hex, b_hex = pair
        # left swatch
        sh = slide.shapes.add_shape(
            1, x, y, sw_w, Inches(1.2)
        )
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor(*hex_to_rgb(a_hex))
        sh.line.fill.background()
        # label
        lbl = slide.shapes.add_textbox(x, y+Inches(1.25), sw_w, Inches(0.3)).text_frame
        lbl.paragraphs[0].text = a_hex
        lbl.paragraphs[0].font.size = Pt(10)
        # right swatch
        x2 = x + sw_w + gap
        sh2 = slide.shapes.add_shape(
            1, x2, y, sw_w, Inches(1.2)
        )
        sh2.fill.solid()
        sh2.fill.fore_color.rgb = RGBColor(*hex_to_rgb(b_hex))
        sh2.line.fill.background()
        lbl2 = slide.shapes.add_textbox(x2, y+Inches(1.25), sw_w, Inches(0.3)).text_frame
        lbl2.paragraphs[0].text = b_hex
        lbl2.paragraphs[0].font.size = Pt(10)
        x = x + (sw_w*2) + (gap*2)
    # Keywords
    kw = slide.shapes.add_textbox(Inches(0.5), Inches(3.3), Inches(9), Inches(1)).text_frame
    kw.paragraphs[0].text = "Keywords: Gender-fluid · Soft-tailoring · Crop tops · Romantic details · Confident · Playful · Modern Boho"
    kw.paragraphs[0].font.size = Pt(14)
    return slide

def add_image_grid_slide(prs, title, img_paths_labels, cols=3):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # Title
    t = slide.shapes.add_textbox(Inches(0.4), Inches(0.2), Inches(9), Inches(0.5)).text_frame
    t.paragraphs[0].text = title
    t.paragraphs[0].font.size = Pt(20)
    # Grid
    img_w = Inches(3)
    img_h = Inches(2)
    x0 = Inches(0.4)
    y0 = Inches(0.9)
    gap_x = Inches(0.3)
    gap_y = Inches(0.2)
    i = 0
    for label, path in img_paths_labels:
        col = i % cols
        row = i // cols
        x = x0 + col*(img_w + gap_x)
        y = y0 + row*(img_h + gap_y + Inches(0.25))
        # insert image
        try:
            slide.shapes.add_picture(path, x, y, width=img_w, height=img_h)
        except Exception as e:
            print("Image insert error:", path, e)
        # label
        lb = slide.shapes.add_textbox(x, y+img_h, img_w, Inches(0.25)).text_frame
        p = lb.paragraphs[0]
        p.text = label
        p.font.size = Pt(11)
        i += 1
    return slide

def add_customer_slide(prs, customer_text, design_direction, credits_text):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    t = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6)).text_frame
    t.paragraphs[0].text = "Customer & Story"
    t.paragraphs[0].font.size = Pt(22)
    # Customer details box
    cbox = slide.shapes.add_textbox(Inches(0.5), Inches(1.1), Inches(4.5), Inches(3)).text_frame
    cbox.paragraphs[0].text = customer_text
    cbox.paragraphs[0].font.size = Pt(12)
    # Design direction
    dbox = slide.shapes.add_textbox(Inches(5.2), Inches(1.1), Inches(4.0), Inches(2.5)).text_frame
    dbox.paragraphs[0].text = "Design Direction:\n" + design_direction
    dbox.paragraphs[0].font.size = Pt(12)
    # Credits (small)
    cred = slide.shapes.add_textbox(Inches(0.5), Inches(4.5), Inches(9), Inches(1)).text_frame
    cred.paragraphs[0].text = "Image sources (Unsplash source queries used): " + credits_text
    cred.paragraphs[0].font.size = Pt(8)
    return slide

def build_ppt(image_local_paths):
    prs = Presentation()
    # Blank default slide size OK
    # Choose palettes (complementary pairs)
    palettes = [
        ("#0B3D91", "#FF7F11"),  # Navy ↔ Orange
        ("#2AA198", "#FF6B6B")   # Teal ↔ Coral
    ]
    # Slide 1: Title
    title = "Male Feminine Men — Modern Boho"
    subtitle = "Gender‑fluid fashion for men: crop tops, soft tailoring, romantic details"
    add_title_slide(prs, title, subtitle, bg_hex="#0B3D91")
    # Slide 2: Palette
    add_palette_slide(prs, palettes)
    # Slide 3: Fabrics & textures
    fabrics = [
        ("Lace & Sheer", image_local_paths.get("lace_texture")),
        ("Soft Silk", image_local_paths.get("silk_texture")),
        ("Boho Prints", image_local_paths.get("boho_print")),
        ("Pastel Interiors (lifestyle)", image_local_paths.get("pastel_interior")),
        ("Textile Detail (suit lapel close-up)", image_local_paths.get("suit_feminine_1")),
        ("Texture: Crochet/Knits", image_local_paths.get("crop_top_1"))
    ]
    add_image_grid_slide(prs, "Fabrics, Textures & Prints", fabrics, cols=3)
    # Slide 4: Outfits & Silhouettes
    outfits = [
        ("Crop Top Styling", image_local_paths.get("crop_top_1")),
        ("Feminine Tailored Suit", image_local_paths.get("suit_feminine_2")),
        ("Street Styling", image_local_paths.get("street_style")),
        ("Runway Inspiration", image_local_paths.get("runway")),
        ("Soft Draping & Layers", image_local_paths.get("crop_top_2")),
        ("Suit with Delicate Details", image_local_paths.get("suit_feminine_1"))
    ]
    add_image_grid_slide(prs, "Outfit & Silhouette Inspiration", outfits, cols=3)
    # Slide 5: Customer & Story + Credits
    customer_text = (
        "Target customer:\n"
        "- Age: 20–35\n"
        "- Lifestyle: Urban, creative professionals or students\n"
        "- Fashion preference: Gender‑fluid, experimental, values craft and comfort\n"
        "- Occasion: Daytime socials, creative workplaces, premium casual events\n"
        "- Buying mindset: Seeks distinctive pieces, willing to invest in statement tailoring and versatile separates."
    )
    design_direction = (
        "Designs balance soft, romantic fabrics with structured tailoring. "
        "Colour pairs use bold complementary contrasts to make feminine details pop against classic menswear shapes. "
        "Silhouettes mix cropped proportions with conservative suit elements to normalize feminine expression for men."
    )
    credits_text = ", ".join([f"{q} (Unsplash source)" for q in image_queries.values()])
    add_customer_slide(prs, customer_text, design_direction, credits_text)
    prs.save(OUTFILE)
    print("Saved PPTX as", OUTFILE)

if __name__ == "__main__":
    print("Downloading images...")
    paths = download_images()
    print("Building PPTX...")
    build_ppt(paths)
    print("Done. Open", OUTFILE)
