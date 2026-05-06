import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor

# --- LAYER 1: PARSER ---
class PPTXParser:
    def extract(self, pptx_file):
        """Extracts raw content into a structured intermediate representation."""
        prs = Presentation(pptx_file)
        model = {
            "width": prs.slide_width,
            "height": prs.slide_height,
            "slides": []
        }

        for slide in prs.slides:
            slide_data = {
                "title": "",
                "elements": []
            }

            if slide.shapes.title:
                slide_data["title"] = slide.shapes.title.text

            for shape in slide.shapes:
                # Text Elements
                if shape.has_text_frame:
                    if shape == slide.shapes.title:
                        continue
                    text = shape.text.strip()
                    if text:
                        slide_data["elements"].append({
                            "type": "text",
                            "content": text,
                            "bbox": {
                                "x": shape.left,
                                "y": shape.top,
                                "w": shape.width,
                                "h": shape.height
                            }
                        })

                # Image Elements
                elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    try:
                        slide_data["elements"].append({
                            "type": "image",
                            "source": io.BytesIO(shape.image.blob),
                            "bbox": {
                                "x": shape.left,
                                "y": shape.top,
                                "w": shape.width,
                                "h": shape.height
                            }
                        })
                    except Exception as e:
                        print(f"Parser error extracting image: {e}")

            model["slides"].append(slide_data)
        return model

# --- LAYER 2: LAYOUT PREPROCESSOR ---
# --- LAYER 2: INITIAL PLACEMENT ---
class InitialPlacement:
    def __init__(self):
        self.margin = Inches(0.5)
        self.title_height = Inches(0.8)

    def intersects(self, r1, r2):
        """Helper to detect intersection between two rectangles (x, y, w, h)."""
        return not (r1[0] + r1[2] <= r2["x"] or r1[0] >= r2["x"] + r2["w"] or
                    r1[1] + r1[3] <= r2["y"] or r1[1] >= r2["y"] + r2["h"])

    def apply(self, model, ai_summarizer=None):
        """Sets up a deterministic 50/50 horizontal split layout (Text Left, Image Right)."""
        processed_slides = []
        margin = Inches(0.5)
        safe_l, safe_t = margin, margin
        safe_r, safe_b = model["width"] - margin, model["height"] - margin
        
        usable_w = safe_r - safe_l
        usable_h = safe_b - safe_t - self.title_height - Inches(0.2)
        
        # Define Zones (Horizontal Split)
        zone_w = usable_w / 2.0
        text_zone_x = safe_l
        img_zone_x = safe_l + zone_w + Inches(0.1) # small gap
        content_y = safe_t + self.title_height + Inches(0.2)

        for slide_data in model["slides"]:
            new_elements = []
            
            # 0. Asset Extraction (Always defined)
            images = [e for e in slide_data["elements"] if e["type"] == "image"]
            text_blocks = [e for e in slide_data["elements"] if e["type"] == "text"]
            # Filter out NF/UF prefixes
            text_blocks = [t for t in text_blocks if not (t["content"].strip().startswith("NF") or t["content"].strip().startswith("UF"))]

            # 1. Title
            title_text = slide_data["title"]
            title_bbox = {"x": safe_l, "y": safe_t, "w": usable_w, "h": self.title_height}
            
            # 2. Layout Mode Detection
            is_image_only = len(text_blocks) == 0 and len(images) > 0
            
            # Layout Zone Parameters
            if is_image_only:
                curr_img_zone_x = safe_l
                curr_zone_w = usable_w
            else:
                curr_img_zone_x = img_zone_x
                curr_zone_w = zone_w

            # 2. Images (GRID LAYOUT)
            if images:
                num_images = len(images)
                cols = 2 if num_images > 1 else 1
                import math
                rows = math.ceil(num_images / cols)
                
                cell_w = curr_zone_w / cols
                cell_h = usable_h / rows
                
                for idx, img in enumerate(images):
                    row = idx // cols
                    col = idx % cols
                    
                    cx = curr_img_zone_x + (col * cell_w)
                    cy = content_y + (row * cell_h)
                    
                    l, t, w, h = img["bbox"]["x"], img["bbox"]["y"], img["bbox"]["w"], img["bbox"]["h"]
                    
                    ratio = min((cell_w - Inches(0.2)) / w, (cell_h - Inches(0.2)) / h)
                    w_final, h_final = int(w * ratio), int(h * ratio)
                    
                    ix = cx + (cell_w - w_final) / 2.0
                    iy = cy + (cell_h - h_final) / 2.0
                    
                    new_elements.append({
                        "type": "image",
                        "source": img["source"],
                        "bbox": {"x": ix, "y": iy, "w": w_final, "h": h_final}
                    })

            # 3. Text (Locked to LEFT ZONE)
            if not is_image_only:
                full_text = "\n".join([t["content"] for t in text_blocks])
                bullets = ai_summarizer(full_text) if ai_summarizer else [line.strip() for line in full_text.split('\n') if line.strip()][:5]

                new_elements.append({
                    "type": "bullet_group",
                    "content": bullets,
                    "font_size": 18,
                    "bbox": {"x": text_zone_x, "y": content_y, "w": zone_w - Inches(0.1), "h": usable_h}
                })

            processed_slides.append({
                "title": title_text,
                "title_bbox": title_bbox,
                "elements": new_elements
            })
        
        return {
            "width": model["width"],
            "height": model["height"],
            "slides": processed_slides
        }

# --- LAYER 3: RENDERER ---
class PPTXRenderer:
    def __init__(self):
        self.accent_color = RGBColor(255, 107, 0)
        self.text_color = RGBColor(0, 0, 0)

    def render(self, processed_model):
        """Translates final geometry into python-pptx objects. NO logic allowed here."""
        prs = Presentation()
        prs.slide_width = processed_model["width"]
        prs.slide_height = processed_model["height"]

        for slide_data in processed_model["slides"]:
            slide = prs.slides.add_slide(prs.slide_layouts[6]) # Blank
            
            # Render Title
            if slide_data["title"]:
                bb = slide_data["title_bbox"]
                box = slide.shapes.add_textbox(bb["x"], bb["y"], bb["w"], bb["h"])
                tf = box.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.text = slide_data["title"].upper()
                p.font.bold = True
                p.font.size = Pt(28)
                p.font.color.rgb = self.accent_color
                p.font.name = 'Calibri'

            # Render Elements
            for el in slide_data["elements"]:
                bb = el["bbox"]
                if el["type"] == "image":
                    slide.shapes.add_picture(el["source"], bb["x"], bb["y"], width=bb["w"], height=bb["h"])
                
                elif el["type"] == "bullet_group":
                    box = slide.shapes.add_textbox(bb["x"], bb["y"], bb["w"], bb["h"])
                    tf = box.text_frame
                    tf.word_wrap = True
                    f_size = el.get("font_size", 18)
                    for i, bullet in enumerate(el["content"]):
                        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                        p.text = f"• {bullet}"
                        p.font.size = Pt(f_size)
                        p.font.color.rgb = self.text_color
                        p.font.name = 'Calibri'
                        p.space_after = Pt(8)

        output = io.BytesIO()
        prs.save(output)
        output.seek(0)
        return output

# --- VALIDATION LAYER ---
def validate_layout(model):
    """Ensures all elements are valid and within bounds before rendering."""
    margin = Inches(0.4) # Slightly more tolerant for check
    for slide_idx, slide in enumerate(model["slides"]):
        # Check title
        if not all(k in slide["title_bbox"] for k in ["x", "y", "w", "h"]):
            return False, f"Slide {slide_idx}: Missing title bbox"
        
        # Check elements
        for el in slide["elements"]:
            bb = el["bbox"]
            if bb["x"] < 0 or bb["y"] < 0 or (bb["x"] + bb["w"]) > model["width"] + 100 or (bb["y"] + bb["h"]) > model["height"] + 100:
                return False, f"Slide {slide_idx}: Element {el['type']} out of bounds"
            
    return True, "Success"
