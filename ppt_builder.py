"""PowerPoint Presentation Builder using python-pptx.

Creates modern, 16:9 widescreen presentations with distinct themes,
professional typography, card-based layouts, and speaker notes.
"""

import os
from pathlib import Path
from typing import Any, Dict, List, Optional
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE


THEMES = {
    "modern_indigo": {
        "name": "Modern Indigo",
        "description": "Deep indigo & vibrant electric blue with cyan accents",
        "primary": "1E1B4B",
        "secondary": "4F46E5",
        "accent": "06B6D4",
        "bg_light": "F8FAFC",
        "bg_dark": "0F172A",
        "card_bg": "FFFFFF",
        "card_border": "E2E8F0",
        "text_primary": "0F172A",
        "text_secondary": "475569",
        "text_light": "FFFFFF",
        "text_muted": "94A3B8",
    },
    "tech_emerald": {
        "name": "Tech Emerald",
        "description": "Deep forest green & luminous emerald with mint accents",
        "primary": "064E3B",
        "secondary": "059669",
        "accent": "10B981",
        "bg_light": "F0FDF4",
        "bg_dark": "022C22",
        "card_bg": "FFFFFF",
        "card_border": "DCFCE7",
        "text_primary": "064E3B",
        "text_secondary": "166534",
        "text_light": "FFFFFF",
        "text_muted": "6EE7B7",
    },
    "executive_slate": {
        "name": "Executive Slate",
        "description": "Premium slate grey, royal corporate blue & gold accents",
        "primary": "0F172A",
        "secondary": "1E40AF",
        "accent": "F59E0B",
        "bg_light": "F8FAFC",
        "bg_dark": "0F172A",
        "card_bg": "FFFFFF",
        "card_border": "CBD5E1",
        "text_primary": "0F172A",
        "text_secondary": "334155",
        "text_light": "FFFFFF",
        "text_muted": "94A3B8",
    },
    "crimson_bold": {
        "name": "Crimson Bold",
        "description": "Vibrant ruby crimson, warm rose & coral highlights",
        "primary": "881337",
        "secondary": "E11D48",
        "accent": "FB7185",
        "bg_light": "FFF1F2",
        "bg_dark": "4C0519",
        "card_bg": "FFFFFF",
        "card_border": "FFE4E6",
        "text_primary": "881337",
        "text_secondary": "9F1239",
        "text_light": "FFFFFF",
        "text_muted": "FDA4AF",
    },
    "cyber_dark": {
        "name": "Cyber Dark",
        "description": "Sleek dark mode with electric purple and cyan glow",
        "primary": "8B5CF6",
        "secondary": "6366F1",
        "accent": "06B6D4",
        "bg_light": "18181B",
        "bg_dark": "09090B",
        "card_bg": "27272A",
        "card_border": "3F3F46",
        "text_primary": "FAFAFA",
        "text_secondary": "D4D4D8",
        "text_light": "FFFFFF",
        "text_muted": "A1A1AA",
    },
}


def hex_to_rgb(hex_code: str) -> RGBColor:
    """Converts a hex color string like '1E1B4B' or '#1E1B4B' to RGBColor."""
    hex_clean = hex_code.lstrip("#")
    if len(hex_clean) != 6:
        hex_clean = "4F46E5"
    r = int(hex_clean[0:2], 16)
    g = int(hex_clean[2:4], 16)
    b = int(hex_clean[4:6], 16)
    return RGBColor(r, g, b)


class PowerPointBuilder:
    def __init__(self, theme_name: str = "modern_indigo"):
        self.theme_name = theme_name if theme_name in THEMES else "modern_indigo"
        self.theme = THEMES[self.theme_name]
        self.prs = Presentation()
        # Set 16:9 widescreen dimensions
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        self.blank_layout = self.prs.slide_layouts[6]
        self.font_main = "Calibri"
        self.font_heading = "Segoe UI"

    def _set_slide_background(self, slide, is_dark: bool = False):
        """Sets full slide background fill."""
        bg_color_hex = self.theme["bg_dark"] if is_dark else self.theme["bg_light"]
        bg = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5)
        )
        bg.fill.solid()
        bg.fill.fore_color.rgb = hex_to_rgb(bg_color_hex)
        bg.line.fill.background()  # No border

    def _add_header(self, slide, title: str, category: Optional[str] = None, subtitle: Optional[str] = None):
        """Adds a consistent professional header to content slides."""
        # Optional category / kicker tag
        if category:
            cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.45), Inches(11.7), Inches(0.35))
            tf_cat = cat_box.text_frame
            tf_cat.word_wrap = True
            p_cat = tf_cat.paragraphs[0]
            p_cat.text = category.upper()
            p_cat.font.name = self.font_heading
            p_cat.font.size = Pt(10)
            p_cat.font.bold = True
            p_cat.font.color.rgb = hex_to_rgb(self.theme["secondary"])

        # Main slide title
        top_offset = Inches(0.75) if category else Inches(0.55)
        title_box = slide.shapes.add_textbox(Inches(0.8), top_offset, Inches(11.7), Inches(0.8))
        tf = title_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = self.font_heading
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgb(self.theme["text_primary"])

        # Subtitle / context if available
        if subtitle:
            sub_box = slide.shapes.add_textbox(Inches(0.8), top_offset + Inches(0.55), Inches(11.7), Inches(0.4))
            tf_sub = sub_box.text_frame
            tf_sub.word_wrap = True
            p_sub = tf_sub.paragraphs[0]
            p_sub.text = subtitle
            p_sub.font.name = self.font_main
            p_sub.font.size = Pt(13)
            p_sub.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])

        # Subtle accent divider line
        line_top = top_offset + (Inches(0.95) if subtitle else Inches(0.7))
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), line_top, Inches(2.5), Inches(0.04))
        line.fill.solid()
        line.fill.fore_color.rgb = hex_to_rgb(self.theme["accent"])
        line.line.fill.background()

    def _add_footer(self, slide, current_idx: int, total_slides: int):
        """Adds slide number and subtle brand accent."""
        foot_box = slide.shapes.add_textbox(Inches(11.5), Inches(6.9), Inches(1.0), Inches(0.35))
        tf = foot_box.text_frame
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT
        p.text = f"{current_idx} / {total_slides}"
        p.font.name = self.font_main
        p.font.size = Pt(10)
        p.font.color.rgb = hex_to_rgb(self.theme["text_muted"])

    def _set_speaker_notes(self, slide, notes: str):
        """Sets speaker notes for the slide."""
        if notes:
            notes_slide = slide.notes_slide
            tf = notes_slide.notes_text_frame
            tf.text = notes

    def create_title_slide(self, title: str, subtitle: str = "", presenter: str = "", date_str: str = "", notes: str = ""):
        """Creates a high-impact title/hero slide."""
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._set_slide_background(slide, is_dark=True)

        # Right geometric accent banner/panel
        accent_shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(8.5), Inches(0), Inches(4.833), Inches(7.5)
        )
        accent_shape.fill.solid()
        accent_shape.fill.fore_color.rgb = hex_to_rgb(self.theme["primary"])
        accent_shape.line.fill.background()

        # Accent decorative bar
        dec_bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(8.35), Inches(0), Inches(0.15), Inches(7.5)
        )
        dec_bar.fill.solid()
        dec_bar.fill.fore_color.rgb = hex_to_rgb(self.theme["accent"])
        dec_bar.line.fill.background()

        # Category pill badge
        pill = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(1.6), Inches(2.6), Inches(0.42)
        )
        pill.fill.solid()
        pill.fill.fore_color.rgb = hex_to_rgb(self.theme["secondary"])
        pill.line.fill.background()
        tf_pill = pill.text_frame
        p_pill = tf_pill.paragraphs[0]
        p_pill.alignment = PP_ALIGN.CENTER
        p_pill.text = "PRESENTATION"
        p_pill.font.name = self.font_heading
        p_pill.font.size = Pt(11)
        p_pill.font.bold = True
        p_pill.font.color.rgb = hex_to_rgb(self.theme["text_light"])

        # Main Title Box
        title_box = slide.shapes.add_textbox(Inches(1.2), Inches(2.2), Inches(6.8), Inches(2.4))
        tf = title_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = self.font_heading
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgb(self.theme["text_light"])

        # Subtitle
        if subtitle:
            sub_box = slide.shapes.add_textbox(Inches(1.2), Inches(4.7), Inches(6.8), Inches(1.2))
            tf_sub = sub_box.text_frame
            tf_sub.word_wrap = True
            p_sub = tf_sub.paragraphs[0]
            p_sub.text = subtitle
            p_sub.font.name = self.font_main
            p_sub.font.size = Pt(16)
            p_sub.font.color.rgb = hex_to_rgb(self.theme["text_muted"])

        # Presenter / Info in right card
        if presenter or date_str:
            info_box = slide.shapes.add_textbox(Inches(9.0), Inches(3.0), Inches(3.8), Inches(2.0))
            tf_info = info_box.text_frame
            tf_info.word_wrap = True
            if presenter:
                p_pres = tf_info.paragraphs[0]
                p_pres.text = f"Presented By:\n{presenter}"
                p_pres.font.name = self.font_heading
                p_pres.font.size = Pt(14)
                p_pres.font.bold = True
                p_pres.font.color.rgb = hex_to_rgb(self.theme["text_light"])
            if date_str:
                p_date = tf_info.add_paragraph()
                p_date.text = f"\nDate:\n{date_str}"
                p_date.font.name = self.font_main
                p_date.font.size = Pt(12)
                p_date.font.color.rgb = hex_to_rgb(self.theme["text_muted"])

        self._set_speaker_notes(slide, notes)
        return slide

    def create_agenda_slide(self, title: str, items: List[Dict[str, str]], subtitle: str = "", notes: str = "", slide_idx: int = 2, total_slides: int = 5):
        """Creates an agenda or roadmap slide."""
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._set_slide_background(slide)
        self._add_header(slide, title, category="OVERVIEW", subtitle=subtitle)

        count = min(len(items), 5)
        if count == 0:
            return slide

        col_width = (Inches(11.7) - Inches(0.3) * (count - 1)) / count
        start_left = Inches(0.8)
        top = Inches(2.3)
        height = Inches(4.2)

        for i, item in enumerate(items[:count]):
            left = start_left + i * (col_width + Inches(0.3))
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, col_width, height)
            card.fill.solid()
            card.fill.fore_color.rgb = hex_to_rgb(self.theme["card_bg"])
            card.line.color.rgb = hex_to_rgb(self.theme["card_border"])
            card.line.width = Pt(1)

            # Card Header Pill with Step Number
            num_pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.2), top + Inches(0.25), Inches(0.7), Inches(0.4))
            num_pill.fill.solid()
            num_pill.fill.fore_color.rgb = hex_to_rgb(self.theme["secondary"])
            num_pill.line.fill.background()
            tf_num = num_pill.text_frame
            p_num = tf_num.paragraphs[0]
            p_num.alignment = PP_ALIGN.CENTER
            p_num.text = f"0{i+1}"
            p_num.font.bold = True
            p_num.font.size = Pt(12)
            p_num.font.color.rgb = hex_to_rgb(self.theme["text_light"])

            # Item Title
            item_title = item.get("title", f"Section {i+1}")
            t_box = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.8), col_width - Inches(0.4), Inches(1.0))
            tf_t = t_box.text_frame
            tf_t.word_wrap = True
            p_t = tf_t.paragraphs[0]
            p_t.text = item_title
            p_t.font.name = self.font_heading
            p_t.font.size = Pt(16)
            p_t.font.bold = True
            p_t.font.color.rgb = hex_to_rgb(self.theme["text_primary"])

            # Item Description
            desc = item.get("description", "")
            if desc:
                d_box = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(1.8), col_width - Inches(0.4), Inches(2.1))
                tf_d = d_box.text_frame
                tf_d.word_wrap = True
                p_d = tf_d.paragraphs[0]
                p_d.text = desc
                p_d.font.name = self.font_main
                p_d.font.size = Pt(12)
                p_d.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])

        self._add_footer(slide, slide_idx, total_slides)
        self._set_speaker_notes(slide, notes)
        return slide

    def create_cards_slide(self, title: str, cards: List[Dict[str, str]], category: str = "", subtitle: str = "", notes: str = "", slide_idx: int = 3, total_slides: int = 5):
        """Creates a 3-card or 4-card metric / key takeaway slide."""
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._set_slide_background(slide)
        self._add_header(slide, title, category=category or "INSIGHTS", subtitle=subtitle)

        count = min(max(len(cards), 1), 4)
        col_width = (Inches(11.7) - Inches(0.4) * (count - 1)) / count
        start_left = Inches(0.8)
        top = Inches(2.2)
        height = Inches(4.3)

        for i, card_info in enumerate(cards[:count]):
            left = start_left + i * (col_width + Inches(0.4))
            
            # Container card
            bg_card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, col_width, height)
            bg_card.fill.solid()
            bg_card.fill.fore_color.rgb = hex_to_rgb(self.theme["card_bg"])
            bg_card.line.color.rgb = hex_to_rgb(self.theme["card_border"])
            bg_card.line.width = Pt(1)

            # Top accent bar on each card
            top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left + Inches(0.2), top + Inches(0.15), Inches(0.8), Inches(0.06))
            top_bar.fill.solid()
            top_bar.fill.fore_color.rgb = hex_to_rgb(self.theme["accent"])
            top_bar.line.fill.background()

            # Metric / Stat (if provided)
            metric = card_info.get("metric", "")
            text_start_top = top + Inches(0.35)
            if metric:
                m_box = slide.shapes.add_textbox(left + Inches(0.25), text_start_top, col_width - Inches(0.5), Inches(0.8))
                tf_m = m_box.text_frame
                tf_m.word_wrap = True
                p_m = tf_m.paragraphs[0]
                p_m.text = metric
                p_m.font.name = self.font_heading
                p_m.font.size = Pt(28)
                p_m.font.bold = True
                p_m.font.color.rgb = hex_to_rgb(self.theme["secondary"])
                text_start_top += Inches(0.8)

            # Card Title
            card_title = card_info.get("title", f"Point {i+1}")
            t_box = slide.shapes.add_textbox(left + Inches(0.25), text_start_top, col_width - Inches(0.5), Inches(0.8))
            tf_t = t_box.text_frame
            tf_t.word_wrap = True
            p_t = tf_t.paragraphs[0]
            p_t.text = card_title
            p_t.font.name = self.font_heading
            p_t.font.size = Pt(16)
            p_t.font.bold = True
            p_t.font.color.rgb = hex_to_rgb(self.theme["text_primary"])

            # Card Description
            card_desc = card_info.get("description", "")
            if card_desc:
                d_box = slide.shapes.add_textbox(left + Inches(0.25), text_start_top + Inches(0.75), col_width - Inches(0.5), Inches(2.2))
                tf_d = d_box.text_frame
                tf_d.word_wrap = True
                p_d = tf_d.paragraphs[0]
                p_d.text = card_desc
                p_d.font.name = self.font_main
                p_d.font.size = Pt(13)
                p_d.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])

        self._add_footer(slide, slide_idx, total_slides)
        self._set_speaker_notes(slide, notes)
        return slide

    def create_two_column_slide(self, title: str, col1: Dict[str, Any], col2: Dict[str, Any], category: str = "", subtitle: str = "", notes: str = "", slide_idx: int = 4, total_slides: int = 5):
        """Creates a side-by-side comparison or 2-column breakdown slide."""
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._set_slide_background(slide)
        self._add_header(slide, title, category=category or "COMPARISON", subtitle=subtitle)

        width = Inches(5.6)
        top = Inches(2.1)
        height = Inches(4.5)

        for idx, (col_data, left_pos) in enumerate([(col1, Inches(0.8)), (col2, Inches(6.9))]):
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_pos, top, width, height)
            card.fill.solid()
            card.fill.fore_color.rgb = hex_to_rgb(self.theme["card_bg"])
            card.line.color.rgb = hex_to_rgb(self.theme["card_border"])
            card.line.width = Pt(1)

            # Column Title Header
            col_title = col_data.get("title", f"Perspective {idx+1}")
            h_box = slide.shapes.add_textbox(left_pos + Inches(0.4), top + Inches(0.3), width - Inches(0.8), Inches(0.6))
            tf_h = h_box.text_frame
            tf_h.word_wrap = True
            p_h = tf_h.paragraphs[0]
            p_h.text = col_title
            p_h.font.name = self.font_heading
            p_h.font.size = Pt(18)
            p_h.font.bold = True
            p_h.font.color.rgb = hex_to_rgb(self.theme["secondary"] if idx == 0 else self.theme["primary"])

            # Column Points / Items
            items = col_data.get("points", [])
            if isinstance(items, str):
                items = [items]

            points_box = slide.shapes.add_textbox(left_pos + Inches(0.4), top + Inches(1.0), width - Inches(0.8), Inches(3.2))
            tf_p = points_box.text_frame
            tf_p.word_wrap = True
            for p_idx, pt_text in enumerate(items[:5]):
                p_item = tf_p.paragraphs[0] if p_idx == 0 else tf_p.add_paragraph()
                p_item.text = f"•  {pt_text}"
                p_item.font.name = self.font_main
                p_item.font.size = Pt(13)
                p_item.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])
                p_item.space_after = Pt(10)

        self._add_footer(slide, slide_idx, total_slides)
        self._set_speaker_notes(slide, notes)
        return slide

    def create_bullets_slide(self, title: str, points: List[Dict[str, str]], category: str = "", subtitle: str = "", notes: str = "", slide_idx: int = 5, total_slides: int = 5):
        """Creates a structured key takeaways / bullet points slide."""
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._set_slide_background(slide)
        self._add_header(slide, title, category=category or "KEY POINTS", subtitle=subtitle)

        count = min(len(points), 4)
        if count == 0:
            return slide

        top_start = Inches(2.2)
        row_height = (Inches(4.4) - Inches(0.2) * (count - 1)) / count

        for i, pt_info in enumerate(points[:count]):
            y_pos = top_start + i * (row_height + Inches(0.2))

            # Row container card
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), y_pos, Inches(11.7), row_height)
            card.fill.solid()
            card.fill.fore_color.rgb = hex_to_rgb(self.theme["card_bg"])
            card.line.color.rgb = hex_to_rgb(self.theme["card_border"])
            card.line.width = Pt(1)

            # Colored accent left pill
            pill = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), y_pos, Inches(0.12), row_height)
            pill.fill.solid()
            pill.fill.fore_color.rgb = hex_to_rgb(self.theme["secondary"] if i % 2 == 0 else self.theme["accent"])
            pill.line.fill.background()

            # Heading & Text in the card
            t_box = slide.shapes.add_textbox(Inches(1.2), y_pos + Inches(0.15), Inches(11.0), row_height - Inches(0.3))
            tf = t_box.text_frame
            tf.word_wrap = True

            pt_title = pt_info.get("title", f"Point {i+1}")
            pt_desc = pt_info.get("description", "")

            p1 = tf.paragraphs[0]
            p1.text = pt_title
            p1.font.name = self.font_heading
            p1.font.size = Pt(15)
            p1.font.bold = True
            p1.font.color.rgb = hex_to_rgb(self.theme["text_primary"])

            if pt_desc:
                p2 = tf.add_paragraph()
                p2.text = pt_desc
                p2.font.name = self.font_main
                p2.font.size = Pt(12)
                p2.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])

        self._add_footer(slide, slide_idx, total_slides)
        self._set_speaker_notes(slide, notes)
        return slide

    def create_conclusion_slide(self, title: str, takeaways: List[str], call_to_action: str = "", notes: str = "", slide_idx: int = 5, total_slides: int = 5):
        """Creates a conclusion, next steps, and Q&A slide."""
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._set_slide_background(slide)
        self._add_header(slide, title, category="CONCLUSION", subtitle="Summary & Path Forward")

        # Left Container: Key Takeaways
        left_card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.2), Inches(7.0), Inches(4.3))
        left_card.fill.solid()
        left_card.fill.fore_color.rgb = hex_to_rgb(self.theme["card_bg"])
        left_card.line.color.rgb = hex_to_rgb(self.theme["card_border"])
        left_card.line.width = Pt(1)

        t_box = slide.shapes.add_textbox(Inches(1.2), Inches(2.5), Inches(6.2), Inches(3.7))
        tf = t_box.text_frame
        tf.word_wrap = True
        p_head = tf.paragraphs[0]
        p_head.text = "Key Takeaways"
        p_head.font.name = self.font_heading
        p_head.font.size = Pt(18)
        p_head.font.bold = True
        p_head.font.color.rgb = hex_to_rgb(self.theme["primary"])
        p_head.space_after = Pt(12)

        for takeaway in takeaways[:4]:
            p_item = tf.add_paragraph()
            p_item.text = f"✔  {takeaway}"
            p_item.font.name = self.font_main
            p_item.font.size = Pt(13)
            p_item.font.color.rgb = hex_to_rgb(self.theme["text_secondary"])
            p_item.space_after = Pt(10)

        # Right Container: Call To Action / Q&A
        right_card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.2), Inches(2.2), Inches(4.3), Inches(4.3))
        right_card.fill.solid()
        right_card.fill.fore_color.rgb = hex_to_rgb(self.theme["bg_dark"])
        right_card.line.fill.background()

        cta_box = slide.shapes.add_textbox(Inches(8.6), Inches(2.7), Inches(3.5), Inches(3.3))
        tf_cta = cta_box.text_frame
        tf_cta.word_wrap = True
        p_cta_head = tf_cta.paragraphs[0]
        p_cta_head.text = "Next Steps & Q&A"
        p_cta_head.font.name = self.font_heading
        p_cta_head.font.size = Pt(20)
        p_cta_head.font.bold = True
        p_cta_head.font.color.rgb = hex_to_rgb(self.theme["accent"])
        p_cta_head.space_after = Pt(14)

        if call_to_action:
            p_cta = tf_cta.add_paragraph()
            p_cta.text = call_to_action
            p_cta.font.name = self.font_main
            p_cta.font.size = Pt(13)
            p_cta.font.color.rgb = hex_to_rgb(self.theme["text_light"])
            p_cta.space_after = Pt(14)

        p_thanks = tf_cta.add_paragraph()
        p_thanks.text = "Thank you! Questions or comments are welcome."
        p_thanks.font.name = self.font_heading
        p_thanks.font.size = Pt(13)
        p_thanks.font.bold = True
        p_thanks.font.color.rgb = hex_to_rgb(self.theme["text_muted"])

        self._add_footer(slide, slide_idx, total_slides)
        self._set_speaker_notes(slide, notes)
        return slide

    def build_from_spec(self, presentation_data: Dict[str, Any], output_path: str) -> str:
        """Builds a complete presentation from a structured dictionary spec."""
        title = presentation_data.get("title", "Executive Presentation")
        subtitle = presentation_data.get("subtitle", "")
        presenter = presentation_data.get("presenter", "Google Antigravity AI")
        date_str = presentation_data.get("date", "")
        slides_spec = presentation_data.get("slides", [])

        # Total slides count includes title slide
        total_slides = len(slides_spec) + 1

        # 1. Title Slide
        title_notes = presentation_data.get("title_notes", "Welcome to this presentation.")
        self.create_title_slide(title, subtitle, presenter, date_str, notes=title_notes)

        # 2. Content Slides
        for idx, slide_info in enumerate(slides_spec, start=2):
            slide_type = slide_info.get("type", "bullets").lower()
            s_title = slide_info.get("title", f"Slide {idx}")
            s_subtitle = slide_info.get("subtitle", "")
            s_category = slide_info.get("category", "")
            s_notes = slide_info.get("speaker_notes", "")

            if slide_type in ("agenda", "overview"):
                items = slide_info.get("items", [])
                self.create_agenda_slide(s_title, items, s_subtitle, notes=s_notes, slide_idx=idx, total_slides=total_slides)

            elif slide_type in ("cards", "metrics", "stats"):
                cards = slide_info.get("cards", [])
                self.create_cards_slide(s_title, cards, category=s_category, subtitle=s_subtitle, notes=s_notes, slide_idx=idx, total_slides=total_slides)

            elif slide_type in ("two_column", "comparison"):
                col1 = slide_info.get("col1", {})
                col2 = slide_info.get("col2", {})
                self.create_two_column_slide(s_title, col1, col2, category=s_category, subtitle=s_subtitle, notes=s_notes, slide_idx=idx, total_slides=total_slides)

            elif slide_type in ("conclusion", "summary"):
                takeaways = slide_info.get("takeaways", [])
                cta = slide_info.get("call_to_action", "")
                self.create_conclusion_slide(s_title, takeaways, call_to_action=cta, notes=s_notes, slide_idx=idx, total_slides=total_slides)

            else:  # bullets / default
                points = slide_info.get("points", [])
                # If points is a list of strings, convert to dicts
                if points and isinstance(points[0], str):
                    points = [{"title": pt, "description": ""} for pt in points]
                self.create_bullets_slide(s_title, points, category=s_category, subtitle=s_subtitle, notes=s_notes, slide_idx=idx, total_slides=total_slides)

        # Ensure output directory exists
        out_file = Path(output_path).resolve()
        out_file.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(str(out_file))
        return str(out_file)

