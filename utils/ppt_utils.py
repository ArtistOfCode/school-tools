from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_VERTICAL_ANCHOR, PP_ALIGN
from pptx.slide import Slide


def add_layout_slide(ppt, layout, title) -> Slide:
    _layout = ppt.slide_layouts[layout]
    _slide: Slide = ppt.slides.add_slide(_layout)
    _slide.shapes.title.text = title
    return _slide


def set_center_cell(cell, value: str, color=None):
    cell.text = value
    cell.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    if color is not None:
        cell.text_frame.paragraphs[0].font.color.rgb = RGBColor.from_string(color)
