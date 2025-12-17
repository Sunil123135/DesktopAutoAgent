import sys
import json
import shutil
import time
import os
from pathlib import Path
from typing import Any, Dict, List, Optional

import win32com.client
from win32com.client import constants

WORK_DIR = Path(r"C:\Users\slalwani\OneDrive - QuidelOrtho\SUNIL\EAG\Session 13\Code\temp")
CM_TO_POINTS = 28.3465  # Conversion factor


def load_instructions(json_path: Path) -> Dict[str, Any]:
    with json_path.open("r", encoding="utf-8") as f:
        return json.load(f)


def backup_pptx(file_path: Path) -> Optional[Path]:
    backup_path = file_path.with_name(file_path.stem + "_backup" + file_path.suffix)
    try:
        shutil.copy2(file_path, backup_path)
        return backup_path
    except PermissionError:
        print(f"  Warning: Could not create backup (file may be locked). Continuing without backup.")
        return None
    except Exception as e:
        print(f"  Warning: Backup failed: {e}. Continuing without backup.")
        return None


def log_entry(
    log: List[Dict[str, Any]],
    op_id: Any,
    action: str,
    slide_index: Optional[int],
    success: bool,
    message: str,
    file_name: Optional[str] = None,
    iteration: Optional[int] = None,
):
    log.append(
        {
            "op_id": op_id,
            "action": action,
            "slide_index": slide_index,
            "success": success,
            "message": message,
            "file": file_name,
            "iteration": iteration,
        }
    )


def ensure_powerpoint():
    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        ppt.Visible = True
        return ppt
    except Exception as e:
        raise Exception(f"PowerPoint not available: {e}")


def open_presentation(ppt_app, path: Path):
    try:
        pres = ppt_app.Presentations.Open(str(path), WithWindow=True)
        return pres
    except Exception as e:
        raise Exception(f"Failed to open presentation: {e}")


def get_layout_index(pres, layout_name: str) -> int:
    """Get layout index by name"""
    layout_map = {
        "Title": 1,
        "Title and Content": 2,
        "Blank": 7,
    }
    # Try common indices
    for idx in range(1, pres.SlideMaster.CustomLayouts.Count + 1):
        try:
            layout = pres.SlideMaster.CustomLayouts(idx)
            if layout_name.lower() in layout.Name.lower() or idx == layout_map.get(layout_name, 2):
                return idx
        except:
            continue
    return layout_map.get(layout_name, 2)  # Default to Title and Content


# ---- ALL OPERATIONS IMPLEMENTED ----

def op_add_slide(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    layout_name = op.get("layout", "Title and Content")
    try:
        layout_idx = get_layout_index(pres, layout_name)
        layout = pres.SlideMaster.CustomLayouts(layout_idx)
        pres.Slides.AddSlide(pres.Slides.Count + 1, layout)
        log_entry(log, op_id, "add_slide", pres.Slides.Count, True, f"Added slide with layout '{layout_name}'", file_name, iteration)
        time.sleep(0.3)
    except Exception as e:
        log_entry(log, op_id, "add_slide", None, False, f"Error: {e}", file_name, iteration)


def op_duplicate_slide(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    slide_idx = op.get("slide_index", pres.Slides.Count)
    try:
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            log_entry(log, op_id, "duplicate_slide", slide_idx, False, "Invalid slide index", file_name, iteration)
            return
        slide = pres.Slides(slide_idx)
        dup = slide.Duplicate()
        dup.Item(1).MoveTo(pres.Slides.Count)
        log_entry(log, op_id, "duplicate_slide", pres.Slides.Count, True, f"Duplicated slide {slide_idx}", file_name, iteration)
        time.sleep(0.3)
    except Exception as e:
        log_entry(log, op_id, "duplicate_slide", slide_idx, False, f"Error: {e}", file_name, iteration)


def op_delete_slide(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    slide_idx = op.get("slide_index")
    try:
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            log_entry(log, op_id, "delete_slide", slide_idx, False, "Invalid slide index", file_name, iteration)
            return
        pres.Slides(slide_idx).Delete()
        log_entry(log, op_id, "delete_slide", slide_idx, True, f"Deleted slide {slide_idx}", file_name, iteration)
        time.sleep(0.3)
    except Exception as e:
        log_entry(log, op_id, "delete_slide", slide_idx, False, f"Error: {e}", file_name, iteration)


def op_move_slide(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    from_idx = op.get("from_index")
    to_idx = op.get("to_index")
    try:
        if from_idx < 1 or from_idx > pres.Slides.Count or to_idx < 1 or to_idx > pres.Slides.Count:
            log_entry(log, op_id, "move_slide", from_idx, False, "Invalid slide indices", file_name, iteration)
            return
        pres.Slides(from_idx).MoveTo(to_idx)
        log_entry(log, op_id, "move_slide", to_idx, True, f"Moved slide {from_idx} to {to_idx}", file_name, iteration)
        time.sleep(0.3)
    except Exception as e:
        log_entry(log, op_id, "move_slide", from_idx, False, f"Error: {e}", file_name, iteration)


def op_change_slide_layout(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    slide_idx = op.get("slide_index", pres.Slides.Count)
    layout_name = op.get("layout", "Title and Content")
    try:
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            log_entry(log, op_id, "change_slide_layout", slide_idx, False, "Invalid slide index", file_name, iteration)
            return
        layout_idx = get_layout_index(pres, layout_name)
        layout = pres.SlideMaster.CustomLayouts(layout_idx)
        pres.Slides(slide_idx).Layout = layout
        log_entry(log, op_id, "change_slide_layout", slide_idx, True, f"Changed layout to '{layout_name}'", file_name, iteration)
        time.sleep(0.3)
    except Exception as e:
        log_entry(log, op_id, "change_slide_layout", slide_idx, False, f"Error: {e}", file_name, iteration)


def op_set_section_name(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    start_idx = op.get("start_index", 1)
    end_idx = op.get("end_index", pres.Slides.Count)
    name = op.get("name", "Section")
    try:
        pres.SectionProperties.AddSection(name, start_idx)
        log_entry(log, op_id, "set_section_name", start_idx, True, f"Added section '{name}'", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "set_section_name", start_idx, False, f"Error: {e}", file_name, iteration)


def op_rename_slide_titles(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    mode = op.get("mode", "title_case")
    changed = 0
    try:
        for s_idx in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(s_idx)
            for shape in slide.Shapes:
                if shape.Type == constants.msoPlaceholder:
                    if shape.PlaceholderFormat.Type in (constants.ppPlaceholderTitle, constants.ppPlaceholderCenterTitle):
                        if shape.HasTextFrame and shape.TextFrame.HasText:
                            tr = shape.TextFrame.TextRange
                            if mode == "title_case":
                                tr.Text = tr.Text.title()
                            changed += 1
                            break
        log_entry(log, op_id, "rename_slide_titles", None, True, f"Renamed {changed} titles", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "rename_slide_titles", None, False, f"Error: {e}", file_name, iteration)


def op_insert_slide_at_index(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    index = op.get("index", pres.Slides.Count + 1)
    layout_name = op.get("layout", "Title and Content")
    try:
        layout_idx = get_layout_index(pres, layout_name)
        layout = pres.SlideMaster.CustomLayouts(layout_idx)
        pres.Slides.AddSlide(index, layout)
        log_entry(log, op_id, "insert_slide_at_index", index, True, f"Inserted slide at index {index}", file_name, iteration)
        time.sleep(0.3)
    except Exception as e:
        log_entry(log, op_id, "insert_slide_at_index", index, False, f"Error: {e}", file_name, iteration)


def op_set_slide_hidden(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    slide_idx = op.get("slide_index", pres.Slides.Count)
    hidden = op.get("hidden", True)
    try:
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            log_entry(log, op_id, "set_slide_hidden", slide_idx, False, "Invalid slide index", file_name, iteration)
            return
        pres.Slides(slide_idx).SlideShowTransition.Hidden = hidden
        log_entry(log, op_id, "set_slide_hidden", slide_idx, True, f"Set slide {slide_idx} hidden={hidden}", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "set_slide_hidden", slide_idx, False, f"Error: {e}", file_name, iteration)


def op_set_slide_background_color(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    slide_idx = op.get("slide_index", pres.Slides.Count)
    rgb_hex = op.get("rgb", "#F2F2F2")
    try:
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            log_entry(log, op_id, "set_slide_background_color", slide_idx, False, "Invalid slide index", file_name, iteration)
            return
        # Convert hex to RGB
        rgb = int(rgb_hex.lstrip("#"), 16)
        slide = pres.Slides(slide_idx)
        slide.Background.Fill.ForeColor.RGB = rgb
        log_entry(log, op_id, "set_slide_background_color", slide_idx, True, f"Set background color {rgb_hex}", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "set_slide_background_color", slide_idx, False, f"Error: {e}", file_name, iteration)


def op_set_slide_title(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    idx = op.get("slide_index", pres.Slides.Count)
    text = op.get("text", "")
    try:
        if idx < 1 or idx > pres.Slides.Count:
            log_entry(log, op_id, "set_slide_title", idx, False, "Invalid slide index", file_name, iteration)
            return
        slide = pres.Slides(idx)
        for shape in slide.Shapes:
            if shape.Type == constants.msoPlaceholder:
                if shape.PlaceholderFormat.Type in (constants.ppPlaceholderTitle, constants.ppPlaceholderCenterTitle):
                    if shape.HasTextFrame:
                        shape.TextFrame.TextRange.Text = text
                        log_entry(log, op_id, "set_slide_title", idx, True, "Title set", file_name, iteration)
                        time.sleep(0.2)
                        return
        log_entry(log, op_id, "set_slide_title", idx, False, "No title placeholder found", file_name, iteration)
    except Exception as e:
        log_entry(log, op_id, "set_slide_title", idx, False, f"Error: {e}", file_name, iteration)


def op_replace_text(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    find = op.get("find", "")
    repl = op.get("replace", "")
    scope = op.get("scope", "all_slides")
    count = 0
    try:
        for s_idx in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(s_idx)
            shapes_to_check = []
            if scope == "titles_only":
                for shape in slide.Shapes:
                    if shape.Type == constants.msoPlaceholder:
                        if shape.PlaceholderFormat.Type in (constants.ppPlaceholderTitle, constants.ppPlaceholderCenterTitle):
                            if shape.HasTextFrame:
                                shapes_to_check.append(shape)
                                break
            else:
                shapes_to_check = [s for s in slide.Shapes if s.HasTextFrame]
            
            for shape in shapes_to_check:
                tf = shape.TextFrame
                if tf.HasText:
                    tr = tf.TextRange
                    if find in tr.Text:
                        tr.Text = tr.Text.replace(find, repl)
                        count += 1
        log_entry(log, op_id, "replace_text", None, True, f"Replaced {count} occurrences", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "replace_text", None, False, f"Error: {e}", file_name, iteration)


def op_change_font_family_slide(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    slide_idx = op.get("slide_index", pres.Slides.Count)
    font_name = op.get("font_name", "Calibri")
    try:
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            log_entry(log, op_id, "change_font_family_slide", slide_idx, False, "Invalid slide index", file_name, iteration)
            return
        slide = pres.Slides(slide_idx)
        count = 0
        for shape in slide.Shapes:
            if shape.HasTextFrame:
                tr = shape.TextFrame.TextRange
                for p in range(1, tr.Paragraphs().Count + 1):
                    para = tr.Paragraphs(p)
                    for r in range(1, para.Runs().Count + 1):
                        run = para.Runs(r)
                        run.Font.Name = font_name
                        count += 1
        log_entry(log, op_id, "change_font_family_slide", slide_idx, True, f"Changed font on {count} runs", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "change_font_family_slide", slide_idx, False, f"Error: {e}", file_name, iteration)


def op_change_title_font_size(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    font_size = op.get("font_size", 32)
    changed = 0
    try:
        for s_idx in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(s_idx)
            for shape in slide.Shapes:
                if shape.Type == constants.msoPlaceholder:
                    if shape.PlaceholderFormat.Type in (constants.ppPlaceholderTitle, constants.ppPlaceholderCenterTitle):
                        if shape.HasTextFrame:
                            tr = shape.TextFrame.TextRange
                            tr.Font.Size = font_size
                            changed += 1
                            break
        log_entry(log, op_id, "change_title_font_size", None, True, f"Updated {changed} titles", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "change_title_font_size", None, False, f"Error: {e}", file_name, iteration)


def op_change_title_font(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    font_name = op.get("font_name", "Calibri")
    font_size = op.get("font_size", 32)
    changed = 0
    try:
        for s_idx in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(s_idx)
            for shape in slide.Shapes:
                if shape.Type == constants.msoPlaceholder:
                    if shape.PlaceholderFormat.Type in (constants.ppPlaceholderTitle, constants.ppPlaceholderCenterTitle):
                        if shape.HasTextFrame:
                            tr = shape.TextFrame.TextRange
                            tr.Font.Name = font_name
                            tr.Font.Size = font_size
                            changed += 1
                            break
        log_entry(log, op_id, "change_title_font", None, True, f"Updated {changed} titles", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "change_title_font", None, False, f"Error: {e}", file_name, iteration)


def op_bullets_to_numbered(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    slide_idx = op.get("slide_index", pres.Slides.Count)
    try:
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            log_entry(log, op_id, "bullets_to_numbered", slide_idx, False, "Invalid slide index", file_name, iteration)
            return
        slide = pres.Slides(slide_idx)
        count = 0
        for shape in slide.Shapes:
            if shape.HasTextFrame:
                tf = shape.TextFrame
                for p in range(1, tf.TextRange.Paragraphs().Count + 1):
                    para = tf.TextRange.Paragraphs(p)
                    if para.ParagraphFormat.Bullet.Type != constants.ppBulletNone:
                        para.ParagraphFormat.Bullet.Type = constants.ppBulletNumbered
                        count += 1
        log_entry(log, op_id, "bullets_to_numbered", slide_idx, True, f"Converted {count} bullets", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "bullets_to_numbered", slide_idx, False, f"Error: {e}", file_name, iteration)


def op_align_text(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    slide_idx = op.get("slide_index", pres.Slides.Count)
    scope = op.get("scope", "all")
    alignment = op.get("alignment", "left")
    align_map = {"left": constants.ppAlignLeft, "center": constants.ppAlignCenter, "right": constants.ppAlignRight, "justify": constants.ppAlignJustify}
    align_const = align_map.get(alignment, constants.ppAlignLeft)
    try:
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            log_entry(log, op_id, "align_text", slide_idx, False, "Invalid slide index", file_name, iteration)
            return
        slide = pres.Slides(slide_idx)
        count = 0
        for shape in slide.Shapes:
            if shape.HasTextFrame:
                is_title = False
                if shape.Type == constants.msoPlaceholder:
                    if shape.PlaceholderFormat.Type in (constants.ppPlaceholderTitle, constants.ppPlaceholderCenterTitle):
                        is_title = True
                
                if (scope == "all") or (scope == "title_only" and is_title) or (scope == "body_only" and not is_title):
                    tf = shape.TextFrame
                    for p in range(1, tf.TextRange.Paragraphs().Count + 1):
                        para = tf.TextRange.Paragraphs(p)
                        para.Alignment = align_const
                        count += 1
        log_entry(log, op_id, "align_text", slide_idx, True, f"Aligned {count} paragraphs", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "align_text", slide_idx, False, f"Error: {e}", file_name, iteration)


def op_title_capitalize(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    mode = op.get("mode", "title_case")
    changed = 0
    try:
        for s_idx in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(s_idx)
            for shape in slide.Shapes:
                if shape.Type == constants.msoPlaceholder:
                    if shape.PlaceholderFormat.Type in (constants.ppPlaceholderTitle, constants.ppPlaceholderCenterTitle):
                        if shape.HasTextFrame and shape.TextFrame.HasText:
                            tr = shape.TextFrame.TextRange
                            if mode == "title_case":
                                tr.Text = tr.Text.title()
                            changed += 1
                            break
        log_entry(log, op_id, "title_capitalize", None, True, f"Capitalized {changed} titles", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "title_capitalize", None, False, f"Error: {e}", file_name, iteration)


def op_set_footer_text(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    text = op.get("text", "")
    try:
        pres.SlideMaster.HeadersFooters.Footer.Text = text
        pres.SlideMaster.HeadersFooters.Footer.Visible = True
        log_entry(log, op_id, "set_footer_text", None, True, f"Set footer text", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "set_footer_text", None, False, f"Error: {e}", file_name, iteration)


def op_insert_slide_number(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    try:
        pres.SlideMaster.HeadersFooters.SlideNumber.Visible = True
        log_entry(log, op_id, "insert_slide_number", None, True, "Enabled slide numbers", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "insert_slide_number", None, False, f"Error: {e}", file_name, iteration)


def op_format_keywords(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    keyword = op.get("keyword", "")
    bold = op.get("bold", False)
    italic = op.get("italic", False)
    count = 0
    try:
        for s_idx in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(s_idx)
            for shape in slide.Shapes:
                if shape.HasTextFrame and shape.TextFrame.HasText:
                    tr = shape.TextFrame.TextRange
                    start_pos = tr.Text.find(keyword)
                    if start_pos >= 0:
                        found_range = tr.Characters(start_pos + 1, len(keyword))
                        found_range.Font.Bold = bold
                        found_range.Font.Italic = italic
                        count += 1
        log_entry(log, op_id, "format_keywords", None, True, f"Formatted {count} occurrences", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "format_keywords", None, False, f"Error: {e}", file_name, iteration)


def op_add_speaker_notes(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    slide_idx = op.get("slide_index", pres.Slides.Count)
    text = op.get("text", "")
    try:
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            log_entry(log, op_id, "add_speaker_notes", slide_idx, False, "Invalid slide index", file_name, iteration)
            return
        slide = pres.Slides(slide_idx)
        notes_page = slide.NotesPage
        # Find notes placeholder (usually index 2)
        if notes_page.Shapes.Placeholders.Count >= 2:
            notes_shape = notes_page.Shapes.Placeholders(2)
            if notes_shape.HasTextFrame:
                notes_shape.TextFrame.TextRange.Text = text
        log_entry(log, op_id, "add_speaker_notes", slide_idx, True, "Added speaker notes", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "add_speaker_notes", slide_idx, False, f"Error: {e}", file_name, iteration)


def op_insert_shape(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    slide_idx = op.get("slide_index", pres.Slides.Count)
    shape_type = op.get("shape_type", "rectangle")
    left = op.get("left", 1.0)
    top = op.get("top", 1.0)
    width = op.get("width", 4.0)
    height = op.get("height", 1.0)
    units = op.get("units", "inches")
    
    shape_map = {
        "rectangle": constants.msoShapeRectangle,
        "circle": constants.msoShapeOval,
        "arrow": constants.msoShapeRightArrow,
    }
    shape_const = shape_map.get(shape_type, constants.msoShapeRectangle)
    
    try:
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            log_entry(log, op_id, "insert_shape", slide_idx, False, "Invalid slide index", file_name, iteration)
            return
        
        if units == "cm":
            left_pt = left * CM_TO_POINTS
            top_pt = top * CM_TO_POINTS
            width_pt = width * CM_TO_POINTS
            height_pt = height * CM_TO_POINTS
        else:  # inches
            left_pt = left * 72
            top_pt = top * 72
            width_pt = width * 72
            height_pt = height * 72
        
        slide = pres.Slides(slide_idx)
        shape = slide.Shapes.AddShape(shape_const, left_pt, top_pt, width_pt, height_pt)
        log_entry(log, op_id, "insert_shape", slide_idx, True, f"Inserted {shape_type}", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "insert_shape", slide_idx, False, f"Error: {e}", file_name, iteration)


def op_resize_shape(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    slide_idx = op.get("slide_index", pres.Slides.Count)
    shape_index = op.get("shape_index", 1)
    width = op.get("width", 4.0)
    height = op.get("height", 1.0)
    units = op.get("units", "inches")
    try:
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            log_entry(log, op_id, "resize_shape", slide_idx, False, "Invalid slide index", file_name, iteration)
            return
        slide = pres.Slides(slide_idx)
        if shape_index < 1 or shape_index > slide.Shapes.Count:
            log_entry(log, op_id, "resize_shape", slide_idx, False, "Invalid shape index", file_name, iteration)
            return
        
        if units == "cm":
            width_pt = width * CM_TO_POINTS
            height_pt = height * CM_TO_POINTS
        else:
            width_pt = width * 72
            height_pt = height * 72
        
        shape = slide.Shapes(shape_index)
        shape.Width = width_pt
        shape.Height = height_pt
        log_entry(log, op_id, "resize_shape", slide_idx, True, f"Resized shape {shape_index}", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "resize_shape", slide_idx, False, f"Error: {e}", file_name, iteration)


def op_align_shapes_evenly(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    slide_idx = op.get("slide_index", pres.Slides.Count)
    shape_indices = op.get("shape_indices", [])
    mode = op.get("mode", "horizontal")
    try:
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            log_entry(log, op_id, "align_shapes_evenly", slide_idx, False, "Invalid slide index", file_name, iteration)
            return
        slide = pres.Slides(slide_idx)
        if not shape_indices or len(shape_indices) < 2:
            log_entry(log, op_id, "align_shapes_evenly", slide_idx, False, "Need at least 2 shapes", file_name, iteration)
            return
        
        shapes = slide.Shapes.Range([slide.Shapes(i).Name for i in shape_indices if 1 <= i <= slide.Shapes.Count])
        if mode == "horizontal":
            shapes.Distribute(constants.msoDistributeHorizontally, False)
        else:
            shapes.Distribute(constants.msoDistributeVertically, False)
        log_entry(log, op_id, "align_shapes_evenly", slide_idx, True, f"Aligned {len(shape_indices)} shapes", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "align_shapes_evenly", slide_idx, False, f"Error: {e}", file_name, iteration)


def op_group_shapes(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    slide_idx = op.get("slide_index", pres.Slides.Count)
    shape_indices = op.get("shape_indices", [])
    try:
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            log_entry(log, op_id, "group_shapes", slide_idx, False, "Invalid slide index", file_name, iteration)
            return
        slide = pres.Slides(slide_idx)
        if len(shape_indices) < 2:
            log_entry(log, op_id, "group_shapes", slide_idx, False, "Need at least 2 shapes", file_name, iteration)
            return
        
        shapes = slide.Shapes.Range([slide.Shapes(i).Name for i in shape_indices if 1 <= i <= slide.Shapes.Count])
        shapes.Group()
        log_entry(log, op_id, "group_shapes", slide_idx, True, f"Grouped {len(shape_indices)} shapes", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "group_shapes", slide_idx, False, f"Error: {e}", file_name, iteration)


def op_ungroup_shapes(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    slide_idx = op.get("slide_index", pres.Slides.Count)
    shape_index = op.get("shape_index", 1)
    try:
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            log_entry(log, op_id, "ungroup_shapes", slide_idx, False, "Invalid slide index", file_name, iteration)
            return
        slide = pres.Slides(slide_idx)
        if shape_index < 1 or shape_index > slide.Shapes.Count:
            log_entry(log, op_id, "ungroup_shapes", slide_idx, False, "Invalid shape index", file_name, iteration)
            return
        shape = slide.Shapes(shape_index)
        if shape.Type == constants.msoGroup:
            shape.Ungroup()
            log_entry(log, op_id, "ungroup_shapes", slide_idx, True, "Ungrouped shape", file_name, iteration)
        else:
            log_entry(log, op_id, "ungroup_shapes", slide_idx, False, "Shape is not a group", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "ungroup_shapes", slide_idx, False, f"Error: {e}", file_name, iteration)


def op_set_shape_fill_color(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    slide_idx = op.get("slide_index", pres.Slides.Count)
    shape_index = op.get("shape_index", 1)
    rgb_hex = op.get("rgb", "#0070C0")
    try:
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            log_entry(log, op_id, "set_shape_fill_color", slide_idx, False, "Invalid slide index", file_name, iteration)
            return
        slide = pres.Slides(slide_idx)
        if shape_index < 1 or shape_index > slide.Shapes.Count:
            log_entry(log, op_id, "set_shape_fill_color", slide_idx, False, "Invalid shape index", file_name, iteration)
            return
        rgb = int(rgb_hex.lstrip("#"), 16)
        shape = slide.Shapes(shape_index)
        shape.Fill.ForeColor.RGB = rgb
        log_entry(log, op_id, "set_shape_fill_color", slide_idx, True, f"Set fill color {rgb_hex}", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "set_shape_fill_color", slide_idx, False, f"Error: {e}", file_name, iteration)


def op_set_shape_border(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    slide_idx = op.get("slide_index", pres.Slides.Count)
    shape_index = op.get("shape_index", 1)
    rgb_hex = op.get("rgb", "#000000")
    width_pt = op.get("width_pt", 1.0)
    try:
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            log_entry(log, op_id, "set_shape_border", slide_idx, False, "Invalid slide index", file_name, iteration)
            return
        slide = pres.Slides(slide_idx)
        if shape_index < 1 or shape_index > slide.Shapes.Count:
            log_entry(log, op_id, "set_shape_border", slide_idx, False, "Invalid shape index", file_name, iteration)
            return
        rgb = int(rgb_hex.lstrip("#"), 16)
        shape = slide.Shapes(shape_index)
        shape.Line.ForeColor.RGB = rgb
        shape.Line.Weight = width_pt
        shape.Line.Visible = True
        log_entry(log, op_id, "set_shape_border", slide_idx, True, f"Set border color {rgb_hex}", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "set_shape_border", slide_idx, False, f"Error: {e}", file_name, iteration)


def op_set_shape_text(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    slide_idx = op.get("slide_index", pres.Slides.Count)
    shape_index = op.get("shape_index", 1)
    text = op.get("text", "")
    try:
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            log_entry(log, op_id, "set_shape_text", slide_idx, False, "Invalid slide index", file_name, iteration)
            return
        slide = pres.Slides(slide_idx)
        if shape_index < 1 or shape_index > slide.Shapes.Count:
            log_entry(log, op_id, "set_shape_text", slide_idx, False, "Invalid shape index", file_name, iteration)
            return
        shape = slide.Shapes(shape_index)
        if shape.HasTextFrame:
            shape.TextFrame.TextRange.Text = text
            log_entry(log, op_id, "set_shape_text", slide_idx, True, "Set shape text", file_name, iteration)
        else:
            log_entry(log, op_id, "set_shape_text", slide_idx, False, "Shape has no text frame", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "set_shape_text", slide_idx, False, f"Error: {e}", file_name, iteration)


def op_copy_shape_between_slides(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    from_slide = op.get("from_slide")
    shape_index = op.get("shape_index", 1)
    to_slide = op.get("to_slide")
    try:
        if from_slide < 1 or from_slide > pres.Slides.Count or to_slide < 1 or to_slide > pres.Slides.Count:
            log_entry(log, op_id, "copy_shape_between_slides", from_slide, False, "Invalid slide indices", file_name, iteration)
            return
        from_slide_obj = pres.Slides(from_slide)
        to_slide_obj = pres.Slides(to_slide)
        if shape_index < 1 or shape_index > from_slide_obj.Shapes.Count:
            log_entry(log, op_id, "copy_shape_between_slides", from_slide, False, "Invalid shape index", file_name, iteration)
            return
        shape = from_slide_obj.Shapes(shape_index)
        shape.Copy()
        to_slide_obj.Shapes.Paste()
        log_entry(log, op_id, "copy_shape_between_slides", to_slide, True, f"Copied shape from slide {from_slide}", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "copy_shape_between_slides", from_slide, False, f"Error: {e}", file_name, iteration)


def op_lock_object_position(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    slide_idx = op.get("slide_index", pres.Slides.Count)
    shape_index = op.get("shape_index", 1)
    try:
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            log_entry(log, op_id, "lock_object_position", slide_idx, False, "Invalid slide index", file_name, iteration)
            return
        slide = pres.Slides(slide_idx)
        if shape_index < 1 or shape_index > slide.Shapes.Count:
            log_entry(log, op_id, "lock_object_position", slide_idx, False, "Invalid shape index", file_name, iteration)
            return
        # Note: PowerPoint doesn't have direct lock, but we can send to back as workaround
        shape = slide.Shapes(shape_index)
        shape.ZOrder(constants.msoSendToBack)
        log_entry(log, op_id, "lock_object_position", slide_idx, True, "Sent shape to back (workaround)", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "lock_object_position", slide_idx, False, f"Error: {e}", file_name, iteration)


def op_insert_image(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    slide_idx = op.get("slide_index", pres.Slides.Count)
    image_path = op.get("image_path", "")
    left = op.get("left", 1.0)
    top = op.get("top", 1.0)
    width = op.get("width", 4.0)
    height = op.get("height", 3.0)
    units = op.get("units", "inches")
    try:
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            log_entry(log, op_id, "insert_image", slide_idx, False, "Invalid slide index", file_name, iteration)
            return
        img_path = WORK_DIR / image_path if image_path else None
        if not img_path or not img_path.exists():
            log_entry(log, op_id, "insert_image", slide_idx, False, f"Image not found: {image_path}", file_name, iteration)
            return
        
        if units == "cm":
            left_pt = left * CM_TO_POINTS
            top_pt = top * CM_TO_POINTS
            width_pt = width * CM_TO_POINTS
            height_pt = height * CM_TO_POINTS
        else:
            left_pt = left * 72
            top_pt = top * 72
            width_pt = width * 72
            height_pt = height * 72
        
        slide = pres.Slides(slide_idx)
        slide.Shapes.AddPicture(str(img_path), False, True, left_pt, top_pt, width_pt, height_pt)
        log_entry(log, op_id, "insert_image", slide_idx, True, f"Inserted image {image_path}", file_name, iteration)
        time.sleep(0.3)
    except Exception as e:
        log_entry(log, op_id, "insert_image", slide_idx, False, f"Error: {e}", file_name, iteration)


def op_resize_image(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    slide_idx = op.get("slide_index", pres.Slides.Count)
    shape_index = op.get("shape_index", 1)
    width = op.get("width", 4.0)
    height = op.get("height", 3.0)
    units = op.get("units", "inches")
    try:
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            log_entry(log, op_id, "resize_image", slide_idx, False, "Invalid slide index", file_name, iteration)
            return
        slide = pres.Slides(slide_idx)
        if shape_index < 1 or shape_index > slide.Shapes.Count:
            log_entry(log, op_id, "resize_image", slide_idx, False, "Invalid shape index", file_name, iteration)
            return
        
        if units == "cm":
            width_pt = width * CM_TO_POINTS
            height_pt = height * CM_TO_POINTS
        else:
            width_pt = width * 72
            height_pt = height * 72
        
        shape = slide.Shapes(shape_index)
        if shape.Type == constants.msoLinkedPicture or shape.Type == constants.msoPicture:
            shape.Width = width_pt
            shape.Height = height_pt
            log_entry(log, op_id, "resize_image", slide_idx, True, f"Resized image", file_name, iteration)
        else:
            log_entry(log, op_id, "resize_image", slide_idx, False, "Shape is not an image", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "resize_image", slide_idx, False, f"Error: {e}", file_name, iteration)


def op_crop_image(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    slide_idx = op.get("slide_index", pres.Slides.Count)
    shape_index = op.get("shape_index", 1)
    crop_left = op.get("crop_left", 0.0)
    crop_right = op.get("crop_right", 0.0)
    crop_top = op.get("crop_top", 0.0)
    crop_bottom = op.get("crop_bottom", 0.0)
    try:
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            log_entry(log, op_id, "crop_image", slide_idx, False, "Invalid slide index", file_name, iteration)
            return
        slide = pres.Slides(slide_idx)
        if shape_index < 1 or shape_index > slide.Shapes.Count:
            log_entry(log, op_id, "crop_image", slide_idx, False, "Invalid shape index", file_name, iteration)
            return
        shape = slide.Shapes(shape_index)
        if shape.Type == constants.msoLinkedPicture or shape.Type == constants.msoPicture:
            shape.PictureFormat.CropLeft = crop_left * 72
            shape.PictureFormat.CropRight = crop_right * 72
            shape.PictureFormat.CropTop = crop_top * 72
            shape.PictureFormat.CropBottom = crop_bottom * 72
            log_entry(log, op_id, "crop_image", slide_idx, True, "Cropped image", file_name, iteration)
        else:
            log_entry(log, op_id, "crop_image", slide_idx, False, "Shape is not an image", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "crop_image", slide_idx, False, f"Error: {e}", file_name, iteration)


def op_align_image_with_title(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    slide_idx = op.get("slide_index", pres.Slides.Count)
    mode = op.get("mode", "center")
    try:
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            log_entry(log, op_id, "align_image_with_title", slide_idx, False, "Invalid slide index", file_name, iteration)
            return
        slide = pres.Slides(slide_idx)
        # Find title shape
        title_left = None
        title_width = None
        for shape in slide.Shapes:
            if shape.Type == constants.msoPlaceholder:
                if shape.PlaceholderFormat.Type in (constants.ppPlaceholderTitle, constants.ppPlaceholderCenterTitle):
                    title_left = shape.Left
                    title_width = shape.Width
                    break
        
        if title_left is None:
            log_entry(log, op_id, "align_image_with_title", slide_idx, False, "No title found", file_name, iteration)
            return
        
        # Find first image
        for shape in slide.Shapes:
            if shape.Type == constants.msoLinkedPicture or shape.Type == constants.msoPicture:
                if mode == "center":
                    shape.Left = title_left + (title_width - shape.Width) / 2
                else:  # left
                    shape.Left = title_left
                log_entry(log, op_id, "align_image_with_title", slide_idx, True, f"Aligned image {mode}", file_name, iteration)
                break
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "align_image_with_title", slide_idx, False, f"Error: {e}", file_name, iteration)


def op_set_image_alt_text(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int):
    op_id = op.get("op_id")
    slide_idx = op.get("slide_index", pres.Slides.Count)
    shape_index = op.get("shape_index", 1)
    alt_text = op.get("alt_text", "")
    try:
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            log_entry(log, op_id, "set_image_alt_text", slide_idx, False, "Invalid slide index", file_name, iteration)
            return
        slide = pres.Slides(slide_idx)
        if shape_index < 1 or shape_index > slide.Shapes.Count:
            log_entry(log, op_id, "set_image_alt_text", slide_idx, False, "Invalid shape index", file_name, iteration)
            return
        shape = slide.Shapes(shape_index)
        shape.AlternativeText = alt_text
        log_entry(log, op_id, "set_image_alt_text", slide_idx, True, "Set alt text", file_name, iteration)
        time.sleep(0.2)
    except Exception as e:
        log_entry(log, op_id, "set_image_alt_text", slide_idx, False, f"Error: {e}", file_name, iteration)


# Placeholder handlers for remaining operations
def op_placeholder(pres, op: Dict[str, Any], log: List[Dict[str, Any]], file_name: str, iteration: int, action_name: str):
    op_id = op.get("op_id")
    log_entry(log, op_id, action_name, None, False, f"Operation '{action_name}' not yet implemented", file_name, iteration)


ACTION_HANDLERS = {
    "add_slide": op_add_slide,
    "duplicate_slide": op_duplicate_slide,
    "delete_slide": op_delete_slide,
    "move_slide": op_move_slide,
    "change_slide_layout": op_change_slide_layout,
    "set_section_name": op_set_section_name,
    "rename_slide_titles": op_rename_slide_titles,
    "insert_slide_at_index": op_insert_slide_at_index,
    "set_slide_hidden": op_set_slide_hidden,
    "set_slide_background_color": op_set_slide_background_color,
    "replace_text": op_replace_text,
    "change_font_family_slide": op_change_font_family_slide,
    "change_title_font_size": op_change_title_font_size,
    "bullets_to_numbered": op_bullets_to_numbered,
    "align_text": op_align_text,
    "title_capitalize": op_title_capitalize,
    "set_footer_text": op_set_footer_text,
    "insert_slide_number": op_insert_slide_number,
    "format_keywords": op_format_keywords,
    "add_speaker_notes": op_add_speaker_notes,
    "insert_shape": op_insert_shape,
    "resize_shape": op_resize_shape,
    "align_shapes_evenly": op_align_shapes_evenly,
    "group_shapes": op_group_shapes,
    "ungroup_shapes": op_ungroup_shapes,
    "set_shape_fill_color": op_set_shape_fill_color,
    "set_shape_border": op_set_shape_border,
    "set_shape_text": op_set_shape_text,
    "copy_shape_between_slides": op_copy_shape_between_slides,
    "lock_object_position": op_lock_object_position,
    "insert_image": op_insert_image,
    "resize_image": op_resize_image,
    "crop_image": op_crop_image,
    "align_image_with_title": op_align_image_with_title,
    "set_image_alt_text": op_set_image_alt_text,
    # Placeholders for remaining
    "apply_picture_style": lambda p, o, l, f, i: op_placeholder(p, o, l, f, i, "apply_picture_style"),
    "replace_image_preserve": lambda p, o, l, f, i: op_placeholder(p, o, l, f, i, "replace_image_preserve"),
    "remove_image_background": lambda p, o, l, f, i: op_placeholder(p, o, l, f, i, "remove_image_background"),
    "set_image_transparency": lambda p, o, l, f, i: op_placeholder(p, o, l, f, i, "set_image_transparency"),
    "add_image_caption": lambda p, o, l, f, i: op_placeholder(p, o, l, f, i, "add_image_caption"),
    "apply_theme": lambda p, o, l, f, i: op_placeholder(p, o, l, f, i, "apply_theme"),
    "apply_master_layout": lambda p, o, l, f, i: op_placeholder(p, o, l, f, i, "apply_master_layout"),
    "standardize_margins": lambda p, o, l, f, i: op_placeholder(p, o, l, f, i, "standardize_margins"),
    "ensure_title_font_consistency": lambda p, o, l, f, i: op_placeholder(p, o, l, f, i, "ensure_title_font_consistency"),
    "apply_color_palette": lambda p, o, l, f, i: op_placeholder(p, o, l, f, i, "apply_color_palette"),
    "export_pdf": lambda p, o, l, f, i: op_placeholder(p, o, l, f, i, "export_pdf"),
    "export_specific_slides": lambda p, o, l, f, i: op_placeholder(p, o, l, f, i, "export_specific_slides"),
    "validate_missing_titles": lambda p, o, l, f, i: op_placeholder(p, o, l, f, i, "validate_missing_titles"),
    "validate_text_overflow": lambda p, o, l, f, i: op_placeholder(p, o, l, f, i, "validate_text_overflow"),
    "generate_slide_summary": lambda p, o, l, f, i: op_placeholder(p, o, l, f, i, "generate_slide_summary"),
    "set_slide_title": op_set_slide_title,
    "insert_bullet_points": lambda p, o, l, f, i: op_placeholder(p, o, l, f, i, "insert_bullet_points"),
    "change_title_font": op_change_title_font,
}


def process_instructions(instructions: Dict[str, Any]):
    file_name = instructions.get("file")
    files_list = instructions.get("files")
    operations = instructions.get("operations", [])
    iterations = int(instructions.get("iterations", 1))
    log: List[Dict[str, Any]] = []

    if files_list and isinstance(files_list, list):
        target_files = files_list
    elif file_name:
        target_files = [file_name]
    else:
        print("No 'file' or 'files' specified in instructions.")
        return

    ppt_app = ensure_powerpoint()
    print("PowerPoint application opened and visible.")

    for fname in target_files:
        file_path = WORK_DIR / fname
        if not file_path.exists():
            print(f"File does not exist: {file_path}")
            continue

        backup_path = backup_pptx(file_path)
        if backup_path:
            print(f"Backup created: {backup_path}")
        else:
            print(f"Continuing without backup (file may be locked)")

        try:
            pres = open_presentation(ppt_app, file_path)
            print(f"Opened presentation: {fname} ({pres.Slides.Count} slides)")

            # Run iterations: each iteration duplicates the previous slide, then performs operations
            for iteration in range(1, iterations + 1):
                print(f"\n--- Iteration {iteration}/{iterations} for {fname} ---")
                
                # STEP 1: Duplicate the last slide (or first slide if only one exists)
                if pres.Slides.Count > 0:
                    last_slide_idx = pres.Slides.Count
                    try:
                        last_slide = pres.Slides(last_slide_idx)
                        dup = last_slide.Duplicate()
                        dup.Item(1).MoveTo(pres.Slides.Count)
                        new_slide_idx = pres.Slides.Count
                        print(f"  Duplicated slide {last_slide_idx} -> new slide {new_slide_idx}")
                        time.sleep(0.5)
                    except Exception as e:
                        print(f"  Warning: Could not duplicate slide: {e}")
                else:
                    # No slides exist, create a blank one
                    try:
                        layout = pres.SlideMaster.CustomLayouts(7)  # Blank layout
                        pres.Slides.AddSlide(1, layout)
                        print(f"  Created initial blank slide")
                        time.sleep(0.5)
                    except Exception as e:
                        print(f"  Error creating initial slide: {e}")
                        continue

                # STEP 2: Perform all operations on the newly duplicated slide (or last slide)
                current_slide_idx = pres.Slides.Count
                
                for op in operations:
                    action = op.get("action")
                    handler = ACTION_HANDLERS.get(action)
                    
                    if not handler:
                        log_entry(
                            log,
                            op.get("op_id"),
                            action or "",
                            op.get("slide_index"),
                            False,
                            "Unknown/unsupported action",
                            fname,
                            iteration,
                        )
                        continue

                    try:
                        # For operations that need a slide_index, default to the current slide
                        if "slide_index" in op and op["slide_index"] is None:
                            op["slide_index"] = current_slide_idx
                        elif "slide_index" not in op:
                            # Only set default for operations that typically need a slide
                            if action in ["set_slide_title", "duplicate_slide", "delete_slide", "change_slide_layout", 
                                         "set_slide_hidden", "set_slide_background_color", "change_font_family_slide",
                                         "bullets_to_numbered", "align_text", "add_speaker_notes", "insert_shape",
                                         "resize_shape", "align_shapes_evenly", "group_shapes", "ungroup_shapes",
                                         "set_shape_fill_color", "set_shape_border", "set_shape_text", "lock_object_position",
                                         "insert_image", "resize_image", "crop_image", "align_image_with_title",
                                         "set_image_alt_text"]:
                                op["slide_index"] = current_slide_idx
                        
                        handler(pres, op, log, fname, iteration)
                        time.sleep(0.1)
                    except Exception as e:
                        log_entry(
                            log,
                            op.get("op_id"),
                            action,
                            op.get("slide_index"),
                            False,
                            f"Unhandled exception: {e}",
                            fname,
                            iteration,
                        )

                # Save after each iteration
                pres.Save()
                print(f"  Saved after iteration {iteration}")

            pres.Close()
            print(f"Closed presentation: {fname}")

        except Exception as e:
            print(f"Error processing {fname}: {e}")
            log_entry(log, None, "file_error", None, False, f"Error processing file: {e}", fname, None)

    print("\nAll files processed. PowerPoint remains open for inspection.")
    log_path = WORK_DIR / "log.json"
    with log_path.open("w", encoding="utf-8") as f:
        json.dump(log, f, indent=2, ensure_ascii=False)
    print(f"Operation log written to: {log_path}")


def main():
    if len(sys.argv) != 2:
        print("Usage: python modify_powerpoint.py instructions.json")
        sys.exit(1)

    json_path = Path(sys.argv[1])
    if not json_path.is_absolute():
        json_path = WORK_DIR / json_path

    if not json_path.exists():
        print(f"Instructions file not found: {json_path}")
        sys.exit(1)

    instructions = load_instructions(json_path)
    process_instructions(instructions)


if __name__ == "__main__":
    main()
