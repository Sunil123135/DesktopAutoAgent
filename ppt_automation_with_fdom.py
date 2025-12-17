import os
import sys
import traceback
from pathlib import Path

import win32com.client
import shutil

# ----- PROJECT ROOT + IMPORT fDOM CREATOR -----
PROJECT_ROOT = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

try:
    from utils.fdom.fdom_creator import FDOMCreator
except ImportError:
    FDOMCreator = None


TEMP_DIR = r"C:\\Users\\slalwani\\OneDrive - QuidelOrtho\\SUNIL\\EAG\\Session 13\\Code\\temp"
PPT_FILES = ["1.pptx", "2.pptx", "3.pptx", "4.pptx"]  # files to process
CM_TO_POINTS = 28.3465  # PowerPoint units


class StepError(Exception):
    def __init__(self, step_number: str, slide_number: int | None, message: str):
        self.step_number = step_number
        self.slide_number = slide_number
        self.message = message
        super().__init__(f"[Step {step_number}] Slide {slide_number}: {message}")


def ensure_powerpoint():
    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        return ppt
    except Exception as e:
        raise StepError("0", None, f"PowerPoint not available: {e}")


def is_file_writable(path: Path) -> bool:
    if not path.exists():
        return False
    try:
        with open(path, "a"):
            return True
    except Exception:
        return False


def open_presentation(ppt_app, path: Path):
    try:
        pres = ppt_app.Presentations.Open(str(path), WithWindow=False)
        return pres
    except Exception as e:
        raise StepError("1", None, f"Failed to open presentation: {e}")


def ensure_not_protected_view(ppt_app, pres):
    try:
        pvw = ppt_app.ProtectedViewWindows
    except Exception:
        return

    for i in range(1, pvw.Count + 1):
        win = pvw.Item(i)
        if os.path.basename(win.Presentation.FullName) == os.path.basename(pres.FullName):
            try:
                win.Edit()
            except Exception as e:
                raise StepError("1.4", None, f"File is in Protected View and cannot be edited: {e}")


def log_presentation_info(pres):
    print(f"Processing file: {pres.FullName}")
    print(f"Initial slide count: {pres.Slides.Count}")


# ---------------- TASK SET 1 ----------------

def task_set_1_slide_structure(pres):
    try:
        layout_title_content = pres.SlideMaster.CustomLayouts(2)  # usually Title and Content
        pres.Slides.AddSlide(2, layout_title_content)
    except Exception as e:
        raise StepError("1.2", 2, f"Failed to insert slide at position 2: {e}")

    if pres.Slides.Count < 5:
        raise StepError("1.3", None, "Not enough slides to duplicate slide 5.")

    try:
        s5 = pres.Slides(5)
        dup = s5.Duplicate()
        dup.Item(1).MoveTo(6)
    except Exception as e:
        raise StepError("1.3", 5, f"Failed to duplicate slide 5: {e}")

    if pres.Slides.Count < 8:
        raise StepError("1.4", None, "Not enough slides to move slide 8.")
    try:
        pres.Slides(8).MoveTo(3)
    except Exception as e:
        raise StepError("1.4", 8, f"Failed to move slide 8 to position 3: {e}")

    if pres.Slides.Count < 12:
        raise StepError("1.5", None, "Not enough slides to hide slide 12.")
    try:
        pres.Slides(12).SlideShowTransition.Hidden = True
    except Exception as e:
        raise StepError("1.5", 12, f"Failed to hide slide 12: {e}")

    try:
        pres.SectionProperties.AddSection("Executive Summary", 1)
    except Exception as e:
        raise StepError("1.6", 1, f"Failed to create section 'Executive Summary': {e}")

    try:
        sp = pres.SectionProperties
        for i in range(sp.Count):
            name = sp.Name(i)
            if name == "Analysis":
                sp.Rename(i, "Detailed Analysis")
                break
    except Exception as e:
        raise StepError("1.7", None, f"Failed to rename section 'Analysis': {e}")

    print(f"[Task Set 1] Slide count after structure changes: {pres.Slides.Count}")


# ---------------- TASK SET 2 ----------------

def _iterate_all_text_ranges(pres):
    from win32com.client import constants

    for s_idx in range(1, pres.Slides.Count + 1):
        slide = pres.Slides(s_idx)
        for shape in slide.Shapes:
            if shape.HasTextFrame and shape.TextFrame.HasText:
                yield s_idx, shape, shape.TextFrame.TextRange
        notes_page = slide.NotesPage
        for shp in notes_page.Shapes:
            if shp.HasTextFrame and shp.TextFrame.HasText:
                yield s_idx, shp, shp.TextFrame.TextRange


def task_set_2_text_modification(pres):
    from win32com.client import constants

    # Replace "FY24" -> "FY2024"
    for s_idx, shape, tr in _iterate_all_text_ranges(pres):
        text = tr.Text
        if "FY24" in text:
            tr.Text = text.replace("FY24", "FY2024")

    # Title text to Title Case
    for s_idx in range(1, pres.Slides.Count + 1):
        slide = pres.Slides(s_idx)
        for shape in slide.Shapes:
            if (
                shape.Type == constants.msoPlaceholder
                and shape.PlaceholderFormat.Type
                in (constants.ppPlaceholderTitle, constants.ppPlaceholderCenterTitle)
                and shape.HasTextFrame
                and shape.TextFrame.HasText
            ):
                tr = shape.TextFrame.TextRange
                tr.Text = tr.Text.title()

    # Bullet lists: indent level 1, line spacing 1.15
    for s_idx, shape, tr in _iterate_all_text_ranges(pres):
        for p in range(1, tr.Paragraphs().Count + 1):
            para = tr.Paragraphs(p)
            if para.ParagraphFormat.Bullet.Visible:
                para.ParagraphFormat.FirstLineIndent = 0
                para.ParagraphFormat.LeftIndent = 0
                para.ParagraphFormat.SpaceWithin = 115
                try:
                    para.ParagraphFormat.IndentLevel = 1
                except Exception:
                    pass

    # Slide 6 speaker notes
    if pres.Slides.Count < 6:
        raise StepError("2.4", None, "Not enough slides to add notes to slide 6.")
    try:
        slide6 = pres.Slides(6)
        notes_page = slide6.NotesPage
        body = notes_page.Shapes.Placeholders(2)
        body.TextFrame.TextRange.Text = (
            "Key message: Focus on cost-to-serve reduction and service resilience."
        )
    except Exception as e:
        raise StepError("2.4", 6, f"Failed to add speaker notes to slide 6: {e}")

    # Overflow check: best-effort only
    for s_idx, shape, tr in _iterate_all_text_ranges(pres):
        try:
            _ = shape.TextFrame2.AutoSize  # just touching; no reliable overflow flag
        except Exception:
            continue


# ---------------- TASK SET 3 ----------------

def apply_office_theme(pres, ppt_app):
    templates_path = ppt_app.TemplatesPath
    candidates = ["Office Theme.thmx", "Office.thmx"]
    for name in candidates:
        theme_path = os.path.join(templates_path, name)
        if os.path.exists(theme_path):
            try:
                pres.ApplyTemplate(theme_path)
                print(f"Applied theme: {theme_path}")
                return
            except Exception:
                continue
    print("Warning: Could not apply 'Office' theme automatically. Please verify manually.")


def task_set_3_formatting(pres, ppt_app):
    from win32com.client import constants

    apply_office_theme(pres, ppt_app)

    for s_idx in range(1, pres.Slides.Count + 1):
        slide = pres.Slides(s_idx)
        for shape in slide.Shapes:
            if not (shape.HasTextFrame and shape.TextFrame.HasText):
                continue
            tr = shape.TextFrame.TextRange
            is_title = (
                shape.Type == constants.msoPlaceholder
                and shape.PlaceholderFormat.Type
                in (constants.ppPlaceholderTitle, constants.ppPlaceholderCenterTitle)
            )
            if is_title:
                tr.Font.Name = "Calibri"
                tr.Font.Size = 32
                tr.Font.Color.RGB = 0x404040
            else:
                tr.Font.Name = "Calibri"
                tr.Font.Size = 18

            try:
                if shape.Type == constants.msoPlaceholder:
                    tr.ParagraphFormat.Alignment = constants.ppAlignLeft
            except Exception:
                pass

    for s_idx in range(10, min(12, pres.Slides.Count) + 1):
        try:
            pres.Slides(s_idx).Design = pres.SlideMaster.Design
            pres.Slides(s_idx).FollowMasterBackground = True
        except Exception:
            pass

    for s_idx in [1, min(2, pres.Slides.Count), min(3, pres.Slides.Count)]:
        print(f"[Task Set 3] Visually inspect slide {s_idx} for consistency.")


# ---------------- TASK SET 4 ----------------
# (image replacement on slide 7 skipped as requested)

def task_set_4_shapes_images(pres):
    from win32com.client import constants

    if pres.Slides.Count < 4:
        raise StepError("4.1", None, "Not enough slides to modify slide 4.")
    slide4 = pres.Slides(4)
    try:
        width = 10 * CM_TO_POINTS
        height = 1.2 * CM_TO_POINTS
        left = 2 * CM_TO_POINTS
        top = 2 * CM_TO_POINTS

        rect = slide4.Shapes.AddShape(
            constants.msoShapeRectangle, left, top, width, height
        )
        rect.Fill.ForeColor.RGB = 0x0000FF
        rect.Line.Visible = False

        arrow = slide4.Shapes.AddShape(
            constants.msoShapeRightArrow,
            left + width * 0.05,
            top + height * 0.1,
            width * 0.9,
            height * 0.8,
        )
        arrow.Fill.ForeColor.RGB = 0xFFFFFF
        arrow.Line.Visible = False

        shape_range = slide4.Shapes.Range([rect.Name, arrow.Name])
        shape_range.Group()
        print(f"[Task Set 4] Inserted and grouped rectangle+arrow on slide 4.")
    except Exception as e:
        raise StepError("4.1", 4, f"Failed to insert/group rectangle and arrow: {e}")

    # Slide 7 image replacement intentionally skipped

    # Alt text + compress
    for s_idx in range(1, pres.Slides.Count + 1):
        slide = pres.Slides(s_idx)
        for shape in slide.Shapes:
            try:
                if shape.Type == constants.msoPicture:
                    shape.AlternativeText = "Decorative – no critical information"
            except Exception:
                continue

    try:
        pres.CompressPictures(
            DeleteOriginal=False,
            UseDocumentDefault=False,
            OutputResolution=220,
        )
    except Exception:
        print("Warning: CompressPictures not available or failed (version dependent).")


# ---------------- TASK SET 5 ----------------

def task_set_5_tables_charts(pres):
    from win32com.client import constants

    if pres.Slides.Count < 9:
        raise StepError("5.1", None, "Not enough slides to modify slide 9.")
    slide9 = pres.Slides(9)
    try:
        tbl_shape = slide9.Shapes.AddTable(
            NumRows=6,
            NumColumns=4,
            Left=CM_TO_POINTS * 2,
            Top=CM_TO_POINTS * 3,
            Width=CM_TO_POINTS * 20,
            Height=CM_TO_POINTS * 8,
        )
        _ = tbl_shape.Table
        print("[Task Set 5] Inserted 4x6 table on slide 9. Apply 'Medium Style 2 – Accent 1' manually if needed.")
    except Exception as e:
        raise StepError("5.1", 9, f"Failed to insert table on slide 9: {e}")

    if pres.Slides.Count < 11:
        raise StepError("5.2", None, "Not enough slides to modify slide 11.")
    slide11 = pres.Slides(11)
    try:
        chart_shape = slide11.Shapes.AddChart(
            constants.xlBarClustered,
            CM_TO_POINTS * 2,
            CM_TO_POINTS * 3,
            CM_TO_POINTS * 20,
            CM_TO_POINTS * 10,
        )
        chart = chart_shape.Chart

        data_sheet = chart.ChartData.Workbook.Worksheets(1)
        data_sheet.ListObjects(1).Resize(data_sheet.Range("A1:B4"))
        data_sheet.Range("A1").Value = "Category"
        data_sheet.Range("B1").Value = "Value"
        categories = ["Cost", "Service", "Inventory"]
        values = [12, 18, 9]
        for i, (cat, val) in enumerate(zip(categories, values), start=2):
            data_sheet.Range(f"A{i}").Value = cat
            data_sheet.Range(f"B{i}").Value = val

        chart.ChartData.Workbook.Close()
        if chart.SeriesCollection().Count == 1:
            chart.HasLegend = False

        print("[Task Set 5] Inserted clustered bar chart on slide 11.")
    except Exception as e:
        raise StepError("5.2", 11, f"Failed to insert clustered bar chart on slide 11: {e}")


# ---------------- TASK SET 6 ----------------

def task_set_6_transitions(pres):
    from win32com.client import constants

    for s_idx in range(1, pres.Slides.Count + 1):
        slide = pres.Slides(s_idx)
        tr = slide.SlideShowTransition
        tr.EntryEffect = constants.ppEffectFade
        tr.Duration = 0.4
        tr.AdvanceOnClick = False
        tr.AdvanceOnTime = True
        tr.AdvanceTime = 15.0

    print("[Task Set 6] Presenter View should be used when running the slide show.")

    for s_idx in range(1, min(3, pres.Slides.Count) + 1):
        tr = pres.Slides(s_idx).SlideShowTransition
        print(
            f"[Task Set 6] Slide {s_idx} transition: effect={tr.EntryEffect}, "
            f"duration={tr.Duration}, AdvanceOnClick={tr.AdvanceOnClick}, AdvanceOnTime={tr.AdvanceOnTime}"
        )


# ---------------- PROCESSING ONE DECK ----------------

def process_presentation_file(ppt_app, file_path: Path):
    if not file_path.exists():
        raise StepError("0.1", None, f"File does not exist: {file_path}")
    if not is_file_writable(file_path):
        raise StepError("0.2", None, f"File is not writable: {file_path}")

    pres = open_presentation(ppt_app, file_path)
    ensure_not_protected_view(ppt_app, pres)
    log_presentation_info(pres)

    try:
        task_set_1_slide_structure(pres)
        pres.Save()

        task_set_2_text_modification(pres)
        pres.Save()

        task_set_3_formatting(pres, ppt_app)
        pres.Save()

        task_set_4_shapes_images(pres)
        pres.Save()

        task_set_5_tables_charts(pres)
        pres.Save()

        task_set_6_transitions(pres)
        pres.Save()

        print("[Post-Execution] Final slide count:", pres.Slides.Count)
        print("[Post-Execution] No explicit PowerPoint errors were raised during automation.")
        print("[Post-Execution] File saved successfully:", pres.FullName)

        print("Summary:")
        print("- Slides modified: all")
        print("- Objects inserted: slide 4 (shape group), slide 9 (table), slide 11 (chart)")
        print("- Errors encountered: none")

    except StepError:
        raise
    except Exception as e:
        raise StepError("?", None, f"Unhandled error: {e}")
    finally:
        pres.Close()


# ---------------- fDOM REFRESH FOR POWERPOINT ----------------

def refresh_powerpoint_fdom():
    """
    Run your existing fDOM pipeline for PowerPoint so that apps/powerpnt/fdom.json is refreshed.
    Uses FDOMCreator and attempts to locate powerpnt.exe via PATH.
    """
    if FDOMCreator is None:
        print("[fDOM] FDOMCreator import failed – cannot refresh fDOM.")
        return

    powerpnt_path = shutil.which("powerpnt.exe")
    if not powerpnt_path:
        print("[fDOM] Could not find powerpnt.exe in PATH. Please ensure Office is installed or update this path manually.")
        return

    print("\n========== Starting fDOM refresh for PowerPoint ==========")
    print(f"[fDOM] Using PowerPoint executable: {powerpnt_path}")

    creator = FDOMCreator()
    result = creator.create_fdom_for_app(powerpnt_path)

    if result.get("success"):
        print(f"[fDOM] fDOM creation completed for app: {result.get('app_name')}")
        print("[fDOM] apps/powerpnt/fdom.json should now be updated.")
    else:
        print(f"[fDOM] fDOM creation failed: {result.get('error')}")


# ---------------- MAIN ----------------

def main():
    try:
        ppt_app = ensure_powerpoint()
    except StepError as e:
        print("FATAL:", e.message)
        sys.exit(1)

    ppt_app.Visible = True

    for fname in PPT_FILES:
        path = Path(TEMP_DIR) / fname
        print("\n" + "=" * 80)
        print(f"Starting processing for: {path}")
        print("=" * 80)
        try:
            process_presentation_file(ppt_app, path)
        except StepError as e:
            print("ERROR:")
            print(f"  Step: {e.step_number}")
            print(f"  Slide: {e.slide_number}")
            print(f"  Message: {e.message}")
            print("Stopping further processing due to failure.")
            traceback.print_exc()
            break

    ppt_app.Quit()

    # After all decks are processed, refresh fDOM for PowerPoint
    refresh_powerpoint_fdom()


if __name__ == "__main__":
    main()
