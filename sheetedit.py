#!/usr/bin/env python3
"""SheetEdit — lightweight .xlsx editor with call-sheet formatting primitives."""

import sys
import os
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QTabWidget,
    QToolBar, QFileDialog, QColorDialog, QMessageBox, QVBoxLayout, QWidget,
    QLabel, QStatusBar, QMenu, QMenuBar, QSizePolicy, QStyledItemDelegate,
    QStyleOptionViewItem,
)
from PySide6.QtPrintSupport import QPrinter, QPrintDialog, QPrintPreviewDialog
from PySide6.QtGui import QPageLayout
from PySide6.QtCore import Qt, QSize, QRect, QRectF, QModelIndex, QPointF
from PySide6.QtGui import (
    QAction, QColor, QFont, QBrush, QIcon, QKeySequence, QPainter, QPen,
    QShortcut, QFontMetrics,
)

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font as XlFont, PatternFill, Alignment, Border, Side


# ── Theme + Color Handling ───────────────────────────────────────────────────

# Office default theme colors (Office 2007+ "Office" theme)
THEME_COLORS = [
    "#FFFFFF",  # 0  lt1 (background)
    "#000000",  # 1  dk1 (text)
    "#44546A",  # 2  dk2
    "#E7E6E6",  # 3  lt2
    "#4472C4",  # 4  accent1
    "#ED7D31",  # 5  accent2
    "#A5A5A5",  # 6  accent3
    "#FFC000",  # 7  accent4
    "#5B9BD5",  # 8  accent5
    "#70AD47",  # 9  accent6
    "#0563C1",  # 10 hyperlink
    "#954F72",  # 11 followed hyperlink
]


def _apply_tint(color: QColor, tint: float) -> QColor:
    """Apply Excel tint (-1.0 to 1.0) to a QColor. Positive = lighter, negative = darker."""
    r, g, b = color.red(), color.green(), color.blue()
    if tint > 0:
        r = int(r + (255 - r) * tint)
        g = int(g + (255 - g) * tint)
        b = int(b + (255 - b) * tint)
    elif tint < 0:
        factor = 1.0 + tint  # e.g. tint=-0.5 → factor=0.5
        r = int(r * factor)
        g = int(g * factor)
        b = int(b * factor)
    return QColor(min(255, max(0, r)), min(255, max(0, g)), min(255, max(0, b)))


def xl_color_to_qcolor(xl_color, fallback=None):
    """Convert openpyxl color to QColor with full theme + tint support."""
    if xl_color is None:
        return fallback

    # Direct RGB
    if xl_color.type == "rgb" and xl_color.rgb and xl_color.rgb != "00000000":
        rgb = xl_color.rgb
        if len(rgb) == 8:
            qc = QColor(f"#{rgb[2:]}")
        else:
            qc = QColor(f"#{rgb}")
        if xl_color.tint and xl_color.tint != 0:
            qc = _apply_tint(qc, xl_color.tint)
        return qc

    # Theme color
    if xl_color.theme is not None:
        try:
            idx = xl_color.theme.__index__() if hasattr(xl_color.theme, '__index__') else int(str(xl_color.theme))
            if 0 <= idx < len(THEME_COLORS):
                qc = QColor(THEME_COLORS[idx])
            else:
                qc = QColor("#000000")
            if xl_color.tint and xl_color.tint != 0:
                tint_val = float(str(xl_color.tint)) if not isinstance(xl_color.tint, (int, float)) else xl_color.tint
                qc = _apply_tint(qc, tint_val)
            return qc
        except (ValueError, TypeError):
            return fallback

    # Indexed color (legacy)
    if xl_color.type == "indexed" and xl_color.indexed is not None:
        # Common indexed colors
        indexed_colors = {
            0: "#000000", 1: "#FFFFFF", 2: "#FF0000", 3: "#00FF00",
            4: "#0000FF", 5: "#FFFF00", 6: "#FF00FF", 7: "#00FFFF",
            8: "#000000", 9: "#FFFFFF", 10: "#FF0000", 11: "#00FF00",
            12: "#0000FF", 13: "#FFFF00", 14: "#FF00FF", 15: "#00FFFF",
            64: "#000000",  # system foreground
        }
        hex_val = indexed_colors.get(xl_color.indexed)
        if hex_val:
            return QColor(hex_val)

    return fallback


def qcolor_to_xl_rgb(qc: QColor) -> str:
    """QColor -> openpyxl hex string like 'FF4472C4'."""
    return f"FF{qc.red():02X}{qc.green():02X}{qc.blue():02X}"


HALIGN_MAP = {"left": Qt.AlignLeft, "center": Qt.AlignHCenter, "right": Qt.AlignRight}
VALIGN_MAP = {"top": Qt.AlignTop, "center": Qt.AlignVCenter, "bottom": Qt.AlignBottom}


# ── Column/Row Size Conversion ───────────────────────────────────────────────

def xl_col_width_to_px(width, default_font_size=11):
    """Convert Excel column width (character units) to pixels.
    Excel formula: pixels = Truncate(((width * max_digit_width + 5) / max_digit_width * 256) / 256 * max_digit_width)
    Simplified: for Calibri 11pt, max_digit_width ~ 7px. For Arial 11pt ~ 7px.
    """
    if width is None or width <= 0:
        return 64  # default
    # Excel standard: width is in character units of the default font
    # At 11pt, one character ~ 7.5 pixels, plus 5px padding on each side
    max_digit_w = 7.0 + (default_font_size - 11) * 0.5
    px = int(width * max_digit_w + 10)  # 5px padding each side
    return max(px, 20)


def xl_row_height_to_px(height):
    """Convert Excel row height (points) to pixels. 1pt = 1.333px at 96 DPI."""
    if height is None or height <= 0:
        return 21  # default row height
    return int(height * 96.0 / 72.0)


def px_to_xl_col_width(px, default_font_size=11):
    """Reverse of xl_col_width_to_px."""
    max_digit_w = 7.0 + (default_font_size - 11) * 0.5
    return max(0, (px - 10) / max_digit_w)


def px_to_xl_row_height(px):
    """Reverse of xl_row_height_to_px."""
    return px * 72.0 / 96.0


# ── Border Data + Delegate ───────────────────────────────────────────────────

# Border style → pen width and style
BORDER_STYLES = {
    "thin": (1, Qt.SolidLine),
    "medium": (2, Qt.SolidLine),
    "thick": (3, Qt.SolidLine),
    "dashed": (1, Qt.DashLine),
    "dotted": (1, Qt.DotLine),
    "double": (1, Qt.SolidLine),  # approximate
    "hair": (1, Qt.DotLine),
    "mediumDashed": (2, Qt.DashLine),
    "dashDot": (1, Qt.DashDotLine),
    "mediumDashDot": (2, Qt.DashDotLine),
    "dashDotDot": (1, Qt.DashDotDotLine),
    "mediumDashDotDot": (2, Qt.DashDotDotLine),
}


def _side_to_pen(side):
    """Convert openpyxl Side to (color_hex, width, style) or None."""
    if side is None or side.style is None or side.style == "none":
        return None
    color = "#000000"
    if side.color:
        qc = xl_color_to_qcolor(side.color)
        if qc and qc.isValid():
            color = qc.name()
    width, style = BORDER_STYLES.get(side.style, (1, Qt.SolidLine))
    return (color, width, style)


# Store border info as a role on QTableWidgetItem
BORDER_ROLE = Qt.UserRole + 100  # dict: {top, bottom, left, right} → (color, width, style)


class BorderDelegate(QStyledItemDelegate):
    """Custom delegate that paints cell borders from stored border data."""

    def paint(self, painter: QPainter, option: QStyleOptionViewItem, index: QModelIndex):
        # Draw default content first
        super().paint(painter, option, index)

        # Get border data
        border_data = index.data(BORDER_ROLE)
        if not border_data:
            return

        painter.save()
        rect = option.rect

        for side_name, edge in [
            ("top", (rect.left(), rect.top(), rect.right(), rect.top())),
            ("bottom", (rect.left(), rect.bottom(), rect.right(), rect.bottom())),
            ("left", (rect.left(), rect.top(), rect.left(), rect.bottom())),
            ("right", (rect.right(), rect.top(), rect.right(), rect.bottom())),
        ]:
            info = border_data.get(side_name)
            if info:
                color_hex, width, style = info
                pen = QPen(QColor(color_hex), width, style)
                painter.setPen(pen)
                painter.drawLine(*edge)

        painter.restore()


# ── SheetView — one tab ─────────────────────────────────────────────────────

class SheetView(QTableWidget):
    """Displays one openpyxl worksheet in a QTableWidget."""

    def __init__(self, ws, parent=None):
        super().__init__(parent)
        self.ws = ws
        self.merges = []
        self.setItemDelegate(BorderDelegate(self))
        self.setShowGrid(True)  # light grid from stylesheet, borders drawn on top
        self._load()

    def _load(self):
        ws = self.ws
        rows = ws.max_row or 1
        cols = ws.max_column or 1
        self.setRowCount(rows)
        self.setColumnCount(cols)

        # Column headers
        self.setHorizontalHeaderLabels(
            [get_column_letter(c + 1) for c in range(cols)]
        )

        # Column widths
        for ci in range(cols):
            col_letter = get_column_letter(ci + 1)
            if col_letter in ws.column_dimensions:
                w = ws.column_dimensions[col_letter].width
                if w:
                    self.setColumnWidth(ci, xl_col_width_to_px(w))

        # Row heights
        for ri in range(rows):
            if (ri + 1) in ws.row_dimensions:
                h = ws.row_dimensions[ri + 1].height
                if h:
                    self.setRowHeight(ri, xl_row_height_to_px(h))

        # Cells
        for ri in range(rows):
            for ci in range(cols):
                cell = ws.cell(row=ri + 1, column=ci + 1)
                item = QTableWidgetItem()

                # Value
                val = cell.value
                if val is not None:
                    item.setText(str(val))

                # Font
                qf = QFont()
                if cell.font:
                    qf.setFamily(cell.font.name or "Arial")
                    qf.setPointSize(cell.font.size or 11)
                    qf.setBold(cell.font.bold or False)
                    qf.setItalic(cell.font.italic or False)
                    qf.setUnderline(
                        cell.font.underline is not None
                        and cell.font.underline != "none"
                    )
                    fc = xl_color_to_qcolor(cell.font.color, QColor("#202124"))
                    item.setForeground(QBrush(fc if fc else QColor("#202124")))
                else:
                    item.setForeground(QBrush(QColor("#202124")))
                item.setFont(qf)

                # Fill
                has_fill = False
                if cell.fill and cell.fill.fgColor:
                    bg = xl_color_to_qcolor(cell.fill.fgColor)
                    if bg and bg.isValid() and bg != QColor(0, 0, 0):
                        item.setBackground(QBrush(bg))
                        has_fill = True
                if not has_fill:
                    item.setBackground(QBrush(QColor("#FFFFFF")))

                # Alignment
                flags = Qt.AlignLeft | Qt.AlignVCenter
                if cell.alignment:
                    h = HALIGN_MAP.get(cell.alignment.horizontal, Qt.AlignLeft)
                    v = VALIGN_MAP.get(cell.alignment.vertical, Qt.AlignVCenter)
                    flags = h | v
                    if cell.alignment.wrap_text:
                        flags |= Qt.TextWordWrap
                item.setTextAlignment(flags)

                # Borders
                if cell.border:
                    bd = {}
                    for side_name in ("top", "bottom", "left", "right"):
                        side = getattr(cell.border, side_name, None)
                        pen_info = _side_to_pen(side)
                        if pen_info:
                            bd[side_name] = pen_info
                    if bd:
                        item.setData(BORDER_ROLE, bd)

                self.setItem(ri, ci, item)

        # Merged cells
        for rng in ws.merged_cells.ranges:
            r1, c1 = rng.min_row - 1, rng.min_col - 1
            r2, c2 = rng.max_row - 1, rng.max_col - 1
            self.merges.append((r1, c1, r2, c2))
            self.setSpan(r1, c1, r2 - r1 + 1, c2 - c1 + 1)

    # ── Write back ───────────────────────────────────────────────────────

    def sync_to_ws(self):
        """Push current table state back to the openpyxl worksheet."""
        ws = self.ws

        # Unmerge all existing
        for rng in list(ws.merged_cells.ranges):
            ws.unmerge_cells(str(rng))

        for ri in range(self.rowCount()):
            for ci in range(self.columnCount()):
                cell = ws.cell(row=ri + 1, column=ci + 1)
                item = self.item(ri, ci)
                if item is None:
                    cell.value = None
                    continue

                cell.value = item.text() or None

                # Font
                qf = item.font()
                fg_brush = item.foreground()
                fg_color = (
                    fg_brush.color()
                    if fg_brush != QBrush()
                    else QColor("#202124")
                )
                cell.font = XlFont(
                    name=qf.family(),
                    size=qf.pointSize(),
                    bold=qf.bold(),
                    italic=qf.italic(),
                    underline="single" if qf.underline() else None,
                    color=qcolor_to_xl_rgb(fg_color),
                )

                # Fill
                bg_brush = item.background()
                if bg_brush != QBrush() and bg_brush.color().isValid():
                    bgc = bg_brush.color()
                    if bgc != QColor("#FFFFFF"):
                        cell.fill = PatternFill(
                            start_color=qcolor_to_xl_rgb(bgc),
                            end_color=qcolor_to_xl_rgb(bgc),
                            fill_type="solid",
                        )
                    else:
                        cell.fill = PatternFill(fill_type=None)
                else:
                    cell.fill = PatternFill(fill_type=None)

                # Alignment
                flags = item.textAlignment()
                ha = "left"
                if flags & Qt.AlignHCenter:
                    ha = "center"
                elif flags & Qt.AlignRight:
                    ha = "right"
                va = "center"
                if flags & Qt.AlignTop:
                    va = "top"
                elif flags & Qt.AlignBottom:
                    va = "bottom"
                wrap = bool(flags & Qt.TextWordWrap)
                cell.alignment = Alignment(
                    horizontal=ha, vertical=va, wrap_text=wrap
                )

                # Borders
                bd = item.data(BORDER_ROLE)
                if bd:
                    sides = {}
                    for side_name in ("top", "bottom", "left", "right"):
                        info = bd.get(side_name)
                        if info:
                            color_hex, width, _ = info
                            style = "thin"
                            if width >= 3:
                                style = "thick"
                            elif width >= 2:
                                style = "medium"
                            sides[side_name] = Side(
                                style=style,
                                color=qcolor_to_xl_rgb(QColor(color_hex)),
                            )
                        else:
                            sides[side_name] = Side(style=None)
                    cell.border = Border(**sides)

        # Re-merge
        for r1, c1, r2, c2 in self.merges:
            ws.merge_cells(
                start_row=r1 + 1,
                start_column=c1 + 1,
                end_row=r2 + 1,
                end_column=c2 + 1,
            )

        # Column widths
        for ci in range(self.columnCount()):
            col_letter = get_column_letter(ci + 1)
            ws.column_dimensions[col_letter].width = px_to_xl_col_width(
                self.columnWidth(ci)
            )

        # Row heights
        for ri in range(self.rowCount()):
            ws.row_dimensions[ri + 1].height = px_to_xl_row_height(
                self.rowHeight(ri)
            )


# ── Main Window ──────────────────────────────────────────────────────────────

class SheetEditWindow(QMainWindow):
    def __init__(self, filepath=None):
        super().__init__()
        self.setWindowTitle("SheetEdit")
        self.resize(1200, 800)
        self.wb = None
        self.filepath = None

        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage("Ready")

        self._build_menu()
        self._build_toolbar()

        if filepath:
            self.open_file(filepath)
        else:
            self._new_workbook()

    # ── Menu bar ─────────────────────────────────────────────────────────

    def _build_menu(self):
        mb = self.menuBar()

        file_menu = mb.addMenu("File")

        new_act = QAction("New", self)
        new_act.setShortcut(QKeySequence.New)
        new_act.triggered.connect(self._new_workbook)
        file_menu.addAction(new_act)

        open_act = QAction("Open...", self)
        open_act.setShortcut(QKeySequence.Open)
        open_act.triggered.connect(self._open_dialog)
        file_menu.addAction(open_act)

        save_act = QAction("Save", self)
        save_act.setShortcut(QKeySequence.Save)
        save_act.triggered.connect(self._save)
        file_menu.addAction(save_act)

        saveas_act = QAction("Save As...", self)
        saveas_act.setShortcut(QKeySequence("Ctrl+Shift+S"))
        saveas_act.triggered.connect(self._save_as)
        file_menu.addAction(saveas_act)

        file_menu.addSeparator()

        preview_act = QAction("Print Preview...", self)
        preview_act.triggered.connect(self._print_preview)
        file_menu.addAction(preview_act)

        print_act = QAction("Print...", self)
        print_act.setShortcut(QKeySequence.Print)
        print_act.triggered.connect(self._print)
        file_menu.addAction(print_act)

        edit_menu = mb.addMenu("Edit")

        ins_row = QAction("Insert Row", self)
        ins_row.triggered.connect(self.cmd_insert_row)
        edit_menu.addAction(ins_row)

        del_row = QAction("Delete Row", self)
        del_row.triggered.connect(self.cmd_delete_row)
        edit_menu.addAction(del_row)

        ins_col = QAction("Insert Column", self)
        ins_col.triggered.connect(self.cmd_insert_col)
        edit_menu.addAction(ins_col)

        del_col = QAction("Delete Column", self)
        del_col.triggered.connect(self.cmd_delete_col)
        edit_menu.addAction(del_col)

    # ── Toolbar ──────────────────────────────────────────────────────────

    def _build_toolbar(self):
        tb = self.addToolBar("Format")
        tb.setIconSize(QSize(20, 20))
        tb.setMovable(False)

        # Bold
        bold_act = QAction("B", self)
        bold_act.setToolTip("Bold (Ctrl+B)")
        bold_act.setShortcut(QKeySequence.Bold)
        bold_act.triggered.connect(self.cmd_bold)
        f = bold_act.font()
        f.setBold(True)
        bold_act.setFont(f)
        tb.addAction(bold_act)

        # Italic
        italic_act = QAction("I", self)
        italic_act.setToolTip("Italic (Ctrl+I)")
        italic_act.setShortcut(QKeySequence.Italic)
        italic_act.triggered.connect(self.cmd_italic)
        f = italic_act.font()
        f.setItalic(True)
        italic_act.setFont(f)
        tb.addAction(italic_act)

        tb.addSeparator()

        # Fill color
        fill_act = QAction("Fill", self)
        fill_act.setToolTip("Fill Color")
        fill_act.triggered.connect(self.cmd_fill_color)
        tb.addAction(fill_act)

        # Font color
        font_act = QAction("Color", self)
        font_act.setToolTip("Font Color")
        font_act.triggered.connect(self.cmd_font_color)
        tb.addAction(font_act)

        tb.addSeparator()

        # Horizontal alignment
        for label, align in [
            ("\u2190", "left"),
            ("\u2194", "center"),
            ("\u2192", "right"),
        ]:
            act = QAction(label, self)
            act.setToolTip(f"Align {align.title()}")
            act.triggered.connect(lambda checked, a=align: self.cmd_halign(a))
            tb.addAction(act)

        tb.addSeparator()

        # Vertical alignment
        for label, align in [
            ("\u2191", "top"),
            ("\u2195", "center"),
            ("\u2193", "bottom"),
        ]:
            act = QAction(label, self)
            act.setToolTip(f"Vertical {align.title()}")
            act.triggered.connect(lambda checked, a=align: self.cmd_valign(a))
            tb.addAction(act)

        tb.addSeparator()

        # Wrap
        wrap_act = QAction("Wrap", self)
        wrap_act.setToolTip("Toggle Wrap Text")
        wrap_act.triggered.connect(self.cmd_wrap)
        tb.addAction(wrap_act)

        # Merge / Unmerge
        merge_act = QAction("Merge", self)
        merge_act.setToolTip("Merge Cells")
        merge_act.triggered.connect(self.cmd_merge)
        tb.addAction(merge_act)

        unmerge_act = QAction("Unmerge", self)
        unmerge_act.setToolTip("Unmerge Cells")
        unmerge_act.triggered.connect(self.cmd_unmerge)
        tb.addAction(unmerge_act)

        tb.addSeparator()

        # Borders
        border_act = QAction("Borders", self)
        border_act.setToolTip("Add Thin Borders to Selection")
        border_act.triggered.connect(self.cmd_borders)
        tb.addAction(border_act)

        thick_border_act = QAction("Thick", self)
        thick_border_act.setToolTip("Add Thick Borders to Selection")
        thick_border_act.triggered.connect(self.cmd_thick_borders)
        tb.addAction(thick_border_act)

        clear_border_act = QAction("No Border", self)
        clear_border_act.setToolTip("Clear Borders")
        clear_border_act.triggered.connect(self.cmd_clear_borders)
        tb.addAction(clear_border_act)

        tb.addSeparator()

        # Force text
        text_act = QAction("Force Text", self)
        text_act.setToolTip("Force selected cells to text format")
        text_act.triggered.connect(self.cmd_force_text)
        tb.addAction(text_act)

    # ── File I/O ─────────────────────────────────────────────────────────

    def _new_workbook(self):
        self.wb = openpyxl.Workbook()
        self.filepath = None
        self.setWindowTitle("SheetEdit — New Workbook")
        self._reload_tabs()

    def _open_dialog(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Spreadsheet", "", "Excel Files (*.xlsx);;All Files (*)"
        )
        if path:
            self.open_file(path)

    def open_file(self, path):
        try:
            self.wb = openpyxl.load_workbook(path)
            self.filepath = path
            self.setWindowTitle(f"SheetEdit \u2014 {Path(path).name}")
            self._reload_tabs()
            self.statusBar().showMessage(f"Opened {Path(path).name}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to open file:\n{e}")

    def _reload_tabs(self):
        self.tabs.clear()
        for name in self.wb.sheetnames:
            sv = SheetView(self.wb[name], self)
            self.tabs.addTab(sv, name)

    def _save(self):
        if self.filepath:
            self._write(self.filepath)
        else:
            self._save_as()

    def _save_as(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Spreadsheet", "", "Excel Files (*.xlsx)"
        )
        if path:
            if not path.endswith(".xlsx"):
                path += ".xlsx"
            self.filepath = path
            self._write(path)

    def _write(self, path):
        try:
            for i in range(self.tabs.count()):
                sv = self.tabs.widget(i)
                sv.sync_to_ws()
            self.wb.save(path)
            self.setWindowTitle(f"SheetEdit \u2014 {Path(path).name}")
            self.statusBar().showMessage(f"Saved to {Path(path).name}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save:\n{e}")

    # ── Print ────────────────────────────────────────────────────────────

    def _print_preview(self):
        sv = self._sheet()
        if sv is None:
            return
        printer = QPrinter(QPrinter.HighResolution)
        printer.setPageOrientation(QPageLayout.Landscape)
        dialog = QPrintPreviewDialog(printer, self)
        dialog.paintRequested.connect(lambda p: self._render_sheet(p, sv))
        dialog.exec()

    def _print(self):
        sv = self._sheet()
        if sv is None:
            return
        printer = QPrinter(QPrinter.HighResolution)
        printer.setPageOrientation(QPageLayout.Landscape)
        dialog = QPrintDialog(printer, self)
        if dialog.exec() == QPrintDialog.Accepted:
            self._render_sheet(printer, sv)

    def _render_sheet(self, printer, sv: "SheetView"):
        """Render the current sheet onto a QPrinter with pagination."""
        painter = QPainter()
        if not painter.begin(printer):
            return

        page_rect = printer.pageRect(QPrinter.DevicePixel)
        pw = page_rect.width()
        ph = page_rect.height()

        # Gather column widths and row heights in screen pixels
        col_widths = [sv.columnWidth(c) for c in range(sv.columnCount())]
        row_heights = [sv.rowHeight(r) for r in range(sv.rowCount())]

        # Scale: fit all columns to page width
        total_col_w = sum(col_widths)
        if total_col_w <= 0:
            painter.end()
            return
        scale = pw / total_col_w
        # Cap scale so rows aren't absurdly large
        max_scale = ph / 40  # at least 40px row equivalent per page
        if scale > max_scale:
            scale = max_scale

        scaled_col_widths = [w * scale for w in col_widths]
        scaled_row_heights = [h * scale for h in row_heights]

        # Paginate rows
        pages = []  # list of (start_row, end_row)
        current_y = 0
        page_start = 0
        for ri, rh in enumerate(scaled_row_heights):
            if current_y + rh > ph and ri > page_start:
                pages.append((page_start, ri - 1))
                page_start = ri
                current_y = rh
            else:
                current_y += rh
        pages.append((page_start, len(scaled_row_heights) - 1))

        for page_idx, (row_start, row_end) in enumerate(pages):
            if page_idx > 0:
                printer.newPage()

            y = 0.0
            for ri in range(row_start, row_end + 1):
                x = 0.0
                rh = scaled_row_heights[ri]

                for ci in range(sv.columnCount()):
                    cw = scaled_col_widths[ci]
                    cell_rect = QRectF(x, y, cw, rh)

                    item = sv.item(ri, ci)
                    if item:
                        # Fill
                        bg = item.background()
                        if bg != QBrush() and bg.color().isValid() and bg.color() != QColor("#FFFFFF"):
                            painter.fillRect(cell_rect, bg)

                        # Borders
                        bd = item.data(BORDER_ROLE)
                        if bd:
                            for side_name, coords in [
                                ("top", (x, y, x + cw, y)),
                                ("bottom", (x, y + rh, x + cw, y + rh)),
                                ("left", (x, y, x, y + rh)),
                                ("right", (x + cw, y, x + cw, y + rh)),
                            ]:
                                info = bd.get(side_name)
                                if info:
                                    color_hex, width, style = info
                                    pen = QPen(QColor(color_hex), width * scale * 0.5, style)
                                    painter.setPen(pen)
                                    painter.drawLine(QPointF(coords[0], coords[1]), QPointF(coords[2], coords[3]))

                        # Text
                        text = item.text()
                        if text:
                            qf = QFont(item.font())
                            qf.setPointSizeF(qf.pointSizeF() * scale / printer.logicalDpiY() * 72.0)
                            painter.setFont(qf)

                            fg = item.foreground()
                            if fg != QBrush():
                                painter.setPen(QPen(fg.color()))
                            else:
                                painter.setPen(QPen(QColor("#202124")))

                            flags = item.textAlignment()
                            text_rect = cell_rect.adjusted(3 * scale, 1 * scale, -3 * scale, -1 * scale)
                            painter.drawText(text_rect, flags, text)

                    # Grid line
                    painter.setPen(QPen(QColor("#E2E2E2"), 0.5))
                    painter.drawRect(cell_rect)

                    x += cw
                y += rh

        painter.end()

    # ── Selection helpers ────────────────────────────────────────────────

    def _sheet(self) -> SheetView:
        return self.tabs.currentWidget()

    def _selected_items(self):
        sv = self._sheet()
        if sv is None:
            return []
        return sv.selectedItems()

    def _selected_range(self):
        sv = self._sheet()
        if sv is None:
            return None
        ranges = sv.selectedRanges()
        if not ranges:
            return None
        r = ranges[0]
        return (r.topRow(), r.leftColumn(), r.bottomRow(), r.rightColumn())

    # ── Format commands ──────────────────────────────────────────────────

    def cmd_bold(self):
        for item in self._selected_items():
            f = item.font()
            f.setBold(not f.bold())
            item.setFont(f)

    def cmd_italic(self):
        for item in self._selected_items():
            f = item.font()
            f.setItalic(not f.italic())
            item.setFont(f)

    def cmd_fill_color(self):
        color = QColorDialog.getColor(QColor(Qt.white), self, "Fill Color")
        if color.isValid():
            for item in self._selected_items():
                item.setBackground(QBrush(color))

    def cmd_font_color(self):
        color = QColorDialog.getColor(QColor(Qt.black), self, "Font Color")
        if color.isValid():
            for item in self._selected_items():
                item.setForeground(QBrush(color))

    def cmd_halign(self, align):
        flag = HALIGN_MAP[align]
        for item in self._selected_items():
            cur = item.textAlignment()
            cur &= ~(Qt.AlignLeft | Qt.AlignHCenter | Qt.AlignRight)
            item.setTextAlignment(cur | flag)

    def cmd_valign(self, align):
        flag = VALIGN_MAP[align]
        for item in self._selected_items():
            cur = item.textAlignment()
            cur &= ~(Qt.AlignTop | Qt.AlignVCenter | Qt.AlignBottom)
            item.setTextAlignment(cur | flag)

    def cmd_wrap(self):
        for item in self._selected_items():
            cur = item.textAlignment()
            item.setTextAlignment(cur ^ Qt.TextWordWrap)

    def cmd_merge(self):
        rng = self._selected_range()
        if rng is None:
            return
        r1, c1, r2, c2 = rng
        if r1 == r2 and c1 == c2:
            return
        sv = self._sheet()
        sv.setSpan(r1, c1, r2 - r1 + 1, c2 - c1 + 1)
        sv.merges.append((r1, c1, r2, c2))
        self.statusBar().showMessage(
            f"Merged {get_column_letter(c1+1)}{r1+1}:{get_column_letter(c2+1)}{r2+1}"
        )

    def cmd_unmerge(self):
        rng = self._selected_range()
        if rng is None:
            return
        r1, c1, r2, c2 = rng
        sv = self._sheet()
        to_remove = []
        for m in sv.merges:
            if m[0] <= r2 and m[2] >= r1 and m[1] <= c2 and m[3] >= c1:
                sv.setSpan(m[0], m[1], 1, 1)
                to_remove.append(m)
        for m in to_remove:
            sv.merges.remove(m)
        self.statusBar().showMessage("Unmerged")

    def _apply_borders(self, color_hex, width):
        """Apply border to all selected cells."""
        rng = self._selected_range()
        if rng is None:
            return
        r1, c1, r2, c2 = rng
        sv = self._sheet()
        style = Qt.SolidLine
        for ri in range(r1, r2 + 1):
            for ci in range(c1, c2 + 1):
                item = sv.item(ri, ci)
                if item is None:
                    item = QTableWidgetItem()
                    sv.setItem(ri, ci, item)
                bd = item.data(BORDER_ROLE) or {}
                # Top edge
                if ri == r1:
                    bd["top"] = (color_hex, width, style)
                # Bottom edge
                if ri == r2:
                    bd["bottom"] = (color_hex, width, style)
                # Left edge
                if ci == c1:
                    bd["left"] = (color_hex, width, style)
                # Right edge
                if ci == c2:
                    bd["right"] = (color_hex, width, style)
                # Inner borders
                if ri < r2:
                    bd["bottom"] = (color_hex, width, style)
                if ri > r1:
                    bd["top"] = (color_hex, width, style)
                if ci < c2:
                    bd["right"] = (color_hex, width, style)
                if ci > c1:
                    bd["left"] = (color_hex, width, style)
                item.setData(BORDER_ROLE, bd)
        sv.viewport().update()
        self.statusBar().showMessage("Borders applied")

    def cmd_borders(self):
        self._apply_borders("#BFBFBF", 1)

    def cmd_thick_borders(self):
        self._apply_borders("#000000", 2)

    def cmd_clear_borders(self):
        for item in self._selected_items():
            item.setData(BORDER_ROLE, None)
        sv = self._sheet()
        if sv:
            sv.viewport().update()
        self.statusBar().showMessage("Borders cleared")

    def cmd_force_text(self):
        for item in self._selected_items():
            txt = item.text()
            item.setText(txt)
        self.statusBar().showMessage("Cells forced to text")

    # ── Row/Column insert/delete ─────────────────────────────────────────

    def cmd_insert_row(self):
        sv = self._sheet()
        if sv is None:
            return
        row = sv.currentRow()
        sv.insertRow(row)
        self.statusBar().showMessage(f"Inserted row at {row + 1}")

    def cmd_delete_row(self):
        sv = self._sheet()
        if sv is None:
            return
        row = sv.currentRow()
        sv.removeRow(row)
        self.statusBar().showMessage(f"Deleted row {row + 1}")

    def cmd_insert_col(self):
        sv = self._sheet()
        if sv is None:
            return
        col = sv.currentColumn()
        sv.insertColumn(col)
        self.statusBar().showMessage(
            f"Inserted column at {get_column_letter(col + 1)}"
        )

    def cmd_delete_col(self):
        sv = self._sheet()
        if sv is None:
            return
        col = sv.currentColumn()
        sv.removeColumn(col)
        self.statusBar().showMessage(
            f"Deleted column {get_column_letter(col + 1)}"
        )


# ── Stylesheet ───────────────────────────────────────────────────────────────

LIGHT_STYLESHEET = """
QMainWindow, QWidget {
    background-color: #FFFFFF;
    color: #202124;
}
QMenuBar {
    background-color: #F8F9FA;
    border-bottom: 1px solid #DADCE0;
    color: #202124;
    padding: 2px;
}
QMenuBar::item:selected {
    background-color: #E8EAED;
    border-radius: 4px;
}
QMenu {
    background-color: #FFFFFF;
    border: 1px solid #DADCE0;
    color: #202124;
}
QMenu::item:selected {
    background-color: #E8F0FE;
}
QToolBar {
    background-color: #F1F3F4;
    border-bottom: 1px solid #DADCE0;
    spacing: 2px;
    padding: 3px 6px;
}
QToolBar QToolButton {
    background-color: transparent;
    border: 1px solid transparent;
    border-radius: 4px;
    padding: 4px 8px;
    color: #444746;
    font-size: 13px;
}
QToolBar QToolButton:hover {
    background-color: #E8EAED;
    border: 1px solid #DADCE0;
}
QToolBar QToolButton:pressed {
    background-color: #D2E3FC;
}
QToolBar::separator {
    width: 1px;
    background-color: #DADCE0;
    margin: 4px 4px;
}
QTabWidget::pane {
    border: none;
    background-color: #FFFFFF;
}
QTabBar::tab {
    background-color: #F1F3F4;
    border: 1px solid #DADCE0;
    border-bottom: none;
    padding: 6px 16px;
    margin-right: 2px;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
    color: #5F6368;
    font-size: 12px;
}
QTabBar::tab:selected {
    background-color: #FFFFFF;
    color: #1A73E8;
    font-weight: bold;
    border-bottom: 2px solid #1A73E8;
}
QTabBar::tab:hover:!selected {
    background-color: #E8EAED;
}
QTableWidget {
    background-color: #FFFFFF;
    gridline-color: #E2E2E2;
    border: none;
    selection-background-color: #D2E3FC;
    selection-color: #202124;
    color: #202124;
    font-family: Arial;
    font-size: 13px;
}
QTableWidget::item {
    padding: 1px 3px;
}
QTableWidget::item:selected {
    background-color: #D2E3FC;
    color: #202124;
}
QTableWidget QTableCornerButton::section {
    background-color: #F8F9FA;
    border: 1px solid #E2E2E2;
}
QHeaderView::section {
    background-color: #F8F9FA;
    color: #5F6368;
    border: 1px solid #E2E2E2;
    padding: 4px;
    font-family: Arial;
    font-size: 12px;
    font-weight: 500;
}
QHeaderView::section:checked {
    background-color: #D2E3FC;
    color: #1A73E8;
}
QStatusBar {
    background-color: #F8F9FA;
    border-top: 1px solid #DADCE0;
    color: #5F6368;
    font-size: 12px;
}
QScrollBar:vertical {
    background: #F8F9FA;
    width: 12px;
    border: none;
}
QScrollBar::handle:vertical {
    background: #BDC1C6;
    border-radius: 6px;
    min-height: 30px;
}
QScrollBar::handle:vertical:hover {
    background: #9AA0A6;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0;
}
QScrollBar:horizontal {
    background: #F8F9FA;
    height: 12px;
    border: none;
}
QScrollBar::handle:horizontal {
    background: #BDC1C6;
    border-radius: 6px;
    min-width: 30px;
}
QScrollBar::handle:horizontal:hover {
    background: #9AA0A6;
}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
    width: 0;
}
"""


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    app = QApplication(sys.argv)
    app.setApplicationName("SheetEdit")
    app.setStyle("Fusion")
    app.setStyleSheet(LIGHT_STYLESHEET)

    filepath = sys.argv[1] if len(sys.argv) > 1 else None
    win = SheetEditWindow(filepath)
    win.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
