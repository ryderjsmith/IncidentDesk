"""PDF export of the incident board via reportlab."""
from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import List

from .dates import fmt_dt
from .db import DB
from .dialogs import dark_info


class Exporter:
    def __init__(self, db: DB):
        self.db = db

    def _notes_text(self, inc_id: int) -> str:
        """Return all notes for an incident as plain text, one per line."""
        notes = self.db.list_notes(inc_id)
        return "\n".join(f"{fmt_dt(n['ts'])}: {n['body']}" for n in notes)

    def _billables_text(self, inc_id: int) -> str:
        """Return all billables for an incident as plain text, one per line."""
        return "\n".join(b["body"] for b in self.db.list_billables(inc_id))

    def export_pdf(self, rows: List[sqlite3.Row], path: Path, parent=None, title: str = "Incident Board"):
        try:
            from reportlab.lib.pagesizes import letter, landscape
            from reportlab.lib import colors
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
            from reportlab.lib.styles import getSampleStyleSheet
        except Exception:
            dark_info(parent, "Missing dependency",
                      "reportlab is not installed. Run\n  pip install reportlab")
            return

        doc = SimpleDocTemplate(str(path), pagesize=landscape(letter),
                                leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36,
                                title=title)
        styles = getSampleStyleSheet()

        from reportlab.lib.styles import ParagraphStyle
        cell_style = ParagraphStyle('cell', fontSize=7, leading=9, wordWrap='LTR')
        hdr_style  = ParagraphStyle('hdr',  fontSize=7, leading=9, fontName='Helvetica-Bold')

        def P(text, style=cell_style):
            return Paragraph(str(text).replace("\n", "<br/>"), style)

        headers = ["Reported", "Dispatched", "Arrived", "Cleared", "Type", "Location", "Car #", "Driver Code", "Unit", "Status", "Notes", "Billables"]
        data = [[P(h, hdr_style) for h in headers]]
        for r in rows:
            notes_text = self._notes_text(r["id"])
            billables_text = self._billables_text(r["id"])
            data.append([
                P(fmt_dt(r["reported_at"])), P(fmt_dt(r["dispatched_at"])), P(fmt_dt(r["arrived_at"])), P(fmt_dt(r["cleared_at"])),
                P(r["type"]), P(r["location_name"] or ""), P(r["car_number"] or ""), P(r["driver_code"] or ""), P(r["primary_units"] or ""),
                P("Cleared" if r["is_cleared"] else "Active"),
                P(notes_text),
                P(billables_text),
            ])

        # Landscape letter usable width ≈ 720pt (11in × 72 − 2×36 margins)
        col_widths = [66, 62, 62, 62, 52, 68, 38, 55, 52, 38, 88, 77]  # sum = 720
        table = Table(data, colWidths=col_widths, repeatRows=1)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#d0d0d0')),
            ('GRID', (0, 0), (-1, -1), 0.25, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#ebebeb')]),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('LEFTPADDING', (0, 0), (-1, -1), 3),
            ('RIGHTPADDING', (0, 0), (-1, -1), 3),
        ]))
        story = [Paragraph(title, styles['Title']), table]
        doc.build(story)
        dark_info(parent, "Exported", f"Saved PDF to\n{path}")
