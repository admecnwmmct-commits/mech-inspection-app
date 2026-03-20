from flask import Flask, render_template, request, session, send_file, redirect, url_for, jsonify
import os, io, json
from datetime import datetime
from checklist_data import SECTIONS

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "mech-bcr-inspection-2026")

# ── PDF generation ────────────────────────────────────────────────────────────
def generate_pdf(header, responses):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, HRFlowable
    from reportlab.lib.enums import TA_CENTER, TA_LEFT

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=1.8*cm, rightMargin=1.8*cm,
                            topMargin=2*cm, bottomMargin=2*cm)

    navy   = colors.HexColor('#0a1628')
    saffron= colors.HexColor('#f4a014')
    steel  = colors.HexColor('#3a7bd5')
    lgray  = colors.HexColor('#f5f5f5')
    mgray  = colors.HexColor('#e0e0e0')
    green  = colors.HexColor('#16a34a')
    red    = colors.HexColor('#dc2626')
    white  = colors.white
    black  = colors.black

    styles = getSampleStyleSheet()
    title_style  = ParagraphStyle('title',  fontSize=14, fontName='Helvetica-Bold',  textColor=white,   alignment=TA_CENTER, spaceAfter=4)
    sub_style    = ParagraphStyle('sub',    fontSize=9,  fontName='Helvetica',        textColor=saffron, alignment=TA_CENTER, spaceAfter=2)
    head2_style  = ParagraphStyle('head2',  fontSize=11, fontName='Helvetica-Bold',   textColor=navy,    spaceBefore=10, spaceAfter=4)
    head3_style  = ParagraphStyle('head3',  fontSize=9,  fontName='Helvetica-Bold',   textColor=steel,   spaceBefore=6, spaceAfter=3)
    body_style   = ParagraphStyle('body',   fontSize=8,  fontName='Helvetica',        textColor=black,   leading=12)
    remark_style = ParagraphStyle('remark', fontSize=7.5,fontName='Helvetica-Oblique',textColor=colors.HexColor('#444444'), leading=11)
    small_style  = ParagraphStyle('small',  fontSize=8,  fontName='Helvetica',        textColor=colors.HexColor('#444444'))

    story = []

    # ── Header banner ──
    W = 17.4*cm
    hdr_data = [[Paragraph("MECHANICAL DEPARTMENT INSPECTION REPORT", title_style)],
                [Paragraph("Office of Sr. DME (Co.) BCT · Western Railway · Mumbai Central", sub_style)]]
    hdr_tbl = Table(hdr_data, colWidths=[W])
    hdr_tbl.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), navy),
        ('TOPPADDING', (0,0), (-1,-1), 10),
        ('BOTTOMPADDING', (0,0), (-1,-1), 10),
        ('LEFTPADDING', (0,0), (-1,-1), 12),
        ('RIGHTPADDING', (0,0), (-1,-1), 12),
        ('LINEBELOW', (0,-1), (-1,-1), 3, saffron),
    ]))
    story.append(hdr_tbl)
    story.append(Spacer(1, 0.4*cm))

    # ── Inspection details table ──
    details = [
        ["Inspector Name:", header.get('name',''),      "Designation:", header.get('designation','')],
        ["Location:",       header.get('location',''),  "Date:",         header.get('date','')],
        ["Inspection Type:",header.get('insp_type',''), "Division:",      "BCT / Western Railway"],
    ]
    det_tbl = Table(details, colWidths=[3.8*cm, 4.5*cm, 3.5*cm, 5.6*cm])
    det_tbl.setStyle(TableStyle([
        ('FONTNAME',  (0,0), (0,-1), 'Helvetica-Bold'),
        ('FONTNAME',  (2,0), (2,-1), 'Helvetica-Bold'),
        ('FONTSIZE',  (0,0), (-1,-1), 8.5),
        ('TEXTCOLOR', (0,0), (0,-1), navy),
        ('TEXTCOLOR', (2,0), (2,-1), navy),
        ('BACKGROUND',(0,0), (-1,-1), lgray),
        ('ROWBACKGROUNDS', (0,0), (-1,-1), [lgray, white]),
        ('GRID', (0,0), (-1,-1), 0.5, mgray),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(det_tbl)
    story.append(Spacer(1, 0.5*cm))

    # ── Score summary ──
    total_yes = sum(1 for v in responses.values() if v.get('answer') == 'yes')
    total_no  = sum(1 for v in responses.values() if v.get('answer') == 'no')
    total_ans = total_yes + total_no
    pct = round(100 * total_yes / total_ans) if total_ans else 0

    score_data = [["Total Items", "Compliant (Yes)", "Non-Compliant (No)", "Compliance %"],
                  [str(total_ans), str(total_yes), str(total_no), f"{pct}%"]]
    score_tbl = Table(score_data, colWidths=[W/4]*4)
    score_tbl.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), steel),
        ('TEXTCOLOR',  (0,0), (-1,0), white),
        ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1,-1), 9),
        ('ALIGN',      (0,0), (-1,-1), 'CENTER'),
        ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,1), (-1,1), lgray),
        ('GRID',       (0,0), (-1,-1), 0.5, mgray),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('FONTNAME',   (0,1), (-1,1), 'Helvetica-Bold'),
        ('TEXTCOLOR',  (1,1), (1,1), green),
        ('TEXTCOLOR',  (2,1), (2,1), red),
        ('FONTSIZE',   (0,1), (-1,1), 12),
    ]))
    story.append(score_tbl)
    story.append(Spacer(1, 0.6*cm))

    # ── Per section checklist ──
    for section in SECTIONS:
        story.append(Paragraph(f"{section['icon']}  {section['title'].upper()}", head2_style))
        story.append(HRFlowable(width=W, thickness=1.5, color=saffron, spaceAfter=6))

        for sub in section['subsections']:
            story.append(Paragraph(sub['title'], head3_style))

            tbl_data = [["#", "Inspection Item", "Ans", "Remarks", "Person(s) Responsible"]]
            for i, item in enumerate(sub['items'], 1):
                resp    = responses.get(item['id'], {})
                ans     = resp.get('answer', '---').upper()
                remark  = resp.get('remark', '')
                persons = resp.get('person_responsible', '') if ans == 'NO' else ''
                ans_color = green if ans == 'YES' else (red if ans == 'NO' else black)

                tbl_data.append([
                    Paragraph(str(i), small_style),
                    Paragraph(item['label'], body_style),
                    Paragraph(f"<b>{ans}</b>", ParagraphStyle('ans', fontSize=8, fontName='Helvetica-Bold', textColor=ans_color, alignment=TA_CENTER)),
                    Paragraph(remark or '', remark_style),
                    Paragraph(persons or '', ParagraphStyle('pers', fontSize=7.5, fontName='Helvetica-Bold', textColor=colors.HexColor('#1a3060'), leading=11)),
                ])

            col_w = [0.8*cm, 7.8*cm, 1.6*cm, 3.8*cm, 3.4*cm]
            tbl = Table(tbl_data, colWidths=col_w, repeatRows=1)
            tbl.setStyle(TableStyle([
                ('BACKGROUND',    (0,0), (-1,0), navy),
                ('TEXTCOLOR',     (0,0), (-1,0), white),
                ('FONTNAME',      (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE',      (0,0), (-1,0), 8),
                ('ALIGN',         (0,0), (0,-1), 'CENTER'),
                ('ALIGN',         (2,0), (2,-1), 'CENTER'),
                ('VALIGN',        (0,0), (-1,-1), 'TOP'),
                ('ROWBACKGROUNDS',(0,1), (-1,-1), [white, lgray]),
                ('GRID',          (0,0), (-1,-1), 0.4, mgray),
                ('TOPPADDING',    (0,0), (-1,-1), 4),
                ('BOTTOMPADDING', (0,0), (-1,-1), 4),
                ('LEFTPADDING',   (0,0), (-1,-1), 5),
                ('LINEBELOW',     (0,-1), (-1,-1), 1, steel),
            ]))
            story.append(tbl)
            story.append(Spacer(1, 0.3*cm))

    # ── Signature block ──
    story.append(Spacer(1, 0.8*cm))
    sig_data = [
        ["Inspector's Signature", "Date & Time of Inspection"],
        ["\n\n" + "_"*35, datetime.now().strftime("%d-%m-%Y  %H:%M")],
    ]
    sig_tbl = Table(sig_data, colWidths=[W/2, W/2])
    sig_tbl.setStyle(TableStyle([
        ('FONTNAME',  (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',  (0,0), (-1,-1), 8.5),
        ('ALIGN',     (0,0), (-1,-1), 'CENTER'),
        ('GRID',      (0,0), (-1,-1), 0.5, mgray),
        ('BACKGROUND',(0,0), (-1,0), lgray),
        ('TOPPADDING',(0,0), (-1,-1), 5),
        ('BOTTOMPADDING',(0,0), (-1,-1), 12),
    ]))
    story.append(sig_tbl)

    doc.build(story)
    buf.seek(0)
    return buf


# ── DOCX generation ───────────────────────────────────────────────────────────
def generate_docx(header, responses):
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_ALIGN_VERTICAL
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    def set_cell_bg(cell, hex_color):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), hex_color)
        tcPr.append(shd)

    doc = Document()
    # Page margins
    for section in doc.sections:
        section.left_margin   = Cm(2)
        section.right_margin  = Cm(2)
        section.top_margin    = Cm(2)
        section.bottom_margin = Cm(2)

    # Title
    title = doc.add_heading('', 0)
    run = title.add_run('MECHANICAL DEPARTMENT INSPECTION REPORT')
    run.font.color.rgb = RGBColor(0x0a, 0x16, 0x28)
    run.font.size = Pt(16)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = doc.add_paragraph('Office of Sr. DME (Co.) BCT · Western Railway · Mumbai Central')
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.runs[0].font.color.rgb = RGBColor(0x3a, 0x7b, 0xd5)
    sub.runs[0].font.size = Pt(10)

    doc.add_paragraph()

    # Header table
    tbl = doc.add_table(rows=3, cols=4)
    tbl.style = 'Table Grid'
    fields = [
        ("Inspector Name:", header.get('name',''),       "Designation:",    header.get('designation','')),
        ("Location:",       header.get('location',''),   "Date:",           header.get('date','')),
        ("Inspection Type:",header.get('insp_type',''),  "Division:",        "BCT / Western Railway"),
    ]
    for r, (l1, v1, l2, v2) in enumerate(fields):
        row = tbl.rows[r]
        for ci, txt in enumerate([l1, v1, l2, v2]):
            cell = row.cells[ci]
            cell.text = txt
            run = cell.paragraphs[0].runs[0]
            run.font.size = Pt(9)
            if ci in (0, 2):
                run.bold = True
                run.font.color.rgb = RGBColor(0x0a, 0x16, 0x28)

    doc.add_paragraph()

    # Score
    total_yes = sum(1 for v in responses.values() if v.get('answer') == 'yes')
    total_no  = sum(1 for v in responses.values() if v.get('answer') == 'no')
    total_ans = total_yes + total_no
    pct = round(100 * total_yes / total_ans) if total_ans else 0

    score_p = doc.add_paragraph()
    score_p.add_run(f'COMPLIANCE SUMMARY: ').bold = True
    score_p.add_run(f'Total Answered: {total_ans}  |  Compliant (Yes): {total_yes}  |  Non-Compliant (No): {total_no}  |  Compliance: {pct}%')
    score_p.runs[0].bold = True

    doc.add_paragraph()

    # Sections
    for section in SECTIONS:
        doc.add_heading(f"{section['icon']} {section['title']}", level=1)
        for sub in section['subsections']:
            doc.add_heading(sub['title'], level=2)
            tbl2 = doc.add_table(rows=1, cols=4)
            tbl2.style = 'Table Grid'
            hdr_cells = tbl2.rows[0].cells
            for ci, hd in enumerate(['#', 'Inspection Item', 'Answer', 'Remarks']):
                hdr_cells[ci].text = hd
                hdr_cells[ci].paragraphs[0].runs[0].bold = True
                hdr_cells[ci].paragraphs[0].runs[0].font.size = Pt(8)
                set_cell_bg(hdr_cells[ci], '0a1628')
                hdr_cells[ci].paragraphs[0].runs[0].font.color.rgb = RGBColor(0xff, 0xff, 0xff)

            for i, item in enumerate(sub['items'], 1):
                resp   = responses.get(item['id'], {})
                ans    = resp.get('answer', '—').upper()
                remark = resp.get('remark', '')
                row = tbl2.add_row().cells
                row[0].text = str(i)
                row[1].text = item['label']
                row[2].text = ans
                row[3].text = remark
                for ci in range(4):
                    row[ci].paragraphs[0].runs[0].font.size = Pt(8)
                run2 = row[2].paragraphs[0].runs[0]
                run2.bold = True
                if ans == 'YES':
                    run2.font.color.rgb = RGBColor(0x16, 0xa3, 0x4a)
                elif ans == 'NO':
                    run2.font.color.rgb = RGBColor(0xdc, 0x26, 0x26)
                if i % 2 == 0:
                    for ci in range(4):
                        set_cell_bg(row[ci], 'f5f5f5')

            # Column widths
            for ri, row in enumerate(tbl2.rows):
                for ci, width in enumerate([0.8, 8.5, 1.8, 4.4]):
                    row.cells[ci].width = Cm(width)
            doc.add_paragraph()

    # Signature
    doc.add_paragraph()
    sig_tbl = doc.add_table(rows=2, cols=2)
    sig_tbl.style = 'Table Grid'
    for ci, hd in enumerate(["Inspector's Signature", "Date & Time of Inspection"]):
        sig_tbl.rows[0].cells[ci].text = hd
        sig_tbl.rows[0].cells[ci].paragraphs[0].runs[0].bold = True
        sig_tbl.rows[0].cells[ci].paragraphs[0].runs[0].font.size = Pt(8)
    sig_tbl.rows[1].cells[1].text = datetime.now().strftime("%d-%m-%Y  %H:%M")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ── WhatsApp summary ──────────────────────────────────────────────────────────
def generate_whatsapp(header, responses):
    total_yes = sum(1 for v in responses.values() if v.get('answer') == 'yes')
    total_no  = sum(1 for v in responses.values() if v.get('answer') == 'no')
    total_ans = total_yes + total_no
    pct = round(100 * total_yes / total_ans) if total_ans else 0

    lines = [
        "🚂 *MECHANICAL DEPT INSPECTION REPORT*",
        f"📍 *Location:* {header.get('location','')}",
        f"👤 *Inspector:* {header.get('name','')} ({header.get('designation','')})",
        f"📅 *Date:* {header.get('date','')}",
        f"🔍 *Type:* {header.get('insp_type','')}",
        "",
        "📊 *COMPLIANCE SUMMARY*",
        f"✅ Compliant (Yes): {total_yes}",
        f"❌ Non-Compliant (No): {total_no}",
        f"📈 Compliance: *{pct}%*",
        "",
        "⚠️ *NON-COMPLIANT ITEMS:*",
    ]
    nc_count = 0
    for section in SECTIONS:
        for sub in section['subsections']:
            for item in sub['items']:
                resp = responses.get(item['id'], {})
                if resp.get('answer') == 'no':
                    nc_count += 1
                    remark = resp.get('remark', '')
                    remark_str = f" | Remark: {remark}" if remark else ""
                    persons_str = f" | Responsible: {resp.get('person_responsible','')}" if resp.get('person_responsible') else ""
                    lines.append(f"• {item['label'][:70]}{remark_str}{persons_str}")
    if nc_count == 0:
        lines.append("None — Full Compliance ✅")

    lines += ["", f"_Generated by Mech Inspection App · BCT Division · {datetime.now().strftime('%d-%m-%Y %H:%M')}_"]
    return "\n".join(lines)


# ── Routes ────────────────────────────────────────────────────────────────────
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/checklist')
def checklist():
    import json
    sections_json = json.dumps(SECTIONS)
    return render_template('checklist.html', sections_json=sections_json)

@app.route('/submit', methods=['POST'])
def submit():
    data = request.get_json()
    session['header']    = data.get('header', {})
    session['responses'] = data.get('responses', {})
    return jsonify({'status': 'ok'})

@app.route('/summary')
def summary():
    header    = session.get('header', {})
    responses = session.get('responses', {})
    total_yes = sum(1 for v in responses.values() if v.get('answer') == 'yes')
    total_no  = sum(1 for v in responses.values() if v.get('answer') == 'no')
    total_ans = total_yes + total_no
    pct = round(100 * total_yes / total_ans) if total_ans else 0
    nc_items = []
    for section in SECTIONS:
        for sub in section['subsections']:
            for item in sub['items']:
                resp = responses.get(item['id'], {})
                if resp.get('answer') == 'no':
                    nc_items.append({
                        'label': item['label'],
                        'remark': resp.get('remark',''),
                        'section': section['title'],
                        'persons': resp.get('person_responsible','')
                    })
    whatsapp = generate_whatsapp(header, responses)
    return render_template('summary.html',
                           header=header,
                           total_yes=total_yes, total_no=total_no, total_ans=total_ans, pct=pct,
                           nc_items=nc_items, whatsapp=whatsapp)

@app.route('/download/pdf')
def download_pdf():
    header    = session.get('header', {})
    responses = session.get('responses', {})
    buf = generate_pdf(header, responses)
    fname = f"Inspection_{header.get('location','BCT').replace(' ','_')}_{header.get('date','').replace('-','')}.pdf"
    return send_file(buf, as_attachment=True, download_name=fname, mimetype='application/pdf')

@app.route('/download/docx')
def download_docx():
    header    = session.get('header', {})
    responses = session.get('responses', {})
    buf = generate_docx(header, responses)
    fname = f"Inspection_{header.get('location','BCT').replace(' ','_')}_{header.get('date','').replace('-','')}.docx"
    return send_file(buf, as_attachment=True, download_name=fname, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
