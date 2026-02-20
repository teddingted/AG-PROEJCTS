# -*- coding: utf-8 -*-
"""
Professional Commissioning Report PDF Generator
전문 시운전 보고서 PDF 생성기
"""
import sys
import os
import re

# UTF-8 Fix
if sys.platform == 'win32':
    if sys.stdout.encoding != 'utf-8':
        sys.stdout.reconfigure(encoding='utf-8')
    if sys.stderr.encoding != 'utf-8':
        sys.stderr.reconfigure(encoding='utf-8')
    os.environ['PYTHONIOENCODING'] = 'utf-8'

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY

def setup_korean_font():
    """Setup Korean font"""
    font_path = r'C:\Windows\Fonts\malgun.ttf'
    if os.path.exists(font_path):
        pdfmetrics.registerFont(TTFont('KoreanFont', font_path))
        return 'KoreanFont'
    return 'Helvetica'

def parse_table(lines, start_idx):
    """Parse markdown table"""
    table_lines = []
    i = start_idx
    
    while i < len(lines) and lines[i].strip().startswith('|'):
        table_lines.append(lines[i])
        i += 1
    
    if len(table_lines) < 2:
        return None, start_idx
    
    table_data = []
    for line in table_lines:
        if re.match(r'^\|[\s:|-]+\|$', line):
            continue
        cells = [cell.strip() for cell in line.split('|')]
        cells = [c for c in cells if c]
        if cells:
            table_data.append(cells)
    
    return table_data, i

def create_table(table_data, font):
    """Create styled table"""
    if not table_data:
        return None
    
    table = Table(table_data)
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#34495e')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), font),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('TOPPADDING', (0, 0), (-1, 0), 8),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#ecf0f1')),
        ('FONTNAME', (0, 1), (-1, -1), font),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ])
    table.setStyle(style)
    return table

def generate_professional_pdf(md_path, pdf_path):
    """Generate professional PDF from markdown"""
    
    print(f"\n{'='*80}")
    print(f"{'전문 시운전 보고서 PDF 생성':^70}")
    print(f"{'='*80}\n")
    
    font = setup_korean_font()
    print(f"✓ 한글 폰트: {font}\n")
    
    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    lines = content.split('\n')
    
    # PDF setup
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        leftMargin=2*cm,
        rightMargin=2*cm,
        topMargin=2.5*cm,
        bottomMargin=2*cm
    )
    
    # Professional styles
    title = ParagraphStyle(
        'Title',
        fontName=font,
        fontSize=18,
        leading=24,
        alignment=TA_CENTER,
        spaceAfter=12,
        textColor=colors.HexColor('#2c3e50'),
        fontWeight='BOLD'
    )
    
    subtitle = ParagraphStyle(
        'Subtitle',
        fontName=font,
        fontSize=14,
        leading=18,
        alignment=TA_CENTER,
        spaceAfter=10,
        textColor=colors.HexColor('#7f8c8d')
    )
    
    h1 = ParagraphStyle(
        'H1',
        fontName=font,
        fontSize=14,
        leading=18,
        spaceAfter=10,
        spaceBefore=15,
        textColor=colors.HexColor('#2c3e50'),
        borderPadding=5
    )
    
    h2 = ParagraphStyle(
        'H2',
        fontName=font,
        fontSize=12,
        leading=16,
        spaceAfter=8,
        spaceBefore=10,
        textColor=colors.HexColor('#34495e')
    )
    
    h3 = ParagraphStyle(
        'H3',
        fontName=font,
        fontSize=10,
        leading=14,
        spaceAfter=6,
        textColor=colors.HexColor('#7f8c8d')
    )
    
    body = ParagraphStyle(
        'Body',
        fontName=font,
        fontSize=9,
        leading=12,
        alignment=TA_JUSTIFY,
        spaceAfter=5
    )
    
    # Build document
    story = []
    i = 0
    stats = {'tables': 0, 'images': 0, 'pages_est': 0}
    
    print(f"문서 처리 중 ({len(lines)}줄)...\n")
    
    while i < len(lines):
        line = lines[i].strip()
        
        if not line:
            story.append(Spacer(1, 0.08*inch))
            i += 1
            continue
        
        if line.startswith('---'):
            story.append(Spacer(1, 0.15*inch))
            i += 1
            continue
        
        # Tables
        if line.startswith('|'):
            table_data, end_idx = parse_table(lines, i)
            if table_data:
                tbl = create_table(table_data, font)
                if tbl:
                    story.append(tbl)
                    story.append(Spacer(1, 0.2*inch))
                    stats['tables'] += 1
                i = end_idx
                continue
        
        # Images
        if line.startswith('!['):
            match = re.match(r'!\[(.*?)\]\((.*?)\)', line)
            if match:
                caption = match.group(1)
                img_path = match.group(2)
                if os.path.exists(img_path):
                    try:
                        img = Image(img_path, width=5.5*inch, height=3.3*inch)
                        story.append(img)
                        cap = ParagraphStyle('Cap', fontName=font, fontSize=8, alignment=TA_CENTER, textColor=colors.grey)
                        story.append(Paragraph(f"<i>그림: {caption}</i>", cap))
                        story.append(Spacer(1, 0.15*inch))
                        stats['images'] += 1
                    except:
                        pass
            i += 1
            continue
        
        # Headings
        if line.startswith('# '):
            story.append(Paragraph(line[2:], title))
            i += 1
            continue
        
        if line.startswith('## '):
            story.append(Paragraph(line[3:], h1))
            i += 1
            continue
        
        if line.startswith('### '):
            story.append(Paragraph(line[4:], h2))
            i += 1
            continue
        
        if line.startswith('#### '):
            story.append(Paragraph(line[5:], h3))
            i += 1
            continue
        
        # Text
        text = line.replace('**', '').replace('*', '').replace('`', '')
        text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        if text and len(text) > 1 and not text.startswith('['):
            try:
                story.append(Paragraph(text[:250], body))
            except:
                pass
        
        i += 1
        
        if i % 200 == 0:
            print(f"  진행: {i}/{len(lines)} 줄...")
    
    # Generate PDF
    print(f"\nPDF 문서 생성 중...")
    doc.build(story)
    
    stats['size_kb'] = os.path.getsize(pdf_path) / 1024
    stats['pages_est'] = len(story) // 35 + 1
    
    print(f"\n{'='*80}")
    print(f"✓ PDF 생성 완료\n")
    print(f"  파일명: {os.path.basename(pdf_path)}")
    print(f"  크기: {stats['size_kb']:.1f} KB")
    print(f"  폰트: {font} (맑은 고딕)")
    print(f"  테이블: {stats['tables']}개")
    print(f"  이미지: {stats['images']}개")
    print(f"  예상 페이지: {stats['pages_est']}페이지\n")
    print(f"{'='*80}\n")

if __name__ == "__main__":
    md_file = r"c:\Users\Admin\Desktop\AG-BEGINNING\시운전_시험성적서_8196_ULTIMATE.md"
    pdf_file = r"c:\Users\Admin\Desktop\AG-BEGINNING\시운전_시험성적서_8196_ULTIMATE.pdf"
    
    if os.path.exists(md_file):
        generate_professional_pdf(md_file, pdf_file)
        print("✓ 전문 시운전 보고서 PDF 생성이 완료되었습니다.")
        print(f"  위치: {pdf_file}\n")
    else:
        print(f"Error: {md_file} not found")
