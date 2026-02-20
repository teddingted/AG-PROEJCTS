# -*- coding: utf-8 -*-
"""
Enhanced Markdown to PDF Converter with Table Support
테이블과 이미지를 지원하는 고급 PDF 변환기
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
from reportlab.lib.enums import TA_CENTER, TA_LEFT

def find_korean_font():
    """Find Korean font on Windows system"""
    font_paths = [
        r'C:\Windows\Fonts\malgun.ttf',
        r'C:\Windows\Fonts\gulim.ttc',
    ]
    
    for path in font_paths:
        if os.path.exists(path):
            return path
    return None

def register_korean_font():
    """Register Korean font"""
    font_path = find_korean_font()
    if font_path and font_path.endswith('.ttf'):
        pdfmetrics.registerFont(TTFont('KoreanFont', font_path))
        return 'KoreanFont'
    return 'Helvetica'

def parse_markdown_table(lines, start_idx):
    """
    Parse markdown table starting from given index
    Returns: (table_data, end_idx)
    """
    table_lines = []
    i = start_idx
    
    while i < len(lines) and lines[i].strip().startswith('|'):
        table_lines.append(lines[i])
        i += 1
    
    if len(table_lines) < 2:
        return None, start_idx
    
    # Parse table
    table_data = []
    for line in table_lines:
        # Skip separator line (|:---|:---|)
        if re.match(r'^\|[\s:|-]+\|$', line):
            continue
        
        # Split by |
        cells = [cell.strip() for cell in line.split('|')]
        # Remove empty first/last cells
        cells = [c for c in cells if c]
        
        if cells:
            table_data.append(cells)
    
    return table_data, i

def create_reportlab_table(table_data, korean_font):
    """Create reportlab Table from parsed data"""
    if not table_data:
        return None
    
    # Create table
    table = Table(table_data)
    
    # Style
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2c3e50')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), korean_font),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('FONTNAME', (0, 1), (-1, -1), korean_font),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ])
    
    table.setStyle(style)
    return table

def create_enhanced_pdf(md_path, pdf_path):
    """Create PDF with table and image support"""
    print(f"\n{'='*70}")
    print("고급 PDF 생성 (테이블 + 이미지 지원)")
    print(f"{'='*70}\n")
    
    korean_font = register_korean_font()
    print(f"✓ 한글 폰트 등록: {korean_font}\n")
    
    # Read markdown
    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    lines = content.split('\n')
    
    # Create PDF
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        leftMargin=1.5*cm,
        rightMargin=1.5*cm,
        topMargin=2*cm,
        bottomMargin=2*cm
    )
    
    # Styles
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'Title',
        fontName=korean_font,
        fontSize=16,
        leading=20,
        alignment=TA_CENTER,
        spaceAfter=10,
        textColor=colors.HexColor('#1a1a1a')
    )
    
    h1_style = ParagraphStyle(
        'H1',
        fontName=korean_font,
        fontSize=13,
        leading=16,
        spaceAfter=8,
        textColor=colors.HexColor('#2c3e50'),
        spaceBefore=10
    )
    
    h2_style = ParagraphStyle(
        'H2',
        fontName=korean_font,
        fontSize=11,
        leading=14,
        spaceAfter=6,
        textColor=colors.HexColor('#34495e')
    )
    
    h3_style = ParagraphStyle(
        'H3',
        fontName=korean_font,
        fontSize=10,
        leading=13,
        spaceAfter=5,
        textColor=colors.HexColor('#7f8c8d')
    )
    
    body_style = ParagraphStyle(
        'Body',
        fontName=korean_font,
        fontSize=8,
        leading=11,
        alignment=TA_LEFT,
        spaceAfter=4
    )
    
    # Build story
    story = []
    i = 0
    table_count = 0
    
    print(f"총 {len(lines)}줄 처리 중...\n")
    
    while i < len(lines):
        line = lines[i].strip()
        
        # Skip empty lines
        if not line:
            story.append(Spacer(1, 0.05*inch))
            i += 1
            continue
        
        # Skip horizontal rules
        if line.startswith('---') or line.startswith('***'):
            story.append(Spacer(1, 0.1*inch))
            i += 1
            continue
        
        # Parse tables
        if line.startswith('|'):
            table_data, end_idx = parse_markdown_table(lines, i)
            if table_data:
                table = create_reportlab_table(table_data, korean_font)
                if table:
                    story.append(table)
                    story.append(Spacer(1, 0.15*inch))
                    table_count += 1
                    print(f"  ✓ 테이블 {table_count}: {len(table_data)}행 × {len(table_data[0])}열")
                i = end_idx
                continue
        
        # Parse images
        if line.startswith('!['):
            # Extract image path
            match = re.match(r'!\[(.*?)\]\((.*?)\)', line)
            if match:
                caption = match.group(1)
                img_path = match.group(2)
                
                # Try to add image
                if os.path.exists(img_path):
                    try:
                        img = Image(img_path, width=5*inch, height=3*inch)
                        story.append(img)
                        if caption:
                            cap_style = ParagraphStyle(
                                'Caption',
                                fontName=korean_font,
                                fontSize=8,
                                alignment=TA_CENTER,
                                textColor=colors.grey
                            )
                            story.append(Paragraph(f"그림: {caption}", cap_style))
                        story.append(Spacer(1, 0.1*inch))
                        print(f"  ✓ 이미지 추가: {caption}")
                    except:
                        # Image load failed, add caption only
                        story.append(Paragraph(f"[이미지: {caption}]", body_style))
                else:
                    # Image not found, add reference
                    story.append(Paragraph(f"[이미지: {caption}]", body_style))
            i += 1
            continue
        
        # Parse headings
        if line.startswith('# '):
            story.append(Paragraph(line[2:], title_style))
            i += 1
            continue
        
        if line.startswith('## '):
            story.append(Paragraph(line[3:], h1_style))
            i += 1
            continue
        
        if line.startswith('### '):
            story.append(Paragraph(line[4:], h2_style))
            i += 1
            continue
        
        if line.startswith('#### '):
            story.append(Paragraph(line[5:], h3_style))
            i += 1
            continue
        
        # Normal text
        # Clean markdown
        text = line.replace('**', '').replace('*', '').replace('`', '')
        
        # Handle special characters
        text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        if text and len(text) > 1:
            try:
                story.append(Paragraph(text[:200], body_style))
            except:
                pass
        
        i += 1
        
        # Progress
        if i % 100 == 0:
            print(f"  처리 중: {i}/{len(lines)} 줄...")
    
    # Build PDF
    print(f"\nPDF 문서 생성 중...")
    doc.build(story)
    
    file_size = os.path.getsize(pdf_path) / 1024
    print(f"\n{'='*70}")
    print(f"✓ PDF 생성 완료!")
    print(f"{'='*70}")
    print(f"  파일: {pdf_path}")
    print(f"  크기: {file_size:.1f} KB")
    print(f"  폰트: {korean_font}")
    print(f"  테이블: {table_count}개")
    print(f"  페이지: 약 {len(story) // 40 + 1}페이지\n")

if __name__ == "__main__":
    md_file = r"c:\Users\Admin\Desktop\AG-BEGINNING\시운전_시험성적서_8196_FULL.md"
    pdf_file = r"c:\Users\Admin\Desktop\AG-BEGINNING\시운전_시험성적서_8196_FINAL.pdf"
    
    if os.path.exists(md_file):
        create_enhanced_pdf(md_file, pdf_file)
    else:
        print(f"Error: {md_file} not found")
