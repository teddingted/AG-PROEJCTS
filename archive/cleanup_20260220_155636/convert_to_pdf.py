# -*- coding: utf-8 -*-
"""
Markdown to PDF Converter with Korean Font Support
한글 폰트를 지원하는 PDF 변환기
"""
import sys
import os

# UTF-8 Fix
if sys.platform == 'win32':
    if sys.stdout.encoding != 'utf-8':
        sys.stdout.reconfigure(encoding='utf-8')
    if sys.stderr.encoding != 'utf-8':
        sys.stderr.reconfigure(encoding='utf-8')
    os.environ['PYTHONIOENCODING'] = 'utf-8'

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY

def find_korean_font():
    """Find Korean font on Windows system"""
    font_paths = [
        r'C:\Windows\Fonts\malgun.ttf',  # 맑은 고딕
        r'C:\Windows\Fonts\malgunbd.ttf',  # 맑은 고딕 Bold
        r'C:\Windows\Fonts\gulim.ttc',  # 굴림
        r'C:\Windows\Fonts\batang.ttc',  # 바탕
        r'C:\Windows\Fonts\NanumGothic.ttf',  # 나눔고딕 (if installed)
    ]
    
    for path in font_paths:
        if os.path.exists(path):
            print(f"Found Korean font: {path}")
            return path
    
    print("Warning: No Korean font found, using default")
    return None

def register_korean_font():
    """Register Korean font with reportlab"""
    font_path = find_korean_font()
    
    if font_path and font_path.endswith('.ttf'):
        try:
            pdfmetrics.registerFont(TTFont('KoreanFont', font_path))
            print(f"✓ Registered Korean font: {font_path}")
            return 'KoreanFont'
        except Exception as e:
            print(f"Error registering font: {e}")
            return 'Helvetica'
    else:
        print("Using default font (Korean may not display correctly)")
        return 'Helvetica'

def create_korean_pdf(md_path, pdf_path):
    """Create PDF with Korean font support"""
    print(f"\n{'='*70}")
    print("PDF 생성 (한글 폰트 지원)")
    print(f"{'='*70}\n")
    
    # Register Korean font
    korean_font = register_korean_font()
    
    # Read markdown
    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Create PDF with margins
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        leftMargin=2*cm,
        rightMargin=2*cm,
        topMargin=2.5*cm,
        bottomMargin=2*cm
    )
    
    # Custom styles with Korean font
    styles = getSampleStyleSheet()
    
    # Title style
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Title'],
        fontName=korean_font,
        fontSize=18,
        leading=24,
        alignment=TA_CENTER,
        spaceAfter=12,
        textColor=colors.HexColor('#1a1a1a')
    )
    
    # Heading styles
    heading1_style = ParagraphStyle(
        'CustomHeading1',
        parent=styles['Heading1'],
        fontName=korean_font,
        fontSize=14,
        leading=18,
        spaceAfter=10,
        textColor=colors.HexColor('#2c3e50')
    )
    
    heading2_style = ParagraphStyle(
        'CustomHeading2',
        parent=styles['Heading2'],
        fontName=korean_font,
        fontSize=12,
        leading=16,
        spaceAfter=8,
        textColor=colors.HexColor('#34495e')
    )
    
    # Body text style
    body_style = ParagraphStyle(
        'CustomBody',
        parent=styles['BodyText'],
        fontName=korean_font,
        fontSize=9,
        leading=12,
        alignment=TA_LEFT,
        spaceAfter=6
    )
    
    # Build story
    story = []
    lines = content.split('\n')
    
    print(f"Processing {len(lines)} lines...")
    
    for i, line in enumerate(lines):
        try:
            line = line.strip()
            
            if not line:
                story.append(Spacer(1, 0.1*inch))
                continue
            
            # Skip image references (can't embed easily)
            if line.startswith('!['):
                continue
            
            # Skip horizontal rules
            if line.startswith('---') or line.startswith('***'):
                story.append(Spacer(1, 0.2*inch))
                continue
            
            # Process headings
            if line.startswith('# '):
                text = line[2:].strip()
                if text:
                    story.append(Paragraph(text, title_style))
                    story.append(Spacer(1, 0.15*inch))
                    
            elif line.startswith('## '):
                text = line[3:].strip()
                if text:
                    story.append(Spacer(1, 0.2*inch))
                    story.append(Paragraph(text, heading1_style))
                    
            elif line.startswith('### '):
                text = line[4:].strip()
                if text:
                    story.append(Paragraph(text, heading2_style))
                    
            elif line.startswith('| '):
                # Skip table lines for now (complex to parse)
                continue
                
            elif line.startswith('**') and line.endswith('**'):
                # Bold text
                text = line.strip('*')
                if text:
                    bold_style = ParagraphStyle(
                        'Bold',
                        parent=body_style,
                        fontName=korean_font,
                        fontSize=10,
                        leading=14
                    )
                    story.append(Paragraph(f"<b>{text}</b>", bold_style))
                    
            else:
                # Normal text
                # Remove markdown formatting
                text = line.replace('**', '').replace('*', '').replace('`', '')
                
                # Limit line length
                if len(text) > 150:
                    text = text[:150]
                
                if text:
                    try:
                        story.append(Paragraph(text, body_style))
                    except:
                        # Skip problematic lines
                        pass
            
            # Progress indicator
            if (i + 1) % 100 == 0:
                print(f"  Processed {i+1}/{len(lines)} lines...")
                
        except Exception as e:
            # Skip problematic lines
            continue
    
    # Build PDF
    print("\nBuilding PDF document...")
    try:
        doc.build(story)
        print(f"\n✓ PDF 생성 성공: {pdf_path}")
        print(f"  파일 크기: {os.path.getsize(pdf_path) / 1024:.1f} KB")
        print(f"  사용 폰트: {korean_font}")
        print(f"  총 페이지: 추정 {len(story) // 50 + 1}페이지")
        
    except Exception as e:
        print(f"\n✗ PDF 생성 실패: {e}")
        raise

if __name__ == "__main__":
    base_dir = r"c:\Users\Admin\Desktop\AG-BEGINNING"
    
    md_file = os.path.join(base_dir, "시운전_시험성적서_8196_FULL.md")
    pdf_file = os.path.join(base_dir, "시운전_시험성적서_8196_FINAL.pdf")
    
    if os.path.exists(md_file):
        create_korean_pdf(md_file, pdf_file)
        print(f"\n{'='*70}")
        print("완료! PDF 파일을 확인해주세요.")
        print(f"{'='*70}")
    else:
        print(f"Error: {md_file} not found")
