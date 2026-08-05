# ============================================================
# 샘플 PDF 파일 생성 스크립트
# ============================================================

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import os

# sample_pdfs 폴더가 없으면 생성
if not os.path.exists('sample_pdfs'):
    os.makedirs('sample_pdfs')
    print("sample_pdfs 폴더 생성됨")

# 샘플 PDF 데이터
pdf_data = [
    {
        'filename': 'R001_김민영_최종논문.pdf',
        'title': 'Deep Learning in Humanities',
        'author': 'Kim Min-young',
        'content': 'This is a sample paper by Kim Min-young about deep learning applications in humanities research. The author would like to acknowledge the National Research Foundation for their support.'
    },
    {
        'filename': 'R003_홍지수_논문.pdf',
        'title': 'Philosophy of Artificial Intelligence',
        'author': 'Hong Ji-su',
        'content': 'A philosophical inquiry into the nature of artificial intelligence and its implications for human society. We thank the Ministry of Education for providing research support.'
    },
    {
        'filename': 'R005_박소연_저서.pdf',
        'title': 'Aesthetics in Digital Age',
        'author': 'Park So-yeon',
        'content': 'Exploring aesthetic principles in the digital age and their application to contemporary art forms. This research was supported by the Korea Arts Council.'
    }
]

# PDF 생성
for pdf_info in pdf_data:
    pdf_path = os.path.join('sample_pdfs', pdf_info['filename'])
    
    try:
        # Canvas 생성
        c = canvas.Canvas(pdf_path, pagesize=letter)
        width, height = letter
        
        # 제목
        c.setFont("Helvetica-Bold", 16)
        c.drawString(50, height - 50, pdf_info['title'])
        
        # 저자
        c.setFont("Helvetica", 12)
        c.drawString(50, height - 80, f"Author: {pdf_info['author']}")
        
        # 내용
        c.setFont("Helvetica", 11)
        y_position = height - 120
        
        words = pdf_info['content'].split(' ')
        for word in words:
            if y_position < 50:
                c.showPage()
                c.setFont("Helvetica", 11)
                y_position = height - 50
            c.drawString(50, y_position, word)
            y_position -= 15
        
        # PDF 저장
        c.save()
        print(f"✅ 생성됨: {pdf_path}")
    
    except Exception as e:
        print(f"❌ 오류 발생 ({pdf_info['filename']}): {str(e)}")

print("\n✅ 모든 샘플 PDF가 생성되었습니다!")