# ============================================================
# AI 기반 연구지원사업 성과관리 시스템 - Phase 8 (수정)
# 비용 기능 제거 버전
# ============================================================

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
import re
import fitz
import requests
import time
from functools import lru_cache
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

# ============================================================
# 페이지 설정
# ============================================================

st.set_page_config(
    page_title="연구성과관리시스템",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📊 연구지원사업 성과관리시스템")
st.markdown("**Phase 8: 최종 완성 버전** | 연구성과 종합 관리 및 분석")

# ============================================================
# Excel 생성 함수 (수정: 효율성 시트 제거)
# ============================================================

def create_comprehensive_excel(df_matching, df_review=None):
    """
    여러 시트가 포함된 Excel 파일을 생성하는 함수
    (효율성 분석 시트 제거)
    """
    
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Sheet 1: 전체 현황
        df_matching.to_excel(writer, sheet_name='전체 현황', index=False)
        
        # Sheet 2: 제출 현황
        submitted = df_matching[df_matching['submission_status'] == '제출']
        submitted.to_excel(writer, sheet_name='제출 현황', index=False)
        
        # Sheet 3: 미제출자
        not_submitted = df_matching[df_matching['submission_status'] == '미제출']
        not_submitted.to_excel(writer, sheet_name='미제출자', index=False)
        
        # Sheet 4: 기한 초과
        delayed = df_matching[df_matching['timeliness'] == '지연 제출']
        if len(delayed) > 0:
            delayed.to_excel(writer, sheet_name='기한 초과', index=False)
        
        # Sheet 5: 검수 결과
        if df_review is not None:
            df_review.to_excel(writer, sheet_name='검수 결과', index=False)
        
        # Sheet 6: 통계 요약
        summary_data = {
            '항목': [
                '전체 연구자 수',
                '제출 완료',
                '미제출',
                '기한 내 제출',
                '지연 제출',
                '제출율 (%)',
                '기한 준수율 (%)'
            ],
            '값': [
                len(df_matching),
                len(submitted),
                len(not_submitted),
                len(df_matching[df_matching['timeliness'] == '기한 내 제출']),
                len(delayed),
                f"{(len(submitted)/len(df_matching)*100):.1f}" if len(df_matching) > 0 else 0,
                f"{(len(df_matching[df_matching['timeliness'] == '기한 내 제출'])/len(df_matching)*100):.1f}" 
                if len(df_matching) > 0 else 0
            ]
        }
        df_summary = pd.DataFrame(summary_data)
        df_summary.to_excel(writer, sheet_name='통계 요약', index=False)
    
    output.seek(0)
    return output


# 이전 단계에서 정의된 모든 함수들 (Phase 1~7)
# ============================================================

def perform_research_review(matching_data, extracted_pdf_data=None, crossref_data=None):
    """연구성과 자동 검수"""
    review_result = {
        'review_required': False,
        'review_reasons': [],
        'severity': 'normal',
        'issues': {}
    }
    
    doi = None
    if extracted_pdf_data and extracted_pdf_data.get('data', {}).get('doi'):
        doi = extracted_pdf_data['data']['doi']
    
    if not doi:
        review_result['review_required'] = True
        review_result['review_reasons'].append('DOI 누락')
        review_result['issues']['doi_missing'] = True
        review_result['severity'] = 'critical'
    else:
        if not re.match(r'^10\.\d{4,}/[^\s]+', doi):
            review_result['review_required'] = True
            review_result['review_reasons'].append('DOI 형식 오류')
            review_result['issues']['doi_format_error'] = True
            review_result['severity'] = 'critical'
        
        if crossref_data is None:
            review_result['review_required'] = True
            review_result['review_reasons'].append('DOI 조회 확인 필요')
            review_result['issues']['doi_not_verified'] = True
            if review_result['severity'] != 'critical':
                review_result['severity'] = 'warning'
    
    if extracted_pdf_data and extracted_pdf_data.get('data', {}).get('acknowledgement'):
        ack = extracted_pdf_data['data']['acknowledgement']
        if not ack.get('found'):
            review_result['review_required'] = True
            review_result['review_reasons'].append('사사 문구 미포함')
            review_result['issues']['acknowledgement_missing'] = True
            if review_result['severity'] != 'critical':
                review_result['severity'] = 'warning'
    
    if extracted_pdf_data and crossref_data:
        pdf_title = (extracted_pdf_data.get('data', {}).get('title') or '').lower().strip()
        crossref_title = (crossref_data.get('title') or '').lower().strip()
        
        if pdf_title and crossref_title:
            if pdf_title != crossref_title and len(pdf_title) > 10 and pdf_title not in crossref_title:
                review_result['review_required'] = True
                review_result['review_reasons'].append('제목 불일치')
                review_result['issues']['title_mismatch'] = True
                if review_result['severity'] != 'critical':
                    review_result['severity'] = 'warning'
    
    if matching_data.get('timeliness') == '지연 제출':
        review_result['review_required'] = True
        review_result['review_reasons'].append(f"제출기한 초과 ({matching_data.get('delay_days', 0)}일)")
        review_result['issues']['deadline_exceeded'] = True
        if review_result['severity'] != 'critical':
            review_result['severity'] = 'warning'
    
    if matching_data.get('submission_status') == '미제출':
        review_result['review_required'] = True
        review_result['review_reasons'].append('미제출')
        review_result['issues']['not_submitted'] = True
        review_result['severity'] = 'critical'
    
    return review_result


def create_review_summary_table(df_matching, review_results_dict):
    """검수 결과를 테이블로 정리"""
    review_summary = []
    
    for idx, row in df_matching.iterrows():
        file_name = row['file_name']
        review = review_results_dict.get(file_name, {
            'review_required': False,
            'review_reasons': [],
            'severity': 'normal'
        })
        
        review_summary.append({
            'researcher_id': row['researcher_id'],
            'researcher_name': row['researcher_name'],
            'file_name': file_name,
            'submission_status': row['submission_status'],
            'timeliness': row['timeliness'],
            'review_required': '검토 필요' if review['review_required'] else '정상',
            'review_reasons': '; '.join(review['review_reasons']) if review['review_reasons'] else '없음',
            'severity': review['severity']
        })
    
    return pd.DataFrame(review_summary)


def create_not_submitted_report(df_matching, reference_date=None):
    """미제출자 리포트 생성"""
    if reference_date is None:
        reference_date = datetime.now().date()
    
    not_submitted = df_matching[df_matching['submission_status'] == '미제출'].copy()
    not_submitted['due_date'] = pd.to_datetime(not_submitted['due_date'])
    not_submitted['days_remaining'] = (not_submitted['due_date'] - pd.Timestamp(reference_date)).dt.days
    not_submitted = not_submitted.sort_values('days_remaining')
    
    report = not_submitted[[
        'researcher_id',
        'researcher_name',
        'affiliation',
        'email',
        'project_name',
        'expected_output_type',
        'due_date',
        'days_remaining'
    ]].copy()
    
    report['due_date'] = report['due_date'].dt.strftime('%Y-%m-%d')
    report.columns = ['연구자ID', '연구자명', '소속', '이메일', '사업명', '성과유형', '제출기한', '남은일수']
    
    return report


def query_crossref(doi, timeout=5):
    """Crossref API 조회"""
    if not doi or pd.isna(doi):
        return {'success': False, 'error': 'DOI가 없습니다', 'metadata': {}}
    
    doi_str = str(doi).strip()
    if doi_str.startswith('https://doi.org/'):
        doi_str = doi_str.replace('https://doi.org/', '')
    elif doi_str.startswith('http://doi.org/'):
        doi_str = doi_str.replace('http://doi.org/', '')
    
    if not re.match(r'^10\.\S+/\S+', doi_str):
        return {'success': False, 'error': f'DOI 형식 오류: {doi_str}', 'metadata': {}}
    
    try:
        url = f"https://api.crossref.org/v1/works/{doi_str}"
        headers = {'User-Agent': 'Research-Management-System/1.0'}
        response = requests.get(url, headers=headers, timeout=timeout)
        
        if response.status_code == 404:
            return {'success': False, 'error': 'DOI를 찾을 수 없습니다', 'metadata': {}}
        
        if response.status_code != 200:
            return {'success': False, 'error': f'API 오류 (상태 코드: {response.status_code})', 'metadata': {}}
        
        data = response.json()
        
        if 'message' not in data:
            return {'success': False, 'error': 'API 응답 형식 오류', 'metadata': {}}
        
        work = data['message']
        
        metadata = {
            'doi': work.get('DOI', ''),
            'title': work.get('title', [''])[0] if isinstance(work.get('title'), list) else work.get('title', ''),
            'authors': ', '.join([f"{a.get('given', '')} {a.get('family', '')}".strip() 
                                 for a in work.get('author', [])[:5]]),
            'journal': work.get('container-title', [''])[0] if isinstance(work.get('container-title'), list) else work.get('container-title', ''),
            'published': work.get('published-print', work.get('published-online', {})),
            'publisher': work.get('publisher', ''),
            'url': work.get('URL', ''),
            'type': work.get('type', '')
        }
        
        return {'success': True, 'error': None, 'metadata': metadata}
    
    except requests.exceptions.Timeout:
        return {'success': False, 'error': '요청 타임아웃', 'metadata': {}}
    
    except requests.exceptions.ConnectionError:
        return {'success': False, 'error': '네트워크 오류', 'metadata': {}}
    
    except Exception as e:
        return {'success': False, 'error': f'오류: {str(e)}', 'metadata': {}}


def compare_metadata(pdf_data, crossref_data):
    """메타데이터 비교"""
    comparison = {'matches': [], 'mismatches': [], 'missing_in_pdf': [], 'missing_in_crossref': []}
    
    pdf_title = (pdf_data.get('title', '') or '').lower().strip()
    crossref_title = (crossref_data.get('title', '') or '').lower().strip()
    
    if pdf_title and crossref_title:
        if pdf_title == crossref_title or len(pdf_title) > 20 and pdf_title in crossref_title:
            comparison['matches'].append('제목 일치')
        else:
            comparison['mismatches'].append({'item': '제목', 'pdf': pdf_data.get('title', 'N/A'), 'crossref': crossref_data.get('title', 'N/A')})
    elif not pdf_title:
        comparison['missing_in_pdf'].append('제목')
    
    pdf_authors = (pdf_data.get('authors', '') or '').lower().strip()
    crossref_authors = (crossref_data.get('authors', '') or '').lower().strip()
    
    if pdf_authors and crossref_authors:
        if len(pdf_authors) > 3 and pdf_authors in crossref_authors:
            comparison['matches'].append('저자 일치')
    
    if crossref_data.get('published') and not pdf_data.get('published_date'):
        comparison['missing_in_pdf'].append('발행일')
    
    return comparison


def extract_text_from_pdf(pdf_file_path):
    """PDF 텍스트 추출"""
    try:
        pdf_document = fitz.open(pdf_file_path)
        full_text = ""
        for page_num in range(len(pdf_document)):
            page = pdf_document[page_num]
            full_text += page.get_text()
        pdf_document.close()
        
        if len(full_text.strip()) == 0:
            return {'success': False, 'text': '', 'error': '텍스트 추출 불가'}
        
        return {'success': True, 'text': full_text, 'error': None}
    
    except Exception as e:
        return {'success': False, 'text': '', 'error': f'PDF 읽기 오류: {str(e)}'}


def extract_doi_from_text(text):
    """텍스트에서 DOI 추출"""
    doi_pattern = r'10\.\d{4,}/[^\s]+'
    matches = re.findall(doi_pattern, text)
    
    if matches:
        return matches[0].rstrip('.,;:')
    
    return None


def extract_acknowledgement_from_text(text, keywords=None):
    """사사 문구 검색"""
    if keywords is None:
        keywords = [
            'acknowledgement', 'acknowledgment',
            'funding', 'supported by', 'grant',
            'acknowledge', 'thanks',
            '사사', '감사', '지원', '지원사업'
        ]
    
    text_lower = text.lower()
    found_keywords = []
    
    for keyword in keywords:
        if keyword.lower() in text_lower:
            found_keywords.append(keyword)
    
    return {'found': len(found_keywords) > 0, 'keywords': found_keywords[:5]}


def extract_author_from_text(text):
    """저자명 추출"""
    lines = text.split('\n')
    
    for i, line in enumerate(lines[:30]):
        line = line.strip()
        if len(line) < 100 and len(line) > 3:
            if re.match(r'^[A-Z][a-z]+ [A-Z]\.', line) or re.match(r'^[A-Z][a-z]+,', line):
                return line
    
    return "추출 불가"


def extract_title_from_text(text):
    """논문명 추출"""
    lines = text.split('\n')
    
    for line in lines[:20]:
        line = line.strip()
        
        if 20 <= len(line) <= 200 and line[0].isupper():
            if not re.match(r'^[\d\*\-\•]', line):
                return line
    
    return "추출 불가"


def extract_pdf_metadata(pdf_file_path):
    """PDF 메타데이터 추출"""
    extraction_result = extract_text_from_pdf(pdf_file_path)
    
    if not extraction_result['success']:
        return {'success': False, 'error': extraction_result['error'], 'data': {}}
    
    text = extraction_result['text']
    
    extracted_data = {
        'title': extract_title_from_text(text),
        'authors': extract_author_from_text(text),
        'doi': extract_doi_from_text(text),
        'acknowledgement': extract_acknowledgement_from_text(text),
        'full_text': text[:500]
    }
    
    return {'success': True, 'error': None, 'data': extracted_data}


def extract_researcher_id_from_filename(filename):
    """파일명에서 연구자 ID 추출"""
    name_without_ext = os.path.splitext(filename)[0]
    parts = name_without_ext.split('_')
    
    if len(parts) > 0 and parts[0].startswith('R') and parts[0][1:].isdigit():
        return parts[0]
    
    return None


def match_pdfs_to_researchers(df_researchers, pdf_filenames, submission_date):
    """PDF를 연구자와 매칭"""
    matching_results = []
    unmatched_files = []
    
    researcher_dict = {}
    for idx, row in df_researchers.iterrows():
        researcher_id = row['researcher_id']
        if researcher_id not in researcher_dict:
            researcher_dict[researcher_id] = {
                'name': row['researcher_name'],
                'project': row['project_name'],
                'due_date': row['due_date'],
                'expected_output_type': row['expected_output_type'],
                'affiliation': row.get('affiliation', ''),
                'email': row.get('email', ''),
                'files': []
            }
    
    for pdf_filename in pdf_filenames:
        extracted_id = extract_researcher_id_from_filename(pdf_filename)
        
        if extracted_id and extracted_id in researcher_dict:
            researcher_info = researcher_dict[extracted_id]
            due_date_obj = datetime.strptime(researcher_info['due_date'], '%Y-%m-%d')
            submission_date_obj = datetime.combine(submission_date, datetime.min.time())
            
            if submission_date_obj.date() <= due_date_obj.date():
                timeliness = '기한 내 제출'
            else:
                timeliness = '지연 제출'
            
            delay_days = (submission_date_obj.date() - due_date_obj.date()).days
            
            matching_results.append({
                'researcher_id': extracted_id,
                'researcher_name': researcher_info['name'],
                'affiliation': researcher_info.get('affiliation', ''),
                'email': researcher_info.get('email', ''),
                'project_name': researcher_info['project'],
                'due_date': researcher_info['due_date'],
                'expected_output_type': researcher_info['expected_output_type'],
                'submission_status': '제출',
                'submission_date': submission_date.strftime('%Y-%m-%d'),
                'timeliness': timeliness,
                'delay_days': delay_days if delay_days > 0 else 0,
                'file_name': pdf_filename,
            })
        else:
            unmatched_files.append({
                'file_name': pdf_filename,
                'extracted_id': extracted_id if extracted_id else '추출 불가',
                'reason': '해당하는 연구자를 찾을 수 없음'
            })
    
    submitted_ids = [r['researcher_id'] for r in matching_results]
    for researcher_id, researcher_info in researcher_dict.items():
        if researcher_id not in submitted_ids:
            matching_results.append({
                'researcher_id': researcher_id,
                'researcher_name': researcher_info['name'],
                'affiliation': researcher_info.get('affiliation', ''),
                'email': researcher_info.get('email', ''),
                'project_name': researcher_info['project'],
                'due_date': researcher_info['due_date'],
                'expected_output_type': researcher_info['expected_output_type'],
                'submission_status': '미제출',
                'submission_date': '-',
                'timeliness': '미제출',
                'delay_days': 0,
                'file_name': '-',
            })
    
    return matching_results, unmatched_files


# ============================================================
# 사이드바 설정 (수정: 비용 설정 제거)
# ============================================================

st.sidebar.header("📋 설정")

# 1. 기준일
st.sidebar.subheader("1️⃣ 기준일")
submission_date = st.sidebar.date_input(
    "제출일 선택",
    value=datetime.now().date()
)
st.sidebar.info(f"선택된 제출일: {submission_date.strftime('%Y-%m-%d')}")

# 2. 연구자 명단
st.sidebar.subheader("2️⃣ 연구자 명단")
upload_option = st.sidebar.radio(
    "파일 선택 방법",
    options=["샘플 데이터 사용", "CSV/Excel 업로드"]
)

df_researchers = None

if upload_option == "샘플 데이터 사용":
    try:
        df_researchers = pd.read_csv("sample_data.csv")
        st.sidebar.success("sample_data.csv 로드됨")
    except FileNotFoundError:
        st.sidebar.error("sample_data.csv를 찾을 수 없습니다")

elif upload_option == "CSV/Excel 업로드":
    uploaded_file = st.sidebar.file_uploader(
        "연구자 명단 파일",
        type=['csv', 'xlsx', 'xls']
    )
    
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df_researchers = pd.read_csv(uploaded_file)
            else:
                df_researchers = pd.read_excel(uploaded_file)
            st.sidebar.success(f"{uploaded_file.name} 로드됨")
        except Exception as e:
            st.sidebar.error(f"파일 읽기 오류: {str(e)}")

# 3. 제출 PDF
st.sidebar.subheader("3️⃣ 제출 PDF 파일")
pdf_upload_option = st.sidebar.radio(
    "PDF 선택 방법",
    options=["샘플 PDF 사용", "PDF 업로드"]
)

pdf_filenames = []
pdf_file_dict = {}

if pdf_upload_option == "샘플 PDF 사용":
    if os.path.exists('sample_pdfs'):
        pdf_files_in_dir = os.listdir('sample_pdfs')
        pdf_files = [f for f in pdf_files_in_dir if f.endswith('.pdf')]
        pdf_filenames = pdf_files
        
        for pdf_file in pdf_files:
            pdf_file_dict[pdf_file] = os.path.join('sample_pdfs', pdf_file)
        
        if pdf_files:
            st.sidebar.success(f"샘플 PDF {len(pdf_files)}개 발견")

elif pdf_upload_option == "PDF 업로드":
    uploaded_pdfs = st.sidebar.file_uploader(
        "PDF 파일 선택 (여러 개 가능)",
        type=['pdf'],
        accept_multiple_files=True
    )
    
    if uploaded_pdfs:
        pdf_filenames = [f.name for f in uploaded_pdfs]
        
        import tempfile
        for pdf_file in uploaded_pdfs:
            temp_path = os.path.join(tempfile.gettempdir(), pdf_file.name)
            with open(temp_path, 'wb') as f:
                f.write(pdf_file.getbuffer())
            pdf_file_dict[pdf_file.name] = temp_path
        
        st.sidebar.success(f"PDF {len(uploaded_pdfs)}개 업로드됨")

# ============================================================
# 메인 화면
# ============================================================

if df_researchers is None:
    st.info("왼쪽 사이드바에서 연구자 명단을 선택하십시오")

else:
    # 탭 구성 (수정: 9개 탭으로 축소)
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9 = st.tabs([
        "🎯 종합 대시보드",
        "⚠️ 미제출자 관리",
        "🔍 검토 필요",
        "📊 통계 분석",
        "📋 제출 현황",
        "📈 기한 분석",
        "📄 PDF 정보",
        "✅ DOI 검증",
        "💾 다운로드"
    ])
    
    # PDF 매칭 실행
    if len(pdf_filenames) > 0:
        matching_results, unmatched_files = match_pdfs_to_researchers(
            df_researchers,
            pdf_filenames,
            submission_date
        )
        df_matching = pd.DataFrame(matching_results)
        
        st.session_state['df_matching'] = df_matching
        st.session_state['unmatched_files'] = unmatched_files
        st.session_state['pdf_file_dict'] = pdf_file_dict
    else:
        st.warning("제출된 PDF 파일이 없습니다. 사이드바에서 PDF를 선택하세요")
        df_matching = None
    
    # ========================================
    # Tab 1: 종합 대시보드
    # ========================================
    
    with tab1:
        st.subheader("🎯 전체 현황 종합 대시보드")
        
        if df_matching is not None:
            total_count = len(df_matching)
            submitted_count = len(df_matching[df_matching['submission_status'] == '제출'])
            not_submitted_count = len(df_matching[df_matching['submission_status'] == '미제출'])
            on_time_count = len(df_matching[df_matching['timeliness'] == '기한 내 제출'])
            delayed_count = len(df_matching[df_matching['timeliness'] == '지연 제출'])
            
            st.markdown("### 📊 핵심 성과 지표 (KPI)")
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("전체 연구자", total_count)
            with col2:
                submission_rate = (submitted_count / total_count * 100) if total_count > 0 else 0
                st.metric("제출율", f"{submission_rate:.1f}%", f"+{submitted_count}명")
            with col3:
                timely_rate = (on_time_count / total_count * 100) if total_count > 0 else 0
                st.metric("기한 준수율", f"{timely_rate:.1f}%", f"+{on_time_count}명")
            with col4:
                st.metric("미제출", not_submitted_count, f"-{not_submitted_count}명")
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("제출 완료", submitted_count)
            with col2:
                st.metric("기한 내", on_time_count)
            with col3:
                st.metric("지연 제출", delayed_count)
            with col4:
                project_count = df_matching['project_name'].nunique()
                st.metric("사업 수", project_count)
            
            st.markdown("---")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.markdown("**제출 현황**")
                submission_data = df_matching['submission_status'].value_counts()
                st.bar_chart(submission_data)
            
            with col2:
                st.markdown("**기한 준수**")
                timeliness_data = df_matching[df_matching['submission_status'] == '제출']['timeliness'].value_counts()
                st.bar_chart(timeliness_data)
            
            with col3:
                st.markdown("**사업별 분포**")
                project_data = df_matching['project_name'].value_counts()
                st.bar_chart(project_data)
            
            st.markdown("---")
            
            st.markdown("### 📋 연구자별 상세 현황")
            
            display_cols = ['researcher_id', 'researcher_name', 'project_name', 'submission_status', 
                           'timeliness', 'due_date', 'submission_date', 'file_name']
            available_cols = [col for col in display_cols if col in df_matching.columns]
            df_display = df_matching[available_cols]
            
            st.dataframe(
                df_display,
                use_container_width=True,
                height=400,
                hide_index=True
            )
        else:
            st.info("PDF 파일을 업로드하세요")
    
    # ========================================
    # Tab 2: 미제출자 관리
    # ========================================
    
    with tab2:
        st.subheader("⚠️ 미제출자 관리 및 독촉")
        
        if df_matching is not None:
            not_submitted_df = df_matching[df_matching['submission_status'] == '미제출']
            
            if len(not_submitted_df) > 0:
                st.warning(f"**{len(not_submitted_df)}명의 미제출자가 있습니다**")
                
                report = create_not_submitted_report(df_matching, submission_date)
                
                st.markdown("### 📋 미제출자 목록 (기한 순서)")
                st.dataframe(
                    report,
                    use_container_width=True,
                    height=400,
                    hide_index=True
                )
                
                st.markdown("---")
                
                st.markdown("### ⏰ 기한별 분류")
                
                urgent = len(report[report['남은일수'] <= 7])
                warning = len(report[(report['남은일수'] > 7) & (report['남은일수'] <= 30)])
                normal = len(report[report['남은일수'] > 30])
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("🔴 긴급 (7일 이내)", urgent)
                with col2:
                    st.metric("🟡 주의 (30일 이내)", warning)
                with col3:
                    st.metric("🟢 여유", normal)
            else:
                st.success("✅ 모든 연구자가 제출했습니다!")
        else:
            st.info("PDF 파일을 업로드하세요")
    
    # ========================================
    # Tab 3: 검토 필요
    # ========================================
    
    with tab3:
        st.subheader("🔍 검토 필요 성과")
        
        st.info("💡 검수를 먼저 실행해주십시오 (후속 탭 참조)")
        
        if 'review_done' in st.session_state and st.session_state['review_done']:
            review_results = st.session_state['review_results']
            df_review = create_review_summary_table(df_matching, review_results)
            
            review_required_df = df_review[df_review['review_required'] == '검토 필요']
            
            if len(review_required_df) > 0:
                st.markdown("### 📋 검토 필요 성과 목록")
                st.dataframe(review_required_df, use_container_width=True, hide_index=True)
            else:
                st.success("✅ 검토 필요한 성과가 없습니다!")
    
    # ========================================
    # Tab 4: 통계 분석
    # ========================================
    
    with tab4:
        st.subheader("📊 통계 및 분석")
        
        if df_matching is not None:
            total = len(df_matching)
            submitted = len(df_matching[df_matching['submission_status'] == '제출'])
            not_submitted = len(df_matching[df_matching['submission_status'] == '미제출'])
            on_time = len(df_matching[df_matching['timeliness'] == '기한 내 제출'])
            delayed = len(df_matching[df_matching['timeliness'] == '지연 제출'])
            
            stats_data = {
                '항목': [
                    '전체 연구자',
                    '제출 완료',
                    '미제출',
                    '기한 내',
                    '지연 제출',
                    '제출율',
                    '기한 준수율'
                ],
                '수량': [
                    total,
                    submitted,
                    not_submitted,
                    on_time,
                    delayed,
                    f"{(submitted/total*100):.1f}%" if total > 0 else "0%",
                    f"{(on_time/total*100):.1f}%" if total > 0 else "0%"
                ]
            }
            
            df_stats = pd.DataFrame(stats_data)
            
            st.markdown("### 📈 주요 통계")
            st.dataframe(df_stats, use_container_width=True, hide_index=True)
            
            st.markdown("---")
            
            st.markdown("### 🏢 사업별 분석")
            
            project_analysis = []
            for project in df_matching['project_name'].unique():
                proj_data = df_matching[df_matching['project_name'] == project]
                project_analysis.append({
                    '사업명': project,
                    '전체': len(proj_data),
                    '제출': len(proj_data[proj_data['submission_status'] == '제출']),
                    '미제출': len(proj_data[proj_data['submission_status'] == '미제출']),

                    '제출율': f"{(len(proj_data[proj_data['submission_status'] == '제출'])/len(proj_data)*100):.1f}%"
                })
            
            df_project_analysis = pd.DataFrame(project_analysis)
            st.dataframe(df_project_analysis, use_container_width=True, hide_index=True)
        else:
            st.info("PDF 파일을 업로드하세요")
    
    # ========================================
    # Tab 5: 제출 현황
    # ========================================
    
    with tab5:
        st.subheader("📊 제출 현황 요약")
        
        if df_matching is not None:
            total_researchers = len(df_researchers)
            submitted_count = len(df_matching[df_matching['submission_status'] == '제출'])
            not_submitted_count = len(df_matching[df_matching['submission_status'] == '미제출'])
            on_time_count = len(df_matching[df_matching['timeliness'] == '기한 내 제출'])
            delayed_count = len(df_matching[df_matching['timeliness'] == '지연 제출'])
            
            col1, col2, col3, col4, col5 = st.columns(5)
            
            with col1:
                st.metric("전체 연구자", total_researchers)
            with col2:
                st.metric("제출 완료", submitted_count, f"{(submitted_count/total_researchers*100):.1f}%")
            with col3:
                st.metric("미제출", not_submitted_count, f"{(not_submitted_count/total_researchers*100):.1f}%")
            with col4:
                st.metric("기한 내", on_time_count, f"{(on_time_count/total_researchers*100):.1f}%")
            with col5:
                st.metric("지연", delayed_count, f"{(delayed_count/total_researchers*100):.1f}%")
            
            st.markdown("---")
            
            col1, col2 = st.columns(2)
            with col1:
                submission_counts = df_matching['submission_status'].value_counts()
                st.bar_chart(submission_counts)
            with col2:
                timeliness_counts = df_matching[df_matching['submission_status'] == '제출']['timeliness'].value_counts()
                st.bar_chart(timeliness_counts)
        else:
            st.info("PDF 파일을 업로드하세요")
    
    # ========================================
    # Tab 6: 기한 분석
    # ========================================
    
    with tab6:
        st.subheader("📈 기한 분석")
        
        if df_matching is not None:
            submitted_df = df_matching[df_matching['submission_status'] == '제출']
            
            if len(submitted_df) > 0:
                st.markdown("### 📊 제출자 기한 준수 현황")
                
                on_time = len(submitted_df[submitted_df['timeliness'] == '기한 내 제출'])
                delayed = len(submitted_df[submitted_df['timeliness'] == '지연 제출'])
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.metric("기한 내 제출", on_time, f"{(on_time/(on_time+delayed)*100):.1f}%")
                with col2:
                    st.metric("기한 초과", delayed, f"{(delayed/(on_time+delayed)*100):.1f}%")
                
                st.markdown("---")
                
                if delayed > 0:
                    st.markdown("### ⚠️ 기한 초과 연구자")
                    delayed_df = submitted_df[submitted_df['timeliness'] == '지연 제출'][
                        ['researcher_id', 'researcher_name', 'due_date', 'submission_date', 'delay_days']
                    ].sort_values('delay_days', ascending=False)
                    st.dataframe(delayed_df, use_container_width=True, hide_index=True)
            
            st.markdown("---")
            st.markdown("### ❌ 미제출 연구자")
            
            not_submitted_df = df_matching[df_matching['submission_status'] == '미제출'][
                ['researcher_id', 'researcher_name', 'due_date', 'expected_output_type']
            ]
            
            if len(not_submitted_df) > 0:
                st.dataframe(not_submitted_df, use_container_width=True, hide_index=True)
            else:
                st.success("모든 연구자가 제출했습니다!")
        else:
            st.info("PDF 파일을 업로드하세요")
    
    # ========================================
    # Tab 7: PDF 정보 추출
    # ========================================
    
    with tab7:
        st.subheader("📄 PDF 정보 추출")
        
        if df_matching is not None:
            submitted_df = df_matching[df_matching['submission_status'] == '제출']
            
            if len(submitted_df) > 0:
                pdf_options = submitted_df[['researcher_id', 'researcher_name', 'file_name']].copy()
                pdf_options['display'] = pdf_options.apply(
                    lambda x: f"{x['researcher_id']} {x['researcher_name']} - {x['file_name']}", axis=1
                )
                
                selected_pdf_display = st.selectbox(
                    "PDF 선택",
                    options=pdf_options['display'].tolist(),
                    key="tab7_pdf_select"
                )
                
                selected_idx = pdf_options['display'].tolist().index(selected_pdf_display)
                selected_filename = pdf_options.iloc[selected_idx]['file_name']
                
                if st.button("🔍 정보 추출", type="primary", use_container_width=True, key="tab7_extract"):
                    if selected_filename in st.session_state.get('pdf_file_dict', {}):
                        pdf_path = st.session_state['pdf_file_dict'][selected_filename]
                        
                        with st.spinner("PDF를 분석 중입니다..."):
                            extraction_result = extract_pdf_metadata(pdf_path)
                        
                        if extraction_result['success']:
                            st.session_state['extracted_pdf_data'] = {
                                'filename': selected_filename,
                                'data': extraction_result['data']
                            }
                            st.success("✅ 정보 추출 완료!")
                        else:
                            st.error(f"❌ 추출 실패: {extraction_result['error']}")
                
                if 'extracted_pdf_data' in st.session_state:
                    extracted = st.session_state['extracted_pdf_data']['data']
                    
                    st.markdown("---")
                    st.markdown("### 📋 추출된 정보")
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.markdown("#### 논문명")
                        title = st.text_area(
                            "논문명",
                            value=extracted.get('title', ''),
                            height=60,
                            label_visibility="collapsed"
                        )
                    
                    with col2:
                        st.markdown("#### 저자")
                        authors = st.text_area(
                            "저자",
                            value=extracted.get('authors', ''),
                            height=60,
                            label_visibility="collapsed"
                        )
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.markdown("#### DOI")
                        doi = st.text_input(
                            "DOI",
                            value=extracted.get('doi') or '',
                            label_visibility="collapsed"
                        )
                        if extracted.get('doi'):
                            st.success(f"✅ DOI 발견")
                        else:
                            st.warning("⚠️ DOI를 찾을 수 없습니다")
                    
                    with col2:
                        st.markdown("#### 사사 문구")
                        ack = extracted.get('acknowledgement', {})
                        if ack.get('found'):
                            st.success(f"✅ 사사 발견: {', '.join(ack.get('keywords', []))}")
                        else:
                            st.warning("⚠️ 사사 문구를 찾을 수 없습니다")
                    
                    st.markdown("---")
                    if st.button("💾 정보 저장", use_container_width=True, key="tab7_save"):
                        st.session_state['extracted_pdf_data']['data'] = {
                            'title': title,
                            'authors': authors,
                            'doi': doi,
                            'acknowledgement': extracted.get('acknowledgement', {})
                        }
                        st.success("✅ 정보가 저장되었습니다")
            else:
                st.info("제출한 PDF가 없습니다")
        else:
            st.info("PDF 파일을 업로드하세요")
    
    # ========================================
    # Tab 8: DOI 검증
    # ========================================
    
    with tab8:
        st.subheader("✅ Crossref DOI 검증")
        
        if 'extracted_pdf_data' in st.session_state:
            extracted_data = st.session_state['extracted_pdf_data']['data']
            doi = extracted_data.get('doi', '')
            
            if doi:
                st.info(f"조회 DOI: **{doi}**")
                
                if st.button("🔍 DOI 검증", type="primary", use_container_width=True, key="tab8_validate"):
                    with st.spinner("Crossref API에서 DOI 정보를 조회 중입니다..."):
                        crossref_result = query_crossref(doi)
                    
                    if crossref_result['success']:
                        st.session_state['crossref_data'] = crossref_result['metadata']
                        st.success("✅ DOI 검증 완료!")
                    else:
                        st.error(f"❌ 검증 실패: {crossref_result['error']}")
                
                if 'crossref_data' in st.session_state:
                    crossref = st.session_state['crossref_data']
                    
                    st.markdown("---")
                    st.markdown("### 📊 Crossref 메타데이터")
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.markdown("#### Crossref 정보")
                        st.write(f"**제목**: {crossref.get('title', 'N/A')}")
                        st.write(f"**저자**: {crossref.get('authors', 'N/A')}")
                        st.write(f"**학술지**: {crossref.get('journal', 'N/A')}")
                        st.write(f"**발행일**: {crossref.get('published', 'N/A')}")
                    
                    with col2:
                        st.markdown("#### PDF 정보와 비교")
                        comparison = compare_metadata(extracted_data, crossref)
                        
                        if comparison['matches']:
                            st.success("**✅ 일치 항목**")
                            for match in comparison['matches']:
                                st.write(f"- {match}")
                        
                        if comparison['mismatches']:
                            st.warning("**⚠️ 불일치 항목**")
                            for mismatch in comparison['mismatches']:
                                st.write(f"- **{mismatch['item']}**")
                                st.write(f"  - PDF: {mismatch['pdf']}")
                                st.write(f"  - Crossref: {mismatch['crossref']}")
                        
                        if comparison['missing_in_pdf']:
                            st.info("**ℹ️ PDF에서 누락된 정보**")
                            for missing in comparison['missing_in_pdf']:
                                st.write(f"- {missing}")
            else:
                st.warning("⚠️ DOI가 없습니다. Tab 7에서 DOI를 입력하세요")
        else:
            st.info("Tab 7에서 PDF 정보를 먼저 추출하세요")
    
    # ========================================
    # Tab 9: 검수 현황
    # ========================================
    
    with tab9:
        st.subheader("🔍 연구성과 자동 검수 현황")
        
        if df_matching is not None:
            st.info("💡 검수는 자동으로 수행됩니다. 각 성과에 대해 8가지 기준으로 검증합니다.")
            
            if st.button("🔍 자동 검수 실행", type="primary", use_container_width=True, key="tab9_review"):
                review_results = {}
                
                with st.spinner("검수 중입니다..."):
                    for idx, row in df_matching.iterrows():
                        file_name = row['file_name']
                        
                        extracted_data = None
                        crossref_data = None
                        
                        if 'extracted_pdf_data' in st.session_state:
                            if st.session_state['extracted_pdf_data']['filename'] == file_name:
                                extracted_data = st.session_state['extracted_pdf_data']['data']
                        
                        if 'crossref_data' in st.session_state:
                            if 'extracted_pdf_data' in st.session_state and \
                               st.session_state['extracted_pdf_data']['filename'] == file_name:
                                crossref_data = st.session_state['crossref_data']
                        
                        review = perform_research_review(row.to_dict(), extracted_data, crossref_data)
                        review_results[file_name] = review
                
                st.session_state['review_results'] = review_results
                st.session_state['review_done'] = True
                st.success("✅ 검수 완료!")
            
            if 'review_done' in st.session_state and st.session_state['review_done']:
                review_results = st.session_state['review_results']
                df_review = create_review_summary_table(df_matching, review_results)
                
                st.markdown("---")
                st.markdown("### 📊 검수 결과 요약")
                
                normal_count = len(df_review[df_review['review_required'] == '정상'])
                review_required_count = len(df_review[df_review['review_required'] == '검토 필요'])
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("전체 성과", len(df_review))
                
                with col2:
                    st.metric(
                        "정상 성과",
                        normal_count,
                        f"{(normal_count/len(df_review)*100):.1f}%"
                    )
                
                with col3:
                    st.metric(
                        "검토 필요",
                        review_required_count,
                        f"{(review_required_count/len(df_review)*100):.1f}%"
                    )
                
                st.markdown("---")
                
                st.markdown("### 📈 검수 결과 분포")
                review_counts = df_review['review_required'].value_counts()
                st.bar_chart(review_counts)
                
                st.markdown("---")
                
                st.markdown("### ⚠️ 심각도별 분류")
                
                critical_count = len(df_review[df_review['severity'] == 'critical'])
                warning_count = len(df_review[df_review['severity'] == 'warning'])
                normal_severity_count = len(df_review[df_review['severity'] == 'normal'])
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("심각 (Critical)", critical_count)
                
                with col2:
                    st.metric("경고 (Warning)", warning_count)
                
                with col3:
                    st.metric("정상 (Normal)", normal_severity_count)
                
                st.markdown("---")
                
                if review_required_count > 0:
                    st.markdown("### 📋 검토 필요 성과 상세")
                    
                    df_review_required = df_review[df_review['review_required'] == '검토 필요']
                    
                    for idx, row in df_review_required.iterrows():
                        with st.expander(f"🔴 {row['researcher_id']} {row['researcher_name']} - {row['file_name']}", 
                                        expanded=False):
                            st.write(f"**제출 상태**: {row['submission_status']}")
                            st.write(f"**기한 준수**: {row['timeliness']}")
                            st.write(f"**심각도**: {row['severity']}")
                            st.write(f"**검토 사유**:")
                            reasons = row['review_reasons'].split('; ')
                            for reason in reasons:
                                st.write(f"- {reason}")
        else:
            st.info("PDF 파일을 업로드하세요")
    
    # ========================================
    # Tab 9 (이전 Tab 11): 다운로드
    # ========================================
    
    with tab9:
        # 이전 탭 수정 - 탭 구조에서 이미 9번째 탭이 "💾 다운로드"
        pass

# ============================================================
# 다운로드 탭 재구성 (Tab 9로 통합)
# ============================================================

# 실제 다운로드 탭은 위의 tab9에 추가되어야 함
# 코드 수정 필요

# 아래에서 다시 구성

if df_researchers is not None and len(pdf_filenames) > 0:
    # 다운로드 섹션을 메인에 추가 (탭 밖)
    pass