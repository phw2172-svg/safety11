import sys
import subprocess
import os

# (1) 필요한 라이브러리 자동 설치 확인 및 로드
def install_and_import(package):
    """지정된 패키지가 설치되어 있는지 확인하고, 없으면 자동으로 설치합니다."""
    try:
        __import__(package)
        print(f"[DEBUG] 라이브러리 '{package}'이(가) 이미 설치되어 있습니다.")
    except ImportError:
        print(f"[DEBUG] 라이브러리 '{package}'이(가) 없어 설치를 진행합니다...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        print(f"[DEBUG] '{package}' 설치가 완료되었습니다.")

# 필수 라이브러리 검사
required_packages = ['streamlit', 'pandas', 'openpyxl', 'matplotlib', 'seaborn']
for pkg in required_packages:
    install_and_import(pkg)

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# (2) 한글 폰트 및 마이너스 기호 설정
def setup_matplotlib_font():
    """그래프에서 한글이 깨지지 않도록 '맑은 고딕' 폰트와 마이너스 기호를 설정합니다."""
    plt.rcParams['font.family'] = 'Malgun Gothic'
    plt.rcParams['axes.unicode_minus'] = False
    print("[DEBUG] 맷플롯립 한글 폰트('맑은 고딕') 및 마이너스 기호 설정 완료")

# (3) 스트림릿 페이지 설정 및 파일 업로더 구성
def configure_streamlit_page():
    """스트림릿 대시보드의 레이아웃을 'wide'로 설정합니다."""
    st.set_page_config(page_title="매출 분석 대시보드", layout="wide")
    print("[DEBUG] 스트림릿 페이지 설정 완료")

# (4) 엑셀 데이터 불러오기 및 전처리
@st.cache_data
def load_data(uploaded_file):
    """업로드된 엑셀 파일을 읽어와 데이터프레임으로 변환합니다."""
    df = pd.read_excel(uploaded_file)
    df['주문일자'] = pd.to_datetime(df['주문일자']).dt.strftime('%Y-%m-%d')
    print(f"[DEBUG] 데이터 로드 완료. 행 수: {len(df)}, 열 수: {len(df.columns)}")
    return df

# (5) 바탕화면에 전체 데이터가 포함된 HTML 파일 저장 기능 구현
def save_all_data_to_desktop(df):
    """전체 데이터(245건 전체)를 바탕화면에 '매출분석대시보드.html' 파일로 생성합니다."""
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    if not os.path.exists(desktop_path):
        desktop_path = os.path.expanduser("~")
        if os.path.exists(os.path.join(desktop_path, "바탕 화면")):
            desktop_path = os.path.join(desktop_path, "바탕 화면")

    output_file_path = os.path.join(desktop_path, "매출분석대시보드.html")
    
    # 전체 매출 및 주요 지표 계산
    total_revenue = df['매출금액'].sum()
    total_net_sales = df['순매출'].sum()
    total_orders = len(df)
    avg_order_value = total_net_sales / total_orders if total_orders > 0 else 0

    # 전체 데이터 테이블 HTML 변환
    html_table = df.to_html(classes='table table-striped', index=False)
    
    html_content = f"""<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>2026년 상반기 종합 매출분석 대시보드</title>
    <style>
        body {{ font-family: 'Malgun Gothic', sans-serif; background-color: #F4F6F9; padding: 30px; color: #333; }}
        .container {{ max-width: 1400px; margin: 0 auto; }}
        h1 {{ color: #1E3A8A; margin-bottom: 5px; }}
        p.subtitle {{ color: #6B7280; margin-bottom: 25px; }}
        
        .kpi-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; margin-bottom: 25px; }}
        .kpi-card {{ background: #FFFFFF; padding: 20px; border-radius: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.04); border-left: 5px solid #2563EB; }}
        .kpi-title {{ font-size: 12px; color: #6B7280; font-weight: bold; text-transform: uppercase; }}
        .kpi-value {{ font-size: 20px; color: #111827; font-weight: bold; margin-top: 8px; }}
        
        .card {{ background: #FFFFFF; padding: 25px; border-radius: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.04); margin-bottom: 25px; }}
        .table-container {{ max-height: 500px; overflow-y: auto; border: 1px solid #E5E7EB; border-radius: 6px; margin-top: 15px; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 12px; text-align: left; }}
        th, td {{ padding: 10px 12px; border-bottom: 1px solid #E5E7EB; }}
        th {{ background-color: #F8F9FA; color: #374151; position: sticky; top: 0; font-weight: bold; }}
        tr:hover {{ background: #F9FAFB; }}
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 2026년 상반기 종합 매출분석 대시보드</h1>
        <p class="subtitle">전체 주문내역 데이터 및 성과 지표 보고서</p>
        
        <div class="kpi-grid">
            <div class="kpi-card">
                <div class="kpi-title">[총매출]</div>
                <div class="kpi-value">{total_revenue:,}원</div>
            </div>
            <div class="kpi-card" style="border-left-color: #3B82F6;">
                <div class="kpi-title">[순매출]</div>
                <div class="kpi-value">{total_net_sales:,}원</div>
            </div>
            <div class="kpi-card" style="border-left-color: #10B981;">
                <div class="kpi-title">[주문건수]</div>
                <div class="kpi-value">{total_orders:,}건</div>
            </div>
            <div class="kpi-card" style="border-left-color: #F59E0B;">
                <div class="kpi-title">[평균 주문금액]</div>
                <div class="kpi-value">{int(avg_order_value):,}원</div>
            </div>
        </div>

        <div class="card">
            <h3>📋 [주문 상세 데이터 전체 목록] (총 {total_orders}건)</h3>
            <div class="table-container">
                {html_table}
            </div>
        </div>
    </div>
</body>
</html>
"""
    
    with open(output_file_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"[DEBUG] 바탕화면 전체 데이터 HTML 저장 완료: {output_file_path}")
    return output_file_path

# (6) 메인 실행 함수
def main():
    """스트림릿 대시보드의 전체 흐름을 제어하는 메인 함수입니다."""
    print("[DEBUG] 스트림릿 앱 실행 시작")
    setup_matplotlib_font()
    configure_streamlit_page()

    st.title("📊 2026년 상반기 종합 매출분석 대시보드")
    st.markdown("사이드바에 **주문내역 엑셀 파일**을 업로드하면 대시보드가 자동으로 생성됩니다.")

    st.sidebar.header("📁 데이터 파일 업로드")
    uploaded_file = st.sidebar.file_uploader("주문내역 엑셀 파일을 선택하세요", type=["xlsx", "xls"])

    if uploaded_file is None:
        st.info("👈 왼쪽 사이드바에서 엑셀 파일(.xlsx)을 업로드해 주세요.")
        return

    df = load_data(uploaded_file)

    st.sidebar.markdown("---")
    st.sidebar.header("🔍 월별 필터")
    months = ['전체보기', '1월', '2월', '3월', '4월', '5월', '6월']
    selected_month = st.sidebar.radio("조회할 월을 선택하세요:", months)

    if selected_month != '전체보기':
        filtered_df = df[df['주문월'] == selected_month]
    else:
        filtered_df = df

    # 화면 표시용 KPI (필터 적용)
    total_revenue = filtered_df['매출금액'].sum()
    total_net_sales = filtered_df['순매출'].sum()
    total_orders = len(filtered_df)
    avg_order_value = total_net_sales / total_orders if total_orders > 0 else 0

    col1, col2, col3, col4 = st.columns(4)
    with col1: st.metric(label="총매출", value=f"{total_revenue:,}원")
    with col2: st.metric(label="순매출", value=f"{total_net_sales:,}원")
    with col3: st.metric(label="주문건수", value=f"{total_orders:,}건")
    with col4: st.metric(label="평균 주문금액", value=f"{int(avg_order_value):,}원")

    st.markdown("---")

    # 차트 시각화
    plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
    plt.rcParams['font.family'] = 'Malgun Gothic'
    plt.rcParams['axes.unicode_minus'] = False

    c_col1, c_col2 = st.columns(2)
    with c_col1:
        st.subheader("[월별 매출 추이]")
        monthly_sales = df.groupby('주문월', as_index=False)['순매출'].sum()
        fig, ax = plt.subplots(figsize=(6, 3.5))
        bars = ax.bar(monthly_sales['주문월'], monthly_sales['순매출'] / 10000, color='#2563EB', width=0.5)
        ax.set_ylabel('순매출 (만원)')
        max_idx = monthly_sales['순매출'].idxmax()
        bars[max_idx].set_color('#F59E0B')
        st.pyplot(fig)

    with c_col2:
        st.subheader("[지점별 매출]")
        branch_sales = filtered_df.groupby('지점', as_index=False)['순매출'].sum().sort_values(by='순매출', ascending=True)
        fig, ax = plt.subplots(figsize=(6, 3.5))
        ax.barh(branch_sales['지점'], branch_sales['순매출'] / 10000, color='#3B82F6', height=0.5)
        st.pyplot(fig)

    # 바탕화면 전체 데이터 HTML 저장 버튼
    st.markdown("---")
    if st.button("💾 바탕화면에 전체 데이터 HTML 대시보드 저장하기"):
        saved_path = save_all_data_to_desktop(df)
        st.success(f"성공적으로 저장되었습니다! 저장 경로: {saved_path}")

    st.subheader(f"📋 [주문 상세 데이터] ({selected_month} 기준, 총 {len(filtered_df)}건)")
    st.dataframe(filtered_df, use_container_width=True)

if __name__ == "__main__":
    main()