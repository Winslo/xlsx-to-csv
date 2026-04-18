import streamlit as st
import pandas as pd
import io

# ── 페이지 설정 ──────────────────────────────────────────────
st.set_page_config(
    page_title="XLSX → CSV 변환기",
    page_icon="📄",
    layout="centered",
)

# ── 커스텀 스타일 ────────────────────────────────────────────
st.markdown(
    """
    <style>
    /* 전체 배경 */
    .stApp {
        background: linear-gradient(135deg, #0f0c29, #302b63, #24243e);
    }

    /* 제목 */
    h1 {
        background: linear-gradient(90deg, #a78bfa, #60a5fa);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800 !important;
        letter-spacing: -0.5px;
    }

    /* 설명 텍스트 */
    .description {
        color: #c4b5fd;
        font-size: 1.05rem;
        margin-bottom: 1.5rem;
        line-height: 1.6;
    }

    /* 카드 컨테이너 */
    .card {
        background: rgba(255, 255, 255, 0.06);
        border: 1px solid rgba(255, 255, 255, 0.10);
        border-radius: 16px;
        padding: 1.5rem 1.8rem;
        margin-bottom: 1rem;
        backdrop-filter: blur(12px);
    }

    /* 시트 이름 배지 */
    .sheet-badge {
        display: inline-block;
        background: linear-gradient(135deg, #6366f1, #8b5cf6);
        color: #fff;
        padding: 4px 14px;
        border-radius: 999px;
        font-size: 0.85rem;
        font-weight: 600;
        margin: 4px 4px 4px 0;
    }

    /* 통계 숫자 */
    .stat-number {
        font-size: 1.8rem;
        font-weight: 700;
        background: linear-gradient(90deg, #a78bfa, #60a5fa);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .stat-label {
        color: #94a3b8;
        font-size: 0.85rem;
    }

    /* 푸터 */
    .footer {
        text-align: center;
        color: #64748b;
        font-size: 0.8rem;
        margin-top: 3rem;
        padding-top: 1rem;
        border-top: 1px solid rgba(255,255,255,0.06);
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ── 타이틀 영역 ──────────────────────────────────────────────
st.markdown("# 📄 XLSX → CSV 변환기")
st.markdown(
    '<p class="description">'
    "엑셀(.xlsx) 파일을 업로드하면 모든 시트를 하나의 CSV 파일로 합쳐 드립니다.<br>"
    "각 시트 앞에 <code>### SHEET: 시트이름 ###</code> 구분자가 자동으로 추가됩니다."
    "</p>",
    unsafe_allow_html=True,
)

# ── 파일 업로드 ──────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "📂 엑셀 파일을 여기에 드래그하거나 클릭하여 업로드하세요",
    type=["xlsx"],
    help="지원 형식: .xlsx (Excel 2007 이상)",
)

if uploaded_file is None:
    st.info("⬆️  .xlsx 파일을 업로드하면 변환이 시작됩니다.")
    st.stop()

# ── 파일 처리 ────────────────────────────────────────────────
try:
    xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
    sheet_names = xls.sheet_names
except Exception as e:
    st.error(f"❌ 파일을 읽는 중 오류가 발생했습니다:\n\n`{e}`")
    st.stop()

if not sheet_names:
    st.warning("⚠️ 파일에 시트가 하나도 없습니다.")
    st.stop()

# ── 시트별 처리 & CSV 생성 ───────────────────────────────────
output = io.StringIO()
total_rows = 0
empty_sheets = []
filled_sheets = []

for idx, sheet_name in enumerate(sheet_names):
    try:
        df = pd.read_excel(xls, sheet_name=sheet_name, engine="openpyxl", header=None)
    except Exception as e:
        st.warning(f"⚠️ '{sheet_name}' 시트를 읽는 중 오류 발생 → 건너뜁니다. ({e})")
        continue

    # 구분자 줄 추가
    if idx > 0:
        output.write("\n")  # 시트 사이 빈 줄
    output.write(f"### SHEET: {sheet_name} ###\n")

    if df.empty:
        empty_sheets.append(sheet_name)
        output.write("(빈 시트)\n")
    else:
        filled_sheets.append(sheet_name)
        total_rows += len(df)
        df.to_csv(output, index=False, header=False, encoding="utf-8-sig")

# ── 결과 통계 ────────────────────────────────────────────────
st.markdown("---")
st.markdown("### 📊 변환 결과")

col1, col2, col3 = st.columns(3)
with col1:
    st.markdown(
        f'<div class="card" style="text-align:center">'
        f'<div class="stat-number">{len(sheet_names)}</div>'
        f'<div class="stat-label">전체 시트</div></div>',
        unsafe_allow_html=True,
    )
with col2:
    st.markdown(
        f'<div class="card" style="text-align:center">'
        f'<div class="stat-number">{len(filled_sheets)}</div>'
        f'<div class="stat-label">데이터 있는 시트</div></div>',
        unsafe_allow_html=True,
    )
with col3:
    st.markdown(
        f'<div class="card" style="text-align:center">'
        f'<div class="stat-number">{total_rows:,}</div>'
        f'<div class="stat-label">총 행 수</div></div>',
        unsafe_allow_html=True,
    )

# ── 시트 목록 ────────────────────────────────────────────────
st.markdown("#### 🗂️ 시트 목록")
badges_html = ""
for name in sheet_names:
    if name in empty_sheets:
        badges_html += (
            f'<span class="sheet-badge" style="background:linear-gradient(135deg,#64748b,#475569)">'
            f"🚫 {name} (빈 시트)</span> "
        )
    else:
        badges_html += f'<span class="sheet-badge">✅ {name}</span> '
st.markdown(f'<div style="margin-bottom:1rem">{badges_html}</div>', unsafe_allow_html=True)

if empty_sheets:
    st.caption(f"ℹ️ 빈 시트 {len(empty_sheets)}개는 '(빈 시트)'로 표기됩니다.")

# ── 미리보기 ─────────────────────────────────────────────────
csv_text = output.getvalue()

with st.expander("👀 CSV 미리보기 (처음 50줄)", expanded=False):
    preview_lines = csv_text.split("\n")[:50]
    st.code("\n".join(preview_lines), language="csv")

# ── 다운로드 버튼 ────────────────────────────────────────────
st.markdown("---")

original_name = uploaded_file.name.rsplit(".", 1)[0]
download_filename = f"{original_name}_합본.csv"

st.download_button(
    label="⬇️  CSV 파일 다운로드",
    data=csv_text.encode("utf-8-sig"),
    file_name=download_filename,
    mime="text/csv",
    use_container_width=True,
)

# ── 푸터 ─────────────────────────────────────────────────────
st.markdown(
    '<div class="footer">XLSX → CSV 변환기 · Streamlit · pandas · openpyxl</div>',
    unsafe_allow_html=True,
)
