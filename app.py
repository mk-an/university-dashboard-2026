# -*- coding: utf-8 -*-
import streamlit as st
import xlrd
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(
    page_title="2026학년도 수시 합격자 분석",
    page_icon="🎓",
    layout="wide"
)

# ── 몬드리안 스타일 CSS ───────────────────────────────────
st.markdown("""
<style>
/* ── 전체 배경 ── */
.stApp { background-color: #F8F6F0; }

/* ── 사이드바 ── */
[data-testid="stSidebar"] {
    background-color: #1A1A1A;
    border-right: 5px solid #000;
}
[data-testid="stSidebar"] * { color: #F0F0F0 !important; }
[data-testid="stSidebar"] .stRadio label,
[data-testid="stSidebar"] .stMultiSelect label,
[data-testid="stSidebar"] .stMarkdown p { color: #CCCCCC !important; }
[data-testid="stSidebar"] h1 {
    color: #F5C900 !important;
    font-size: 1.3rem !important;
    border-left: 5px solid #D40000;
    padding-left: 10px;
    letter-spacing: 1px;
}
[data-testid="stSidebar"] hr {
    border: none;
    border-top: 2px solid #333;
    margin: 10px 0;
}

/* ── 메인 타이틀 ── */
h1 {
    font-size: 1.7rem !important;
    font-weight: 900 !important;
    letter-spacing: 1px;
    padding: 14px 18px !important;
    background: #fff;
    border-left: 10px solid #D40000;
    border-bottom: 4px solid #000;
    margin-bottom: 4px !important;
}

/* ── 섹션 서브헤더 ── */
h2 {
    font-weight: 800 !important;
    padding: 6px 12px !important;
    border-left: 6px solid #1E3FA4;
    background: #fff;
    border-bottom: 2px solid #000;
    margin-top: 8px !important;
}
h3 {
    font-weight: 700 !important;
    border-left: 4px solid #F5C900;
    padding-left: 10px !important;
    color: #222 !important;
}

/* ── 구분선 ── */
hr {
    border: none !important;
    border-top: 3px solid #000 !important;
    margin: 18px 0 !important;
}

/* ── KPI 메트릭 카드 ── */
[data-testid="stMetric"] {
    background: #fff;
    border: 3px solid #000;
    border-radius: 0px !important;
    padding: 14px 10px !important;
}
[data-testid="stMetricLabel"] {
    font-weight: 700 !important;
    font-size: 0.75rem !important;
    color: #555 !important;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}
[data-testid="stMetricValue"] {
    font-size: 1.6rem !important;
    font-weight: 900 !important;
    color: #111 !important;
}

/* 카드 별 상단 컬러 바 (1~5번째) */
[data-testid="column"]:nth-child(1) [data-testid="stMetric"] { border-top: 7px solid #D40000; }
[data-testid="column"]:nth-child(2) [data-testid="stMetric"] { border-top: 7px solid #1E3FA4; }
[data-testid="column"]:nth-child(3) [data-testid="stMetric"] { border-top: 7px solid #F5C900; }
[data-testid="column"]:nth-child(4) [data-testid="stMetric"] { border-top: 7px solid #000;    }
[data-testid="column"]:nth-child(5) [data-testid="stMetric"] { border-top: 7px solid #D40000; }

/* ── 차트 컨테이너 ── */
[data-testid="stPlotlyChart"] {
    background: #fff;
    border: 2px solid #000;
    padding: 4px;
}

/* ── 탭 ── */
[data-baseweb="tab-list"] {
    border-bottom: 3px solid #000 !important;
    gap: 2px;
}
[data-baseweb="tab"] {
    background: #E8E4DC !important;
    border: 2px solid #000 !important;
    border-bottom: none !important;
    border-radius: 0 !important;
    font-weight: 700 !important;
    color: #222 !important;
    padding: 8px 18px !important;
}
[aria-selected="true"][data-baseweb="tab"] {
    background: #1E3FA4 !important;
    color: #fff !important;
}

/* ── 드릴다운 Expander ── */
[data-testid="stExpander"] {
    border: 3px solid #000 !important;
    border-left: 8px solid #D40000 !important;
    border-radius: 0 !important;
    background: #fff;
}
[data-testid="stExpander"] summary {
    font-weight: 700 !important;
    font-size: 1rem !important;
    background: #FFF9E6;
}

/* ── 데이터프레임 ── */
[data-testid="stDataFrame"] {
    border: 2px solid #000 !important;
}

/* ── 멀티셀렉트 태그 ── */
[data-baseweb="tag"] {
    background-color: #1E3FA4 !important;
    border-radius: 0 !important;
}

/* ── 라디오 선택됨 ── */
[data-testid="stRadio"] [aria-checked="true"] + div {
    color: #F5C900 !important;
    font-weight: 700 !important;
}

/* ── 상태 뱃지 텍스트 ── */
[data-testid="stMarkdownContainer"] p {
    font-size: 0.92rem;
}

/* ── 몬드리안 장식 블록 (헤더 우측) ── */
.mondrian-bar {
    display: flex;
    height: 8px;
    margin-bottom: 16px;
}
.mondrian-bar .r { flex: 3; background: #D40000; border-right: 3px solid #000; }
.mondrian-bar .b { flex: 2; background: #1E3FA4; border-right: 3px solid #000; }
.mondrian-bar .y { flex: 1; background: #F5C900; border-right: 3px solid #000; }
.mondrian-bar .w { flex: 4; background: #fff; }
</style>
""", unsafe_allow_html=True)

# ── 등급대 설정 ───────────────────────────────────────────
GRADE_ORDER = ['1등급대', '2등급대', '3등급대', '4등급대', '5등급대', '6등급대 이상', '미입력']
GRADE_COLORS = {
    '1등급대':    '#1A5276', '2등급대': '#2471A3', '3등급대': '#5DADE2',
    '4등급대':    '#A9CCE3', '5등급대': '#F5CBA7', '6등급대 이상': '#E59866',
    '미입력':     '#BDC3C7',
}

def grade_band(val):
    try:
        v = float(val)
        if v < 2.0: return '1등급대'
        if v < 3.0: return '2등급대'
        if v < 4.0: return '3등급대'
        if v < 5.0: return '4등급대'
        if v < 6.0: return '5등급대'
        return '6등급대 이상'
    except:
        return '미입력'

# 계열 → 3개 그룹 매핑
GYEOL_MAP = {
    '인문':   '인문', '인문사회': '인문', '교육': '인문', '공통': '인문',
    '자연':   '자연', '공학':    '자연', '간호보건': '자연',
    '예체능': '예체능',
}
GYEOL_ORDER = ['인문', '자연', '예체능']
GYEOL_COLORS = {'인문': '#5B8DB8', '자연': '#4CAF82', '예체능': '#E88C3A'}

# ── 데이터 로딩 ───────────────────────────────────────────
@st.cache_data
def load_data():
    wb = xlrd.open_workbook(
        'data.xls',
        encoding_override='cp949'
    )
    ws = wb.sheet_by_index(0)
    headers = []
    for c in range(ws.ncols):
        h = str(ws.cell_value(1, c)).replace('\n', ' ').strip()
        headers.append(h if h else f"Col_{c}")
    rows = []
    for r in range(2, ws.nrows):
        row = [ws.cell_value(r, c) for c in range(ws.ncols)]
        rows.append(row)
    df = pd.DataFrame(rows, columns=headers)
    df['전형 분류']  = df['전형 분류'].astype(str).str.strip()
    df['최종']      = df['최종'].astype(str).str.strip()
    df['전형방법']   = df['전형방법'].astype(str).str.replace('\r\n', ' / ').str.replace('\n', ' / ').str.strip()
    df['최저학력기준'] = df['최저학력기준'].astype(str).str.strip().replace('nan', '-')
    df = df[df['전형 분류'].str.strip() != '']
    df['국수영']    = pd.to_numeric(df['국수영'], errors='coerce')
    df['국어']     = pd.to_numeric(df['국어'],   errors='coerce')
    df['수학']     = pd.to_numeric(df['수학'],   errors='coerce')
    df['영어']     = pd.to_numeric(df['영어'],   errors='coerce')
    df['국영수평균'] = df[['국어', '영어', '수학']].mean(axis=1).round(3)
    df['국영수 등급대'] = df['국영수평균'].apply(grade_band)
    # 계열 그룹 컬럼 추가
    df['계열'] = df['계열'].astype(str).str.strip()
    df['계열구분'] = df['계열'].map(GYEOL_MAP).fillna('기타')
    return df

df      = load_data()
df_pass = df[df['최종'] == '합'].copy()
col_type  = '전형 분류'
col_grade = '국영수 등급대'
col_gyeol = '계열구분'
COLORS    = px.colors.qualitative.Set2

# 몬드리안 plotly 공통 레이아웃
PLOT_LAYOUT = dict(
    paper_bgcolor='#ffffff',
    plot_bgcolor='#F8F6F0',
    font=dict(family='Arial Black, Arial, sans-serif', color='#111'),
    title_font=dict(size=13, color='#333'),
    margin=dict(t=40, b=10, l=10, r=10),
    xaxis=dict(showgrid=False, linecolor='#000', linewidth=2, ticks='outside', tickcolor='#000'),
    yaxis=dict(showgrid=True, gridcolor='#E0E0E0', linecolor='#000', linewidth=2),
)

# 드릴다운에서 보여줄 기본 컬럼
DRILL_COLS = ['지역', '대학', '전형명', '모집단위', col_type, col_gyeol, col_grade,
              '국영수평균', '최종', '최저학력기준', '전형방법']

# ── 사이드바 ─────────────────────────────────────────────
with st.sidebar:
    st.title("🎓 필터")
    view = st.radio("데이터 범위", ["전체 지원자", "최종 합격자만"], index=0)
    st.markdown("---")

    all_types = sorted(df[col_type].unique().tolist())
    selected_types = st.multiselect("전형분류 (미선택 시 전체)", options=all_types, default=[])
    st.markdown("---")

    avail_grades = [g for g in GRADE_ORDER if g in df[col_grade].unique()]
    selected_grades = st.multiselect(
        "내신 국영수 등급대 (미선택 시 전체)", options=avail_grades, default=[],
        help="국어·영어·수학 평균 등급\n1등급대: 1.00~1.99 / 2등급대: 2.00~2.99 ..."
    )
    st.markdown("---")

    # 계열구분 필터 ★ 신규
    selected_gyeol = st.multiselect(
        "계열 구분 (미선택 시 전체)",
        options=GYEOL_ORDER,
        default=[],
        help="인문: 인문·인문사회·교육·공통\n자연: 자연·공학·간호보건\n예체능: 예체능"
    )
    st.markdown("---")

    if '지역' in df.columns:
        all_regions = sorted(r for r in df['지역'].astype(str).str.strip().unique() if r and r != 'nan')
        selected_regions = st.multiselect("지역 (미선택 시 전체)", options=all_regions, default=[])
    else:
        selected_regions = []

    st.markdown("---")
    st.caption("2026학년도 수시합격자\n지원결과 분석 대시보드")

# ── 필터 적용 ─────────────────────────────────────────────
data = df_pass.copy() if view == "최종 합격자만" else df.copy()
if selected_types:   data = data[data[col_type].isin(selected_types)]
if selected_grades:  data = data[data[col_grade].isin(selected_grades)]
if selected_gyeol:   data = data[data[col_gyeol].isin(selected_gyeol)]
if selected_regions and '지역' in data.columns:
    data = data[data['지역'].astype(str).str.strip().isin(selected_regions)]
total = len(data)

# ── 집계 ─────────────────────────────────────────────────
counts = data[col_type].value_counts().reset_index()
counts.columns = [col_type, '건수']
counts = counts.sort_values('건수', ascending=False).reset_index(drop=True)
counts['비율(%)'] = (counts['건수'] / total * 100).round(2) if total > 0 else 0

# ── 헤더 ─────────────────────────────────────────────────
st.title("🎓 2026학년도 수시 합격자 지원결과 분석")
st.markdown('<div class="mondrian-bar"><div class="r"></div><div class="b"></div><div class="y"></div><div class="w"></div></div>', unsafe_allow_html=True)
badge       = "✅ 최종 합격자만" if view == "최종 합격자만" else "📋 전체 지원자"
grade_badge = f" &nbsp;|&nbsp; 등급대: **{', '.join(selected_grades)}**" if selected_grades else ""
gyeol_badge = f" &nbsp;|&nbsp; 계열: **{', '.join(selected_gyeol)}**" if selected_gyeol else ""
st.markdown(f"**현재 보기:** {badge}{grade_badge}{gyeol_badge} &nbsp;|&nbsp; **총 {total:,}건**")
st.markdown("---")

# ── KPI 카드 ─────────────────────────────────────────────
k1, k2, k3, k4, k5 = st.columns(5)
total_all  = len(df)
total_pass = len(df_pass)
pass_rate  = round(total_pass / total_all * 100, 1) if total_all > 0 else 0
type_count = data[col_type].nunique()
has_grade  = data[data[col_grade] != '미입력']
avg_grade  = round(has_grade['국영수평균'].mean(), 3) if len(has_grade) > 0 else 0
k1.metric("전체 지원 건수",      f"{total_all:,}건")
k2.metric("최종 합격 건수",      f"{total_pass:,}건")
k3.metric("전체 합격률",         f"{pass_rate}%")
k4.metric("전형분류 수",         f"{type_count}개")
k5.metric("국영수 평균 (필터 기준)", f"{avg_grade}등급")
st.markdown("---")

# ── 드릴다운 헬퍼 ────────────────────────────────────────
def show_drilldown(subset: pd.DataFrame, title: str):
    """선택된 카테고리의 상세 데이터 표시"""
    cols = [c for c in DRILL_COLS if c in subset.columns]
    disp = subset[cols].sort_values('국영수평균', na_position='last').reset_index(drop=True)
    disp.index = disp.index + 1
    with st.expander(f"🔍 {title} — {len(disp):,}건 상세 데이터", expanded=True):
        # 미니 통계
        m1, m2, m3 = st.columns(3)
        pass_cnt = (disp['최종'] == '합').sum()
        m1.metric("건수", f"{len(disp):,}")
        m2.metric("합격", f"{pass_cnt}")
        m3.metric("합격률", f"{pass_cnt/len(disp)*100:.1f}%" if len(disp) > 0 else "-")
        st.dataframe(disp, width='stretch', height=380)

# ── [섹션1] 전형분류 차트 ────────────────────────────────
st.subheader("📊 전형분류별 분포")
col_l, col_r = st.columns(2)

with col_l:
    fig_bar = px.bar(counts, x=col_type, y='건수', text='건수',
                     color=col_type, color_discrete_sequence=COLORS)
    fig_bar.update_traces(textposition='outside')
    fig_bar.update_layout(**{**PLOT_LAYOUT,
                          'showlegend': False, 'xaxis_tickangle': -30, 'height': 380,
                          'title': "건수 막대 (클릭 → 상세보기)"})
    ev_bar = st.plotly_chart(fig_bar, on_select="rerun",
                             key="bar_type", width='stretch')

with col_r:
    fig_pie = px.pie(counts, names=col_type, values='건수',
                     color_discrete_sequence=COLORS, hole=0.35)
    fig_pie.update_traces(textinfo='label+percent', pull=[0.03]*len(counts))
    fig_pie.update_layout(paper_bgcolor='#ffffff', plot_bgcolor='#F8F6F0',
                          font=dict(family='Arial Black, Arial', color='#111'),
                          margin=dict(t=40, b=10), height=380,
                          title="비율 파이 (클릭 → 상세보기)")
    ev_pie = st.plotly_chart(fig_pie, on_select="rerun",
                             key="pie_type", width='stretch')

# 전형분류 드릴다운
clicked_type = None
if ev_bar and ev_bar.selection and ev_bar.selection.points:
    clicked_type = ev_bar.selection.points[0].get('x') or ev_bar.selection.points[0].get('label')
elif ev_pie and ev_pie.selection and ev_pie.selection.points:
    clicked_type = ev_pie.selection.points[0].get('label')

if clicked_type:
    subset = data[data[col_type] == clicked_type]
    show_drilldown(subset, f"전형분류: {clicked_type}")

st.markdown("---")

# ── [섹션2] 내신 등급대 분포 ─────────────────────────────
st.subheader("📚 내신 국영수 등급대 분포")

grade_cnt = data[col_grade].value_counts().reset_index()
grade_cnt.columns = [col_grade, '건수']
grade_cnt[col_grade] = pd.Categorical(grade_cnt[col_grade], categories=GRADE_ORDER, ordered=True)
grade_cnt = grade_cnt.sort_values(col_grade)
grade_cnt['비율(%)'] = (grade_cnt['건수'] / total * 100).round(1) if total > 0 else 0

g_l, g_r = st.columns(2)

with g_l:
    fig_gb = px.bar(grade_cnt, x=col_grade, y='건수', text='건수',
                    color=col_grade, color_discrete_map=GRADE_COLORS)
    fig_gb.update_traces(textposition='outside')
    fig_gb.update_layout(**{**PLOT_LAYOUT,
                         'showlegend': False, 'height': 360,
                         'title': "등급대별 건수 (클릭 → 상세보기)"})
    ev_gb = st.plotly_chart(fig_gb, on_select="rerun",
                            key="bar_grade", width='stretch')

with g_r:
    fig_gp = px.pie(grade_cnt, names=col_grade, values='건수',
                    color=col_grade, color_discrete_map=GRADE_COLORS, hole=0.35)
    fig_gp.update_traces(textinfo='label+percent')
    fig_gp.update_layout(paper_bgcolor='#ffffff', plot_bgcolor='#F8F6F0',
                         font=dict(family='Arial Black, Arial', color='#111'),
                         margin=dict(t=40, b=10), height=360,
                         title="등급대별 비율 (클릭 → 상세보기)")
    ev_gp = st.plotly_chart(fig_gp, on_select="rerun",
                            key="pie_grade", width='stretch')

# 등급대 드릴다운
clicked_grade = None
if ev_gb and ev_gb.selection and ev_gb.selection.points:
    clicked_grade = ev_gb.selection.points[0].get('x') or ev_gb.selection.points[0].get('label')
elif ev_gp and ev_gp.selection and ev_gp.selection.points:
    clicked_grade = ev_gp.selection.points[0].get('label')

if clicked_grade:
    subset = data[data[col_grade] == clicked_grade]
    show_drilldown(subset, f"등급대: {clicked_grade}")

# 계열구분 분포 차트
st.markdown("**계열 구분 분포 (클릭 → 상세보기)**")
gy_cnt = data[col_gyeol].value_counts().reset_index()
gy_cnt.columns = [col_gyeol, '건수']
gy_cnt[col_gyeol] = pd.Categorical(gy_cnt[col_gyeol], categories=GYEOL_ORDER + ['기타'], ordered=True)
gy_cnt = gy_cnt.sort_values(col_gyeol)

gy_l, gy_r = st.columns(2)
with gy_l:
    fig_gy = px.bar(gy_cnt, x=col_gyeol, y='건수', text='건수',
                    color=col_gyeol, color_discrete_map=GYEOL_COLORS)
    fig_gy.update_traces(textposition='outside')
    fig_gy.update_layout(**{**PLOT_LAYOUT, 'showlegend': False, 'height': 320})
    ev_gy = st.plotly_chart(fig_gy, on_select="rerun", key="bar_gyeol", width='stretch')

with gy_r:
    fig_gyp = px.pie(gy_cnt, names=col_gyeol, values='건수',
                     color=col_gyeol, color_discrete_map=GYEOL_COLORS, hole=0.35)
    fig_gyp.update_traces(textinfo='label+percent')
    fig_gyp.update_layout(paper_bgcolor='#ffffff', plot_bgcolor='#F8F6F0',
                          font=dict(family='Arial Black, Arial', color='#111'),
                          margin=dict(t=30, b=10), height=320)
    ev_gyp = st.plotly_chart(fig_gyp, on_select="rerun", key="pie_gyeol", width='stretch')

clicked_gyeol = None
if ev_gy and ev_gy.selection and ev_gy.selection.points:
    clicked_gyeol = ev_gy.selection.points[0].get('x') or ev_gy.selection.points[0].get('label')
elif ev_gyp and ev_gyp.selection and ev_gyp.selection.points:
    clicked_gyeol = ev_gyp.selection.points[0].get('label')
if clicked_gyeol:
    show_drilldown(data[data[col_gyeol] == clicked_gyeol], f"계열: {clicked_gyeol}")

# 전형분류 × 등급대 히트맵
st.markdown("**전형분류 × 등급대 히트맵 (클릭 → 상세보기)**")
heat_data   = data.groupby([col_type, col_grade]).size().reset_index(name='건수')
heat_pivot  = heat_data.pivot(index=col_type, columns=col_grade, values='건수').fillna(0)
heat_cols   = [c for c in GRADE_ORDER if c in heat_pivot.columns]
heat_pivot  = heat_pivot[heat_cols]

fig_heat = px.imshow(heat_pivot, text_auto=True, color_continuous_scale='Blues',
                     aspect='auto', labels=dict(x='등급대', y='전형분류', color='건수'))
fig_heat.update_layout(paper_bgcolor='#ffffff', plot_bgcolor='#F8F6F0',
                       font=dict(family='Arial Black, Arial', color='#111'),
                       height=300, margin=dict(t=20, b=10))
ev_heat = st.plotly_chart(fig_heat, on_select="rerun", key="heat", width='stretch')

if ev_heat and ev_heat.selection and ev_heat.selection.points:
    pt = ev_heat.selection.points[0]
    ht_type  = pt.get('y')
    ht_grade = pt.get('x')
    if ht_type and ht_grade:
        subset = data[(data[col_type] == ht_type) & (data[col_grade] == ht_grade)]
        show_drilldown(subset, f"{ht_type} × {ht_grade}")

st.markdown("---")

# ── [섹션3] 전체 vs 합격자 비교 ──────────────────────────
st.subheader("⚖️ 전체 vs 최종합격자 비교")

cnt_all  = df[col_type].value_counts().reset_index(); cnt_all.columns  = [col_type, '전체건수']
cnt_pass = df_pass[col_type].value_counts().reset_index(); cnt_pass.columns = [col_type, '합격건수']
compare  = cnt_all.merge(cnt_pass, on=col_type, how='outer').fillna(0)
compare['전체건수']  = compare['전체건수'].astype(int)
compare['합격건수']  = compare['합격건수'].astype(int)
compare['합격률(%)'] = (compare['합격건수'] / compare['전체건수'].replace(0,1) * 100).round(1)
compare = compare.sort_values('전체건수', ascending=False)

c_l, c_r = st.columns(2)

with c_l:
    fig_cmp = go.Figure()
    fig_cmp.add_trace(go.Bar(name='전체 지원', x=compare[col_type], y=compare['전체건수'],
                             marker_color='#AED6F1', text=compare['전체건수'], textposition='outside'))
    fig_cmp.add_trace(go.Bar(name='최종 합격', x=compare[col_type], y=compare['합격건수'],
                             marker_color='#2ECC71', text=compare['합격건수'], textposition='outside'))
    fig_cmp.update_layout(**{**PLOT_LAYOUT,
                          'barmode': 'group', 'xaxis_tickangle': -30, 'height': 400,
                          'legend': dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1)})
    st.plotly_chart(fig_cmp, width='stretch')

with c_r:
    fig_rate = px.bar(compare.sort_values('합격률(%)'), x='합격률(%)', y=col_type,
                      orientation='h', text='합격률(%)', color='합격률(%)',
                      color_continuous_scale='RdYlGn', range_color=[0,100])
    fig_rate.update_traces(texttemplate='%{text}%', textposition='outside')
    fig_rate.update_layout(**{**PLOT_LAYOUT,
                           'coloraxis_showscale': False, 'height': 400,
                           'yaxis': dict(showgrid=False, linecolor='#000', linewidth=2)})
    st.plotly_chart(fig_rate, width='stretch')

st.markdown("---")

# ── [섹션4] 등급대별 합격률 ──────────────────────────────
st.subheader("🎯 내신 등급대별 합격률")

grade_all  = df.groupby(col_grade).size().reset_index(name='전체')
grade_pass_cnt = df_pass.groupby(col_grade).size().reset_index(name='합격')
grade_rate = grade_all.merge(grade_pass_cnt, on=col_grade, how='left').fillna(0)
grade_rate['합격'] = grade_rate['합격'].astype(int)
grade_rate['합격률(%)'] = (grade_rate['합격'] / grade_rate['전체'] * 100).round(1)
grade_rate[col_grade] = pd.Categorical(grade_rate[col_grade], categories=GRADE_ORDER, ordered=True)
grade_rate = grade_rate.sort_values(col_grade)

fig_gr = go.Figure()
fig_gr.add_trace(go.Bar(name='전체', x=grade_rate[col_grade], y=grade_rate['전체'],
                        marker_color='#AED6F1', text=grade_rate['전체'], textposition='outside'))
fig_gr.add_trace(go.Bar(name='최종 합격', x=grade_rate[col_grade], y=grade_rate['합격'],
                        marker_color='#2ECC71', text=grade_rate['합격'], textposition='outside'))
fig_gr.add_trace(go.Scatter(name='합격률(%)', x=grade_rate[col_grade], y=grade_rate['합격률(%)'],
                            yaxis='y2', mode='lines+markers+text',
                            text=grade_rate['합격률(%)'].astype(str)+'%',
                            textposition='top center',
                            line=dict(color='#E74C3C', width=2), marker=dict(size=8)))
fig_gr.update_layout(**{**PLOT_LAYOUT,
                     'barmode': 'group', 'height': 420,
                     'yaxis': dict(title='건수', showgrid=True, gridcolor='#E0E0E0', linecolor='#000', linewidth=2),
                     'yaxis2': dict(title='합격률(%)', overlaying='y', side='right', range=[0,100]),
                     'legend': dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1)})
st.plotly_chart(fig_gr, width='stretch')

st.markdown("---")

# ── [섹션5] 하단 테이블 탭 ───────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs(["📋 전형분류 집계", "📚 등급대 집계", "⚖️ 전체 vs 합격 비교", "🗂️ 원본 데이터"])

with tab1:
    disp = counts.copy(); disp.index = disp.index + 1
    st.dataframe(disp, width='stretch', height=350)

with tab2:
    disp2 = grade_cnt[[col_grade, '건수', '비율(%)']].copy(); disp2.index = disp2.index + 1
    st.dataframe(disp2, width='stretch', height=300)

with tab3:
    cmp_disp = compare[[col_type,'전체건수','합격건수','합격률(%)']].sort_values('합격률(%)', ascending=False).reset_index(drop=True)
    cmp_disp.index = cmp_disp.index + 1
    st.dataframe(cmp_disp, width='stretch', height=350)

with tab4:
    st.caption(f"현재 필터 기준 {total:,}건 | 최저학력기준·전형방법 포함")
    show_cols = [col_type, col_gyeol, '계열', col_grade, '국영수평균', '최종', '지역', '대학',
                 '전형명', '모집단위', '모집 인원', '최저학력기준', '전형방법']
    show_cols = [c for c in show_cols if c in data.columns]
    st.dataframe(data[show_cols].reset_index(drop=True), width='stretch', height=450)
