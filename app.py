"""
안전팀 연간계획·실적 대시보드 (Streamlit)

실행 방법 (VS Code 터미널):
    pip install streamlit pandas plotly openpyxl
    streamlit run safety_dashboard.py

엑셀 파일:
    - 이 파일과 같은 폴더에 '연간계획조회*.xlsx' 를 두면 가장 최근 파일을 자동으로 불러옵니다.
      (파일명 속 날짜·시각 12자리, 예: 202610151320 을 '데이터 기준 시점'으로 사용)
    - 화면 왼쪽 사이드바에서 엑셀을 직접 올릴 수도 있습니다.
"""
from __future__ import annotations

import html
import io
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# =====================================================================
# 1. 설정값 (이름·색상은 여기서만 바꾸면 됩니다)
# =====================================================================
APP_TITLE = "안전팀 연간계획·실적 대시보드"
ORG_NAME = "대성에너지 안전팀"
FILE_PATTERN = "연간계획조회*.xlsx"
MONTHS = [f"{i}월" for i in range(1, 13)]

C_PLAN = "#898781"   # 계획(중립 회색)
C_DONE = "#2a78d6"   # 처리(파랑)
C_GOOD = "#0ca30c"   # 양호
C_WARN = "#fab219"   # 주의(당월 진행)
C_BAD = "#d03b3b"    # 지연
# 구분(정기검사·자율검사 …)별 고정 색상 - 필터를 바꿔도 같은 구분은 같은 색
CAT_COLORS = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#6250d6", "#008300", "#e34948"]

S_DONE, S_LATE, S_DUE, S_OK, S_NA = "✅ 완료", "🔴 지연", "🟡 당월 진행", "🔵 정상", "⚪ 계획 없음"
STATUS_ORDER = [S_LATE, S_DUE, S_OK, S_DONE, S_NA]

st.set_page_config(page_title=APP_TITLE, page_icon="🛡️", layout="wide")

st.markdown(
    """
<style>
.kpi {border:1px solid rgba(128,128,128,.28); border-top:4px solid var(--accent,#2a78d6);
      border-radius:10px; padding:14px 16px; background:rgba(128,128,128,.06); height:100%;}
.kpi .label {font-size:.82rem; opacity:.75;}
.kpi .value {font-size:1.85rem; font-weight:700; line-height:1.3;}
.kpi .sub {font-size:.78rem; opacity:.7; margin-top:2px;}
.brief {border-left:5px solid #2a78d6; background:rgba(42,120,214,.08);
        border-radius:6px; padding:12px 18px; line-height:1.85; margin:6px 0 14px 0;}
.stTabs [data-baseweb="tab"] {font-weight:600;}
</style>
""",
    unsafe_allow_html=True,
)


# =====================================================================
# 2. 엑셀 읽기 - 구조가 조금 바뀌어도 동작하도록 '제목 줄'을 찾아서 해석
# =====================================================================
def file_timestamp(name: str):
    """파일명 속 12자리(YYYYMMDDhhmm)를 날짜·시각으로 변환. 없으면 None."""
    m = re.search(r"(20\d{2})(\d{2})(\d{2})(\d{2})(\d{2})", name)
    if not m:
        return None
    try:
        return datetime(*[int(x) for x in m.groups()])
    except ValueError:
        return None


def find_files() -> list[Path]:
    """이 스크립트 폴더와 현재 폴더에서 연간계획조회 엑셀을 찾아 최신순으로 반환."""
    folders: list[Path] = []
    for f in (Path(__file__).resolve().parent, Path.cwd()):
        if f not in folders:
            folders.append(f)
    found: dict[str, Path] = {}
    for folder in folders:
        for p in folder.glob(FILE_PATTERN):
            if not p.name.startswith("~$"):  # 엑셀이 열려 있을 때 생기는 임시파일 제외
                found[str(p.resolve())] = p
    return sorted(
        found.values(),
        key=lambda p: (file_timestamp(p.name) or datetime.min, p.stat().st_mtime),
        reverse=True,
    )


def _txt(v):
    return None if pd.isna(v) else str(v).strip()


def _has(v, *words) -> bool:
    # pandas 버전에 따라 빈 칸이 None 또는 NaN(float)이 되므로 '문자열인지'를 먼저 확인
    return isinstance(v, str) and all(w in v for w in words)


def parse_plan(raw: pd.DataFrame) -> dict:
    """엑셀 원본(header=None)을 세부항목 / 소계 / 총계 / 원본보기로 분리."""
    # (1) '계획'·'처리' 글자가 있는 줄 = 제목 줄
    header_idx = None
    for i in range(len(raw)):
        vals = [str(v).strip() for v in raw.iloc[i].tolist()]
        if "계획" in vals and "처리" in vals:
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("'계획'과 '처리' 제목이 있는 줄을 찾지 못했습니다. 연간계획조회 엑셀이 맞는지 확인해 주세요.")

    header = ["" if pd.isna(v) else str(v).strip() for v in raw.iloc[header_idx].tolist()]
    body = raw.iloc[header_idx + 1:].reset_index(drop=True)
    pos: dict[str, int] = {}
    for j, name in enumerate(header):
        if j >= 2 and name and name not in pos:
            pos[name] = j

    # (2) 필요한 열만 뽑아 숫자로 변환 (없는 월 열은 0 처리)
    out = pd.DataFrame({"구분": body.iloc[:, 0], "세부항목": body.iloc[:, 1]})
    for c in ["계획", "처리", "처리율"] + MONTHS:
        out[c] = pd.to_numeric(body.iloc[:, pos[c]], errors="coerce") if c in pos else float("nan")
    out = out.dropna(how="all", subset=["구분", "세부항목", "계획", "처리"]).reset_index(drop=True)
    for c in ["계획", "처리"] + MONTHS:
        out[c] = out[c].fillna(0).round().astype(int)

    # (3) 행 종류 구분: 세부 / 소계('~ 합계') / 총계('총 합계')
    g = out["구분"].map(_txt)
    it = out["세부항목"].map(_txt)
    is_total = g.map(lambda v: _has(v, "총", "합계")) | it.map(lambda v: _has(v, "총", "합계"))
    is_sub = (g.map(lambda v: _has(v, "합계")) | it.map(lambda v: _has(v, "합계"))) & ~is_total
    is_detail = ~is_total & ~is_sub & it.notna()

    label = g.where(g.notna(), it)
    # 병합셀 때문에 비어 있는 '구분'을 위 행 값으로 채움 (합계 행 값은 이어받지 않음)
    group = g.where(~g.map(lambda v: _has(v, "합계"))).ffill()

    details = out[is_detail].copy()
    details["구분"] = group[is_detail].fillna("(미분류)")
    details["항목"] = details["구분"] + " · " + details["세부항목"]
    details = details.reset_index(drop=True)

    subtotals = out[is_sub].copy()
    subtotals["구분"] = label[is_sub].map(lambda v: v.replace("합계", "").strip())
    subtotals = subtotals.reset_index(drop=True)

    total = out[is_total].iloc[0] if is_total.any() else None

    kind = pd.Series("기타", index=out.index).mask(is_detail, "세부").mask(is_sub, "소계").mask(is_total, "총계")
    view = out.copy()
    view["구분"] = group.where(is_detail, label)
    view.insert(0, "행구분", kind)

    return {"details": details, "subtotals": subtotals, "total": total, "view": view}


@st.cache_data(show_spinner=False)
def load_plan(data: bytes) -> dict:
    return parse_plan(pd.read_excel(io.BytesIO(data), header=None))


# =====================================================================
# 3. 분석 계산 (기준월을 넣으면 지연·당월 목표를 계산)
# =====================================================================
def _judge(r) -> str:
    if r["계획"] <= 0:
        return S_NA
    if r["처리"] >= r["계획"]:
        return S_DONE
    if r["처리"] < r["전월말누계"]:
        return S_LATE
    if r["처리"] < r["당월말누계"]:
        return S_DUE
    return S_OK


def add_progress(details: pd.DataFrame, month: int) -> pd.DataFrame:
    """
    월별 열 = 그 달에 처리하기로 한 '계획'(합계가 연간 계획과 같음).
    처리는 연초부터의 누적 한 개 값이므로, 기준월까지의 누계 계획과 비교해 지연을 판단한다.
      전월말누계 : 기준월 이전 달까지 도래한 계획
      당월말누계 : 기준월 말까지 도래하는 계획
    """
    d = details.copy()
    d["전월말누계"] = d[MONTHS[: month - 1]].sum(axis=1)
    d["당월계획"] = d[MONTHS[month - 1]]
    d["당월말누계"] = d["전월말누계"] + d["당월계획"]
    d["처리율_건수"] = (d["처리"] / d["계획"].where(d["계획"] > 0)).fillna(0)
    d["잔여"] = (d["계획"] - d["처리"]).clip(lower=0)
    d["지연"] = (d["전월말누계"] - d["처리"]).clip(lower=0)
    d["당월말필요"] = (d["당월말누계"] - d["처리"]).clip(lower=0)
    d["상태"] = d.apply(_judge, axis=1)
    first = d[MONTHS].gt(0).idxmax(axis=1)
    d["첫예정월"] = first.where(d["계획"] > 0, "")
    return d


def summarize(d: pd.DataFrame) -> dict:
    plan, done = int(d["계획"].sum()), int(d["처리"].sum())
    return {
        "plan": plan,
        "done": done,
        "rate": done / plan * 100 if plan else 0.0,
        "remain": int(d["잔여"].sum()),
        "late": int(d["지연"].sum()),
        "late_items": int((d["지연"] > 0).sum()),
        "cur": int(d["당월계획"].sum()),
        "gap": int(d["당월말필요"].sum()),
        "cum_prev": int(d["전월말누계"].sum()),
        "cum_cur": int(d["당월말누계"].sum()),
    }


def by_category(d: pd.DataFrame) -> pd.DataFrame:
    g = d.groupby("구분", sort=False).agg(
        항목수=("세부항목", "count"),
        계획=("계획", "sum"),
        처리=("처리", "sum"),
        지연=("지연", "sum"),
        당월계획=("당월계획", "sum"),
        당월말필요=("당월말필요", "sum"),
    ).reset_index()
    g["처리율"] = (g["처리"] / g["계획"].where(g["계획"] > 0) * 100).fillna(0)
    return g


def to_pct(v) -> float:
    """엑셀 처리율은 0~1 비율. 이미 % 단위(1.5 초과)면 그대로 사용."""
    v = float(v)
    return v * 100 if v <= 1.5 else v


def make_briefing(d: pd.DataFrame, month: int, ts: datetime, s: dict) -> list[str]:
    """숫자로부터 자동 생성되는 '오늘의 한눈 요약'."""
    e = html.escape
    prev = month - 1
    lines = [
        f"<b>{ts:%Y.%m.%d} 기준</b> 연간 계획 <b>{s['plan']:,}건</b> 중 <b>{s['done']:,}건</b> 처리 "
        f"(처리율 <b>{s['rate']:.1f}%</b>, 잔여 {s['remain']:,}건)"
    ]

    if prev == 0:
        lines.append("1월은 전월 말까지 도래한 계획이 없어 지연 판정 대상이 없음")
    else:
        late = d[d["지연"] > 0].sort_values("지연", ascending=False)
        if late.empty:
            lines.append(f"{prev}월 말까지 도래한 계획은 <b>모두 처리</b>됨 (지연 0건) ✅")
        else:
            names = ", ".join(f"{e(n)} {int(v):,}건" for n, v in zip(late["항목"].head(3), late["지연"].head(3)))
            more = f" 외 {len(late) - 3}개" if len(late) > 3 else ""
            lines.append(
                f"{prev}월 말 기준 <b>지연 {s['late']:,}건 ({s['late_items']}개 항목)</b> 🔴 : {names}{more}"
            )

    diff = s["done"] - s["cum_cur"]
    lines.append(
        f"{month}월 계획 <b>{s['cur']:,}건</b> · {month}월 말 누계 목표 {s['cum_cur']:,}건 대비 처리 "
        f"{s['done']:,}건 (<b>{diff:+,}건</b>), 당월 말까지 추가 처리 필요 <b>{s['gap']:,}건</b>"
    )

    pending = d[(d["계획"] > 0) & (d["처리"] == 0) & (d["지연"] == 0)]
    if not pending.empty:
        txt = ", ".join(
            f"{e(n)}({m} 예정 {int(p):,}건)"
            for n, m, p in zip(pending["항목"].head(3), pending["첫예정월"].head(3), pending["계획"].head(3))
        )
        lines.append(f"아직 착수 전(예정대로 진행 중): {txt}")

    if month < 12:
        col = MONTHS[month]
        nxt = d[d[col] > 0].sort_values(col, ascending=False)
        if not nxt.empty:
            txt = ", ".join(f"{e(n)} {int(v):,}건" for n, v in zip(nxt["항목"].head(3), nxt[col].head(3)))
            lines.append(f"다음 달({col}) 예정 <b>{int(nxt[col].sum()):,}건</b> : {txt}")
    return lines


def run_checks(parsed: dict) -> pd.DataFrame:
    """엑셀 자체의 숫자가 서로 맞는지 점검 (신뢰도 확인용)."""
    det, sub, total = parsed["details"], parsed["subtotals"], parsed["total"]
    rows: list[dict] = []

    def add(name: str, ok: bool, text: str, warn_label: str = "⚠️ 확인 필요"):
        rows.append({"점검 항목": name, "결과": "✅ 정상" if ok else warn_label, "내용": text})

    plan, done = int(det["계획"].sum()), int(det["처리"].sum())
    if total is not None:
        tp, td = int(total["계획"]), int(total["처리"])
        add("세부항목 합계 = 엑셀 '총 합계'(계획·처리)", plan == tp and done == td,
            f"세부 합계 {plan:,}/{done:,} · 엑셀 총 합계 {tp:,}/{td:,}")
        if pd.notna(total["처리율"]):
            xl = to_pct(total["처리율"])
            real = done / plan * 100 if plan else 0
            simple = det["처리율"].mean() * 100 if det["처리율"].notna().any() else None
            why = " (세부항목 처리율의 단순평균과 일치)" if simple is not None and abs(simple - xl) < 0.1 else ""
            add("엑셀 '총 합계' 처리율 표기 방식", abs(xl - real) < 0.5,
                f"엑셀 표기 {xl:.1f}% / 건수 기준 {real:.1f}%{why} → 대시보드는 건수 기준을 사용", "ℹ️ 참고")
    if not sub.empty:
        grp = det.groupby("구분", sort=False)[["계획", "처리"]].sum().reset_index()
        m = sub.merge(grp, on="구분", how="left", suffixes=("_엑셀", "_세부"))
        bad = m[(m["계획_엑셀"] != m["계획_세부"]) | (m["처리_엑셀"] != m["처리_세부"])]
        add("구분별 소계 = 세부항목 합", bad.empty,
            "모두 일치" if bad.empty else "불일치: " + ", ".join(bad["구분"].astype(str)))
    mism = det[det[MONTHS].sum(axis=1) != det["계획"]]
    add("월별 계획 합 = 연간 계획", mism.empty,
        "모든 항목 일치" if mism.empty else "불일치: " + ", ".join(mism["항목"]))
    over = det[det["처리"] > det["계획"]]
    add("처리가 계획을 초과한 항목", over.empty,
        "없음" if over.empty else ", ".join(f"{a}({int(b)}/{int(c)})" for a, b, c in zip(over["항목"], over["처리"], over["계획"])))
    return pd.DataFrame(rows)


def to_excel_bytes(sheets: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for name, df in sheets.items():
            name = name[:31]
            df.to_excel(w, sheet_name=name, index=False)
            ws = w.sheets[name]
            ws.freeze_panes = "A2"
            for col_cells in ws.columns:
                width = max((len(str(c.value)) for c in col_cells if c.value is not None), default=0)
                ws.column_dimensions[col_cells[0].column_letter].width = min(max(width * 1.6 + 2, 8), 42)
    return buf.getvalue()


# =====================================================================
# 4. 화면 부품 (카드·차트)
# =====================================================================
def kpi(label: str, value: str, sub: str = "", accent: str = C_DONE) -> str:
    return (
        f'<div class="kpi" style="--accent:{accent}"><div class="label">{html.escape(label)}</div>'
        f'<div class="value">{value}</div><div class="sub">{sub}</div></div>'
    )


def show(fig: go.Figure, height: int = 380):
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=30, b=10),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="h", y=1.1, x=0),
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="rgba(128,128,128,.2)", zeroline=False)
    st.plotly_chart(fig, use_container_width=True)


def fig_cumulative(d: pd.DataFrame, month: int, s: dict) -> go.Figure:
    xs = list(range(1, 13))
    cum = d[MONTHS].sum().cumsum().tolist()
    fig = go.Figure()
    fig.add_vrect(x0=month - 0.5, x1=month + 0.5, fillcolor="rgba(250,178,25,.16)", line_width=0, layer="below")
    fig.add_trace(go.Scatter(
        x=xs, y=cum, mode="lines+markers", name="누적 계획(월말)",
        line=dict(color=C_PLAN, width=2), marker=dict(size=8),
        hovertemplate="%{x}월 말 누적 계획 %{y:,.0f}건<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        x=[month], y=[s["done"]], mode="markers+text", name="현재 누적 처리",
        marker=dict(size=15, color=C_DONE, line=dict(width=2, color="white")),
        text=[f"처리 {s['done']:,}건"], textposition="bottom right",
        hovertemplate="누적 처리 %{y:,.0f}건<extra></extra>",
    ))
    fig.update_xaxes(tickmode="array", tickvals=xs, ticktext=MONTHS)
    fig.update_yaxes(rangemode="tozero", tickformat=",")
    return fig


def fig_rate_by_category(g: pd.DataFrame, colors: dict) -> go.Figure:
    labels = [f"{c}  ({int(dn):,}/{int(p):,}건)" for c, dn, p in zip(g["구분"], g["처리"], g["계획"])]
    rates = g["처리율"].clip(upper=100).tolist()
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=[100] * len(g), y=labels, orientation="h", name="계획(100%)", showlegend=False,
        marker=dict(color="rgba(128,128,128,.22)"), hoverinfo="skip",
    ))
    fig.add_trace(go.Bar(
        x=rates, y=labels, orientation="h", name="처리율", showlegend=False,
        marker=dict(color=[colors.get(c, C_DONE) for c in g["구분"]]),
        text=[f"{r:.1f}%" for r in g["처리율"]], textposition="outside",
        hovertemplate="%{y}<br>처리율 %{x:.1f}%<extra></extra>",
    ))
    fig.update_layout(barmode="overlay", bargap=0.35)
    fig.update_xaxes(range=[0, 125], ticksuffix="%", tickvals=[0, 25, 50, 75, 100])
    fig.update_yaxes(autorange="reversed")
    return fig


def fig_monthly_stack(d: pd.DataFrame, month: int, colors: dict) -> go.Figure:
    xs = list(range(1, 13))
    fig = go.Figure()
    fig.add_vrect(x0=month - 0.5, x1=month + 0.5, fillcolor="rgba(250,178,25,.16)", line_width=0, layer="below")
    for cat, sub in d.groupby("구분", sort=False):
        fig.add_trace(go.Bar(
            x=xs, y=sub[MONTHS].sum().tolist(), name=cat, marker_color=colors.get(cat, C_DONE),
            hovertemplate=cat + " %{x}월 %{y:,.0f}건<extra></extra>",
        ))
    totals = d[MONTHS].sum().tolist()
    fig.add_trace(go.Scatter(
        x=xs, y=totals, mode="text", text=[f"{t:,}" if t else "" for t in totals],
        textposition="top center", showlegend=False, hoverinfo="skip",
    ))
    fig.update_layout(barmode="stack")
    fig.update_xaxes(tickmode="array", tickvals=xs, ticktext=MONTHS)
    fig.update_yaxes(tickformat=",", rangemode="tozero")
    return fig


def fig_gap(d: pd.DataFrame) -> go.Figure:
    t = d[d["당월말필요"] > 0].sort_values("당월말필요", ascending=True)
    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=t["항목"].tolist(), x=t["지연"].tolist(), orientation="h", name="🔴 지연(전월 말까지)",
        marker_color=C_BAD, hovertemplate="%{y}<br>지연 %{x:,.0f}건<extra></extra>",
    ))
    fig.add_trace(go.Bar(
        y=t["항목"].tolist(), x=(t["당월말필요"] - t["지연"]).tolist(), orientation="h", name="🟡 당월 도래분",
        marker_color=C_WARN, hovertemplate="%{y}<br>당월 도래 %{x:,.0f}건<extra></extra>",
    ))
    fig.update_layout(barmode="stack")
    fig.update_xaxes(tickformat=",")
    return fig


def fig_heatmap(d: pd.DataFrame, month: int) -> go.Figure:
    hm = d[MONTHS].astype(float)
    z = hm.pow(0.5).where(hm > 0)  # 물량 편차가 커서 제곱근 스케일로 색 농도 표현 (0은 빈칸)
    text = hm.apply(lambda col: col.map(lambda v: f"{v:,.0f}" if v > 0 else ""))
    fig = go.Figure(go.Heatmap(
        z=z.values.tolist(), x=list(range(1, 13)), y=d["항목"].tolist(),
        text=text.values.tolist(), texttemplate="%{text}", customdata=hm.values.tolist(),
        hovertemplate="%{y}<br>%{x}월 계획 %{customdata:,.0f}건<extra></extra>",
        colorscale=[[0, "#cde2fb"], [0.5, "#3987e5"], [1, "#0d366b"]],
        showscale=False, xgap=2, ygap=2, hoverongaps=False, zmin=0,
    ))
    fig.add_vrect(x0=month - 0.5, x1=month + 0.5, line_width=2, line_color=C_WARN, fillcolor="rgba(0,0,0,0)")
    fig.update_xaxes(tickmode="array", tickvals=list(range(1, 13)), ticktext=MONTHS, side="top")
    fig.update_yaxes(autorange="reversed", gridcolor="rgba(0,0,0,0)")
    return fig


def fig_month_items(m: pd.DataFrame, col: str, colors: dict) -> go.Figure:
    t = m.sort_values(col, ascending=True)
    fig = go.Figure(go.Bar(
        y=t["항목"].tolist(), x=t[col].tolist(), orientation="h",
        marker_color=[colors.get(c, C_DONE) for c in t["구분"]],
        text=[f"{v:,}" for v in t[col]], textposition="outside",
        hovertemplate="%{y}<br>%{x:,.0f}건<extra></extra>",
    ))
    fig.update_xaxes(tickformat=",")
    return fig


# =====================================================================
# 5. 데이터 불러오기 (사이드바)
# =====================================================================
source_name, data = None, None
with st.sidebar:
    st.header("⚙️ 설정")
    uploaded = st.file_uploader("엑셀 직접 올리기 (선택)", type=["xlsx"])
    files = find_files()
    if uploaded is not None:
        source_name, data = uploaded.name, uploaded.getvalue()
    elif files:
        names = [p.name for p in files]
        pick = st.selectbox("📄 자동 인식된 파일", names, index=0) if len(names) > 1 else names[0]
        if len(names) == 1:
            st.caption(f"📄 {pick}")
        source_name, data = pick, files[names.index(pick)].read_bytes()

if data is None:
    st.title(f"🛡️ {APP_TITLE}")
    st.info(
        f"'{FILE_PATTERN}' 엑셀 파일을 이 프로그램과 **같은 폴더**에 두거나, "
        "왼쪽 사이드바의 '엑셀 직접 올리기'로 올려 주세요."
    )
    st.stop()

try:
    parsed = load_plan(data)
except Exception as exc:  # 초보자가 원인을 알 수 있도록 화면에 그대로 안내
    st.error(f"엑셀을 해석하는 중 문제가 생겼습니다: {exc}")
    st.stop()

details_all: pd.DataFrame = parsed["details"]
if details_all.empty:
    st.error("세부항목 행을 찾지 못했습니다. 엑셀 형식을 확인해 주세요.")
    st.stop()

ts = file_timestamp(source_name) or datetime.now()
cats_all = list(dict.fromkeys(details_all["구분"]))
cat_color = {c: CAT_COLORS[i % len(CAT_COLORS)] for i, c in enumerate(cats_all)}

with st.sidebar:
    month = st.selectbox(
        "📅 기준월", list(range(1, 13)), index=ts.month - 1,
        format_func=lambda x: f"{x}월", key=f"month_{source_name}",
    )
    sel_cats = st.multiselect("🏷️ 구분 (전체 화면에 적용)", cats_all, default=cats_all, key=f"cats_{source_name}") or cats_all
    st.caption("기준월은 파일명의 날짜에서 자동 설정됩니다. 바꾸면 지연·당월 목표가 다시 계산됩니다.")
    if month != ts.month:
        st.warning(
            f"파일 기준월({ts.month}월)과 다릅니다. '현재 처리 실적'을 {month}월 목표와 비교하는 "
            "시뮬레이션 화면입니다."
        )

all_selected = set(sel_cats) == set(cats_all)
d = add_progress(details_all[details_all["구분"].isin(sel_cats)].reset_index(drop=True), month)
s = summarize(d)
cat_df = by_category(d)
prev = month - 1

# =====================================================================
# 6. 화면
# =====================================================================
st.title(f"🛡️ {APP_TITLE}")
st.caption(f"{ORG_NAME} · 데이터 기준 {ts:%Y-%m-%d %H:%M} · 파일: {source_name} · 기준월 {month}월")

st.markdown('<div class="brief">' + "<br>".join("• " + x for x in make_briefing(d, month, ts, s)) + "</div>",
            unsafe_allow_html=True)

# --- KPI 카드 ---
xl_note = "건수 기준 (처리 ÷ 계획)"
total = parsed["total"]
if all_selected and total is not None and pd.notna(total["처리율"]):
    xl = to_pct(total["처리율"])
    simple = details_all["처리율"].mean() * 100 if details_all["처리율"].notna().any() else None
    if abs(xl - s["rate"]) >= 0.5:
        how = " (항목별 처리율 단순평균)" if simple is not None and abs(simple - xl) < 0.1 else ""
        xl_note = f"엑셀 표기 {xl:.1f}%{how}와 다름"
    else:
        xl_note = "엑셀 표기와 일치"

c1, c2, c3, c4, c5 = st.columns(5)
c1.markdown(kpi("📊 연간 계획", f"{s['plan']:,}건", f"{len(d)}개 세부항목"), unsafe_allow_html=True)
c2.markdown(kpi("✅ 누적 처리", f"{s['done']:,}건", f"{ts:%m/%d} 기준"), unsafe_allow_html=True)
c3.markdown(kpi("📈 처리율", f"{s['rate']:.1f}%", html.escape(xl_note), C_GOOD if s["rate"] >= 95 else C_DONE),
            unsafe_allow_html=True)
c4.markdown(kpi("⏳ 잔여", f"{s['remain']:,}건", f"{month}월 말까지 필요 {s['gap']:,}건", C_WARN),
            unsafe_allow_html=True)
late_label = f"{prev}월 말 기준" if prev else "판정 대상 없음"
c5.markdown(
    kpi("🔴 지연", f"{s['late']:,}건" + (" ✅" if s["late"] == 0 else ""),
        f"{s['late_items']}개 항목 · {late_label}", C_BAD if s["late"] else C_GOOD),
    unsafe_allow_html=True,
)
st.write("")

tab1, tab2, tab3, tab4 = st.tabs(["📈 종합 현황", "🚨 이행·지연 관리", "🗓️ 월별 일정", "🗂️ 데이터·검증"])

# ---------------------------------------------------------------- 탭1
with tab1:
    left, right = st.columns([3, 2])
    with left:
        st.subheader("누적 계획 vs 현재 처리")
        show(fig_cumulative(d, month, s), 360)
        st.caption(f"회색 선: 월말까지 도래하는 누적 계획 · 파란 점: 현재 누적 처리 · 노란 띠: 기준월({month}월)")
    with right:
        st.subheader("구분별 처리율")
        show(fig_rate_by_category(cat_df, cat_color), 360)
    st.subheader("월별 계획 물량 (구분별)")
    show(fig_monthly_stack(d, month, cat_color), 380)

# ---------------------------------------------------------------- 탭2
with tab2:
    cnt = d["상태"].value_counts()
    cols = st.columns(4)
    for col, stt, color in zip(cols, [S_LATE, S_DUE, S_OK, S_DONE], [C_BAD, C_WARN, C_DONE, C_GOOD]):
        col.markdown(kpi(stt, f"{int(cnt.get(stt, 0))}개 항목", "", color), unsafe_allow_html=True)
    st.write("")

    with st.expander("판정 기준 보기"):
        st.markdown(
            f"""
- 엑셀의 월별 숫자는 **그 달에 처리하기로 한 계획**이며(합계 = 연간 계획), 처리는 **연초부터의 누적 한 값**입니다.
- 그래서 **기준월({month}월)** 에 맞춰 누계 계획과 처리를 비교합니다.
- ✅ **완료**: 처리 ≥ 연간 계획
- 🔴 **지연**: 처리 < {prev}월 말까지 도래한 누계 계획
- 🟡 **당월 진행**: 지연은 아니지만 처리 < {month}월 말 누계 계획
- 🔵 **정상**: {month}월 말 목표까지 충족(남은 물량은 이후 월 예정)
- 처리율 0%라도 예정월이 오지 않았으면 지연이 아니라 🔵 정상으로 표시됩니다.
"""
        )

    if d["당월말필요"].sum() > 0:
        st.subheader(f"{month}월 말까지 추가 처리가 필요한 항목")
        show(fig_gap(d), max(220, 60 + 48 * int((d["당월말필요"] > 0).sum())))
    else:
        st.success(f"{month}월 말까지의 목표를 모두 충족했습니다. 🎉")

    st.subheader("세부항목 이행 현황")
    f1, f2 = st.columns([3, 2])
    present = [x for x in STATUS_ORDER if x in set(d["상태"])]
    with f1:
        st_sel = st.multiselect("상태 필터", present, default=present)
    with f2:
        kw = st.text_input("항목 검색", "", key="item_search", placeholder="예: 밸브, 정압기")

    t = d.copy()
    t["_rank"] = t["상태"].map({x: i for i, x in enumerate(STATUS_ORDER)})
    t = t.sort_values(["_rank", "지연", "당월말필요"], ascending=[True, False, False], kind="stable")
    t = t[t["상태"].isin(st_sel or present)]
    if kw:
        t = t[t["항목"].str.contains(kw, case=False, regex=False)]

    COL_LATE = f"지연 ({prev}월 말 기준)" if prev else "지연"
    COL_GAP = f"{month}월 말까지 추가 필요"
    COL_CUR = f"{month}월 계획"
    table = pd.DataFrame({
        "상태": t["상태"], "구분": t["구분"], "세부항목": t["세부항목"],
        "연간 계획": t["계획"], "처리": t["처리"],
        "처리율(%)": (t["처리율_건수"] * 100).round(1),
        COL_LATE: t["지연"], COL_CUR: t["당월계획"], COL_GAP: t["당월말필요"], "연간 잔여": t["잔여"],
    })
    st.dataframe(
        table, use_container_width=True, hide_index=True, height=min(36 * (len(table) + 1) + 3, 620),
        column_config={
            "처리율(%)": st.column_config.ProgressColumn("처리율(%)", min_value=0, max_value=100, format="%.1f"),
            "연간 계획": st.column_config.NumberColumn(format="%d"),
            "처리": st.column_config.NumberColumn(format="%d"),
            COL_LATE: st.column_config.NumberColumn(format="%d"),
            COL_CUR: st.column_config.NumberColumn(format="%d"),
            COL_GAP: st.column_config.NumberColumn(format="%d"),
            "연간 잔여": st.column_config.NumberColumn(format="%d"),
        },
    )

# ---------------------------------------------------------------- 탭3
with tab3:
    st.subheader("월별 점검 일정 히트맵")
    show(fig_heatmap(d, month), 110 + 34 * len(d))
    st.caption("칸 안의 숫자 = 해당 월 계획 건수 · 색이 진할수록 물량이 많음(제곱근 스케일) · 노란 테두리 = 기준월")

    st.subheader("월 선택 상세")
    sel_m = st.selectbox("조회할 월", MONTHS, index=month - 1, key=f"detail_month_{source_name}")
    mi = MONTHS.index(sel_m) + 1
    m_total = int(d[sel_m].sum())
    m_cum = int(d[MONTHS[:mi]].sum().sum())
    k1, k2, k3 = st.columns(3)
    k1.markdown(kpi(f"{sel_m} 계획", f"{m_total:,}건", f"{int((d[sel_m] > 0).sum())}개 항목"), unsafe_allow_html=True)
    k2.markdown(kpi("연간 계획 대비", f"{m_total / s['plan'] * 100 if s['plan'] else 0:.1f}%", "해당 월 비중", C_PLAN),
                unsafe_allow_html=True)
    k3.markdown(kpi(f"{sel_m} 말 누계 계획", f"{m_cum:,}건", f"현재 처리 {s['done']:,}건", C_WARN), unsafe_allow_html=True)
    st.write("")

    md = d[d[sel_m] > 0].copy()
    if md.empty:
        st.info(f"{sel_m}에는 계획된 점검이 없습니다.")
    else:
        lc, rc = st.columns([3, 2])
        with lc:
            cum_col = f"{sel_m} 말 누계 계획"
            md[cum_col] = md[MONTHS[:mi]].sum(axis=1)
            tm = pd.DataFrame({
                "구분": md["구분"], "세부항목": md["세부항목"], f"{sel_m} 계획": md[sel_m],
                "월 내 비중(%)": (md[sel_m] / m_total * 100).round(1),
                "연간 계획 대비(%)": (md[sel_m] / md["계획"].where(md["계획"] > 0) * 100).round(1).fillna(0),
                cum_col: md[cum_col], "현재 처리": md["처리"], "처리−누계": md["처리"] - md[cum_col],
            }).sort_values(f"{sel_m} 계획", ascending=False)
            st.dataframe(tm, use_container_width=True, hide_index=True, height=min(36 * (len(tm) + 1) + 3, 460))
            st.caption("'처리−누계'가 음수면 해당 월 말까지의 누계 계획에 현재 처리가 못 미친다는 뜻입니다.")
        with rc:
            show(fig_month_items(md, sel_m, cat_color), max(240, 70 + 38 * len(md)))

# ---------------------------------------------------------------- 탭4
with tab4:
    st.subheader("엑셀 숫자 검증")
    checks = run_checks(parsed)
    st.dataframe(checks, use_container_width=True, hide_index=True)

    st.subheader("분석 결과 내려받기")
    out_detail = pd.DataFrame({
        "구분": d["구분"], "세부항목": d["세부항목"], "연간 계획": d["계획"], "처리": d["처리"],
        "처리율(%)": (d["처리율_건수"] * 100).round(1), "상태": d["상태"],
        f"지연({prev}월 말 기준)": d["지연"], f"{month}월 계획": d["당월계획"],
        f"{month}월 말까지 추가 필요": d["당월말필요"], "연간 잔여": d["잔여"],
    })
    out_month = d.groupby("구분", sort=False)[MONTHS].sum().reset_index()
    out_cat = cat_df.rename(columns={"처리율": "처리율(%)"}).round({"처리율(%)": 1})
    dl1, dl2, _ = st.columns([1, 1, 2])
    dl1.download_button(
        "📥 엑셀로 저장 (분석 결과)",
        data=to_excel_bytes({"세부현황": out_detail, "구분별요약": out_cat, "월별계획": out_month, "검증결과": checks}),
        file_name=f"안전팀_연간계획_분석_{ts:%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    dl2.download_button(
        "📥 CSV로 저장 (세부현황)",
        data=out_detail.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"안전팀_세부현황_{ts:%Y%m%d}.csv",
        mime="text/csv",
    )

    st.subheader("원본 데이터 탐색기")
    view: pd.DataFrame = parsed["view"]
    x1, x2, x3 = st.columns([2, 2, 2])
    with x1:
        kind_sel = st.radio("행 종류", ["전체", "세부만", "소계·총계만"], horizontal=True, key="kind_sel")
    with x2:
        cat_sel = st.multiselect("구분 필터", list(dict.fromkeys(view["구분"].dropna())))
    with x3:
        raw_kw = st.text_input("세부항목 검색", "", key="raw_kw")
    v = view
    if kind_sel == "세부만":
        v = v[v["행구분"] == "세부"]
    elif kind_sel == "소계·총계만":
        v = v[v["행구분"].isin(["소계", "총계"])]
    if cat_sel:
        v = v[v["구분"].isin(cat_sel)]
    if raw_kw:
        v = v[v["세부항목"].fillna("").astype(str).str.contains(raw_kw, case=False, regex=False)]
    st.dataframe(v, use_container_width=True, hide_index=True, height=480)

st.markdown("---")
st.markdown(
    f"<div style='text-align:center; opacity:.6; font-size:.85rem'>© {ts.year} {ORG_NAME} | Safety Management Dashboard</div>",
    unsafe_allow_html=True,
)