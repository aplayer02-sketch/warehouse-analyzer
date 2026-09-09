import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import json
import os
import urllib.request
import urllib.error

st.set_page_config(page_title="库房数据分析与规划工具", page_icon="📦", layout="wide")
st.title("📦 库房出入库数据分析与规划工具")

# ── 加载配置文件 ──
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")


def load_config():
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"入库": {}, "出库": {}, "主数据": {}, "业务规则": {}}


CFG = load_config()
RULE = CFG.get("业务规则", {})
PALLET_CAPACITY = RULE.get("一托盘装多少箱", 36)
PALLET_THRESHOLD = RULE.get("收货超过多少箱需用托盘", 12)
RACK_LEVELS = RULE.get("货架层数", 3)
BUFFER_RATE = RULE.get("库容缓冲系数", 0.25)
IN_EFF = RULE.get("收货每人每小时处理箱数", 15)
OUT_EFF = RULE.get("发货每人每小时处理箱数", 12)

FIELD_MAP_IN = CFG.get("入库", {})
FIELD_MAP_OUT = CFG.get("出库", {})
FIELD_MAP_MAT = CFG.get("主数据", {})
AI_CFG = CFG.get("大模型", {})
AI_ENABLED = bool(AI_CFG.get("启用", True))
AI_MODEL = AI_CFG.get("模型", "gpt-4o-mini")
AI_URL = AI_CFG.get("接口地址", "https://api.openai.com/v1/chat/completions")
AI_TIMEOUT = int(AI_CFG.get("超时秒数", 30))
AI_KEY = os.environ.get("OPENAI_API_KEY") or AI_CFG.get("密钥", "")

if "version" not in st.session_state:
    st.session_state.version = 0
v = st.session_state.version

# ── 上传区 ──
st.markdown("---")
c1, c2, c3 = st.columns(3)
with c1:
    fin = st.file_uploader("📥 上传入库文件（Excel）", type=["xlsx", "xls"], key=f"in_f_{v}")
with c2:
    fout = st.file_uploader("📤 上传出库文件（Excel）", type=["xlsx", "xls"], key=f"out_f_{v}")
with c3:
    fmat = st.file_uploader("🧰 物料主数据（可选）", type=["xlsx", "xls"], key=f"mat_f_{v}")
st.markdown("---")

if fin is not None or fout is not None:
    with st.sidebar:
        if st.button("🗑️ 清除所有数据并重新开始"):
            st.session_state.version += 1
            st.rerun()


def find_col_by_map(df, field_map, std_name, candidates):
    """优先用配置文件映射找列，找不到再用关键词猜测"""
    mapped = field_map.get(std_name)
    if mapped:
        if mapped in df.columns:
            return mapped
        mapped_norm = mapped.replace(" ", "").replace("\n", "").replace("　", "")
        for c in df.columns:
            if mapped_norm == c.replace(" ", "").replace("\n", "").replace("　", ""):
                return c
        for c in df.columns:
            if mapped.lower() in str(c).lower():
                return c
    for c in df.columns:
        cl = str(c).lower()
        for cand in candidates:
            if cand in cl:
                return c
    return None


def read_any(file, sheet_name=None, nrows=None):
    """读取 Excel，统一清理列名去 BOM/空格"""
    if sheet_name is None:
        df = pd.read_excel(file, nrows=nrows)
    else:
        df = pd.read_excel(file, sheet_name=sheet_name, nrows=nrows)
    if df is not None:
        df.columns = [str(c).strip().replace("﻿", "") for c in df.columns]
    return df


def load_sheet(file, sheet_name, time_col, qty_col, field_map=None, time_std="入库时间", qty_std="入库箱数"):
    df = read_any(file, sheet_name)
    if df is None or df.empty:
        return None, None, None
    field_map = field_map or {}
    tc = time_col if time_col and time_col != "（自动）" else find_col_by_map(
        df, field_map, time_std, ["时间", "time", "日期", "date"])
    qc = qty_col if qty_col and qty_col != "（自动）" else find_col_by_map(
        df, field_map, qty_std, ["箱数", "箱", "数量", "qty", "quantity", "数"])
    if tc is None:
        tc = df.columns[0]
    if qc is None:
        qc = df.columns[1] if len(df.columns) > 1 else df.columns[0]
    result = df.copy()
    result["时间"] = pd.to_datetime(result[tc], errors="coerce")
    result["数量"] = pd.to_numeric(
        result[qc].astype(str).str.replace(",", "").str.strip(),
        errors="coerce"
    ).fillna(0)
    result = result.dropna(subset=["时间"]).sort_values("时间").reset_index(drop=True)
    return result, tc, qc


def sheet_list(file):
    return pd.ExcelFile(file).sheet_names


def compute_pallets(箱数):
    import math
    if 箱数 <= PALLET_THRESHOLD:
        return 0.0
    return math.ceil(箱数 / PALLET_CAPACITY)


def call_llm(prompt, model, url, api_key, timeout):
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": "你是供应链仓储数据分析师，请用中文、简洁、分点回答，不编造数据。"},
            {"role": "user", "content": prompt},
        ],
        "temperature": 0.3,
    }
    req = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"HTTP {exc.code} {exc.reason}：{detail}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"网络连接失败：{exc.reason}") from exc
    return data["choices"][0]["message"]["content"].strip()


def build_rule_summary(ctx):
    lines = [
        "**规则版解读（未调用大模型或调用失败）**",
        f"- 分析区间共 {ctx['days']} 天，入库 {ctx['total_in']:,} 箱，出库 {ctx['total_out']:,} 箱。",
    ]
    if ctx.get("avg_d") is not None:
        lines.append(f"- 平均在库约 {ctx['avg_d']:.1f} 天；日均入库约 {ctx['daily_pallet']:.1f} 托。")
        lines.append(f"- 建议储位约 {ctx['recommend_pallet']:,.0f} 托（含 {int(BUFFER_RATE * 100)}% 缓冲），按 {RACK_LEVELS} 层货架约每层 {ctx['per_level']:,.0f} 个托盘位。")
    lines.append(f"- 收货区建议 {ctx['in_min']} ~ {ctx['in_max']} 人；发货区建议 {ctx['out_min']} ~ {ctx['out_max']} 人。")
    lines.append("- 以上为规则计算结果，请结合业务波动进一步校验。")
    return "\n".join(lines)


# ── 物料主数据加载 ──
mat_df = None
mat_map = {}
if fmat is not None:
    mat_sheets = sheet_list(fmat)
    mat_sheet = next((s for s in mat_sheets if any(k in s.lower() for k in ["主数据", "master", "物料"])), mat_sheets[0])
    mat_df = read_any(fmat, mat_sheet)
    if mat_df is not None and not mat_df.empty:
        def mat_find(std_name, candidates):
            return find_col_by_map(mat_df, FIELD_MAP_MAT, std_name, candidates)
        code_col = mat_find("零件图号", ["图号", "零件", "物料", "编码", "code", "sku"])
        qty_per_box_col = mat_find("每箱件数", ["每箱件数", "箱件数", "件/箱", "件数", "箱量"])
        if code_col and qty_per_box_col:
            mat_df = mat_df[[code_col, qty_per_box_col]].dropna()
            mat_df.columns = ["零件图号", "每箱件数"]
            mat_df["零件图号"] = mat_df["零件图号"].astype(str).str.strip()
            mat_map = dict(zip(mat_df["零件图号"], pd.to_numeric(mat_df["每箱件数"], errors="coerce")))
            st.sidebar.success(f"🧰 主数据：{len(mat_map)} 个零件")


# ── 数据变量初始化 ──
df_in = df_out = None
has_in = has_out = False
has_inp = fin is not None
has_outp = fout is not None

# ── 入库加载 ──
if has_inp:
    in_sh = sheet_list(fin)
    auto_in = next((s for s in in_sh if any(k in s.lower() for k in ["入库", "in"])), in_sh[0])
    with st.sidebar:
        st.markdown("### 📥 入库")
        s_in = st.selectbox("Sheet", in_sh, index=in_sh.index(auto_in), key=f"s_in_{v}")
        prev_in = read_any(fin, s_in, nrows=1)
        cols_in = ["（自动）"] + list(prev_in.columns)
        t_in = st.selectbox("时间列", cols_in, key=f"t_in_{v}")
        q_in = st.selectbox("数量列", cols_in, key=f"q_in_{v}")
    df_in, _, _ = load_sheet(fin, s_in, t_in, q_in, field_map=FIELD_MAP_IN, time_std="入库时间", qty_std="入库箱数")
    has_in = df_in is not None and len(df_in) > 0
    if has_in:
        pallet_col = find_col_by_map(df_in, FIELD_MAP_IN, "托盘数", ["托盘", "托盘数", "pallet"])
        if pallet_col:
            df_in["托盘数"] = pd.to_numeric(df_in[pallet_col], errors="coerce").fillna(0)
        else:
            df_in["托盘数"] = df_in["数量"].apply(compute_pallets)
        part_col = find_col_by_map(df_in, FIELD_MAP_IN, "零件图号", ["图号", "零件", "物料", "code", "sku"])
        if part_col:
            df_in["零件图号"] = df_in[part_col].astype(str).str.strip()
        if mat_map and part_col:
            df_in["每箱件数"] = df_in["零件图号"].map(mat_map)
            df_in["件数"] = df_in["数量"] * df_in["每箱件数"].fillna(0)
else:
    has_in = False

# ── 出库加载 ──
if has_outp:
    out_sh = sheet_list(fout)
    auto_out = next((s for s in out_sh if any(k in s.lower() for k in ["出库", "out"])), out_sh[0])
    with st.sidebar:
        st.markdown("### 📤 出库")
        s_out = st.selectbox("Sheet", out_sh, index=out_sh.index(auto_out), key=f"s_out_{v}")
        prev_out = read_any(fout, s_out, nrows=1)
        cols_out = ["（自动）"] + list(prev_out.columns)
        t_out = st.selectbox("时间列", cols_out, key=f"t_out_{v}")
        q_out = st.selectbox("数量列", cols_out, key=f"q_out_{v}")
    df_out, t_out_used, q_out_used = load_sheet(fout, s_out, t_out, q_out, field_map=FIELD_MAP_OUT, time_std="出库时间", qty_std="出库数量")
    has_out = df_out is not None and len(df_out) > 0
    if has_out:
        st.sidebar.info(f"📤 出库时间列：{t_out_used} | 数量列：{q_out_used} | {len(df_out):,} 条")
    part_col_o = find_col_by_map(df_out, FIELD_MAP_OUT, "零件图号", ["图号", "零件", "物料", "code", "sku"])
    if part_col_o:
        df_out["零件图号"] = df_out[part_col_o].astype(str).str.strip()
    if mat_map and part_col_o:
        df_out["每箱件数"] = df_out["零件图号"].map(mat_map)
        df_out["件数"] = df_out["数量"] * df_out["每箱件数"].fillna(0)
else:
    has_out = False

# ── 无数据提示 ──
if not has_in and not has_out:
    st.info("### 👆 请在左侧上传数据文件开始分析")
    st.markdown("支持：\n- 上传**两个独立文件**（入库 + 出库）\n- 上传**一个文件**（内含入库和出库两个 Sheet）\n- 仅上传**入库**或**出库**单独分析")
    st.stop()

# ── 侧边栏汇总 ──
with st.sidebar:
    st.markdown("---")
    if has_in:
        st.success(f"✅ 入库：{len(df_in):,} 条")
    if has_out:
        st.success(f"✅ 出库：{len(df_out):,} 条")

total_in = 0
total_out = 0

# ── 时间筛选 ──
all_times = []
if has_in: all_times.extend(df_in["时间"].tolist())
if has_out: all_times.extend(df_out["时间"].tolist())

if all_times:
    min_d = min(all_times).date()
    max_d = max(all_times).date()
    dr = st.sidebar.date_input("📅 时间范围", [min_d, max_d], min_value=min_d, max_value=max_d, key=f"dr_{v}")
    if len(dr) == 2:
        sd, ed = dr[0], dr[1]
        if has_in:
            df_in = df_in[(df_in["时间"].dt.date >= sd) & (df_in["时间"].dt.date <= ed)]
        if has_out:
            df_out = df_out[(df_out["时间"].dt.date >= sd) & (df_out["时间"].dt.date <= ed)]

has_in = has_in and (df_in is not None and len(df_in) > 0)
has_out = has_out and (df_out is not None and len(df_out) > 0)

# ── 按物件筛选（多选） ──
item_col_in = find_col_by_map(df_in, FIELD_MAP_IN, "物件", ["物件", "品名", "名称", "item", "name"]) if has_in else None
item_col_out = find_col_by_map(df_out, FIELD_MAP_OUT, "物件", ["物件", "品名", "名称", "item", "name"]) if has_out else None
if item_col_in is not None or item_col_out is not None:
    item_values = set()
    if item_col_in is not None:
        item_values.update(df_in[item_col_in].astype(str).dropna().unique())
    if item_col_out is not None:
        item_values.update(df_out[item_col_out].astype(str).dropna().unique())
    if item_values:
        item_list = sorted(item_values)
        selected_items = st.sidebar.multiselect("🔍 按物件筛选（可多选）", item_list, key=f"item_{v}")
        if selected_items:
            if item_col_in is not None:
                df_in = df_in[df_in[item_col_in].astype(str).isin(selected_items)]
            if item_col_out is not None:
                df_out = df_out[df_out[item_col_out].astype(str).isin(selected_items)]
    has_in = has_in and (df_in is not None and len(df_in) > 0)
    has_out = has_out and (df_out is not None and len(df_out) > 0)

# ── 按单据种类筛选（多选） ──
doc_col_out = find_col_by_map(df_out, FIELD_MAP_OUT, "单据种类", ["单据种类", "单据类型", "单据"]) if has_out else None
doc_col_in = find_col_by_map(df_in, FIELD_MAP_IN, "单据种类", ["单据种类", "单据类型", "单据"]) if has_in else None
if doc_col_out is not None or doc_col_in is not None:
    doc_values = set()
    if doc_col_out is not None:
        doc_values.update(df_out[doc_col_out].astype(str).dropna().unique())
    if doc_col_in is not None:
        doc_values.update(df_in[doc_col_in].astype(str).dropna().unique())
    if doc_values:
        doc_list = sorted(doc_values)
        selected_docs = st.sidebar.multiselect("📋 按单据种类筛选（可多选）", doc_list, key=f"doc_{v}")
        if selected_docs:
            if doc_col_out is not None:
                df_out = df_out[df_out[doc_col_out].astype(str).isin(selected_docs)]
            if doc_col_in is not None:
                df_in = df_in[df_in[doc_col_in].astype(str).isin(selected_docs)]
    has_in = has_in and (df_in is not None and len(df_in) > 0)
    has_out = has_out and (df_out is not None and len(df_out) > 0)

# ── 筛选后的总量 ──
total_in = int(df_in["数量"].sum()) if has_in else 0
total_out = int(df_out["数量"].sum()) if has_out else 0

days = (max_d - min_d).days + 1 if all_times else 0

total_pallets_in = int(df_in["托盘数"].sum()) if has_in and "托盘数" in df_in.columns else 0

# ── 每日聚合 ──
daily_stats = pd.DataFrame({"date": pd.date_range(min_d, max_d)}).set_index("date")
if has_in:
    daily_in_sum = df_in.set_index("时间").resample("D")["数量"].sum()
    daily_in_cnt = df_in.set_index("时间").resample("D").size()
    daily_stats["入库量"] = daily_in_sum
    daily_stats["入库笔数"] = daily_in_cnt
if has_out:
    daily_out_sum = df_out.set_index("时间").resample("D")["数量"].sum()
    daily_out_cnt = df_out.set_index("时间").resample("D").size()
    daily_stats["出库量"] = daily_out_sum
    daily_stats["出库笔数"] = daily_out_cnt
if has_in and "托盘数" in df_in.columns:
    daily_stats["托盘量"] = df_in.set_index("时间").resample("D")["托盘数"].sum()
daily_stats = daily_stats.fillna(0)

# ═══════════════════════════════════════════
# 概览
# ═══════════════════════════════════════════
st.subheader("📊 数据概览")
c1, c2, c3, c4 = st.columns(4)
c1.metric("入库总数量", f"{total_in:,} 箱")
c2.metric("出库总数量", f"{total_out:,} 箱" if has_out else "-")
c3.metric("入库笔数", f"{len(df_in):,}" if has_in else "0")
c4.metric("出库笔数", f"{len(df_out):,}" if has_out else "0")

if total_pallets_in > 0:
    st.markdown("#### 🧱 托盘维度")
    p1, p2, p3 = st.columns(3)
    p1.metric("入库托盘总量", f"{total_pallets_in:,} 托")
    p2.metric("日均托盘量", f"{total_pallets_in/days:,.1f} 托" if days else "-")
    p3.metric("最大单日托盘", f"{df_in.set_index('时间').resample('D')['托盘数'].sum().max():,.0f} 托" if has_in else "-")

st.markdown("---")
c1, c2, c3 = st.columns(3)
c1.metric("分析天数", f"{days} 天")
c2.metric("日均入库", f"{total_in/days:,.0f} 箱" if days else "0")
c3.metric("日均出库", f"{total_out/days:,.0f} 箱" if days and has_out else "-")

# ── 单日极值统计 ──
st.markdown("#### 📈 单日极值统计")
if has_in:
    in_daily = df_in.groupby(df_in["时间"].dt.date)["数量"].sum()
    peak_in_val = in_daily.max()
    peak_in_row = in_daily.idxmax()
    nonzero_in = in_daily[in_daily > 0]
    if len(nonzero_in) > 0:
        min_in_val = nonzero_in.min()
        min_in_row = nonzero_in.idxmin()
    else:
        min_in_val, min_in_row = peak_in_val, peak_in_row
    e1, e2, e3, e4 = st.columns(4)
    e1.metric("单日入库峰值", f"{peak_in_val:,.0f} 箱")
    e2.metric("峰值日期", str(peak_in_row))
    e3.metric("单日入库谷值", f"{min_in_val:,.0f} 箱")
    e4.metric("谷值日期", str(min_in_row))
else:
    st.info("无入库数据")

if has_out:
    out_daily = df_out.groupby(df_out["时间"].dt.date)["数量"].sum()
    peak_out_val = out_daily.max()
    peak_out_row = out_daily.idxmax()
    nonzero_out = out_daily[out_daily > 0]
    if len(nonzero_out) > 0:
        min_out_val = nonzero_out.min()
        min_out_row = nonzero_out.idxmin()
    else:
        min_out_val, min_out_row = peak_out_val, peak_out_row
    e5, e6, e7, e8 = st.columns(4)
    e5.metric("单日出库峰值", f"{peak_out_val:,.0f} 箱")
    e6.metric("峰值日期", str(peak_out_row))
    e7.metric("单日出库谷值", f"{min_out_val:,.0f} 箱")
    e8.metric("谷值日期", str(min_out_row))
    with st.expander("📋 每日出库箱数明细（按量排序）"):
        detail = out_daily.sort_values(ascending=False).reset_index()
        detail.columns = ["日期", "出库箱数"]
        detail["出库箱数"] = detail["出库箱数"].map(lambda x: f"{x:,.0f}")
        st.dataframe(detail, use_container_width=True)

# ── 每日趋势 ──
st.subheader("每日趋势")
plot_daily = daily_stats.reset_index().rename(columns={"入库量": "入库", "出库量": "出库", "托盘量": "入库托盘"})
fig = go.Figure()
if has_in:
    fig.add_trace(go.Scatter(x=plot_daily["date"], y=plot_daily["入库"], mode="lines+markers", name="入库(箱)", line=dict(color="#2E86AB", width=2)))
if has_in and "入库托盘" in plot_daily.columns:
    fig.add_trace(go.Bar(x=plot_daily["date"], y=plot_daily["入库托盘"], name="入库托盘", marker_color="rgba(44,111,179,0.3)", yaxis="y2"))
if has_out:
    fig.add_trace(go.Scatter(x=plot_daily["date"], y=plot_daily["出库"], mode="lines+markers", name="出库(箱)", line=dict(color="#A23B72", width=2)))
fig.update_layout(height=400, hovermode="x unified", yaxis=dict(title="箱数"), yaxis2=dict(title="托盘数", overlaying="y", side="right", showgrid=False))
st.plotly_chart(fig, use_container_width=True)

# ═══════════════════════════════════════════
# 频次分析
# ═══════════════════════════════════════════
st.subheader("⏱ 频次分析")
if has_in:
    st.markdown("#### 入库")
    df_in["小时"] = df_in["时间"].dt.hour
    hi = df_in.groupby("小时").size().reset_index(name="笔数")
    fig = px.bar(hi, x="小时", y="笔数", title="各小时入库频次", color="笔数", color_continuous_scale="Blues")
    fig.update_xaxes(dtick=1)
    st.plotly_chart(fig, use_container_width=True)

if has_out:
    st.markdown("#### 出库")
    df_out["小时"] = df_out["时间"].dt.hour
    ho = df_out.groupby("小时").size().reset_index(name="笔数")
    fig = px.bar(ho, x="小时", y="笔数", title="各小时出库频次", color="笔数", color_continuous_scale="Reds")
    fig.update_xaxes(dtick=1)
    st.plotly_chart(fig, use_container_width=True)

# ═══════════════════════════════════════════
# 批次分布分析
# ═══════════════════════════════════════════
out_order_col = find_col_by_map(df_out, FIELD_MAP_OUT, "生产订单号", ["生产订单", "订单", "order"]) if has_out else None
out_part_col = find_col_by_map(df_out, FIELD_MAP_OUT, "零件图号", ["图号", "料号", "零件", "物料", "code", "sku"]) if has_out else None
out_batch_col = find_col_by_map(df_out, FIELD_MAP_OUT, "批次号", ["批次", "batch", "lot"]) if has_out else None
has_batch_analysis = has_out and all([out_order_col, out_part_col, out_batch_col])

if has_batch_analysis:
    st.subheader("🧬 批次分布分析（生产订单 × 料号 → 批次数）")
    batch_raw = df_out[[out_order_col, out_part_col, out_batch_col]].dropna()
    batch_raw[out_order_col] = batch_raw[out_order_col].astype(str)
    batch_raw[out_part_col] = batch_raw[out_part_col].astype(str)
    batch_raw[out_batch_col] = batch_raw[out_batch_col].astype(str)
    grp = batch_raw.groupby([out_order_col, out_part_col], dropna=False).agg(
        批次数量=pd.NamedAgg(column=out_batch_col, aggfunc="nunique"),
        出库笔数=pd.NamedAgg(column=out_batch_col, aggfunc="count")
    ).reset_index()
    grp.columns = ["生产订单", "料号", "批次数量", "出库笔数"]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("订单×料号组合数", f"{len(grp):,}")
    c2.metric("平均批次数量", f"{grp['批次数量'].mean():.2f}")
    c3.metric("最多批次", f"{grp['批次数量'].max()}")
    c4.metric("多批次组合占比", f"{(grp['批次数量']>1).mean()*100:.1f}%")

    st.markdown("#### 批次数量分布")
    hist = grp["批次数量"].value_counts().sort_index().reset_index()
    hist.columns = ["批次数量", "组合数"]
    fig = px.bar(hist, x="批次数量", y="组合数", title="批次数量分布", color="组合数", color_continuous_scale="Viridis")
    fig.update_xaxes(dtick=1)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### 每日多批次发生情况")
    if "时间" in df_out.columns:
        daily_key = df_out[[out_order_col, out_part_col, out_batch_col, "时间"]].copy()
        daily_key[out_order_col] = daily_key[out_order_col].astype(str)
        daily_key[out_part_col] = daily_key[out_part_col].astype(str)
        daily_key[out_batch_col] = daily_key[out_batch_col].astype(str)
        daily_key["日期"] = pd.to_datetime(daily_key["时间"], errors="coerce").dt.date
        daily_key = daily_key.dropna(subset=["日期"])
        daily_grp = daily_key.groupby(["日期", out_order_col, out_part_col]).agg(
            批次数量=pd.NamedAgg(column=out_batch_col, aggfunc="nunique")
        ).reset_index()
        daily_batch = daily_grp.groupby("日期").agg(
            全部组合数=("批次数量", "count"),
            多批次组合数=("批次数量", lambda s: (s > 1).sum())
        ).reset_index()
        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(x=daily_batch["日期"], y=daily_batch["全部组合数"], mode="lines+markers", name="全部组合数", line=dict(color="#2E86AB", width=2)))
        fig2.add_trace(go.Scatter(x=daily_batch["日期"], y=daily_batch["多批次组合数"], mode="lines+markers", name="多批次组合数", line=dict(color="#E74C3C", width=2)))
        fig2.update_layout(height=380, hovermode="x unified", xaxis_title="日期", yaxis_title="组合数")
        st.plotly_chart(fig2, use_container_width=True)
        multi_days = (daily_batch["多批次组合数"] > 0).sum()
        total_days = len(daily_batch)
        if total_days:
            st.info(f"📅 多批次现象：{total_days} 天中有 {multi_days} 天发生（占比 {multi_days/total_days*100:.0f}%）")
        st.markdown("#### 批次最多的 Top 15 组合")
        st.dataframe(grp.nlargest(15, "批次数量")[["生产订单", "料号", "批次数量", "出库笔数"]], use_container_width=True)

# ═══════════════════════════════════════════
# 在库时间分析（FIFO）
# ═══════════════════════════════════════════
if has_in and has_out:
    st.subheader("🏭 在库时间分析（FIFO）")
    dwell = []
    part_stats = []
    in_part_col = "零件图号" if "零件图号" in df_in.columns else None
    out_part_col = "零件图号" if "零件图号" in df_out.columns else None
    if in_part_col and out_part_col:
        in_series = df_in[in_part_col].fillna("").astype(str).str.strip()
        out_series = df_out[out_part_col].fillna("").astype(str).str.strip()
        parts = sorted(set(in_series[in_series != ""]) | set(out_series[out_series != ""]))
        st.caption("已按零件图号分别进行先进先出匹配")
        for part in parts:
            iq = df_in[in_series == part][["时间", "数量"]].copy().sort_values("时间").reset_index(drop=True)
            oq = df_out[out_series == part][["时间", "数量"]].copy().sort_values("时间").reset_index(drop=True)
            idx = 0
            part_dwell = []
            for _, row in oq.iterrows():
                qty = row["数量"]
                ot = row["时间"]
                while qty > 0 and idx < len(iq):
                    r = iq.iloc[idx]
                    if r["数量"] <= 0:
                        idx += 1
                        continue
                    m = min(qty, r["数量"])
                    rec = {"零件图号": part, "在库天数": (ot - r["时间"]).total_seconds() / 86400, "匹配数量": m}
                    dwell.append(rec)
                    part_dwell.append(rec)
                    iq.at[idx, "数量"] -= m
                    qty -= m
                    if iq.iloc[idx]["数量"] <= 0:
                        idx += 1
            if part_dwell:
                part_stats.append({
                    "零件图号": part,
                    "出库箱数": oq["数量"].sum(),
                    "平均在库天数": np.average([d["在库天数"] for d in part_dwell], weights=[d["匹配数量"] for d in part_dwell])
                })
    else:
        iq = df_in[["时间", "数量"]].copy().sort_values("时间")
        oq = df_out[["时间", "数量"]].copy().sort_values("时间")
        idx = 0
        for _, row in oq.iterrows():
            qty = row["数量"]
            ot = row["时间"]
            while qty > 0 and idx < len(iq):
                r = iq.iloc[idx]
                if r["数量"] <= 0:
                    idx += 1
                    continue
                m = min(qty, r["数量"])
                dwell.append({"在库天数": (ot - r["时间"]).total_seconds() / 86400, "匹配数量": m})
                iq.at[iq.index[idx], "数量"] -= m
                qty -= m
                if iq.iloc[idx]["数量"] <= 0:
                    idx += 1
    if dwell:
        df_d = pd.DataFrame(dwell)
        avg_d = np.average(df_d["在库天数"], weights=df_d["匹配数量"])
        c1, c2, c3 = st.columns(3)
        c1.metric("平均在库", f"{avg_d:.1f} 天")
        c2.metric("中位在库", f"{np.median(np.repeat(df_d['在库天数'].values, df_d['匹配数量'].astype(int))):.1f} 天")
        c3.metric("最长在库", f"{df_d['在库天数'].max():.1f} 天")
        fig = px.histogram(df_d, x="在库天数", y="匹配数量", nbins=30, title="在库时间分布")
        fig.add_vline(x=avg_d, line_dash="dash", line_color="red")
        st.plotly_chart(fig, use_container_width=True)
        if part_stats:
            part_df = pd.DataFrame(part_stats)
            part_df["平均在库天数"] = part_df["平均在库天数"].round(2)
            st.markdown("#### 各零件在库时间")
            st.dataframe(part_df, use_container_width=True)

    # ── 规划建议 ──
    st.subheader("📋 规划建议")
    daily_in_rate = total_in / days if days else 0
    if has_in and "托盘数" in df_in.columns and total_pallets_in > 0:
        daily_pallet = total_pallets_in / days if days else 0
        if dwell:
            req_pallet = daily_pallet * avg_d
            buffer = req_pallet * BUFFER_RATE
            st.info(f"🧱 **库容规划（托盘级）**\n- 日均入库：{daily_pallet:.1f} 托\n- 平均在库：{avg_d:.1f} 天\n- 基础储位需求：{req_pallet:,.0f} 托\n- 建议储位（含{int(BUFFER_RATE*100)}%缓冲）：{req_pallet + buffer:,.0f} 托\n- 按 {RACK_LEVELS} 层货架 → 约 {((req_pallet + buffer)/RACK_LEVELS):,.0f} 个托盘位/层")
        else:
            st.info(f"🧱 日均入库 {daily_pallet:.1f} 托")
    elif days > 0:
        st.info(f"日均入库 {daily_in_rate:.0f} 箱")

    avg_h = df_in.groupby("小时").size().mean()
    max_h = df_in.groupby("小时").size().max()
    st.info(f"收货区：{max(1, round(avg_h/IN_EFF))} ~ {max(1, round(max_h/IN_EFF))} 人（每人每小时处理 {IN_EFF} 箱）")
    if has_out:
        avg_ho = df_out.groupby("小时").size().mean()
        max_ho = df_out.groupby("小时").size().max()
        st.info(f"发货区：{max(1, round(avg_ho/OUT_EFF))} ~ {max(1, round(max_ho/OUT_EFF))} 人（每人每小时处理 {OUT_EFF} 箱）")

    st.markdown("---")
    st.subheader("🤖 AI 智能解读")
    in_min = max(1, round(avg_h / IN_EFF))
    in_max = max(1, round(max_h / IN_EFF))
    out_min = max(1, round(avg_ho / OUT_EFF))
    out_max = max(1, round(max_ho / OUT_EFF))
    ai_daily_pallet = ai_recommend = ai_per_level = None
    if dwell and "托盘数" in df_in.columns and total_pallets_in > 0:
        ai_daily_pallet = total_pallets_in / days if days else 0
        ai_req_pallet = ai_daily_pallet * avg_d
        ai_buffer = ai_req_pallet * BUFFER_RATE
        ai_recommend = ai_req_pallet + ai_buffer
        ai_per_level = ai_recommend / RACK_LEVELS
    ctx = {
        "days": days,
        "total_in": total_in,
        "total_out": total_out,
        "avg_d": avg_d if dwell else None,
        "daily_pallet": ai_daily_pallet,
        "recommend_pallet": ai_recommend,
        "per_level": ai_per_level,
        "in_min": in_min,
        "in_max": in_max,
        "out_min": out_min,
        "out_max": out_max,
    }
    rule_text = build_rule_summary(ctx)
    if AI_ENABLED and AI_KEY:
        prompt = (
            "你是供应链仓储数据分析师。请基于以下已计算好的指标，输出中文分析结论。"
            "要求：分点、简洁、不编造数据，重点给出仓储规划洞察。\n"
            f"分析天数：{days} 天；入库总量：{total_in:,} 箱；出库总量：{total_out:,} 箱。\n"
        )
        if dwell:
            prompt += f"平均在库：{avg_d:.1f} 天；日均入库托盘：{ai_daily_pallet:.1f} 托；建议储位：{ai_recommend:,.0f} 托（{RACK_LEVELS} 层货架，每层约 {ai_per_level:,.0f} 个托盘位）。\n"
        prompt += f"收货区建议 {in_min}~{in_max} 人；发货区建议 {out_min}~{out_max} 人。\n"
        prompt += "请给出：1) 整体概况；2) 周转效率判断；3) 库容与人员建议；4) 潜在风险或改进方向。"
        ai_error = ""
        try:
            ai_text = call_llm(prompt, AI_MODEL, AI_URL, AI_KEY, AI_TIMEOUT)
        except Exception as exc:
            ai_text = None
            ai_error = str(exc)
        if ai_text:
            st.markdown(ai_text)
            st.caption(f"由大模型生成 · 模型：{AI_MODEL}")
        else:
            st.info(rule_text)
            if ai_error:
                st.error(f"大模型调用失败：{ai_error}")
            st.caption("已回退到规则版解读")
    else:
        st.info(rule_text)
        st.caption("未配置大模型密钥，显示规则版解读")

# ── 原始数据 ──
with st.expander("查看原始数据"):
    ta, tb = st.tabs(["入库", "出库"])
    with ta:
        if has_in: st.dataframe(df_in, use_container_width=True)
    with tb:
        if has_out: st.dataframe(df_out, use_container_width=True)
