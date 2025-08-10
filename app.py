import pandas as pd
import streamlit as st
from io import BytesIO
import hashlib
import time

st.set_page_config(page_title="PaperOn Sales Mini", page_icon="🧾", layout="wide")

# ======================================================
# Global CSS: white bg + high-contrast text + compact UI
# ＋ Data Editor（編集中セル）の視認性を強化
# ======================================================
st.markdown("""
<style>
/* Base: force white background and dark text everywhere */
:root, body, .stApp { background:#FFFFFF !important; color:#0B0B0C !important; }
.block-container { padding-top: .6rem !important; padding-bottom: 1.0rem !important; }

/* Sidebar visibility */
section[data-testid="stSidebar"], [data-testid="stSidebar"] {
  background:#FFFFFF !important; color:#0B0B0C !important;
}
[data-testid="stSidebar"] * { color:#0B0B0C !important; }
[data-testid="stSidebar"] .stMarkdown, [data-testid="stSidebar"] p, 
[data-testid="stSidebar"] span, [data-testid="stSidebar"] label { color:#0B0B0C !important; }

/* Headings & text */
h1,h2,h3,h4,h5,h6 { color:#0B0B0C !important; letter-spacing:.2px; }
h1 { font-weight:800; margin: .2rem 0 .6rem !important; }
h2 { font-weight:700; margin: .2rem 0 .5rem !important; }
h3 { font-weight:700; margin: .2rem 0 .4rem !important; }
p, span, div, small, label { color:#0B0B0C; }

/* Inputs */
.stTextInput input, .stNumberInput input, textarea, select {
  color:#0B0B0C !important; background:#FFFFFF !important; border-radius: 10px !important;
}
.stSelectbox div[role="combobox"] { color:#0B0B0C !important; }
[data-baseweb="select"] * { color:#0B0B0C !important; }

/* Data components */
[data-testid="stDataFrame"] *, [data-testid="stDataEditorGrid"] *, [data-testid="stTable"] * {
  color:#0B0B0C !important;
}

/* Data Editor: 編集中セルの視認性（白文字＋濃い青背景）*/
[data-testid="stDataEditorGrid"] [role="gridcell"][aria-selected="true"] {
  background:#1E3A8A !important; color:#FFFFFF !important; /* 選択セル背景・文字 */
}
[data-testid="stDataEditorGrid"] div[contenteditable="true"] {
  color:#FFFFFF !important; caret-color:#FFFFFF !important; /* 入力中文字色・カーソル */
}
[data-testid="stDataEditorGrid"] div[contenteditable="true"]::selection {
  background:#1D4ED8 !important; color:#FFFFFF !important; /* 入力中文字の選択配色 */
}

/* Buttons */
.stButton > button { 
  border-radius: 12px !important; font-weight: 700 !important; 
  background:#2563EB !important; color:#FFFFFF !important; border: none !important;
  padding: .5rem .9rem !important;
}
.stButton > button:hover { background:#1D4ED8 !important; color:#FFFFFF !important; }

/* Cards */
.paperon-card {
  background:#FAFAFA; color:#0B0B0C;
  border: 1px solid #ECECEC; border-radius: 14px; 
  box-shadow: 0 4px 14px rgba(0,0,0,0.04); padding: 14px 14px; margin-bottom: 12px;
}

/* KPI tiles */
.kpi {
  background:#FFFFFF; border:1px solid #EAEAEA; border-radius:14px; padding: 10px 12px;
}
.kpi .label { color:#475569; font-size: 12px; font-weight: 700; }
.kpi .value { color:#0B0B0C; font-size: 20px; font-weight: 800; }
.kpi .sub   { color:#334155; font-size: 11px; font-weight: 600; }

/* Order cards */
.order-card {
  border: 1px solid #EDEDED; border-radius: 16px; padding: 12px 14px; 
  background: #FFFFFF; box-shadow: 0 3px 12px rgba(0,0,0,0.06);
}
.order-id   { font-weight:800; color:#2563EB; font-size: 13px; }
.order-cust { font-weight:700; font-size: 16px; color:#0B0B0C; }
.order-meta { color:#334155; font-size:12px; }
.badge { 
  display:inline-block; padding:2px 8px; border-radius:9999px; 
  background:#EFF6FF; color:#1D4ED8; font-size:11px; font-weight:800; border:1px solid #DBEAFE;
}

/* Sticky CTA */
.cta-bar {
  position: sticky; bottom: 6px; z-index: 50;
  background:#FFFFFFE6; backdrop-filter: blur(6px);
  border: 1px solid #EDEDED; border-radius: 12px;
  padding: 8px 10px; box-shadow: 0 8px 24px rgba(0,0,0,0.06);
}

/* Compact tables/editor spacing */
[data-testid="stDataFrame"] > div { padding: 0 !important; }
</style>
""", unsafe_allow_html=True)

# =========================
# Config
# =========================
HEADER_COLS = [
  "supplier.companyName","supplier.department","supplier.personName","supplier.postalCode","supplier.address","supplier.tel","supplier.fax",
  "orderer.companyName","orderer.department","orderer.personName","orderer.owner","orderer.postalCode","orderer.address","orderer.tel","orderer.fax","orderer.email","orderer.homepage",
  "totalPriceInfo.totalPrice","totalPriceInfo.subTotalPrice","totalPriceInfo.taxAmount","totalPriceInfo.taxInfo",
  "subTotals.totalPriceInclude08","subTotals.totalPriceInclude10","subTotals.totalPriceIncludeEtc",
  "subTotals.subTotalPriceExclude08","subTotals.subTotalPriceExclude10","subTotals.subTotalPriceExcludeEtc",
  "subTotals.discount","subTotals.taxAmount08","subTotals.taxAmount10","subTotals.taxAmountEtc","subTotals.taxInfo","subTotals.etcAmount"
]
ITEM_COLS = [
  "items.name","items.num","items.count","items.date","items.discount","items.etc","items.quantityUnit",
  "items.taxExcludedUnitPrice","items.taxExcludedPrice","items.taxIncludedUnitPrice","items.taxIncludedPrice","items.taxAmount","items.taxInfo"
]
MONEY_COLS = [
  "items.taxExcludedUnitPrice","items.taxExcludedPrice","items.taxIncludedUnitPrice","items.taxIncludedPrice","items.taxAmount",
  "totalPriceInfo.totalPrice","totalPriceInfo.subTotalPrice","totalPriceInfo.taxAmount",
  "subTotals.totalPriceInclude08","subTotals.totalPriceInclude10","subTotals.totalPriceIncludeEtc",
  "subTotals.subTotalPriceExclude08","subTotals.subTotalPriceExclude10","subTotals.subTotalPriceExcludeEtc",
  "subTotals.discount","subTotals.taxAmount08","subTotals.taxAmount10","subTotals.taxAmountEtc","subTotals.etcAmount"
]
JAPANESE_LABELS = {
    "orderId": "注文ID",
    "orderer.companyName": "得意先名",
    "orderer.personName": "担当者",
    "totalPriceInfo.subTotalPrice": "小計",
    "totalPriceInfo.taxAmount": "消費税",
    "totalPriceInfo.totalPrice": "合計",
    "items.name":"商品名",
    "items.num":"品番",
    "items.count":"数量",
    "items.quantityUnit":"単位",
    "items.taxExcludedUnitPrice":"単価(税抜)",
    "items.taxExcludedPrice":"金額(税抜)",
    "items.taxIncludedUnitPrice":"単価(税込)",
    "items.taxIncludedPrice":"金額(税込)",
    "items.taxAmount":"税額",
}

# =========================
# Helpers
# =========================
def detect_encoding(file):
    for enc in ["utf-8-sig","cp932","shift_jis","utf-8"]:
        try:
            file.seek(0)
            _ = pd.read_csv(file, nrows=5, encoding=enc)
            file.seek(0)
            return enc
        except Exception:
            continue
    return None

def to_number(v):
    if pd.isna(v): return pd.NA
    s = str(v).replace(",", "").replace("円", "").strip()
    try:
        return pd.to_numeric(s, errors="coerce")
    except Exception:
        return pd.NA

def make_order_key(row, present_cols):
    key_cols = [c for c in [
        "orderer.companyName","orderer.personName",
        "totalPriceInfo.subTotalPrice","totalPriceInfo.taxAmount","totalPriceInfo.totalPrice"
    ] if c in present_cols]
    parts = [str(row.get(c, "")) for c in key_cols]
    return "|".join(parts)

def to_order_id(key):
    h = hashlib.sha1(key.encode("utf-8")).hexdigest()[:8].upper()
    return f"ORD-{h}"

def normalize_tables(df):
    present_headers = [c for c in HEADER_COLS if c in df.columns]
    df_ff = df.copy()
    if present_headers:
        df_ff[present_headers] = df_ff[present_headers].ffill()

    if "items.name" not in df_ff.columns:
        raise ValueError("CSVに 'items.name' 列がありません。PaperOnの出力列名をご確認ください。")
    items = df_ff[df_ff["items.name"].notna()].copy()

    present_money = [c for c in MONEY_COLS if c in items.columns]
    for c in present_money:
        items[c] = items[c].apply(to_number)

    present_cols = items.columns
    keys = items.apply(lambda r: make_order_key(r, present_cols), axis=1)
    if (keys == "").all():
        keys = items.index.astype(str)
    order_ids = keys.apply(to_order_id)
    items.insert(0, "orderId", order_ids)

    agg_map = {}
    for c in ["totalPriceInfo.subTotalPrice","totalPriceInfo.taxAmount","totalPriceInfo.totalPrice"]:
        if c in items.columns:
            agg_map[c] = "max"

    base_cols = [c for c in ["orderer.companyName","orderer.personName"] if c in items.columns]
    orders = items.groupby("orderId", as_index=False).agg({**{c:"first" for c in base_cols}, **agg_map})

    order_cols = ["orderId"] + base_cols + [c for c in ["totalPriceInfo.subTotalPrice","totalPriceInfo.taxAmount","totalPriceInfo.totalPrice"] if c in orders.columns]
    orders = orders[order_cols]
    return orders, items

def yen_fmt(x):
    try:
        x = float(x)
        return f"{x:,.0f} 円"
    except Exception:
        return "-"

# =========================
# State init
# =========================
if "orders" not in st.session_state:
    st.session_state["orders"] = pd.DataFrame()
if "items" not in st.session_state:
    st.session_state["items"] = pd.DataFrame()
if "selected_order" not in st.session_state:
    st.session_state["selected_order"] = None
if "tax_rate" not in st.session_state:
    st.session_state["tax_rate"] = 0.10  # 10%
if "page" not in st.session_state:
    st.session_state["page"] = "① アップロード＆注文一覧"

# =========================
# Navigation (programmatically switchable)
# =========================
st.sidebar.header("メニュー")
page = st.sidebar.radio(
    "ページを選択",
    ["① アップロード＆注文一覧", "② 注文詳細（編集）"],
    index=["① アップロード＆注文一覧", "② 注文詳細（編集）"].index(st.session_state["page"]),
    key="page_radio",
)
# keep session_state["page"] in sync
st.session_state["page"] = page

# =========================
# Hand-entry (page1 inline)
# =========================
def hand_entry_form():
    st.subheader("＋ 注文を手入力で登録")
    with st.form("hand_entry_form", clear_on_submit=False):
        c1, c2, c3, c4 = st.columns([1,1,1,1])
        with c1:
            cust = st.text_input("得意先名")
        with c2:
            person = st.text_input("担当者")
        with c3:
            tax_rate = st.number_input("消費税率", min_value=0.0, max_value=0.3, step=0.01, value=float(st.session_state["tax_rate"]))
        with c4:
            order_date = st.date_input("注文日")

        st.caption("明細行（必要なだけ追加してください）")
        items_df = st.data_editor(
            pd.DataFrame([{"商品名":"","品番":"","数量":1,"単価(税抜)":0.0,"単位":""}]),
            num_rows="dynamic",
            use_container_width=True,
            key="hand_items_editor"
        )

        submitted = st.form_submit_button("🖫 注文登録（手入力）", type="primary")
        if submitted:
            # Build internal items
            tmp = items_df.rename(columns={
                "商品名":"items.name","品番":"items.num","数量":"items.count","単価(税抜)":"items.taxExcludedUnitPrice","単位":"items.quantityUnit"
            })
            tmp["items.taxExcludedPrice"] = pd.to_numeric(tmp["items.taxExcludedUnitPrice"], errors="coerce").fillna(0) * pd.to_numeric(tmp["items.count"], errors="coerce").fillna(0)

            # Create a pseudo row to derive orderId
            key = "|".join([cust or "", person or "", str(tmp["items.taxExcludedPrice"].sum())])
            order_id = to_order_id(key)
            tmp.insert(0, "orderId", order_id)
            tmp["orderer.companyName"] = cust
            tmp["orderer.personName"] = person

            subtotal = float(tmp["items.taxExcludedPrice"].sum())
            tax = round(subtotal * tax_rate, 0)
            total = subtotal + tax

            orders_row = {
                "orderId": order_id,
                "orderer.companyName": cust,
                "orderer.personName": person,
                "totalPriceInfo.subTotalPrice": subtotal,
                "totalPriceInfo.taxAmount": tax,
                "totalPriceInfo.totalPrice": total
            }

            # append
            st.session_state["items"] = pd.concat([st.session_state["items"], tmp], ignore_index=True)
            st.session_state["orders"] = pd.concat([st.session_state["orders"], pd.DataFrame([orders_row])], ignore_index=True)
            st.session_state["tax_rate"] = tax_rate
            st.success(f"手入力の注文を登録しました。注文ID: {order_id}  合計: {int(total):,} 円")

            # Move to detail page immediately
            st.session_state["selected_order"] = order_id
            st.session_state["page"] = "② 注文詳細（編集）"
            st.experimental_rerun()

# =========================
# Page 1
# =========================
if st.session_state["page"] == "① アップロード＆注文一覧":
    st.markdown('<div class="paperon-card">', unsafe_allow_html=True)
    st.title("🧾 アップロード & 注文リスト")
    st.caption("CSV一括登録 と 手入力登録 の両方に対応。")
    st.markdown('</div>', unsafe_allow_html=True)

    # Top row: hand-entry (left) and CSV upload (right)
    left, right = st.columns([1,1])
    with left:
        st.markdown('<div class="paperon-card">', unsafe_allow_html=True)
        hand_entry_form()
        st.markdown('</div>', unsafe_allow_html=True)

    with right:
        st.markdown('<div class="paperon-card">', unsafe_allow_html=True)
        st.subheader("📥 CSVから注文登録")
        up = st.file_uploader("PaperOnのCSVを選択（Shift_JIS / UTF-8）", type=["csv"], label_visibility="collapsed")
        if up is not None:
            enc = detect_encoding(up)
            if enc is None:
                st.error("エンコーディング判定に失敗しました。UTF-8 または Shift_JIS で保存して再試行してください。")
            else:
                df = pd.read_csv(up, encoding=enc)
                st.info(f"読み込み成功: rows={len(df)}, cols={len(df.columns)}, encoding={enc}")
                cols_btn = st.columns([1,1])
                if cols_btn[0].button("📥 注文登録（CSV）", type="primary"):
                    prog = st.progress(0, text="正規化しています...")
                    time.sleep(0.2)
                    try:
                        orders, items = normalize_tables(df)
                        time.sleep(0.2)
                        for i in range(10):
                            prog.progress(int((i+1)/10*100), text=f"登録中... {i+1}/10")
                            time.sleep(0.03)
                        st.success(f"登録完了：注文 {len(orders)} 件 / 明細 {len(items)} 件")
                        st.session_state["orders"] = pd.concat([st.session_state["orders"], orders], ignore_index=True)
                        st.session_state["items"]  = pd.concat([st.session_state["items"], items], ignore_index=True)
                    except Exception as e:
                        st.error(f"登録に失敗しました: {e}")
        st.markdown('</div>', unsafe_allow_html=True)

    # Summary
    st.subheader("サマリー")
    if st.session_state["orders"].empty:
        st.info("まだ注文が登録されていません。")
    else:
        orders = st.session_state["orders"].copy()
        total_orders = len(orders)
        total_amount = pd.to_numeric(orders.get("totalPriceInfo.totalPrice"), errors="coerce").fillna(0).sum() if "totalPriceInfo.totalPrice" in orders.columns else 0
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f'<div class="kpi"><div class="label">登録済み注文</div><div class="value">{total_orders:,}</div></div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="kpi"><div class="label">合計金額（合算）</div><div class="value">{yen_fmt(total_amount)}</div></div>', unsafe_allow_html=True)
        with c3:
            if "orderer.companyName" in orders.columns and len(orders):
                top = orders["orderer.companyName"].value_counts()
                top_name = top.index[0] if len(top) else "-"
                cnt = int(top.iloc[0]) if len(top) else 0
                st.markdown(f'<div class="kpi"><div class="label">最多の得意先</div><div class="value">{top_name}</div><div class="sub">{cnt} 件</div></div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="kpi"><div class="label">最多の得意先</div><div class="value">-</div></div>', unsafe_allow_html=True)

    # Orders as cards
    st.subheader("注文リスト")
    if st.session_state["orders"].empty:
        st.caption("※ 登録が完了するとここに表示されます。")
    else:
        orders = st.session_state["orders"].copy()
        cols = st.columns(3)
        for i, (_, r) in enumerate(orders.iterrows()):
            with cols[i % 3]:
                total = yen_fmt(r.get("totalPriceInfo.totalPrice", None))
                sub = yen_fmt(r.get("totalPriceInfo.subTotalPrice", None))
                tax = yen_fmt(r.get("totalPriceInfo.taxAmount", None))
                cust = r.get("orderer.companyName","-")
                person = r.get("orderer.personName","-")

                st.markdown(f'''
                <div class="order-card">
                  <div class="order-id">{r["orderId"]}</div>
                  <div class="order-cust">{cust}</div>
                  <div class="order-meta">担当: {person if person else "-"}</div>
                  <div style="margin-top:8px;">
                    <span class="badge">小計 {sub}</span>
                    <span class="badge" style="margin-left:6px;">税 {tax}</span>
                    <span class="badge" style="margin-left:6px;background:#DCFCE7;border-color:#BBF7D0;color:#166534;">合計 {total}</span>
                  </div>
                </div>
                ''', unsafe_allow_html=True)

                if st.button("詳細を開く", key=f"open_{r['orderId']}"):
                    st.session_state["selected_order"] = r["orderId"]
                    st.session_state["page"] = "② 注文詳細（編集）"  # ← ページ2へ
                    st.experimental_rerun()

    st.markdown("---")
    colL, colR = st.columns([1,1])
    with colL:
        if st.button("🧼 全データ削除（デモ用）"):
            st.session_state["orders"] = pd.DataFrame()
            st.session_state["items"]  = pd.DataFrame()
            st.session_state["selected_order"] = None
            st.experimental_rerun()
    with colR:
        st.caption("カードのクリックで詳細を開く実装も可能です。")

# =========================
# Page 2: Order Detail
# =========================
elif st.session_state["page"] == "② 注文詳細（編集）":
    st.title("📄 注文詳細（入力・編集）")

    if st.session_state["orders"].empty:
        st.info("注文がありません。『① アップロード＆注文一覧』でCSV/手入力から登録してください。")
        st.stop()

    order_ids = st.session_state["orders"]["orderId"].tolist()
    default_id = st.session_state.get("selected_order", order_ids[0])
    # keep dropdown synced but allow change
    order_id = st.selectbox("注文IDを選択", order_ids, index=order_ids.index(default_id) if default_id in order_ids else 0)
    st.session_state["selected_order"] = order_id

    orders = st.session_state["orders"].copy()
    items = st.session_state["items"].copy()
    o_row = orders[orders["orderId"]==order_id].head(1).to_dict("records")[0]

    # Header summary card
    total = yen_fmt(o_row.get("totalPriceInfo.totalPrice", None))
    sub = yen_fmt(o_row.get("totalPriceInfo.subTotalPrice", None))
    tax = yen_fmt(o_row.get("totalPriceInfo.taxAmount", None))
    st.markdown(f"""
    <div class="paperon-card">
      <div style="display:flex; justify-content: space-between; align-items:center; gap: 10px; flex-wrap: wrap;">
        <div>
          <div class="order-id">{order_id}</div>
          <div class="order-cust" style="margin-top:4px;">{o_row.get("orderer.companyName","-")}</div>
          <div class="order-meta">担当: {o_row.get("orderer.personName","-")}</div>
        </div>
        <div style="display:flex; gap:10px;">
          <div class="kpi"><div class="label">小計</div><div class="value">{sub}</div></div>
          <div class="kpi"><div class="label">消費税</div><div class="value">{tax}</div></div>
          <div class="kpi"><div class="label">合計</div><div class="value">{total}</div></div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # Header form
    st.subheader("ヘッダ情報")
    with st.form("header_form"):
        col1, col2, col3 = st.columns(3)
        with col1:
            cust = st.text_input("得意先名", o_row.get("orderer.companyName",""))
        with col2:
            person = st.text_input("担当者", o_row.get("orderer.personName",""))
        with col3:
            st.session_state["tax_rate"] = st.number_input("消費税率", min_value=0.0, max_value=0.3, step=0.01, value=float(st.session_state["tax_rate"]))

        col4, col5, col6 = st.columns(3)
        with col4:
            sub_total = st.text_input("小計（数値）", str(o_row.get("totalPriceInfo.subTotalPrice","") or ""))
        with col5:
            tax_v = st.text_input("消費税（数値）", str(o_row.get("totalPriceInfo.taxAmount","") or ""))
        with col6:
            total_v = st.text_input("合計（数値）", str(o_row.get("totalPriceInfo.totalPrice","") or ""))

        saved_header = st.form_submit_button("ヘッダを保存")
        if saved_header:
            idx = orders.index[orders["orderId"]==order_id][0]
            orders.at[idx,"orderer.companyName"] = cust
            orders.at[idx,"orderer.personName"] = person
            def _num(x):
                try: return float(str(x).replace(","," ").replace("円"," ").strip())
                except: return None
            if "totalPriceInfo.subTotalPrice" in orders.columns:
                orders.at[idx,"totalPriceInfo.subTotalPrice"] = _num(sub_total)
            if "totalPriceInfo.taxAmount" in orders.columns:
                orders.at[idx,"totalPriceInfo.taxAmount"] = _num(tax_v)
            if "totalPriceInfo.totalPrice" in orders.columns:
                orders.at[idx,"totalPriceInfo.totalPrice"] = _num(total_v)
            st.session_state["orders"] = orders
            st.success("ヘッダを保存しました。")

    # Items editor
    st.subheader("明細（編集可能）")
    item_cols_show = ["items.name","items.num","items.count","items.quantityUnit","items.taxExcludedUnitPrice","items.taxExcludedPrice"]
    item_cols_show = [c for c in item_cols_show if c in items.columns]
    df_items = items[items["orderId"]==order_id][["orderId"]+item_cols_show].copy().reset_index(drop=True)
    j_cols = {k:JAPANESE_LABELS.get(k,k) for k in df_items.columns}
    df_show = df_items.rename(columns=j_cols)
    edited = st.data_editor(
        df_show.drop(columns=["注文ID"]),
        num_rows="dynamic",
        use_container_width=True,
        key=f"editor_{order_id}"
    )

    # Sticky CTA
    st.markdown('<div class="cta-bar">', unsafe_allow_html=True)
    colA, colB, colC = st.columns([1,1,2])
    with colA:
        save_clicked = st.button("🖫 明細を保存して再計算", type="primary")
    with colB:
        if st.button("💾 Excelエクスポート"):
            def _download_excel(orders, items):
                buf = BytesIO()
                with pd.ExcelWriter(buf, engine="xlsxwriter") as xw:
                    orders.to_excel(xw, index=False, sheet_name="Orders")
                    items.to_excel(xw, index=False, sheet_name="OrderItems")
                return buf.getvalue()
            data = _download_excel(st.session_state["orders"], st.session_state["items"]) 
            st.download_button("ダウンロード", data=data, file_name="sales_after_edit.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with colC:
        st.caption("税率や数量を編集して保存すると小計/税/合計を再計算します。")
    st.markdown('</div>', unsafe_allow_html=True)

    if save_clicked:
        inv_map = {v:k for k,v in j_cols.items()}
        edited_internal = edited.rename(columns=inv_map)
        new_items = items[items["orderId"]!=order_id].copy()
        edited_internal.insert(0, "orderId", order_id)

        if "items.taxExcludedUnitPrice" in edited_internal.columns and "items.count" in edited_internal.columns:
            try:
                edited_internal["items.taxExcludedPrice"] = pd.to_numeric(edited_internal["items.taxExcludedUnitPrice"], errors="coerce") * pd.to_numeric(edited_internal["items.count"], errors="coerce")
            except Exception:
                pass

        subtotal = 0.0
        if "items.taxExcludedPrice" in edited_internal.columns:
            subtotal = pd.to_numeric(edited_internal["items.taxExcludedPrice"], errors="coerce").fillna(0).sum()

        tax_rate = float(st.session_state["tax_rate"])
        tax = round(subtotal * tax_rate, 0)
        total = subtotal + tax

        if "totalPriceInfo.subTotalPrice" in orders.columns:
            orders.loc[orders["orderId"]==order_id, "totalPriceInfo.subTotalPrice"] = subtotal
        if "totalPriceInfo.taxAmount" in orders.columns:
            orders.loc[orders["orderId"]==order_id, "totalPriceInfo.taxAmount"] = tax
        if "totalPriceInfo.totalPrice" in orders.columns:
            orders.loc[orders["orderId"]==order_id, "totalPriceInfo.totalPrice"] = total

        st.session_state["orders"] = orders
        st.session_state["items"]  = pd.concat([new_items, edited_internal], ignore_index=True)
        st.success(f"保存しました。小計: {subtotal:,.0f} 円 / 税: {tax:,.0f} 円 / 合計: {total:,.0f} 円")
