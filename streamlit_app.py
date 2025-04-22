import streamlit as st
import pandas as pd
import re
import io
import os

# — ページ設定 —
st.set_page_config(page_title="G‑Change｜Googleリスト整形＋NGリスト除外", layout="wide")

# — タイトル＆スタイル —
st.markdown("""
    <style>
    h1 { color: #800000; }
    </style>
""", unsafe_allow_html=True)
st.title("🚗 G‑Change｜企業情報自動整形＆NG除外ツール（VerX.X）")

# — ファイルアップロード —
uploaded_file = st.file_uploader("📤 整形対象のExcelファイルをアップロードしてください", type=["xlsx"])

# — NGリスト選択肢を GitHub 直下から自動取得 —
nglist_files = [
    f for f in os.listdir()
    if f.endswith(".xlsx")
       and f != (uploaded_file.name if uploaded_file else "")
       and "template" not in f.lower()
]
nglist_options = ["なし"] + [os.path.splitext(f)[0] for f in nglist_files]
nglist_choice = st.selectbox("👥 クライアントNGリストを選択してください", nglist_options)

# — 前処理用 normalize —
def normalize(text):
    if pd.isna(text):
        return ""
    s = str(text).strip().replace(" ", " ").replace("　", " ")
    return re.sub(r"[−–—―]", "-", s)

# — 縦型リスト抽出用 —
def extract_info(lines):
    company = normalize(lines[0]) if lines else ""
    industry = address = phone = ""
    for line in lines[1:]:
        s = normalize(line)
        # 業種
        if "·" in s or "⋅" in s:
            industry = re.split(r"[·⋅]", s)[-1].strip()
        # 電話
        m = re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", s)
        if m:
            phone = m.group()
        # 住所
        if not address and any(tok in s for tok in ["丁目","町","番","区","-","−"]):
            address = s
    return pd.Series([company, industry, address, phone],
                     index=["企業名","業種","住所","電話番号"])

def is_company_line(line):
    return not re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", normalize(line))

# — メイン処理 —
if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    # 「入力マスター」シートを優先
    sheet = next((s for s in xls.sheet_names if "入力マスター" in s), xls.sheet_names[0])
    raw = pd.read_excel(uploaded_file, sheet_name=sheet, header=None)

    # ヘッダー行検出（先頭10行）
    header_row = None
    for i, row in raw.head(10).iterrows():
        if any(str(c).strip() in ("企業様名称","企業名","職種") for c in row):
            header_row = i
            break

    # 縦型 vs 横型 の分岐
    if header_row is None and raw.shape[1] == 1:
        # —— 縦型リスト
        lines = raw[0].dropna().astype(str).tolist()
        groups, cur = [], []
        for ln in lines:
            if is_company_line(ln):
                if cur: groups.append(cur)
                cur = [ln]
            else:
                cur.append(ln)
        if cur: groups.append(cur)
        df = pd.DataFrame([extract_info(g) for g in groups],
                          columns=["企業名","業種","住所","電話番号"])
    else:
        # —— 横型リスト（入力マスター形式）
        if header_row is not None:
            df0 = pd.read_excel(uploaded_file, sheet_name=sheet, header=header_row)
        else:
            df0 = pd.read_excel(uploaded_file, sheet_name=sheet)
        # 列名クリーニング
        df0.columns = [str(c).strip() for c in df0.columns]
        # 必要なリネーム
        df0.rename(columns={
            "企業様名称":"企業名",
            "職種":"業種"          # ← ここを追加
        }, inplace=True)
        # 必須４列を確実に揃える
        for col in ("企業名","業種","住所","電話番号"):
            if col not in df0.columns:
                df0[col] = ""
        df = df0[["企業名","業種","住所","電話番号"]].copy()

    st.success(f"✅ 整形完了！（企業数：{len(df)} 件）")
    st.dataframe(df, use_container_width=True)

    # — NG除外 —
    if nglist_choice != "なし":
        ng = pd.read_excel(f"{nglist_choice}.xlsx")
        ng.rename(columns={"企業様名称":"企業名","職種":"業種"}, inplace=True)
        ngc = ng.get("企業名",pd.Series()).dropna().astype(str).tolist()
        ngp = ng.get("電話番号",pd.Series()).dropna().astype(str).tolist()

        m1 = df["企業名"].astype(str).apply(lambda x: any(n in x for n in ngc))
        m2 = df["電話番号"].astype(str).isin(ngp)
        rem = df[m1|m2]
        df = df[~(m1|m2)]

        st.success(f"🧹 NG除外完了！（除外件数：{len(rem)} 件）")
        if not rem.empty:
            st.error("🚫 除外された企業一覧")
            st.dataframe(rem, use_container_width=True)

    # — ダウンロード —
    base = os.path.splitext(uploaded_file.name)[0]
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        df.to_excel(w, index=False, sheet_name="リスト")
    st.download_button("📥 Excel をダウンロード",
                       data=buf.getvalue(),
                       file_name=f"{base}：リスト.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
