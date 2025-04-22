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
    .stSelectbox > div[data-testid="stMarkdownContainer"] { margin-bottom: 0.5rem; }
    </style>
""", unsafe_allow_html=True)
st.title("🚗 G‑Change｜企業情報自動整形＆NG除外ツール（VerX.X）")

# — ファイルアップロード —
uploaded_file = st.file_uploader("📤 整形対象のExcelファイルをアップロードしてください", type=["xlsx"])

# — NGリスト選択肢を GitHub 直下から自動取得 —
nglist_files = [
    f for f in os.listdir()
    if f.endswith(".xlsx")
       and f != os.path.basename(uploaded_file.name)  # アップロードファイル自身は除外
       and "template" not in f.lower()
]
nglist_options = ["なし"] + [os.path.splitext(f)[0] for f in nglist_files]
nglist_choice = st.selectbox("👥 クライアントNGリストを選択してください", nglist_options)

# — 正規化関数（前処理） —
def normalize(text):
    if pd.isna(text):
        return ""
    s = str(text).strip().replace(" ", " ").replace("　", " ")
    return re.sub(r"[−–—―]", "-", s)

# — 縦型リスト用の抽出関数 —
def extract_info(lines):
    company = normalize(lines[0]) if lines else ""
    industry = address = phone = ""
    for line in lines[1:]:
        s = normalize(line)
        # industry
        if "·" in s or "⋅" in s:
            industry = s.split("·")[-1].split("⋅")[-1].strip()
        # phone
        m = re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", s)
        if m:
            phone = m.group()
        # address
        if not address and any(tok in s for tok in ["丁目","町","番","区","‐","-"]):
            address = s
    return pd.Series([company, industry, address, phone],
                     index=["企業名","業種","住所","電話番号"])

# — 縦型リストの行判定 —
def is_company_line(line):
    s = normalize(line)
    # 電話番号が含まれないものを「企業名行」とみなす
    return not re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", s)

# — メイン処理 —
if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    # 「入力マスター」シートを優先、なければ先頭シート
    sheet_to_use = next((s for s in xls.sheet_names if "入力マスター" in s), xls.sheet_names[0])

    # まずヘッダー無しで読み込み、ヘッダー行を検出
    raw = pd.read_excel(uploaded_file, sheet_name=sheet_to_use, header=None)
    header_row = None
    for i, row in raw.head(10).iterrows():
        if any(str(cell).strip() in ("企業様名称","企業名") for cell in row):
            header_row = i
            break

    # 縦型 or 横型判定
    if header_row is None and raw.shape[1] == 1:
        # —— 縦型リスト
        lines = raw[0].dropna().tolist()
        groups, current = [], []
        for ln in lines:
            s = normalize(ln)
            if is_company_line(s):
                if current:
                    groups.append(current)
                current = [s]
            else:
                current.append(s)
        if current:
            groups.append(current)
        df = pd.DataFrame([extract_info(g) for g in groups],
                          columns=["企業名","業種","住所","電話番号"])
    else:
        # —— 横型リスト（テンプレート同様式）
        if header_row is not None:
            df_temp = pd.read_excel(uploaded_file,
                                    sheet_name=sheet_to_use,
                                    header=header_row)
        else:
            df_temp = pd.read_excel(uploaded_file, sheet_name=sheet_to_use)

        # 列名クリーニング＆リネーム
        df_temp.columns = [str(c).strip() for c in df_temp.columns]
        df_temp.rename(columns={"企業様名称":"企業名"}, inplace=True)

        # 必要列を確実に確保
        for col in ("企業名","業種","住所","電話番号"):
            if col not in df_temp.columns:
                df_temp[col] = ""
        df = df_temp[["企業名","業種","住所","電話番号"]].copy()

    st.success(f"✅ 整形完了！（企業数：{len(df)} 件）")
    st.dataframe(df, use_container_width=True)

    # — NGリスト除外処理 —
    if nglist_choice != "なし":
        ng_path = nglist_choice + ".xlsx"
        ng_df = pd.read_excel(ng_path)
        # NG列名が違うケース対応
        ng_df.rename(columns={"企業様名称":"企業名"}, inplace=True)
        ng_companies = ng_df.get("企業名", pd.Series()).dropna().astype(str).tolist()
        ng_phones    = ng_df.get("電話番号", pd.Series()).dropna().astype(str).tolist()

        # 部分一致：企業名 / 完全一致：電話番号
        mask_c = df["企業名"].astype(str).apply(
            lambda x: any(ng in x for ng in ng_companies)
        )
        mask_p = df["電話番号"].astype(str).isin(ng_phones)

        removed = df[mask_c | mask_p]
        df = df[~(mask_c | mask_p)]
        st.success(f"🧹 NG除外完了！（除外件数：{len(removed)} 件）")
        if not removed.empty:
            st.error("🚫 除外された企業一覧")
            st.dataframe(removed, use_container_width=True)

    # — 出力ファイル名＆ダウンロードボタン —
    base = os.path.splitext(uploaded_file.name)[0]
    out_name = f"{base}：リスト.xlsx"
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="リスト")
    st.download_button(
        "📥 Excelファイルをダウンロード",
        data=buf.getvalue(),
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
