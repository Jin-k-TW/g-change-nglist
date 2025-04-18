import streamlit as st
import pandas as pd
import os
import re
import io

# ページ設定
st.set_page_config(page_title="G-Change｜リスト整形＋NG除外", layout="wide")

# ヘッダー装飾
st.markdown("""
    <style>
    h1 { color: #800000; }
    </style>
""", unsafe_allow_html=True)

st.title("🚗 G-Change｜Googleリスト整形＋NGリスト自動除外")

# フォルダパス（事前にnglists/を置いておく）
NG_FOLDER = "nglists"

# 企業名NGリスト・電話番号NGリストを取得
def load_nglists():
    ng_options = []
    if os.path.exists(NG_FOLDER):
        files = os.listdir(NG_FOLDER)
        ng_options = [f.replace(".xlsx", "") for f in files if f.endswith(".xlsx")]
    return ng_options

# NGリスト読み込み
def read_nglist(company_name):
    file_path = os.path.join(NG_FOLDER, company_name + ".xlsx")
    if not os.path.exists(file_path):
        return pd.DataFrame(columns=["企業名", "電話番号"])
    df = pd.read_excel(file_path)
    return df

# テキスト整形
def normalize(text):
    text = str(text).strip().replace(" ", " ").replace("　", " ")
    return re.sub(r'[−–—―]', '-', text)

# 整形ロジック
def extract_info(lines):
    company = normalize(lines[0]) if lines else ""
    industry, address, phone = "", "", ""
    review_keywords = ["楽しい", "親切", "人柄", "感じ", "スタッフ", "雰囲気", "交流", "お世話", "ありがとう", "です", "ました", "🙇"]
    ignore_keywords = ["ウェブサイト", "ルート", "営業中", "閉店", "口コミ"]

    for line in lines[1:]:
        line = normalize(line)
        if any(kw in line for kw in ignore_keywords):
            continue
        if any(kw in line for kw in review_keywords):
            continue
        if "·" in line or "⋅" in line:
            parts = re.split(r"[·⋅]", line)
            industry = parts[-1].strip()
        elif re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line):
            phone = re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line).group()
        elif not address and any(x in line for x in ["丁目", "町", "番", "区", "−", "-"]):
            address = line

    return pd.Series([company, industry, address, phone])

def is_company_line(line):
    line = normalize(str(line))
    review_keywords = ["楽しい", "親切", "人柄", "感じ", "スタッフ", "雰囲気", "交流", "お世話", "ありがとう", "です", "ました", "🙇"]
    ignore_keywords = ["ウェブサイト", "ルート", "営業中", "閉店", "口コミ"]
    return not any(kw in line for kw in ignore_keywords + review_keywords) and not re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line)

# --- Streamlit画面 ---

uploaded_file = st.file_uploader("📤 整形対象のリストをアップロード", type=["xlsx"])

nglist_options = load_nglists()

selected_nglist = st.selectbox("🚫 クライアントNGリストを選択してください", ["なし"] + nglist_options)

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, header=None)
    lines = df_raw[0].dropna().tolist()

    groups = []
    current = []
    for line in lines:
        line = normalize(str(line))
        if is_company_line(line):
            if current:
                groups.append(current)
            current = [line]
        else:
            current.append(line)
    if current:
        groups.append(current)

    result_df = pd.DataFrame([extract_info(group) for group in groups],
                             columns=["企業名", "業種", "住所", "電話番号"])

    st.success(f"✅ 整形完了：{len(result_df)}件取得。次にNG除外処理へ進みます。")

    if selected_nglist != "なし":
        ng_df = read_nglist(selected_nglist)
        # NG除外
        before = len(result_df)
        result_df = result_df[
            ~(
                result_df["企業名"].isin(ng_df["企業名"].dropna())
                | result_df["電話番号"].isin(ng_df["電話番号"].dropna())
            )
        ]
        after = len(result_df)
        st.info(f"🚫 NGリスト適用：{before-after}件除外しました。")

    # ダウンロード準備
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        result_df.to_excel(writer, index=False, startrow=1, startcol=1, header=False, sheet_name="入力マスター")
    filename = uploaded_file.name.replace(".xlsx", "：リスト.xlsx")

    st.download_button("📥 整形＋NG除外済リストをダウンロード", data=output.getvalue(),
                       file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
