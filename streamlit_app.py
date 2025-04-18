import os
import streamlit as st
import pandas as pd
import io
import re

# ページ設定
st.set_page_config(page_title="G-Change｜Googleリスト整形＋NG除外", layout="wide")

# タイトル
st.markdown("""
    <h1 style='color: #800000;'>🚗 G-Change｜Googleリスト整形＋NGリスト自動除外</h1>
    <p>ファイルをアップロードし、クライアントNGリストを選択すると整形・除去できます。</p>
""", unsafe_allow_html=True)

# nglistsディレクトリの中を確認
NG_DIR = os.path.join(os.path.dirname(__file__), 'nglists')
nglist_files = []
if os.path.exists(NG_DIR):
    nglist_files = [f for f in os.listdir(NG_DIR) if f.endswith('.xlsx')]

# ファイルアップロード
uploaded_file = st.file_uploader("📤 整形対象のリストをアップロード", type=["xlsx"])

# クライアントNGリストの選択（プルダウン）
selected_nglist = st.selectbox("👥 クライアントNGリストを選択してください", ["なし"] + nglist_files)

# テンプレート使用フラグ
use_template = st.checkbox("🗂 テンプレートファイルとして処理します（入力マスターから抽出）", value=False)

# 抽出ルールキーワード
review_keywords = ["楽しい", "親切", "人柄", "感じ", "スタッフ", "雰囲気", "交流", "お世話", "ありがとう", "です", "ました", "🙇"]
ignore_keywords = ["ウェブサイト", "ルート", "営業中", "閉店", "口コミ"]

def normalize(text):
    text = str(text).strip().replace(" ", " ").replace("　", " ")
    return re.sub(r'[−–—―]', '-', text)

def extract_info(lines):
    company = normalize(lines[0]) if lines else ""
    industry, address, phone = "", "", ""

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
    return not any(kw in line for kw in ignore_keywords + review_keywords) and not re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line)

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, header=None)

    if use_template:
        # テンプレートファイルの場合
        df = df_raw.copy()
        result_df = pd.DataFrame({
            "企業名": df.iloc[1:, 1].dropna(),
            "業種": df.iloc[1:, 2].dropna(),
            "住所": df.iloc[1:, 3].dropna(),
            "電話番号": df.iloc[1:, 4].dropna()
        })
    else:
        # 通常の縦型リストの場合
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

    # NGリスト適用処理
    if selected_nglist != "なし":
        nglist_path = os.path.join(NG_DIR, selected_nglist)
        ng_df = pd.read_excel(nglist_path)

        ng_companies = ng_df['企業名'].dropna().astype(str).tolist()
        ng_phones = ng_df['電話番号'].dropna().astype(str).tolist()

        result_df['電話番号'] = result_df['電話番号'].astype(str)

        result_df = result_df[
            ~((result_df['企業名'].isin(ng_companies)) | (result_df['電話番号'].isin(ng_phones)))
        ]

    # 成形後の出力
    st.success(f"✅ 整形＆NG除外完了！企業数：{len(result_df)}件")
    st.dataframe(result_df, use_container_width=True)

    # ダウンロードボタン
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        result_df.to_excel(writer, index=False, sheet_name="整形済みデータ")
    st.download_button("📥 Excelファイルをダウンロード", data=output.getvalue(),
                       file_name="整形済み_企業リスト.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
