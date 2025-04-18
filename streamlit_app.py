import streamlit as st
import pandas as pd
import re
import io
import os

# ページ設定
st.set_page_config(page_title="G-Change｜Googleリスト整形＋NGリスト除外", layout="wide")

# タイトル＆スタイル
st.markdown("""
    <style>
    h1 { color: #800000; }
    </style>
""", unsafe_allow_html=True)
st.title("🚗 G-Change｜Googleリスト整形＋NGリスト除外")

# ファイルアップロード
uploaded_file = st.file_uploader("📤 整形対象のリストをアップロード", type=["xlsx"])

# NGリスト読み込み（nglists/フォルダから自動読み込み）
nglist_folder = "nglists"
nglist_files = [f for f in os.listdir(nglist_folder) if f.endswith(".xlsx")]

nglist_options = ["なし"] + [os.path.splitext(f)[0] for f in nglist_files]
nglist_choice = st.selectbox("👥 クライアントNGリストを選択してください", nglist_options)

# 抽出ルール用のキーワード
review_keywords = ["楽しい", "親切", "人柄", "感じ", "スタッフ", "雰囲気", "交流", "お世話", "ありがとう", "です", "ました", "🙇"]
ignore_keywords = ["ウェブサイト", "ルート", "営業中", "閉店", "口コミ"]

# 正規化関数
def normalize(text):
    if pd.isna(text):
        return ""
    text = str(text).strip().replace(" ", " ").replace("　", " ")
    text = re.sub(r'[−–—―]', '-', text)
    return text

# 企業情報抽出
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

# 企業名の行判定
def is_company_line(line):
    line = normalize(line)
    return not any(kw in line for kw in ignore_keywords + review_keywords) and not re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line)

# メイン処理
if uploaded_file:
    # 読み込み
    df_raw = pd.read_excel(uploaded_file, header=None)

    try:
        # 1列しかない＝縦型リストと判断
        lines = df_raw[0].dropna().tolist()

        groups = []
        current = []
        for line in lines:
            line = normalize(line)
            if is_company_line(line):
                if current:
                    groups.append(current)
                current = [line]
            else:
                current.append(line)
        if current:
            groups.append(current)

        df = pd.DataFrame([extract_info(group) for group in groups],
                          columns=["企業名", "業種", "住所", "電話番号"])
        
    except Exception:
        # 複数列あり＝すでに整形済みリストと判断
        df = pd.read_excel(uploaded_file)

        if "企業名" not in df.columns or "業種" not in df.columns or "住所" not in df.columns or "電話番号" not in df.columns:
            st.error("❌ ファイル形式が正しくありません。（必要な列：企業名、業種、住所、電話番号）")
            st.stop()

    st.success(f"✅ 整形完了！（企業数：{len(df)} 件）")

    # --- NGリスト除外処理 ---
    if nglist_choice != "なし":
        ng_file_path = os.path.join(nglist_folder, nglist_choice + ".xlsx")
        ng_df = pd.read_excel(ng_file_path)

        ng_companies = ng_df["企業名"].dropna().unique().tolist() if "企業名" in ng_df.columns else []
        ng_phones = ng_df["電話番号"].dropna().unique().tolist() if "電話番号" in ng_df.columns else []

        original_count = len(df)

        df = df[~(df["企業名"].isin(ng_companies) | df["電話番号"].isin(ng_phones))]

        removed_count = original_count - len(df)
        st.success(f"🧹 NGリスト除外完了！（除外件数：{removed_count} 件）")

    # 出力ファイル名設定
    uploaded_filename = uploaded_file.name.replace(".xlsx", "")
    final_filename = uploaded_filename + "：リスト.xlsx"

    # Excel保存
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="リスト")

    st.download_button(
        label="📥 Excelファイルをダウンロード",
        data=output.getvalue(),
        file_name=final_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
