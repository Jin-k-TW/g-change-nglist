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
st.title("🚗 G-Change｜Googleリスト整形＋NGリスト除外（部分一致対応版）")

# ファイルアップロード
uploaded_file = st.file_uploader("📤 整形対象のリストをアップロード", type=["xlsx"])

# NGリスト（リポジトリ直下から選択）
nglist_files = [f for f in os.listdir() if f.endswith(".xlsx") and "リスト" not in f and "template" not in f]
nglist_options = ["なし"] + [os.path.splitext(f)[0] for f in nglist_files]
nglist_choice = st.selectbox("👥 クライアントNGリストを選択してください", nglist_options)

# キーワード
review_keywords = ["楽しい", "親切", "人柄", "感じ", "スタッフ", "雰囲気", "交流", "お世話", "ありがとう", "です", "ました", "🙇"]
ignore_keywords = ["ウェブサイト", "ルート", "営業中", "閉店", "口コミ"]

def normalize(text):
    if pd.isna(text):
        return ""
    text = str(text).strip().replace(" ", " ").replace("　", " ")
    text = re.sub(r'[−–—―]', '-', text)
    return text

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

# 部分一致用クリーニング関数
def clean_text(text):
    if pd.isna(text):
        return ""
    return str(text).strip().replace(" ", "").replace("　", "").lower()

def clean_phone(phone):
    if pd.isna(phone):
        return ""
    return re.sub(r"[-ー－−]", "", str(phone))

def is_partial_match(value, series):
    if pd.isna(value):
        return False
    value = str(value).strip()
    return series.dropna().astype(str).str.contains(re.escape(value), na=False).any()

# --- メイン処理 ---
if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, header=None)

    try:
        # 縦型リスト判定
        lines = df_raw.iloc[:, 0].dropna().tolist()

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

        df = pd.DataFrame([extract_info(group) for group in groups],
                          columns=["企業名", "業種", "住所", "電話番号"])

    except Exception:
        # 整形済みリスト判定
        df = pd.read_excel(uploaded_file)

        # 企業様名称の変換
        rename_map = {}
        if "企業様名称" in df.columns:
            rename_map["企業様名称"] = "企業名"
        if rename_map:
            df.rename(columns=rename_map, inplace=True)

        required_cols = ["企業名", "業種", "住所", "電話番号"]
        if not all(col in df.columns for col in required_cols):
            st.error("❌ ファイル形式が正しくありません。（必要列：企業名・業種・住所・電話番号）")
            st.stop()

    st.success(f"✅ 整形完了！（企業数：{len(df)} 件）")

    # --- NGリスト除外処理 ---
    if nglist_choice != "なし":
        ng_file_path = nglist_choice + ".xlsx"
        ng_df = pd.read_excel(ng_file_path)

        # 企業様名称の変換
        rename_map_ng = {}
        if "企業様名称" in ng_df.columns:
            rename_map_ng["企業様名称"] = "企業名"
        if rename_map_ng:
            ng_df.rename(columns=rename_map_ng, inplace=True)

        # クリーニング（比較用）
        df["企業名_clean"] = df["企業名"].apply(clean_text)
        df["電話番号_clean"] = df["電話番号"].apply(clean_phone)
        ng_df["企業名_clean"] = ng_df["企業名"].apply(clean_text)
        ng_df["電話番号_clean"] = ng_df["電話番号"].apply(clean_phone)

        # 部分一致マッチング
        ng_companies = ng_df["企業名_clean"].dropna().unique().tolist()
        ng_phones = ng_df["電話番号_clean"].dropna().unique().tolist()

        mask_exclude = df["企業名_clean"].apply(lambda x: any(part in x for part in ng_companies)) | \
                       df["電話番号_clean"].apply(lambda x: any(part in x for part in ng_phones))

        df_excluded = df[mask_exclude]
        df = df[~mask_exclude]

        removed_count = len(df_excluded)
        st.success(f"🧹 NGリスト除外完了！（除外件数：{removed_count} 件）")

        if removed_count > 0:
            st.subheader("🚫 除外された企業一覧")
            st.dataframe(df_excluded[["企業名", "住所", "電話番号"]])

    # 出力ファイル名設定
    uploaded_filename = uploaded_file.name.replace(".xlsx", "")
    final_filename = uploaded_filename + "：リスト.xlsx"

    # Excel保存
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.drop(columns=["企業名_clean", "電話番号_clean"], errors='ignore').to_excel(writer, index=False, sheet_name="リスト")

    st.download_button(
        label="📥 Excelファイルをダウンロード",
        data=output.getvalue(),
        file_name=final_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
