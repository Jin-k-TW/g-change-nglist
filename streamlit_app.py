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
st.title("🚗 G-Change｜Googleリスト整形＋NGリスト除外（GitHub直下NGリスト版）")

# --- 整形対象ファイルアップロード ---
uploaded_file = st.file_uploader("📤 整形対象のリストをアップロードしてください", type=["xlsx"])

# --- NGリストプルダウン（GitHub直下） ---
nglist_files = [f for f in os.listdir() if f.endswith(".xlsx") and "リスト" not in f and "template" not in f]
nglist_options = ["なし"] + [os.path.splitext(f)[0] for f in nglist_files]
nglist_choice = st.selectbox("👥 クライアントNGリストを選択してください", nglist_options)

# 正規化関数
def normalize(text):
    if pd.isna(text):
        return ""
    text = str(text).strip().replace(" ", " ").replace("　", " ")
    text = re.sub(r'[−–—―]', '-', text)
    return text

# 縦型リスト判定用
review_keywords = ["楽しい", "親切", "人柄", "感じ", "スタッフ", "雰囲気", "交流", "お世話", "ありがとう", "です", "ました", "🙇"]
ignore_keywords = ["ウェブサイト", "ルート", "営業中", "閉店", "口コミ"]

def extract_info(lines):
    company = normalize(lines[0]) if lines else ""
    industry, address, phone = "", "", ""
    for line in lines[1:]:
        line = normalize(line)
        if any(kw in line for kw in ignore_keywords + review_keywords):
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

# --- メイン処理 ---
if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file, header=None)

        # 縦型リスト判定
        if df_raw.shape[1] == 1:
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

            df = pd.DataFrame([extract_info(group) for group in groups],
                              columns=["企業名", "業種", "住所", "電話番号"])
        
        else:
            # 整形済みリスト判定
            df = pd.read_excel(uploaded_file)

            # 「企業様名称」列があればリネーム
            if "企業名" not in df.columns:
                if "企業様名称" in df.columns:
                    df.rename(columns={"企業様名称": "企業名"}, inplace=True)

            # 必要列が揃っているか確認
            required_cols = ["企業名", "業種", "住所", "電話番号"]
            if not all(col in df.columns for col in required_cols):
                st.error("❌ ファイル形式が正しくありません。（必要列：企業名・業種・住所・電話番号）")
                st.stop()

        st.success(f"✅ 整形完了！（企業数：{len(df)} 件）")

        # --- NGリスト除外処理 ---
        if nglist_choice != "なし":
            ng_file_path = nglist_choice + ".xlsx"
            ng_df = pd.read_excel(ng_file_path)

            ng_company_list = ng_df["企業名"].dropna().tolist() if "企業名" in ng_df.columns else []
            ng_phone_list = ng_df["電話番号"].dropna().tolist() if "電話番号" in ng_df.columns else []

            # 部分一致（企業名）
            remove_mask_company = df["企業名"].apply(
                lambda x: any(ng_word in str(x) for ng_word in ng_company_list)
            )

            # 完全一致（電話番号）
            remove_mask_phone = df["電話番号"].isin(ng_phone_list)

            remove_mask = remove_mask_company | remove_mask_phone
            removed_df = df[remove_mask]
            df = df[~remove_mask]

            st.success(f"🧹 NGリスト除外完了！（除外件数：{len(removed_df)} 件）")

            if not removed_df.empty:
                st.subheader("🚫 除外された企業一覧")
                st.dataframe(removed_df, use_container_width=True)

        # --- Excel保存 ---
        uploaded_filename = uploaded_file.name.replace(".xlsx", "")
        final_filename = uploaded_filename + "：リスト.xlsx"

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name="リスト")

        st.download_button(
            label="📥 Excelファイルをダウンロード",
            data=output.getvalue(),
            file_name=final_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ 処理中にエラーが発生しました：{e}")
