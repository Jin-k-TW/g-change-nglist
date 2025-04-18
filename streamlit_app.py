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
st.title("🚗 G-Change｜Googleリスト整形＋NGリスト除外（部分一致＋電話番号完全一致版）")

# ファイルアップロード
uploaded_file = st.file_uploader("📤 整形対象のリストをアップロード", type=["xlsx"])

# NGリスト読み込み（nglists/フォルダから自動読み込み）
nglist_folder = "nglists"
nglist_files = [f for f in os.listdir(nglist_folder) if f.endswith(".xlsx")]
nglist_options = ["なし"] + [os.path.splitext(f)[0] for f in nglist_files]
nglist_choice = st.selectbox("👥 クライアントNGリストを選択してください", nglist_options)

# 正規化関数
def normalize(text):
    if pd.isna(text):
        return ""
    text = str(text).strip()
    text = re.sub(r'[−–—―]', '-', text)
    return text

# アップロードされたら処理開始
if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, header=None)

    try:
        # 1列のみ → 縦型リストと判断
        lines = df_raw[0].dropna().tolist()

        groups = []
        current = []
        for line in lines:
            line = normalize(str(line))
            if not re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line):
                if current:
                    groups.append(current)
                current = [line]
            else:
                current.append(line)
        if current:
            groups.append(current)

        df = pd.DataFrame(groups, columns=["企業名", "業種", "住所", "電話番号"])
        
    except Exception:
        # 複数列あり → 整形済みと判断
        df = pd.read_excel(uploaded_file)

        # 必要な列チェック
        required_cols = ["企業名", "業種", "住所", "電話番号"]
        if not all(col in df.columns for col in required_cols):
            st.error("❌ ファイル形式が正しくありません。（必要列：企業名・業種・住所・電話番号）")
            st.stop()

    st.success(f"✅ 整形完了！（企業数：{len(df)} 件）")

    # --- NGリスト除外処理 ---
    if nglist_choice != "なし":
        ng_file_path = os.path.join(nglist_folder, nglist_choice + ".xlsx")
        ng_df = pd.read_excel(ng_file_path)

        # NGリストから取得（「企業名」列と「電話番号」列）
        ng_names = ng_df["企業名"].dropna().astype(str).tolist() if "企業名" in ng_df.columns else []
        ng_phones = ng_df["電話番号"].dropna().astype(str).tolist() if "電話番号" in ng_df.columns else []

        original_count = len(df)

        # 部分一致判定用の関数
        def is_ng_company(company_name):
            return any(ng_name in str(company_name) for ng_name in ng_names)

        # 電話番号一致判定用の関数
        def is_ng_phone(phone_number):
            return str(phone_number) in ng_phones

        # 除外対象抽出
        mask_remove = df["企業名"].apply(is_ng_company) | df["電話番号"].apply(is_ng_phone)

        removed_df = df[mask_remove]
        df = df[~mask_remove]

        removed_count = len(removed_df)

        st.success(f"🧹 NGリスト除外完了！（除外件数：{removed_count} 件）")

        if removed_count > 0:
            st.warning("🚫 除外された企業一覧")
            st.dataframe(removed_df, use_container_width=True)

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
