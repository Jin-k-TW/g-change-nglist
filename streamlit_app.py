import streamlit as st
import pandas as pd
import os
import re
import io

# ページ設定
st.set_page_config(page_title="G-Change Next｜企業情報整形＋NGリスト除外", layout="wide")

# タイトル＆スタイル
st.markdown("""
    <style>
    h1 { color: #800000; }
    </style>
""", unsafe_allow_html=True)
st.title("🚗 G-Change Next｜企業情報整形＋NGリスト除外（万能版）")

# ファイルアップロード
uploaded_file = st.file_uploader("📤 整形対象のExcelファイルをアップロード", type=["xlsx"])

# NGリスト選択肢（GitHub直下から探す）
nglist_files = [f for f in os.listdir() if f.endswith(".xlsx") and "リスト" not in f and "template" not in f]
nglist_options = ["なし"] + [os.path.splitext(f)[0] for f in nglist_files]
nglist_choice = st.selectbox("👥 使用するNGリストを選択してください", nglist_options)

# 正規化関数
def normalize(text):
    if pd.isna(text):
        return ""
    text = str(text).strip().replace(" ", " ").replace("　", " ")
    return re.sub(r'[−–—―]', '-', text)

# 縦型リストの情報抽出
def extract_info(lines):
    company = normalize(lines[0]) if lines else ""
    industry, address, phone = "", "", ""
    for line in lines[1:]:
        line = normalize(line)
        if "·" in line or "⋅" in line:
            parts = re.split(r"[·⋅]", line)
            industry = parts[-1].strip()
        elif re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line):
            phone = re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line).group()
        elif not address and any(x in line for x in ["丁目", "町", "番", "区", "−", "-"]):
            address = line
    return pd.Series([company, industry, address, phone])

# 縦型判定
def is_company_line(line):
    line = normalize(str(line))
    return not re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", line)

# メイン処理
if uploaded_file:
    # まずヘッダーなしで読む（縦型判定のため）
    df_raw = pd.read_excel(uploaded_file, header=None)

    if df_raw.shape[1] == 1:
        # --- 縦型リスト処理 ---
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
    else:
        # --- 複数列ファイル処理（入力マスターシート探し）---
        xls = pd.ExcelFile(uploaded_file)
        sheet_to_use = None

        for sheet in xls.sheet_names:
            if "入力マスター" in sheet:
                sheet_to_use = sheet
                break

        if not sheet_to_use:
            sheet_to_use = xls.sheet_names[0]  # 最初のシートを使う

        df = pd.read_excel(uploaded_file, sheet_name=sheet_to_use)

        # 列名補正（スペース削除）
        df.columns = [str(col).strip() for col in df.columns]

        # 必要列のリネーム
        rename_map = {"企業様名称": "企業名", "業種": "業種", "住所": "住所", "電話番号": "電話番号"}
        df.rename(columns=rename_map, inplace=True)

        # 必須列がなければ追加
        for col in ["企業名", "業種", "住所", "電話番号"]:
            if col not in df.columns:
                df[col] = ""

        # 必須列だけ残す
        df = df[["企業名", "業種", "住所", "電話番号"]]

    st.success(f"✅ 整形完了！（企業数：{len(df)} 件）")

    # --- NGリスト除外処理 ---
    if nglist_choice != "なし":
        ng_file_path = nglist_choice + ".xlsx"
        ng_df = pd.read_excel(ng_file_path)

        ng_companies = ng_df["企業名"].dropna().tolist() if "企業名" in ng_df.columns else []
        ng_phones = ng_df["電話番号"].dropna().tolist() if "電話番号" in ng_df.columns else []

        original_count = len(df)

        # 部分一致（企業名）＋ 完全一致（電話番号）
        mask_company = df["企業名"].apply(lambda x: any(ng in str(x) for ng in ng_companies))
        mask_phone = df["電話番号"].apply(lambda x: str(x) in [str(p) for p in ng_phones])

        removed_df = df[mask_company | mask_phone]
        df = df[~(mask_company | mask_phone)]

        removed_count = original_count - len(df)

        st.success(f"🧹 NGリスト除外完了！（除外件数：{removed_count} 件）")

        if not removed_df.empty:
            st.error("🚫 除外された企業一覧")
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
