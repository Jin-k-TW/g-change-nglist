import streamlit as st
import pandas as pd
import re
import io
import os

# ── ページ設定 ──
st.set_page_config(page_title="G-Change｜Googleリスト整形＋NG除外（Ver3.1 入力マスター対応）", layout="wide")

# ── タイトル＆スタイル ──
st.markdown("""
    <style>
    h1 { color: #800000; }
    </style>
""", unsafe_allow_html=True)
st.title("🚗 G-Change｜Googleリスト整形＋NGリスト除外（入力マスター対応）")

# ── ファイルアップロード ──
uploaded_file = st.file_uploader("📤 整形対象のExcelファイルをアップロードしてください", type=["xlsx"])

# ── NGリスト選択肢の取得（GitHub直下） ──
nglist_files = [
    f for f in os.listdir()
    if f.endswith(".xlsx")
    and f not in (uploaded_file.name if uploaded_file else [])  # アップロードファイルは除外
    and "template" not in f
]
nglist_options = ["なし"] + [os.path.splitext(f)[0] for f in nglist_files]
nglist_choice = st.selectbox("👥 クライアントNGリストを選択してください", nglist_options)

# ── 正規化関数 ──
def normalize(text):
    if pd.isna(text):
        return ""
    t = str(text).strip().replace(" ", " ").replace("　", " ")
    return re.sub(r"[−–—―]", "-", t)

# ── 縦型リスト抽出ヘルパ ──
def extract_info(lines):
    company = normalize(lines[0]) if lines else ""
    industry = address = phone = ""
    for line in lines[1:]:
        s = normalize(line)
        if "·" in s or "⋅" in s:
            industry = re.split(r"[·⋅]", s)[-1].strip()
        elif m := re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", s):
            phone = m.group()
        elif not address and any(tok in s for tok in ["丁目","町","番","区","−","-"]):
            address = s
    return pd.Series([company, industry, address, phone])

def is_company_line(line):
    s = normalize(line)
    return not re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", s)

# ── メイン処理 ──
if uploaded_file:
    # 1) シート選択：「入力マスター」があればそちら、なければ先頭
    xls = pd.ExcelFile(uploaded_file)
    sheet_to_use = next((s for s in xls.sheet_names if "入力マスター" in s), xls.sheet_names[0])

    # 2) まずヘッダーなしで読んでみる
    df_temp = pd.read_excel(uploaded_file, sheet_name=sheet_to_use, header=None)

    try:
        if df_temp.shape[1] == 1:
            # ── 縦型リスト整形 ──
            lines = df_temp[0].dropna().tolist()
            groups = []
            cur = []
            for ln in lines:
                s = normalize(ln)
                if is_company_line(s):
                    if cur:
                        groups.append(cur)
                    cur = [s]
                else:
                    cur.append(s)
            if cur:
                groups.append(cur)
            df = pd.DataFrame(
                [extract_info(g) for g in groups],
                columns=["企業名","業種","住所","電話番号"]
            )
        else:
            # ── 横型「入力マスター」形式 ──
            df_full = pd.read_excel(uploaded_file, sheet_name=sheet_to_use)
            # 列名の前後空白削除
            df_full.columns = [str(c).strip() for c in df_full.columns]
            # 「企業様名称」を「企業名」に
            df_full.rename(columns={"企業様名称":"企業名"}, inplace=True)
            # 必要な４列を最低限確保
            for col in ["企業名","業種","住所","電話番号"]:
                if col not in df_full:
                    df_full[col] = ""
            df = df_full[["企業名","業種","住所","電話番号"]]
    except Exception as e:
        st.error(f"❌ ファイル読み込みエラー：{e}")
        st.stop()

    # ── 重複削除（企業名・電話番号） ──
    df_before = len(df)
    df = df.drop_duplicates(subset=["企業名","電話番号"])
    dropped = df_before - len(df)

    st.success(f"✅ 整形完了！（企業数：{len(df)} 件, 重複削除：{dropped} 件）")

    # ── NGリスト除外 ──
    if nglist_choice != "なし":
        ng_path = nglist_choice + ".xlsx"
        ng_df = pd.read_excel(ng_path)
        # NG側のキー抽出
        ng_comp = ng_df.get("企業名", pd.Series()).dropna().astype(str).tolist()
        ng_tel  = ng_df.get("電話番号", pd.Series()).dropna().astype(str).tolist()
        # 部分一致／完全一致マスク
        mask_c = df["企業名"].astype(str).apply(lambda x: any(n in x for n in ng_comp))
        mask_t = df["電話番号"].astype(str).isin(ng_tel)
        removed_df = df[mask_c | mask_t]
        df = df[~(mask_c | mask_t)]
        st.success(f"🧹 NG除外完了！（除外件数：{len(removed_df)} 件）")
        if not removed_df.empty:
            st.error("🚫 除外された企業一覧")
            st.dataframe(removed_df, use_container_width=True)

    # ── 出力CSV生成 ──
    base = os.path.splitext(uploaded_file.name)[0]
    out_name = f"{base}：リスト.xlsx"
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as w:
        df.to_excel(w, index=False, sheet_name="リスト")
    st.download_button("📥 Excelをダウンロード", bio.getvalue(), file_name=out_name,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
