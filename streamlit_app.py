import streamlit as st
import pandas as pd
import re
import io
import os

# ------------------------------
# 1. ページ設定 & タイトル
# ------------------------------
st.set_page_config(page_title="G-Change｜企業情報整形＆NG除外ツール", layout="wide")
st.markdown("""
    <style>
    h1 { color: #800000; }
    </style>
""", unsafe_allow_html=True)
st.title("🚗 G-Change Next｜企業情報整形＆NG除外ツール（Ver3.9）")

# ------------------------------
# 2. ファイルアップロードと NG 選択
# ------------------------------
uploaded_file = st.file_uploader("📤 整形対象の Excel ファイルをアップロード", type="xlsx")
# GitHub 直下に置かれた .xlsx (template.xlsx / 出力ファイル を除く) を NG リスト候補として列挙
nglist_files = [
    f for f in os.listdir() if f.endswith(".xlsx")
    and f not in ("template.xlsx",)
    and not f.startswith("整形済み_")
]
nglist_options = ["なし"] + [os.path.splitext(f)[0] for f in nglist_files]
nglist_choice = st.selectbox("👥 クライアント NG リストを選択してください", nglist_options)

if not uploaded_file:
    st.info("まずは上のファイルアップローダーから Excel ファイルを選択してください。")
    st.stop()

# ------------------------------
# 3. 前処理：normalize / extract_info / is_company_line
# ------------------------------
def normalize(text):
    if pd.isna(text):
        return ""
    s = str(text).strip().replace(" ", " ").replace("　", " ")
    return re.sub(r"[−–—―]", "-", s)

def extract_info(lines):
    """縦型リストの 1 社分ブロックから (企業名, 業種, 住所, 電話番号) を抜き出す"""
    company = normalize(lines[0]) if lines else ""
    industry = address = phone = ""
    for line in lines[1:]:
        t = normalize(line)
        if "·" in t or "⋅" in t:
            industry = re.split(r"[·⋅]", t)[-1].strip()
        elif m := re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", t):
            phone = m.group()
        elif not address and any(x in t for x in ["丁目","町","番","区","-", "−"]):
            address = t
    return pd.Series([company, industry, address, phone])

def is_company_line(line):
    """レビューキーワード／電話番号を含まない行＝企業名行とみなす"""
    t = normalize(line)
    # 縦型展開時にレビュー文やリンク行を飛ばしたい場合は keywords を追加
    return not re.search(r"\d{2,4}-\d{2,4}-\d{3,4}", t)

# ------------------------------
# 4. シート自動検知 & 読み込み
# ------------------------------
xls = pd.ExcelFile(uploaded_file)
if "入力マスター" in xls.sheet_names:
    sheet = "入力マスター"
    st.info("✅ 「入力マスター」シートを検知しました。こちらを処理します。")
else:
    sheet = xls.sheet_names[0]
    # 数列チェックのためヘッダーなしでいったん読み込む
    tmp = pd.read_excel(xls, sheet_name=sheet, header=None)
    if tmp.shape[1] == 1:
        st.info("⚠️ 縦型リストを検知しました。展開処理を行います。")
    else:
        st.info("➡️ 横型リスト（既に整形済み、または入力マスター以外）をそのまま読み込みます。")

# 実際の DataFrame 読み込み
# └ 縦型か横型かで header=None／0 を切り替え
df0 = pd.read_excel(xls, sheet_name=sheet, header=None)
if df0.shape[1] == 1:
    # ── 縦型リスト展開
    lines = df0[0].dropna().tolist()
    groups = []
    cur = []
    for L in lines:
        L2 = normalize(L)
        if is_company_line(L2):
            if cur:
                groups.append(cur)
            cur = [L2]
        else:
            cur.append(L2)
    if cur:
        groups.append(cur)
    df = pd.DataFrame([extract_info(g) for g in groups],
                      columns=["企業名","業種","住所","電話番号"])
else:
    # ── 横型リスト or 入力マスター
    df_raw = pd.read_excel(xls, sheet_name=sheet)
    # 列名クリーニング
    df_raw.columns = [str(c).strip() for c in df_raw.columns]
    # 入力マスター固有の列名マップ
    if "企業様名称" in df_raw.columns and "企業名" not in df_raw.columns:
        df_raw.rename(columns={"企業様名称":"企業名"}, inplace=True)
    # 必要列がなければ自動追加（空文字）
    for c in ["企業名","業種","住所","電話番号"]:
        if c not in df_raw.columns:
            df_raw[c] = ""
    df = df_raw[["企業名","業種","住所","電話番号"]]

st.success(f"✅ 整形完了！（企業数：{len(df)} 件）")
st.dataframe(df, use_container_width=True)

# ------------------------------
# 5. NG リスト除外（部分一致：企業名／完全一致：電話番号）
# ------------------------------
if nglist_choice != "なし":
    path = f"{nglist_choice}.xlsx"
    if not os.path.exists(path):
        st.error(f"❌ NG リストファイルが見つかりません： {path}")
        st.stop()
    ng = pd.read_excel(path)
    # 「企業名」「電話番号」があれば拾う
    ng_comp = ng["企業名"].dropna().astype(str).tolist() if "企業名" in ng.columns else []
    ng_tel  = ng["電話番号"].dropna().astype(str).tolist() if "電話番号" in ng.columns else []
    N0 = len(df)
    mask_c = df["企業名"].astype(str).apply(lambda x: any(n in x for n in ng_comp))
    mask_t = df["電話番号"].astype(str).isin(ng_tel)
    removed = df[mask_c | mask_t]
    df = df[~(mask_c | mask_t)]
    st.success(f"🧹 NG 除外完了！（除外件数：{N0 - len(df)} 件）")
    if not removed.empty:
        st.error("🚫 除外された企業一覧")
        st.dataframe(removed, use_container_width=True)

# ------------------------------
# 6. 出力ダウンロード
# ------------------------------
fname = uploaded_file.name.replace(".xlsx","") + "：リスト.xlsx"
out = io.BytesIO()
with pd.ExcelWriter(out, engine="xlsxwriter") as w:
    df.to_excel(w, index=False, sheet_name="リスト")
st.download_button("📥 Excelファイルをダウンロード", data=out.getvalue(),
                   file_name=fname,
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
