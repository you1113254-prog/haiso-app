import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1VPO7xDMz_HPXyWuy8YVhHQQvmrmfGFG5bSuRDdtQ3DM"

@st.cache_resource
def connect_sheets():
    scope = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
    return gspread.authorize(creds)

@st.cache_data(ttl=60)
def load_all_data():
    client = connect_sheets()
    spreadsheet = client.open_by_url(SPREADSHEET_URL)
    all_data = []
    skip = ["金武宜野座コメント","1月北部","1月北部 のリスト","1月北部 のリスト のコピー"]
    for sheet in spreadsheet.worksheets():
        if sheet.title in skip:
            continue
        try:
            records = sheet.get_all_records()
            for row in records:
                row["エリア"] = sheet.title
            all_data.extend(records)
        except Exception as e:
            st.warning(f"⚠️ {sheet.title} スキップ：{e}")
    return all_data

st.set_page_config(page_title="灯油配送アプリ", page_icon="🛢️", layout="centered")
st.title("🛢️ 灯油配送アプリ")
st.write("名前・顧客コード・住所の一部を入れて検索できます。")

try:
    data = load_all_data()
    st.success(f"✅ 全シートから {len(data)}件 読み込み完了")
except Exception as e:
    st.error(f"❌ 読み込みエラー：{e}")
    st.stop()

tab1, tab2, tab3 = st.tabs(["🔎 名前検索", "🔢 顧客コード検索", "📍 住所検索"])

def show_results(results):
    if results:
        st.success(f"{len(results)}件見つかりました")
        for row in results:
            st.write(f"**顧客コード:** {row.get('顧客コード','---')}")
            st.write(f"**名前:** {row.get('名前','---')}")
            st.write(f"**住所:** {row.get('住所','---')}")
            st.write(f"**エリア:** {row.get('エリア','---')}")
            st.markdown("---")
    else:
        st.warning("該当なし")

with tab1:
    k = st.text_input("名前の一部を入力", key="name")
    if k:
        show_results([r for r in data if k in str(r.get("名前",""))])

with tab2:
    k = st.text_input("顧客コードの一部を入力", key="code")
    if k:
        show_results([r for r in data if k in str(r.get("顧客コード",""))])

with tab3:
    k = st.text_input("住所の一部を入力", key="addr")
    if k:
        show_results([r for r in data if k in str(r.get("住所",""))])
