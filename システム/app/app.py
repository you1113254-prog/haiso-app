import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

# ── 設定 ──
CREDENTIALS_FILE = "/Users/youichi/Desktop/配送AI/credentials.json"
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1IX29ymdpGmrfYgoUxT64FjS21OADX3F8"

# ── Google Sheets 接続 ──
@st.cache_resource
def connect_sheets():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scope)
    client = gspread.authorize(creds)
    return client

# ── データ読み込み ──
@st.cache_data(ttl=60)
def load_data():
    client = connect_sheets()
    sheet = client.open_by_url(SPREADSHEET_URL).sheet1
    records = sheet.get_all_records()
    return records

# ── 画面設定 ──
st.set_page_config(page_title="顧客検索", page_icon="🔎", layout="centered")
st.title("顧客検索アプリ")
st.write("名前・顧客コード・住所の一部を入れて検索できます。")

# ── データ取得 ──
try:
    data = load_data()
    st.success(f"✅ スプレッドシートから {len(data)}件 読み込み完了")
except Exception as e:
    st.error(f"❌ 読み込みエラー：{e}")
    st.stop()

# ── 検索タブ ──
tab1, tab2, tab3 = st.tabs(["🔎 名前検索", "🔢 顧客コード検索", "📍 住所検索"])

def show_results(results):
    if results:
        st.success(f"{len(results)}件見つかりました")
        for row in results:
            st.write(f"**顧客コード:** {row.get('顧客コード', '---')}")
            st.write(f"**名前:** {row.get('名前', '---')}")
            st.write(f"**住所:** {row.get('住所', '---')}")
            st.markdown("---")
    else:
        st.warning("該当なし")

with tab1:
    keyword_name = st.text_input("名前の一部を入力", key="name")
    if keyword_name:
        results = [r for r in data if keyword_name in str(r.get("名前", ""))]
        show_results(results)

with tab2:
    keyword_code = st.text_input("顧客コードの一部を入力", key="code")
    if keyword_code:
        results = [r for r in data if keyword_code in str(r.get("顧客コード", ""))]
        show_results(results)

with tab3:
    keyword_addr = st.text_input("住所の一部を入力", key="addr")
    if keyword_addr:
        results = [r for r in data if keyword_addr in str(r.get("住所", ""))]
        show_results(results)