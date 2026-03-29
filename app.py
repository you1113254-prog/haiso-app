import streamlit as st
import gspread

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1VPO7xDMz_HPXyWuy8YVhHQQvmrmfGFG5bSuRDdtQ3DM"

SHEET_NAMES = [
    "宜野座 と金武1～3",
    "恩納村",
    "石川1 ～4",
    "読谷",
    "うるま",
    "本部、今帰仁",
    "勝連",
    "沖縄市",
]

@st.cache_resource
def connect_sheets():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    from google.oauth2.service_account import Credentials
    scope = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(creds)

@st.cache_data(ttl=60)
def load_all_data():
    client = connect_sheets()
    spreadsheet = client.open_by_url(SPREADSHEET_URL)
    all_data = []
    for sheet_name in SHEET_NAMES:
        try:
            sheet = spreadsheet.worksheet(sheet_name)
            records = sheet.get_all_records()
            for row in records:
                # 空白行を除外（顧客コードが空のものをスキップ）
                if not row.get("顧客コード") and not row.get("名前"):
                    continue
                row["エリア"] = sheet_name
                all_data.append(row)
        except Exception as e:
            st.warning(f"⚠️ {sheet_name} スキップ：{e}")
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
