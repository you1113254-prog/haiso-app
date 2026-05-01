import streamlit as st
import anthropic
import gspread
from google.oauth2.service_account import Credentials
import re
import os
from dotenv import load_dotenv
from collections import defaultdict

load_dotenv("/Users/youichi/Desktop/MYAI/.env")

# 新しいスプレッドシートID
SPREADSHEET_ID = "1gX2Iyg7bODxuKZI1ENwMbmmPzTW7KjaWNwjxAQnaWM4"

# 読み込むシート（配送記録・現金フリーは除外）
TARGET_SHEETS = [
    "ぎのざきん",
    "いしかわ",
    "よみたん",
    "うるま",
    "もとぶなきじん",
    "かつれん",
    "おきなわし",
    "おんなそん",
    "なご",
    "くにがみひがしおおぎみ",
    "やがじまきやいさがわ",
    "うむさやぶびいまた",
    "へのこおおうら",
]

st.set_page_config(
    page_title="灯油配送アシスタント",
    page_icon="🛢️",
    layout="centered"
)

@st.cache_resource
def get_spreadsheet():
    if "gcp_service_account" in st.secrets:
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=[
                "https://spreadsheets.google.com/feeds",
                "https://www.googleapis.com/auth/drive"
            ]
        )
    else:
        creds = Credentials.from_service_account_file(
            "/Users/youichi/Desktop/配送AI/credentials.json",
            scopes=[
                "https://spreadsheets.google.com/feeds",
                "https://www.googleapis.com/auth/drive"
            ]
        )
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)

def is_valid_customer_code(code):
    code = str(code).strip()
    if not code:
        return False
    if re.match(r'^\d+$', code):
        return True
    if re.match(r'^\d+[\s\u3000]+[RrＲｒ]$', code):
        return True
    return False

@st.cache_data(ttl=300)
def get_all_customers():
    try:
        spreadsheet = get_spreadsheet()
        all_customers = []
        for sheet_name in TARGET_SHEETS:
            try:
                sheet = spreadsheet.worksheet(sheet_name)
                values = sheet.get_all_values()
                if len(values) < 2:
                    continue
                headers = values[0]
                code_col = None
                for i, h in enumerate(headers):
                    if "顧客コード" in str(h):
                        code_col = i
                        break
                for row_values in values[1:]:
                    if code_col is not None:
                        code = row_values[code_col] if code_col < len(row_values) else ""
                        if not is_valid_customer_code(code):
                            continue
                    else:
                        if not any(v.strip() for v in row_values if isinstance(v, str)):
                            continue
                    row = dict(zip(headers, row_values))
                    row["エリア"] = sheet_name
                    all_customers.append(row)
            except Exception:
                continue
        return all_customers
    except Exception:
        return []

def build_customer_context(customers):
    if not customers:
        return "（お客様データを読み込めませんでした）"

    area_counts = defaultdict(int)
    for c in customers:
        area_counts[c.get("エリア", "不明")] += 1

    lines = [f"【お客様データ：全{len(customers)}件】"]
    lines.append("")
    lines.append("■ エリア別件数")
    for area in TARGET_SHEETS:
        if area in area_counts:
            lines.append(f"  {area}：{area_counts[area]}件")
    lines.append(f"  合計：{len(customers)}件")
    lines.append("")

    lines.append("■ お客様一覧（名前・住所・エリア）")
    name_col = None
    addr_col = None
    for c in customers[:1]:
        for k in c.keys():
            if "名前" in str(k) or "氏名" in str(k):
                name_col = k
            if "住所" in str(k):
                addr_col = k

    for c in customers:
        name = c.get(name_col, "") if name_col else ""
        addr = c.get(addr_col, "") if addr_col else ""
        area = c.get("エリア", "")
        code = c.get("顧客コード", "")
        if name or addr:
            lines.append(f"  [{code}] {name} / {addr} / {area}")

    return "\n".join(lines)

def build_system_prompt(customer_context):
    return f"""
あなたはゆういちさんの灯油配送業務を助けるAIアシスタントです。
以下の業務知識とお客様データをもとに、配送作業中のスマホからの質問に答えてください。
回答は短く・わかりやすく・箇条書きで答えてください。

【基本情報】
- 担当エリア：沖縄県（うるま市、読谷村、名護市、恩納村など）
- お客様数：約400件
- 訪問頻度：月1回（燃料残量を確認して補充）
- 一人で全件まわっている

【灯油の価格・単価】
- 現在の単価：1リットルあたり約○○円
- 配送最低量：18リットル（1缶）

【地区ごとの特徴】
- うるま：件数が多い。市街地と農村が混在
- よみたん：道が細い場所あり
- なご：距離が遠い。まとめてまわる
- おんなそん：リゾート地帯

【繁忙期パターン】
- 12月〜2月：最繁忙期
- 3月・11月：中繁忙期
- 4月〜10月：閑散期

【燃料残量の目安】
- 残量20%以下：至急訪問
- 残量20〜40%：今月中に訪問
- 残量40%以上：来月でOK

{customer_context}
"""

st.title("🛢️ 灯油配送アシスタント")
st.caption("配送中の質問は何でも聞いてください")

customers = get_all_customers()
if customers:
    st.success(f"✅ お客様データ読み込み済み：{len(customers)}件")
else:
    st.warning("⚠️ お客様データを読み込めませんでした")

st.markdown("#### よく使う質問")
col1, col2, col3 = st.columns(3)
shortcut = None
with col1:
    if st.button("📍 今日のルート相談"):
        shortcut = "今日のルートを効率よくまわる順番を教えてください"
with col2:
    if st.button("🔋 残量少ない地区は？"):
        shortcut = "残量が少なくて至急訪問が必要なお客様はどう判断すればいい？"
with col3:
    if st.button("💰 売上計算"):
        shortcut = "今日の配送量から売上を計算する方法を教えてください"

if "messages" not in st.session_state:
    st.session_state.messages = []

if shortcut:
    st.session_state.messages.append({"role": "user", "content": shortcut})

for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

if prompt := st.chat_input("質問を入力してください..."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

if st.session_state.messages and st.session_state.messages[-1]["role"] == "user":
    with st.chat_message("assistant"):
        with st.spinner("考え中..."):
            api_key = st.secrets.get("ANTHROPIC_API_KEY", os.getenv("ANTHROPIC_API_KEY"))
            client = anthropic.Anthropic(api_key=api_key)
            customer_context = build_customer_context(customers)
            response = client.messages.create(
                model="claude-opus-4-5",
                max_tokens=1024,
                system=build_system_prompt(customer_context),
                messages=st.session_state.messages
            )
            reply = response.content[0].text
            st.markdown(reply)
            st.session_state.messages.append({"role": "assistant", "content": reply})

if st.button("🔄 会話をリセット"):
    st.session_state.messages = []
    st.rerun()