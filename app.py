import streamlit as st
import gspread
import datetime

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

# ── パスワード認証 ────────────────────────────────────────────
PASSWORD = "haiso2026"

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🛢️ 灯油配送アプリ")
    st.markdown("---")
    st.subheader("🔒 ログイン")
    with st.form("login_form"):
        pw = st.text_input("パスワードを入力してください", type="password")
        login_btn = st.form_submit_button("ログイン", use_container_width=True)
    if login_btn:
        if pw == PASSWORD:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("❌ パスワードが違います")
    st.stop()

# ログアウトボタン（サイドバー）
with st.sidebar:
    st.write("🛢️ 灯油配送アプリ")
    if st.button("🔓 ログアウト", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()

# ── メインアプリ ──────────────────────────────────────────────
st.title("🛢️ 灯油配送アプリ")
st.write("名前・顧客コード・住所の一部を入れて検索できます。")

try:
    data = load_all_data()
    st.success(f"✅ 全シートから {len(data)}件 読み込み完了")
except Exception as e:
    st.error(f"❌ 読み込みエラー：{e}")
    st.stop()

tab1, tab2, tab3, tab4, tab5 = st.tabs(["🔎 名前検索", "🔢 顧客コード検索", "📍 住所検索", "📋 今日の配送リスト", "📷 日報読み取り"])

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

# ── 今日の配送リスト ──────────────────────────────────────────
with tab4:
    st.subheader("📋 今日の配送リスト")
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    st.caption(f"📅 配送日：{today_str}")

    area = st.selectbox("エリアを選択してください", SHEET_NAMES, key="delivery_area")
    area_data = [r for r in data if r.get("エリア") == area]

    if not area_data:
        st.warning("このエリアに顧客データがありません")
    else:
        st.info(f"📦 {len(area_data)} 件の顧客が見つかりました")

        with st.form("delivery_form"):
            records = []
            for i, row in enumerate(area_data):
                st.markdown(
                    f"**{row.get('名前', '---')}**　｜　"
                    f"顧客コード: `{row.get('顧客コード', '---')}`"
                )
                st.caption(f"📍 {row.get('住所', '---')}")

                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    visited = st.checkbox("✅ 訪問済み", key=f"visited_{i}")
                with col2:
                    supply = st.number_input(
                        "🛢 補給量 (L)",
                        min_value=0,
                        step=18,
                        key=f"supply_{i}"
                    )
                with col3:
                    absent = st.checkbox("🚪 不在", key=f"absent_{i}")
                with col4:
                    rental = st.checkbox("📄 レンタル伝票投函", key=f"rental_{i}")

                st.markdown("---")

                records.append({
                    "日付":           today_str,
                    "エリア":         area,
                    "顧客コード":     str(row.get("顧客コード", "")),
                    "名前":           str(row.get("名前", "")),
                    "住所":           str(row.get("住所", "")),
                    "訪問済み":       visited,
                    "補給量(L)":      supply,
                    "不在":           absent,
                    "レンタル伝票投函": rental,
                })

            submitted = st.form_submit_button(
                "💾 スプレッドシートに保存する",
                use_container_width=True
            )

        if submitted:
            try:
                client = connect_sheets()
                spreadsheet = client.open_by_url(SPREADSHEET_URL)

                # 「配送記録」シートを取得。なければ新規作成
                try:
                    record_sheet = spreadsheet.worksheet("配送記録")
                except Exception:
                    record_sheet = spreadsheet.add_worksheet(
                        title="配送記録", rows=1000, cols=10
                    )
                    # ヘッダー行を追加
                    record_sheet.append_row([
                        "日付", "エリア", "顧客コード", "名前", "住所",
                        "訪問済み", "補給量(L)", "不在", "レンタル伝票投函"
                    ])

                # 全顧客分をまとめて追記
                rows_to_append = []
                for rec in records:
                    rows_to_append.append([
                        rec["日付"],
                        rec["エリア"],
                        rec["顧客コード"],
                        rec["名前"],
                        rec["住所"],
                        "✓" if rec["訪問済み"] else "",
                        rec["補給量(L)"] if rec["補給量(L)"] > 0 else "",
                        "✓" if rec["不在"] else "",
                        "✓" if rec["レンタル伝票投函"] else "",
                    ])

                record_sheet.append_rows(rows_to_append)
                st.success(f"✅ {len(records)} 件の配送記録をスプレッドシートに保存しました！")

            except Exception as e:
                st.error(f"❌ 保存エラー：{e}")

# ── 日報読み取り（AI） ────────────────────────────────────────
with tab5:
    st.subheader("📷 日報読み取り（AI）")

    # APIキー確認
    anthropic_api_key = None
    try:
        anthropic_api_key = st.secrets["ANTHROPIC_API_KEY"]
    except Exception:
        pass

    if not anthropic_api_key:
        st.warning("⚠️ APIキーを設定してください（.streamlit/secrets.toml に ANTHROPIC_API_KEY を追加）")
    else:
        uploaded_file = st.file_uploader(
            "日報の写真またはPDFをアップロードしてください",
            type=["jpg", "jpeg", "png", "pdf"],
            key="daily_report_upload"
        )

        if uploaded_file is not None:
            if uploaded_file.type.startswith("image"):
                st.image(uploaded_file, caption="アップロードされた日報", use_container_width=True)
            else:
                st.info("📄 PDFがアップロードされました")

            if st.button("🤖 AIで読み取る", use_container_width=True, key="ai_read_btn"):
                with st.spinner("🔍 AIが日報を解析中...しばらくお待ちください"):
                    try:
                        import anthropic as _anthropic
                        import base64
                        import json
                        import re

                        client_ai = _anthropic.Anthropic(api_key=anthropic_api_key)

                        file_bytes = uploaded_file.getvalue()
                        file_b64 = base64.standard_b64encode(file_bytes).decode("utf-8")

                        prompt_text = """この日報画像から配送記録を読み取り、以下のJSON形式のみで返してください。
各顧客の行ごとに1つのオブジェクトを作成してください。

[
  {
    "顧客コード": "顧客コードの値",
    "名前": "顧客名",
    "補給量(L)": 数値（リットル数、0の場合は0）,
    "時間": "HH:MM形式、不明な場合は空文字",
    "レンタル伝票": "有" または "無"
  }
]

読み取りの注意事項：
- レンタル伝票は「オイル」欄の値が1.00（または1）の場合は「有」、それ以外は「無」にしてください
- 補給量はリットル数の数値のみ（単位なし、整数または小数）
- 読み取れない・不明な項目は空文字または0にしてください
- JSONの配列のみを返してください。前後の説明文・コードブロック記号は不要です"""

                        if uploaded_file.type == "application/pdf":
                            content = [
                                {
                                    "type": "document",
                                    "source": {
                                        "type": "base64",
                                        "media_type": "application/pdf",
                                        "data": file_b64,
                                    },
                                },
                                {"type": "text", "text": prompt_text},
                            ]
                        else:
                            content = [
                                {
                                    "type": "image",
                                    "source": {
                                        "type": "base64",
                                        "media_type": uploaded_file.type,
                                        "data": file_b64,
                                    },
                                },
                                {"type": "text", "text": prompt_text},
                            ]

                        message = client_ai.messages.create(
                            model="claude-opus-4-6",
                            max_tokens=3000,
                            messages=[{"role": "user", "content": content}],
                        )

                        response_text = message.content[0].text

                        # JSONを抽出（コードブロックや余分なテキストに対応）
                        json_match = re.search(r"\[.*?\]", response_text, re.DOTALL)
                        if json_match:
                            parsed_data = json.loads(json_match.group())
                            st.session_state.daily_report_data = parsed_data
                            st.success(f"✅ {len(parsed_data)}件のデータを読み取りました")
                        else:
                            st.error("❌ JSONデータの抽出に失敗しました。AIの返答を確認してください。")
                            with st.expander("AIの返答（デバッグ用）"):
                                st.text(response_text)

                    except Exception as e:
                        st.error(f"❌ AI読み取りエラー：{e}")

        # ── 読み取り結果の表示・編集・保存 ──────────────────────
        if "daily_report_data" in st.session_state and st.session_state.daily_report_data:
            st.markdown("---")
            st.subheader("📝 読み取り結果（編集可能）")
            st.caption("内容を確認・修正してから保存ボタンを押してください")

            import pandas as pd

            df = pd.DataFrame(st.session_state.daily_report_data)

            # 必要な列が存在しない場合は追加
            required_cols = ["顧客コード", "名前", "補給量(L)", "時間", "レンタル伝票"]
            for col in required_cols:
                if col not in df.columns:
                    df[col] = ""
            df = df[required_cols]

            # 型を整える
            df["補給量(L)"] = pd.to_numeric(df["補給量(L)"], errors="coerce").fillna(0)

            edited_df = st.data_editor(
                df,
                use_container_width=True,
                num_rows="dynamic",
                column_config={
                    "顧客コード":  st.column_config.TextColumn("顧客コード"),
                    "名前":        st.column_config.TextColumn("名前"),
                    "補給量(L)":   st.column_config.NumberColumn("補給量(L)", min_value=0, step=1),
                    "時間":        st.column_config.TextColumn("時間"),
                    "レンタル伝票": st.column_config.SelectboxColumn(
                        "レンタル伝票", options=["無", "有"]
                    ),
                },
                key="daily_report_editor",
            )

            today_str = datetime.date.today().strftime("%Y-%m-%d")

            if st.button("💾 スプレッドシートに保存", use_container_width=True, key="save_daily_report"):
                try:
                    client = connect_sheets()
                    spreadsheet = client.open_by_url(SPREADSHEET_URL)

                    # 「配送記録」シートを取得。なければ新規作成
                    try:
                        record_sheet = spreadsheet.worksheet("配送記録")
                    except Exception:
                        record_sheet = spreadsheet.add_worksheet(
                            title="配送記録", rows=1000, cols=10
                        )
                        record_sheet.append_row([
                            "日付", "エリア", "顧客コード", "名前", "住所",
                            "訪問済み", "補給量(L)", "不在", "レンタル伝票投函",
                        ])

                    rows_to_append = []
                    for _, row in edited_df.iterrows():
                        supply_val = row.get("補給量(L)", 0)
                        try:
                            supply_val = float(supply_val) if supply_val else 0
                        except (ValueError, TypeError):
                            supply_val = 0

                        rental_val = "✓" if str(row.get("レンタル伝票", "")).strip() == "有" else ""

                        rows_to_append.append([
                            today_str,
                            "",          # エリア（日報からは取得しない）
                            str(row.get("顧客コード", "")),
                            str(row.get("名前", "")),
                            "",          # 住所（日報からは取得しない）
                            "✓",         # 訪問済み
                            supply_val if supply_val > 0 else "",
                            "",          # 不在
                            rental_val,
                        ])

                    record_sheet.append_rows(rows_to_append)
                    st.success(f"✅ {len(rows_to_append)}件をスプレッドシートに保存しました！")
                    st.balloons()
                    # 保存後に読み取りデータをクリア
                    del st.session_state.daily_report_data

                except Exception as e:
                    st.error(f"❌ 保存エラー：{e}")
