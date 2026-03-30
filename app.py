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

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "🔎 名前検索", "🔢 顧客コード検索", "📍 住所検索",
    "📋 今日の配送リスト", "📝 配送記録入力",
    "⚠️ アラート", "👤 顧客管理",
])

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

# ── 配送記録入力 ──────────────────────────────────────────────
with tab5:
    st.subheader("📝 配送記録入力")

    # 日付選択（デフォルトは今日）
    selected_date = st.date_input("📅 配送日を選択", value=datetime.date.today(), key="record_date")
    record_date_str = selected_date.strftime("%Y-%m-%d")
    st.caption(f"配送日：{record_date_str}")

    # エリア選択
    record_area = st.selectbox("エリアを選択してください", SHEET_NAMES, key="record_area")
    record_area_data = [r for r in data if r.get("エリア") == record_area]

    if not record_area_data:
        st.warning("このエリアに顧客データがありません")
    else:
        st.info(f"📦 {len(record_area_data)} 件の顧客が見つかりました")

        with st.form("record_input_form"):
            record_entries = []
            for i, row in enumerate(record_area_data):
                st.markdown(
                    f"**{row.get('名前', '---')}**　｜　"
                    f"顧客コード: `{row.get('顧客コード', '---')}`"
                )
                st.caption(f"📍 {row.get('住所', '---')}")

                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    rec_visited = st.checkbox("✅ 訪問済み", key=f"rec_visited_{i}")
                with col2:
                    rec_supply = st.number_input(
                        "🛢 補給量 (L)",
                        min_value=0,
                        step=18,
                        key=f"rec_supply_{i}"
                    )
                with col3:
                    rec_absent = st.checkbox("🚪 不在", key=f"rec_absent_{i}")
                with col4:
                    rec_rental = st.checkbox("📄 レンタル伝票投函", key=f"rec_rental_{i}")

                st.markdown("---")

                record_entries.append({
                    "日付":            record_date_str,
                    "エリア":          record_area,
                    "顧客コード":      str(row.get("顧客コード", "")),
                    "名前":            str(row.get("名前", "")),
                    "住所":            str(row.get("住所", "")),
                    "訪問済み":        rec_visited,
                    "補給量(L)":       rec_supply,
                    "不在":            rec_absent,
                    "レンタル伝票投函": rec_rental,
                })

            rec_submitted = st.form_submit_button(
                "💾 スプレッドシートに保存する",
                use_container_width=True
            )

        if rec_submitted:
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
                        "訪問済み", "補給量(L)", "不在", "レンタル伝票投函"
                    ])

                # 全顧客分をまとめて追記
                rows_to_append = []
                for rec in record_entries:
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
                st.success(f"✅ {len(record_entries)} 件の配送記録をスプレッドシートに保存しました！")

            except Exception as e:
                st.error(f"❌ 保存エラー：{e}")

# ── アラート ──────────────────────────────────────────────────
with tab6:
    st.subheader("⚠️ 未訪問・伝票漏れアラート")

    alert_date = st.date_input(
        "📅 確認する日付を選択", value=datetime.date.today(), key="alert_date"
    )
    alert_date_str = alert_date.strftime("%Y-%m-%d")

    if st.button("🔍 アラートを確認する", use_container_width=True, key="check_alerts"):
        try:
            client = connect_sheets()
            spreadsheet = client.open_by_url(SPREADSHEET_URL)

            try:
                record_sheet = spreadsheet.worksheet("配送記録")
                all_records = record_sheet.get_all_records()
            except Exception:
                st.warning("⚠️ 配送記録シートが見つかりません。先に配送記録を保存してください。")
                st.stop()

            # 選択した日付の記録だけ絞り込む
            day_records = [
                r for r in all_records
                if str(r.get("日付", "")).strip() == alert_date_str
            ]

            if not day_records:
                st.info(f"📭 {alert_date_str} の配送記録はありません")
            else:
                st.info(f"📋 {alert_date_str} の記録：{len(day_records)} 件")
                st.markdown("---")

                # 未訪問（訪問済みが ✓ でない）
                unvisited = [
                    r for r in day_records
                    if str(r.get("訪問済み", "")).strip() != "✓"
                ]

                # レンタル顧客（顧客コード末尾がR）で伝票投函が ✓ でない
                rental_missed = [
                    r for r in day_records
                    if str(r.get("顧客コード", "")).strip().upper().endswith("R")
                    and str(r.get("レンタル伝票投函", "")).strip() != "✓"
                ]

                # ── 未訪問リスト ──
                st.markdown("### 🚶 今日まだ行っていない顧客")
                if unvisited:
                    for r in unvisited:
                        st.warning(
                            f"**{r.get('名前', '---')}**　｜　"
                            f"顧客コード: `{r.get('顧客コード', '---')}`　｜　"
                            f"エリア: {r.get('エリア', '---')}"
                        )
                else:
                    st.success("✅ 全員訪問済みです！")

                st.markdown("---")

                # ── 伝票投函漏れリスト ──
                st.markdown("### 📄 伝票投函漏れの顧客（レンタル）")
                if rental_missed:
                    for r in rental_missed:
                        st.error(
                            f"**{r.get('名前', '---')}**　｜　"
                            f"顧客コード: `{r.get('顧客コード', '---')}`　｜　"
                            f"エリア: {r.get('エリア', '---')}"
                        )
                else:
                    st.success("✅ 伝票投函漏れはありません！")

        except Exception as e:
            st.error(f"❌ 読み込みエラー：{e}")

# ── 顧客管理 ──────────────────────────────────────────────────
with tab7:
    st.subheader("👤 顧客管理")

    mgmt_area = st.selectbox("エリアを選択してください", SHEET_NAMES, key="mgmt_area")
    mgmt_area_data = [r for r in data if r.get("エリア") == mgmt_area]
    st.info(f"📋 {len(mgmt_area_data)} 件の顧客が登録されています")

    # ── 新規顧客追加 ──────────────────────────────────────────
    with st.expander("➕ 新規顧客を追加する"):
        with st.form("add_customer_form"):
            new_code = st.text_input("顧客コード *", key="new_code")
            new_name = st.text_input("名前 *", key="new_name")
            new_addr = st.text_input("住所", key="new_addr")
            add_btn = st.form_submit_button("追加する", use_container_width=True)

        if add_btn:
            if not new_code.strip() or not new_name.strip():
                st.error("❌ 顧客コードと名前は必須です")
            else:
                try:
                    client = connect_sheets()
                    spreadsheet = client.open_by_url(SPREADSHEET_URL)
                    sheet = spreadsheet.worksheet(mgmt_area)

                    # ヘッダーを確認して列順に合わせて追加
                    headers = sheet.row_values(1)
                    new_row = [""] * max(len(headers), 3)
                    for col_name, val in [
                        ("顧客コード", new_code.strip()),
                        ("名前", new_name.strip()),
                        ("住所", new_addr.strip()),
                    ]:
                        if col_name in headers:
                            new_row[headers.index(col_name)] = val
                    sheet.append_row(new_row)

                    st.success(f"✅ {new_name.strip()} を追加しました！")
                    load_all_data.clear()
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 追加エラー：{e}")

    st.markdown("---")
    st.subheader("✏️ 既存顧客の編集")

    if not mgmt_area_data:
        st.warning("このエリアに顧客データがありません")
    else:
        for i, row in enumerate(mgmt_area_data):
            code = str(row.get("顧客コード", "")).strip()
            name = str(row.get("名前", ""))
            with st.expander(f"**{name}**　（コード: {code}）"):
                with st.form(f"edit_form_{i}"):
                    edit_name = st.text_input("名前", value=name, key=f"edit_name_{i}")
                    edit_addr = st.text_input("住所", value=str(row.get("住所", "")), key=f"edit_addr_{i}")
                    save_btn = st.form_submit_button("💾 この顧客を保存", use_container_width=True)

                if save_btn:
                    try:
                        client = connect_sheets()
                        spreadsheet = client.open_by_url(SPREADSHEET_URL)
                        sheet = spreadsheet.worksheet(mgmt_area)

                        # 全データを取得して顧客コードで行番号を特定
                        all_values = sheet.get_all_values()
                        headers = all_values[0] if all_values else []
                        code_col = headers.index("顧客コード") if "顧客コード" in headers else 0
                        name_col = headers.index("名前")      if "名前"      in headers else 1
                        addr_col = headers.index("住所")      if "住所"      in headers else 2

                        target_row = None
                        for row_idx, row_vals in enumerate(all_values[1:], start=2):
                            if len(row_vals) > code_col and str(row_vals[code_col]).strip() == code:
                                target_row = row_idx
                                break

                        if target_row:
                            sheet.update_cell(target_row, name_col + 1, edit_name)
                            sheet.update_cell(target_row, addr_col + 1, edit_addr)
                            st.success(f"✅ {edit_name} の情報を更新しました！")
                            load_all_data.clear()
                            st.rerun()
                        else:
                            st.error("❌ 顧客コードがシートで見つかりませんでした")

                    except Exception as e:
                        st.error(f"❌ 更新エラー：{e}")
