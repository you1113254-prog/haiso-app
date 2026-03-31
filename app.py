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

@st.cache_data(ttl=60)
def load_delivery_records():
    """配送記録シートを全件取得。シートが存在しない場合は空リストを返す。"""
    try:
        client = connect_sheets()
        spreadsheet = client.open_by_url(SPREADSHEET_URL)
        sheet = spreadsheet.worksheet("配送記録")
        return sheet.get_all_records()
    except Exception:
        return []

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

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
    "🔎 名前検索", "🔢 顧客コード検索", "📍 住所検索",
    "📋 今日の配送リスト", "📝 配送記録入力",
    "⚠️ アラート", "👤 顧客管理", "📊 日報",
])

def show_results(results):
    if not results:
        st.warning("該当なし")
        return

    st.success(f"{len(results)}件見つかりました")
    all_delivery = load_delivery_records()   # キャッシュ済みなので高速

    for row in results:
        code = str(row.get("顧客コード", "")).strip()
        st.write(f"**顧客コード:** {code or '---'}")
        st.write(f"**名前:** {row.get('名前','---')}")
        st.write(f"**住所:** {row.get('住所','---')}")
        st.write(f"**エリア:** {row.get('エリア','---')}")

        # ── 配送履歴（この顧客の全記録） ─────────────────────────
        history = [
            r for r in all_delivery
            if code and str(r.get("顧客コード", "")).strip() == code
        ]

        # ── 最終補給情報 ──────────────────────────────────────
        supply_records = []
        for _h in history:
            try:
                _sp = float(_h.get("補給量(L)", 0) or 0)
            except (ValueError, TypeError):
                _sp = 0.0
            if _sp > 0:
                supply_records.append((str(_h.get("日付", "")), _sp))

        if supply_records:
            supply_records.sort(key=lambda x: x[0], reverse=True)   # 新しい順
            _last_date_str, _last_supply = supply_records[0]
            try:
                _last_date = datetime.date.fromisoformat(_last_date_str)
                _days_ago  = (datetime.date.today() - _last_date).days
                _date_jp   = f"{_last_date.year}年{_last_date.month}月{_last_date.day}日"
                st.write(f"🗓️ **最終補給日：** {_date_jp}（{_days_ago}日前）")
            except ValueError:
                st.write(f"🗓️ **最終補給日：** {_last_date_str}")
            st.write(f"🛢️ **最終補給量：** {_last_supply:.2f} L")
        else:
            st.caption("📭 補給記録なし")

        if history:
            with st.expander(f"📦 配送履歴（{len(history)}件）を見る"):
                # 月別集計
                monthly: dict = {}
                for h in history:
                    date_str = str(h.get("日付", ""))
                    try:
                        supply = float(h.get("補給量(L)", 0) or 0)
                    except (ValueError, TypeError):
                        supply = 0.0
                    if len(date_str) >= 7:
                        ym = date_str[:7]                 # "YYYY-MM"
                        y, m = ym[:4], ym[5:7].lstrip("0") or "0"
                        label = f"{y}年{m}月"
                    else:
                        label = "不明"
                    monthly[label] = monthly.get(label, 0.0) + supply

                # 月別補給量（新しい月順）
                st.markdown("**📅 月別補給量**")
                for label in sorted(monthly.keys(), reverse=True):
                    st.write(f"　・{label}：**{monthly[label]:.2f} L**")

                # 年間合計・月平均
                total_all = sum(monthly.values())
                active_months = len([v for v in monthly.values() if v > 0])
                monthly_avg = total_all / active_months if active_months else 0.0
                st.info(
                    f"📊 年間合計：**{total_all:.2f} L**　｜　"
                    f"月平均：**{monthly_avg:.2f} L**（{active_months}か月分）"
                )

                st.markdown("---")

                # 個別履歴一覧（新しい順）
                st.markdown("**📋 訪問履歴一覧**")
                for h in sorted(history, key=lambda x: str(x.get("日付", "")), reverse=True):
                    d = str(h.get("日付", "---"))
                    try:
                        sp = float(h.get("補給量(L)", 0) or 0)
                    except (ValueError, TypeError):
                        sp = 0.0
                    absent = "🚪不在　" if str(h.get("不在", "")).strip() == "✓" else ""
                    rental = "📄伝票あり" if str(h.get("レンタル伝票投函", "")).strip() == "✓" else ""
                    st.write(f"**{d}**　補給量: {sp:.2f} L　{absent}{rental}")
        else:
            st.caption("（配送記録なし）")

        st.markdown("---")

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
    today_str    = datetime.date.today().strftime("%Y-%m-%d")
    _t4_month    = datetime.date.today().strftime("%Y-%m")
    st.caption(f"📅 配送日：{today_str}")

    area = st.selectbox("エリアを選択してください", SHEET_NAMES, key="delivery_area")
    area_data = [r for r in data if r.get("エリア") == area]

    if not area_data:
        st.warning("このエリアに顧客データがありません")
    else:
        # ── 今月の訪問状況を配送記録から取得 ──────────────────
        _t4_all = load_delivery_records()

        # 今月1度でも訪問済み（✓）の顧客コードセット
        _t4_visited = {
            str(r.get("顧客コード", "")).strip()
            for r in _t4_all
            if str(r.get("日付", "")).strip().startswith(_t4_month)
            and str(r.get("訪問済み", "")).strip() == "✓"
        }

        # 各顧客の最終補給情報（code → (date_str, supply)）
        _t4_last: dict = {}
        for _r in _t4_all:
            _c = str(_r.get("顧客コード", "")).strip()
            try:
                _s = float(_r.get("補給量(L)", 0) or 0)
            except (ValueError, TypeError):
                _s = 0.0
            if _c and _s > 0:
                _d = str(_r.get("日付", ""))
                if _c not in _t4_last or _d > _t4_last[_c][0]:
                    _t4_last[_c] = (_d, _s)

        # ── エリアサマリー ─────────────────────────────────────
        _t4_v_cnt = sum(
            1 for r in area_data
            if str(r.get("顧客コード", "")).strip() in _t4_visited
        )
        _t4_u_cnt = len(area_data) - _t4_v_cnt

        _sc1, _sc2, _sc3 = st.columns(3)
        _sc1.info(f"📦 {len(area_data)} 件")
        _sc2.success(f"✅ 今月訪問済み：{_t4_v_cnt} 件")
        if _t4_u_cnt > 0:
            _sc3.warning(f"⚠️ 今月未訪問：{_t4_u_cnt} 件")
        else:
            _sc3.success(f"✅ 今月未訪問：0 件")

        # ── 入力フォーム ───────────────────────────────────────
        with st.form("delivery_form"):
            records = []
            for i, row in enumerate(area_data):
                code       = str(row.get("顧客コード", "")).strip()
                name       = row.get("名前", "---")
                is_visited = code in _t4_visited
                is_rental  = code.upper().endswith("R")

                # 訪問状態バッジ（HTML）
                if is_visited:
                    badge = (
                        "<span style='background:#d4edda;color:#155724;"
                        "border-radius:4px;padding:2px 8px;"
                        "font-size:0.82em;font-weight:bold;'>"
                        "✅ 今月訪問済み</span>"
                    )
                else:
                    badge = (
                        "<span style='background:#fff3cd;color:#856404;"
                        "border-radius:4px;padding:2px 8px;"
                        "font-size:0.82em;font-weight:bold;'>"
                        "⚠️ 今月未訪問</span>"
                    )

                rental_mark = "　🔴" if is_rental else ""
                st.markdown(
                    f"**{name}{rental_mark}**　`{code}`　{badge}",
                    unsafe_allow_html=True,
                )

                # 最終補給情報（小さく）
                if code in _t4_last:
                    _ld, _ls = _t4_last[code]
                    try:
                        _ldt  = datetime.date.fromisoformat(_ld)
                        _ldjp = f"{_ldt.year}年{_ldt.month}月{_ldt.day}日"
                        _days = (datetime.date.today() - _ldt).days
                        st.caption(f"🗓️ 最終補給：{_ldjp}（{_days}日前）　{_ls:.2f} L")
                    except ValueError:
                        st.caption(f"🗓️ 最終補給：{_ld}　{_ls:.2f} L")

                st.caption(f"📍 {row.get('住所', '---')}")

                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    visited = st.checkbox("✅ 訪問済み", key=f"visited_{i}")
                with col2:
                    supply = st.number_input(
                        "🛢 補給量 (L)",
                        min_value=0.0,
                        step=0.01,
                        format="%.2f",
                        key=f"supply_{i}"
                    )
                with col3:
                    absent = st.checkbox("🚪 不在", key=f"absent_{i}")
                with col4:
                    rental = st.checkbox("📄 レンタル伝票投函", key=f"rental_{i}")

                st.markdown("---")

                records.append({
                    "日付":            today_str,
                    "エリア":          area,
                    "顧客コード":      code,
                    "名前":            str(row.get("名前", "")),
                    "住所":            str(row.get("住所", "")),
                    "訪問済み":        visited,
                    "補給量(L)":       supply,
                    "不在":            absent,
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
                load_delivery_records.clear()   # キャッシュをリフレッシュ

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
                        min_value=0.0,
                        step=0.01,
                        format="%.2f",
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

    _today        = datetime.date.today()
    _month_prefix = _today.strftime("%Y-%m")          # 例: "2026-03"
    st.caption(f"📅 対象月：{_today.year}年{_today.month}月（今月固定）")

    if st.button("🔍 今月のアラートを確認する", use_container_width=True, key="check_alerts"):
        try:
            client      = connect_sheets()
            spreadsheet = client.open_by_url(SPREADSHEET_URL)

            try:
                record_sheet = spreadsheet.worksheet("配送記録")
                all_records  = record_sheet.get_all_records()
            except Exception:
                st.warning("⚠️ 配送記録シートが見つかりません。先に配送記録を保存してください。")
                st.stop()

            # 今月の記録だけ絞り込む
            month_records = [
                r for r in all_records
                if str(r.get("日付", "")).strip().startswith(_month_prefix)
            ]

            # 今月1度でも訪問済み（✓）になった顧客コードのセット
            visited_codes = {
                str(r.get("顧客コード", "")).strip()
                for r in month_records
                if str(r.get("訪問済み", "")).strip() == "✓"
            }

            # 今月1度でも伝票投函済み（✓）になったレンタル顧客コードのセット
            rental_done_codes = {
                str(r.get("顧客コード", "")).strip()
                for r in month_records
                if str(r.get("レンタル伝票投函", "")).strip() == "✓"
            }

            # エリア別に未訪問・伝票漏れを集計
            area_results = {}
            total_unvisited     = 0
            total_rental_missed = 0

            for area_name in SHEET_NAMES:
                area_customers = [r for r in data if r.get("エリア") == area_name]
                unvisited = [
                    c for c in area_customers
                    if str(c.get("顧客コード", "")).strip() not in visited_codes
                ]
                rental_missed = [
                    c for c in area_customers
                    if str(c.get("顧客コード", "")).strip().upper().endswith("R")
                    and str(c.get("顧客コード", "")).strip() not in rental_done_codes
                ]
                area_results[area_name] = {
                    "unvisited":     unvisited,
                    "rental_missed": rental_missed,
                }
                total_unvisited     += len(unvisited)
                total_rental_missed += len(rental_missed)

            # ══ ① 大きなサマリー表示 ═══════════════════════════════
            visited_count = len(data) - total_unvisited
            st.markdown(
                f"<div style='text-align:center;padding:16px 0 8px;'>"
                f"<span style='font-size:1.1rem;color:#555;'>今月の訪問状況</span><br>"
                f"<span style='font-size:2.4rem;font-weight:bold;color:#1f77b4;'>"
                f"全 {len(data)} 件中</span>"
                f"<span style='font-size:2.4rem;font-weight:bold;color:#d62728;'>"
                f"　残り {total_unvisited} 件</span>"
                f"<span style='font-size:1.6rem;color:#555;'> 未訪問</span><br>"
                f"<span style='font-size:1.1rem;color:#2ca02c;'>✅ 訪問済み {visited_count} 件</span>"
                f"</div>",
                unsafe_allow_html=True,
            )
            st.markdown("---")

            # ══ ② 未訪問顧客リスト（エリア別） ══════════════════════
            st.markdown("### 🚶 今月まだ訪問記録がない顧客")
            if total_unvisited == 0:
                st.success("✅ 全顧客に訪問済みです！")
            else:
                for area_name in SHEET_NAMES:
                    unvisited = area_results[area_name]["unvisited"]
                    if not unvisited:
                        continue
                    with st.expander(
                        f"📍 {area_name}　（{len(unvisited)} 件）",
                        expanded=True,
                    ):
                        for c in unvisited:
                            code = str(c.get("顧客コード", "")).strip()
                            name = c.get("名前", "---")
                            addr = c.get("住所", "---")
                            is_rental = code.upper().endswith("R")
                            line = (
                                f"**{name}**　｜　"
                                f"コード: `{code}`　｜　{addr}"
                            )
                            if is_rental:
                                st.error(f"🔴 {line}")   # レンタル顧客は赤
                            else:
                                st.warning(line)

            st.markdown("---")

            # ══ ③ 伝票投函漏れリスト（エリア別） ════════════════════
            st.markdown("### 📄 今月伝票投函記録がないレンタル顧客")
            if total_rental_missed == 0:
                st.success("✅ 伝票投函漏れはありません！")
            else:
                for area_name in SHEET_NAMES:
                    rental_missed = area_results[area_name]["rental_missed"]
                    if not rental_missed:
                        continue
                    with st.expander(
                        f"📍 {area_name}　（{len(rental_missed)} 件）",
                        expanded=True,
                    ):
                        for c in rental_missed:
                            code = str(c.get("顧客コード", "")).strip()
                            st.error(
                                f"**{c.get('名前', '---')}**　｜　"
                                f"コード: `{code}`　｜　{c.get('住所', '---')}"
                            )

            # ══ ④ 集計サマリー ════════════════════════════════════
            st.markdown("---")
            st.info(
                f"📊 {_today.year}年{_today.month}月の配送記録：{len(month_records)} 件　｜　"
                f"未訪問 **{total_unvisited}** 件　｜　"
                f"伝票投函漏れ **{total_rental_missed}** 件"
            )

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

# ── 日報 ──────────────────────────────────────────────────────
with tab8:
    st.subheader("📊 日報")

    # 印刷時に Streamlit の UI 要素を非表示にする CSS
    st.markdown("""
    <style>
    @media print {
        header, footer, section[data-testid="stSidebar"],
        div[data-testid="stToolbar"], div[data-testid="stDecoration"],
        .stButton, .stDateInput, .stSelectbox, .block-container > div:first-child {
            display: none !important;
        }
        #nippo-print-area { margin: 0; padding: 0; }
    }
    </style>
    """, unsafe_allow_html=True)

    nippo_date = st.date_input(
        "📅 日付を選択", value=datetime.date.today(), key="nippo_date"
    )
    nippo_date_str = nippo_date.strftime("%Y-%m-%d")
    nippo_date_jp  = f"{nippo_date.year}年{nippo_date.month}月{nippo_date.day}日"

    if st.button("📊 日報を表示", use_container_width=True, key="show_nippo"):
        try:
            # 訪問済み or 補給量>0 かつ 不在でない行だけ抽出
            day_records = []
            for _r in load_delivery_records():
                if str(_r.get("日付", "")).strip() != nippo_date_str:
                    continue
                if str(_r.get("不在", "")).strip() == "✓":
                    continue
                try:
                    _sp = float(_r.get("補給量(L)", 0) or 0)
                except (ValueError, TypeError):
                    _sp = 0.0
                if str(_r.get("訪問済み", "")).strip() == "✓" or _sp > 0:
                    day_records.append(_r)

            if not day_records:
                st.info(f"📭 {nippo_date_jp} の配送記録はありません")
            else:
                # 合計補給量
                total_supply = 0.0
                for r in day_records:
                    try:
                        total_supply += float(r.get("補給量(L)", 0) or 0)
                    except (ValueError, TypeError):
                        pass

                # テーブル行 HTML を組み立てる
                rows_html = ""
                for idx, r in enumerate(day_records, 1):
                    try:
                        sp = float(r.get("補給量(L)", 0) or 0)
                    except (ValueError, TypeError):
                        sp = 0.0
                    sp_str   = f"{sp:.2f}" if sp > 0 else "―"
                    rental   = "✓" if str(r.get("レンタル伝票投函", "")).strip() == "✓" else ""
                    time_str = str(r.get("時間", "")).strip()

                    rows_html += f"""
                    <tr style="background:white;">
                      <td style="text-align:center;padding:7px 10px;border:1px solid #ccc;">{idx}</td>
                      <td style="padding:7px 10px;border:1px solid #ccc;">{r.get('顧客コード','')}</td>
                      <td style="padding:7px 10px;border:1px solid #ccc;font-weight:bold;">{r.get('名前','')}</td>
                      <td style="text-align:center;padding:7px 10px;border:1px solid #ccc;">{time_str}</td>
                      <td style="text-align:right;padding:7px 10px;border:1px solid #ccc;">{sp_str}</td>
                      <td style="text-align:center;padding:7px 10px;border:1px solid #ccc;">{rental}</td>
                    </tr>"""

                nippo_html = f"""
                <div id="nippo-print-area" style="font-family:'Hiragino Sans','Meiryo',sans-serif;max-width:800px;margin:0 auto;padding:20px;">
                  <h2 style="text-align:center;border-bottom:3px solid #1f77b4;padding-bottom:10px;color:#1f77b4;">
                    🛢️ 灯油配送　日報
                  </h2>
                  <p style="text-align:right;font-size:16px;margin-bottom:16px;">
                    <strong>配送日：{nippo_date_jp}</strong>
                  </p>
                  <table style="width:100%;border-collapse:collapse;font-size:14px;">
                    <thead>
                      <tr style="background:#1f77b4;color:white;">
                        <th style="padding:8px 10px;border:1px solid #ccc;width:40px;">No.</th>
                        <th style="padding:8px 10px;border:1px solid #ccc;">顧客コード</th>
                        <th style="padding:8px 10px;border:1px solid #ccc;">名前</th>
                        <th style="padding:8px 10px;border:1px solid #ccc;width:80px;">時間</th>
                        <th style="padding:8px 10px;border:1px solid #ccc;width:110px;">補給量 (L)</th>
                        <th style="padding:8px 10px;border:1px solid #ccc;width:90px;">レンタル伝票</th>
                      </tr>
                    </thead>
                    <tbody>
                      {rows_html}
                      <tr style="background:#e8f0fe;font-weight:bold;">
                        <td colspan="4" style="text-align:right;padding:8px 10px;border:1px solid #ccc;">合計</td>
                        <td style="text-align:right;padding:8px 10px;border:1px solid #ccc;">{total_supply:.2f}</td>
                        <td style="border:1px solid #ccc;"></td>
                      </tr>
                    </tbody>
                  </table>
                  <p style="margin-top:16px;font-size:15px;">
                    件数：<strong>{len(day_records)} 件</strong>
                    合計補給量：<strong>{total_supply:.2f} L</strong>
                  </p>
                </div>"""

                # components.html() は Markdown を介さず直接 HTML レンダリングするため
                # インデントによるコードブロック誤解釈が発生しない
                import streamlit.components.v1 as _components

                _full_html = (
                    "<!DOCTYPE html><html lang='ja'><head><meta charset='UTF-8'>"
                    "<style>"
                    "body{font-family:'Hiragino Sans','Meiryo','Yu Gothic',sans-serif;margin:16px;}"
                    "@media print{.no-print{display:none!important;}}"
                    "</style></head><body>"
                    + nippo_html
                    + "<div class='no-print' style='margin-top:20px;'>"
                    "<button onclick='window.print()' style='"
                    "background-color:#1f77b4;color:white;padding:12px 36px;"
                    "border:none;border-radius:6px;cursor:pointer;"
                    "font-size:16px;font-weight:bold;'>🖨️ 印刷</button>"
                    "</div></body></html>"
                )

                _components.html(
                    _full_html,
                    height=max(520, len(day_records) * 46 + 320),
                    scrolling=True,
                )

        except Exception as e:
            st.error(f"❌ 読み込みエラー：{e}")
