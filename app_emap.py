"""
NLSC 臺灣通用電子地圖 異動情資蒐整 — Streamlit 互動介面
=====================================================
執行方式：
    streamlit run app_emap.py

功能：
  - 側欄設定搜尋地區、時間範圍、金額下限
  - 新聞爬搜（JSON + Google News RSS）
  - 工程會清冊分析
  - 互動式表格與圖表視覺化
"""

import sys
import importlib
import os
from pathlib import Path
from datetime import date, datetime, timedelta

import streamlit as st
import pandas as pd

# ── 頁面設定 ──────────────────────────────────────────────────
st.set_page_config(
    page_title="EMAP 異動情資蒐整",
    page_icon="🗺️",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).parent

# ─────────────────────────────────────────────────────────────
# 載入核心模組（動態覆寫全域參數）
# ─────────────────────────────────────────────────────────────
@st.cache_resource(show_spinner="載入核心模組...")
def _load_module():
    """只載入一次，避免重複 import"""
    sys.path.insert(0, str(BASE_DIR))
    import demo_emap_intelligence as m
    return m

mod = _load_module()


def _get_runtime_secret(name: str) -> str:
    """優先讀取 Streamlit Secrets（含巢狀/大小寫差異），其次讀取環境變數。"""
    target = name.strip().upper()
    try:
        direct = st.secrets.get(name)
        if direct is not None and str(direct).strip():
            return str(direct).strip()

        for k, v in st.secrets.items():
            if str(k).strip().upper() == target and str(v).strip():
                return str(v).strip()
            if hasattr(v, "items"):
                for nk, nv in v.items():
                    if str(nk).strip().upper() == target and str(nv).strip():
                        return str(nv).strip()
    except Exception:
        pass
    for env_key in (name, name.lower(), name.upper()):
        val = os.getenv(env_key, "").strip()
        if val:
            return val
    return ""


def _sync_llm_runtime_config():
    """每次 rerun 同步 LLM 設定，避免 import/caching 吃到舊值。"""
    api_key = _get_runtime_secret("EMAP_LLM_API_KEY")
    base_url = _get_runtime_secret("EMAP_LLM_BASE_URL")
    model = _get_runtime_secret("EMAP_LLM_MODEL")
    timeout_raw = _get_runtime_secret("EMAP_LLM_TIMEOUT")

    if api_key:
        mod.LLM_CFG["api_key"] = api_key
    if base_url:
        mod.LLM_CFG["base_url"] = base_url
    if model:
        mod.LLM_CFG["model"] = model
    if timeout_raw:
        try:
            mod.LLM_CFG["timeout"] = int(timeout_raw)
        except ValueError:
            pass


def _dedup_runtime_mode() -> str:
    has_api_key = bool(str(mod.LLM_CFG.get("api_key", "")).strip())
    return "LLM" if (mod.LLM_DEDUP and has_api_key) else "fallback文字"


_sync_llm_runtime_config()


def _llm_key_loaded() -> bool:
    return bool(str(mod.LLM_CFG.get("api_key", "")).strip())

# 所有支援的縣市清單
ALL_CITIES = [
    "臺北市", "新北市", "桃園市", "臺中市", "臺南市", "高雄市",
    "基隆市", "新竹市", "嘉義市", "新竹縣", "苗栗縣", "彰化縣",
    "南投縣", "雲林縣", "嘉義縣", "屏東縣", "宜蘭縣", "花蓮縣",
    "臺東縣", "澎湖縣", "金門縣", "連江縣",
]
CITY_SHORT_MAP = {
    "臺北市": ["臺北", "台北"],
    "新北市": ["新北"],
    "桃園市": ["桃園"],
    "臺中市": ["臺中", "台中"],
    "臺南市": ["臺南", "台南"],
    "高雄市": ["高雄"],
    "基隆市": ["基隆"],
    "新竹市": ["新竹"],
    "嘉義市": ["嘉義"],
    "新竹縣": ["新竹"],
    "苗栗縣": ["苗栗"],
    "彰化縣": ["彰化"],
    "南投縣": ["南投"],
    "雲林縣": ["雲林"],
    "嘉義縣": ["嘉義"],
    "屏東縣": ["屏東"],
    "宜蘭縣": ["宜蘭"],
    "花蓮縣": ["花蓮"],
    "臺東縣": ["臺東", "台東"],
    "澎湖縣": ["澎湖"],
    "金門縣": ["金門"],
    "連江縣": ["馬祖", "連江"],
}

# ─────────────────────────────────────────────────────────────
# 側欄：使用者輸入
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.image("https://www.nlsc.gov.tw/img/logo.png", width=120) if False else None
    st.title("EMAP 情資蒐整")
    st.markdown("---")

    st.subheader("搜尋地區")
    selected_cities = st.multiselect(
        "選擇縣市（可複選）",
        options=ALL_CITIES,
        default=["臺北市", "桃園市"],
        help="選擇要搜尋的縣市範圍，工程會清冊與新聞爬搜同步套用"
    )
    if not selected_cities:
        st.warning("請至少選擇一個縣市")
        st.stop()

    st.subheader("時間範圍")
    news_days = st.slider(
        "新聞近幾天",
        min_value=7,
        max_value=180,
        value=30,
        step=7,
        help="只納入距今幾天內的新聞（JSON 來源 A）"
    )
    date_from = date.today() - timedelta(days=news_days)
    st.caption(f"涵蓋：{date_from} ～ {date.today()}")

    st.subheader("金額下限")
    min_price_m = st.number_input(
        "工程金額下限（萬元）",
        min_value=0,
        max_value=100000,
        value=500,
        step=100,
        help="工程會清冊篩選用，新聞搜尋不受限"
    )
    min_price_k = min_price_m * 100  # 轉千元

    st.subheader("進階設定")
    web_qty = st.slider(
        "Web Search 每次查詢筆數",
        min_value=5,
        max_value=30,
        value=10,
        step=5,
    )
    show_dedup = st.checkbox("顯示去重追蹤（Debug）", value=False)
    auto_update = st.checkbox(
        "分析完成後自動更新案管檔",
        value=True,
        help="新聞與工程會結果都存在時，按下執行按鈕後自動重寫 115EMAP案管_更新版.xlsx。"
    )
    if _llm_key_loaded():
        st.success("LLM Key 狀態：已載入")
    else:
        st.error("LLM Key 狀態：未載入（目前會使用 fallback 去重）")

    st.markdown("---")
    run_news = st.button("執行新聞爬搜", width="stretch")
    run_pcc  = st.button("執行工程會分析", width="stretch")
    run_update = st.button("產出案管更新版", width="stretch")
    run_all  = st.button("全部執行", type="primary", width="stretch")

# ─────────────────────────────────────────────────────────────
# 套用使用者參數到模組全域變數
# ─────────────────────────────────────────────────────────────
def apply_params():
    """把 UI 設定寫回 demo_emap_intelligence 的全域變數"""
    _sync_llm_runtime_config()
    shorts = []
    for c in selected_cities:
        shorts.extend(CITY_SHORT_MAP.get(c, []))
    shorts = list(dict.fromkeys(shorts))  # 去重保序

    mod.NEWS_CITIES       = list(selected_cities)
    mod.NEWS_CITIES_SHORT = shorts
    mod.TARGET_CITY       = selected_cities[0]
    mod.NEWS_DAYS         = news_days
    mod.MIN_PRICE_K       = min_price_k
    mod.WEB_SEARCH_QTY    = web_qty
    mod.WEB_DEBUG         = False          # Streamlit 不需要 print 細節
    mod.LLM_DEDUP         = True           # 去重固定使用 LLM；失敗才由核心模組保守 fallback

# ─────────────────────────────────────────────────────────────
# 快取執行結果（依參數組合 cache）
# ─────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def run_news_search(cities_key, news_days_val, web_qty_val):
    apply_params()
    mod.LLM_DEDUP = True
    nlsc_df = pd.read_excel(mod.NLSC_FORM, sheet_name="工作表2")
    kw       = mod.load_keywords(mod.KEYWORD_FILE)
    build_kws = kw["include"]
    noise_kws = kw["exclude"]
    loc_kws   = mod.derive_news_loc_keywords(
        mod.NEWS_CITIES, mod.NEWS_CITIES_SHORT, nlsc_df)
    json_news, json_dedup = mod.parse_news_json(loc_kws, build_kws, noise_kws)
    web_news,  web_dedup  = mod.web_search_news(loc_kws, build_kws, noise_kws)
    all_news  = mod.merge_news(json_news, web_news)
    return all_news, json_news, web_news, json_dedup, web_dedup, loc_kws

@st.cache_data(show_spinner=False)
def run_pcc_analysis(cities_key, min_price_k_val):
    apply_params()
    nlsc_df = pd.read_excel(mod.NLSC_FORM, sheet_name="工作表2")
    return mod.process_pcc(nlsc_df)

def run_nlsc_update_workbook(pcc_result, news_result):
    """寫出 115EMAP案管_更新版.xlsx，並回傳更新日誌 list。"""
    all_news = news_result[0] if news_result else []
    return mod.write_nlsc_update(pcc_result or {}, all_news)

def load_update_log_df():
    out_path = BASE_DIR / "115EMAP案管_更新版.xlsx"
    if not out_path.exists():
        return pd.DataFrame()
    try:
        return pd.read_excel(out_path, sheet_name="更新日誌", header=1)
    except Exception:
        return pd.DataFrame()

# ─────────────────────────────────────────────────────────────
# 主頁面標題
# ─────────────────────────────────────────────────────────────
st.title("臺灣通用電子地圖 異動情資蒐整")
cols_h = st.columns([3,1,1,1])
with cols_h[0]:
    st.markdown(f"**搜尋範圍：** {'、'.join(selected_cities)}　｜　**時間：** 近 {news_days} 天　｜　**金額下限：** {min_price_m:,} 萬元")
with cols_h[3]:
    st.markdown(f"*{date.today()}*")

st.divider()

# ─────────────────────────────────────────────────────────────
# Session state 管理
# ─────────────────────────────────────────────────────────────
if "news_result"  not in st.session_state: st.session_state.news_result  = None
if "pcc_result"   not in st.session_state: st.session_state.pcc_result   = None
if "update_log"   not in st.session_state: st.session_state.update_log   = None
if "cities_key"   not in st.session_state: st.session_state.cities_key   = None
if "error_msg"    not in st.session_state: st.session_state.error_msg    = None
if "last_news_refresh" not in st.session_state: st.session_state.last_news_refresh = None
if "last_pcc_refresh" not in st.session_state: st.session_state.last_pcc_refresh = None
if "last_update_refresh" not in st.session_state: st.session_state.last_update_refresh = None

cities_key = ",".join(sorted(selected_cities))

# ─────────────────────────────────────────────────────────────
# 執行搜尋
# ─────────────────────────────────────────────────────────────
if run_news:
    run_news_search.clear()
    st.session_state.news_result = None
if run_pcc:
    run_pcc_analysis.clear()
    st.session_state.pcc_result = None
if run_all:
    run_news_search.clear()
    run_pcc_analysis.clear()
    st.session_state.news_result = None
    st.session_state.pcc_result = None
    st.session_state.update_log = None

if run_news or run_all:
    apply_params()
    with st.spinner("正在搜尋新聞（JSON + Google News RSS）..."):
        try:
            result = run_news_search(cities_key, news_days, web_qty)
            st.session_state.news_result = result
            st.session_state.cities_key  = cities_key
            st.session_state.last_news_refresh = datetime.now()
            st.session_state.error_msg   = None
        except Exception as e:
            st.session_state.error_msg = f"新聞搜尋失敗：{e}"

if run_pcc or run_all:
    apply_params()
    with st.spinner("正在分析工程會清冊..."):
        try:
            pcc = run_pcc_analysis(cities_key, min_price_k)
            st.session_state.pcc_result = pcc
            st.session_state.last_pcc_refresh = datetime.now()
            st.session_state.error_msg  = None
        except Exception as e:
            st.session_state.error_msg = f"工程會分析失敗：{e}"

should_auto_update = (
    auto_update
    and (run_news or run_pcc or run_all)
    and st.session_state.pcc_result is not None
    and st.session_state.news_result is not None
)

if run_update or run_all or should_auto_update:
    apply_params()
    if run_update:
        run_news_search.clear()
        run_pcc_analysis.clear()
        st.session_state.news_result = None
        st.session_state.pcc_result = None

    if st.session_state.news_result is None:
        with st.spinner("重新執行新聞爬搜，確保更新版使用最新 LLM 去重結果..."):
            try:
                st.session_state.news_result = run_news_search(cities_key, news_days, web_qty)
                st.session_state.last_news_refresh = datetime.now()
            except Exception as e:
                st.session_state.error_msg = f"新聞搜尋失敗，無法產出案管更新版：{e}"
    if st.session_state.pcc_result is None and st.session_state.error_msg is None:
        with st.spinner("重新執行工程會分析，確保更新版使用最新清冊結果..."):
            try:
                st.session_state.pcc_result = run_pcc_analysis(cities_key, min_price_k)
                st.session_state.last_pcc_refresh = datetime.now()
            except Exception as e:
                st.session_state.error_msg = f"工程會分析失敗，無法產出案管更新版：{e}"
    if st.session_state.pcc_result is None or st.session_state.news_result is None:
        if st.session_state.error_msg is None:
            st.session_state.error_msg = "產出更新版前，請先執行新聞爬搜與工程會分析。"
    elif st.session_state.error_msg is None:
        with st.spinner("正在產出 115EMAP 案管更新版..."):
            try:
                st.session_state.update_log = run_nlsc_update_workbook(
                    st.session_state.pcc_result,
                    st.session_state.news_result,
                )
                st.session_state.last_update_refresh = datetime.now()
                st.session_state.error_msg = None
            except Exception as e:
                st.session_state.error_msg = f"案管更新版產出失敗：{e}"

if st.session_state.error_msg:
    st.error(st.session_state.error_msg)

# ─────────────────────────────────────────────────────────────
# 頁籤
# ─────────────────────────────────────────────────────────────
tab_flow, tab_news, tab_pcc, tab_update_tab, tab_chart, tab_about = st.tabs(
    ["實作流程", "新聞爬搜", "工程會清冊", "案管更新", "圖表分析", "說明"])

# ══════════════════════════════════════════════════════════════
# TAB 0：實作流程
# ══════════════════════════════════════════════════════════════
with tab_flow:
    st.subheader("互動式實作流程總覽")
    flow_steps = pd.DataFrame([
        {"步驟": "1. 載入資料", "輸入": "NLSC案管表、工程會JSON、新聞JSON、關鍵字TXT", "處理": "統一讀檔並保留原始表單", "輸出": "可追蹤資料來源"},
        {"步驟": "2. 動態衍生關鍵字", "輸入": "縣市設定 + NLSC參考門牌", "處理": "產生縣市、縮寫、道路位置詞", "輸出": "新聞與工程會共用篩選條件"},
        {"步驟": "3. 工程會清冊", "輸入": "20260127 / 20260415 JSON", "處理": "金額、關鍵字、既有案、前次篩選交叉比對", "輸出": "新增建議、進度異動、摘要"},
        {"步驟": "4. 新聞爬搜", "輸入": "每日新聞JSON + Google News RSS", "處理": "位置/工程詞/鄰近性/LLM去重整併", "輸出": "案名、網址、備註、進度、完工日"},
        {"步驟": "5. 案管更新", "輸入": "工程會異動 + 新聞結果", "處理": "既有案更新；未命中新聞新增候選列", "輸出": "115EMAP案管_更新版.xlsx + 更新日誌"},
    ])
    st.dataframe(flow_steps, width="stretch", hide_index=True, height=220)

    st.markdown("#### 目前產出狀態")
    f1, f2, f3, f4 = st.columns(4)
    news_status = "已刷新" if st.session_state.news_result else "待執行"
    pcc_status = "已刷新" if st.session_state.pcc_result else "待執行"
    f1.metric("新聞結果", news_status)
    f2.metric("工程會結果", pcc_status)
    log_df_now = load_update_log_df()
    news_log_n = 0
    pcc_log_n = 0
    if not log_df_now.empty and "資料來源" in log_df_now.columns:
        news_log_n = log_df_now["資料來源"].astype(str).str.contains("新聞", na=False).sum()
        pcc_log_n = log_df_now["資料來源"].astype(str).str.contains("PCC|工程會", na=False).sum()
    f3.metric("PCC更新日誌", f"{pcc_log_n} 筆")
    f4.metric("新聞更新日誌", f"{news_log_n} 筆")

    def _fmt_ts(ts):
        return ts.strftime("%Y-%m-%d %H:%M:%S") if ts else "尚未於本頁刷新"

    st.caption(
        "刷新時間｜"
        f"新聞：{_fmt_ts(st.session_state.last_news_refresh)}　"
        f"工程會：{_fmt_ts(st.session_state.last_pcc_refresh)}　"
        f"案管更新版：{_fmt_ts(st.session_state.last_update_refresh)}"
    )

    st.markdown("#### 一鍵操作")
    st.write("左側的「全部執行」會清除本頁快取，依序重跑新聞、工程會分析，並產出案管更新版。若只想重新寫 Excel，可使用「產出案管更新版」，它也會先清除快取並重跑資料。")
    st.write("若「分析完成後自動更新案管檔」保持開啟，只要本次按鈕執行後新聞與工程會結果都存在，就會自動重寫更新版 Excel。")
    key_status = "已偵測到 LLM Key" if str(mod.LLM_CFG.get("api_key", "")).strip() else "未偵測到 LLM Key"
    st.caption(f"目前新聞去重模式：{_dedup_runtime_mode()}（{key_status}）。")

# ══════════════════════════════════════════════════════════════
# TAB 1：新聞爬搜
# ══════════════════════════════════════════════════════════════
with tab_news:
    if st.session_state.news_result is None:
        st.info("請在左側設定搜尋條件後，點選「執行新聞爬搜」或「全部執行」")
    else:
        all_news, json_news, web_news, json_dedup, web_dedup, loc_kws = st.session_state.news_result

        # ── 指標列 ──────────────────────────────────────────────
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("整合結果", f"{len(all_news)} 筆")
        m2.metric("JSON 來源", f"{len(json_news)} 筆")
        m3.metric("Web 來源",  f"{len(web_news)} 筆")
        m4.metric("搜尋縣市",  f"{len(selected_cities)} 縣市")

        st.divider()

        # ── 搜尋/過濾列 ─────────────────────────────────────────
        col_f1, col_f2, col_f3 = st.columns([2, 1, 1])
        with col_f1:
            kw_filter = st.text_input("關鍵字篩選（標題 / 案名）", placeholder="輸入關鍵字...")
        with col_f2:
            city_filter = st.multiselect("縣市篩選", options=selected_cities, default=[])
        with col_f3:
            src_filter = st.multiselect(
                "來源篩選",
                options=["JSON新聞資料", "Web Search", "JSON新聞資料+Web Search"],
                default=[]
            )

        # ── 套用過濾 ─────────────────────────────────────────────
        df_news = pd.DataFrame(all_news)
        if df_news.empty:
            st.warning("沒有找到符合條件的新聞。請嘗試調整搜尋條件（擴大地區或時間範圍）。")
        else:
            if kw_filter:
                mask = (df_news["新聞標題"].str.contains(kw_filter, na=False) |
                        df_news["案名"].str.contains(kw_filter, na=False))
                df_news = df_news[mask]
            if city_filter:
                df_news = df_news[df_news["縣市"].isin(city_filter)]
            if src_filter:
                df_news = df_news[df_news["來源分類"].isin(src_filter)]

            st.caption(f"顯示 {len(df_news)} 筆（共 {len(all_news)} 筆）")

            # ── 互動表格 ──────────────────────────────────────────
            display_cols = ["通報日期", "縣市", "新聞標題", "案名", "工程進度",
                            "(預計)完工日", "來源分類", "網址"]
            df_display = df_news[[c for c in display_cols if c in df_news.columns]].copy()
            df_display.index = range(1, len(df_display) + 1)

            st.dataframe(
                df_display,
                width="stretch",
                height=420,
                column_config={
                    "網址": st.column_config.LinkColumn("網址", display_text="🔗"),
                    "通報日期": st.column_config.TextColumn("通報日期", width="small"),
                    "縣市": st.column_config.TextColumn("縣市", width="small"),
                    "工程進度": st.column_config.TextColumn("工程進度", width="medium"),
                    "(預計)完工日": st.column_config.TextColumn("完工日", width="small"),
                    "來源分類": st.column_config.TextColumn("來源", width="medium"),
                    "新聞標題": st.column_config.TextColumn("標題", width="large"),
                    "案名": st.column_config.TextColumn("案名", width="large"),
                },
                hide_index=False,
            )

            # ── 展開詳細備註 ──────────────────────────────────────
            with st.expander("顯示內文摘要（備註欄）"):
                for _, row in df_news.reset_index(drop=True).iterrows():
                    title = row.get("新聞標題", "")
                    note  = row.get("備註", "")
                    loc   = row.get("參考位置", "")
                    st.markdown(f"**{title}**")
                    st.caption(f"{loc}")
                    st.write(note[:300] if note else "（無摘要）")
                    st.divider()

            # ── 下載 ──────────────────────────────────────────────
            csv = df_news.to_csv(index=False, encoding="utf-8-sig")
            st.download_button(
                "下載 CSV",
                data=csv,
                file_name=f"emap_news_{date.today()}.csv",
                mime="text/csv",
            )

        # ── 去重追蹤（Debug）─────────────────────────────────────
        if show_dedup and (json_dedup or web_dedup):
            with st.expander("🔍 去重追蹤記錄（Debug）"):
                all_dedup = (json_dedup or []) + (web_dedup or [])
                df_dedup = pd.DataFrame(all_dedup)
                st.dataframe(df_dedup, width="stretch", height=300)

# ══════════════════════════════════════════════════════════════
# TAB 2：工程會清冊
# ══════════════════════════════════════════════════════════════
with tab_pcc:
    if st.session_state.pcc_result is None:
        st.info("請在左側點選「執行工程會分析」或「全部執行」")
    else:
        pcc = st.session_state.pcc_result

        # pcc 是 dict，包含 t1_new, t1_exist, t2, t3, t4 等
        t1_new  = pcc.get("t1_new",  [])
        t1_exist= pcc.get("t1_exist",[])
        t2      = pcc.get("t2",      [])
        t3      = pcc.get("t3",      [])
        t4      = pcc.get("t4",      [])

        # ── 指標列 ──────────────────────────────────────────────
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("建議新增納管", f"{len(t1_new)} 件",  delta="★ 新增")
        m2.metric("已在管控/前次篩選", f"{len(t1_exist)} 件")
        m3.metric("測試2 異動追蹤", f"{len(t2)} 件")
        t2_changed = sum(1 for r in t2 if "有異動" in r.get("有無異動",""))
        m4.metric("其中有異動", f"{t2_changed} 件",
                  delta=f"+{t2_changed}" if t2_changed else None,
                  delta_color="normal")

        st.divider()

        pcc_tab1, pcc_tab2, pcc_tab3 = st.tabs(
            ["測試1：新增建議", "測試2：異動追蹤", "測試3/4：Web補充"])

        with pcc_tab1:
            st.subheader("★ 建議新增納入管控")
            if t1_new:
                df_new = pd.DataFrame(t1_new)
                keep_cols = [c for c in ["基本資料_案名","坐落_縣市","坐落_區",
                             "核對追蹤_執行單位","核對追蹤_預算金額【千元】",
                             "核對追蹤_(預計)完工日","篩選說明"] if c in df_new.columns]
                df_new_disp = df_new[keep_cols].copy()
                df_new_disp.index = range(1, len(df_new_disp)+1)
                if "核對追蹤_預算金額【千元】" in df_new_disp.columns:
                    df_new_disp["預算（萬元）"] = (
                        pd.to_numeric(df_new_disp["核對追蹤_預算金額【千元】"], errors="coerce")
                        .fillna(0).astype(int) // 10
                    )
                st.dataframe(df_new_disp, width="stretch", height=360)
                st.download_button(
                    "下載建議新增清單",
                    data=df_new.to_csv(index=False, encoding="utf-8-sig"),
                    file_name=f"emap_pcc_new_{date.today()}.csv",
                    mime="text/csv",
                )
            else:
                st.info("無建議新增案件（所有案件均已在管控或前次已篩選）")

            st.subheader("已在管控 / 前次篩選 Y")
            if t1_exist:
                df_exist = pd.DataFrame(t1_exist)
                keep_e = [c for c in ["基本資料_案名","篩選結果","篩選說明",
                          "坐落_縣市","坐落_區"] if c in df_exist.columns]
                st.dataframe(df_exist[keep_e], width="stretch", height=260)
            else:
                st.info("無相關記錄")

        with pcc_tab2:
            st.subheader("測試2：0127 vs 0415 異動追蹤")
            if t2:
                df_t2 = pd.DataFrame(t2)
                # 高亮異動列
                def _color_changed(val):
                    return "background-color: #fff3cd" if "有異動" in str(val) else ""

                st.dataframe(
                    df_t2.style.applymap(_color_changed, subset=["有無異動"])
                    if "有無異動" in df_t2.columns else df_t2,
                    width="stretch", height=400
                )
                st.download_button(
                    "下載異動追蹤",
                    data=df_t2.to_csv(index=False, encoding="utf-8-sig"),
                    file_name=f"emap_pcc_t2_{date.today()}.csv",
                    mime="text/csv",
                )
            else:
                st.info("無可比對的異動資料（需要兩份不同日期的工程會 JSON）")

        with pcc_tab3:
            st.subheader("測試3/4：Web Search 補充")
            for label, data in [("測試3", t3), ("測試4", t4)]:
                st.markdown(f"**{label}**")
                if data:
                    st.dataframe(pd.DataFrame(data), width="stretch", height=220)
                else:
                    st.info(f"{label}：無資料")

# ══════════════════════════════════════════════════════════════
# TAB 3：案管更新
# ══════════════════════════════════════════════════════════════
with tab_update_tab:
    st.subheader("115EMAP 案管更新版")
    out_path = BASE_DIR / "115EMAP案管_更新版.xlsx"
    log_df = load_update_log_df()

    if out_path.exists():
        c1, c2, c3 = st.columns(3)
        c1.metric("更新版檔案", out_path.name)
        c2.metric("更新日誌", f"{len(log_df)} 筆" if not log_df.empty else "0 筆")
        if not log_df.empty and "資料來源" in log_df.columns:
            c3.metric("新聞來源", f"{log_df['資料來源'].astype(str).str.contains('新聞', na=False).sum()} 筆")
        else:
            c3.metric("新聞來源", "0 筆")

        if not log_df.empty:
            src_options = sorted(log_df["資料來源"].dropna().astype(str).unique()) if "資料來源" in log_df.columns else []
            src_pick = st.multiselect("資料來源篩選", options=src_options, default=[])
            show_df = log_df.copy()
            if src_pick and "資料來源" in show_df.columns:
                show_df = show_df[show_df["資料來源"].isin(src_pick)]
            st.dataframe(show_df, width="stretch", height=420, hide_index=True)

            with open(out_path, "rb") as f:
                st.download_button(
                    "下載 115EMAP案管_更新版.xlsx",
                    data=f.read(),
                    file_name=out_path.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        else:
            st.info("目前更新版存在，但沒有讀到更新日誌。請重新執行「產出案管更新版」。")
    else:
        st.info("尚未產出更新版。請先執行新聞爬搜與工程會分析，再按左側「產出案管更新版」。")

# ══════════════════════════════════════════════════════════════
# TAB 3：圖表分析
# ══════════════════════════════════════════════════════════════
with tab_chart:
    # 需要新聞資料才能畫圖
    has_news = (st.session_state.news_result is not None and
                len(st.session_state.news_result[0]) > 0)
    has_pcc  = (st.session_state.pcc_result  is not None)

    if not has_news and not has_pcc:
        st.info("請先執行搜尋以產生圖表")
    else:
        chart_col1, chart_col2 = st.columns(2)

        # ── 新聞時間趨勢 ──────────────────────────────────────────
        if has_news:
            all_news = st.session_state.news_result[0]
            df_chart = pd.DataFrame(all_news)

            with chart_col1:
                st.subheader("新聞時間分佈")
                if "通報日期" in df_chart.columns and not df_chart.empty:
                    df_chart["通報日期_dt"] = pd.to_datetime(df_chart["通報日期"], errors="coerce")
                    daily = (df_chart.dropna(subset=["通報日期_dt"])
                             .groupby(df_chart["通報日期_dt"].dt.strftime("%Y-%m-%d"))
                             .size()
                             .reset_index(name="篇數")
                             .rename(columns={"通報日期_dt": "日期"})
                             .sort_values("日期"))
                    if not daily.empty:
                        st.bar_chart(daily.set_index("日期")["篇數"], height=250)
                    else:
                        st.info("無日期資料")

            with chart_col2:
                st.subheader("新聞縣市分佈")
                if "縣市" in df_chart.columns and not df_chart.empty:
                    city_cnt = df_chart["縣市"].value_counts().reset_index()
                    city_cnt.columns = ["縣市", "篇數"]
                    st.bar_chart(city_cnt.set_index("縣市")["篇數"], height=250)

            col_c3, col_c4 = st.columns(2)

            with col_c3:
                st.subheader("工程進度分佈")
                if "工程進度" in df_chart.columns and not df_chart.empty:
                    prog_cnt = df_chart["工程進度"].value_counts().reset_index()
                    prog_cnt.columns = ["工程進度", "篇數"]
                    st.bar_chart(prog_cnt.set_index("工程進度")["篇數"], height=250)

            with col_c4:
                st.subheader("新聞來源分佈")
                if "來源分類" in df_chart.columns and not df_chart.empty:
                    src_cnt = df_chart["來源分類"].value_counts().reset_index()
                    src_cnt.columns = ["來源分類", "篇數"]
                    st.bar_chart(src_cnt.set_index("來源分類")["篇數"], height=250)

        # ── 工程會圖表 ──────────────────────────────────────────
        if has_pcc:
            pcc = st.session_state.pcc_result
            t1_new   = pcc.get("t1_new",  [])
            t1_exist = pcc.get("t1_exist",[])
            t2       = pcc.get("t2",      [])

            st.divider()
            st.subheader("工程會清冊統計")

            pcc_col1, pcc_col2 = st.columns(2)

            with pcc_col1:
                st.markdown("測試1 篩選結果占比")
                t1_labels = (
                    ["★ 建議新增"] * len(t1_new) +
                    [r.get("篩選結果", "其他") for r in t1_exist]
                )
                if t1_labels:
                    cnt_df = pd.Series(t1_labels).value_counts().reset_index()
                    cnt_df.columns = ["類別", "件數"]
                    st.bar_chart(cnt_df.set_index("類別")["件數"], height=220)

            with pcc_col2:
                st.markdown("測試2 異動 vs 無異動")
                if t2:
                    changed_cnt    = sum(1 for r in t2 if "有異動" in r.get("有無異動",""))
                    unchanged_cnt  = len(t2) - changed_cnt
                    chg_df = pd.DataFrame({
                        "狀態": ["有異動", "無異動"],
                        "件數": [changed_cnt, unchanged_cnt]
                    })
                    st.bar_chart(chg_df.set_index("狀態")["件數"], height=220)
                else:
                    st.info("無異動資料")

# ══════════════════════════════════════════════════════════════
# TAB 4：說明
# ══════════════════════════════════════════════════════════════
with tab_about:
    st.subheader("系統說明")
    st.markdown("""
    本系統為 **臺灣通用電子地圖（EMAP）異動情資蒐整** 的互動展示介面，
    整合工程會清冊分析與多來源新聞爬搜功能。

    #### 資料來源
    - **來源 A（JSON）**：`20260503_vw_8DNewsAI_noTag新聞資料.json`
      每日新聞資料，依時間窗口與關鍵字過濾
    - **來源 B（Web Search）**：Google News RSS
      即時爬搜，不需要 API key
    - **工程會清冊**：`output_file_all_case_format_v2_*.json`
      兩份不同日期的工程會資料，用於比對異動

    #### 搜尋邏輯
    1. **地區過濾**：縣市名稱與縮寫需出現在標題或內文
    2. **工程詞比對**：關鍵字 TXT 定義的應納入 / 非納入詞彙
    3. **鄰近性過濾**：位置詞與工程詞需在 300 字元內共現
    4. **去重邏輯**：使用LLM判斷是否為同一事件

    #### 如何使用
    1. 在左側設定地區、時間、金額條件
    2. 點選「執行新聞爬搜」取得新聞資料
    3. 點選「執行工程會分析」取得清冊分析
    4. 點選「全部執行」同時執行兩者
    5. 在「圖表分析」頁籤查看視覺化結果

    #### 注意事項
    - Web Search 需要網路連線（Google News RSS）
    - LLM 後分類功能需設定 `.emap_llm_config.json` 或環境變數
    - 結果可下載為 CSV 格式
    """)

    st.subheader("資料檔案狀態")
    files = {
        "關鍵字檔": mod.KEYWORD_FILE,
        "NLSC 管控表單": mod.NLSC_FORM,
        "案件管理 Excel": mod.PCC_CASE_MGR,
        "工程會 JSON (0127)": mod.JSON_127,
        "工程會 JSON (0415)": mod.JSON_415,
    }
    for name, path in files.items():
        exists = Path(path).exists()
        icon   = "✅" if exists else "❌"
        size   = f"（{Path(path).stat().st_size // 1024:,} KB）" if exists else ""
        st.write(f"{icon} **{name}**：`{Path(path).name}` {size}")
