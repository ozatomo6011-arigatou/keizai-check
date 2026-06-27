import streamlit as st
import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta, timezone
import anthropic
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from pathlib import Path
import os
import re
import requests
import io
from dotenv import load_dotenv

load_dotenv()  # .envファイルから自動でAPIキーを読み込む

st.set_page_config(
    page_title="毎日の経済チェック",
    page_icon="📊",
    layout="wide",
)

EXCEL_FILE = Path("経済日誌.xlsx")

TICKERS = {
    "日経平均":       "^N225",
    "ダウ平均":       "^DJI",
    "S&P500":        "^GSPC",
    "NASDAQ":        "^IXIC",
    "SOX（半導体）":  "^SOX",
    "ドル円":         "USDJPY=X",
    "米国10年債(%)":  "^TNX",
    "日本10年債(%)":  None,  # 財務省から別途取得
    "原油(WTI)":     "CL=F",
    "金":            "GC=F",
    "ビットコイン":   "BTC-USD",
}

# ──────────────────────────────
# データ取得
# ──────────────────────────────
def reiwa_to_seireki(date_str: str) -> str:
    """財務省CSVの和暦表記(例: R8.6.25)を西暦(例: 2026-06-25)に変換する"""
    m = re.match(r"R(\d+)\.(\d+)\.(\d+)", date_str.strip())
    if not m:
        return date_str
    year = 2018 + int(m.group(1))
    return f"{year}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"


def fetch_jp10y() -> dict | None:
    """財務省から日本10年債利回りを取得"""
    try:
        url = "https://www.mof.go.jp/jgbs/reference/interest_rate/jgbcm.csv"
        r = requests.get(url, timeout=10, headers={"User-Agent": "Mozilla/5.0"})
        content = r.content.decode("shift_jis", errors="replace")
        df = pd.read_csv(io.StringIO(content), skiprows=1, header=0)
        df = df.dropna(subset=[df.columns[0]])
        df = df[pd.to_numeric(df.iloc[:, 1], errors="coerce").notna()]
        if len(df) >= 2:
            last = float(df.iloc[-1].iloc[10])
            prev = float(df.iloc[-2].iloc[10])
            date_str = reiwa_to_seireki(str(df.iloc[-1].iloc[0]))
            return {"value": last, "change": last - prev, "pct": (last - prev) / prev * 100, "date": date_str}
        elif len(df) == 1:
            date_str = reiwa_to_seireki(str(df.iloc[-1].iloc[0]))
            return {"value": float(df.iloc[-1].iloc[10]), "change": None, "pct": None, "date": date_str}
    except Exception:
        pass
    return None


def _fetch_one_ticker(ticker: str) -> dict | None:
    """個別銘柄を取得し直す(株価系の日付ズレ対策の再取得用)"""
    try:
        prices = yf.download(ticker, period="2d", interval="1d", progress=False, auto_adjust=True)["Close"]
        prices = prices.iloc[:, 0].dropna() if hasattr(prices, "columns") else prices.dropna()
        if len(prices) == 0:
            return None
        ticker_date = str(prices.index[-1].date()) if hasattr(prices.index, "date") else None
        if len(prices) >= 2:
            prev, last = float(prices.iloc[-2]), float(prices.iloc[-1])
            return {"value": last, "change": last - prev, "pct": (last - prev) / prev * 100, "date": ticker_date}
        return {"value": float(prices.iloc[-1]), "change": None, "pct": None, "date": ticker_date}
    except Exception:
        return None


@st.cache_data(ttl=300)
def fetch_market_data():
    results = {}
    yf_tickers = {name: ticker for name, ticker in TICKERS.items() if ticker is not None}
    try:
        raw = yf.download(list(yf_tickers.values()), period="2d", interval="1d", progress=False, auto_adjust=True)
        close = raw["Close"] if "Close" in raw.columns else raw
        for name, ticker in yf_tickers.items():
            try:
                prices = close[ticker].dropna()
                # 銘柄ごとに実際にデータがある最後の日付を使う
                # (ビットコインは土日も動くため、共通の日付インデックスを使うと
                # 株式・為替など平日のみ動く指標にも誤って今日の日付がついてしまう)
                ticker_date = str(prices.index[-1].date()) if len(prices) and hasattr(prices.index, "date") else None
                if len(prices) >= 2:
                    prev, last = float(prices.iloc[-2]), float(prices.iloc[-1])
                    change = last - prev
                    pct = change / prev * 100
                    results[name] = {"value": last, "change": change, "pct": pct, "date": ticker_date}
                elif len(prices) == 1:
                    results[name] = {"value": float(prices.iloc[-1]), "change": None, "pct": None, "date": ticker_date}
                else:
                    results[name] = None
            except Exception:
                results[name] = None
    except Exception as e:
        st.error(f"データ取得エラー: {e}")

    # 銘柄ごとにYahoo側のデータ配信が遅れて日付がズレることがあるため、
    # 他の銘柄より明らかに古い日付のものだけ個別に再取得する
    valid_dates = [r["date"] for r in results.values() if r and r.get("date")]
    if valid_dates:
        newest_date = max(valid_dates)
        for name, ticker in yf_tickers.items():
            r = results.get(name)
            if r and r.get("date") and r["date"] < newest_date:
                retry = _fetch_one_ticker(ticker)
                if retry and retry.get("date") and retry["date"] >= newest_date:
                    results[name] = retry

    # 日本10年債は財務省から取得
    results["日本10年債(%)"] = fetch_jp10y()
    return results


def fmt_value(name, val):
    if "債" in name or "%" in name:
        return f"{val:.2f}%"
    if "ドル円" in name:
        return f"{val:.2f}円"
    if "ビットコイン" in name:
        return f"${val:,.0f}"
    if "原油" in name or "金" in name:
        return f"${val:,.2f}"
    return f"{val:,.2f}"


def metric_card(name, info):
    if info is None:
        st.metric(name, "取得失敗")
        return
    val_str = fmt_value(name, info["value"])
    if info["pct"] is not None:
        delta_str = f"{info['pct']:+.2f}%"
        st.metric(name, val_str, delta=delta_str)
    else:
        st.metric(name, val_str)
    if info.get("date"):
        st.caption(f"📅 {info['date']}")


# ──────────────────────────────
# Claude AIコメント生成
# ──────────────────────────────
def generate_comment(data: dict, api_key: str) -> str:
    lines = []
    for name, info in data.items():
        if info:
            v = fmt_value(name, info["value"])
            if info["pct"] is not None:
                lines.append(f"{name}: {v}（{info['pct']:+.2f}%）")
            else:
                lines.append(f"{name}: {v}")

    market_text = "\n".join(lines)
    prompt = f"""以下は本日の主要市場データです。

{market_text}

これを見て、初心者の経済学習者向けに、必ず次の見出しをそのまま使った形式で日本語で出力してください。

【まとめ】
（市場全体のムード：リスクオン/リスクオフを100字程度で）

【注目ポイント】
（注目すべき動きとその理由を100字程度で、推測でOK）

【今日の市場まとめを読んで疑問に思ったこと】
（メンター面談で聞くと良さそうなポイントを1つ、1文で）

【質問への回答】
（上記の質問に対する、あなた自身の回答・考え方を150字程度で。ここだけは小学生にもわかるレベルまでやさしい言葉で、難しい専門用語を避けて説明してください）"""

    client = anthropic.Anthropic(api_key=api_key)
    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=600,
        messages=[{"role": "user", "content": prompt}],
    )
    return message.content[0].text


def parse_ai_comment(text: str) -> dict:
    """AIコメントを見出しごとに分割する。形式に従っていない場合はsummaryに全文を入れる。"""
    sections = {"summary": "", "question": "", "answer": ""}
    heading_map = {
        "まとめ": "summary",
        "注目ポイント": "summary",
        "今日の市場まとめを読んで疑問に思ったこと": "question",
        "質問への回答": "answer",
    }
    parts = re.split(r"【(.+?)】", text)
    if len(parts) <= 1:
        sections["summary"] = text.strip()
        return sections
    for i in range(1, len(parts), 2):
        heading = parts[i].strip()
        content = parts[i + 1].strip() if i + 1 < len(parts) else ""
        key = heading_map.get(heading)
        if key == "summary" and sections["summary"]:
            sections["summary"] += "\n" + content
        elif key:
            sections[key] = content
    return sections


# ──────────────────────────────
# Google Sheets
# ──────────────────────────────
def get_gsheet():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    import google.oauth2.service_account as sa
    creds = sa.Credentials.from_service_account_info(
        creds_dict,
        scopes=["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"],
    )
    import gspread
    gc = gspread.authorize(creds)
    spreadsheet_id = st.secrets.get("SPREADSHEET_ID", "") or os.environ.get("SPREADSHEET_ID", "")
    return gc.open_by_key(spreadsheet_id).sheet1


def save_to_gsheet(data: dict, comment: str, picks: str, today: str):
    ws = get_gsheet()
    desired_headers = ["日付"] + list(TICKERS.keys()) + ["AIコメント", "気になる銘柄"]
    all_values = ws.get_all_values()
    if not all_values or all_values[0][0] != "日付":
        ws.insert_row(desired_headers, 1)
        all_values = ws.get_all_values()

    headers = all_values[0]
    if headers != desired_headers:
        # 列の並び順が変わった場合、既存データも新しい並び順に揃え直す
        reordered_rows = []
        for r in all_values[1:]:
            row_dict = dict(zip(headers, r))
            reordered_rows.append([row_dict.get(h, "") for h in desired_headers])
        ws.clear()
        ws.update("A1", [desired_headers] + reordered_rows)
        headers = desired_headers
        all_values = [desired_headers] + reordered_rows

    # 今日の日付がすでにあれば、AIコメントが未記入の場合のみ追記する
    existing_dates = [r[0] for r in all_values[1:]]
    if today in existing_dates:
        row_index = existing_dates.index(today) + 2  # 1行目はヘッダーなので+2
        comment_col = desired_headers.index("AIコメント") + 1
        existing_comment = all_values[existing_dates.index(today) + 1][comment_col - 1]
        if not existing_comment and comment:
            ws.update_cell(row_index, comment_col, comment)
            return True
        return False
    row = [today]
    for name in TICKERS:
        info = data.get(name)
        row.append(round(info["value"], 4) if info else "N/A")
    row.append(comment)
    row.append(picks)
    ws.append_row(row)
    return True


def load_from_gsheet() -> pd.DataFrame | None:
    try:
        ws = get_gsheet()
        records = ws.get_all_values()
        if len(records) < 2:
            return None
        df = pd.DataFrame(records[1:], columns=records[0])
        df = df.set_index("日付")
        for col in df.columns:
            if col not in ["AIコメント", "気になる銘柄"]:
                df[col] = pd.to_numeric(df[col], errors="coerce")
        return df
    except Exception:
        return None


# ──────────────────────────────
# UI
# ──────────────────────────────
st.markdown("##### 📊 毎日の経済チェック")
JST = timezone(timedelta(hours=9))
st.caption(f"最終更新: {datetime.now(JST).strftime('%Y年%m月%d日 %H:%M')}")

# APIキー取得（Streamlit Secrets → 環境変数の順で読み込む）
try:
    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
except Exception:
    api_key = ""
api_key = api_key or os.environ.get("ANTHROPIC_API_KEY", "")

# サイドバー
with st.sidebar:
    st.header("⚙️ 設定")
    if api_key:
        st.success("AIコメント：利用可能")
    else:
        st.warning("APIキー未設定")
    st.divider()
    if st.button("🔄 データを再取得"):
        st.cache_data.clear()
        st.rerun()

# データ取得
with st.spinner("市場データを取得中..."):
    data = fetch_market_data()

# ──── 指標カードのフォントを縮小
st.markdown("""
<style>
div[data-testid="stMetricValue"] { font-size: 1.1rem; }
div[data-testid="stMetricLabel"] { font-size: 0.75rem; }
div[data-testid="stMetricDelta"] { font-size: 0.75rem; }
div[data-testid="stAlertContentInfo"] p,
div[data-testid="stExpander"] p { font-size: 0.85rem; }
</style>
""", unsafe_allow_html=True)

# ──── セクション1: 日米主要指数
st.markdown("###### 📈 日米主要指数")
cols = st.columns(5)
for i, name in enumerate(["日経平均", "ダウ平均", "S&P500", "NASDAQ", "SOX（半導体）"]):
    with cols[i]:
        metric_card(name, data.get(name))

# ──── セクション2: 為替・債券
st.markdown("###### 💱 為替・債券利回り")
cols = st.columns(3)
for i, name in enumerate(["ドル円", "米国10年債(%)", "日本10年債(%)"]):
    with cols[i]:
        metric_card(name, data.get(name))

# ──── セクション3: 資金逃避先
st.markdown("###### 🛡️ 資金逃避先（有事の動き）")
cols = st.columns(3)
for i, name in enumerate(["原油(WTI)", "金", "ビットコイン"]):
    with cols[i]:
        metric_card(name, data.get(name))

st.divider()

# ──── セクション4: AIコメント
st.markdown("###### 🤖 今日の市場まとめ")

from datetime import timezone, timedelta
JST = timezone(timedelta(hours=9))
now_jst = datetime.now(JST)
today_str = (now_jst if now_jst.hour >= 9 else now_jst - timedelta(days=1)).strftime("%Y-%m-%d")

# 今日のコメントをスプレッドシートから読み込む
if "ai_comment" not in st.session_state:
    df_existing = load_from_gsheet()
    if df_existing is not None and today_str in df_existing.index:
        saved = df_existing.loc[today_str, "AIコメント"]
        st.session_state.ai_comment = saved if isinstance(saved, str) else ""
    else:
        st.session_state.ai_comment = ""

if st.session_state.ai_comment:
    parsed = parse_ai_comment(st.session_state.ai_comment)
    if parsed["summary"]:
        st.info(parsed["summary"])
    if parsed["question"]:
        st.markdown(f"**🙋 今日の市場まとめを読んで疑問に思ったこと：** {parsed['question']}")
        if parsed["answer"]:
            with st.expander("💡 回答を見る"):
                st.write(parsed["answer"])
    else:
        st.info(st.session_state.ai_comment)
    st.caption("※ 今日はすでに生成済みです")

    note_text = "🤖 今日の市場まとめ\n\n" + parsed["summary"]
    if parsed["question"]:
        note_text += f"\n\n🙋 今日の市場まとめを読んで疑問に思ったこと\n{parsed['question']}"
        if parsed["answer"]:
            note_text += f"\n\n💡 回答\n{parsed['answer']}"
    with st.expander("📝 note投稿用テキスト（コピーして使う）"):
        st.caption("上の画面全体をスクショして画像に、このテキストを本文に貼り付けてください")
        st.code(note_text, language=None)
else:
    if st.button("💬 AIコメントを生成", type="primary", disabled=not api_key):
        with st.spinner("Claudeが市場を分析中..."):
            try:
                st.session_state.ai_comment = generate_comment(data, api_key)
                st.rerun()
            except Exception as e:
                st.error(f"コメント生成エラー: {e}")

st.divider()

# ──── セクション5: 気になる銘柄
st.markdown("###### 📌 気になる銘柄（メンター面談用）")
st.caption("メンターとの面談で話したい銘柄やテーマを入力してください")

picks_input = st.text_area(
    "銘柄・テーマメモ",
    height=100,
    placeholder="例: NVIDIA（AI需要で急騰）、TSMC（SOX連動）、ドル円の動きが気になる",
    help="ここだけは手入力です。市場データを見て気になったことを書き留めましょう",
)

st.divider()

# ──── セクション6: 記録ボタン
st.markdown("###### 📁 今日のデータを記録する")
if st.button("💾 Googleスプレッドシートに保存", type="primary"):
    try:
        saved = save_to_gsheet(data, st.session_state.ai_comment, picks_input, today_str)
        if saved:
            st.success("✅ 保存しました！")
        else:
            st.warning("今日はすでに保存済みです")
    except Exception as e:
        st.error(f"保存エラー: {e}")

st.divider()

# ──── セクション7: 過去グラフ
st.markdown("###### 📈 過去の推移グラフ")
df = load_from_gsheet()
if df is not None and len(df) >= 2:
    chart_options = [c for c in df.columns if c not in ["AIコメント", "気になる銘柄"]]
    selected = st.multiselect("表示する項目を選択", chart_options, default=chart_options[:3])
    if selected:
        import plotly.graph_objects as go
        fig = go.Figure()
        for col in selected:
            fig.add_trace(go.Scatter(x=df.index, y=df[col], name=col, mode="lines+markers"))
        fig.update_layout(height=400, margin=dict(l=0, r=0, t=30, b=0))
        st.plotly_chart(fig, use_container_width=True)
else:
    st.info("まだデータがありません。「Googleスプレッドシートに保存」ボタンで記録を始めましょう！")
