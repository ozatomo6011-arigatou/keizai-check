import streamlit as st
import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta
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
            date_str = str(df.iloc[-1].iloc[0])
            return {"value": last, "change": last - prev, "pct": (last - prev) / prev * 100, "date": date_str}
        elif len(df) == 1:
            date_str = str(df.iloc[-1].iloc[0])
            return {"value": float(df.iloc[-1].iloc[10]), "change": None, "pct": None, "date": date_str}
    except Exception:
        pass
    return None


@st.cache_data(ttl=300)
def fetch_market_data():
    results = {}
    yf_tickers = {name: ticker for name, ticker in TICKERS.items() if ticker is not None}
    data_date = None
    try:
        raw = yf.download(list(yf_tickers.values()), period="2d", interval="1d", progress=False, auto_adjust=True)
        close = raw["Close"] if "Close" in raw.columns else raw
        if hasattr(close.index, "date"):
            data_date = str(close.index[-1].date())
        for name, ticker in yf_tickers.items():
            try:
                prices = close[ticker].dropna()
                if len(prices) >= 2:
                    prev, last = float(prices.iloc[-2]), float(prices.iloc[-1])
                    change = last - prev
                    pct = change / prev * 100
                    results[name] = {"value": last, "change": change, "pct": pct, "date": data_date}
                elif len(prices) == 1:
                    results[name] = {"value": float(prices.iloc[-1]), "change": None, "pct": None, "date": data_date}
                else:
                    results[name] = None
            except Exception:
                results[name] = None
    except Exception as e:
        st.error(f"データ取得エラー: {e}")

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

【メンターへの質問】
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
        "メンターへの質問": "question",
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

    # 今日の日付がすでにあればはじく
    existing_dates = [r[0] for r in all_values[1:]]
    if today in existing_dates:
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
st.title("📊 毎日の経済チェック")
st.caption(f"最終更新: {datetime.now().strftime('%Y年%m月%d日 %H:%M')}")

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

# ──── セクション1: 日米主要指数
st.subheader("📈 日米主要指数")
cols = st.columns(5)
for i, name in enumerate(["日経平均", "ダウ平均", "S&P500", "NASDAQ", "SOX（半導体）"]):
    with cols[i]:
        metric_card(name, data.get(name))

st.divider()

# ──── セクション2: 為替・債券
st.subheader("💱 為替・債券利回り")
cols = st.columns(3)
for i, name in enumerate(["ドル円", "米国10年債(%)", "日本10年債(%)"]):
    with cols[i]:
        metric_card(name, data.get(name))

st.divider()

# ──── セクション3: 資金逃避先
st.subheader("🛡️ 資金逃避先（有事の動き）")
cols = st.columns(3)
for i, name in enumerate(["原油(WTI)", "金", "ビットコイン"]):
    with cols[i]:
        metric_card(name, data.get(name))

st.divider()

# ──── セクション4: AIコメント
st.subheader("🤖 今日の市場コメント（AI自動生成）")

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
        st.markdown(f"**🙋 メンターへの質問：** {parsed['question']}")
        if parsed["answer"]:
            with st.expander("💡 回答を見る"):
                st.write(parsed["answer"])
    else:
        st.info(st.session_state.ai_comment)
    st.caption("※ 今日はすでに生成済みです")
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
st.subheader("📌 気になる銘柄（メンター面談用）")
st.caption("メンターとの面談で話したい銘柄やテーマを入力してください")

picks_input = st.text_area(
    "銘柄・テーマメモ",
    height=100,
    placeholder="例: NVIDIA（AI需要で急騰）、TSMC（SOX連動）、ドル円の動きが気になる",
    help="ここだけは手入力です。市場データを見て気になったことを書き留めましょう",
)

st.divider()

# ──── セクション6: 記録ボタン
st.subheader("📁 今日のデータを記録する")
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
st.subheader("📈 過去の推移グラフ")
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
