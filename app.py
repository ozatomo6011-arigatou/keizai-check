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
    "S&P500":        "^GSPC",
    "NASDAQ":        "^IXIC",
    "ダウ平均":       "^DJI",
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
            return {"value": last, "change": last - prev, "pct": (last - prev) / prev * 100}
        elif len(df) == 1:
            return {"value": float(df.iloc[-1].iloc[10]), "change": None, "pct": None}
    except Exception:
        pass
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
                if len(prices) >= 2:
                    prev, last = float(prices.iloc[-2]), float(prices.iloc[-1])
                    change = last - prev
                    pct = change / prev * 100
                    results[name] = {"value": last, "change": change, "pct": pct}
                elif len(prices) == 1:
                    results[name] = {"value": float(prices.iloc[-1]), "change": None, "pct": None}
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

これを見て、初心者の経済学習者向けに以下をまとめてください（日本語・200字以内）：
1. 市場全体のムード（リスクオン/リスクオフ）
2. 注目すべき動きとその理由（推測でOK）
3. メンター面談で聞くと良さそうなポイント1つ"""

    client = anthropic.Anthropic(api_key=api_key)
    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=400,
        messages=[{"role": "user", "content": prompt}],
    )
    return message.content[0].text


# ──────────────────────────────
# Excel記録
# ──────────────────────────────
def save_to_excel(data: dict, comment: str, picks: str):
    today = datetime.now().strftime("%Y-%m-%d")
    headers = ["日付"] + list(TICKERS.keys()) + ["AIコメント", "気になる銘柄"]

    if EXCEL_FILE.exists():
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "経済日誌"
        ws.append(headers)
        header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

    row = [today]
    for name in TICKERS:
        info = data.get(name)
        row.append(round(info["value"], 4) if info else "N/A")
    row.append(comment)
    row.append(picks)
    ws.append(row)

    # 前日比で赤/緑色付け（値のみの列）
    last_row = ws.max_row
    for col_idx, name in enumerate(TICKERS.keys(), start=2):
        info = data.get(name)
        if info and info["pct"] is not None:
            cell = ws.cell(row=last_row, column=col_idx)
            if info["pct"] > 0:
                cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            elif info["pct"] < 0:
                cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    wb.save(EXCEL_FILE)


# ──────────────────────────────
# UI
# ──────────────────────────────
st.title("📊 毎日の経済チェック")
st.caption(f"最終更新: {datetime.now().strftime('%Y年%m月%d日 %H:%M')}")

# APIキー取得（Streamlit Secrets → 環境変数の順で読み込む）
api_key = st.secrets.get("ANTHROPIC_API_KEY", "") or os.environ.get("ANTHROPIC_API_KEY", "")

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

# ──── セクション1: 米主要指数
st.subheader("🇺🇸 米主要指数")
cols = st.columns(4)
for i, name in enumerate(["S&P500", "NASDAQ", "ダウ平均", "SOX（半導体）"]):
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

if "ai_comment" not in st.session_state:
    st.session_state.ai_comment = ""

if st.button("💬 AIコメントを生成", type="primary", disabled=not api_key):
    if not api_key:
        st.warning("サイドバーにAPIキーを入力してください")
    else:
        with st.spinner("Claudeが市場を分析中..."):
            try:
                st.session_state.ai_comment = generate_comment(data, api_key)
            except Exception as e:
                st.error(f"コメント生成エラー: {e}")

if st.session_state.ai_comment:
    st.info(st.session_state.ai_comment)

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

# ──── セクション6: Excel記録ボタン
st.subheader("📁 Excel に記録する")
col1, col2 = st.columns([1, 3])
with col1:
    if st.button("💾 今日のデータを保存", type="primary"):
        try:
            save_to_excel(data, st.session_state.ai_comment, picks_input)
            st.success(f"✅ 保存しました → {EXCEL_FILE.resolve()}")
        except Exception as e:
            st.error(f"保存エラー: {e}")
with col2:
    if EXCEL_FILE.exists():
        with open(EXCEL_FILE, "rb") as f:
            st.download_button(
                "⬇️ Excelをダウンロード",
                data=f,
                file_name="経済日誌.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

st.divider()

# ──── セクション7: 過去グラフ
st.subheader("📈 過去の推移グラフ")
if EXCEL_FILE.exists():
    try:
        df = pd.read_excel(EXCEL_FILE, index_col=0, parse_dates=True)
        if len(df) >= 2:
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
            st.info("グラフは2日分以上のデータが貯まると表示されます")
    except Exception as e:
        st.warning(f"グラフ表示エラー: {e}")
else:
    st.info("まだデータがありません。「Excelに保存」ボタンで記録を始めましょう！")
