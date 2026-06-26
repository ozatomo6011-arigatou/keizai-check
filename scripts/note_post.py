"""毎日決まった時刻に本番アプリを開き、スクショとnote用テキストをメールで送る"""
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from pathlib import Path

from playwright.sync_api import sync_playwright
from PIL import Image

APP_URL = os.environ["APP_URL"]
ICLOUD_EMAIL = os.environ["ICLOUD_EMAIL"]
ICLOUD_APP_PASSWORD = os.environ["ICLOUD_APP_PASSWORD"]
SCREENSHOT_PATH = Path("note_screenshot.png")


def capture() -> str:
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": 420, "height": 2200})
        page.goto(APP_URL, wait_until="networkidle", timeout=60000)
        page.wait_for_timeout(8000)
        try:
            page.wait_for_load_state("networkidle", timeout=15000)
        except Exception:
            pass

        # Streamlit Cloudは実際のアプリ本体を内部のiframeで配信しているため、
        # そのフレームを探して操作・スタイル適用する必要がある
        # 起床直後は一時的に複数フレームが存在することがあるため、testid要素が
        # 最も多い(=最終的に表示される)フレームを選ぶ
        app_frame = max(page.frames, key=lambda f: f.locator("[data-testid]").count())
        print("DEBUG chosen frame url:", app_frame.url, "testid count:", app_frame.locator("[data-testid]").count())

        app_frame.add_style_tag(content="""
            [data-testid='stSidebar'] { display: none !important; }
            div[data-testid="stMetricValue"] { font-size: 2.4rem !important; }
            div[data-testid="stMetricLabel"] { font-size: 1.4rem !important; }
            div[data-testid="stMetricDelta"] { font-size: 1.4rem !important; }
            h5 { font-size: 2.4rem !important; }
            h6 { font-size: 1.8rem !important; }
            div[data-testid="stCaptionContainer"] p { font-size: 1.3rem !important; }
        """)

        expander = app_frame.locator("text=note投稿用テキスト")

        gen_button = app_frame.locator("button", has_text="AIコメントを生成")
        if gen_button.count() and gen_button.first.is_visible():
            gen_button.first.click()
            try:
                expander.first.wait_for(state="visible", timeout=60000)
            except Exception:
                pass

        save_button = app_frame.locator("button", has_text="Googleスプレッドシートに保存")
        if save_button.count() and save_button.first.is_visible():
            save_button.first.click()
            page.wait_for_timeout(3000)

        comment_heading = app_frame.get_by_role("heading", name="今日の市場まとめ")
        try:
            comment_heading.first.wait_for(state="visible", timeout=30000)
        except Exception:
            pass

        # Streamlitの本体は独自のスクロール領域(stAppViewContainer等)を持つため、
        # window.scrollTo ではなく該当要素のscrollTopを直接リセットする
        reset_scroll_js = """
            document.querySelectorAll('*').forEach(el => {
                if (el.scrollTop > 0) el.scrollTop = 0;
            });
        """
        app_frame.evaluate(reset_scroll_js)
        page.evaluate(reset_scroll_js)
        page.evaluate("window.scrollTo(0, 0)")
        page.wait_for_timeout(500)

        comment_box = comment_heading.first.bounding_box() if comment_heading.count() else None
        scroll_y = page.evaluate("window.scrollY")
        absolute_y = comment_box["y"] + scroll_y if comment_box else None
        print("DEBUG comment_heading count:", comment_heading.count(), "box:", comment_box, "scroll_y:", scroll_y, "absolute_y:", absolute_y)

        page.screenshot(path=str(SCREENSHOT_PATH), full_page=True)

        if absolute_y and absolute_y > 50:
            img = Image.open(SCREENSHOT_PATH)
            cropped = img.crop((0, 0, img.width, int(absolute_y)))
            cropped.save(SCREENSHOT_PATH)
        else:
            print("DEBUG could not crop, keeping full_page screenshot")

        note_text = ""
        if expander.count():
            expander.first.click()
            page.wait_for_timeout(1000)
            code_block = app_frame.locator("pre").last
            if code_block.count():
                note_text = code_block.inner_text()

        browser.close()
        return note_text


def send_email(note_text: str):
    msg = MIMEMultipart()
    msg["Subject"] = "毎日の経済チェック - note投稿用"
    msg["From"] = ICLOUD_EMAIL
    msg["To"] = ICLOUD_EMAIL
    msg.attach(MIMEText(note_text or "（note用テキストを取得できませんでした。アプリを確認してください）", "plain"))

    with open(SCREENSHOT_PATH, "rb") as f:
        img = MIMEImage(f.read())
        img.add_header("Content-Disposition", "attachment", filename="screenshot.png")
        msg.attach(img)

    with smtplib.SMTP("smtp.mail.me.com", 587) as server:
        server.starttls()
        server.login(ICLOUD_EMAIL, ICLOUD_APP_PASSWORD)
        server.send_message(msg)


if __name__ == "__main__":
    text = capture()
    send_email(text)
