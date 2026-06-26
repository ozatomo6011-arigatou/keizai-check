"""毎日決まった時刻に本番アプリを開き、スクショとnote用テキストをメールで送る"""
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from pathlib import Path

from playwright.sync_api import sync_playwright

APP_URL = os.environ["APP_URL"]
ICLOUD_EMAIL = os.environ["ICLOUD_EMAIL"]
ICLOUD_APP_PASSWORD = os.environ["ICLOUD_APP_PASSWORD"]
SCREENSHOT_PATH = Path("note_screenshot.png")


def capture() -> str:
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": 1280, "height": 1600})
        page.goto(APP_URL, wait_until="networkidle", timeout=60000)
        page.wait_for_timeout(5000)
        page.add_style_tag(content="[data-testid='stSidebar'] { display: none !important; }")

        expander = page.locator("text=note投稿用テキスト")

        gen_button = page.locator("button", has_text="AIコメントを生成")
        if gen_button.count() and gen_button.first.is_visible():
            gen_button.first.click()
            try:
                expander.first.wait_for(state="visible", timeout=60000)
            except Exception:
                pass

        save_button = page.locator("button", has_text="Googleスプレッドシートに保存")
        if save_button.count() and save_button.first.is_visible():
            save_button.first.click()
            page.wait_for_timeout(3000)

        comment_heading = page.locator("text=今日の市場コメント")
        comment_box = comment_heading.first.bounding_box() if comment_heading.count() else None
        if comment_box:
            page.screenshot(
                path=str(SCREENSHOT_PATH),
                clip={"x": 0, "y": 0, "width": 1280, "height": comment_box["y"]},
            )
        else:
            page.screenshot(path=str(SCREENSHOT_PATH), full_page=True)

        note_text = ""
        if expander.count():
            expander.first.click()
            page.wait_for_timeout(1000)
            code_block = page.locator("pre").last
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
