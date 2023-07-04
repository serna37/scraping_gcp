from gapi import GApi
from webdriver import WDriver
from time import sleep
import os

driver = WDriver().getDriver()

# ======================
# const setting
# ======================
# Disneyランドの公式ページの、7月カレンダーを開き、7/6の状況を確認する
target_url = "https://www.tokyodisneyresort.jp/ticket/index/202307/"
selector1 = "#pbBlock5813935 > div.tab-detail.mt50 > div.tab-detail-section.is-active > table > tbody > tr:nth-child(2) > td:nth-child(5) > a > div.remaining"
atr1 = "className"
selector2 = "#pbBlock5813935 > div.tab-detail.mt50 > div.tab-detail-section.is-active > table > tbody > tr:nth-child(2) > td:nth-child(5) > a > div.type"

# ======================
# open
result = ""
screenshot_file = "watch_disney.png"
driver.get(target_url)
sleep(3)

# ======================
# get1

# どうやらクラス名でチケットの残りがどうかの表示を制御している
elm1 = driver.find_element_by_css_selector(selector1)
elm1_atr = elm1.get_attribute(atr1)
result += "\r\n混雑状況: "
result += "よゆーです・ω・v" if elm1_atr == "remaining" else "混んでいます"

# ======================
# get2
elm2 = driver.find_element_by_css_selector(selector2)
result += f"\r\n\r\n料金: {elm2.text}"

# ======================
# screen shot
# スクロールをして、スクショをとっておく
os.chdir(os.path.dirname(os.path.abspath(__file__)))
driver.execute_script("document.body.style.zoom='100%'")
driver.execute_script('window.scrollTo(0, 1250)')
driver.save_screenshot(screenshot_file)
GApi().upMailAttFile(screenshot_file, 'image/jpeg')
os.remove(screenshot_file)

# ======================
# end
driver.quit()

# ======================
# common mail GAS sync
to = "mailaddress"
cc = "mailaddress"
subject = "【watch】Disney 7/6"
body = f"Disney land 7/6 status\r\n{result}"

# スプシとDriveへの連携で、メール送信キューに任せる
GApi().sendMail(to, cc, subject, body, screenshot_file)

