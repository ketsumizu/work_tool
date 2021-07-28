import time
import chromedriver_binary
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

#ブラウザ制御のためにここでドライバ読み込む
driver = webdriver.Chrome()

#目的のサイトへ行ってログイン
driver.get('http://〇.〇.〇.〇/login.html')                                                         #ウェブ上でアクセスするためのURL（デフォルトゲートウェイのIPを含めたURL）
driver.find_element_by_id('パスワードフォームのHTML要素ID').send_keys('ルータ管理画面パスワード')       #フォームを取得してパス入力
driver.find_element_by_class_name('ログインボタンのHTML要素クラス名').click()
driver.find_element_by_id('クリックしたいパネルのHTML要素ID').click()
time.sleep(1)                                                                                       #画面が実際に切り替わる前に要素をクリックしてしまうとエラーになるので少し待ち時間を作る
driver.find_element_by_class_name('クリックしたいパネルのHTML要素クラス名').click()
time.sleep(1)
driver.find_element_by_class_name('クリックしたいパネルのHTML要素クラス名').click()
time.sleep(5)
iframe = driver.find_element_by_id('入りたいiframeのHTML要素ID')                                     #iframeに入る処理が必要だったので以下２行で入る
driver.switch_to_frame(iframe)
driver.find_element_by_xpath('再起動ボタンのHTML要素xpath').click()                                #再起動ボタンをクリック