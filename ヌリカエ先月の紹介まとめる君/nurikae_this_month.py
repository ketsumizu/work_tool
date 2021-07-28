# #今回使用するライブラリ
import os
import csv
import glob
import time
import calendar
import openpyxl
from datetime import datetime
from selenium import webdriver
import chromedriver_binary
from selenium.webdriver.common.keys import Keys
from openpyxl.styles.fonts import Font
from openpyxl.styles.alignment import Alignment
from openpyxl.styles.borders import Border, Side
import win32com.client


#ブラウザ制御のためにここでドライバ読み込む
driver = webdriver.Chrome()
  # ↑　chromeブラウザを制御するためのドライバ「chromeedriver」のインストール先パスを指定


#目的のサイトに行き、ログイン認証を済ませる
driver.get('')                                           #行きたいサイト、今回はヌリカエのウェブサイトのURLを指定
driver.find_element_by_id('client_tool_account_identifier').send_keys('')        #フォームを取得してID入力
driver.find_element_by_id('client_tool_account_password').send_keys('')         #フォームを取得してパス入力　←これセキュリティ的にやべーね
driver.find_element_by_name('commit').click()                                           #フォームを送信してログイン


############################csvをサイトからとってきて開く################################################################################################################

#プログラム時の年月日などを保持
dt = datetime.now()


driver.find_element_by_link_text('紹介一覧').click()    #紹介一覧のページへ移動

driver.find_element_by_class_name('panel-title').click()    #「条件で絞り込む」のパネルを開く
time.sleep(0.1)   #パネルが開くまで明示的に待たないと怒られる仕様なので待つ


#例えばプログラム実行日が10月だったら9/1~9/30　11月だったら10/1~10/31を入力するようにしてる
start_date = driver.find_element_by_id('q_created_at_gteq')
if dt.month == 1:
    start_date.send_keys(dt.year - 1)
    start_date.send_keys(Keys.RIGHT)
else:
    start_date.send_keys(dt.year)
    start_date.send_keys(Keys.RIGHT)
if dt.month == 1:   #先月を参照するこの部分、今月が「1月」である場合の対策。抜かりないぜ。
    start_date.send_keys('12')
else:
    start_date.send_keys(dt.month - 1)
start_date.send_keys(Keys.RIGHT)
start_date.send_keys(1)

end_date = driver.find_element_by_id('q_created_at_lteq_end_of_day')
if dt.month == 1:
    end_date.send_keys(dt.year - 1)
    end_date.send_keys(Keys.RIGHT)
    end_date.send_keys(12)
    end_date.send_keys(Keys.RIGHT)
    end_date.send_keys(calendar.monthrange(dt.year - 1, 12)[1])   #先月末日を取得して入力
else:
    end_date.send_keys(dt.year)
    end_date.send_keys(Keys.RIGHT)
    end_date.send_keys(dt.month - 1)
    end_date.send_keys(Keys.RIGHT)
    end_date.send_keys(calendar.monthrange(dt.year, dt.month - 1)[1])   #先月末日を取得して入力

driver.find_element_by_id('csv-output-btn').find_element_by_class_name('btn').click()
time.sleep(3)


#csv開いてExcelに変換
csv_list = glob.glob('')  #ダウンロードcsvを探してる  ここはもうちょっと上手いやり方を考える余地あり
target_csv = csv_list[0]

wb2 = openpyxl.Workbook()
ws2 = wb2.active

with open('{}'.format(target_csv)) as f:    #この辺でcsvの中身をExcelに移動させる
    reader = csv.reader(f)
    for row in reader:
        ws2.append(row)


############################Excelファイルの中身を整える################################################################################################################

#セル幅揃える
ws2.column_dimensions['A'].width = 33
ws2.column_dimensions['B'].width = 30
ws2.column_dimensions['C'].width = 40
ws2.column_dimensions['D'].width = 7

#不要な列を削除
ws2.delete_cols(5, 5)

#罫線の準備
border = Border(top=Side(style='medium', color='c0c0c0'),
                bottom=Side(style='medium', color='c0c0c0'),
                left=Side(style='medium', color='c0c0c0'),
                right=Side(style='medium', color='c0c0c0'))

border2 = Border(top=Side(style='thick', color='000000'),
                bottom=Side(style='thick', color='000000'),
                left=Side(style='thick', color='000000'),
                right=Side(style='thick', color='000000'))

#A1から一番下までの範囲を中央ぞろえした
for cell1 in ws2['A1:D{}'.format(ws2.max_row)]:
    for cell2 in cell1:
        cell2.alignment = Alignment(horizontal='center', vertical = 'center', wrapText=True)
        cell2.font = Font(name = 'メイリオ', size = 14, bold = True)
        cell2.border = border


for cell3 in ws2['A1:D1']:
    for cell4 in cell3:
        cell4.border = border2


#印刷設定   pdfに変換するときにも影響する
# ws1.print_title_rows = '1:1'              # 印刷タイトル行
ws2.page_setup.orientation = 'portrait'  # 横は'landscape'
ws2.page_setup.fitToWidth = 1
ws2.page_setup.fitToHeight = 0
ws2.sheet_properties.pageSetUpPr.fitToPage = True


#Excelシートをセーブする
wb2.save('.xlsx')     #vscodeでこのPythonファイルを起動するならパスは 'ヌリカエ先月の紹介まとめる君/thisMonth.xlsx'  とする


#要らないのでcsv削除
os.remove(target_csv)


#ブラウザ閉じる
driver.close()


# #Excelファイルをpdfに変換     したければどうぞ的な
# excel = win32com.client.Dispatch("Excel.Application")
# file = excel.Workbooks.Open("C:/Users/taked/Desktop/unmanaged_python/thisMonth.xlsx")
# file.WorkSheets(1).Select()
# file.ActiveSheet.ExportAsFixedFormat(0,"C:/Users/taked/Desktop/unmanaged_python/thisMonth.pdf")
# file.close()