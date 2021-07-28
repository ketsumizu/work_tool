#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
# import subprocess
import PyPDF2
import glob


def deco(func):
    def wrapper(*args, **kwargs):
        print('--開始--')
        func(*args, **kwargs)
        print('--終了--')
    return wrapper

@deco
def main():

   #argvでコマンドライン引数を受け取ってプログラム内で使える
   args = sys.argv
   '''
   $ python ファイル名 a b
   args = sys.argv
   print args[0] #ファイル名
   print args[1] #a
   print args[2] #b
   '''

   #CLIで受け取った値をページの分割に使うためにリストにする             ＊＊＊＊＊＊＊バグ防止のためのちに数値バリデーション必要＊＊＊＊＊＊＊
   page_number = []
   for i in range(2, len(args)):
      page_number.append(args[i])

   #スキャンしたpdfを探して変数に格納。ここはもうちょっと上手いやり方を考える余地あり
   f = glob.glob("C:/Users/taked/Desktop/scan/*.pdf") #pdfのあるフォルダを指定
   target_pdf = f[0]

   #ファイル命名の為の材料を用意。具体的には処理しているファイルのタイトルをbasenameで受け取って、それをバラバラにして部品にする
   # basename = os.path.basename(target_pdf)[:-4]
   basename = args[1]
   basename_split = basename.split('_')

   #抽出したいページ数を受け取って抽出。名前をつけて１つのファイルとして保存する関数。命名規則は関数内のfile_name変数の行を参照
   def merge_page():
      merger = PyPDF2.PdfFileMerger()
      merger.append(target_pdf, pages=(start_page, end_page))
      file_name = left + '_' + '0' + str(right) + '.pdf'
      merger.write(file_name)
      merger.close

   #処理の開始ページと終了ページを指定しつつCLIで受け取った数値の個数分ループを走らせる。実際のpdf上ではstart_page+1枚目からend_page枚目までを分割処理する。
   start_page = 0
   for i, j in enumerate(range(1, len(page_number))):
      if page_number[j] == 'pass':
         continue
      else:
         left = basename_split[0]
         right = int(basename_split[1]) - i
         end_page = int(page_number[j])
         merge_page()
         start_page = end_page

'''
あとで使えそうなやつメモ
    Python内でコマンドを実行する
      cmd = subprocess.run(['command', 'argument']) # commandがコマンド、argumentが引数。必要なだけ並べる。
'''

if __name__ == "__main__":
   main()