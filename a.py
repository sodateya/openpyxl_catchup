import openpyxl

# Excelファイルの読み込み
file_path = '/Users/takuyaishikawa/Desktop/OPE室/asasa.xlsm'
load_book = openpyxl.load_workbook(file_path,keep_vba=True)

# シートを指定（'最新リスト'のシートを仮定）
sheet = load_book['最新リスト']

# A1セルに新しい値を設定
sheet['A1'] = 's111kdjファsd?'

# 変更を保存
load_book.save(file_path)

print("A1セルの値が '12314?' に変更されました。")
