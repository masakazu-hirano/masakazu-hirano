## 🖥️ Portfolio（ポートフォリオ）

#### 📌 Excelファイルのシートを複製して、日付（降順）に並び替える処理

> [!NOTE]
> ■ 毎日（一日 1シート）複製元となるシートからシート名（YYYY-MM-DD）で複製する。  
> ■ 複製されたシート一覧を日付（降順）で並び替える。

```Python
import logging
import openpyxl

from datetime import date, timedelta
from openpyxl import Workbook

# シート一覧を日付（降順）で並び替える。
def Sort_Date_DESC(excel_file: Workbook) -> Workbook:
	new_sheetname_list: list = []
	sheetname_list: list[str] = excel_file.sheetnames
	for sheetname in sheetname_list:
		try:
			sheetname_datetime_object = datetime.strptime(sheetname, '%Y-%m-%d')
			new_sheetname_list.append(
				date(sheetname_datetime_object.year, sheetname_datetime_object.month, sheetname_datetime_object.day)
			)
		except ValueError:
			# 複製元となるシートをスキップ
			continue

	new_sheetname_list = sorted(new_sheetname_list, reverse = False)
	for sheetname in new_sheetname_list:
		sheetname = sheetname.strftime(format = '%Y-%m-%d')
		if sheetname in sheetname_list:
			sheet = excel_file[sheetname]
			excel_file._sheets.remove(sheet)
			excel_file._sheets.insert(0, sheet)

	return excel_file

if __name__ == '__main__':
	logging.basicConfig(
		level = logging.INFO,
		format = '[{levelname}]: {message}',
		style = '{'
	)

	excel_file: Workbook = openpyxl.load_workbook(
		filename = "ファイルパス（.xlsx）"
		rich_text = True,
		keep_links = False,
		data_only = False,
		read_only = False,
		keep_vba = False
	)

	current_date: date = date.today()	# 本ファイルを実行した日付
	execute_date: date = current_date	# シートを複製する日付
	while execute_date > (current_date - timedelta(days = "何日前までのシートを作成するか？")):
		sheet_name: str = execute_date.strftime(format = '%Y-%m-%d')
		try:
			excel_file[sheet_name]
		except KeyError:
			# シートが存在しなければ、複製して保存。
			duplicate_sheet: Worksheet = excel_file.copy_worksheet(excel_file["シート名（複製元）"])
			duplicate_sheet.title = sheet_name
			excel_file._sheets.insert(0, report_file._sheets.pop())
			excel_file.save(filename = "ファイルパス（.xlsx）"

		execute_date -= timedelta(days = 1)

	excel_file = Sort_Date_DESC(excel_file):
	excel_file.save(filename = "ファイルパス（.xlsx）"

	logging.info(msg = '処理が正常に終了しました。')
```
