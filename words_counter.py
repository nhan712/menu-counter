from xlrd import open_workbook
import re

class Form(object):
  def __init__(self, ten_san_pham, ten_mon, ten_sheet, thu, bua):
    self.ten_san_pham = ten_san_pham
    self.ten_mon = ten_mon
    self.ten_sheet = ten_sheet
    self.thu = thu
    self.bua = bua
  def __str__(self):
    return ("Form:\n"
            " Tên sản phẩm: {0}\n"
            " Tên món: {1}\n"
            " Tên sheet: {2}\n"
            " Thứ: {3}\n"
            " Bữa: {4}\n".format(self.ten_san_pham, self.ten_mon, self.ten_sheet, self.thu, self.bua))

# Find sheet name by keyword
def find_recipe_name_by_keyword(sheets, keyword):
  for sheet in sheets:
    number_of_rows = sheet.nrows
    number_of_cols = sheet.ncols
    if sheet_names.index(sheet.name) > 3:
      for row in range(1, number_of_rows):
        if number_of_cols > 3:
          cell_obj__other = sheet.cell(row,3)
          value = str(cell_obj__other.value).lower()
          if value == keyword:
            return sheet.name

# Find item in menu
def find_item_in_menu(product_name, sheets, item, rt):
  for sheet in sheets:
    number_of_rows = sheet.nrows
    number_of_cols = sheet.ncols
    for row in range(3, number_of_rows):
      for col in range(3, number_of_cols):
        cell_obj__other = sheet.cell(row,col)
        value = str(cell_obj__other.value).lower()
        if value.split("\n")[0].strip().find(item.strip()) > -1:
          day_of_week = ""
          if sheet.cell(1,col).value.find("Thứ") > -1:
            day_of_week = sheet.cell(1,col).value
          if sheet.cell(2,col).value.find("Thứ") > -1:
            day_of_week = sheet.cell(2,col).value
          day = str(sheet.cell(0,3).value) or str(sheet.cell(1,3).value)
          rt.append(Form(product_name, item.strip(), sheet.name, day_of_week, day))

# Open the workbook
wb = open_workbook('TVTĐ with recipes 03022018.xlsx')
# List sheet names, and pull a sheet by name
#1
sheets = wb.sheets()
sheet_names = wb.sheet_names()
products = []
# items = []
for sheet in sheets:
  number_of_rows = sheet.nrows
  number_of_cols = sheet.ncols
  if sheet_names.index(sheet.name) > 3:
    for row in range(1, number_of_rows):
      if number_of_cols > 3:
        cell_obj__other = sheet.cell(row,3)
        value = str(cell_obj__other.value)
        products.append(value.lower())
#2
file = open("form.txt","w",encoding='utf-8')
wb__form = open_workbook('Form tính tần suất thực phẩm.xlsx')
sheet__form = wb__form.sheet_by_index(0)
for row in range(2, sheet__form.nrows):
  items = []
  product_name = str(sheet__form.cell(row,0).value).lower().strip()
  for p in products:
    if p.find(product_name) > -1:
      items.append(find_recipe_name_by_keyword(sheets, p))
      #
      wb__menu = open_workbook('Rotation menu 4 Circle 250118.xlsx')
      sheets__menu = wb__menu.sheets()
      rt = []
      for item in items:
        item = re.sub(r'\d|[.]', '', item)
        find_item_in_menu(product_name, sheets__menu, item.lower(), rt)
      #
      for item in rt:
        file.write(str(item) + "----------------------\n")
file.close()