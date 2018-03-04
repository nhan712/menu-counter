from xlrd import open_workbook

class Menu_By_DayOfWeek(object):
  def __init__(self, day_of_week, name):
    self.day_of_week = day_of_week
    self.name = name

  def __str__(self):
    return ("Menu By Day of week:\n"
            " Day_of_week: {0}\n"
            " Name: {1}".format(self.day_of_week, self.name))

# Open the workbook
wb = open_workbook('Rotation menu 4 Circle 250118.xlsx')
# List sheet names, and pull a sheet by name
#
sheet_names = wb.sheet_names()
print('Sheet Names: ', sheet_names)
#
sheet = wb.sheet_by_name(sheet_names[0])
number_of_rows = sheet.nrows
number_of_columns = sheet.ncols
# collect day_of_week
day_of_week = []
for col in range(3, number_of_columns):
  cell_obj__day_of_week = sheet.cell(1,col)
  day_of_week.append(cell_obj__day_of_week.value)
# collect course
courses = []
for row in range(3, number_of_rows):
    for col in range(3, number_of_columns):
      cell_obj__courses = sheet.cell(row,col)
      if (cell_obj__courses.value != ""):
        courses.append(cell_obj__courses.value)
#
items = []
for d in day_of_week:
  for course in courses:
    item = Menu_By_DayOfWeek(d, course)
    items.append(item)
#
for item in items:
  print(item)