import math
import datetime

from copy import copy

from openpyxl.styles import Alignment, Font, Border, Side

CELL_STYLES = {
  'HEADER':{
    'alignment': Alignment(horizontal='center', vertical='center', wrap_text="true"),
    'font': Font(bold=True),
    'border':  Border(left=Side(style='thin'),
                       right=Side(style='thin'),
                       top=Side(style='thin'),
                       bottom=Side(style='thin'))
  },
  'TEXT':{
    'alignment': Alignment(horizontal='center', vertical='center'),
    'border':  Border(left=Side(style='thin'),
                       right=Side(style='thin'),
                       top=Side(style='thin'),
                       bottom=Side(style='thin'))
  },
  'TITLE':{
    'alignment': Alignment(horizontal='center', vertical='center'),
    'font': Font(bold=True, size=15)
  }
}

def insert_row(sheet, data, row_num, styles={}):
  col_idx = 1
  for d in data:
    cell = sheet.cell(row=row_num, column=col_idx, value=d or "")
    if "alignment" in styles.keys():
      cell.alignment = styles["alignment"]
    if "font" in styles.keys():
      cell.font = styles["font"]
    if 'border' in styles.keys():
      cell.border = styles['border']
    col_idx +=1


def copy_style(new_cell, cell):
  new_cell.font = copy(cell.font)
  new_cell.border = copy(cell.border)
  new_cell.fill = copy(cell.fill)
  new_cell.number_format = copy(cell.number_format)
  new_cell.protection = copy(cell.protection)
  new_cell.alignment = copy(cell.alignment)

def display_from_to(from_val, to_val):
  if from_val is None or to_val is None:
    return None
  if from_val == to_val:
    return str(from_val)
  else:
    return str(from_val) + ' - ' + str(to_val)

def display_time(t):
  a = ' AM'
  if t.hour > 12:
    a = ' PM'
  return str(t.hour%12) + ':' + str(t.minute) + a

def display_month(m):
  if 'MAR' in m :
    return 'MARCH'
  return m

def round_up(x, n):
  n = float(n)
  return int(math.ceil(x / n) * n)

def get_string_sub_code(code):
  c = str(code)
  if len(c) != 2:
    c = '0'+ c
  return c

def zerofy(n):
  if n < 10:
    return '0' + str(n)
  else:
    return str(n)

def time_plus(time, timedelta):
    start = datetime.datetime(
        2000, 1, 1,
        hour=time.hour, minute=time.minute, second=time.second)
    end = start + timedelta
    return end.time()



def split_arr_n(arr, x):
  n = int(math.ceil(len(arr)/float(x)))

  new_arr = []
  for i in range(0, n):
    for j in range(0, x):
      new_index = i+j*n
      if new_index < len(arr):
        new_arr.append(arr[new_index])
      # else:
      #   new_arr.append(None)

  return new_arr

def sort_by(items, columns):
  return multikeysort(items, columns)

def multikeysort(items, columns):
  from operator import itemgetter
  comparers = [ ((itemgetter(col[1:].strip()), -1) if col.startswith('-') else (itemgetter(col.strip()), 1)) for col in columns]
  def comparer(left, right):
      for fn, mult in comparers:
          result = cmp(fn(left), fn(right))
          if result:
              return mult * result
      else:
          return 0
  return sorted(items, cmp=comparer)

def group_by(input_data, keys):
  def _group(datum, k):
    data = {}
    for d in datum:
      if d[k] not in data.keys():
        data[d[k]] = []
      data[d[k]].append(d)
    return data

  if type(keys) in [str, unicode]:
    return _group(input_data, keys)

  data = _group(input_data, keys[0])
  for k, datum in data.iteritems():
    data[k] = _group(datum, keys[1])

  return data

def replace_all(string, d):
  s = string
  for k,v in d.items():
    s = s.replace(k, str(v))
  return s
