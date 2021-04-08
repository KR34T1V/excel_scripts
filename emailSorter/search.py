import sys
import time

def search(wb, column, value):
  if (not wb or not column or not value):
    return None
  value = value.strip()
  ws = wb.active
  mr = ws.max_row
  i = 1
  while (i <= mr):
    current_cell = str(ws[column + str(i)].value)
    if (current_cell == None):
     i = i + 1
     continue
    current_cell = current_cell.strip()
    if (current_cell == value):
      print('Match Found:', current_cell)
      return (str(i))
    i = i + 1
