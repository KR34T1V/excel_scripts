def source(wb, next):
  if (not wb):
    return None
  ws = wb.active
  mr = ws.max_row
  if (next <= mr):
    if ( not ws['A' + str(next)].value):
      return None
    else:
      return ws['A' + str(next)].value
  else:
    return -1