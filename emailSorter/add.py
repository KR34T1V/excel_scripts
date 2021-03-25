from search import search

def add(wb_to, wb_fra, col, email):
  if (not wb_to or not wb_fra or not col or not email):
    return None
  
  ws_fra = wb_fra.active
  ws_to = wb_to.active
  mc_fra = ws_fra.max_column
  mr_to = ws_to.max_row
  entry = search(wb_to, col, email)
  if (not entry):
    fra_entry = search(wb_fra, col, email)
    if (not fra_entry):
      print('Could not find ' + email)
      return None
    fra_entry = int(fra_entry)
    print('Inserting 1st Instace')
    for j in range (1, mc_fra + 1):
      read = ws_fra.cell(row = fra_entry, column = j)
      ws_to.cell(row = mr_to +1, column = j).value = read.value
    ws_to.cell(row = mr_to +1, column = mc_fra + 1).value = 1
  else:
    entry = int(entry)
    print('Updating Listed Entry')
    ws_to.cell(row = entry, column = mc_fra + 1).value = int(ws_to.cell(row = entry, column = mc_fra + 1).value) + 1
  print('')