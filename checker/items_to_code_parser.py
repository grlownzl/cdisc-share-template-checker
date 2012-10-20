import xlrd
import openpyxl
import logging

def si(content):
  """
  If it gets a None, return '', else return cleaned string
  """
  import types
  if isinstance(content, openpyxl.cell.Cell):
    if isinstance(content.value, types.NoneType):
      return unicode('')
    else:
      return unicode(content.value).strip()
  elif isinstance(content, xlrd.sheet.Cell):
    if content.ctype == xlrd.sheet.XL_CELL_EMPTY:
      return unicode('')
    else:
      return unicode(content.value).strip()
  return unicode(content).strip()
    
class UniqueItemsToCodeParser(object):
  
  def __init__(self, codes=[]):
    # takes a list of dicts containing the current state, adds processing, but cuts db
    self.codes = dict([[x.get('name'), x] for x in codes])
    
  def load_from_mem(self, mem_obj):
    codes = {}
    try:
      workbook = xlrd.open_workbook(file_contents=mem_obj.file.read())
    except Exception, e:
      logging.exception("Couldn't open terminology sheet: %s" % e)
      return {}
    for sheet_name in workbook.sheet_names():
      sheet = workbook.sheet_by_name(sheet_name)
      cols = [si(x) for x in sheet.row_values(0)]
      for row in range(1, sheet.nrows):
        content = dict(zip(cols, [si(x) for x in sheet.row_values(row)]))
        code = codes.setdefault(content['Field'], {})
        code['name'] = content['Field']
        if 'Context' in content:
          code['terminology_type'] = content['Context']
        elif 'Code' in content:
          code['code'] = content['Code']
    return codes
