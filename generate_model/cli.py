# generate_model
# coding=utf-8

import sys
import codecs
import logging
import argparse
import json

from io import open
import openpyxl as xl
from six import iteritems
from itertools import islice

import pyel

log = logging.getLogger()

def write_tsv(file_name, rows_list):
     log.info("writing tsv %s..." % file_name)
     tsv = open(file_name, encoding='utf-8', mode='w')
     for rows in rows_list:
       for row in rows:
         for col, cell in enumerate(row):
            if col != 0:
                tsv.write("\t")
            if cell.value != None:
                tsv.write(str(cell.value))
            else:
                tsv.write('')
         tsv.write("\n")
     tsv.close()

def extract_row(slides, xls, slide_id, el, value, range_name):
  log.debug("%s %s %s %s" % (slide_id, el, value, range_name))
  if value == None and range_name != None:
      file_name = u"%s-%s.tsv" % (slide_id, el)
      destinations = xls.defined_names[range_name].destinations
      write_tsv(file_name, [xls[sheet][cords] for sheet, cords in destinations])
      value = {"file_name": file_name}
  return pyel.set_value(slides, u"%s.%s" % (slide_id, el), value)

def main():
  if sys.version_info[0] == 2:
    # sys.stdout = codecs.getwriter('utf-8')(sys.stdout)
    reload(sys)
    sys.setdefaultencoding('utf-8')

  log = logging.getLogger()
  handler = logging.StreamHandler()
  handler.setLevel(logging.DEBUG)
  log.addHandler(handler)

  parser = argparse.ArgumentParser(description = 'Generate model.json from Excel')
  parser.add_argument('--xlsx',      help='file name of Excel book', required=True)
  parser.add_argument('--debug',      action='store_true', help='output verbose log')
  opts = parser.parse_args()

  if opts.debug:
    log.setLevel(logging.DEBUG)
  else:
    log.setLevel(logging.INFO)

  xls = xl.load_workbook(opts.xlsx, read_only=True, data_only=True)
  model_sheet = xls['model']

  slides = {}
  for row in islice(model_sheet.rows, 1, None):
      slide_id, el, value, range_name = row[0].value, row[1].value, row[2].value, row[3].value
      slides = extract_row(slides, xls, slide_id, el, value, range_name)

  log.info(u"writing model data:%s" % {"slides": slides})
  model_file = open('model.json', mode='w', encoding='utf-8')
  json.dump({"slides":slides}, model_file)
  model_file.close()


if __name__ == '__main__':
  main()
