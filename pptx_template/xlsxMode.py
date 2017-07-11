# generate_model
# coding=utf-8

import sys
import codecs
import logging
import argparse
import json

import re
from io import open

import openpyxl as xl
from six import iteritems, moves
from itertools import islice

import pptx_template.pyel as pyel

log = logging.getLogger()

def build_tsv(rect_list, side_by_side=False, transpose=False):
    """
    Excel の範囲名（複数セル範囲）から一つの二次元配列を作る
    rect_list:    セル範囲自体の配列
    side_by_side: 複数のセル範囲を横に並べる。指定ない場合はタテに並べる
    transpose:    結果を行列入れ替えする(複数範囲を結合した後で処理する)
    """
    result = []
    for rect_index, rect in enumerate(rect_list):
        for row_index, row in enumerate(rect):
            if side_by_side and rect_index > 0:
                line = result[row_index]
                for cell in row:
                    line.append(cell)
            else:
                result.append(list(row))

    if transpose:
        result = [list(row) for row in moves.zip_longest(*result, fillvalue=None)] # idiom for transpose

    return result

def write_tsv(file_name, list_of_list):
     log.info("writing tsv %s..." % file_name)
     tsv = open(file_name, encoding='utf-8', mode='w')
     for row in list_of_list:
         for col, cell in enumerate(row):
            if col != 0:
                tsv.write("\t")
            if cell.value != None:
                tsv.write(str(cell.value))
            else:
                tsv.write('')
         tsv.write("\n")
     tsv.close()

FRACTIONAL_PART_RE = re.compile(u"\.(0+)")

def format_cell_value(cell):
    """
    for expample, a value 123.4567 will be formatted along with its cell.number_format:
      0     -> "123"
      0.00  -> "123.46"
      0%    -> "12345%"
      0.00% -> "12345.68%"
      other -> 123.4567     # numeric type
    """
    value = cell.value
    format = cell.number_format
    unit = ''
    if '%' in cell.format:
        value = value * 100
        unit = '%'

    match = FRACTIONAL_PART_RE.search(format)
    if match:
        fraction_format = "%%.%df%%s" % len(match.group(1))
        return fraction_format % (value, unit)
    elif '0' in format:
        return "%d%s" % (value, unit)
    else:
        return value

def extract_row(slides, xls, slide_id, el, value, range_name, options):
  log.debug("slide_id:%s EL:%s value:%s range:%s options:%s" % (slide_id, el, value, range_name, options))
  if value == None and range_name != None:
      file_name = u"%s-%s.tsv" % (slide_id, el)
      destinations = xls.defined_names[range_name].destinations
      tsv = build_tsv([xls[sheet][cords] for sheet, cords in destinations], side_by_side = 'S' in options, transpose = 'T' in options)
      write_tsv(file_name, tsv)
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
      slide_id, el, value, range_name, options = row[0].value, row[1].value, row[2].value, row[3].value, row[4].value
      options = options.split(' ,') if options else []
      slides = extract_row(slides, xls, slide_id, el, value, range_name, options)

  log.info(u"writing model data:%s" % {"slides": slides})
  model_file = open('model.json', mode='w', encoding='utf-8')
  json.dump({"slides":slides}, model_file)
  model_file.close()


if __name__ == '__main__':
  main()
