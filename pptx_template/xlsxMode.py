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

from six import string_types
import numbers

import pptx_template.pyel as pyel

log = logging.getLogger()

def build_tsv(rect_list, side_by_side=False, transpose=False, format_cell=False):
    """
    Excel の範囲名（複数セル範囲）から一つの二次元配列を作る
    rect_list:    セル範囲自体の配列
    side_by_side: 複数のセル範囲を横に並べる。指定ない場合はタテに並べる
    transpose:    結果を行列入れ替えする(複数範囲を結合した後で処理する)
    """
    result = []
    for rect_index, rect in enumerate(rect_list):
        for row_index, row in enumerate(rect):
            line = []
            for cell in row:
                value = cell
                if not cell:
                    value = None
                elif hasattr(cell, 'value'):
                    value = format_cell_value(cell) if format_cell else cell.value
                else:
                    raise ValueError("Unknown type %s for %s" % (type(cell), cell))
                line.append(value)
            if side_by_side and rect_index > 0:
                result[row_index].extend(line)
            else:
                result.append(line)

    if transpose:
        result = [list(row) for row in moves.zip_longest(*result, fillvalue=None)] # idiom for transpose

    return result

def write_tsv(file_name, list_of_list):
     log.info("writing tsv %s..." % file_name)
     tsv = open(file_name, encoding='utf-8', mode='w')
     for row in list_of_list:
         for col, value in enumerate(row):
            if col != 0:
                tsv.write("\t")
            if value != None:
                tsv.write(str(value))
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
    value, unit = cell.value, ''
    format = cell.number_format if cell.number_format else ''

    if not isinstance(value, numbers.Number):
        return value

    if '%' in format:
        value, unit = value * 100, '%'

    match = FRACTIONAL_PART_RE.search(format)
    if match:
        fraction_format = "%%.%df%%s" % len(match.group(1))
        return fraction_format % (value, unit)
    elif '0' in format:
        return "%d%s" % (value, unit)
    else:
        return value

def extract_row(slides, xls, slide_id, el, cell, range_name, options):
  log.debug("slide_id:%s EL:%s value:%s range:%s options:%s" % (slide_id, el, cell.value, range_name, options))

  model_value = None
  if cell.value:
      model_value = format_cell_value(cell)
  elif range_name:
      file_name = u"%s-%s.tsv" % (slide_id, el)
      rects = [xls[sheet][cords] for sheet, cords in xls.defined_names[range_name].destinations]
      array_mode = u"Array" in options
      tsv = build_tsv(rects, side_by_side = u"SideBySide" in options, transpose = u"Transpose" in options, format_cell = array_mode)

      if array_mode:
          model_value = tsv
      else:
          write_tsv(file_name, tsv)
          model_value = {"file_name": file_name}
  else:
       raise ValueError("One of value or range_name required.")

  return pyel.set_value(slides, u"%s.%s" % (slide_id, el), model_value)

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
      slide_id, el, cell, range_name, options = row[0].value, row[1].value, row[2], row[3].value, row[4].value
      if not slide_id or slide_id[0] == '#':
          continue
      options = options.split(' ,') if options else []
      slides = extract_row(slides, xls, slide_id, el, cell, range_name, options)

  log.info(u"writing model data:%s" % {"slides": slides})
  model_file = open('model.json', mode='w', encoding='utf-8')
  json.dump({"slides":slides}, model_file)
  model_file.close()


if __name__ == '__main__':
  main()
