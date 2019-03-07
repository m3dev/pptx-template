# generate_model files
# coding=utf-8

import sys
import codecs
import logging
import argparse
import json

import re
from io import open, StringIO

import openpyxl as xl
from six import iteritems, moves
from itertools import islice

from six import string_types
import numbers

import pptx_template.pyel as pyel

log = logging.getLogger()

def _build_tsv(rect_list, side_by_side=False, transpose=False, format_cell=False):
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
                    value = _format_cell_value(cell) if format_cell else cell.value
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

def _write_tsv(tsv, list_of_list):
     for row in list_of_list:
         for col, value in enumerate(row):
            if col != 0:
                tsv.write(u"\t")
            if value != None:
                tsv.write(u"%s" % value)
            else:
                tsv.write(u"")
         tsv.write(u"\n")

FRACTIONAL_PART_RE = re.compile(u"\.(0+)")

def _format_cell_value(cell):
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

def _extract_row(slides, xls, slide_id, el, value_cell, range_name, options):
    log.debug(" loading from xlsx: slide_id:%s EL:%s value:%s range:%s options:%s" % (slide_id, el, value_cell.value, range_name, options))
    number_format = _get_number_format_from_options(options)

    model_value = None
    if value_cell.value:
        model_value = _format_cell_value(value_cell)
    elif range_name:
        if range_name[0] == '=':
            rects = []
            for one_range in range_name[1:].split(','):
                parts = one_range.split('!')                # sheet!A1:C99 style
                sheet, coords = parts[0], parts[1]
                rects.append(xls[sheet][coords])
        else:
            rects = [xls[sheet][coords] for sheet, coords in xls.defined_names[range_name].destinations]

        array_mode = u"Array" in options
        tsv = _build_tsv(rects, side_by_side = u"SideBySide" in options, transpose = u"Transpose" in options, format_cell = array_mode)

        if array_mode:
            model_value = tsv
        else:
            tsv_body = StringIO()
            _write_tsv(tsv_body, tsv)
            model_value = {"tsv_body": tsv_body.getvalue(), "number_format": number_format, "xy_transpose": u"XYTranspose" in options}
            tsv_body.close()
    else:
         raise ValueError("One of value or range_name required.")

    return pyel.set_value(slides, u"%s.%s" % (slide_id, el), model_value)


def generate_whole_model(xls, slides):
    (xls, xls_formula, rows) = build_model_sheet_rows(xls)
    for data, formula in islice(rows, 1, None):
        slide_id, el, cell, range_name, options = data[0].value, data[1].value, data[2], formula[3].value, data[4].value
        if not slide_id or slide_id[0] == '#':
            continue
        options = options.split(' ,') if options else []
        slides = _extract_row(slides, xls, slide_id, el, cell, range_name, options)
    xls.close()
    xls_formula.close()
    return slides

def build_model_sheet_rows(xls_filename):
    """
    Builds row iterator for values from data_only mode and not data_only mode.
    Why we need this function is - To read cell value and its formula at the same time,
    Openpyxl needs to create two instances in different mode.
    """
    xls = xl.load_workbook(xls_filename, read_only=True, data_only=True)
    xls_formula = xl.load_workbook(xls_filename, read_only=True, data_only=False)
    model_sheet_data = xls['model']
    model_sheet_formula = xls_formula['model']
    rows = zip(model_sheet_data.rows, model_sheet_formula.rows)
    return (xls, xls_formula, rows)

def _get_number_format_from_options(options):
    for option in options:
        NUMBER_FORMAT_KEY = "NumberFormat:"
        if option.startswith(NUMBER_FORMAT_KEY):
            return option[len(NUMBER_FORMAT_KEY):]
    return None