#
# coding=utf-8
import unittest
import sys
import os
import tempfile
import shutil
from io import open

from itertools import islice

from pptx_template.xlsxMode import build_tsv, format_cell_value, generate_whole_model
import openpyxl as xl

class Cell:
  def __init__(self, value, number_format):
      self.value = value
      self.number_format = number_format

def _to_cells(list_of_list):
  return [[Cell(value, '') for value in list] for list in list_of_list]

class MyTest(unittest.TestCase):

  def test_build_tsv(self):
     tsv = build_tsv([_to_cells([["Year","A","B"],["2016",100,200]])])
     self.assertEqual([["Year","A","B"],["2016",100,200]], tsv)

  def test_build_tsv_tranapose(self):
     tsv = build_tsv([_to_cells([["Year","A","B"],["2016",100,200]])], transpose=True)
     self.assertEqual([["Year","2016"],["A",100],["B",200]], tsv)

  def test_build_tsv_side_by_side(self):
     tsv = build_tsv([_to_cells([["Year","A"],["2016",100]]), _to_cells([["B"],[200]])], side_by_side=True)
     self.assertEqual([["Year","A","B"],["2016",100,200]], tsv)

  def test_format_cell_value(self):
      self.assertEqual(123.45678, format_cell_value(Cell(123.45678, '')))
      self.assertEqual("123", format_cell_value(Cell(123.45678, '0')))
      self.assertEqual("123.46", format_cell_value(Cell(123.45678, '0.00')))
      self.assertEqual("123.5", format_cell_value(Cell(123.45678, '0.0_')))
      self.assertEqual("12345.7%", format_cell_value(Cell(123.45678, '0.0%_')))
      self.assertEqual("12345%", format_cell_value(Cell(123.45678, '0%_')))

  def test_generate_whole_model(self):
      def read_expect(name):
          file_name = os.path.join(os.path.dirname(__file__), 'data2', name)
          f = open(file_name, mode = 'r', encoding = 'utf-8')
          result = f.read()
          f.close()
          return result

      def read_result(name):
          f = open(os.path.join(temp_dir, name), mode = 'r', encoding = 'utf-8')
          result = f.read()
          f.close()
          return result

      xls_file = os.path.join(os.path.dirname(__file__), 'data2', 'in.xlsx')
      xls = xl.load_workbook(xls_file, read_only=True, data_only=True)
      model_sheet = xls['model']

      temp_dir = tempfile.mkdtemp()
      current_dir = os.getcwd()
      try:
          os.chdir(temp_dir)

          slides = generate_whole_model(xls, islice(model_sheet.rows, 1, None), {})

          self.assertEqual({u'file_name': 'p02-normal.tsv'}, slides['p02']['normal'])
          self.assertEqual({u'file_name': 'p02-sidebyside.tsv'}, slides['p02']['sidebyside'])
          self.assertEqual({u'file_name': 'p02-transpose.tsv'}, slides['p02']['transpose'])
          self.assertEqual(u'Hello!', slides['p01']['greeting']['en'])
          self.assertEqual(u'こんにちは！', slides['p01']['greeting']['ja'])
          self.assertEqual([['Season', u'売り上げ', u'利益', u'利益率'],[u'春', 100, 50, 0.5],[u'夏', 110, 60, 0.5],[u'秋', 120, 70, 0.5]], slides['p02']['array'])

          self.assertEqual(read_expect('p02-normal.tsv'), read_result('p02-normal.tsv'))
          self.assertEqual(read_expect('p02-transpose.tsv'), read_result('p02-transpose.tsv'))
          self.assertEqual(read_expect('p02-sidebyside.tsv'), read_result('p02-sidebyside.tsv'))

      finally:
          os.chdir(current_dir)
          shutil.rmtree(temp_dir)

if __name__ == '__main__':
    unittest.main()
