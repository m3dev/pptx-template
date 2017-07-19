#
# coding=utf-8
import unittest
import sys
import os
from io import open

import openpyxl as xl

from pptx_template.xlsx_model import _build_tsv, _format_cell_value, generate_whole_model

class Cell:
    def __init__(self, value, number_format):
        self.value = value
        self.number_format = number_format

def _to_cells(list_of_list):
    return [[Cell(value, '') for value in list] for list in list_of_list]

class MyTest(unittest.TestCase):

    def test_build_tsv(self):
         tsv = _build_tsv([_to_cells([["Year","A","B"],["2016",100,200]])])
         self.assertEqual([["Year","A","B"],["2016",100,200]], tsv)

    def test_build_tsv_tranapose(self):
         tsv = _build_tsv([_to_cells([["Year","A","B"],["2016",100,200]])], transpose=True)
         self.assertEqual([["Year","2016"],["A",100],["B",200]], tsv)

    def test_build_tsv_side_by_side(self):
         tsv = _build_tsv([_to_cells([["Year","A"],["2016",100]]), _to_cells([["B"],[200]])], side_by_side=True)
         self.assertEqual([["Year","A","B"],["2016",100,200]], tsv)

    def test_format_cell_value(self):
        self.assertEqual(123.45678, _format_cell_value(Cell(123.45678, '')))
        self.assertEqual("123", _format_cell_value(Cell(123.45678, '0')))
        self.assertEqual("123.46", _format_cell_value(Cell(123.45678, '0.00')))
        self.assertEqual("123.5", _format_cell_value(Cell(123.45678, '0.0_')))
        self.assertEqual("12345.7%", _format_cell_value(Cell(123.45678, '0.0%_')))
        self.assertEqual("12345%", _format_cell_value(Cell(123.45678, '0%_')))

    def test_generate_whole_model(self):
        def read_expect(name):
            file_name = os.path.join(os.path.dirname(__file__), 'data2', name)
            f = open(file_name, mode = 'r', encoding = 'utf-8')
            result = f.read()
            f.close()
            return result

        xls_file = os.path.join(os.path.dirname(__file__), 'data2', 'in.xlsx')
        slides = generate_whole_model(xls_file, {})

        self.assertEqual(u'Hello!', slides['p01']['greeting']['en'])
        self.assertEqual(u'こんにちは！', slides['p01']['greeting']['ja'])
        self.assertEqual([
                    ['Season', u'売り上げ', u'利益', u'利益率'],
                    [u'春', 100, 50, 0.5],
                    [u'夏', 110, 60, 0.5],
                    [u'秋', 120, 70, 0.5],
                    [u'冬', 130,    0, 0.6],
        ], slides['p02']['array'])

        self.assertEqual(read_expect('p02-normal.tsv'), slides['p02']['normal']['tsv_body'])
        self.assertEqual(read_expect('p02-transpose.tsv'), slides['p02']['transpose']['tsv_body'])
        self.assertEqual(read_expect('p02-sidebyside.tsv'), slides['p02']['sidebyside']['tsv_body'])

if __name__ == '__main__':
    unittest.main()
