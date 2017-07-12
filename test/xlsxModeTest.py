import unittest
import sys
import os

from pptx_template.xlsxMode import build_tsv, format_cell_value

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


if __name__ == '__main__':
    unittest.main()
