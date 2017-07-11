import unittest
import sys
import os

from pptx_template.xlsxMode import build_tsv

class MyTest(unittest.TestCase):

  def test_build_tsv(self):
     tsv = build_tsv([[["Year","A","B"],["2016",100,200]]])
     self.assertEqual([["Year","A","B"],["2016",100,200]], tsv)

  def test_build_tsv_tranapose(self):
     tsv = build_tsv([[["Year","A","B"],["2016",100,200]]], transpose=True)
     self.assertEqual([["Year","2016"],["A",100],["B",200]], tsv)

  def test_build_tsv_side_by_side(self):
     tsv = build_tsv([[["Year","A"],["2016",100]],[["B"],[200]]], side_by_side=True)
     self.assertEqual([["Year","A","B"],["2016",100,200]], tsv)

if __name__ == '__main__':
    unittest.main()
