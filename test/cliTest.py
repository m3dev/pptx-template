#
# coding=utf-8
import unittest
import sys
import os

from pptx_template.cli import main

def chdir(dir):
  current_dir = os.getcwd()
  try:
      os.chdir(os.path.join(os.path.dirname(__file__), dir))
      yield
  finally:
      os.chdir(current_dir)

class MyTest(unittest.TestCase):

  def test_simple(self):
      with chdir('data'):
        sys.argv = ['myprog', '--out', 'out.pptx', '--template', 'in.pptx', '--model', 'model.json', '--debug']
        main()

  def test_xlsx_mode(self):
      with chdir('data2'):
          sys.argv = ['myprog', '--out', 'out.pptx', '--template', 'in.pptx', '--model', 'in.xlsx', '--debug']
          main()

if __name__ == '__main__':
    unittest.main()
