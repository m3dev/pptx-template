import unittest
import sys
import os

from pptx_template.cli import main

class MyTest(unittest.TestCase):

  def test_simple(self):
      os.chdir(os.path.join(os.path.dirname(__file__), 'data'))
      sys.argv = ['myprog', '--out', 'out.pptx', '--template', 'in.pptx', '--model', 'model.json']
      main()

if __name__ == '__main__':
    unittest.main()
