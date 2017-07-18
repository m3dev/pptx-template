#
# coding=utf-8
import unittest
import sys
import os
import logging

from pptx_template.cli import main

class MyTest(unittest.TestCase):

    def test_simple(self):
        current_dir = os.getcwd()
        os.chdir(os.path.join(os.path.dirname(__file__), 'data'))
        sys.argv = ['myprog', '--out', 'out.pptx', '--template', 'in.pptx', '--model', 'model.json', '--debug']
        main()
        os.chdir(current_dir)

    def test_xlsx_mode(self):
        current_dir = os.getcwd()
        os.chdir(os.path.join(os.path.dirname(__file__), 'data2'))
        sys.argv = ['myprog', '--out', 'out.pptx', '--template', 'in.pptx', '--model', 'in.xlsx', '--debug']
        main()
        os.chdir(current_dir)

if __name__ == '__main__':
    unittest.main()
