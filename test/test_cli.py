#
# coding=utf-8
import unittest
import sys
import os
import logging

from pptx_template.cli import main

BASE_DIR = os.getcwd()

log = logging.getLogger()
handler = logging.StreamHandler()
handler.setLevel(logging.DEBUG)
log.addHandler(handler)

class MyTest(unittest.TestCase):
    def tearDown(self):
        os.chdir(BASE_DIR)

    def test_simple(self):
        os.chdir(os.path.join(BASE_DIR, 'test', 'data'))
        sys.argv = ['myprog', '--out', 'out.pptx', '--template', 'in.pptx', '--model', 'model.json', '--debug']
        main()

    def test_invalid_csv(self):
        os.chdir(os.path.join(BASE_DIR, 'test', 'data'))
        sys.argv = ['myprog', '--out', 'out.pptx', '--template', 'in.pptx', '--model', 'model_invalid.json', '--debug']
        with self.assertRaises(ValueError):
            main()

    def test_xlsx_mode(self):
        os.chdir(os.path.join(BASE_DIR, 'test', 'data2'))
        sys.argv = ['myprog', '--out', 'out.pptx', '--template', 'in.pptx', '--model', 'in.xlsx', '--debug']
        main()

if __name__ == '__main__':
    unittest.main()
