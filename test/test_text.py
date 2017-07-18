import unittest

from pptx_template.text import iterate_els, find_el_position

class MyTest(unittest.TestCase):

    def test_iterate_els(self):
        self.assertEqual(['1','2','3'], [id for id in iterate_els('{1}{2}{3}')])
        self.assertEqual(['def'], [id for id in iterate_els('abc{def}ghi')])
        self.assertEqual(['def'], [id for id in iterate_els('abc{{def}}def')])
        self.assertEqual([], [id for id in iterate_els('abcdef')])
        self.assertEqual([], [id for id in iterate_els('{}')])
        self.assertEqual([], [id for id in iterate_els('')])

    def test_find_el_position(self):
        self.assertEqual(((0,0),(0,2)), find_el_position(['{a}'],'a'))
        self.assertEqual(((0,0),(0,2)), find_el_position(['{a}{a}'],'a'))
        self.assertEqual(((0,3),(2,1)), find_el_position(['abc{','de', 'f}ghi'],'def'))
        with self.assertRaises(ValueError):
            find_el_position([],'def')

if __name__ == '__main__':
    unittest.main()
