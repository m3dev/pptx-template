import unittest

from pptx_template.text import iterate_els

class MyTest(unittest.TestCase):

  def test_iterate_els(self):
    self.assertEqual(['1','2','3'], [id for id in iterate_els('{1}{2}{3}')])
    self.assertEqual(['def'], [id for id in iterate_els('abc{def}ghi')])
    self.assertEqual(['def'], [id for id in iterate_els('abc{{def}}def')])
    self.assertEqual([], [id for id in iterate_els('abcdef')])
    self.assertEqual([], [id for id in iterate_els('{}')])
    self.assertEqual([], [id for id in iterate_els('')])

if __name__ == '__main__':
    unittest.main()
