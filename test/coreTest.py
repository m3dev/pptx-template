import unittest

from pptx_template.core import iterate_ids

class MyTest(unittest.TestCase):

  def test_iterate_ids(self):
    self.assertEqual(['1','2','3'], [id for id in iterate_ids('{1}{2}{3}')])
    self.assertEqual(['def'], [id for id in iterate_ids('abc{def}ghi')])
    self.assertEqual(['def'], [id for id in iterate_ids('abc{{def}}def')])
    self.assertEqual([], [id for id in iterate_ids('abcdef')])
    self.assertEqual([], [id for id in iterate_ids('{}')])
    self.assertEqual([], [id for id in iterate_ids('')])

if __name__ == '__main__':
    unittest.main()
