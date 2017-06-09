import unittest

from pptx_template.pyel import eval_el, build_el

class MyTest(unittest.TestCase):

  def test_simple(self):
    self.assertEqual(2, eval_el('hoge.ora.1', {"hoge": {"ora": [1,2,3] }} ))
    self.assertEqual(3, eval_el('hoge.ora.2', {"hoge": {"ora": [1,2,3] }} ))

  def test_array(self):
    result = build_el( {"hoge": {"ora": [1,2,3] }, "bar": "Hello" })
    self.assertEqual(4, len(result))
    self.assertTrue(('hoge.ora.0',1) in result)
    self.assertTrue(('hoge.ora.1',2) in result)
    self.assertTrue(('hoge.ora.2',3) in result)
    self.assertTrue(('bar','Hello') in result)

  def test_root_array(self):
    result = build_el( [ 'abc', 1 ])
    self.assertEqual(("0", "abc"), result[0])

if __name__ == '__main__':
    unittest.main()
