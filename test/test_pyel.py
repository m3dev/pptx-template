import unittest

from pptx_template.pyel import eval_el, build_el, set_value

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

    def test_gen_model(self):
        result = set_value({}, 'greeting', 'hello')
        self.assertEqual(result['greeting'], 'hello')

    def test_gen_model_array(self):
        result = set_value({}, 'greeting.0', 'hello')
        self.assertEqual(result['greeting'][0], 'hello')

    def test_gen_model_array_index_not_exist(self):
        result = set_value({}, 'greeting.5', 'hello')
        self.assertEqual(result['greeting'][0], None)
        self.assertEqual(result['greeting'][5], 'hello')

    def test_gen_model_array_add(self):
        result = set_value({ 'greeting': []}, 'greeting.0', 'hello')
        self.assertEqual(result['greeting'][0], 'hello')

    def test_gen_model_array_add_index_not_exist(self):
        result = set_value({ 'greeting': [1,2]}, 'greeting.5', 'hello')
        self.assertEqual(result['greeting'][5], 'hello')

    def test_gen_model_dict(self):
        result = set_value({}, 'greeting.en', 'hello')
        self.assertEqual(result['greeting']['en'], 'hello')

    def test_gen_model_dict_add(self):
        result = set_value({'greeting': {'en': 'hola'}}, 'greeting.en', 'hello')
        self.assertEqual(result['greeting']['en'], 'hello')

    def test_gen_model_dict_start_with_dot(self):
        with self.assertRaises(ValueError):
            set_value({}, '.en', 'hello')

    def test_gen_model_with_invalid_dict(self):
        with self.assertRaises(ValueError):
            set_value({'greeting': 'hello'}, 'greeting.0', 'error')

    def test_gen_model_with_invalid_array(self):
        with self.assertRaises(ValueError):
            set_value({'greeting': [1,2]}, 'greeting.en', 'error')


if __name__ == '__main__':
    unittest.main()
