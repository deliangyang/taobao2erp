import unittest
from parse.parse import Parse


class MyTestCase(unittest.TestCase):

    def test_something(self):
        parse = Parse('ExportOrderList201909081324.csv', 'xxx.xlsx', 'xxxx')
        parse.do_parse()
        self.assertEqual(True, True)

    def test_extract_code(self):
        title = '央美清华素描书入门基础教程9787229019570王建才著李家友主编'
        parse = Parse('ExportOrderList201909081324.csv', 'xxx.xlsx', 'xxxx')
        code, ok = parse.extract_erp_code(title)
        self.assertEqual('9787229019570', code)
        self.assertEqual(ok, True)
        title = '央美清华素描书入门基础教程王建才著李家友主编'
        code, ok = parse.extract_erp_code(title)
        self.assertEqual(None, code)
        self.assertEqual(ok, False)


if __name__ == '__main__':
    unittest.main()
