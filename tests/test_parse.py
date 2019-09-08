import unittest
from parser.parse import Parse


class MyTestCase(unittest.TestCase):

    def test_something(self):
        parse = Parse('ExportOrderList201909081324.csv', 'xxx.xlsx', 'xxxx')
        parse.do_parse()
        self.assertEqual(True, True)


if __name__ == '__main__':
    unittest.main()
