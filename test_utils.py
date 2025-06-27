import unittest
from utils import parse_header, get_list

class TestUtils(unittest.TestCase):

    def test_parse_header(self):
        # Mock OpenPyXL row object
        class MockColumn:
            def __init__(self, value):
                self.value = value

        row = [MockColumn('dc.title'), MockColumn('dc.contributor.author'), MockColumn('dc.subject(;)')]
        header = parse_header(row)
        self.assertEqual(header[0]['element'], 'title')
        self.assertEqual(header[1]['qualifier'], 'author')
        self.assertEqual(header[2]['delimit'], ';')

    def test_get_list(self):
        self.assertEqual(get_list('a;b;c', ';'), ['a', 'b', 'c'])
        self.assertEqual(get_list('a|b|c', '|'), ['a', 'b', 'c'])
        self.assertEqual(get_list('a, b, c', ','), ['a', 'b', 'c'])
        self.assertEqual(get_list('a', ';'), ['a'])
        self.assertEqual(get_list('', ';'), [])
        self.assertEqual(get_list(None, ';'), [])

if __name__ == '__main__':
    unittest.main()
