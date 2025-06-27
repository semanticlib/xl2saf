import unittest
from utils import parse_header, get_list

class TestUtils(unittest.TestCase):

    def test_parse_header(self):
        class MockColumn:
            def __init__(self, value):
                self.value = value

        # Test with various header formats
        row = [
            MockColumn('dc.title'),
            MockColumn('dc.contributor.author'),
            MockColumn('dc.subject(;)'),
            MockColumn('dc.description.abstract'),
            MockColumn('dc.date.issued'),
            MockColumn('dc.publisher')
        ]
        header = parse_header(row)
        
        # Assertions for each header
        self.assertEqual(header[0]['element'], 'title')
        self.assertEqual(header[0]['qualifier'], 'none')
        self.assertEqual(header[0]['delimit'], '')
        
        self.assertEqual(header[1]['element'], 'contributor')
        self.assertEqual(header[1]['qualifier'], 'author')
        self.assertEqual(header[1]['delimit'], '')
        
        self.assertEqual(header[2]['element'], 'subject')
        self.assertEqual(header[2]['qualifier'], 'none')
        self.assertEqual(header[2]['delimit'], ';')
        
        self.assertEqual(header[3]['element'], 'description')
        self.assertEqual(header[3]['qualifier'], 'abstract')
        self.assertEqual(header[3]['delimit'], '')

    def test_get_list(self):
        # Test with different delimiters
        self.assertEqual(get_list('a;b;c', ';'), ['a', 'b', 'c'])
        self.assertEqual(get_list('a|b|c', '|'), ['a', 'b', 'c'])
        self.assertEqual(get_list('a, b, c', ','), ['a', 'b', 'c'])
        
        # Test with single value
        self.assertEqual(get_list('a', ';'), ['a'])
        
        # Test with empty and None values
        self.assertEqual(get_list('', ';'), [])
        self.assertEqual(get_list(None, ';'), [])
        
        # Test with extra whitespace
        self.assertEqual(get_list(' a ; b ; c ', ';'), ['a', 'b', 'c'])

if __name__ == '__main__':
    unittest.main()
