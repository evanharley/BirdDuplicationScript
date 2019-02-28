import unittest
from BirdDuplicationScript import BirdDuplicator
import openpyxl

class Test_test1(unittest.TestCase):
    def setUp(self):
        self.dupe = BirdDuplicator()
        return super().setUp()

    def test_workbook_exists(self):
        self.assertIsInstance(self.dupe.file, openpyxl.Workbook)

    def test_data_dictionary(self):
        self.dupe._create_data_dictionary()
        test_values = self.dupe.item_ids
        values = {800002:[821536],
                  816865: [821537, 821538, 821539],
                  820109: [821540, 821541, 821542],
                  821535:[821543, 821544]}
        self.assertEqual(test_values, values)

    def test_parse_spreadsheet(self):
        self.dupe.parse_spreadsheet()


if __name__ == '__main__':
    unittest.main()
