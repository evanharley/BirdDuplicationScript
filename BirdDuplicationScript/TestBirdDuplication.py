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
        values = {800002:[821356],
                  816865: [821357, 821358, 821359],
                  820109: [821360, 821361, 821362],
                  821535:[821363, 821364]}
        self.assertEqual(test_values, values)

if __name__ == '__main__':
    unittest.main()
