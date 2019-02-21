import openpyxl
import tkinter


class BirdDuplicator:

    def __init__(self, *args, **kwargs):
        root = tkinter()
        self.input_folder = file_dialog.GetDirectory()
        self.output_folder = self.input_folder + '\\new_versions\\'
        self.file = openpyxl.load_workbook(file_dialog.GetPath)
        self.item_ids = self._create_data_dictionary()
        return super().__init__(*args, **kwargs)

    def parse_spreadsheet(self):
        # takes each of the rows of the spreadsheet and does some logic to inform which helper
        # methods need to be run to handle the row
        print('stuff')
        return 0

    def _create_data_dictionary(self):
        # Takes the Item Spreadsheet and generates a dictionary in the form
        # {item_id: [ids for the component records]}

        return 0

    def _duplicate_record(self, row):
        # Main method for duplication of records
        return 0

    def duplicate_birds(self):
        # Main method for duplication of all records
        return 0