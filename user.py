import os
from sheet import Sheet
from typing import Dict


class User:
    def __init__(self, name: str):
        """A constructor for a User in Excel"""
        self.__name = name
        self.__sheets: Dict[str, Sheet] = {}
        self.starting_sheet()

    def get_name(self) -> str:
        """Returns the name of the user"""
        return self.__name

    def starting_sheet(self) -> None:
        """Creates a starting sheet (default sheet) for the user"""
        self.__sheets["1"] = Sheet()

    def add_sheet(self, sheet_name: str) -> None:
        """Adds a new sheet to the user's sheets list"""
        self.__sheets[sheet_name] = Sheet()

    def upload_sheet(self, file_path: str) -> bool:
        """Uploads a sheet (from file gotten) to the user's sheets list"""
        new_sheet = Sheet()
        if ".xlsx" in file_path:
            new_sheet.upload_excel(file_path)
        elif ".json" in file_path:
            new_sheet.upload_json(file_path)
        else:
            return False
        file_name = os.path.basename(file_path)
        self.__sheets[file_name] = new_sheet
        return True

    def get_sheets(self) -> Dict[str, Sheet]:
        """Returns the user's sheets list"""
        return self.__sheets

    def save_all_data(self, folder_name: str) -> None:
        for num_sheet in self.__sheets:
            file_name = self.__name + "sheet" + num_sheet + ".xlsx"
            self.__sheets[num_sheet].save_excel(folder_name, file_name)
            file_name = self.__name + "sheet" + num_sheet + ".json"
            self.__sheets[num_sheet].save_to_json(folder_name, file_name)



