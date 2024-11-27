import json
import os
import pandas as pd
from typing import List, Dict
import formula as Formula
import validation as Validation
from pandas import Index

COLUMNS: List[str] = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T",
           "U", "V", "W", "X", "Y", "Z"]


class Sheet:
    def __init__(self) -> None:
        """
        A constructor for a Sheet in Excel
        """
        # starting state of sheet in Excel
        self.__len_rows = 40
        self.__len_cols = 26
        self.__df: pd.DataFrame = self.make_sheet()
        self.__part_of_formula: Dict[str, List[str]] = {}
        self.__are_formula: Dict[str, str] = {}

    def starting_lst(self) -> List[List[str]]:
        """Restarts a list of lists for the sheet, putting in '_' in each cell"""
        # initialize list of lists
        main_lst = []
        new_lst = []
        for i in range(self.__len_rows):
            for j in range(self.__len_cols):
                new_lst.append("_")
            main_lst.append(new_lst)
            new_lst = []
        return main_lst

    def make_sheet(self) -> pd.DataFrame:
        """Makes a new sheet, uses starting_lst function to restart sheet, puts data in df"""
        columns = COLUMNS
        data_lst = self.starting_lst()
        # Create the pandas DataFrame
        df = pd.DataFrame(data_lst, columns=columns)
        new_index = [str(i + 1) for i in range(self.__len_rows)]  # Generating new row index starting from 1
        df = df.rename(index=dict(zip(df.index, new_index)))
        return df

    def upload_json(self, file_path: str) -> bool:
        """Uploads a JSON file to the DataFrame"""
        try:
            # Read data from the JSON file into a DataFrame
            with open(file_path, 'r') as file:
                data = json.load(file)
            self.__df = pd.DataFrame(data)
            # Replace NaN with underscores
            self.__df.fillna("_", inplace=True)
            # Trim or add columns to match the expected number of columns
            if len(self.__df.columns) > len(COLUMNS):
                self.__df = self.__df.iloc[:, :len(COLUMNS)]
            elif len(self.__df.columns) < len(COLUMNS):
                missing_columns = len(COLUMNS) - len(self.__df.columns)
                for i in range(missing_columns):
                    self.__df[f'Column_{i + len(self.__df.columns) + 1}'] = "_"
            # Rename columns with default names
            # assert isinstance(self.__df.columns, list)
            self.__df.columns = list(COLUMNS)
            # Rename index
            new_index = [str(i + 1) for i in range(len(self.__df))]
            self.__df.index = new_index
            # Iterate through each cell in the DataFrame
            for index, row in self.__df.iterrows():
                for col in self.__df.columns:
                    cell_value = self.__df.at[index, col]
                    if isinstance(cell_value, list):
                        # Check if the cell value is a formula
                        if Formula.is_formula(cell_value[0]):
                        # Update the cell with the result
                            self.__df.at[index, col] = cell_value[1]
                            params = index, col
                            self.update_df_formula(cell_value[0], params)

            return True
        except Exception as e:
            return False


    def upload_excel(self, file_path: str) -> bool:
        """Uploads an excel file to the sheet"""
        try:
            # Read data from the Excel file into a DataFrame
            self.__df = pd.read_excel(file_path, header=None)
            # Replace NaN with underscores
            self.__df.fillna("_", inplace=True)
            # Trim or add columns to match the expected number of columns
            if len(self.__df.columns) > len(COLUMNS):
                self.__df = self.__df.iloc[:, :len(COLUMNS)]
            elif len(self.__df.columns) < len(COLUMNS):
                missing_columns = len(COLUMNS) - len(self.__df.columns)
                for i in range(missing_columns):
                    self.__df[f'Column_{i + len(self.__df.columns) + 1}'] = "_"
            # Rename columns with default names
            #assert isinstance(self.__df.columns, list)
            self.__df.columns = list(COLUMNS)
            # Rename index
            new_index = [str(i + 1) for i in range(len(self.__df))]
            self.__df.index = new_index
            return True
        except Exception as e:
            return False


    def add_new_row(self) -> None:
        """Adds a new row to the sheet"""
        num_rows, num_cols = self.__df.shape
        new_row_data = ["_" for i in range(len(self.__df.columns))]
        # Appending the new row to the original DataFrame
        self.__df.loc[str(num_rows + 1)] = new_row_data

    def save_excel(self, folder_path: str, file_name: str) -> None:
        """Saves the sheet as an excel file"""
        try:
            file_path = os.path.join(r'', folder_path, file_name)
            # saving the excel
            self.__df.to_excel(file_path, index=False, header=False)
        except Exception as e:
            pass

    def save_to_json(self, folder_path: str, file_name: str):
        # Save DataFrame to JSON file
        file_path = os.path.join(r'', folder_path, file_name)
        df_upload = self.__df
        for cell in self.__are_formula:
            params = cell[1:], cell[0]
            value = self.__df.loc[params]
            df_upload.loc[params] = [self.__are_formula[cell], value]
        df_upload.to_json(file_path, orient='records')

    def get_value(self, params: tuple) -> str:
        """Returns value of a specific cell"""
        try:
            value = self.__df.loc[params]  # Using iloc with integer positions
            return value
        except KeyError:
            return "_"

    def validata_new_value(self, new_value: str) -> tuple:
        """Returns if value is valid (separates to two situations;
        one being a formula and one being a number or string"""
        is_math = Formula.is_formula(new_value)
        if new_value == "":
            return False, "Empty cell"
        # is formula
        if is_math:
            # if formula is valid
            is_valid, reason_not_valid = Validation.valid_formula(new_value, self.__df)
            if not is_valid:
                # if formula isn't valid tell gui to tell user invalid formula
                return False, reason_not_valid
        return True, ""

    def update_df_formula(self, current_formula: str, params: tuple) -> None:
        """Updates dictionaries of class according to formula gotten - are_formula and part_of_formula"""
        old_value = self.get_value(params)
        current_params = self.make_params(params)
        if Formula.is_formula(old_value):
            cell_lst = Formula.get_params_formula(old_value, self.__df)
            for cell in cell_lst:
                cell = cell.upper()
                self.__part_of_formula[cell].remove(current_params)
        self.__are_formula[current_params] = current_formula
        new_cells_lst = Formula.get_params_formula(current_formula, self.__df)

        for new_cell in new_cells_lst:
            new_cell = new_cell.upper()
            if new_cell in self.__part_of_formula:
                self.__part_of_formula[new_cell].append(current_params)
                for cell in self.__part_of_formula:
                    if new_cell in self.__part_of_formula[cell]:
                        self.__part_of_formula[cell].append(current_params)
            else:
                self.__part_of_formula[new_cell] = [current_params]
                for cell in self.__part_of_formula:
                    if new_cell in self.__part_of_formula[cell]:
                        self.__part_of_formula[cell].append(current_params)


    def set_new_value(self, got_value: str, params: tuple) -> None:
        """Sets a new value to a specific cell in the sheet, updates the sheet accordingly.
        If the new value is a formula, updates the dictionaries of the class accordingly, else checks if the cell is part of a
        formula and updates the sheet accordingly."""
        if Formula.is_formula(got_value):
            current_formula = got_value
            new_value = Formula.get_math_answer(got_value, self.__df)
            self.update_df_formula(current_formula, params)
            # Using loc with index labels
            self.__df.loc[params] = str(new_value)
            new_value = str(new_value)
            self.update_is_formula_cell(new_value, params)
        else:
            # Using loc with index labels
            self.__df.loc[params] = str(got_value)
            # if cell is part of a formula, update other cells
            self.update_cell_formulas(got_value, params)

    def make_params(self, params: tuple) -> str:
        """Change params from being tuple to be string"""
        return params[1].upper() + params[0]

    def str_to_param(self, str_param: str) -> tuple:
        """Shifts params from str to tuple"""
        return str_param[1:], str_param[0].upper()

    def change_single_cell(self, cell_param: str, is_digit: bool) -> None:
        """Updates single cell, if new value is a digit, updates the cell accordingly,
         else changes all cell formulas to be error"""
        # new value isn't digit so need to change all cell formulas to be error
        if is_digit:
            # get the formula of cell from dictionary
            new_formula = self.__are_formula[cell_param]
            # after we updated the chart, get new answer to formula
            math_answer = Formula.get_math_answer(new_formula, self.__df)
            # update data frame
            # Using loc with index labels
            tuple_cell_param = self.str_to_param(cell_param)
            self.__df.loc[tuple_cell_param] = str(math_answer)
        else:
            # Using loc with index labels
            tuple_cell_param = self.str_to_param(cell_param)
            self.__df.loc[tuple_cell_param] = str("ERROR")

    def get_df(self) -> pd.DataFrame:
        """Returns the data frame of the sheet"""
        return self.__df

    def update_is_formula_cell(self, new_value: str, params: tuple) -> None:
        """Updates df when changing a specific cell to a different value"""
        new_param = self.make_params(params)
        new_value = str(new_value)
        if new_value == "_":
            new_value = "0"
        if new_param in self.__part_of_formula:
            # if value isn't digit so all the formulas will be not valid
            if not Formula.check_float(new_value):
                for cell in self.__part_of_formula[new_param]:
                    self.change_single_cell(cell, False)
            else:
                for cell in self.__part_of_formula[new_param]:
                    self.change_single_cell(cell, True)

    def update_cell_formulas(self, new_value: str, params: tuple) -> None:
        """Updates df when changing a specific cell to a different value, if the cell is part of a formula"""
        new_param = self.make_params(params)
        if new_value == "_":
            new_value = "0"
        if new_param.upper() in self.__part_of_formula:
            # if value isn't digit so all the formulas will be not valid
            if not Formula.check_float(new_value):
                for cell in self.__part_of_formula[new_param]:
                    self.change_single_cell(cell, False)
            else:
                for cell in self.__part_of_formula[new_param]:
                    self.change_single_cell(cell, True)
        for key in self.__part_of_formula:
            if new_param in self.__part_of_formula[key]:
                self.__part_of_formula[key].remove(new_param)
        # deletes the old formula
        if new_param in self.__are_formula:
            del self.__are_formula[new_param]
