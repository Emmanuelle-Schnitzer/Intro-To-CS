import re
from typing import List, Union
import formula as Formula

VALID_LETTER = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

def check_range(lst_value)-> bool:
    """Returns if a given range is valid"""
    if lst_value == ["Invalid range"]:
        return True
    return False

def check_cell_in_formula(params, new_value) -> bool:
    params = params[1].upper() + params[0]
    if params in new_value.upper():
        return False
    return True


def valid_formula(formula, df) -> tuple:
    """Returns if a gotten formula is valid, and if not, the reason why"""
    if "min" in formula or "MIN" in formula:
        lst_value = split_by_formula(formula, "min")
        if check_range(lst_value):
            return False, "Invalid range"
        if not lst_value:
            return False, "Invalid parameters"
        is_valid, reason_not_valid = check_len_formula(lst_value, "min", df)
        if not is_valid:
            return False, reason_not_valid
        return True, ""
    if "max" in formula or "MAX" in formula:
        lst_value = split_by_formula(formula, "max")
        if check_range(lst_value):
            return False, "Invalid range"
        if not lst_value:
            return False, "Invalid parameters"
        is_valid, reason_not_valid = check_len_formula(lst_value, "max", df)
        if not is_valid:
            return False, reason_not_valid
        return True, ""
    if "sqrt" in formula or "SQRT" in formula:
        lst_value = split_by_formula(formula, "sqrt")
        if check_range(lst_value):
            return False, "Too many parameters for sqrt"
        if not lst_value:
            return False, "Invalid parameters"
        is_valid, reason_not_valid = check_len_formula(lst_value, "sqrt", df)
        if not is_valid:
            return False, reason_not_valid
        return True, ""
    if "sum" in formula or "SUM" in formula:
        lst_value = split_by_formula(formula, "sum")
        if check_range(lst_value):
            return False, "Invalid range"
        if not lst_value:
            return False, "Invalid parameters"
        is_valid, reason_not_valid = check_len_formula(lst_value, "sum", df)
        if not is_valid:
            return False, reason_not_valid
        return True, ""
    if "avg" in formula or "AVG" in formula:
        lst_value = split_by_formula(formula, "avg")
        if check_range(lst_value):
            return False, "Invalid range"
        if not lst_value:
            return False, "Invalid parameters"
        is_valid, reason_not_valid = check_len_formula(lst_value, "avg", df)
        if not is_valid:
            return False, reason_not_valid
        return True, ""
    if "+" in formula or "-" in formula or "/" in formula or "*" in formula:
        return complex_expression(formula, df)
    return True, ""

def get_lst_value(formula, df) -> tuple:
    """Returns if each param gotten is valid"""
    pattern = r'[+\-*/=]'
    # Use re.split() to split the string based on the pattern
    parts = re.split(pattern, formula)
    # Remove empty strings from the list
    parts = [part for part in parts if part]
    integer_valid, reason_not_valid = check_len_formula(parts, "", df)
    if not integer_valid:
        return False, reason_not_valid
    is_valid, reason_not_valid = check_int(parts, formula, df)
    if is_valid:
        return True, ""
    return False, reason_not_valid

def get_value_param(param, df) -> str:
    """Returns value of param in df, if not in df, returns param"""
    param_exc = param
    try:
        param = str(param[1:]), param[0].upper()
        value = df.loc[param]
        return value
    except KeyError:
        return param_exc

def check_int(lst_formula, formula, df) -> tuple:
    """Makes a string of the formula"""
    try:
        for item in lst_formula:
            if Formula.check_float(item):
                pass
            else:
                value = get_value_param(item, df)
                formula = formula.replace(item, str(value))
        # changes formula from string to int in order to do math
        formula = formula[1:]
        return check_valid_expression(formula)
    except Exception as e:
        print(e)
        return False, "An error occurred during formula validation"


def check_valid_expression(formula) -> tuple:
    """Checks if a formula (now not having params) is valid, if not, returns why"""
    try:
        # Check for division by zero
        if "/0" in formula:
            return False, "Can't divide by zero"
        # Check if the formula is not empty
        if not formula.strip():
            return False, "Formula is empty"

        # If all checks pass, the formula is valid
        return True, ""
    except:
        return False, "An error occurred during formula validation"


def complex_expression(formula, df) -> tuple:
    """Checks if an expression with +, -, *, /, is valid"""
    try:
        return get_lst_value(formula, df)
    except:
        return False, "An error occurred during formula validation"


def check_len_formula(formula_lst, operand, df) -> tuple:
    """Checks if gotten correct number of cells and that cells are in df"""
    num_rows, num_cols = df.shape
    # needs atleast two numbers in order to do math equation
    if len(formula_lst) < 1 or formula_lst == [""]:
        return False, "Not enough parameters"
    for param in formula_lst:
        # invalid column letter
        if Formula.check_float(param):
            continue
        if not param[0].upper() in VALID_LETTER:
            if operand.upper() == "SQRT":
                return False, "Invalid parameter for sqrt"
            return False, "Invalid column letter"
        if not param[1:].isdigit():
            if operand.upper() == "SQRT":
                return False, "Invalid parameter for sqrt"
            return False, "Invalid row number"
        # is digit so can convert to int
        if int(param[1:]) == 0:
            return False, "Invalid row number"
        if int(param[1:]) > num_rows:
            return False, "Row number too big"
        # means params are valid, lets see if values in df of params are valid
        if not check_is_int([str(param[1:]), param[0]], df):
            return False, "Invalid value in cell"
    if operand == "sqrt" or operand == "SQRT":
        if len(formula_lst) != 1:
            return False, "Too many parameters for sqrt"
        if Formula.check_float(formula_lst[0]):
            return True, ""
        param = formula_lst[0]
        if not can_sqrt([str(param[1:]), param[0]], df):
            return False, "Can't sqrt negative number"
    return True, ""


def can_sqrt(params, df) -> bool:
    """Returns true if value of param can be sqrt"""
    try:
        params[1] = params[1].upper()
        value = df.loc[tuple(params)]
        if value == "_":
            return True
        if float(value) < 0:
            return False
        return True
    except:
        return False


def check_is_int(params, df) -> bool:
    """Checks if gotten cells are integers or if "_"""
    try:
        params[1] = params[1].upper()
        value = df.loc[tuple(params)]
        if len(value) > 0:
            if value[0] == "-":
                value = value[1:]
        if "." in value:
            divide_lst = value.split(".")
            if divide_lst[0].isdigit() and divide_lst[1].isdigit():
                return True
            return False
        if value == "_":
            return True
        if value.isdigit():
            return True
        if len(value) > 1:
            if value[0] == "-" and value[1:].isdigit():
                return True
        return False
    except:
        return False


def can_divide_zero(params, df) -> bool:
    """Returns false if value of param is zero"""
    value = df.iloc[params]
    if value == 0 or value == "_":
        return False
    return True


def split_by_formula(formula, operand) -> List:
    """Returns a list splits by min, max, sqrt formula"""
    if operand != "sqrt":
        # erase =min/max and (
        new_lst = formula[5:]
    else:
        # erase =sqrt and (
        return [formula[6:-1]]
    new_lst = new_lst.split(",")
    # the last one is ) we dont want that
    new_lst[-1] = new_lst[-1][:-1]
    return_lst = []
    for param in new_lst:
        if ':' in param:
            check = [part.strip() for part in param.split(":") if part.strip()]
            if check[0][0] != check[1][0]:
                if check[0][1:]!=check[1][1:]:
                    return ["Invalid range"]
            range_param = get_range(param)
            for r in range_param:
                return_lst.append(r)
        else:
            param = param.strip()
            return_lst.append(param)
    return return_lst


def get_range(param) -> List:
    """Returns all cells in a given range"""
    try:
        new_lst = []
        range_lst = [part.strip() for part in param.split(":") if part.strip()]
        if range_lst[0][0] == range_lst[1][0]:
            letter = range_lst[0][0]
            start = int(range_lst[0][1:])
            end = int(range_lst[1][1:])
            for i in range(start, end +1):
                new_lst.append(letter + str(i))
            return new_lst
        else:
            num = range_lst[0][1:]
            first_let = range_lst[0][0]
            second_let = range_lst[1][0]
            num_first = get_number(first_let)
            num_second = get_number(second_let)
            if num_first>=num_second:
                start=num_second
                end = num_first
            else:
                start = num_first
                end = num_second
            for i in range(start, end +1):
                new_lst.append(get_letter(i) + num)
            return new_lst


    except:
        return []

def get_letter(num: int) -> Union[str, None]:
    """Gets the letter of the number gotten (by what it represents in the data frame)"""
    try:
        number_letter_map = {1: 'A',2: 'B',3: 'C',4: 'D',5: 'E',6: 'F',7: 'G',8: 'H',9: 'I',10: 'J',11: 'K',12: 'L',13: 'M',14: 'N',15: 'O',16: 'P',17: 'Q',18: 'R',19: 'S',20: 'T',21: 'U',22: 'V',23: 'W',24: 'X',25: 'Y',26: 'Z'}
        return number_letter_map[num]
    except:
        return None

def get_number(letter) -> Union[int, None]:
    """Gets the number of the letter gotten by what it represents in the data frame)"""
    try:
        letter = letter.upper()
        letter_number_map = {'A': 1,'B': 2,'C': 3,'D': 4,'E': 5,'F': 6,'G': 7,'H': 8,'I': 9,'J': 10,'K': 11,'L': 12,'M': 13,'N': 14,'O': 15,'P': 16,'Q': 17,'R': 18,'S': 19,'T': 20,'U': 21,'V': 22,'W': 23,'X':24,'Y': 25,'Z': 26}
        # Check if the number is in the dictionary, otherwise return None
        return letter_number_map[letter]
    except:
        return None




def check_valid_place(row,col) -> bool:
    """Function for excel_gui, checks if row and column gotten are valid"""
    if not row.isdigit() or row == "":
        return False
    if not (col.upper() in "ABCDEFGHIJKLMNOPQRSTUVWXYZ"):
        return False
    return True
