import re
from typing import List, Union, Any
import pandas as pd
import validation
import math


VALID_LETTER = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


def get_params_formula(formula: str, df: pd.DataFrame) -> List[str]:
    """Gets all cells used in formula - their coordination"""
    num_rows, num_cols = df.shape
    if "+" in formula or "-" in formula or "/" in formula or "*" in formula:
        pattern = r'[+\-*/=]'
        # Use re.split() to split the string based on the pattern
        parts = re.split(pattern, formula)
        # Remove empty strings from the list
        lst_value = [part for part in parts if part]
        return are_params(lst_value)
    if "min" in formula or "MIN" in formula:
        lst_value = validation.split_by_formula(formula, "min")
        return are_params(lst_value)
    if "max" in formula or "MAX" in formula:
        lst_value = validation.split_by_formula(formula, "max")
        return are_params(lst_value)
    if "sqrt" in formula or "SQRT" in formula:
        lst_value = validation.split_by_formula(formula, "sqrt")
        return are_params(lst_value)
    if "sum" in formula or "SUM" in formula:
        lst_value = validation.split_by_formula(formula, "sum")
        return are_params(lst_value)
    if "avg" in formula or "AVG" in formula:
        lst_value = validation.split_by_formula(formula, "avg")
        return are_params(lst_value)
    # maybe gotten one integer
    else:
        value = formula[1:]
        if value[0].upper() in VALID_LETTER:
            if value[1:].isdigit():
                if int(value[1:]) > num_rows:
                    return []
                return [value]
            else:
                return []
        else:
            return []


def get_math_answer(formula: str, df: pd.DataFrame) -> Union[float, int, str]:
    """Returns the answer of the formula given"""
    try:
        num_rows, num_cols = df.shape
        if "+" in formula or "-" in formula or "/" in formula or "*" in formula:
            return get_complex_formula(formula, df)
        if "min" in formula or "MIN" in formula:
            lst_value = validation.split_by_formula(formula, "min")
            return min_formula(lst_value, df)
        if "max" in formula or "MAX" in formula:
            lst_value = validation.split_by_formula(formula, "max")
            return max_formula(lst_value, df)
        if "sqrt" in formula or "SQRT" in formula:
            lst_value = validation.split_by_formula(formula, "sqrt")
            return sqrt_formula(lst_value, df)
        if "sum" in formula or "SUM" in formula:
            lst_value = validation.split_by_formula(formula, "sum")
            return sum_formula(lst_value, df)
        if "avg" in formula or "AVG" in formula:
            lst_value = validation.split_by_formula(formula, "avg")
            return avg_formula(lst_value, df)
        else:
            value = formula[1:]
            if value[0].upper() in VALID_LETTER:
                if value[1:].isdigit():
                    if int(value[1:]) > num_rows:
                        return value
                    else:
                        return get_value_param([str(value[1:]), value[0]], df)
                else:
                    return value
            else:
                return value
    except:
        return "ERROR"


def sum_formula(lst_value: List, df: pd.DataFrame) -> float:
    """Returns sum of all cell values"""
    sum_formula = 0.0
    for param in lst_value:
        if check_float(param):
            value = float(param)
        else:
            value = get_value_param([str(param[1:]), param[0]], df)
        # first time in we want to make sure to adjust min_value
        sum_formula += value
    return sum_formula


def check_float(num: str) -> bool:
    """Checks if a number is a float"""
    try:
        if num == "":
            return False
        if num[0] == "-":
            num = num[1:]
        if "." in num:
            divide_lst = num.split(".")
            if divide_lst[0].isdigit() and divide_lst[1].isdigit():
                return True
        if num.isdigit():
            return True
        return False
    except ValueError:
        return False


def avg_formula(lst_value: List, df: pd.DataFrame) -> Union[float, int]:
    """Returns average of all cell values"""
    sum_value = sum_formula(lst_value, df)
    return sum_value/len(lst_value)


def seperate_formula(lst_formula: List, formula: str, df: pd.DataFrame) -> str:
    """Turns formula from having params to being a math formula with numbers and signs"""
    for item in lst_formula:
        if check_float(item):
            pass
        else:
            value = get_value_param([str(item[1:]), item[0]], df)
            formula = formula.replace(item, str(value))
    # changes formula from string to int in order to do math
    formula = formula[1:]
    is_valid, reason_not_valid = validation.check_valid_expression(formula)
    if not is_valid:
        return "DIVISION ZERO"
    result = eval(formula)
    return result


def get_complex_formula(formula: str, df: pd.DataFrame) -> Union[float, int, str]:
    """Returns a list of all the params in formula gotten"""
    pattern = r'[+\-*/=]'
    # Use re.split() to split the string based on the pattern
    parts = re.split(pattern, formula)
    # Remove empty strings from the list
    parts = [part for part in parts if part]
    seperated_formula = seperate_formula(parts, formula, df)
    return seperated_formula


def min_formula(lst_value: List, df: pd.DataFrame) -> float:
    """Returns minimum of all cell values"""
    min_value = 0.0
    i = 0
    for param in lst_value:
        if check_float(param):
            value = float(param)
        else:
            value = get_value_param([str(param[1:]), param[0]], df)
            if value == 0:
                param = [str(param[1:]), param[0]]
                param[1] = param[1].upper()
                check_value = df.loc[tuple(param)]
                if check_value == "_":
                    value = min_value
        # first time in we want to make sure to adjust min_value
        if i == 0:
            min_value = value
        else:
            if value < min_value:
                min_value = value
        i += 1
    return min_value


def max_formula(lst_value: List, df: pd.DataFrame) -> Union[float, int]:
    """Returns maximum of all cell values"""
    max_value = 0.0
    i = 0
    for param in lst_value:
        if check_float(param):
            value = float(param)
        else:
            value = get_value_param([str(param[1:]), param[0]], df)
            if value == 0:
                param = [str(param[1:]), param[0]]
                param[1] = param[1].upper()
                check_value = df.loc[tuple(param)]
                if check_value == "_":
                    value = max_value
        # first time in we want to make sure to adjust max_value
        if i == 0:
            max_value = value
        else:
            if value > max_value:
                max_value = value
        i += 1
    return max_value


def sqrt_formula(lst_value: List, df: pd.DataFrame) -> Union[float, int]:
    """Returns sqrt of cell value"""
    if check_float(lst_value[0]):
        return math.sqrt(float(lst_value[0]))
    else:
        value = get_value_param([str(lst_value[0][1:]), lst_value[0][0]], df)
    return math.sqrt(value)


def are_params(lst_value: List) -> List[str]:
    """Returns which variables in list are params"""
    new_lst = []
    for item in lst_value:
        if not check_float(item):
            new_lst.append(item)
    return new_lst


def get_value_param(param: List, df: pd.DataFrame) -> float:
    """Gets the value of a specific cell"""
    param_exc = param
    try:
        param[1] = param[1].upper()
        value = df.loc[tuple(param)]
        if value == "_":
            return 0
        else:
            return float(value)
    except KeyError:
        return 0


def is_formula(new_value: str) -> bool:
    """Checks if the new value is a formula"""
    if new_value == "":
        return False
    if new_value[0] == "=":
        return True
    return False



