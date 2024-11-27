import math
import formula
import validation
import pandas as pd

COLUMNS = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T",
           "U", "V", "W", "X", "Y", "Z"]

LEN_ROWS = 40
LEN_COLS = 26


def make_data_frame() -> pd.DataFrame:
    """makes a test dataframe, in each test we will use this dataframe to test
    the functions. So the data frame consists of various types of data, including str, int, float"""
    main_lst = []
    new_lst = []
    for i in range(LEN_ROWS):
        for j in range(LEN_COLS):
            if j == 1:
                new_lst.append("random")
            elif j == 2:
                new_lst.append("2")
            elif j == 3:
                new_lst.append("0.5")
            elif j == 4:
                new_lst.append("0")
            elif j == 5:
                new_lst.append("1")
            elif j == 20:
                new_lst.append("h")
            elif j == 25:
                new_lst.append("32.2")
            else:
                new_lst.append("_")
        main_lst.append(new_lst)
        new_lst = []
    main_lst[0][0] = "2"
    main_lst[1][0] = "3"
    main_lst[2][0] = "4"
    main_lst[3][0] = "5"
    main_lst[4][0] = "6"
    main_lst[5][0] = "7"
    main_lst[6][0] = "8"
    main_lst[7][0] = "9"
    main_lst[8][0] = "10"
    columns = COLUMNS

    # Create the pandas DataFrame
    df1 = pd.DataFrame(main_lst, columns=columns)
    new_index = [str(i + 1) for i in range(LEN_ROWS)]  # Generating new row index starting from 1
    df1 = df1.rename(index=dict(zip(df1.index, new_index)))
    return df1


def test_get_params_formula() -> None:
    """tests function get params - test that function knows to separate between coordinates,
    knows to differentiate between numbers and coordinates and etc"""
    df = make_data_frame()
    assert formula.get_params_formula("=min(A1:A8)", df) == ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8']
    assert formula.get_params_formula("=avg(a1:a8, a10)", df) == ['a1', 'a2', 'a3', 'a4', 'a5', 'a6', 'a7', 'a8', 'a10']
    assert formula.get_params_formula("=min(a1, a4, b3, b4:b5)", df) == ['a1', 'a4', 'b3', 'b4', 'b5']
    assert formula.get_params_formula("=sum(a4, a4, b3, b4:b5)", df) == ['a4', 'a4', 'b3', 'b4', 'b5']


def test_get_math_answer() -> None:
    """tests function get_math_answer - test that function knows to calculate the correct answer
    as well as differentiate between different formulas and etc."""
    df = make_data_frame()
    assert formula.get_math_answer("=min(A1:A8)", df) == 2 / 1
    assert formula.get_math_answer("=avg(a1:a8, a10)", df) == 44 / 9
    assert formula.get_math_answer("=min(a1, a4, D1)", df) == 0.5
    assert formula.get_math_answer("=sum(a4, a1, D1, e1)", df) == 7.5
    assert formula.get_math_answer("=max(a1, a4, D1, e1)", df) == 5
    assert formula.get_math_answer("=sqrt(a1)", df) == math.sqrt(2)
    assert formula.get_math_answer("=max(A1:A8)", df) == 9
    assert formula.get_math_answer("=avg(A1:A8)", df) == 44 / 8
    # Assertion tests for valid_formula function with valid formulas on the provided DataFrame
    assert formula.get_math_answer("=A1+A2-A3/A4", df) == 4.2
    assert formula.get_math_answer("=e2/C3-D4", df) == -0.5
    assert formula.get_math_answer("=C1*C3/A4+a5", df) == 6.8
    assert formula.get_math_answer("=89", df) == "89"
    assert formula.get_math_answer("=min(5,3,4)", df) == 3
    assert formula.get_math_answer("=avg(5,3,4, Z4, A1:A8)", df) == 88.2 / 12
    assert formula.get_math_answer("=avg(5,3,4, y4, A1:A8)", df) == 56 / 12


def test_valid_formula() -> None:
    """tests function valid_formula - test that function knows to recognize valid formulas
    checks that error gotten is appropriate to the actual error"""
    # Assertion tests for valid_formula function
    df = make_data_frame()
    assert validation.valid_formula("=min(A1:B1)", df) == (
    False, 'Invalid value in cell'), "Test failed: Minimum number of parameters not met"
    assert validation.valid_formula("=max(A1:A8)", df) == (
    True, ""), "Test failed: Maximum number of parameters not met"
    assert validation.valid_formula("=sqrt(A1,B2)", df) == (
    False, "Invalid parameter for sqrt"), "Test failed: Incorrect number of parameters for sqrt"
    assert validation.valid_formula("=sum(A1:B2)", df) == (
    False, "Invalid range"), "Test failed: sum function not recognized"
    assert validation.valid_formula("=avg(A1:A8)", df) == (True, ""), "Test failed: avg function not recognized"
    assert validation.valid_formula("=sum(A1,B2)", df) == (
    False, "Invalid value in cell"), "Test failed: Invalid column letter"
    assert validation.valid_formula("=min(A1:B)", df) == (False, "Invalid range"), "Test failed: Invalid row number"
    assert validation.valid_formula("=max(A1:Z9)", df) == (
    False, "Invalid range"), "Test failed: Row number exceeds dataframe size"
    # Assertion tests for valid_formula function with complex formulas on the provided DataFrame
    assert validation.valid_formula("=A1+A2-C3+AB/A5", df) == (
    False, "Invalid row number"), "Test failed: row not right"
    assert validation.valid_formula("=A1-A2*B3/C4+D5", df) == (
    False, "Invalid value in cell"), "Test failed: Non-integer value in formula"
    assert validation.valid_formula("=A1+A2-A3/A4+B5*B6", df) == (
    False, "Invalid value in cell"), "Test failed: Non-integer value in formula"
    assert validation.valid_formula("=A1+A2/A3-C4*C5/D6", df) == (
    False, "Can't divide by zero"), "Test failed: Division by zero"
    assert validation.valid_formula("=A1+E2-C3*D4/A5+A6-C7/E8+E9*F10", df) == (
    False, "Can't divide by zero"), "Test failed: Division by zero"
    # Assertion tests for valid_formula function with valid formulas on the provided DataFrame
    assert validation.valid_formula("=A1+A2-A3/A4", df) == (
    True, ""), "Test failed: Valid formula recognized as invalid"
    assert validation.valid_formula("=e2/C3-D4", df) == (True, ""), "Test failed: Valid formula recognized as invalid"
    assert validation.valid_formula("=C1-C2*C3/A4+a5", df) == (
    True, ""), "Test failed: Valid formula recognized as invalid"
    assert validation.valid_formula("=D1+D2*D3/D4-D5", df) == (
    False, "Can't divide by zero"), "Test failed: Valid formula recognized as invalid"
    assert validation.valid_formula("=avg()", df) == (
    False, "Not enough parameters"), "Test failed: Invalid formula recognized as valid"
    assert validation.valid_formula("=min(++++)", df) == (
    False, 'Invalid column letter'), "Test failed: Invalid formula recognized as valid"
    assert validation.valid_formula("=max(5+3+4)", df) == (
    False, 'Invalid column letter'), "Test failed: Invalid formula recognized as valid"
    assert validation.valid_formula("=89", df) == (True, ""), "Test failed: Valid formula recognized as invalid"
    assert validation.valid_formula("=min(5,3,4)", df) == (True, ""), "Test failed: Valid formula recognized as invalid"
    assert validation.valid_formula("=avg(5,3,4, A4, A1:A8)", df) == (
    True, ""), "Test failed: Valid formula recognized as invalid"
    assert validation.valid_formula("=sqrt(5,3,4)", df) == (
    False, "Invalid parameter for sqrt"), "Test failed: Invalid formula recognized as valid"
    assert validation.valid_formula("=sqrt(4)", df) == (True, ''), "Test failed: Valid formula recognized as invalid"

    assert validation.valid_formula("=min(A1:B1:B2)", df) == (
    False, 'Invalid value in cell'), "Test failed: Range with multiple colons recognized as invalid"
    assert validation.valid_formula("=max(A1:A0)", df) == (
    False, "Invalid parameters"), "Test failed: Invalid row number (zero)"
    assert validation.valid_formula("=sqrt(A1:B2)", df) == (
    False, "Invalid parameter for sqrt"), "Test failed: Invalid parameter for sqrt"
    assert validation.valid_formula("=sum(A1:B2:C3)", df) == (
    False, "Invalid range"), "Test failed: Range with multiple columns recognized as invalid"
    assert validation.valid_formula("=avg(A1:A8, B1:B8)", df) == (
    False, 'Invalid value in cell'), "Test failed: Cant add up strings"
    assert validation.valid_formula("=sum(A1:B2)+D4", df) == (
    False, "Invalid range"), "Test failed: Formula with valid range and additional cell recognized as invalid"
    assert validation.valid_formula("=avg(5,3,4, A4, A1:A8)", df) == (
    True, ""), "Test failed: Multiple ranges and individual cells recognized as invalid"
    assert validation.valid_formula("=sqrt(A1+A2-A3/A4+B5*B6)", df) == (
    False, "Invalid parameter for sqrt"), "Test failed: Complex formula recognized as invalid"
    assert validation.valid_formula("=sqrt(4) + 5", df) == (
    False, 'Invalid parameter for sqrt'), "Test failed: Formula with constants recognized as invalid"
