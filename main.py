
import excel_gui
import argparse

if __name__ == '__main__':
    """ 
    Running the entire excel project
    """
    parser = argparse.ArgumentParser(description='INSTRUCTIONS: From the terminal, run the main.py (python main.py) file to start '
                                                 'the Excel project. The GUI will open and you can start using the '
                                                 'Excel application.')
    args = parser.parse_args()
    excel_gui.ExcelGui()
