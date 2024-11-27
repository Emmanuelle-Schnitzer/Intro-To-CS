import os
import queue
import sys
import tkinter as tk
from tkinter import ttk
import user
from tkinter.ttk import Treeview
from typing import List, Union

from PIL import Image as PILImage
# type: ignore
import customtkinter
import validation as Validation
from app import App

class ExcelGui:
    SCREEN_SIZE = '1000x600'
    LETS_PLAY = 'WELCOME TO EXCEL!'
    START_SCREEN = 'Twitter_Photo.png'
    UNDO_IMAGE = "undo_img.png"
    REDO_IMAGE = "redo_img.png"
    FONT = "Arial"

    def __init__(self) -> None:
        """
        This method initializes all necessary variables to run the Excel app, An App object is started.
        Uses helper functions.
        """

        self.un_re = None
        self.game = App()
        # holds current user using the app
        self.__current_user: Union[user.User, None] = None
        self.__root = self.set_tk_root()
        # holds data of df
        self.__treeview = self.create_treeview
        # holds information of which row the cell chosen is in
        self.row_entry = None
        # holds information of which column the cell chosen is in
        self.col_entry = None
        # holds information of which data the user wants to put in cell
        self.input_entry = None
        self.style = None
        self.choose_user: Union[ttk.Button, None] = None
        # name of current sheet user is using
        self.__current_sheet = "1"
        # holds input of user of number of sheet he wants to be in
        self.sheet_num = None
        # when clicked saves user sheets
        self.save_folder = None
        # when clicked uploads file
        self.file_upload = None
        # holds the redo

        self.__redo: queue.Queue = queue.Queue()
        # holds the undo
        self.__undo: List = []
        # starts the app
        self.__create_start_screen_label()
        self.widgets_frame = None
        self.frame = None
        self.mode_switch = None
        self.switch_var = "on"
        self.__root.mainloop()


    def set_tk_root(self) -> tk.Tk:
        """This method sets up TKinter window root"""
        root = tk.Tk()
        root.resizable(False, False)
        root.geometry(self.SCREEN_SIZE)
        root.wm_title(self.LETS_PLAY)
        return root

    def create_treeview(self) -> ttk.Treeview:
        """Creates a Treeview widget and returns it"""
        tree_frame = ttk.Frame(self.__root)
        tree_frame.grid(row=0, column=0)
        tree = ttk.Treeview(tree_frame, columns=("Name", "Age"))
        tree.grid(row=0, column=0)
        return tree

    def __continue_with_name(self):
        """Starts the main screen with specific name entered by user"""
        # Get the text entered in the Entry widget
        name = self.__start_entry.get()
        # Call the continue function with the entered name
        self.start_game(name)

    def __create_start_screen_label(self):
        """
        This method creates the Excel Start screen (background image, labels
        and buttons).
        """
        self.__background_img = tk.PhotoImage(file=self.START_SCREEN)
        self.__label = tk.Label(self.__root, image=self.__background_img)
        self.__label.place(x=0, y=0)
        # Create a text label
        self.__text_label = tk.Label(self.__root, text="Place your username:",
                                     font=(self.FONT, 16), fg='#00008B')
        self.__text_label.place(x=240, y=280)  # Positioned just above the widget
        # Create an Entry widget for user input
        self.__start_entry = tk.Entry(self.__root, font=(self.FONT, 16), bg='white', fg='#00008B', width=13)
        self.__start_entry.place(x=240, y=320)  # Positioned the widget
        # Create a Button widget for continuing
        self.__continue_button = tk.Button(self.__root, text="Continue",
                                           command=self.__continue_with_name,
                                           font=(self.FONT, 16), bg='#00008B', fg='white')
        self.__continue_button.place(x=430, y=320)  # Positioned next to the widget


    def start_game(self, name: str):
        """
        This method is the event that occurs when continue button is
         pressed. The music starts playing, labels and game variables are set,
         an the start screen gets replaced by the App screen.
        """
        self.__text_label.destroy()
        self.__start_entry.destroy()
        self.__continue_button.destroy()
        if not self.game.check_user_in(name):
            current_user = self.game.add_new_user(name)
        else:
            current_user = self.game.get_user(name)
        self.__current_user = current_user
        self.__current_sheet = "1"
        self.create_main_board()

    def change_current_user(self, name: str):
        """Changes current user according to name gotten.
        If the user doesn't exist creates a new user (using App class). Resets labels and entry widgets."""
        if not self.game.check_user_in(name):
            current_user = self.game.add_new_user(name)
        else:
            current_user = self.game.get_user(name)
        assert self.__current_user is not None, "self.__current_user should not be None at this point"
        self.__current_user = current_user
        self.__current_sheet = "1"
        assert self.input_entry is not None, "self.input_entry should not be None at this point"
        self.input_entry.delete(0, tk.END)
        self.input_entry.insert(0, "INPUT")
        # Disable treeview
        self.__treeview.bind("<ButtonRelease-1>", lambda event: None)

        self.row_entry.config(state="normal")
        self.col_entry.config(state="normal")
        self.row_entry.delete("1.0", "end")
        self.row_entry.insert("1.0", "ROW")
        self.col_entry.delete("1.0", "end")
        self.col_entry.insert("1.0", "COL")
        self.row_entry.config(state="disabled")
        self.col_entry.config(state="disabled")
        self.sheet_num.delete(0, tk.END)
        self.sheet_num.insert(0, "SHEET")
        self.save_folder.delete(0, tk.END)
        self.save_folder.insert(0, "SAVE FOLDER")
        # Bind the click event to the treeview
        self.__treeview.bind("<ButtonRelease-1>", self.on_item_click)
        self.refresh_screen()

    def change_user(self):
        """When clicked on change user button changes current user using the App.
        Calls change_current_user function with the names received from widget"""
        # delete queue
        self.__undo = []
        self.__redo = queue.Queue()
        # Create a new window
        new_window = tk.Toplevel(self.__root)
        new_window.title("New User Name")

        # Add labels and entry widgets for user input
        label1 = ttk.Label(new_window, text="Enter name:")
        label1.grid(row=0, column=0, padx=10, pady=5)

        input_entry = ttk.Entry(new_window)
        input_entry.grid(row=0, column=1, padx=10, pady=5)

        # Function to get the input value
        def get_input():
            new_value = input_entry.get()
            # You can now use the 'new_value' obtained from the input
            # Close the new window
            new_window.destroy()
            self.change_current_user(new_value)

        # Add a button to submit the input
        submit_button = ttk.Button(new_window, text="Submit", command=get_input)
        submit_button.grid(row=1, column=0, columnspan=2, padx=10, pady=5)

    def make_row_entry(self):
        """Row of cell user clicked on"""
        self.row_entry = tk.Text(self.widgets_frame, height=1.5, width=10, font=(self.FONT, 12, "bold"))
        self.row_entry.insert("1.0", "COL")
        self.row_entry.configure(state="disabled")
        self.row_entry.place(relx=0.0, rely=0.1, anchor="nw")

    def make_col_entry(self):
        """column clicked on"""
        self.col_entry = tk.Text(self.widgets_frame, height=1.5, width=10, font=(self.FONT, 12, "bold"))
        self.col_entry.insert("1.0", "ROW")
        self.col_entry.configure(state="disabled")
        self.col_entry.place(relx=0.1, rely=0.1, anchor="nw")

    def make_input_entry(self):
        """Input of user what value wants in specific cell entry"""
        self.input_entry = ttk.Entry(self.widgets_frame, width=20, font=(self.FONT, 15, "bold"))
        self.input_entry.insert(0, "INPUT")
        self.input_entry.bind("<FocusIn>", lambda e: self.input_entry.delete('0', 'end'))
        self.input_entry.place(relx=0.0, rely=0.3, anchor="nw")

    def make_button_change(self):
        """change cell value button"""
        button_change = tk.Button(self.widgets_frame, text="CHANGE", command=self.apply_button, width=10,
                                  font=(self.FONT, 14, "bold"), bg='lightblue')
        button_change.place(relx= 0.0, rely=0.5, anchor="nw")
    def make_button_delete_cell(self):
        """Delete cell button"""
        button_delete_cell = tk.Button(self.widgets_frame, text="DELETE CELL", command=self.delete_cell, width=12,
                                       font=(self.FONT, 12, "bold"), bg='lightblue')
        button_delete_cell.place(relx=0.26, rely=0.3, anchor="nw")

    def make_file_upload(self):
        """file to upload entry"""
        self.file_upload = ttk.Entry(self.widgets_frame, width=20)
        self.file_upload.insert(0, "#FILE TO UPLOAD")
        self.file_upload.bind("<FocusIn>", lambda e: self.file_upload.delete('0', 'end'))
        self.file_upload.place(relx=0.5, rely=0.11, anchor="nw")
    def make_button_upload_file(self):
        """Upload file button"""
        button_upload_file = tk.Button(self.widgets_frame, text="UPLOAD", command=self.upload_df, width=10)
        button_upload_file.place(relx=0.5, rely=0.25, anchor="nw")
    def make_button_add_row(self):
        """Add row button"""
        button_add_row = tk.Button(self.widgets_frame, text="ADD ROW", command=self.add_row, width=12,
                                   font=(self.FONT, 12, "bold"), bg='lightblue')
        button_add_row.place(relx=0.26, rely=0.5, anchor="nw")

    def make_sheet_num(self):
        """Sheet number entry"""
        self.sheet_num = ttk.Entry(self.widgets_frame, width=30)
        self.sheet_num.insert(0, "#PUT IN SHEET NUMBER")
        self.sheet_num.bind("<FocusIn>", lambda e: self.sheet_num.delete('0', 'end'))
        self.sheet_num.place(relx=0.75, rely=0.05, anchor="nw")
    def make_button_change_sheet(self):
        """Change sheet button"""
        button_change_sheet = ttk.Button(self.widgets_frame, text="CHANGE SHEET", command=self.change_sheet, width=30)
        button_change_sheet.place(relx=0.75, rely=0.2, anchor="nw")

    def make_save_folder(self):
        """Save folder entry"""
        self.save_folder = ttk.Entry(self.widgets_frame, width=30)
        self.save_folder.insert(0, "#NAME OF FOLDER TO SAVE IN")
        self.save_folder.bind("<FocusIn>", lambda e: self.save_folder.delete('0', 'end'))
        self.save_folder.place(relx=0.75, rely=0.35, anchor="nw")

    def make_save_button(self) -> None:
        """Save button"""
        button_save = ttk.Button(self.widgets_frame, text="SAVE", command=self.save_data, width=30)
        button_save.place(relx=+ 0.75, rely=0.5, anchor="nw")

    def make_choose_user(self) -> None:
        """Add a button at the top-left corner"""
        self.choose_user = ttk.Button(self.widgets_frame, text="CHANGE USER", command=self.change_user)
        assert self.choose_user is not None, "self.input_entry should not be None at this point"
        self.choose_user.place(relx=0.5, rely=0, anchor="center")

    def make_exit_button(self) -> None:
        """exit app button"""
        button_exit = tk.Button(self.widgets_frame, text="EXIT", command=self.__end_app, width=5, bg='red', fg="black")
        button_exit.place(relx=0.95, rely=0.63, anchor="nw")

    def make_button_reun(self):
        """Makes buttons redo and undo and entry if successfully done"""
        add_folder_image1 = customtkinter.CTkImage(PILImage.open(self.REDO_IMAGE).resize((30, 30)))
        add_folder_image2 = customtkinter.CTkImage(PILImage.open(self.UNDO_IMAGE).resize((30, 30)))
        # redo button
        button_redo = customtkinter.CTkButton(self.widgets_frame, image=add_folder_image1, text="", command=self.redo,
                                              width=10, compound="left")
        button_redo.place(relx=0.55, rely=0.4, anchor="nw")
        # undo button
        button_undo = customtkinter.CTkButton(self.widgets_frame, image=add_folder_image2, text="", command=self.undo,
                                              width=10, compound="left")
        button_undo.place(relx=0.5, rely=0.4, anchor="nw")
        # if redo or undo successfully
        self.un_re = tk.Text(self.widgets_frame, height=1, width=20)
        self.un_re.insert("1.0", "")
        self.un_re.configure(state="disabled")
        self.un_re.place(relx=0.5, rely=0.6, anchor="nw")

    def make_treeview(self):
        """Makes TreeView according to dataframe"""
        ## Treeview Widget
        self.__treeview = ttk.Treeview(self.frame)
        # set the height and width of the widget to 100% of its container (frame)
        self.__treeview.place(relheight=1, relwidth=1)
        self.style = ttk.Style(self.__root)
        self.style.configure('Treeview', background='white',font=(self.FONT, 13, 'bold'), foreground='black')
        self.style.configure('Treeview.Heading', foreground='black', font=(self.FONT, 13, 'bold'), background='lightgrey')
        # command means update the yaxis view of the widget
        treescrolly = tk.Scrollbar(self.frame, orient="vertical",command=self.__treeview.yview)
        # command means update the xaxis view of the widget
        treescrollx = tk.Scrollbar(self.frame, orient="horizontal",command=self.__treeview.xview)
        # assign the scrollbars to the Treeview Widget
        self.__treeview.configure(xscrollcommand=treescrollx.set,yscrollcommand=treescrolly.set)
        # make the scrollbar fill the x axis of the Treeview widget
        treescrollx.pack(side="bottom", fill="x")
        # make the scrollbar fill the y axis of the Treeview widget
        treescrolly.pack(side="right", fill="y")
        # gets df of current user sheet
        df = self.__current_user.get_sheets()[self.__current_sheet].get_df()
        cols = df.columns
        cols.insert(0, "")
        # Insert empty heading as the first column
        self.__treeview["column"] = [""] + list(df.columns)
        # Show headings
        self.__treeview["show"] = "headings"
        for i, col in enumerate(self.__treeview["column"]):
            if i == 0:
                self.__treeview.column(col, width=20, anchor='center')  # Set smaller width for the first heading
            self.__treeview.column(col, width=60, anchor='center')
            self.__treeview.heading(col, text=col)
            self.__treeview.heading(col, anchor='center')
            self.__treeview.column(col, minwidth=120, anchor='center')
        # inserts each list into the treeview.
        for i, (index, row) in enumerate(df.iterrows(), 1):
            self.__treeview.insert("", "end", text=str(index), values=[index] + row.tolist())
        # Apply the 'col_0' tag to all items in column 0
        for child in self.__treeview.get_children():
            self.__treeview.item(child, tags=('col_0',))
        self.__treeview.pack(fill='both', expand=True)
        # Disable column resizing for the first column
        self.__treeview.bind("<B1-Motion>", lambda e: "break", add="+")  # Prevent column resizing on mouse drag
        self.__treeview.bind("<ButtonRelease-1>", self.on_item_click)

    def toggle_mode(self):
        # state of switch
        switch_state = self.switch_var.get()
        if switch_state == "on":
            # if on makes color green
            self.widgets_frame.config(bg="#5F9EA0")  # Change to the desired color
            self.frame.config(bg="#5F9EA0")  # Change to the desired color
        else:
            # if off makes color purple
            self.widgets_frame.config(bg="#CBC3E3")  # Change to the desired color
            self.frame.config(bg="#CBC3E3")  # Change to the desired color
    def create_main_board(self):
        """Creates the main board of the game, including the Treeview widget and all the buttons
        Restarts the data frame to be the default sheet - sheet number 1"""
        self.__root.configure(background='#00008B')  # Set background color to green
        self.__root.pack_propagate(False)  # tells the root to not let the widgets inside it determine its size.

        # Frame for TreeView
        self.frame = tk.LabelFrame(self.__root, text="Excel Data", highlightthickness=2, highlightbackground="black", bg="#5F9EA0")
        self.frame.place(height=390, width=1000)
        self.widgets_frame = tk.LabelFrame(self.__root, text="Manual", highlightthickness=2, highlightbackground="black", bg="#5F9EA0")
        self.widgets_frame.place(height=280, width=1000, rely=0.65, relx=0)
        # X-coordinate for the first button
        # Row entry - which row user clicked on
        self.make_row_entry()
        # Col entry - which column user clicked on
        self.make_col_entry()
        # Input of user what value wants in specific cell entry
        self.make_input_entry()
        # Change value of cell button
        self.make_button_change()
        # Delete cell button
        self.make_button_delete_cell()
        # file to upload entry
        self.make_file_upload()
        # Upload File button
        self.make_button_upload_file()
        # Add row button
        self.make_button_add_row()
        # Sheet number entry
        self.make_sheet_num()
        # Change sheet button
        self.make_button_change_sheet()
        # Save folder entry
        self.make_save_folder()
        # save button
        self.make_save_button()
        # change user button
        self.make_choose_user()
        # exit app button
        self.make_exit_button()
        # undo and redo buttons as well as entry
        self.make_button_reun()
        # make treeview
        self.make_treeview()
        # mode switch - to change theme of app
        self.switch_var = customtkinter.StringVar(value="on")
        self.mode_switch = customtkinter.CTkSwitch(self.widgets_frame, text="COLOR", command=self.toggle_mode,
                                                   variable=self.switch_var, onvalue="on", offvalue="off")
        self.mode_switch.place(relx=0.8, rely=0.63, anchor="nw")

    def invalid_file(self):
        """puts in upload file entry invalid file message"""
        self.file_upload.delete(0, "end")
        self.file_upload.insert(0, "INVALID_FILE")


    def upload_df(self):
        """When clicked on upload file, uploads the file gotten to the app.
        makes it the current sheet"""
        file_path = self.file_upload.get()
        try:
            # checks that file path is valid
            if os.path.exists(file_path):
                is_valid_upload = self.__current_user.upload_sheet(file_path)
                if is_valid_upload:
                    self.__current_sheet = os.path.basename(file_path)
                    self.refresh_screen()
                else:
                    self.invalid_file()
            else:
                self.invalid_file()
        except Exception as e:
            self.invalid_file()


    def undo(self):
        """undos action"""
        # checks if list of undo is empty, otherwise undoes action
        if self.__undo == []:
            self.un_re.config(state="normal")
            self.un_re.delete("1.0", "end")
            self.un_re.insert("1.0", "No actions to undo")
            self.un_re.config(state="disabled")
        else:
            self.action_back()
            self.un_re.config(state="normal")
            self.un_re.delete("1.0", "end")
            self.un_re.insert("1.0", "undo successfully")
            self.un_re.config(state="disabled")

    def redo(self):
        """redos action"""
        # checks if list of undo is empty, otherwise redoes action
        if self.__redo.empty():
            self.un_re.config(state="normal")
            self.un_re.delete("1.0", "end")
            self.un_re.insert("1.0", "No actions to redo")
            self.un_re.config(state="disabled")
        else:
            self.action_front()
            self.un_re.config(state="normal")
            self.un_re.delete("1.0", "end")
            self.un_re.insert("1.0", "redo successfully")
            self.un_re.config(state="disabled")


    def action_back(self) -> None:
        """"Pops the last action and put it in queue of redo"""
        action = self.__undo.pop()
        params = action[0]
        old_value = action[1]
        # after popping the action, puts it in redo queue
        self.__redo.put(action)
        # change the value of the cell
        self.apply_undo_redo(params, old_value)

    def action_front(self) -> None:
        """Gets the last action done from redo queue and puts it in undo lst"""
        action = self.__redo.get()
        self.__undo.append(action)
        # after redoing the action, puts it in undo list
        self.apply_undo_redo(action[0], action[2])
    def save_data(self):
        """When clicked on save data button, saves all the users sheet to the file gotten,
        saves by user name and name of sheet"""
        try:
            # gets folder wanted to save in
            folder_save = self.save_folder.get()
            # checks if valid folder
            if os.path.exists(folder_save):
                self.__current_user.save_all_data(folder_save)
                self.save_folder.delete(0, "end")
                self.save_folder.insert(0, "SAVED SUCCESSFULLY")
            else:
                self.save_folder.delete(0, "end")
                self.save_folder.insert(0, "INVALID FOLDER")

        except:
            self.save_folder.delete(0, "end")
            self.save_folder.insert(0, "INVALID_FOLDER")

    def add_row(self) -> None:
        """Adds a row to the data frame using Sheet object.
         After the data frame is changed, refreshes screen in order to implement the change in the current screen"""
        assert self.__current_user is not None, "self.input_entry should not be None at this point"
        self.__current_user.get_sheets()[self.__current_sheet].add_new_row()
        self.refresh_screen()

    def change_sheet(self):
        """When clicked on change sheet button, changes the current sheet to the sheet number gotten"""
        # delete queue
        self.__redo = queue.Queue()
        self.__undo = []
        self.__current_sheet = self.sheet_num.get()
        num_sheet = self.__current_sheet
        #check if user has sheet
        if not (num_sheet in self.__current_user.get_sheets()):
            self.__current_user.add_sheet(num_sheet)
        self.refresh_screen()

    def delete_cell(self):
        """When clicking on delete, deletes the cell gotten and refreshes the screen"""
        try:
            params = (self.col_entry.get("1.0", "end-1c"), self.row_entry.get("1.0", "end-1c")[0])
            row = self.col_entry.get("1.0", "end-1c")
            col = self.row_entry.get("1.0", "end-1c")
            old_value = self.get_cell_value(col, int(row))
            new_value = "_"
            # change the value of the cell to _
            self.__current_user.get_sheets()[self.__current_sheet].set_new_value(new_value, params)
            if old_value is None:
                old_value = "_"
            # updates undo list and redo queue
            self.__undo.append([params, old_value, "_"])
            self.change_value()
        except:
            self.input_entry.delete(0, "end")
            self.input_entry.insert(0, "Didnt choose cell to delete")

    def on_item_click(self, event):
        """When clicking on a cell in the treeview, gets the row and column of the cell clicked"""
        self.input_entry.delete(0, "end")
        self.input_entry.insert(0, "INPUT")
        try:
            item_id = self.__treeview.identify_row(event.y)  # Get the ID of the clicked row
            column_id = self.__treeview.identify_column(event.x)  # Get the ID of the clicked column
            # Convert the column ID to a numerical index
            column_index = int(column_id.replace("#", "")) - 1
            # Get the values of the clicked row
            values = self.__treeview.item(item_id, 'values')
            if values == "" or column_index == 0:
                pass
            else:
                # Map the Treeview row ID to the row number in your dataset
                row_number = self.__treeview.index(item_id)
                # changes the values in the entry
                self.row_entry.config(state="normal")
                self.col_entry.config(state="normal")
                self.row_entry.delete("1.0", "end")
                self.row_entry.insert("1.0", self.get_letter(int(column_index)))
                self.col_entry.delete("1.0","end")
                self.col_entry.insert("1.0", (row_number + 1))
                self.row_entry.config(state="disabled")
                self.col_entry.config(state="disabled")
        except:
            self.input_entry.delete(0, "end")
            self.input_entry.insert(0, "INVALID_SPOT")

    def get_letter(self, num: int) -> Union[str, None]:
        """Gets the letter of the number gotten (by what it represents in the data frame)"""
        try:
            number_letter_map = {1: 'A',2: 'B',3: 'C',4: 'D',5: 'E',6: 'F',7: 'G',8: 'H',9: 'I',10: 'J',11: 'K',12: 'L',13: 'M',14: 'N',15: 'O',16: 'P',17: 'Q',18: 'R',19: 'S',20: 'T',21: 'U',22: 'V',23: 'W',24: 'X',25: 'Y',26: 'Z'}
            return number_letter_map[num]
        except:
            return None

    def get_number(self, letter) -> Union[int, None]:
        """Gets the number of the letter gotten by what it represents in the data frame)"""
        try:
            letter = letter.upper()
            letter_number_map = {'A': 1,'B': 2,'C': 3,'D': 4,'E': 5,'F': 6,'G': 7,'H': 8,'I': 9,'J': 10,'K': 11,'L': 12,'M': 13,'N': 14,'O': 15,'P': 16,'Q': 17,'R': 18,'S': 19,'T': 20,'U': 21,'V': 22,'W': 23,'X':24,'Y': 25,'Z': 26}
            # Check if the number is in the dictionary, otherwise return None
            return letter_number_map[letter]
        except:
            return None

    def get_cell_value(self, row_index, col_index) -> Union[str, None]:
        """Gets the value of the cell in the data frame by the row and column gotten"""
        try:
            # Get the item ID (row ID) corresponding to the row index
            assert isinstance(self.__treeview, Treeview), "Expected self.__treeview to be a Treeview instance"

            item_id = self.__treeview.get_children()[col_index-1]
            # Get the values of the specified row
            values = self.__treeview.item(item_id, 'values')
            # Get the value of the specified column
            row_index = self.get_number(row_index)
            cell_value = values[row_index]
            return cell_value
        except:
            return "_"

    def apply_undo_redo(self, params, value):
        """When clicking on redo or undo changes the values of the cells.
                Uses sheet class to check validation on to change the value of the cells.
                If the value is not valid, changes the value to the reason it is not valid"""
        try:
            is_valid = Validation.check_valid_place(params[0], params[1])
            if not is_valid:
                self.input_entry.delete(0, "end")
                self.input_entry.insert(0, "INVALID_SPOT")
                self.input_entry.delete(0, "end")
                self.input_entry.insert(0, "INPUT")
            else:
                # checks that user input is valid
                is_valid, reason_not_valid = self.__current_user.get_sheets()[self.__current_sheet].validata_new_value(value)
                if is_valid:
                    # changes the value of the cell in the data frame of the user (in sheet class)
                    self.__current_user.get_sheets()[self.__current_sheet].set_new_value(value, params)
                    self.change_value()
                else:
                    self.input_entry.delete(0, "end")
                    self.input_entry.insert(0, reason_not_valid)
                    self.input_entry.delete(0, "end")
                    self.input_entry.insert(0, "INPUT")
        except Exception as e:
            self.input_entry.delete(0, "end")
            self.input_entry.insert(0, e)

    def apply_button(self):
        """When clicking on apply changes the values of the cells.
        Uses sheet class to check validation on to change the value of the cells.
        If the value is not valid, changes the value to the reason it is not valid"""
        try:
            row = self.col_entry.get("1.0", "end-1c")
            col = self.row_entry.get("1.0", "end-1c")
            # checks the place user clicked on the treeview is valid
            is_valid = Validation.check_valid_place(row,col)
            if not is_valid:
                self.input_entry.delete(0, "end")
                self.input_entry.insert(0, "INVALID_SPOT")
            else:
                params = (self.col_entry.get("1.0", "end-1c"), self.row_entry.get("1.0", "end-1c")[0])
                # recieves users input and validates it
                new_value = self.input_entry.get()
                if not Validation.check_cell_in_formula(params, new_value):
                    self.input_entry.delete(0, "end")
                    self.input_entry.insert(0, "Can't chose own cell")
                    return
                is_valid, reason_not_valid = self.__current_user.get_sheets()[self.__current_sheet].validata_new_value(new_value)
                if is_valid:
                    old_value = self.get_cell_value(col, int(row))
                    self.__current_user.get_sheets()[self.__current_sheet].set_new_value(new_value, params)
                    if old_value is None:
                        old_value = "_"
                    self.__undo.append([params, old_value, new_value])
                    self.change_value()
                else:
                    self.input_entry.delete(0, "end")
                    self.input_entry.insert(0, reason_not_valid)
        except Exception as e:
            self.input_entry.delete(0, "end")
            self.input_entry.insert(0, "Invalid formula")

    def refresh_screen(self):
        """Refreshes the data frame according to df of Sheet class"""
        # Clear existing items in the Treeview
        for item in self.__treeview.get_children():
            self.__treeview.delete(item)
        df = self.__current_user.get_sheets()[self.__current_sheet].get_df()
        # Get the DataFrame
        # Insert new values into the Treeview
        for i, (index, row) in enumerate(df.iterrows(), 1):
            self.__treeview.insert("", "end", text=str(index), values=[index] + row.tolist())
        self.input_entry.delete(0, "end")
        self.input_entry.insert(0, "INPUT")
        self.__treeview.bind("<ButtonRelease-1>", self.on_item_click)
        self.row_entry.config(state="normal")
        self.col_entry.config(state="normal")
        self.row_entry.delete("1.0", "end")
        self.row_entry.insert("1.0", "ROW")
        self.col_entry.delete("1.0", "end")
        self.col_entry.insert("1.0", "COL")
        self.row_entry.config(state="disabled")
        self.col_entry.config(state="disabled")

    def change_value(self):
        """Changes the value of the cell gotten by the row and column gotten"""
        self.refresh_screen()  # Update the values
        # Re-enable the ButtonRelease-1 event binding
        self.input_entry.delete(0, "end")
        self.input_entry.insert(0, "INPUT")
        self.__treeview.bind("<ButtonRelease-1>", self.on_item_click)

    def __end_app(self) -> None:
        """
        When time is up, a message prompt pops out, showing the user his final
        score, and asks if he wants to play another game. if so, another game
        starts. else, the program is terminated
        """
        sys.exit()
