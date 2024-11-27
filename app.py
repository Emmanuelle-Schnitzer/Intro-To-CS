from sheet import Sheet
from user import User
from typing import Dict, Union


class App:
    def __init__(self) -> None:
        """
        A constructor for an App - holds all user session of Excel
        """
        # starting state of sheet in Excel
        self.__lst_users: Dict[User, Dict[str, Sheet]] = {}

    def check_user_in(self, name: str) -> bool:
        """Returns true if user is in the app's list of users, false otherwise"""
        for user in self.__lst_users:
            if user.get_name() == name:
                return True
        return False

    def add_new_user(self, name) -> User:
        """Adds a new user to the app's list of users, returns the new user added"""
        new_user = User(name)
        self.__lst_users[new_user] = new_user.get_sheets()
        return new_user

    def get_users(self) -> Dict[User, Dict[str, Sheet]]:
        """Returns the app's list of users"""
        return self.__lst_users

    def get_user(self, name) -> User:
        """Returns the user with the given name from the app's list of users, if it exists, None otherwise"""
        for key in self.__lst_users:
            if key.get_name() == name:
                return key
        return User("")

