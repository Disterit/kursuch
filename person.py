from datetime import date


class Person:
    def __init__(self, full_name: str, Date: date, department: str,
                 family_status: str, position: str, work_exp: int):

        self.__full_name = full_name
        self.__Date = Date
        self.__department = department
        self.__family_status = family_status
        self.__position = position
        self.__work_exp = work_exp

    @property
    def full_name(self):
        return self.__full_name

    @property
    def Date(self):
        return self.__Date

    @property
    def departament(self):
        return self.__department

    @property
    def family_status(self):
        return self.__family_status

    @property
    def position(self):
        return self.__position

    @property
    def work_exp(self):
        return self.__work_exp