import os.path
from src import *
import xlwings as xlw


def saveLastRow(path:str, sheet_name: int, index:int, phone_number:str, last_index:int) -> None:
    path = path.replace("/", "\\")
    file_name = path.split("\\")[-1].split('.')[0] + "-" + str(sheet_name) + ".txt"
    file_path = os.getcwd() + "\\src\\laststopbk\\" + file_name

    with open(file_path, "w+") as file:
        data = str(index) + "|" + phone_number + "|" + str(last_index)
        file.write(data)


def fetchLastRow(path:str, sheet_name: int) -> list:
    path = path.replace("/", "\\")
    file_name = path.split("\\")[-1].split('.')[0] + "-" + str(sheet_name) + ".txt"
    file_path = os.getcwd() + "\\src\\laststopbk\\" + file_name

    try:
        with open(file_path, "r") as file:
            data = file.read().split("|")
            return [int(data[0]), int(data[-1])]
    except:
        return [0, -1]