import argparse
from enum import Enum

class ExcelTypeNotSupportError(Exception):
    def __init__(self, fileType:str) -> None:
        self.__fileType = fileType
    def __str__(self) -> str:
        return f"{self.__fileType}类型文件不支持\n"

class ExcelFileType(Enum):
    NotSupport = 0
    Xls = 1
    Xlsx = 2

class Configure:
    def __init__(self) -> None:
        self.__parser = argparse.ArgumentParser()
        self.__parser.add_argument("source", type=str)
        self.__parser.add_argument('--out', type=str)
        self.__parser.add_argument('--sheets', type=int, nargs='+',)
        self.__args = self.__parser.parse_args()

        self.__fileType = self.__setFileType()
        self.__tableTemplate = "./template/table_template"

    def __setFileType(self) -> ExcelFileType:
        temp = self.sourceFile().split('.')
        if temp[-1] == 'xls':
            return ExcelFileType.Xls
        elif temp[-1] == 'xlsx':
            return ExcelFileType.Xlsx
        else:
            raise ExcelTypeNotSupportError(temp[-1])
    
    def tableTemplate(self)->str:
        return self.__tableTemplate

    def sourceFile(self) -> str:
        return self.__args.source

    def outFile(self) -> str:
        return self.__args.out

    def sheets(self) -> list:
        return self.__args.sheets

    def fileType(self)->ExcelFileType:
        return self.__fileType
