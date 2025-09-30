import pandas as pd
import os
import numpy as np
from scipy.signal import savgol_filter
from scipy.interpolate import interp1d
from scipy.optimize import curve_fit
import sys

__all__ = ["Writer", "Reader", "derivative", "dm"]



class Reader:
    def __init__(self, filename: str, **kwargs) -> None:
        self.__filename = filename
        self.__sheets = []
        self.__sheetsN = 0
        self.__listOfCols = []
        self.__dataContainer = []
        self.__notifications = kwargs.get('notifications', False)

        # check if there is spec file
        if not os.path.isfile(self.__filename):
            raise FileNotFoundError(f"Файл '{self.__filename}' не найден!")

        # read table to __dataContainer to use it by getters
        self.__print_msg(f'Reading \'{os.path.basename(self.__filename)}\'')
        with pd.ExcelFile(self.__filename) as table:
            self.__sheets = table.sheet_names
            self.__sheetsN = len(self.__sheets)
            
            for sheet_name in self.__sheets:
                sheet = pd.read_excel(table, sheet_name=sheet_name, header=0)
                columns = sheet.columns.tolist()
                self.__listOfCols.append([str(c) for c in columns])
                self.__dataContainer.append([sheet[col].tolist() for col in columns])
                self.__print_msg(f'Sheet \'{sheet_name}\' has ' + str(len(columns)) + ' columns')


    def __print_msg(self, message: str) -> None:
        if self.__notifications:
            print(message)

    # return list (of sheet) of list (of columns) of list (of data)...
    def container(self):
        return self.__dataContainer
    
    # return column by name or index from sheet
    def column(self, sheet: int | str, column: int | str) -> list:
        """ Get column as list by (sheet,column) pair. Both sheet and column could be indexies or strings."""
        if isinstance(sheet, str) and isinstance(column, str):
            sheetIndex = self.__sheets.index(sheet)
            return self.__dataContainer[sheetIndex][self.__listOfCols[sheetIndex].index(column)]
        elif isinstance(sheet, int) and isinstance(column, int):
            return self.__dataContainer[sheet][column]
    
    # return names of sheets as a list
    def sheetsNames(self) -> list:
        return self.__sheets
    
    # return number of sheets
    def sheetsN(self) -> int:
        return self.__sheetsN
    
    # return colunms names from sheet as list
    def columnsNames(self, sheet: int | str) -> list[str]:
        if isinstance(sheet, str):
            return self.__listOfCols[self.__sheets.index(sheet)]
        elif isinstance(sheet, int):
            return self.__listOfCols[sheet]
            
    # return colunms names from sheet as list except 0
    def columnsNamesYs(self, sheet) -> list[str]:
        return np.delete(self.columnsNames(sheet), 0)
    
    # return number of cols from sheet
    def columnsN(self, sheet: int | str) -> int:
        if isinstance(sheet, str):
            return len(self.__listOfCols[self.__sheets.index(sheet)])
        elif isinstance(sheet, int):
            return len(self.__listOfCols[sheet])



class Writer:
    def __init__(self, **kwargs):
        self.sheets = {}
        self.__notifications = kwargs.get('notifications', False)

    def __print_msg(self, message: str) -> None:
        if self.__notifications:
            print(message)

    # turn any (almoost) input data to list
    def __tolist(self, data) -> list:
        if isinstance(data, list):
            return data
        elif isinstance(data, np.ndarray):
            return data.tolist()
        elif isinstance(data, (int, float, complex, str)):
            return [data]
        else:
            return [None]

    # add values to column (rewrite if there is data)
    def paste(self, sheet_name: str, column_name: str, data: list | int | float | complex | str | np.ndarray) -> None:
        sheet = self.sheets.setdefault(sheet_name, {})
        sheet[column_name] = self.__tolist(data)

    # add values to column (append if there is data)
    def add(self, sheet_name: str, column_name: str, data: list | int | float | complex | str | np.ndarray) -> None:
        sheet = self.sheets.setdefault(sheet_name, {})
        if column_name in sheet:
            sheet[column_name] += self.__tolist(data)
        else:
            sheet[column_name] = self.__tolist(data)

    # add values to column starting from row
    def add_to_row(self, sheet_name: str, column_name: str, data: list | int | float | complex | str | np.ndarray, start_row: int) -> None:
        if start_row < 0:
            raise ValueError("Row must be non-negative")
        data_list = self.__tolist(data)
        if not data_list:
            return
        sheet = self.sheets.setdefault(sheet_name, {})
        if column_name not in sheet:
            sheet[column_name] = [None] * start_row + data_list
        else:
            current_list = sheet[column_name]
            required_len = start_row + len(data_list)
            current_len = len(current_list)
            if required_len > current_len:
                current_list += [None] * (required_len - current_len)
            for i, value in enumerate(data_list):
                current_list[start_row + i] = value

    # saves table
    def save(self, filename: str) -> None:
        self.__print_msg(f'Writing \'{filename}\'')
        with pd.ExcelWriter(filename) as writer:
            for sheet_name, columns in self.sheets.items():
                if not columns:
                    continue
                max_len = max(len(v) for v in columns.values())
                aligned = {
                    name: values + [None]*(max_len - len(values))
                    for name, values in columns.items()
                }
                pd.DataFrame(aligned).to_excel(
                    writer,
                    sheet_name=sheet_name,
                    index=False
                )
                
        self.__print_msg(f'File \'{filename}\' has been written')



def _findExtermum(x, y, findmax):
    if findmax:
        return x[np.argmax(y)]
    else:
        return x[np.argmin(y)]

def _parabola(x, a, b, c):
    return a*x**2 + b*x + c

def derivative(*args, **kwargs):
    # обработка входных параметров
    if len(args) == 2:
        x = args[0]
        y = args[1]

    findmax = kwargs.get('findmax', True)
    doparabola = kwargs.get('useparabola', True)
    log = kwargs.get('log', False)
    x_start = kwargs.get('start', min(x))
    x_end = kwargs.get('end', max(x))
    interpolationtype = kwargs.get('interpolationtype', 'quadratic')
    interpolationpoints = kwargs.get('interpolationpoints', 1000)
    smoothsignal = kwargs.get('smoothsignal', True)
    if smoothsignal:
        smoothsignalwindow = kwargs.get('smoothsignalwindow', 50)
        polyordersignal = kwargs.get('polyordersignal', 3)
    smoothderiv = kwargs.get('smoothderivative', True)
    if smoothderiv:
        smoothderivwindow = kwargs.get('smoothderivativewindow', 50)
        polyorderderiv = kwargs.get('polyorderderivative', 3)
    if doparabola:
        parabolahalfwidth = kwargs.get('parabolahalfwidth', 4)

    # генерация нового x под интерполяцию
    x_interp = np.linspace(np.min(x), np.max(x), interpolationpoints)

    # логарифмирование
    if log:
        y = [np.log(n) for n in y]

    # интерполяция
    if smoothsignal:
        y_smoothed = savgol_filter(y, smoothsignalwindow, polyordersignal)
    else:
        y_smoothed = y
    Interpolation = interp1d(x, y_smoothed, kind = interpolationtype)
    y_interp = Interpolation(x_interp)

    # производная
    dy = np.gradient(y_interp, x_interp)
    if smoothderiv:
        dy_smoothed = savgol_filter(dy, smoothderivwindow, polyorderderiv)
    else:
        dy_smoothed = dy

    # Поиск экстремума
    start = (np.abs(x_interp - x_start)).argmin()
    end = (np.abs(x_interp - x_end)).argmin()
    x_extremum = _findExtermum(x_interp[start:end], dy_smoothed[start:end], findmax)

    if doparabola:
        par_start = (np.abs(x_interp - (x_extremum - parabolahalfwidth))).argmin()
        par_end = (np.abs(x_interp - (x_extremum + parabolahalfwidth))).argmin()
        parabolaPars, _  = curve_fit(_parabola, x_interp[par_start:par_end], dy[par_start:par_end])
        x_extremum = -parabolaPars[1]/2/parabolaPars[0]

    y_extremum = y_interp[(np.abs(x_interp - x_extremum)).argmin()]
        
    return (x_interp, dy), (x_extremum, y_extremum)
    
    
dm = sys.modules[__name__]