from subprocess import run
import win32com.client as win32
import win32process
import time
from contextlib import contextmanager


@contextmanager
def open_excel(visible=False):
    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        thread_id, process_id = win32process.GetWindowThreadProcessId(excel.Hwnd)
        excel.Visible = visible
        wrapper = ActivationWrapper(excel)
        yield wrapper
    except Exception as e:
        pass
    finally:
        # first try to shut Excel down nicely
        for book in excel.Workbooks:
            book.Close(False)
        excel.Quit()
        wrapper._active_range = None
        wrapper._wrapped_obj = None

        # then make sure to kill Excel if process is hanging for some reason
        # noinspection SpellCheckingInspection
        time.sleep(1)
        cmd = f'taskkill /f /PID {process_id}'
        run(cmd, capture_output=False, universal_newlines=True)


class ActivationWrapper:

    def __init__(self, wrapped_obj):
        self._wrapped_obj = wrapped_obj
        self._active_range = None

    def __getattr__(self, attr):
        if attr in self.__dict__:
            return getattr(self, attr)
        return getattr(self._wrapped_obj, attr)

    def calculate(self):
        self.ActiveSheet.Calculate()
        return self

    def sheet(self, name):
        self.Worksheets(name).Activate()
        return self

    def book(self, name):
        self.Workbooks(name).Activate()
        return self

    def new_workbook(self):
        workbook = self.Workbooks.Add()
        workbook.Activate()
        return self

    def open(self, name):
        workbook = self.Workbooks.Open(name)
        workbook.Activate()
        return self

    def range(self, r):
        self._active_range = self.Range(r)
        return self

    def cell(self, x, y):
        self._active_range = self.Cells(x, y)
        return self

    def set(self, value):
        self._active_range.Value = value
        return self

    def set_formula(self, formula):
        self._active_range.FormulaArray = formula
        return self

    @property
    def value(self):
        return self._active_range.Value
