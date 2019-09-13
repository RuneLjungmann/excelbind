from subprocess import run
import win32com.client as win32
import win32process
import time


def _kill_process_if_it_doesnt_shut_down(process_id, wait_time):
    for i in range(wait_time):
        if not _is_process_running(process_id):
            break
        time.sleep(1)
    else:
        _kill_process(process_id)


def _kill_process(process_id):
    cmd = f'taskkill /f /PID {process_id}'
    output = run(cmd, capture_output=True, universal_newlines=True)


def _is_process_running(process_id):
    cmd = f'tasklist /FI "PID eq {process_id}"'
    output = run(cmd, capture_output=True, universal_newlines=True)
    return str(process_id) in output.stdout


class Excel:
    """ Wrapper around the Excel com object.

    Handles clean shutdown of excel and exposes workbooks, worksheets and ranges in a 'builder' like pattern,
    so small test example sheets can be build and calculated easily.
    """
    def __init__(self, visible=False):
        self._visible = visible
        self._excel = None
        self._active_range = None
        self._process_id = None

    def __enter__(self):
        self._excel = win32.Dispatch('Excel.Application')
        thread_id, self._process_id = win32process.GetWindowThreadProcessId(self._excel.Hwnd)
        self._excel.Visible = self._visible
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        # Note the excessive use of 'del' is to make sure all python references to the pywin32 com objects are deleted
        # Otherwise a lot of errors are reported because there are references to dead com objects
        # which are later cleaned up by the gc
        del self._active_range
        wbs = self._excel.Workbooks
        for book in wbs:
            book.Close(False)
            del book
        del wbs

        # try to shut down excel nicely
        self._excel.Quit()
        del self._excel
        # then make sure to kill Excel if process is hanging for some reason - usually if test failed
        _kill_process_if_it_doesnt_shut_down(self._process_id, 5)

    def register_xll(self, xll_addin_path):
        self._excel.RegisterXLL(xll_addin_path)

    def calculate(self):
        ws = self._excel.ActiveSheet
        ws.Calculate()
        del ws
        return self

    def sheet(self, name):
        ws = self._excel.Worksheets(name)
        ws.Activate()
        del ws
        return self

    def book(self, name):
        wb = self._excel.Workbooks(name)
        wb.Activate()
        del wb
        return self

    def new_workbook(self):
        wb = self._excel.Workbooks.Add()
        wb.Activate()
        del wb
        return self

    def open(self, name):
        wb = self._excel.Workbooks.Open(name)
        wb.Activate()
        del wb
        return self

    def range(self, r):
        self._active_range = self._excel.Range(r)
        return self

    def cell(self, x, y):
        self._active_range = self._excel.Cells(x, y)
        return self

    def set(self, value):
        self._active_range.Value = value
        return self

    def set_formula(self, formula):
        if self._active_range.Count > 1:
            self._active_range.FormulaArray = formula
        else:
            self._active_range.Formula = formula
        return self

    @property
    def value(self):
        return self._active_range.Value
