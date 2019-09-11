import os
import sys
import pathlib
from contextlib import contextmanager
from .excel_test_utilities import open_excel


@contextmanager
def set_env_vars(module_name):
    try:
        backup_env = os.environ.copy()
        os.environ['EXCELBIND_MODULEDIR'] = str(pathlib.Path(__file__).parent / '..' / 'examples')
        os.environ['EXCELBIND_MODULENAME'] = module_name
        if hasattr(sys, 'real_prefix'):
            os.environ['EXCELBIND_VIRTUALENV'] = sys.prefix

        yield
    finally:
        os.environ = backup_env


def test_simple_math_function_with_floats(xll_addin_path):
    with set_env_vars('basic_functions'):
        with open_excel() as excel:
            excel.RegisterXLL(xll_addin_path)

            (
                excel.new_workbook()
                .range('A1').set(3.0)
                .range('A2').set(4.0)
                .range('B1').set_formula("=xll.add(A1, A2)")
                .range('B2').set_formula("=xll.mult(A1, A2)")
                .calculate()
            )

            assert excel.range('B1').value == 7.0
            assert excel.range('B2').value == 12.0


def test_simple_string_concatenation(xll_addin_path):
    with set_env_vars('basic_functions'):
        with open_excel() as excel:
            excel.RegisterXLL(xll_addin_path)

            (
                excel.new_workbook()
                .range('A1').set('Hello ')
                .range('A2').set('World!')
                .range('B1').set_formula("=xll.concat(A1, A2)")
                .calculate()
            )

            assert excel.range('B1').value == 'Hello World!'


def test_matrix_operations_with_np_ndarray(xll_addin_path):
    with set_env_vars('basic_functions'):
        with open_excel() as excel:
            excel.RegisterXLL(xll_addin_path)

            (
                excel.new_workbook()
                .range('A1').set(1.0)
                .range('B1').set(2.0)
                .range('A2').set(1.0)
                .range('B2').set(4.0)
                .range('A3').set_formula("=xll.det(A1:B2)")
                .range('A4:B5').set_formula("=xll.inv(A1:B2)")
                .calculate()
            )

            assert excel.range('A3').value == 2.0

            assert excel.range('A4').value == 2.0
            assert excel.range('B4').value == -1.0
            assert excel.range('A5').value == -0.5
            assert excel.range('B5').value == 0.5
