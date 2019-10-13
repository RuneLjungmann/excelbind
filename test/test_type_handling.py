import pytest
import time
from test.utilities.env_vars import set_env_vars
from test.utilities.excel import Excel


def test_simple_math_function_with_floats(xll_addin_path):
    with set_env_vars('basic_functions'):
        with Excel() as excel:
            excel.register_xll(xll_addin_path)

            (
                excel.new_workbook()
                .range('A1').set(3.0)
                .range('A2').set(4.0)
                .range('B1').set_formula('=excelbind.add(A1, A2)')
                .range('B2').set_formula('=excelbind.mult(A1, A2)')
                .calculate()
            )

            assert excel.range('B1').value == 7.0
            assert excel.range('B2').value == 12.0


def test_simple_string_concatenation(xll_addin_path):
    with set_env_vars('basic_functions'):
        with Excel() as excel:
            excel.register_xll(xll_addin_path)

            (
                excel.new_workbook()
                .range('A1').set('Hello ')
                .range('A2').set('World!')
                .range('B1').set_formula('=excelbind.concat(A1, A2)')
                .calculate()
            )

            assert excel.range('B1').value == 'Hello World!'


def test_matrix_operations_with_np_ndarray(xll_addin_path):
    with set_env_vars('basic_functions'):
        with Excel() as excel:
            excel.register_xll(xll_addin_path)

            (
                excel.new_workbook()
                .range('A1').set(1.0)
                .range('B1').set(2.0)
                .range('A2').set(1.0)
                .range('B2').set(4.0)
                .range('A3').set_formula('=excelbind.det(A1:B2)')
                .range('A4:B5').set_formula('=excelbind.inv(A1:B2)')
                .calculate()
            )

            assert excel.range('A3').value == 2.0

            assert excel.range('A4').value == 2.0
            assert excel.range('B4').value == -1.0
            assert excel.range('A5').value == -0.5
            assert excel.range('B5').value == 0.5


def test_add_without_type_info(xll_addin_path):
    with set_env_vars('basic_functions'):
        with Excel() as excel:
            excel.register_xll(xll_addin_path)

            (
                excel.new_workbook()
                .range('A1').set(3.0)
                .range('A2').set(4.0)
                .range('B1').set_formula('=excelbind.add_without_type_info(A1, A2)')
                .range('A3').set('Hello ')
                .range('A4').set('world!')
                .range('B2').set_formula('=excelbind.add_without_type_info(A3, A4)')
                .calculate()
            )

            assert excel.range('B1').value == 7.0
            assert excel.range('B2').value == 'Hello world!'


def test_list_output(xll_addin_path):
    with set_env_vars('basic_functions'):
        with Excel() as excel:
            excel.register_xll(xll_addin_path)

            (
                excel.new_workbook()
                .range('A1').set(3.0)
                .range('A2').set(4.0)
                .range('A3').set(5.0)
                .range('D1:D3').set_formula('=excelbind.listify(A1, A2, A3)')
                .calculate()
            )

            assert excel.range('D1').value == 3.0
            assert excel.range('D2').value == 4.0
            assert excel.range('D3').value == 5.0


def test_dict_type(xll_addin_path):
    with set_env_vars('basic_functions'):
        with Excel() as excel:
            excel.register_xll(xll_addin_path)

            (
                excel.new_workbook()
                .range('A1').set('x')
                .range('A2').set('y')
                .range('A3').set('z')
                .range('B1').set(5.0)
                .range('B2').set(6.0)
                .range('B3').set(7.0)
                .range('D1:E2').set_formula('=excelbind.filter_dict(A1:B3, A2)')
                .calculate()
            )

            expected_dict = {'x': 5.0, 'z': 7.0}

            res_keys = sorted([excel.range('D1').value, excel.range('D2').value])
            assert res_keys == sorted(expected_dict.keys())
            assert excel.range('E1').value == expected_dict[excel.range('D1').value]
            assert excel.range('E2').value == expected_dict[excel.range('D2').value]


def test_list_in_various_directions(xll_addin_path):
    with set_env_vars('basic_functions'):
        with Excel() as excel:
            excel.register_xll(xll_addin_path)

            (
                excel.new_workbook()
                .range('A1').set(1.0)
                .range('A2').set(2.0)
                .range('A3').set(3.0)
                .range('B1').set(4.0)
                .range('C1').set(5.0)
                .range('D1').set_formula('=excelbind.dot(A1:A3, A1:C1)')
                .calculate()
            )

            assert excel.range('D1').value == 24.0


def test_no_arg(xll_addin_path):
    with set_env_vars('basic_functions'):
        with Excel() as excel:
            excel.register_xll(xll_addin_path)

            (
                excel.new_workbook()
                .range('A1').set_formula('=excelbind.no_arg()')
                .calculate()
            )

            assert excel.range('A1').value == 'Hello world!'


def test_date_to_python(xll_addin_path):
    with set_env_vars('basic_functions'):
        with Excel() as excel:
            excel.register_xll(xll_addin_path)

            (
                excel.new_workbook()
                .range('A1').set(29261)  # 1980-02-10
                .range('A2').set(33090)  # 1990-08-05
                .range('A3').set(25204)  # 1969-01-01
                .range('A4').set(25569)  # 1970-01-01
                .range('A5').set(36193.9549882755000000)  # 1999-02-02 22:55:11.987
                .range('B1').set_formula('=excelbind.date_as_string(A1)')
                .range('B2').set_formula('=excelbind.date_as_string(A2)')
                .range('B3').set_formula('=excelbind.date_as_string(A3)')
                .range('B4').set_formula('=excelbind.date_as_string(A4)')
                .range('B5').set_formula('=excelbind.date_as_string(A5)')
                .calculate()
            )

            assert excel.range('B1').value == '1980-02-10 00:00:00'
            assert excel.range('B2').value == '1990-08-05 00:00:00'
            assert excel.range('B3').value == '1969-01-01 00:00:00'
            assert excel.range('B4').value == '1970-01-01 00:00:00'
            assert excel.range('B5').value[:-3] == '1999-02-02 22:55:10.987'


def test_date_conversion(xll_addin_path):
    with set_env_vars('basic_functions'):
        with Excel() as excel:
            excel.register_xll(xll_addin_path)

            (
                excel.new_workbook()
                .range('A1').set(29261)  # 1980-02-10
                .range('B1').set_formula('=excelbind.just_the_date(A1)')
                .calculate()
            )

            assert excel.range('B1').value == excel.range('A1').value


def test_pandas_series(xll_addin_path):
    with set_env_vars('basic_functions'):
        with Excel() as excel:
            excel.register_xll(xll_addin_path)

            (
                excel.new_workbook()
                .range('A1').set(1)
                .range('B1').set(1.1)
                .range('A2').set(2)
                .range('B2').set(2.1)
                .range('A3').set(3)
                .range('B3').set(-1.1)
                .range('C1:D3').set_formula('=excelbind.pandas_series(A1:B3)')
                .calculate()
            )

            assert excel.range('D1').value == excel.range('B1').value
            assert excel.range('D2').value == excel.range('B2').value
            assert excel.range('D3').value == -excel.range('B3').value


def test_pandas_series_from_list(xll_addin_path):
    with set_env_vars('basic_functions'):
        with Excel() as excel:
            excel.register_xll(xll_addin_path)

            (
                excel.new_workbook()
                .range('A1').set(1)
                .range('B1').set(1.2)
                .range('A2').set(2)
                .range('B2').set(2.1)
                .range('A3').set(3)
                .range('B3').set(-1.1)
                .range('C1').set_formula('=excelbind.pandas_series_sum(A1:B3)')
                .range('D1').set_formula('=excelbind.pandas_series_sum(B1:B3)')
                .calculate()
            )

            assert excel.range('D1').value == excel.range('C1').value
            assert excel.range('C1').value == pytest.approx(2.2, abs=1e-10)


def test_pandas_dataframe(xll_addin_path):
    with set_env_vars('basic_functions'):
        with Excel() as excel:
            excel.register_xll(xll_addin_path)

            (
                excel.new_workbook()
                .range('B1').set("A")
                .range('C1').set("B")
                .range('A2').set(1)
                .range('B2').set(1.2)
                .range('C2').set(2.2)
                .range('A3').set(2)
                .range('B3').set(2.1)
                .range('C3').set(3.1)
                .range('A4').set(3)
                .range('B4').set(-1.1)
                .range('C4').set(-2.1)
                .range('D1:F4').set_formula('=excelbind.pandas_dataframe(A1:C4)')
                .range('G1').set_formula('=sum(B2:C4)')
                .range('H1').set_formula('=sum(E2:F4)')
                .calculate()
            )

            assert excel.range('D1').value == ""
            assert excel.range('A2').value == excel.range('D2').value
            assert excel.range('A3').value == excel.range('D3').value
            assert excel.range('A4').value == excel.range('D4').value

            assert excel.range('B1').value == excel.range('E1').value
            assert excel.range('C1').value == excel.range('F1').value

            assert excel.range('H1').value == excel.range('G1').value + 6*2


def test_volatile_function(xll_addin_path):
    with set_env_vars('basic_functions'):
        with Excel() as excel:
            excel.register_xll(xll_addin_path)

            (
                excel.new_workbook()
                .range('A1').set_formula('=excelbind.the_time()')
                .calculate()
            )

            time1 = excel.range('A1').value
            time.sleep(0.1)
            excel.calculate()
            time2 = excel.range('A1').value

            assert time1 != time2
