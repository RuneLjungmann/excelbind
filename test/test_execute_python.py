from test.utilities.env_vars import set_env_vars
from test.utilities.excel import open_excel


def test_simple_script_for_addition(xll_addin_path):
    with set_env_vars('basic_functions'):
        with open_excel() as excel:
            excel.RegisterXLL(xll_addin_path)

            (
                excel.new_workbook()
                .range('A1').set(3.0)
                .range('A2').set(4.0)
                .range('B1').set_formula('=excelbind.execute_python("return arg0 + arg1", A1, A2)')
                .calculate()
            )

            assert excel.range('B1').value == 7.0


def test_combination_str_n_float(xll_addin_path):
    with set_env_vars('basic_functions'):
        with open_excel() as excel:
            excel.RegisterXLL(xll_addin_path)

            (
                excel.new_workbook()
                .range('A1').set("Hello times ")
                .range('A2').set(3.0)
                .range('B1').set_formula('=excelbind.execute_python("return arg0 + str(arg1)", A1, A2)')
                .calculate()
            )

            assert excel.range('B1').value == 'Hello times 3.0'
