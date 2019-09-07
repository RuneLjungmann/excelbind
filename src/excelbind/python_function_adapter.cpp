#include "script_manager.h"
#include "python_function_adapter.h"


PythonFunctionAdapter::PythonFunctionAdapter(const std::string& python_function_name, BindTypes argument_type)
{
	python_function_name_ = python_function_name;
	argument_type_ = argument_type;
}

xll::LPOPER PythonFunctionAdapter::fct(void* p)
{
	static xll::OPER res_xll;

	py::object argument = convert_to_py_type(p, argument_type_);
	const py::module& scripts = ScriptManager::get_scripts();
	py::object res_py = scripts.attr(python_function_name_.c_str())(argument);
	convert_to_xll_type(res_py, res_xll, BindTypes::STRING);
	return &res_xll;
}
