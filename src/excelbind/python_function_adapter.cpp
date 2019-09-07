#include "script_manager.h"
#include "python_function_adapter.h"


PythonFunctionAdapter::PythonFunctionAdapter(const std::string& python_function_name, std::vector<BindTypes> argument_types, BindTypes return_type)
{
	python_function_name_ = python_function_name;
	argument_types_ = argument_types;
	return_type_ = return_type;
}

xll::LPOPER PythonFunctionAdapter::fct(void* p)
{
	static xll::OPER res_xll;

	py::object argument = convert_to_py_type(p, argument_types_[0]);
	const py::module& scripts = ScriptManager::get_scripts();
	py::object res_py = scripts.attr(python_function_name_.c_str())(argument);
	convert_to_xll_type(res_py, res_xll, return_type_);
	return &res_xll;
}
