#include "script_manager.h"
#include "python_function_adapter.h"


PythonFunctionAdapter::PythonFunctionAdapter(const std::string& python_function_name, std::vector<BindTypes> argument_types, BindTypes return_type)
{
	python_function_name_ = python_function_name;
	argument_types_ = argument_types;
	return_type_ = return_type;
}

xll::LPOPER PythonFunctionAdapter::fct1(void* p0)
{
	static xll::OPER res_xll;

	py::object argument = convert_to_py_type(p0, argument_types_[0]);
	const py::module& scripts = ScriptManager::get_scripts();
	py::object res_py = scripts.attr(python_function_name_.c_str())(argument);
	convert_to_xll_type(res_py, res_xll, return_type_);
	return &res_xll;
}

xll::LPOPER PythonFunctionAdapter::fct2(void* p0, void* p1)
{
	static xll::OPER res_xll;

	py::object argument0 = convert_to_py_type(p0, argument_types_[0]);
	py::object argument1 = convert_to_py_type(p1, argument_types_[1]);
	const py::module& scripts = ScriptManager::get_scripts();
	py::object res_py = scripts.attr(python_function_name_.c_str())(argument0, argument1);
	convert_to_xll_type(res_py, res_xll, return_type_);
	return &res_xll;
}

xll::LPOPER PythonFunctionAdapter::fct3(void* p0, void* p1, void* p2)
{
	static xll::OPER res_xll;

	py::object argument0 = convert_to_py_type(p0, argument_types_[0]);
	py::object argument1 = convert_to_py_type(p1, argument_types_[1]);
	py::object argument2 = convert_to_py_type(p2, argument_types_[2]);
	const py::module& scripts = ScriptManager::get_scripts();
	py::object res_py = scripts.attr(python_function_name_.c_str())(argument0, argument1, argument2);
	convert_to_xll_type(res_py, res_xll, return_type_);
	return &res_xll;
}

xll::LPOPER PythonFunctionAdapter::fct4(void* p0, void* p1, void* p2, void* p3)
{
	static xll::OPER res_xll;

	py::object argument0 = convert_to_py_type(p0, argument_types_[0]);
	py::object argument1 = convert_to_py_type(p1, argument_types_[1]);
	py::object argument2 = convert_to_py_type(p2, argument_types_[2]);
	py::object argument3 = convert_to_py_type(p3, argument_types_[3]);
	const py::module& scripts = ScriptManager::get_scripts();
	py::object res_py = scripts.attr(python_function_name_.c_str())(argument0, argument1, argument2, argument3);
	convert_to_xll_type(res_py, res_xll, return_type_);
	return &res_xll;
}

xll::LPOPER PythonFunctionAdapter::fct5(void* p0, void* p1, void* p2, void* p3, void* p4)
{
	static xll::OPER res_xll;

	py::object argument0 = convert_to_py_type(p0, argument_types_[0]);
	py::object argument1 = convert_to_py_type(p1, argument_types_[1]);
	py::object argument2 = convert_to_py_type(p2, argument_types_[2]);
	py::object argument3 = convert_to_py_type(p3, argument_types_[3]);
	py::object argument4 = convert_to_py_type(p4, argument_types_[4]);
	const py::module& scripts = ScriptManager::get_scripts();
	py::object res_py = scripts.attr(python_function_name_.c_str())(argument0, argument1, argument2, argument3, argument4);
	convert_to_xll_type(res_py, res_xll, return_type_);
	return &res_xll;
}

#define FCT_POINTER(num_args, ...) \
xll::LPOPER(__thiscall PythonFunctionAdapter:: * pFunc)(##__VA_ARGS__) = &PythonFunctionAdapter::fct##num_args; \
return (void*&)pFunc;

void* create_function_ptr(unsigned num_arguments)
{
	switch (num_arguments)
	{
	case 1:
	{
		FCT_POINTER(1, void*)
	}
	case 2:
	{
		FCT_POINTER(2, void*, void*)
	}
	case 3:
	{
		FCT_POINTER(3, void*, void*, void*)
	}
	case 4:
	{
		FCT_POINTER(4, void*, void*, void*, void*)
	}
	case 5:
	{
		FCT_POINTER(5, void*, void*, void*, void*, void*)
	}
	default:
		return nullptr;
	}
}
