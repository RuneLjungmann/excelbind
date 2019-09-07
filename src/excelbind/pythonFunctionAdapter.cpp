#include "scriptManager.h"
#include "pythonFunctionAdapter.h"


PythonFunctionAdapter::PythonFunctionAdapter(const std::string& pythonFunctionName, BindTypes argumentType)
{
	pythonFunctionName_ = pythonFunctionName;
	argumentType_ = argumentType;
}

xll::LPOPER PythonFunctionAdapter::fct(void* p)
{
	static xll::OPER res_xll;

	py::object argument = convert_to_py_type(p, argumentType_);
	const py::module& scripts = ScriptManager::getScripts();
	py::object res_py = scripts.attr(pythonFunctionName_.c_str())(argument);
	convert_to_xll_type(res_py, res_xll, BindTypes::STRING);
	return &res_xll;
}

