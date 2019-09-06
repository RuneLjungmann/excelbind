#include <Python.h>
#include "pybind11/embed.h"

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

	py::object argument;
	if (argumentType_ == BindTypes::STRING)
	{
		py::str s = cast_string(std::wstring((XCHAR*)(p)));
		argument = s;
	}
	else
	{
		py::float_ d = *(double*)(p);
		argument = d;
	}

	const py::module& scripts = ScriptManager::getScripts();
	py::object res_py = scripts.attr(pythonFunctionName_.c_str())(argument);
	res_xll = res_py.cast<std::wstring>().c_str();
	return &res_xll;
}

