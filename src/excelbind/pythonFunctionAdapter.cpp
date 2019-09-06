#include <Python.h>
#include "pybind11/embed.h"

#include "scriptManager.h"

#include "pythonFunctionAdapter.h"


PythonFunctionAdapter::PythonFunctionAdapter(const std::string& pythonFunctionName)
{
	pythonFunctionName_ = pythonFunctionName;
}

xll::LPOPER PythonFunctionAdapter::fct(void* p)
{
	static xll::OPER res_xll;

	py::float_ dbl = *(double*)(p);
	const py::module& scripts = ScriptManager::getScripts();
	py::object res_py = scripts.attr(pythonFunctionName_.c_str())(dbl);
	res_xll = res_py.cast<double>();
	return &res_xll;
}

