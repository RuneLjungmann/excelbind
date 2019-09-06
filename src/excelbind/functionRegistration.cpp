#define _SILENCE_CXX17_CODECVT_HEADER_DEPRECATION_WARNING

#include <codecvt>

#include "pythonFunctionAdapter.h"
#include "functionRegistration.h"


// thunk table 
extern "C"
{
	void* thunksObjects[10];
	void* thunksMethods[10];
}

void registerPythonFunction(const py::str& pyFunctionName)
{
	static int functionIndex = 0;

	const std::string functionName = pyFunctionName;
	std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>> convert;
	const std::wstring functionName_wide = convert.from_bytes(functionName);

	const std::wstring exportName = L"f" + std::to_wstring(functionIndex);
	const std::wstring xllName = L"xll." + functionName_wide;

	// create function object and register it in thunks
	PythonFunctionAdapter* pythonFunction = new PythonFunctionAdapter(functionName);

	xll::LPOPER(__thiscall PythonFunctionAdapter:: * pFunc)(void*) = &PythonFunctionAdapter::fct;

	thunksMethods[functionIndex] = (void*&)pFunc;
	thunksObjects[functionIndex] = pythonFunction;

	// Information Excel needs to register add-in.
	xll::Args functionBuilder = xll::Function(XLL_LPOPER, exportName.c_str(), xllName.c_str())
		.Arg(XLL_DOUBLE_, L"x", L"input")
		.Category(L"XLL")
		.FunctionHelp(L"some help text");

	xll::AddIn function = xll::AddIn(functionBuilder);
	++functionIndex;
}

PYBIND11_EMBEDDED_MODULE(excelbind, m) {
	// `m` is a `py::module` which is used to bind functions and classes
	m.def("register", &registerPythonFunction);
}


#define expf(i) extern "C" __declspec(dllexport,naked)	void f##i(void) {  \
__asm mov ecx, thunksObjects + i * 4 \
__asm jmp thunksMethods + i * 4 \
} \


expf(0)
expf(1)
expf(2)
expf(3)
expf(4)
expf(5)
expf(6)
expf(7)
expf(8)
expf(9)
