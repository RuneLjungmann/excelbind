/*
This code handles dynamic registration of a python function.

The approach is very much inspired by the setup in ExcelDNA. 
However, to support the more dynamic approach to registering which you would want from python, 
the idea is extended a bit, using a non-standard C++ feature.
So, when a python function is registered, a pythonFunctionAdapter object is created, 
which holds a reference to the python function name and a non-virtual method used to invoke the function.
For this to work from Excel, we need to store both a pointer to the object and the member function.
The Visual Studio calling conventions for non-virtual member functions are similar to calling free functions, 
except that the pointer to the object is in register ecx.

More concretley, when a python function is registered, an adapter object is created, and 
pointers to both the object and the member function are stored in thunksObjects and thunksMethods respectively.
The free function exposed to Excel (the expf functions created by macros below) then moves the object adress to ecx and jumps to member function.
*/

#include "type_conversion.h"
#include "python_function_adapter.h"
#include "function_registration.h"


// thunk tables 
extern "C"
{
	void* thunks_objects[10];
	void* thunks_methods[10];
}

void register_python_function(const py::str& py_function_name, const py::str& argument_type)
{
	static int function_index = 0;

	const std::string function_name = py_function_name;
	const std::wstring function_name_wide = cast_string(function_name);

	const std::wstring export_name = L"f" + std::to_wstring(function_index);
	const std::wstring xll_name = L"xll." + function_name_wide;

	BindTypes arg_type_internal = get_bind_type(argument_type);
	std::wstring arg_type_xll = get_xll_type(arg_type_internal);

	// create function object and register it in thunks
	PythonFunctionAdapter* python_function = new PythonFunctionAdapter(function_name, arg_type_internal);

	xll::LPOPER(__thiscall PythonFunctionAdapter:: * p_func)(void*) = &PythonFunctionAdapter::fct;

	thunks_objects[function_index] = python_function;
	thunks_methods[function_index] = (void*&)p_func;


	// Information Excel needs to register add-in.
	xll::Args functionBuilder = xll::Function(XLL_LPOPER, export_name.c_str(), xll_name.c_str())
		.Arg(arg_type_xll.c_str(), L"x", L"input")
		.Category(L"XLL")
		.FunctionHelp(L"some help text");

	xll::AddIn function = xll::AddIn(functionBuilder);
	++function_index;
}

PYBIND11_EMBEDDED_MODULE(excelbind, m) {
	// `m` is a `py::module` which is used to bind functions and classes
	m.def("register", &register_python_function);
}


#define expf(i) extern "C" __declspec(dllexport,naked)	void f##i(void) {  \
__asm mov ecx, thunks_objects + i * 4 \
__asm jmp thunks_methods + i * 4 \
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
