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

void register_python_function(const py::str& function_name_py, const py::list& argument_names_py, const py::list& argument_types_py, const py::str& return_type_py)
{
	static int function_index = 0;

	const std::string function_name = function_name_py;
	const std::wstring function_name_wide = cast_string(function_name);

	const std::wstring export_name = L"f" + std::to_wstring(function_index);
	const std::wstring xll_name = L"xll." + function_name_wide;

	std::vector<BindTypes> argument_types;
	for (auto& i : argument_types_py)
	{
		argument_types.push_back(get_bind_type(i.cast<std::string>()));
	}
	
	std::vector<std::wstring> argument_names;
	for (auto& i : argument_names_py)
	{
		argument_names.push_back(i.cast<std::wstring>());
	}
	
	BindTypes return_type = get_bind_type(return_type_py);

	// create function object and register it in thunks
	PythonFunctionAdapter* python_function = new PythonFunctionAdapter(function_name, argument_types, return_type);

	xll::LPOPER(__thiscall PythonFunctionAdapter:: * p_func)(void*) = &PythonFunctionAdapter::fct;
	thunks_objects[function_index] = python_function;
	thunks_methods[function_index] = (void*&)p_func;

	// Information Excel needs to register add-in.
	xll::Args functionBuilder = xll::Function(XLL_LPOPER, export_name.c_str(), xll_name.c_str())
		.Category(L"XLL")
		.FunctionHelp(L"some help text");

	for (size_t i = 0; i < argument_names.size(); ++i)
	{
		functionBuilder.Arg(get_xll_type(argument_types[i]).c_str(), argument_names[i].c_str(), L"Input");
	}

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
