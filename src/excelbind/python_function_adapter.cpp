#define NOMINMAX
#include <exception>
#include "xll12/xll/xll.h"

#include "script_manager.h"
#include "python_function_adapter.h"
#include "configuration.h"

PythonFunctionAdapter::PythonFunctionAdapter(const std::string& python_function_name, std::vector<BindTypes> argument_types, BindTypes return_type)
{
	python_function_name_ = python_function_name;
	argument_types_ = argument_types;
	return_type_ = return_type;
}

xll::LPOPER PythonFunctionAdapter::fct0() { return fct({ }); }
xll::LPOPER PythonFunctionAdapter::fct1(void* p0) { return fct({ p0 }); }
xll::LPOPER PythonFunctionAdapter::fct2(void* p0, void* p1) { return fct({ p0, p1 }); }
xll::LPOPER PythonFunctionAdapter::fct3(void* p0, void* p1, void* p2) { return fct({ p0, p1, p2 }); }
xll::LPOPER PythonFunctionAdapter::fct4(void* p0, void* p1, void* p2, void* p3) { return fct({ p0, p1, p2, p3 }); }
xll::LPOPER PythonFunctionAdapter::fct5(void* p0, void* p1, void* p2, void* p3, void* p4) { return fct({ p0, p1, p2, p3, p4 }); }
xll::LPOPER PythonFunctionAdapter::fct6(void* p0, void* p1, void* p2, void* p3, void* p4, void* p5) { return fct({ p0, p1, p2, p3, p4, p5 }); }
xll::LPOPER PythonFunctionAdapter::fct7(void* p0, void* p1, void* p2, void* p3, void* p4, void* p5, void* p6) { return fct({ p0, p1, p2, p3, p4, p5, p6 }); }
xll::LPOPER PythonFunctionAdapter::fct8(void* p0, void* p1, void* p2, void* p3, void* p4, void* p5, void* p6, void* p7) { return fct({ p0, p1, p2, p3, p4, p5, p6, p7 }); }
xll::LPOPER PythonFunctionAdapter::fct9(void* p0, void* p1, void* p2, void* p3, void* p4, void* p5, void* p6, void* p7, void* p8) { return fct({ p0, p1, p2, p3, p4, p5, p6, p7, p8 }); }
xll::LPOPER PythonFunctionAdapter::fct10(void* p0, void* p1, void* p2, void* p3, void* p4, void* p5, void* p6, void* p7, void* p8, void* p9) { return fct({ p0, p1, p2, p3, p4, p5, p6, p7, p8, p9 }); }

xll::LPOPER PythonFunctionAdapter::fct(const std::initializer_list<void*>& args)
{
	static xll::OPER res_xll;
	try
	{
		py::tuple args_py(args.size());
		auto j = args.begin();
		for (size_t i = 0; i < args.size(); ++i, ++j)
		{
			args_py[i] = cast_xll_to_py(*j, argument_types_[i]);
		}

		const py::module& scripts = ScriptManager::get_scripts();

		py::object res_py = scripts.attr(python_function_name_.c_str())(*args_py);

		cast_py_to_xll(res_py, res_xll, return_type_);
	}
	catch (const std::exception& e)
	{
		if (Configuration::is_error_messages_enabled())
		{
			XLL_ERROR(e.what());
		}
		res_xll = xll::OPER(xlerr::Num);
	}
	return &res_xll;
}


#define FCT_POINTER(num_args, ...) \
xll::LPOPER(__thiscall PythonFunctionAdapter:: * pFunc)(##__VA_ARGS__) = &PythonFunctionAdapter::fct##num_args; \
return (void*&)pFunc;

void* create_function_ptr(size_t num_arguments)
{
	switch (num_arguments)
	{
    case 0: { FCT_POINTER(0) }
    case 1: { FCT_POINTER(1, void*) }
    case 2: { FCT_POINTER(2, void*, void*) }
	case 3: { FCT_POINTER(3, void*, void*, void*) }
	case 4: { FCT_POINTER(4, void*, void*, void*, void*) }
	case 5: { FCT_POINTER(5, void*, void*, void*, void*, void*) }
	case 6: { FCT_POINTER(6, void*, void*, void*, void*, void*, void*) }
	case 7: { FCT_POINTER(7, void*, void*, void*, void*, void*, void*, void*) }
	case 8: { FCT_POINTER(8, void*, void*, void*, void*, void*, void*, void*, void*) }
	case 9: { FCT_POINTER(9, void*, void*, void*, void*, void*, void*, void*, void*, void*) }
	case 10: { FCT_POINTER(10, void*, void*, void*, void*, void*, void*, void*, void*, void*, void*) }
	default:
		throw std::runtime_error("A maximum of 10 arguments is supported for python functions.");
	}
}
