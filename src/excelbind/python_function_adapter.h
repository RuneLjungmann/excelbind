#pragma once
#include "xll12/xll/xll.h"
#include "type_conversion.h"

class PythonFunctionAdapter
{
public:
	PythonFunctionAdapter(const std::string& python_function_name, std::vector<BindTypes> argument_types, BindTypes return_type);

    xll::LPOPER fct0();
    xll::LPOPER fct1(void* p0);
    xll::LPOPER fct2(void* p0, void* p1);
	xll::LPOPER fct3(void* p0, void* p1, void* p2);
	xll::LPOPER fct4(void* p0, void* p1, void* p2, void* p3);
	xll::LPOPER fct5(void* p0, void* p1, void* p2, void* p3, void* p4);
	xll::LPOPER fct6(void* p0, void* p1, void* p2, void* p3, void* p4, void* p5);
	xll::LPOPER fct7(void* p0, void* p1, void* p2, void* p3, void* p4, void* p5, void* p6);
	xll::LPOPER fct8(void* p0, void* p1, void* p2, void* p3, void* p4, void* p5, void* p6, void* p7);
	xll::LPOPER fct9(void* p0, void* p1, void* p2, void* p3, void* p4, void* p5, void* p6, void* p7, void* p8);
	xll::LPOPER fct10(void* p0, void* p1, void* p2, void* p3, void* p4, void* p5, void* p6, void* p7, void* p8, void* p9);

private:
	xll::LPOPER fct(const std::initializer_list<void*>& args);

	std::string python_function_name_;
	std::vector<BindTypes> argument_types_;
	BindTypes return_type_;
};

void* create_function_ptr(size_t num_arguments);
