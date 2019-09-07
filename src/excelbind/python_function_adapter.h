#pragma once
#include "xll12/xll/xll.h"
#include "type_conversion.h"

class PythonFunctionAdapter
{
public:
	PythonFunctionAdapter(const std::string& python_function_name, std::vector<BindTypes> argument_types, BindTypes return_type);

	xll::LPOPER fct1(void* p0);
	xll::LPOPER fct2(void* p0, void* p1);
	xll::LPOPER fct3(void* p0, void* p1, void* p2);
	xll::LPOPER fct4(void* p0, void* p1, void* p2, void* p3);
	xll::LPOPER fct5(void* p0, void* p1, void* p2, void* p3, void* p4);

private:
	std::string python_function_name_;
	std::vector<BindTypes> argument_types_;
	BindTypes return_type_;
};

void* create_function_ptr(unsigned num_arguments);
