#pragma once
#include "xll12/xll/xll.h"
#include "type_conversion.h"

class PythonFunctionAdapter
{
public:
	PythonFunctionAdapter(const std::string& python_function_name, std::vector<BindTypes> argument_types, BindTypes return_type);

	xll::LPOPER fct(void* p0);

private:
	std::string python_function_name_;
	std::vector<BindTypes> argument_types_;
	BindTypes return_type_;
};
