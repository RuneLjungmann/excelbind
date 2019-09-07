#pragma once
#include "xll12/xll/xll.h"
#include "type_conversion.h"

class PythonFunctionAdapter
{
public:
	PythonFunctionAdapter(const std::string& python_function_name, BindTypes argument_type);

	xll::LPOPER fct(void* p0);

private:
	std::string python_function_name_;
	BindTypes argument_type_;
};
