#pragma once
#include "xll12/xll/xll.h"

class PythonFunctionAdapter
{
public:
	PythonFunctionAdapter(const std::string& pythonFunctionName);

	xll::LPOPER fct(void* p0);

private:
	std::string pythonFunctionName_;
};
