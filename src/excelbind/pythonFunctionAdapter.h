#pragma once
#include "xll12/xll/xll.h"
#include "typeConversion.h"

class PythonFunctionAdapter
{
public:
	PythonFunctionAdapter(const std::string& pythonFunctionName, BindTypes argumentType);

	xll::LPOPER fct(void* p0);

private:
	std::string pythonFunctionName_;
	BindTypes argumentType_;
};
