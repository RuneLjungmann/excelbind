#include "excelbind.h"

// Information Excel needs to register add-in.
xll::AddIn xai_function(
	// Function returning a pointer to an OPER with C++ name xll_function and Excel name XLL.FUNCTION.
	// Don't forget prepend a question mark to the C++ name.
	//                     |
	//                     v
	xll::Function(XLL_LPOPER, L"?xll_function", L"XLL.FUNCTION")
	// First argument is a double called x with an argument description.
	.Arg(XLL_DOUBLE, L"x", L"is the first double argument.")
	// Paste function category.
	.Category(CATEGORY)
	// Insert Function description.
	.FunctionHelp(L"Help on XLL.FUNCTION goes here.")
);

// Calling convention *must* be WINAPI (aka __stdcall) for Excel.
xll::LPOPER WINAPI xll_function(double x)
{
	// Be sure to export your function.
#pragma XLLEXPORT
	static xll::OPER result;

	try {
		ensure(x >= 0);
		result = sqrt(x); // OPER's act like Excel cells.
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		result = xll::OPER(xlerr::Num);
	}

	return &result;
}
