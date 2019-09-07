#define _SILENCE_CXX17_CODECVT_HEADER_DEPRECATION_WARNING

#include <codecvt>

#include "pybind11/numpy.h"
#include "type_conversion.h"



std::string cast_string(const std::wstring& in)
{
	std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>> convert;
	return convert.to_bytes(in);
}

std::wstring cast_string(const std::string& in)
{
	std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>> convert;
	return convert.from_bytes(in);
}

py::object convert_to_py_type(void* p, BindTypes type)
{
	switch (type)
	{
	case BindTypes::DOUBLE:
	{
		return py::float_(*(double*)(p));
	}
	case BindTypes::STRING:
	{
		return py::str(cast_string(std::wstring((XCHAR*)(p))));
	}
	case BindTypes::ARRAY:
	{
		FP12* fp = (FP12*)(p);
		py::buffer_info data =
			py::buffer_info(
				fp->array, sizeof(double), py::format_descriptor<double>::format(), 2,
				{ fp->rows, fp->columns }, { sizeof(double) * fp->columns, sizeof(double) }
		);
		py::array_t<double> m(data);
		return (py::object)(m);
	}
	default:
		return py::object();
	}
}

void convert_to_xll_type(py::object in, xll::OPER& out, BindTypes type)
{
	switch (type)
	{
	case BindTypes::DOUBLE:
		out = in.cast<double>();
		break;
	case BindTypes::STRING:
		out = in.cast<std::wstring>().c_str();
		break;
	default:
		break;
	}
}

BindTypes get_bind_type(const std::string& py_type_name)
{
	static const std::map<std::string, BindTypes> typeConversionMap =
	{
		{ "float", BindTypes::DOUBLE },
		{ "str", BindTypes::STRING },
		{ "ndarray", BindTypes::ARRAY },
		{ "np.ndarray", BindTypes::ARRAY },
		{ "numpy.ndarray", BindTypes::ARRAY }
	};

	auto i = typeConversionMap.find(py_type_name);
	if (i == typeConversionMap.end())
	{
		return BindTypes::DOUBLE;
	}
	return i->second;
}

std::wstring get_xll_type(BindTypes type)
{
	static const std::map<BindTypes, const wchar_t*> conversionMap =
	{
		{ BindTypes::DOUBLE, XLL_DOUBLE_ },
		{ BindTypes::STRING, XLL_CSTRING },
		{ BindTypes::ARRAY, XLL_FP }
	};
	return conversionMap.find(type)->second;
}
