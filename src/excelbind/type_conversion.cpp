#define _SILENCE_CXX17_CODECVT_HEADER_DEPRECATION_WARNING

#include <codecvt>

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
		py::float_ f = *(double*)(p);
		return f;
	}
	case BindTypes::STRING:
	{
		py::str s = cast_string(std::wstring((XCHAR*)(p)));
		return s;
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
		{"float", BindTypes::DOUBLE},
		{"str", BindTypes::STRING}
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
		{BindTypes::DOUBLE, XLL_DOUBLE_},
		{BindTypes::STRING, XLL_CSTRING}
	};
	return conversionMap.find(type)->second;
}
