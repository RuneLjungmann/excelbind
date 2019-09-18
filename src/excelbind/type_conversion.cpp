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

std::string cast_string(const xll::OPER& oper)
{
    return cast_string(std::wstring(oper.val.str + 1, oper.val.str[0]));
}


py::object cast_oper_to_py(const xll::OPER& in)
{
    if (in.isBool())
    {
        return py::bool_(static_cast<bool>(in));
    }
    else if (in.isInt())
    {
        return py::int_(static_cast<int>(in));
    }
    else if (in.isNum())
    {
        return py::float_(static_cast<double>(in));
    }
    else if (in.isStr())
    {
        return py::str(cast_string(in));
    }

    return py::none();
}

void cast_py_to_oper(py::object in, xll::OPER& out)
{
    if (py::isinstance<py::str>(in))
    {
        out = in.cast<std::wstring>().c_str();
    }
    else if (py::isinstance<py::int_>(in))
    {
        out = in.cast<int>();
    }
    else if (py::isinstance<py::bool_>(in))
    {
        out = in.cast<bool>();
    }
    else if (py::isinstance<py::float_>(in))
    {
        out = in.cast<double>();
    }
    else if (py::isinstance<py::none>(in))
    {
        out = xll::OPER();
    }
    else
    {
        out = in.cast<double>();
    }
}


py::object cast_xll_to_py(void* p, BindTypes type)
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
		auto data = py::buffer_info(
			fp->array, sizeof(double), py::format_descriptor<double>::format(), 2,
			{ fp->rows, fp->columns }, { sizeof(double) * fp->columns, sizeof(double) }
		);
		return static_cast<py::object>(py::array_t<double>(data));
	}
    case BindTypes::BOOLEAN:
    {
        return py::bool_(*(bool*)(p));
    }
    case BindTypes::OPER:
    {
        return cast_oper_to_py(*(xll::OPER*)(p));
    }
    default:
		return py::object();
	}
}

void cast_py_to_xll(py::object in, xll::OPER& out, BindTypes type)
{
	switch (type)
	{
	case BindTypes::DOUBLE:
		out = in.cast<double>();
		break;
	case BindTypes::STRING:
		out = in.cast<std::wstring>().c_str();
		break;
	case BindTypes::ARRAY:
	{
		py::buffer_info buffer_info = static_cast<py::array_t<double>>(in).request();
		out = xll::OPER(static_cast<int>(buffer_info.shape[0]), static_cast<int>(buffer_info.shape[1]));

		double* src = static_cast<double*>(buffer_info.ptr);
		auto i = out.begin();
		while (i != out.end())
		{
			*i++ = *src++;
		}
		break;
	}
    case BindTypes::BOOLEAN:
        out = in.cast<bool>();
        break;
    case BindTypes::OPER:
        cast_py_to_oper(in, out);
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
		{ "numpy.ndarray", BindTypes::ARRAY },
        { "bool", BindTypes::BOOLEAN },
        { "Any", BindTypes::OPER }
    };

	auto i = typeConversionMap.find(py_type_name);
	if (i == typeConversionMap.end())
	{
		return BindTypes::OPER;
	}
	return i->second;
}

std::wstring get_xll_type(BindTypes type)
{
	static const std::map<BindTypes, const wchar_t*> conversionMap =
	{
		{ BindTypes::DOUBLE, XLL_DOUBLE_ },
		{ BindTypes::STRING, XLL_CSTRING },
		{ BindTypes::ARRAY, XLL_FP },
        { BindTypes::BOOLEAN, XLL_BOOL_ },
        { BindTypes::OPER, XLL_LPOPER }

	};
	return conversionMap.find(type)->second;
}
