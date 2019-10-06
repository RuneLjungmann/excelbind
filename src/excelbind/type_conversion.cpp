#define _SILENCE_CXX17_CODECVT_HEADER_DEPRECATION_WARNING

#include <codecvt>

#include "pybind11/numpy.h"

#include "date.h"
#include "chrono.h"
#include "type_conversion.h"

using namespace date;

typedef std::chrono::duration<double, std::ratio<60 * 60 * 24>> excel_datetime;
auto excel_base_time_point = 1899_y / December / 30;

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

bool has_dict_shape(const xll::OPER& in)
{
    return in.isMulti() && in.rows() > 1 && in.columns() > 1 && (in.rows() == 2 || in.columns() == 2);
}

bool has_list_shape(const xll::OPER& in)
{
    return in.isMulti() && (in.rows() == 1 || in.columns() == 1);
}


py::dict cast_oper_to_dict(const xll::OPER& in)
{
    if (in.columns() == 2)
    {
        py::dict out;
        for (int i = 0; i < in.rows(); ++i)
        {
            py::object key = cast_oper_to_py(in(i, 0));
            py::object val = cast_oper_to_py(in(i, 1));
            out[key] = val;
        }
        return out;
    }
    if (in.rows() == 2)
    {
        py::dict out;
        for (int i = 0; i < in.columns(); ++i)
        {
            py::object key = cast_oper_to_py(in(0, i));
            py::object val = cast_oper_to_py(in(1, i));
            out[key] = val;
        }
        return out;
    }
    return py::none();
}

void cast_dict_to_oper(const py::dict& in, xll::OPER& out)
{
    out = xll::OPER(static_cast<int>(in.size()), 2);
    int i = 0;
    for (auto item : in)
    {
        cast_py_to_oper(item.first, out(i, 0));
        cast_py_to_oper(item.second, out(i++, 1));
    }
}

py::list cast_oper_to_list(const xll::OPER& in)
{
    py::list out;
    for (auto item : in)
    {
        out.append(cast_oper_to_py(item));
    }
    return out;
}

void cast_list_to_oper(const py::list& in, xll::OPER& out)
{
    out = xll::OPER(static_cast<int>(in.size()), 1);
    auto oper_item = out.begin();   
    for (auto py_item : in)
    {
        cast_py_to_oper(py_item, *oper_item++);
    }
}

py::object cast_oper_to_dataframe(const xll::OPER& in)
{
    bool hasIndices = in(0, 0).isMissing() || in(0, 0).isNil();
    py::object indices = py::none();
    if (hasIndices)
    {
        py::list indices_list;
        for (int i = 1; i < in.rows(); ++i)
        {
            indices_list.append(cast_oper_to_py(in(i, 0)));
        }
        indices = indices_list;
    }
    py::list columns;
    for (int i = hasIndices ? 1 : 0; i < in.columns(); ++i)
    {
        columns.append(cast_oper_to_py(in(0, i)));
    }
    py::list data;
    for (int i = 1; i<in.rows(); ++i)
    {
        py::list row;
        for (int j = hasIndices ? 1 : 0; j < in.columns(); ++j)
        {
            row.append(cast_oper_to_py(in(i, j)));
        }
        data.append(row);
    }
    py::module pandas = py::module::import("pandas");
    return pandas.attr("DataFrame")(data, indices, columns);
}

void cast_dataframe_to_oper(const py::object& in, xll::OPER& out)
{
    py::list columns = py::list(in.attr("columns"));
    py::list index = py::list(in.attr("index"));
    out = xll::OPER(static_cast<int>(index.size() + 1), static_cast<int>(columns.size() + 1));
    for (size_t i = 0; i < index.size(); ++i)
    {
        cast_py_to_oper(index[i], out(static_cast<int>(i + 1), 0));
    }
    for (size_t i = 0; i < columns.size(); ++i)
    {
        cast_py_to_oper(columns[i], out(0, static_cast<int>(i + 1)));
    }
    for (size_t i = 0; i < index.size(); ++i)
    {
        for (size_t j = 0; j < columns.size(); ++j)
        {
            cast_py_to_oper(in.attr("iat")[py::make_tuple(i, j)], out(static_cast<int>(i + 1), static_cast<int>(j + 1)));
        }
    }
    out(0, 0) = L"";
}

py::object cast_oper_to_py(const xll::OPER& in)
{
    if (in.isBool())
    {
        return py::bool_(in == TRUE);
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
    else if (has_dict_shape(in))
    {
        return cast_oper_to_dict(in);
    }
    else if (has_list_shape(in))
    {
        return cast_oper_to_list(in);
    }

    return py::none();
}

void cast_py_to_oper(const py::handle& in, xll::OPER& out)
{
    if (py::isinstance<py::str>(in))
    {
        out = in.cast<std::wstring>().c_str();
    }
    else if (py::isinstance<py::bool_>(in))
    {
        out = in.cast<bool>();
    }
    else if (py::isinstance<py::dict>(in))
    {
        cast_dict_to_oper(in.cast<py::dict>(), out);
    }
    else if (py::isinstance<py::list>(in))
    {
        cast_list_to_oper(in.cast<py::list>(), out);
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
    case BindTypes::DICT:
    {
        return cast_oper_to_dict(*(xll::OPER*)(p));
    }
    case BindTypes::LIST:
    {
        return cast_oper_to_list(*(xll::OPER*)(p));
    }
    case BindTypes::DATETIME:
    {
        double excel_date = (*(double*)(p));
        excel_datetime ndays = excel_datetime(excel_date);
        return py::cast(std::chrono::system_clock::time_point(sys_days(excel_base_time_point)) + round<std::chrono::system_clock::duration>(ndays));
    }
    case BindTypes::PD_SERIES:
    {
        const auto& oper = *(xll::OPER*)(p);
        py::object data = has_list_shape(oper) ? static_cast<py::object>(cast_oper_to_list(oper)) : static_cast<py::object>(cast_oper_to_dict(oper));
        py::module pandas = py::module::import("pandas");
        return pandas.attr("Series")(data);
    }
    case BindTypes::PD_DATAFRAME:
    {
        return cast_oper_to_dataframe(*(xll::OPER*)(p));
    }
    default:
		return py::object();
	}
}

void cast_py_to_xll(const py::object& in, xll::OPER& out, BindTypes type)
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
    case BindTypes::DICT:
        cast_dict_to_oper(in.cast<py::dict>(), out);
        break;
    case BindTypes::LIST:
        cast_list_to_oper(in.cast<py::list>(), out);
        break;
    case BindTypes::DATETIME:
    {
        std::chrono::system_clock::time_point time_point = in.cast<std::chrono::system_clock::time_point>();
        auto duration = time_point - sys_days(excel_base_time_point);
        excel_datetime ndays = std::chrono::duration_cast<excel_datetime>(duration);
        out = ndays.count();
        break;
    }
    case BindTypes::PD_SERIES:
    {
        cast_dict_to_oper((in.attr("to_dict")()).cast<py::dict>(), out);
        break;
    }
    case BindTypes::PD_DATAFRAME:
    {
        cast_dataframe_to_oper(in, out);
        break;
    }
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
        { "Any", BindTypes::OPER },
        { "Dict", BindTypes::DICT },
        { "dict", BindTypes::DICT },
        { "List", BindTypes::LIST },
        { "list", BindTypes::LIST },
        { "datetime", BindTypes::DATETIME },
        { "datetime.datetime", BindTypes::DATETIME },
        { "pd.Series", BindTypes::PD_SERIES },
        { "pandas.Series", BindTypes::PD_SERIES },
        { "Series", BindTypes::PD_SERIES },
        { "pd.DataFrame", BindTypes::PD_DATAFRAME},
        { "pandas.DataFrame", BindTypes::PD_DATAFRAME },
        { "DataFrame", BindTypes::PD_DATAFRAME }
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
        { BindTypes::OPER, XLL_LPOPER },
        { BindTypes::DICT, XLL_LPOPER },
        { BindTypes::LIST, XLL_LPOPER },
        { BindTypes::DATETIME, XLL_DOUBLE_ },
        { BindTypes::PD_SERIES, XLL_LPOPER },
        { BindTypes::PD_DATAFRAME, XLL_LPOPER }
    };
	return conversionMap.find(type)->second;
}
