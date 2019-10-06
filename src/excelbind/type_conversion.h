#pragma once

#include <string>
#include "pybind11/embed.h"
#include "xll12/xll/xll.h"


namespace py = pybind11;


enum class BindTypes { DOUBLE, STRING, ARRAY, BOOLEAN, OPER, DICT, LIST, DATETIME, PD_SERIES, PD_DATAFRAME };


std::string cast_string(const std::wstring& in);

std::wstring cast_string(const std::string& in);

BindTypes get_bind_type(const std::string& py_type_name);

std::wstring get_xll_type(BindTypes type);

py::object cast_oper_to_py(const xll::OPER& in);

void cast_py_to_oper(const py::handle& in, xll::OPER& out);

py::object cast_xll_to_py(void* p, BindTypes type);

void cast_py_to_xll(const py::object& in, xll::OPER& out, BindTypes type);
