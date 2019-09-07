#pragma once

#include <string>
#include <Python.h>
#include "pybind11/embed.h"
#include "xll12/xll/xll.h"


namespace py = pybind11;


enum class BindTypes { DOUBLE, STRING };



std::string cast_string(const std::wstring& in);

std::wstring cast_string(const std::string& in);

BindTypes get_bind_type(const std::string& py_type_name);

std::wstring get_xll_type(BindTypes type);

py::object convert_to_py_type(void* p, BindTypes type);

void convert_to_xll_type(py::object in, xll::OPER& out, BindTypes type);
