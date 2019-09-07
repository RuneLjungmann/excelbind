#pragma once
#include <Python.h>
#include "pybind11/embed.h"


namespace py = pybind11;


void register_python_function(const py::str& function_name_py, const py::list& argument_names_py, const py::list& argument_types_py, const py::str& return_type_py);
