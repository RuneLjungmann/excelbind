#pragma once
#include <Python.h>
#include "pybind11/embed.h"


namespace py = pybind11;


void register_python_function(const py::str& py_function_name, const py::str& argument_type);
