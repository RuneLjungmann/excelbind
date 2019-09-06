#pragma once
#include <Python.h>
#include "pybind11/embed.h"


namespace py = pybind11;


void registerPythonFunction(const py::str& pyFunctionName);
