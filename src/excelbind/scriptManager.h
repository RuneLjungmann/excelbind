#pragma once
#include <Python.h>
#include "pybind11/embed.h"


namespace py = pybind11;


class ScriptManager
{
private:
	ScriptManager() {}

	py::module scripts;

	static ScriptManager& get();

public:
	static const py::module& getScripts();
	static int initPython();
	static int finalizePython();
	
};

