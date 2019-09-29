#pragma once
#include "pybind11/embed.h"


namespace py = pybind11;


class ScriptManager
{
private:
	ScriptManager() {}

	py::module scripts;

	static ScriptManager& get();

public:
	static const py::module& get_scripts();
	static int init_python();
	static int finalize_python();
	
};

