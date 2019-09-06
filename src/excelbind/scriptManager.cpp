#include "scriptManager.h"
#include "xll12/xll/xll.h"

ScriptManager& ScriptManager::get()
{
	static ScriptManager instance;
	return instance;
}

int ScriptManager::finalizePython()
{
	get().scripts = py::module();
	py::finalize_interpreter();
	return 1;
}

int ScriptManager::initPython()
{
	py::initialize_interpreter();
	get().scripts = py::module::import("temp_script");

	return 1;
}

const py::module& ScriptManager::getScripts()
{ 
	return get().scripts; 
}

xll::Auto<xll::OpenBefore> init = xll::Auto<xll::OpenBefore>(&ScriptManager::initPython);
xll::Auto<xll::Close> finalize = xll::Auto<xll::Close>(&ScriptManager::finalizePython);
