#include "xll12/xll/xll.h"
#include "configuration.h"
#include "type_conversion.h"
#include "scriptManager.h"


ScriptManager& ScriptManager::get()
{
	static ScriptManager instance;
	return instance;
}

int ScriptManager::finalize_python()
{
	get().scripts = py::module();
	py::finalize_interpreter();
	return 1;
}

void init_virtual_env_paths()
{
    std::wstring virtual_env = cast_string(Configuration::virtual_env());
    static std::wstring python_home = virtual_env + L"/Scripts";
    Py_SetPythonHome(python_home.c_str());

    static std::wstring program_name = virtual_env + L"/Scripts/python.exe";
    Py_SetProgramName(program_name.c_str());
    static std::wstring python_path = virtual_env + L"/Lib;" + virtual_env + L"/Lib/site-packages";
    Py_SetPath(python_path.c_str());
}

void add_module_dir_to_python_path()
{
    static std::wstring python_path = cast_string(Configuration::module_dir()) + L";" + Py_GetPath();
    Py_SetPath(python_path.c_str());
}

int ScriptManager::init_python()
{
    if (Configuration::is_virtual_env_set())
    {
        init_virtual_env_paths();
    }
    if (Configuration::is_module_dir_set())
    {
        add_module_dir_to_python_path();
    }

	py::initialize_interpreter();
	get().scripts = py::module::import(Configuration::module_name().c_str());

	return 1;
}

const py::module& ScriptManager::get_scripts()
{ 
	return get().scripts; 
}

xll::Auto<xll::OpenBefore> init = xll::Auto<xll::OpenBefore>(&ScriptManager::init_python);
xll::Auto<xll::Close> finalize = xll::Auto<xll::Close>(&ScriptManager::finalize_python);
