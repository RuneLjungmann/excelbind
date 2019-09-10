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

void set_virtual_env_python_interpreter()
{
    std::wstring virtual_env = cast_string(Configuration::virtual_env());
    static std::wstring python_home = virtual_env + L"/Scripts";
    Py_SetPythonHome(python_home.c_str());

    static std::wstring program_name = virtual_env + L"/Scripts/python.exe";
    Py_SetProgramName(program_name.c_str());
    static std::wstring python_path = virtual_env + L"/Lib;" + virtual_env + L"/Lib/site-packages";
    Py_SetPath(python_path.c_str());
}

void init_virtual_env_paths()
{
    py::module::import("sys").attr("path").cast<py::list>().append(Configuration::virtual_env() + "/Scripts");
    py::module m = py::module::import("activate_this");
}

void add_module_dir_to_python_path()
{
    py::module::import("sys").attr("path").cast<py::list>().append(Configuration::module_dir());
}

int ScriptManager::init_python()
{
    if (Configuration::is_virtual_env_set())
    {
        set_virtual_env_python_interpreter();
    }
    
    py::initialize_interpreter();

    if (Configuration::is_virtual_env_set())
    {
        init_virtual_env_paths();
    }
    if (Configuration::is_module_dir_set())
    {
        add_module_dir_to_python_path();
    }

	get().scripts = py::module::import(Configuration::module_name().c_str());

	return 1;
}

const py::module& ScriptManager::get_scripts()
{ 
	return get().scripts; 
}

xll::Auto<xll::OpenBefore> init = xll::Auto<xll::OpenBefore>(&ScriptManager::init_python);
xll::Auto<xll::Close> finalize = xll::Auto<xll::Close>(&ScriptManager::finalize_python);
