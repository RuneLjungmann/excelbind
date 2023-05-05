// to support vs2015 - we import filesytem as experimental
#define _SILENCE_EXPERIMENTAL_FILESYSTEM_DEPRECATION_WARNING

#include <experimental/filesystem>
#include <filesystem>
#include "xll12/xll/xll.h"
#include "configuration.h"
#include "type_conversion.h"
#include "script_manager.h"


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

void add_python_helper_functions_to_excelbind_module()
{
    py::dict excelbind_scope = py::module::import("excelbind").attr("__dict__").cast<py::dict>();
    excelbind_scope[py::str("__builtins__")] = py::globals()[py::str("__builtins__")];

    py::exec(R"(
import os as _os
import re as _re
def _parse_doc_string(doc_string):
    param_regex = _re.compile(
        r'^:param (?P<param>\w+): (?P<doc>.*)$'
    )
    args_docs = {}
    function_doc_lines = []
    for item in doc_string.splitlines():
        g = param_regex.search(item.strip())
        if g:
            args_docs[g.group('param')] = g.group('doc')
        elif ':return:' not in item:
            function_doc_lines.append(item)

    while function_doc_lines and function_doc_lines[-1].strip() == '':
        function_doc_lines = function_doc_lines[:-1]

    return _os.linesep.join(function_doc_lines), args_docs


def _get_type_name(t):
    if hasattr(t, '__name__'):
        return t.__name__
    else:
        return t._name


def _wrapper(f, is_volatile=False):
    arguments = {arg_name: _get_type_name(f.__annotations__[arg_name]) if arg_name in f.__annotations__ else 'Any' for arg_name in f.__code__.co_varnames}
    return_type = _get_type_name(f.__annotations__['return']) if 'return' in f.__annotations__ else 'Any'
    raw_doc = f.__doc__ or ' '
    function_doc, arg_docs_dict = _parse_doc_string(raw_doc)
    arg_docs = [arg_docs_dict.get(item, ' ') for item in arguments.keys()]
    register(f.__name__, list(arguments.keys()), list(arguments.values()), arg_docs, return_type, function_doc, is_volatile)
    return f


def function(*args, **kwargs):
    if len(args) == 1 and len(kwargs) == 0 and callable(args[0]):
        return _wrapper(args[0])
    else:
        return lambda f: _wrapper(f, *args, **kwargs)

)", excelbind_scope);
}

void set_virtual_env_python_interpreter()
{
    const std::string virtual_env = Configuration::virtual_env();
    if (!std::experimental::filesystem::exists(virtual_env + "/Scripts/python.exe"))
    {
        std::string error_msg = "Failed to find python interpreter " + virtual_env + "/Scripts/python.exe";
        throw std::runtime_error(error_msg);
    }
             
    PyStatus status;
    PyConfig config;
    PyConfig_InitPythonConfig(&config);

    std::wstring virtual_env_w = cast_string(virtual_env);
    static std::wstring python_home = virtual_env_w + L"/Scripts";
    status = PyConfig_SetString(&config, &config.home, const_cast<wchar_t*>(python_home.c_str()));

    static std::wstring program_name = virtual_env_w + L"/Scripts/python.exe";
    status = PyConfig_SetString(&config, &config.program_name, const_cast<wchar_t*>(program_name.c_str()));

    static std::wstring python_path = virtual_env_w + L"/Lib;" + virtual_env_w + L"/Lib/site-packages";
    status = PyConfig_SetString(&config, &config.pythonpath_env, const_cast<wchar_t*>(python_path.c_str()));
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
    add_python_helper_functions_to_excelbind_module();
	get().scripts = py::module::import(Configuration::module_name().c_str());

	return 1;
}

const py::module& ScriptManager::get_scripts()
{ 
	return get().scripts; 
}

xll::Auto<xll::OpenBefore> init = xll::Auto<xll::OpenBefore>(&ScriptManager::init_python);
xll::Auto<xll::Close> finalize = xll::Auto<xll::Close>(&ScriptManager::finalize_python);
