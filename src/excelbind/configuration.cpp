#pragma warning(push)
#pragma warning(disable: 4996)
#pragma warning(disable: 4244)

#include <cstdlib>
#include "inih/INIReader.h"
#include "configuration.h"


INIReader load_config_file()
{
    std::string user_home_path = std::getenv("HOMEPATH");
    INIReader reader(user_home_path + "/excelbind.conf");
    return reader;
}

Configuration::Configuration()
{
    INIReader reader = load_config_file();
    is_error_messages_enabled_ = reader.GetBoolean("", "EnableErrorMessages", false);
    virtual_env_ = reader.Get("", "VirtualEnv", "");
    module_dir_ = reader.Get("", "ModuleDir", "");
    module_name_ = reader.Get("", "ModuleName", "excelbind_export");
}

Configuration& Configuration::get()
{
	static Configuration instance;
	return instance;
}

#pragma warning(pop)
