#pragma warning(push)
#pragma warning(disable: 4996)
#pragma warning(disable: 4244)

#include <cstdlib>
#include "inih/INIReader.h"
#include "configuration.h"


INIReader load_config_file()
{
    std::string user_home_path = std::getenv("USERPROFILE");
    INIReader reader(user_home_path + "/excelbind.conf");
    return reader;
}

std::string get_single_setting(const INIReader& reader, const std::string& settings_name, const std::string& default_setting)
{
    std::string env_var_name = "EXCELBIND_" + settings_name;
    std::transform(env_var_name.begin(), env_var_name.end(), env_var_name.begin(), [](unsigned char c) { return std::toupper(c); });
    char* env_var = std::getenv(env_var_name.c_str());
    return env_var ? env_var : reader.Get("excelbind", settings_name, default_setting);
}

Configuration::Configuration()
 : ini_header("excelbind")
{
    INIReader reader = load_config_file();
    is_error_messages_enabled_ = reader.GetBoolean(ini_header, "EnableErrorMessages", false);
    virtual_env_ = get_single_setting(reader, "VirtualEnv", "");
    module_dir_ = get_single_setting(reader, "ModuleDir", "");
    module_name_ = get_single_setting(reader, "ModuleName", "excelbind_export");
}

Configuration& Configuration::get()
{
	static Configuration instance;
	return instance;
}

#pragma warning(pop)
