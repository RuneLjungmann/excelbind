#pragma warning(push)
#pragma warning(disable: 4996)
#pragma warning(disable: 4244)

#include <cstdlib>
#include "inih/INIReader.h"
#include "configuration.h"

Configuration::Configuration()
{
	std::string user_home_path = std::getenv("HOMEPATH");
	INIReader reader(user_home_path + "/excelbind.conf");
	is_error_messages_enabled_ = reader.GetBoolean("", "EnableErrorMessages", false);
}

Configuration& Configuration::get()
{
	static Configuration instance;
	return instance;
}

#pragma warning(pop)
