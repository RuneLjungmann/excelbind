#pragma once
class Configuration
{
private:
	Configuration();

	static Configuration& get();

public:

	static bool is_error_messages_enabled() { return get().is_error_messages_enabled_; }

    static const std::string& virtual_env() { return get().virtual_env_; }
    static bool is_virtual_env_set() { return !virtual_env().empty(); }

    static const std::string& module_dir() { return get().module_dir_; }
    static bool is_module_dir_set() { return !module_dir().empty(); }

    static const std::string& module_name() { return get().module_name_; }

private:
	bool is_error_messages_enabled_;
    std::string virtual_env_;
    std::string module_dir_;
    std::string module_name_;
};

