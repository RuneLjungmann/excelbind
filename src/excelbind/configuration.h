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

    static bool is_function_prefix_set() { return !function_prefix().empty(); }
    static const std::string& function_prefix() { return get().function_prefix_; }

    static const std::string& excel_category() { return get().excel_category_; }


private:
    const std::string ini_header;
	bool is_error_messages_enabled_;
    std::string virtual_env_;
    std::string module_dir_;
    std::string module_name_;
    std::string function_prefix_;
    std::string excel_category_;
};

