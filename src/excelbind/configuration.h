#pragma once
class Configuration
{
private:
	Configuration();

	static Configuration& get();

public:

	static bool is_error_messages_enabled() { return get().is_error_messages_enabled_; }

private:
	bool is_error_messages_enabled_;
};

