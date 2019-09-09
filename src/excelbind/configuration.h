#pragma once
class Configuration
{
private:
	Configuration();

	static Configuration& get();

public:

	static bool isErrorMessagesEnabled() { return get().isErrorMessagesEnabled_; }

private:
	bool isErrorMessagesEnabled_;
};

