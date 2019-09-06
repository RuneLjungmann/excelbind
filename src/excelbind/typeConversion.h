#pragma once
#include <string>

enum class BindTypes { DOUBLE, STRING };

std::string cast_string(const std::wstring& in);

std::wstring cast_string(const std::string& in);
