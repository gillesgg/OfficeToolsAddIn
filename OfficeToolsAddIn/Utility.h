#pragma once
class Utility
{
public:
	static std::vector<std::wstring> tokenize(std::wstring str, const char delim);
	static std::wstring GetSIDInfoFromUser(std::wstring user_name);
	static std::wstring GetUserInfo(std::wstring sid_str);
	static std::wstring get_system_user_name();
	static std::wstring Utility::FormatMessage(HRESULT hresult);
	static void DeleteFile(std::wstring str_file);
	static fs::path GetDirectoryWithCurrentExecutable();
};

