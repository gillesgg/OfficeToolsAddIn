#include "pch.h"
#include "Utility.h"
#include "Logger.h"


std::wstring Utility::FormatMessage(HRESULT hresult)
{
	if (HRESULT_FACILITY(hresult) == FACILITY_WINDOWS)
	{
		hresult = HRESULT_CODE(hresult);
	}
	DWORD	dwret;
	wchar_t* szError = 0;
	dwret = ::FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM, NULL, hresult, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPWSTR)&szError, 0, NULL);	
	if (dwret) 
	{
		return szError;
	}
	return (std::wstring());
}

std::wstring Utility::get_system_user_name()
{
	std::wstring result;
	const size_t initial_buf_size = 128;
	std::vector<wchar_t> buffer(initial_buf_size, 0);
	ULONG char_count = static_cast<ULONG>(buffer.size() - 1);
	const EXTENDED_NAME_FORMAT fmt = NameSamCompatible;
	if (!GetUserNameEx(fmt, &buffer[0], &char_count) && ERROR_MORE_DATA == ::GetLastError() && char_count > 0)
	{
		buffer.resize(char_count + 1, 0);
		if (!GetUserNameEx(fmt, &buffer[0], &char_count))
			return L"";
	}

	if (char_count > 0)
		result = &buffer[0];

	return result;
}


std::vector<std::wstring> Utility::tokenize(std::wstring str, const char delim)
{
	size_t start;
	size_t end = 0;
	std::vector<std::wstring> outlist;

	while ((start = str.find_first_not_of(delim, end)) != std::wstring::npos)
	{
		end = str.find(delim, start);
		outlist.push_back(str.substr(start, end - start));
	}
	return outlist;
}

int seh_filter(unsigned int code, struct _EXCEPTION_POINTERS* ep)
{
	// Generate error report
	// Execute exception handler
	return EXCEPTION_EXECUTE_HANDLER;
}


std::wstring Utility::GetSIDInfoFromUser(std::wstring user_name)
{
	DWORD cbSid = 0;
	DWORD cbDomain = 0;
	PSID pSid = nullptr;
	wchar_t* pszDomain = nullptr;
	SID_NAME_USE snu;

	const char delim = '\\';

	std::vector<std::wstring> outlist = tokenize(user_name, delim);
	
	if (outlist.size() == 2)
	{
		LookupAccountName(NULL, outlist[1].c_str(), pSid, &cbSid, pszDomain, &cbDomain, &snu);
		if (GetLastError() == ERROR_INSUFFICIENT_BUFFER)
		{
			pSid = new BYTE[cbSid];
			pszDomain = new wchar_t[cbDomain];
			if (LookupAccountName(NULL, outlist[1].c_str(), pSid, &cbSid, pszDomain, &cbDomain, &snu))
			{
				wchar_t* pszSid = NULL;
				if (ConvertSidToStringSid(pSid, &pszSid))
				{
					std::wstring str_sid = pszSid;
					LocalFree(pszSid);
					return (str_sid);
				}
			}
			else
			{
				LOG_ERROR << __FUNCTION__ << "-LookupAccountName failed, GetLastError=" << GetLastError();
			}
			delete[]pszDomain;
			delete[]pSid;
		}
		else
		{
			LOG_ERROR << __FUNCTION__ << "-LookupAccountName failed, not enought memory" << GetLastError();
		}
	}
	else
	{
		LOG_ERROR << __FUNCTION__ << "-LookupAccountName failed error=" << GetLastError();
	}
	
		
	return std::wstring();
}

void Utility::DeleteFile(std::wstring str_file)
{
	try
	{
		if (fs::exists(str_file) == true)
			fs::remove(fs::path(str_file));
	}
	catch (fs::filesystem_error& ex)
	{
		LOG_ERROR << __FUNCTION__ << "-unable to delete the file=" << ex.what();

	}
	
}

std::wstring Utility::GetUserInfo(std::wstring sid_str)
{
	SID_NAME_USE user_type;
	PSID sid = NULL;
	HRESULT ret = E_FAIL;
	if (ConvertStringSidToSid(sid_str.c_str(), &sid))
	{
		DWORD name_size = 0, domain_size = 0;
		if (!LookupAccountSid(NULL, sid, NULL, &name_size, NULL, &domain_size, &user_type) && ERROR_INSUFFICIENT_BUFFER != GetLastError())
		{
			LocalFree(sid);
			return std::wstring();
		}
		wchar_t* c_name = new wchar_t[name_size];

		if (!c_name)
		{
			LocalFree(sid);
			return std::wstring();
		}
		wchar_t* c_domain = new wchar_t[domain_size];
		if (!c_domain)
		{
			delete[] c_name;
			LocalFree(sid);
			return std::wstring();
		}
		if (LookupAccountSid(NULL, sid, c_name, &name_size, c_domain, &domain_size, &user_type))
		{
			ret = S_OK;
			std::wstring str_name = c_name;
			std::wstring str_domain = c_domain;
			return str_domain + L"\\" + str_name;
		}
		else
		{
			return (std::wstring());
		}
		delete[] c_name;
		delete[] c_domain;
		LocalFree(sid);
	}
	return std::wstring();
}

fs::path Utility::GetDirectoryWithCurrentExecutable()
{
	int size = 256;
	std::vector<wchar_t> charBuffer;
	// Let's be safe, and find the right buffer size programmatically.
	do {
		size *= 2;
		charBuffer.resize(size);

	} while (GetModuleFileName(NULL, charBuffer.data(), size) == size);

	fs::path path(charBuffer.data());  // Contains the full path including .exe
	return path.remove_filename();
}