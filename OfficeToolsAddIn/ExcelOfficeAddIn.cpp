#include "pch.h"
#include "ExcelOfficeAddIn.h"
#include "Logger.h"
#include "Utility.h"
#include "XLSingleton.h"

#define	IS_KEY_LEN 512

/// <summary>
/// Enum All Excel AddIn & Disable All Excel COM Add using the registry
/// </summary>
/// <returns></returns>
HRESULT ExcelCOMAddIn::DisableAllAddin()
{
	LOG_DEBUG << __FUNCTION__;

	HRESULT hr = S_OK;
	ProcessInformation processinformation = XLSingleton::getInstance()->Get_Addin_info();

	for (auto it = processinformation.addininformation_.begin(); it != processinformation.addininformation_.end(); it++)
	{
		if (it->second.addType_ == AddInType::OFFICE)
		{
			if ((it->second.parent_ != HKEY_LOCAL_MACHINE))
			{
				LOG_INFO << "Set Addin LoadBehavior Addin Name=" << it->second.sid_account + L"_" + it->second.ProgId_ << "Description=" << it->second.Description_ << " progID=" << it->second.ProgId_ << it->second.LoadBehavior_;
				hr = WriteLoadBehaviorRegistryInformation(it->second.parent_, it->second.key_, it->second.LoadBehavior_);
				if (FAILED(hr))
					return hr;
			}
		}
		
	}
	return (hr);
}
/// <summary>
/// Disable one Excel COM Add using the registry
/// </summary>
/// <param name="parent">the HKEY parent HKEY_CURRENT_USER or HKEY_USERS</param>
/// <param name="str_key">the key name</param>
/// <param name="dw_value">the LoadBehavior value</param>
/// <returns></returns>
HRESULT ExcelCOMAddIn::WriteLoadBehaviorRegistryInformation(HKEY parent, std::wstring str_key, DWORD dw_value)
{
	LOG_DEBUG << __FUNCTION__;

	LSTATUS status;
	CRegKey key;
	HRESULT hr = S_OK;

	status = key.Open(parent, str_key.c_str(), KEY_WRITE | KEY_READ);
	if (ERROR_SUCCESS == status)
	{
		DWORD dw_currentloadbehavior = -1;
		status = key.QueryDWORDValue(L"LoadBehavior", dw_currentloadbehavior);
		if (ERROR_SUCCESS == status && dw_currentloadbehavior != dw_value)
		{
			status = key.SetDWORDValue(L"LoadBehavior", dw_value);
			LOG_INFO << "Set Addin information Key=" << str_key << " LoadBehavior=" << dw_value;
			if (ERROR_SUCCESS != status)
			{
				LOG_ERROR << __FUNCTION__ << "-unable to save the value" << " key:" << str_key << " value:" << dw_value << " status:" << status;
				return HRESULT_FROM_WIN32(GetLastError());
			}
		}
		else
		{
			if (ERROR_SUCCESS != status)
			{
				LOG_ERROR << __FUNCTION__ << "-unable to read the value" << " key:" << str_key << " value:" << dw_value << " status:" << status;
				return HRESULT_FROM_WIN32(GetLastError());
			}
		}
	}
	else
	{
		LOG_ERROR << __FUNCTION__ << "-unable to write the value" << " key:" << str_key << " value:" << dw_value << " status:" << status;
		return HRESULT_FROM_WIN32(GetLastError());
	}
	return hr;
}
/// <summary>
/// Disable all Excel COM Add in using an elevated process & the registry
/// </summary>
/// <param name="temp_file_export">AddIn name & settings</param>
/// <returns></returns>
HRESULT ExcelCOMAddIn::DisableOfficeAddinAdmin(fs::path  temp_file_export)
{
	auto pathexecutable = Utility::GetDirectoryWithCurrentExecutable();

	std::wstring executable = pathexecutable.append(L"OfficeToolsAddInAdmin.exe").native();
	std::wstring param = temp_file_export.native();
	SHELLEXECUTEINFO ShExecInfo = { 0 };
	ShExecInfo.cbSize = sizeof(SHELLEXECUTEINFO);
	ShExecInfo.fMask = SEE_MASK_NOCLOSEPROCESS;
	ShExecInfo.hwnd = NULL;
	ShExecInfo.lpVerb = NULL;
	ShExecInfo.lpFile = executable.c_str();
	ShExecInfo.lpParameters = param.c_str();
	ShExecInfo.lpDirectory = NULL;
	ShExecInfo.nShow = SW_SHOW;
	ShExecInfo.hInstApp = NULL;
	DWORD dExitCode = 0;

	if (ShellExecuteEx(&ShExecInfo) == TRUE)
	{
		WaitForSingleObject(ShExecInfo.hProcess, INFINITE);
		GetExitCodeProcess(ShExecInfo.hProcess, &dExitCode);
	}
	else
	{
		LOG_INFO << "OfficeToolsAddInAdmin.exe do not start properly, GetLastError=" << GetLastError();
		return HRESULT_FROM_WIN32(GetLastError());
	}
	LOG_INFO << "OfficeToolsAddInAdmin.exe exit code=" << dExitCode;
	CloseHandle(ShExecInfo.hProcess);

	return dExitCode;
}
/// <summary>
/// Read one Excel COM Add in information using the registry
/// </summary>
/// <param name="parent">the HKEY parent HKEY_CURRENT_USER HKEY_LOCAL_MACHINE or HKEY_USERS</param>
/// <param name="siduserkey">if the information come from another profile</param>
/// <param name="rootKey">the key name</param>
void ExcelCOMAddIn::ReadAddInformation(std::map<std::wstring, AddinInformation>&addinsInfo, HKEY parent, std::wstring siduserkey, std::wstring str_key)
{
	LOG_DEBUG << __FUNCTION__;

	CRegKey key;
	CRegKey keyinformation;
	DWORD dwIndex = 0;
	DWORD cbName = IS_KEY_LEN;
	WCHAR szSubKeyName[IS_KEY_LEN] = { 0 };
	LONG lRet;

	LSTATUS status;

	status = key.Open(parent, str_key.c_str(), KEY_READ);
	if (ERROR_SUCCESS == status)
	{
		LOG_TRACE << __FUNCTION__ << "-the key exist" << " key:" << str_key;

		while ((lRet = key.EnumKey(dwIndex, szSubKeyName, &cbName)) != ERROR_NO_MORE_ITEMS)
		{

			AddinInformation info;
			DWORD dvalue;

			if (lRet == ERROR_SUCCESS)
			{
				info.Progid_ = szSubKeyName;
				info.Software_ = L"Excel";

				std::wstring infoPlugIn = str_key + L"\\" + info.Progid_;
				status = keyinformation.Open(parent, infoPlugIn.c_str(), KEY_READ);
				if (ERROR_SUCCESS == status)
				{
					WCHAR path_buff[MAX_PATH] = { 0 };
					ULONG len = sizeof(path_buff);

					if (keyinformation.QueryStringValue(L"Description", path_buff, &len) == ERROR_SUCCESS)
					{
						info.Description_ = path_buff;
					}
					len = sizeof(path_buff);
					if (keyinformation.QueryDWORDValue(L"LoadBehavior", dvalue) == ERROR_SUCCESS)
					{
						info.Startmode_ = dvalue;
					}

					len = sizeof(path_buff);
					std::wstring str_name;
					std::wstring str_domain;

					if (keyinformation.QueryStringValue(L"FriendlyName", path_buff, &len) == ERROR_SUCCESS)
					{
						info.FriendlyName_ = path_buff;
					}

					info.Key_ = infoPlugIn;
					info.Parent_ = parent;



					if (siduserkey.empty())
					{
						info.str_account_ = Utility::get_system_user_name();
					}
					else
					{
						info.str_account_ = Utility::GetUserInfo(siduserkey);
					}

					info.sid_account = Utility::GetSIDInfoFromUser(Utility::get_system_user_name());


					keyinformation.Close();
				}
			}
			dwIndex++;
			cbName = IS_KEY_LEN;


			if (addinsInfo.count(info.Progid_) == 0)
			{
				LOG_TRACE << __FUNCTION__ << "-add to collection" << " key:" << info.sid_account + L"_" + info.Progid_;
				addinsInfo.insert(std::make_pair(info.sid_account + L"_" + info.Progid_, info));
			}
			else
			{
				LOG_TRACE << __FUNCTION__ << "-already exist to the collection" << " key:" << info.Progid_;
				int x = 0;
			}
		}
		key.Close();
	}
	else
	{
		if (status != ERROR_FILE_NOT_FOUND)
		{
			if (status == ERROR_ACCESS_DENIED)
			{
				LOG_TRACE << __FUNCTION__ << "-access denied key=" << str_key;
			}
			else
			{
				LOG_TRACE << __FUNCTION__ << "-unable to open the key, key= " << str_key << " error=" << status;
			}
		}
	}
}
/// <summary>
/// Enumetate the list of user profiles from the registry
/// </summary>
/// <returns>the list of profile</returns>
std::list<std::wstring> ExcelCOMAddIn::EnumerateProfileNames()
{
	LOG_DEBUG << __FUNCTION__;

	CRegKey key;
	CRegKey keyinformation;
	DWORD dwIndex = 0;
	DWORD cbName = IS_KEY_LEN;
	WCHAR szSubKeyName[IS_KEY_LEN] = { 0 };
	LONG lRet;

	std::list<std::wstring> listusers;
	LSTATUS status;

	status = key.Open(HKEY_USERS, nullptr, KEY_READ);
	if (ERROR_SUCCESS == status)
	{
		while ((lRet = key.EnumKey(dwIndex, szSubKeyName, &cbName)) != ERROR_NO_MORE_ITEMS)
		{
			listusers.push_back(szSubKeyName);
			dwIndex++;
			cbName = IS_KEY_LEN;
		}
		key.Close();
	}
	else
	{
		LOG_ERROR << __FUNCTION__ << " unable to open the key:" << "HKEY_USERS" << " status:" << status;
	}
	return listusers;
}
