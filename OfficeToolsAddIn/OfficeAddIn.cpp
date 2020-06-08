#include "pch.h"
#include "OfficeAddIn.h"
#include "XLSingleton.h"
#include "OfficeAddinInformation.h"
#include "ExcelAutomation.h"
#include "Logger.h"


#define	IS_KEY_LEN 512

void OfficeAddIn::DisableAllOfficeAddIn()
{
	LOG_TRACE << __FUNCTION__;
	ProcessInformation processinformation = XLSingleton::getInstance()->Get_Addin_info();

	for (auto it = processinformation.addininformation_.begin(); it != processinformation.addininformation_.end(); it++)
	{
		if (it->second.addType_ == AddInType::OFFICE)
		{
			LOG_TRACE << "Disable Addin Name=" << it->second.Name_ << " Description=" << it->second.Description_;

			it->second.LoadBehavior_ = 0; //Unloaded-Do not load automatically
			WriteInformation(it->second.parent_, it->second.key_, it->second.LoadBehavior_);
		}
		if (it->second.addType_ == AddInType::XL)
		{
			it->second.Installed_ = L"False";
			LOG_TRACE << "Disable Addin Name=" << it->second.Name_ << " Description=" << it->second.Description_;
		}
	}
	XLSingleton::getInstance()->Set_Addin_info(processinformation);

	ExcelAutomation xl;
	xl.DisableAllAddin(processinformation);
}
void OfficeAddIn::DisableOfficeAddin(ProcessInformation processinformation)
{
	LOG_TRACE << __FUNCTION__;	
	ExcelAutomation xl;
	xl.DisableAddin(processinformation);
}

void OfficeAddIn::WriteInformation(HKEY parent, std::wstring str_key, DWORD dw_value)
{
	LOG_TRACE << __FUNCTION__;

	LSTATUS status;
	CRegKey key;

	status = key.Open(parent, str_key.c_str(), KEY_WRITE);
	if (ERROR_SUCCESS == status)
	{
		status = key.SetDWORDValue(L"LoadBehavior", dw_value);
		if (ERROR_SUCCESS != status)
		{
			LOG_ERROR << __FUNCTION__ << " unable to save the value" << " key:" << str_key << " value:" << dw_value << " status:" << status;
		}
	}
	else
	{
		LOG_ERROR << __FUNCTION__ << " unable to write the value" << " key:" << str_key << " value:" << dw_value << " status:" << status;
	}
}
std::list<std::wstring> EnumerateUserNames()
{
	LOG_TRACE << __FUNCTION__;

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

void OfficeAddIn::ReadAddinInformation()
{
	LOG_TRACE << __FUNCTION__;

	ProcessInformation processinformation = XLSingleton::getInstance()->Get_Addin_info();

	processinformation.addininformation_.clear();
	processinformation.modules_.clear();

	ExcelAutomation automation;
	automation.ListInformations(processinformation);

	auto users = EnumerateUserNames();

	for (auto user : users)
	{
		ReadAddInformation(HKEY_USERS,user, user + L"\\SOFTWARE\\Microsoft\\Office\\Excel\\Addins");
		ReadAddInformation(HKEY_USERS, user, user + L"\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Microsoft\\Office\\Excel\\Addins");
		ReadAddInformation(HKEY_USERS, user, user + L"\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Wow6432Node\\Microsoft\\Office\\Excel\\Addins");

	}
	ReadAddInformation(HKEY_CURRENT_USER,L"", L"SOFTWARE\\Microsoft\\Office\\Excel\\Addins");
	ReadAddInformation(HKEY_LOCAL_MACHINE, L"", L"SOFTWARE\\Microsoft\\Office\\Excel\\Addins");
	ReadAddInformation(HKEY_LOCAL_MACHINE, L"", L"SOFTWARE\\Microsoft\\Office\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Microsoft\\Office\\Excel\\Addins");
	ReadAddInformation(HKEY_LOCAL_MACHINE, L"", L"SOFTWARE\\Microsoft\\Office\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Wow6432Node\\Microsoft\\Office\\Excel\\Addins");
	
	ComputeAddInInformation(processinformation);
	XLSingleton::getInstance()->Set_Addin_info(processinformation);
}


std::wstring get_system_user_name()
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

HRESULT GetUserInfo(std::wstring sid_str, std::wstring& str_name, std::wstring& str_domain)
{
	SID_NAME_USE user_type;
	PSID sid = NULL;
	HRESULT ret = E_FAIL;
	if (ConvertStringSidToSid(sid_str.c_str(), &sid)) 
	{
		DWORD name_size = 0, domain_size = 0;
		if (!LookupAccountSid(NULL, sid, NULL, &name_size, NULL,&domain_size, &user_type) && ERROR_INSUFFICIENT_BUFFER != GetLastError()) 
		{
			LocalFree(sid);
			return(HRESULT_FROM_WIN32(GetLastError()));
		}
		wchar_t* c_name = new wchar_t[name_size];

		if (!c_name) 
		{
			LocalFree(sid);
			return E_OUTOFMEMORY;
		}

		wchar_t* c_domain = new wchar_t[domain_size];		
		if (!c_domain) 
		{
			delete[] c_name;
			LocalFree(sid);
			return E_OUTOFMEMORY;
		}

		if (LookupAccountSid(NULL, sid, c_name, &name_size, c_domain,&domain_size, &user_type)) 
		{
			ret = S_OK;
			str_name = c_name;
			str_domain = c_domain;
		}
		else 
		{
			return(HRESULT_FROM_WIN32(GetLastError()));
		}

		delete[] c_name;
		delete[] c_domain;
		LocalFree(sid);
	}
	return ret;
}

void OfficeAddIn::ReadAddInformation(HKEY parent,std::wstring siduserkey, std::wstring rootKey)
{
	LOG_TRACE << __FUNCTION__;

	CRegKey key;
	CRegKey keyinformation;
	DWORD dwIndex = 0;
	DWORD cbName = IS_KEY_LEN;
	WCHAR szSubKeyName[IS_KEY_LEN] = { 0 };
	LONG lRet;

	LSTATUS status;

	status = key.Open(parent, rootKey.c_str(), KEY_READ);
	if (ERROR_SUCCESS == status)
	{
		while ((lRet = key.EnumKey(dwIndex, szSubKeyName, &cbName)) != ERROR_NO_MORE_ITEMS)
		{

			AddinInformation info;
			DWORD dvalue;

			if (lRet == ERROR_SUCCESS)
			{
				info.Progid_ = szSubKeyName;
				info.Software_ = L"Excel";

				std::wstring infoPlugIn = rootKey + L"\\" + info.Progid_;
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
						info.str_account_ = get_system_user_name();
					}
					else
					{
						if (SUCCEEDED(GetUserInfo(siduserkey, str_name, str_domain)))
						{
							info.str_account_ = str_domain + L"\\" + str_name;
						}
					}
					

					keyinformation.Close();
				}
			}
			dwIndex++;
			cbName = IS_KEY_LEN;

			if (addinsInfo_.count(info.FriendlyName_) == 0)
			{
				addinsInfo_.insert(std::make_pair(info.Progid_, info));
			}
		}
		key.Close();
	}
	else
	{
		LOG_ERROR << __FUNCTION__ << " unable to open the key:" << rootKey << " status:" << status;
	}
}
void OfficeAddIn::ComputeAddInInformation(ProcessInformation& processinformation)
{
	LOG_TRACE << __FUNCTION__;

	std::map<std::wstring, XLaddinInformation>::iterator it;

	DumpAddIninfo();

	for (it = processinformation.addininformation_.begin(); it != processinformation.addininformation_.end(); it++)
	{
		LOG_TRACE << L"Description=" << it->second.Description_ << L" --ProgId=" << it->second.ProgId_ << L" --type" << (int)it->second.addType_;

		it->second.LoadBehavior_ = -1;

		if (it->second.addType_ == AddInType::OFFICE)
		{
			it->second.Installed_ = L"Not Applicable";
			if (!it->second.ProgId_.empty())
			{

				std::map<std::wstring, AddinInformation >::iterator i = addinsInfo_.find(it->second.ProgId_);
				if (i != addinsInfo_.end())
				{
					it->second.LoadBehavior_ = i->second.Startmode_;
					it->second.parent_ = i->second.Parent_;
					it->second.key_ = i->second.Key_;
					it->second.str_account = i->second.str_account_;
				}
				else
				{
					LOG_ERROR << "Unable to map ProgID=" << it->second.ProgId_;
				}
			}
			else
			{
				LOG_ERROR << "Unable to map Description is empty";
			}
		}
		else
		{
			it->second.LoadBehavior_ = 100;
			it->second.parent_ = nullptr;
			it->second.key_ = L"";					
			it->second.str_account = get_system_user_name();
			
		}
	}
}


void OfficeAddIn::SaveAddinInformation()
{
	LOG_TRACE << __FUNCTION__;
	ProcessInformation processinformation = XLSingleton::getInstance()->Get_Addin_info();
	for (auto it = processinformation.addininformation_.begin(); it != processinformation.addininformation_.end(); it++)
	{
		if (it->second.addType_ == AddInType::OFFICE)
		{
			WriteInformation(it->second.parent_, it->second.key_, it->second.LoadBehavior_);
		}
	}
	DisableOfficeAddin(processinformation);
}

void OfficeAddIn::DumpAddIninfo()
{
	for (auto addin : addinsInfo_)
	{
		LOG_TRACE << L"Description=" << addin.second.Description_ << L" --progID=" << addin.second.Progid_ << L" --FriendlyName=" << addin.second.FriendlyName_ << L" --type:" << L"Office";
	}
}
