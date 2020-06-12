#include "pch.h"
#include "OfficeAddIn.h"
#include "XLSingleton.h"
#include "OfficeAddinInformation.h"
#include "ExcelAutomation.h"
#include "ExcelOfficeAddIn.h"
#include "Logger.h"
#include "Utility.h"

#pragma region DisableAllAdd-in
/// <summary>
/// Disable all COM Add-in and Excel Add-in
/// </summary>
/// <param name="temp_file_export"></param>
/// <returns></returns>
HRESULT OfficeAddIn::DisableAllOfficeAddIn(std::wstring temp_file_export)
{
	LOG_DEBUG << __FUNCTION__;

	ExcelAddIn			xlAddin;
	ExcelCOMAddIn		xlComAddin;
	HRESULT				hr = S_OK;

	DisableAllOfficeAddinToMemory();
	SaveAddinInformationToFile(temp_file_export);

	DumpAddIninfo();

	hr = xlAddin.DisableAllAddin();
	hr = xlComAddin.DisableAllAddin();
	hr = xlComAddin.DisableOfficeAddinAdmin(temp_file_export);
	
	return hr;
}
/// <summary>
/// Disable COM Add-in and Excel Add-in to the singleton
/// </summary>
void OfficeAddIn::DisableAllOfficeAddinToMemory()
{
	LOG_DEBUG << __FUNCTION__;

	ProcessInformation processinformation = XLSingleton::getInstance()->Get_Addin_info();
	for (auto it = processinformation.addininformation_.begin(); it != processinformation.addininformation_.end(); it++)
	{
		if (it->second.addType_ == AddInType::OFFICE)
		{
			LOG_INFO << "Disable Addin Name=" << it->second.sid_account + L"_" + it->second.ProgId_ << "Description=" << it->second.Description_ << " progID=" << it->second.ProgId_;
			it->second.LoadBehavior_ = 0; //Unloaded-Do not load automatically
		}
		if (it->second.addType_ == AddInType::XL)
		{
			it->second.Installed_ = L"False";
			LOG_INFO << "Disable Addin Name=" << it->second.sid_account + L"_" + it->second.ProgId_ << "Description=" << it->second.Description_ << " progID=" << it->second.ProgId_;
		}
	}
	XLSingleton::getInstance()->Set_Addin_info(processinformation);
}

/// <summary>
/// Save AddIn state to file for interprocess elevation scenario
/// </summary>
/// <param name="temp_file_export"></param>
void OfficeAddIn::SaveAddinInformationToFile(fs::path temp_file_export)
{
	pt::wptree root;

	ProcessInformation processinformation = XLSingleton::getInstance()->Get_Addin_info();

	try
	{
		root.put(L"ImageType", processinformation.imagetype_ == ImageType::x64 ? L"X64" : L"X86");
		root.put(L"Name", processinformation.Name_);
		pt::wptree children;
		for (auto& addininfo : processinformation.addininformation_)
		{
			pt::wptree child;
			child.put(L"Description", addininfo.second.Description_);
			child.put(L"Installed", addininfo.second.Installed_);
			child.put(L"key", addininfo.second.key_);
			if (addininfo.second.parent_ == HKEY_LOCAL_MACHINE)
			{
				child.put(L"parent", L"HKEY_LOCAL_MACHINE");
			}
			else if (addininfo.second.parent_ == HKEY_CURRENT_USER)
			{
				child.put(L"parent", L"HKEY_CURRENT_USER");
			}
			else
			{
				child.put(L"parent", L"HKEY_USERS");
			}

			child.put(L"AddInType", addininfo.second.addType_ == AddInType::OFFICE ? L"Office" : L"Excel");
			child.put(L"LoadBehavior", std::to_wstring(addininfo.second.LoadBehavior_));
			child.put(L"ProgId", addininfo.second.ProgId_);

			child.put(L"SID", addininfo.second.sid_account);
			child.put(L"UserName", addininfo.second.str_account);
			children.push_back(std::make_pair(L"", child));
		}
		root.add_child(L"AddinInformation", children);
		pt::write_json(temp_file_export.generic_string(), root);
	}
	catch (boost::exception const& ex)
	{
		LOG_ERROR << __FUNCTION__ << "-Unable to write the json file, file=" << temp_file_export << " exception=" << diagnostic_information(ex);
	}

	//LOG_DEBUG << __FUNCTION__;
	//ProcessInformation processinformation = XLSingleton::getInstance()->Get_Addin_info();
	//for (auto it = processinformation.addininformation_.begin(); it != processinformation.addininformation_.end(); it++)
	//{
	//	if (it->second.addType_ == AddInType::OFFICE)
	//	{
	//		if ((it->second.parent_ != HKEY_LOCAL_MACHINE))
	//		{
	//			//WriteLoadBehaviorRegistryInformation(it->second.parent_, it->second.key_, it->second.LoadBehavior_);
	//		}
	//		else
	//		{
	//			LOG_TRACE << __FUNCTION__ << " Admin key " << " Key=" << it->second.key_ << " LoadBehavior=" << it->second.LoadBehavior_;
	//		}
	//	}
	//}
}

#pragma endregion
//void OfficeAddIn::DisableXLAddin(ProcessInformation processinformation)
//{
//	LOG_DEBUG << __FUNCTION__;	
//	ExcelAddIn xl;
//	xl.DisableAddin(processinformation);
//}
#pragma region Get Add-in information
/// <summary>
/// Get Office AddinInformation from the XL object model and from the registy
/// </summary>
/// <returns></returns>
HRESULT OfficeAddIn::ReadAddinInformation()
{
	LOG_DEBUG << __FUNCTION__;
	HRESULT hr = S_OK;

	ProcessInformation processinformation = XLSingleton::getInstance()->Get_Addin_info();
	processinformation.addininformation_.clear();
	processinformation.modules_.clear();

	ExcelAddIn	xl;
	ExcelCOMAddIn xlComAddin;

	if FAILED(hr = xl.ListInformations(processinformation))
	{
		return hr;
	}
	
	auto sid_users = xlComAddin.EnumerateProfileNames();

	LOG_TRACE << __FUNCTION__ << "-read Addins information for the current user" << ", user name =" << Utility::get_system_user_name() << " sid name=" << Utility::GetSIDInfoFromUser(Utility::get_system_user_name());

	xlComAddin.ReadAddInformation(addinsInfo_, HKEY_CURRENT_USER,L"", L"SOFTWARE\\Microsoft\\Office\\Excel\\Addins");
	xlComAddin.ReadAddInformation(addinsInfo_, HKEY_LOCAL_MACHINE, L"", L"SOFTWARE\\Microsoft\\Office\\Excel\\Addins");
	xlComAddin.ReadAddInformation(addinsInfo_, HKEY_LOCAL_MACHINE, L"", L"SOFTWARE\\Microsoft\\Office\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Microsoft\\Office\\Excel\\Addins");
	xlComAddin.ReadAddInformation(addinsInfo_, HKEY_LOCAL_MACHINE, L"", L"SOFTWARE\\Microsoft\\Office\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Wow6432Node\\Microsoft\\Office\\Excel\\Addins");

	for (auto sid_user : sid_users)
	{

		LOG_TRACE << __FUNCTION__ << "-read Addins information for the user" << ", user name =" << Utility::GetUserInfo(sid_user) << ", sid name=" << sid_user;
		xlComAddin.ReadAddInformation(addinsInfo_, HKEY_USERS, sid_user, sid_user + L"\\SOFTWARE\\Microsoft\\Office\\Excel\\Addins");
		xlComAddin.ReadAddInformation(addinsInfo_, HKEY_USERS, sid_user, sid_user + L"\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Microsoft\\Office\\Excel\\Addins");
		xlComAddin.ReadAddInformation(addinsInfo_, HKEY_USERS, sid_user, sid_user + L"\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\REGISTRY\\MACHINE\\Software\\Wow6432Node\\Microsoft\\Office\\Excel\\Addins");
	}
	ComputeAddInInformation(processinformation);
	XLSingleton::getInstance()->Set_Addin_info(processinformation);

	return (hr);
}
/// <summary>
/// Merge Office AddinInformation - XL Object model + COM registry information
/// </summary>
/// <param name="processinformation"></param>
void OfficeAddIn::ComputeAddInInformation(ProcessInformation& processinformation)
{
	LOG_DEBUG << __FUNCTION__;

	std::map<std::wstring, XLaddinInformation>::iterator it;


	for (it = processinformation.addininformation_.begin(); it != processinformation.addininformation_.end(); it++)
	{
		LOG_TRACE << L"Description=" << it->second.Description_ << L" --ProgId=" << it->second.ProgId_ << L" --type" << (int)it->second.addType_;

		it->second.LoadBehavior_ = -1;

		if (it->second.addType_ == AddInType::OFFICE)
		{
			it->second.Installed_ = L"Not Applicable";
			if (!it->second.ProgId_.empty())
			{

				std::map<std::wstring, AddinInformation >::iterator i = addinsInfo_.find(it->first);
				if (i != addinsInfo_.end())
				{
					it->second.LoadBehavior_ = i->second.Startmode_;
					it->second.parent_ = i->second.Parent_;
					it->second.key_ = i->second.Key_;
					it->second.str_account = i->second.str_account_;
					it->second.sid_account = i->second.sid_account;
				}
				else
				{
					LOG_DEBUG << "Unable to map ProgID=" << it->second.ProgId_;
				}
			}
			else
			{
				LOG_DEBUG << "Unable to map Description is empty";
			}
		}
		else
		{
			it->second.LoadBehavior_ = 100;
			it->second.parent_ = nullptr;
			it->second.key_ = L"";
		}
	}
}
#pragma endregion



HRESULT OfficeAddIn::DisableCurrentOfficeAddIn(std::wstring temp_file_export)
{
	LOG_DEBUG << __FUNCTION__;

	ExcelAddIn			xlAddin;
	ExcelCOMAddIn		xlComAddin;
	HRESULT				hr = S_OK;

	DumpAddIninfo();

	SaveAddinInformationToFile(temp_file_export);

	

	hr = xlAddin.DisableAddin();
	hr = xlComAddin.DisableAllAddin();
	hr = xlComAddin.DisableOfficeAddinAdmin(temp_file_export);

	return (hr);
}

//std::wstring OfficeAddIn::makeKey(std::wstring siduserkey, std::wstring Progid)
//{
//	std::wstring username;
//
//	if (siduserkey.empty() == true)
//	{
//		username = Utility::get_system_user_name();
//
//	}
//	else
//	{
//		std::wstring str_name;
//		std::wstring str_domain;
//		username = Utility::GetUserInfo(siduserkey);
//	}
//	return username + L"\\" + Progid;
//}

void OfficeAddIn::DumpAddIninfo()
{
	ProcessInformation processinformation = XLSingleton::getInstance()->Get_Addin_info();
	for (auto item : processinformation.addininformation_)
	{
		LOG_TRACE	<< "Key=" 
					<< item.first 
					<< " --AddinType=" 
					<< (int) item.second.addType_ 
					<< " --Description=" 
					<< item.second.Description_ 
					<< " --FullName=" 
					<< item.second.FullName_ 
					<< " --key" 
					<< item.second.key_ 
					<< " --Name" 
					<< item.second.Name_ 
					<< " --sid=" 
					<< item.second.sid_account;
	}
}

