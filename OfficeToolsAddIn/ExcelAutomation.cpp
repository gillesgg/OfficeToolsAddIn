#include "pch.h"
#include "ExcelAutomation.h"
#include "OfficeAddinInformation.h"
#include "ExcelProcessInformation.h"
#include "Logger.h"

HRESULT ExcelAutomation::DisableAllXLAddinInformation(ProcessInformation processinformation)
{
	LOG_TRACE << __FUNCTION__;
	HRESULT hr = S_OK;
	try
	{
		auto pOfficeAddIns = _pXL->GetAddIns();
		if (pOfficeAddIns == nullptr)
		{
			LOG_ERROR << __FUNCTION__ << " pOfficeAddIns is null";
			return E_FAIL;
		}
	
		for (int x = 1; x <= pOfficeAddIns->Count; x++)
		{
			VARIANT v;
			VariantInit(&v);
			v.vt = VT_I4;
			v.iVal = x;

			auto pOfficeAddin = pOfficeAddIns->Item[x];		
			if (pOfficeAddin != nullptr)
				pOfficeAddin->Installed = VARIANT_FALSE;			
			else
				LOG_ERROR << __FUNCTION__ << " pOfficeAddIn is null";
			pOfficeAddin = nullptr;
			VariantClear(&v);
		}
		pOfficeAddIns = nullptr;
	}
	catch (_com_error& ex)
	{
		LOG_ERROR << "error=" << ex.Description();
		return ex.Error();
	}
	return (hr);
}

HRESULT ExcelAutomation::DisableAddin(ProcessInformation processinformation)
{
	LOG_TRACE << __FUNCTION__;

	HRESULT hr = S_OK;
	
	try
	{
		CoInitialize(0L);
		hr = _pXL.CreateInstance(L"Excel.Application");
		if (FAILED(hr))
		{
			LOG_ERROR << __FUNCTION__ << " unable to create an Excel.Application instance" << " hr=" << hr;
			return hr;
		}
		_pXL->PutVisible(0, VARIANT_FALSE);
		hr = DisableXLAddinInformation(processinformation);
		_pXL->Quit();
		_pXL = nullptr;
	}
	catch (_com_error& e)
	{
		LOG_ERROR << "error=" << e.Description();
		return e.Error();
	}
	return (hr);
}


HRESULT ExcelAutomation::DisableXLAddinInformation(ProcessInformation processinformation)
{
	LOG_TRACE << __FUNCTION__;

	try
	{
		auto pOfficeAddIns = _pXL->GetAddIns();
		if (pOfficeAddIns == nullptr)
		{
			LOG_ERROR << __FUNCTION__ << " pOfficeAddIns is null";
			return E_FAIL;
		}

		for (int x = 1; x <= pOfficeAddIns->Count; x++)
		{
			VARIANT v;
			VariantInit(&v);
			v.vt = VT_I4;
			v.iVal = x;

			auto pOfficeAddin = pOfficeAddIns->Item[x];
			if (pOfficeAddin != nullptr)
			{
				auto bstr_clsid = pOfficeAddin->CLSID;
				auto bstr_name = pOfficeAddin->Name;
				std::wstring str_clsid = bstr_clsid.length() == 0 ? bstr_name : bstr_clsid;

				auto it = processinformation.addininformation_.find(str_clsid);
				if (it != processinformation.addininformation_.end())
				{
					pOfficeAddin->Installed = it->second.Installed_ == L"True" ? VARIANT_TRUE : VARIANT_FALSE;
				}
			}
			else
			{
				LOG_ERROR << __FUNCTION__ << " pOfficeAddIn is null";
			}
			pOfficeAddin = nullptr;
			VariantClear(&v);
		}
		pOfficeAddIns = nullptr;
	}
	catch (_com_error& e)
	{
		LOG_ERROR << "error=" << e.Description();
		return e.Error();
	}
	return (S_OK);
}


HRESULT ExcelAutomation::DisableAllAddin(ProcessInformation processinformation)
{
	LOG_TRACE << __FUNCTION__;
	HRESULT hr = S_OK;

	try
	{
		CoInitialize(0L);
		hr = _pXL.CreateInstance(L"Excel.Application");
		if (FAILED(hr))
		{
			LOG_ERROR << __FUNCTION__ << " unable to create an Excel.Application instance" << " hr=" << hr;
			return hr;
		}
		_pXL->PutVisible(0, VARIANT_FALSE);
		hr = DisableAllXLAddinInformation(processinformation);
		_pXL->Quit();
		_pXL = nullptr;
	}
	catch (_com_error& e)
	{
		LOG_ERROR << "error=" << e.Description();
		return e.Error();
	}
	return (hr);
}
HRESULT ExcelAutomation::ReadXLAddinInformation(ProcessInformation& processinformation)
{
	LOG_TRACE << __FUNCTION__;

	try
	{
		auto pOfficeAddIns = _pXL->GetAddIns();
		if (pOfficeAddIns == nullptr)
		{
				LOG_ERROR << __FUNCTION__ << " pOfficeAddIns is null";
				return E_FAIL;
		}

		for (int x = 1; x <= pOfficeAddIns->Count; x++)
		{
			VARIANT v;
			VariantInit(&v);
			v.vt = VT_I4;
			v.iVal = x;

			auto pOfficeAddin = pOfficeAddIns->Item[x];
			if (pOfficeAddin == nullptr)
			{
				LOG_ERROR << __FUNCTION__ << " pOfficeAddIn is null";
				return E_FAIL;
			}
			XLaddinInformation addinInformation;
		

			addinInformation.ProgId_ = pOfficeAddin->CLSID;
			addinInformation.FullName_ = pOfficeAddin->FullName;
			addinInformation.Name_ = pOfficeAddin->Name;
			addinInformation.Description_ = pOfficeAddin->Title;
			addinInformation.Installed_ = (pOfficeAddin->Installed == VARIANT_TRUE) ? L"True" : L"False";
			addinInformation.addType_ = AddInType::XL;		

			if (addinInformation.ProgId_.empty() == true)
			{
				addinInformation.ProgId_ = pOfficeAddin->Name;
			}

			processinformation.addininformation_.insert(std::make_pair(addinInformation.ProgId_, addinInformation));
			LOG_TRACE << "description=" << addinInformation.Description_ << " progID=" << addinInformation.ProgId_;
			pOfficeAddin = nullptr;
			VariantClear(&v);
		}
		pOfficeAddIns = nullptr;
	}
	catch (_com_error& e)
	{
		LOG_ERROR << "error=" << e.Description();
		return e.Error();
	}
	return (S_OK);
}
HRESULT ExcelAutomation::ReadOfficeAddinInformation(ProcessInformation& processinformation)
{
	try
	{
		LOG_TRACE << __FUNCTION__;
		auto pXLAddIns = _pXL->GetCOMAddIns();
		if (pXLAddIns == nullptr)
		{
			LOG_ERROR << __FUNCTION__ << " pOfficeAddIns is null";
			return E_FAIL;
		}
		for (int x = 1; x <= pXLAddIns->Count; x++)
		{
			VARIANT v;
			VariantInit(&v);
			v.vt = VT_I4;
			v.iVal = x;

			auto pXLAddin = pXLAddIns->Item(&v);
			if (pXLAddin == nullptr)
			{
				LOG_ERROR << __FUNCTION__ << " pOfficeAddIn is null";
				return E_FAIL;
			}

			XLaddinInformation addinInformation;

			addinInformation.Description_ = pXLAddin->Description;			
			addinInformation.ProgId_ = pXLAddin->ProgId;			
			addinInformation.Installed_ = (pXLAddin->Connect == VARIANT_TRUE) ? L"True" : L"False";
			addinInformation.addType_ = AddInType::OFFICE;		
			processinformation.addininformation_.insert(std::make_pair(addinInformation.ProgId_, addinInformation));
			LOG_TRACE << "description=" << addinInformation.Description_ << " progID=" << addinInformation.ProgId_;
			VariantClear(&v);
			pXLAddin = nullptr;
		}
		pXLAddIns = nullptr;
	}
	catch (_com_error& e)
	{
		LOG_ERROR << "error=" << e.Error();
		return e.Error();
	}
	return (S_OK);
}
HRESULT ExcelAutomation::ListInformations(ProcessInformation& processinformation)
{
	HRESULT hr =S_OK;

	CoInitialize(0L);

	ExcelProcessInformation process;

	Office::COMAddInsPtr pXLAddIns;

	try
	{
		hr = _pXL.CreateInstance(L"Excel.Application");
		if (FAILED(hr))
		{
			LOG_ERROR << __FUNCTION__ << " unable to create an Excel.Application instance" << " hr=" << hr;
			return hr;
		}
		_pXL->PutVisible(0, VARIANT_FALSE);
		ReadOfficeAddinInformation(processinformation);
		ReadXLAddinInformation(processinformation);

		process.ShowProcesses(L"excel.exe", processinformation);

		_pXL->Quit();
		_pXL = nullptr;
	}
	catch (_com_error& e)
	{
		LOG_ERROR << "error=" << e.Description();
		return e.Error();
	}
	return (hr);
}