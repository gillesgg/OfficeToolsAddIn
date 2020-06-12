#pragma once
#include "OfficeAddinInformation.h"

class ExcelCOMAddIn
{
public:
	HRESULT DisableAllAddin();
	HRESULT DisableOfficeAddinAdmin(fs::path  temp_file_export);
	void ReadAddInformation(std::map<std::wstring, AddinInformation>& addinsInfo, HKEY parent, std::wstring siduserkey, std::wstring str_key);
	std::list<std::wstring> EnumerateProfileNames();
	
private:
	HRESULT WriteLoadBehaviorRegistryInformation(HKEY parent, std::wstring str_key, DWORD dw_value);
};

