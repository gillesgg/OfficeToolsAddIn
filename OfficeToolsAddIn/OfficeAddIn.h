#pragma once
#include "OfficeAddinInformation.h"

class OfficeAddIn
{
public:
	void DisableAllOfficeAddIn();
	void ReadAddinInformation();
	void SaveAddinInformation();
private:
	void DisableOfficeAddin(ProcessInformation processinformation);
	void ReadAddInformation(HKEY parent, std::wstring rootKey);
	void ComputeAddInInformation(ProcessInformation& processinformation);
	void DumpAddIninfo();
	void WriteInformation(HKEY parent, std::wstring str_key, DWORD dw_value);
private:
	std::map<std::wstring, AddinInformation>		addinsInfo_;
};

