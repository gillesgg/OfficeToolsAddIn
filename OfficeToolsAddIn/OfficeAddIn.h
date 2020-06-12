#pragma once
#include "OfficeAddinInformation.h"

class OfficeAddIn
{
public:
	HRESULT		DisableAllOfficeAddIn(std::wstring temp_file_export);
	HRESULT		ReadAddinInformation();
	void		SaveAddinInformationToFile(fs::path temp_file_export);
	fs::path	getDirectoryWithCurrentExecutable();
	void		DisableAllOfficeAddinToMemory();
	HRESULT		DisableCurrentOfficeAddIn(std::wstring temp_file_export);
private:
	void DisableXLAddin(ProcessInformation processinformation);
	void ReadAddInformation(HKEY parent, std::wstring siduserkey, std::wstring rootKey);
	void ComputeAddInInformation(ProcessInformation& processinformation);
	void DumpAddIninfo();
	void WriteLoadBehaviorRegistryInformation(HKEY parent, std::wstring str_key, DWORD dw_value);
	std::wstring makeKey(std::wstring siduserkey, std::wstring Progid);
	void DisableOfficeAddinAdmin(fs::path  temp_file_export);

private:
	std::map<std::wstring, AddinInformation>		addinsInfo_;
};

