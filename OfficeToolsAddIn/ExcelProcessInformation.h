#pragma once
#include "OfficeAddinInformation.h"

class ExcelProcessInformation
{
public:
	HRESULT ShowProcesses(std::wstring filter, ProcessInformation& processinformation);
private:
	int ShowModules(DWORD processID, ProcessInformation& processinformation);
};	