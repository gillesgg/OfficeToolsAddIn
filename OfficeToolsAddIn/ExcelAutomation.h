#pragma once

#include "OfficeAddinInformation.h"

class ExcelAutomation
{
public:
	HRESULT ListInformations(ProcessInformation& processinformation);
	HRESULT DisableAllAddin(ProcessInformation processinformation);
	HRESULT DisableAddin(ProcessInformation processinformation);

private:
	HRESULT ReadXLAddinInformation(ProcessInformation& processinformation);
	HRESULT ReadOfficeAddinInformation(ProcessInformation& processinformation);
	HRESULT DisableAllXLAddinInformation(ProcessInformation processinformation);
	HRESULT DisableXLAddinInformation(ProcessInformation processinformation);

private:
	Excel::_ApplicationPtr _pXL;
};

