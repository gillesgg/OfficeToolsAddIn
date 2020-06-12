#pragma once

#include "OfficeAddinInformation.h"

class ExcelAddIn
{
public:
	HRESULT ListInformations(ProcessInformation& processinformation);
	HRESULT DisableAllAddin();
	HRESULT DisableAddin();
	HRESULT DisableXLAddinInformation();

private:
	HRESULT ReadXLAddinInformation(ProcessInformation& processinformation);
	HRESULT ReadOfficeAddinInformation(ProcessInformation& processinformation);
	HRESULT DisableAllXLAddinInformation();


private:
	Excel::_ApplicationPtr _pXL;
};

