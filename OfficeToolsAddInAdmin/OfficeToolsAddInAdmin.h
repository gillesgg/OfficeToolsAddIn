
// OfficeToolsAddInAdmin.h : main header file for the PROJECT_NAME application
//

#pragma once

#ifndef __AFXWIN_H__
	#error "include 'pch.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols


// COfficeToolsAddInAdminApp:
// See OfficeToolsAddInAdmin.cpp for the implementation of this class
//

class COfficeToolsAddInAdminApp : public CWinApp
{
private:
	std::wstring GetFileNameParameter();
	HRESULT SaveRegistrySettings(std::wstring str_filename);
public:
	COfficeToolsAddInAdminApp();

// Overrides
public:
	virtual BOOL InitInstance();

// Implementation

	DECLARE_MESSAGE_MAP()
};

extern COfficeToolsAddInAdminApp theApp;
