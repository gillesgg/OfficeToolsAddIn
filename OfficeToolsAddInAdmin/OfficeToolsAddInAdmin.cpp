
// OfficeToolsAddInAdmin.cpp : Defines the class behaviors for the application.
//

#include "pch.h"
#include "framework.h"
#include "OfficeToolsAddInAdmin.h"
#include "OfficeToolsAddInAdminDlg.h"
#include "OfficeAddinInformation.h"
#include "logger.h"
#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// COfficeToolsAddInAdminApp

BEGIN_MESSAGE_MAP(COfficeToolsAddInAdminApp, CWinApp)
	ON_COMMAND(ID_HELP, &CWinApp::OnHelp)
END_MESSAGE_MAP()


// COfficeToolsAddInAdminApp construction

COfficeToolsAddInAdminApp::COfficeToolsAddInAdminApp()
{
	// support Restart Manager
	m_dwRestartManagerSupportFlags = AFX_RESTART_MANAGER_SUPPORT_RESTART;

	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}


// The one and only COfficeToolsAddInAdminApp object

COfficeToolsAddInAdminApp theApp;


// COfficeToolsAddInAdminApp initialization

BOOL COfficeToolsAddInAdminApp::InitInstance()
{
	// InitCommonControlsEx() is required on Windows XP if an application
	// manifest specifies use of ComCtl32.dll version 6 or later to enable
	// visual styles.  Otherwise, any window creation will fail.
	INITCOMMONCONTROLSEX InitCtrls;
	InitCtrls.dwSize = sizeof(InitCtrls);
	// Set this to include all the common control classes you want to use
	// in your application.
	InitCtrls.dwICC = ICC_WIN95_CLASSES;
	InitCommonControlsEx(&InitCtrls);

	CWinApp::InitInstance();


	AfxEnableControlContainer();

	// Create the shell manager, in case the dialog contains
	// any shell tree view or shell list view controls.
	CShellManager *pShellManager = new CShellManager;

	// Activate "Windows Native" visual manager for enabling themes in MFC controls
	CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerWindows));


	std::wstring str_filename = GetFileNameParameter();

	if (str_filename.empty() == true)
	{
		LOG_ERROR << "unable to read the plugIn setting file";
		return -1;
	}

	auto hr = SaveRegistrySettings(str_filename);
	LOG_DEBUG << "SaveRegistrySettings error count=" << " hr=" << hr;
		if (pShellManager != nullptr)
	{
		delete pShellManager;
	}

	#if !defined(_AFXDLL) && !defined(_AFX_NO_MFC_CONTROLS_IN_DIALOGS)
		ControlBarCleanUp();
	#endif

	return hr;
}

std::wstring COfficeToolsAddInAdminApp::GetFileNameParameter()
{
	LOG_DEBUG << __FUNCTION__;
	LPWSTR* szArglist = nullptr;
	std::wstring filename;
	int iNumArgs = 0;
	szArglist = CommandLineToArgvW(GetCommandLine(), &iNumArgs);

	filename = iNumArgs == 2 ? szArglist[1] : std::wstring();
	LocalFree(szArglist);
	return filename;
}

HRESULT COfficeToolsAddInAdminApp::SaveRegistrySettings(std::wstring str_filename)
{
	LOG_DEBUG << __FUNCTION__;
	LSTATUS status;
	HRESULT failure = S_OK;
	try
	{
		if (str_filename.empty() == false && fs::exists(fs::path(str_filename)))
		{
			pt::wptree root;
			pt::read_json(fs::path(str_filename).generic_string(), root);				

			for (auto addinfo : root.get_child(L"AddinInformation"))
			{
				auto str_addin_type = addinfo.second.get<std::wstring>(L"AddInType");
				auto str_parent = addinfo.second.get<std::wstring>(L"parent");
				auto SID = addinfo.second.get<std::wstring>(L"SID");
				auto strLoadBehavior = addinfo.second.get<std::wstring>(L"LoadBehavior");

				auto strkey = addinfo.second.get<std::wstring>(L"key");

				if (str_addin_type == L"Office" && str_parent == L"HKEY_LOCAL_MACHINE")
				{
					CRegKey key;
					CRegKey keyinformation;

					std::wstring struserkey = strkey;

					status = key.Open(HKEY_LOCAL_MACHINE, struserkey.c_str(), KEY_WRITE | KEY_READ);
					if (ERROR_SUCCESS == status)
					{
						DWORD dw_loadbehavior = std::stoi(strLoadBehavior);

						DWORD dw_currentloadbehavior = -1;

						status = key.QueryDWORDValue(L"LoadBehavior", dw_currentloadbehavior);
						if (ERROR_SUCCESS == status && dw_currentloadbehavior != dw_loadbehavior)
						{
							status = key.SetDWORDValue(L"LoadBehavior", dw_loadbehavior);
							if (ERROR_SUCCESS != status)
							{
								failure++;
								LOG_ERROR << __FUNCTION__ << "-unable to save the value" << " key:" << struserkey << " LoadBehavior value:" << dw_loadbehavior << " status:" << status;
							}
						}
						else
						{
							if (status != ERROR_SUCCESS)
							{
								failure++;
								LOG_ERROR << __FUNCTION__ << "-unable to read the LoadBehavior value, the key=" << struserkey;
							}
						}
					key.Close();
					}
					else
					{
						failure++;
						LOG_ERROR << __FUNCTION__ << "-unable to open the key=" << struserkey << " with KEY_WRITE";
					}
				}
			}
		}
	}
	catch (boost::exception const& ex)
	{
		LOG_ERROR << __FUNCTION__ << "Unable to read the json file, file=" << str_filename << "exception:" << diagnostic_information(ex);
		return 0XC00CE556; // return a parsing error
	}
	return S_OK;
}