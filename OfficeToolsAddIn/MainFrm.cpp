// MainFrm.cpp : implementation of the CMainFrame class
//

#include "pch.h"
#include "framework.h"
#include "OfficeToolsAddIn.h"
#include "ExcelAutomation.h"
#include "MainFrm.h"
#include "OfficeToolsAddInDoc.h"
#include "OfficeToolsAddInView.h"
#include "Logger.h"
#include "XLSingleton.h"
#include "OfficeAddIn.h"
#include "Utility.h"
XLSingleton* XLSingleton::instance = 0;

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNCREATE(CMainFrame, CFrameWndEx)

BEGIN_MESSAGE_MAP(CMainFrame, CFrameWndEx)
	ON_WM_CREATE()
	ON_WM_SETTINGCHANGE()
	ON_COMMAND(ID_BUTTON_REFRESH, &CMainFrame::OnButtonRefresh)
	ON_COMMAND(ID_BUTTON_DISABLE, &CMainFrame::OnButtonDisable)
	ON_COMMAND(ID_BUTTON_EXPORT, &CMainFrame::OnButtonExport)
	ON_COMMAND(ID_BUTTON_IMPORT, &CMainFrame::OnButtonImport)
	ON_COMMAND(ID_BUTTON_EXPORTLOGS, &CMainFrame::OnButtonExportLogs)
	ON_COMMAND(ID_BUTTON_SAVE_CHANGE, &CMainFrame::OnButtonSaveChange)
END_MESSAGE_MAP()

CMainFrame::CMainFrame() noexcept
{}

CMainFrame::~CMainFrame()
{
	if (fs::exists(temp_file_export_))
	{
		try
		{
			fs::remove(temp_file_export_);
		}
		catch (std::filesystem::filesystem_error const& ex)
		{
			LOG_ERROR << __FUNCTION__ << "Unable to remove the file, file=" << temp_file_export_ << "exception:" << ex.what();
		}
	}
	::ExitProcess(0);
}


void CMainFrame::OnUpdateFrameTitle(BOOL Nada)
{
	CString csAppName;
	csAppName.Format(AFX_IDS_APP_TITLE);
	SetWindowText(csAppName);
}

int CMainFrame::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	LOG_DEBUG << __FUNCTION__;

	if (CFrameWndEx::OnCreate(lpCreateStruct) == -1)
		return -1;

	//BOOL bNameValid;
	

	m_wndRibbonBar.Create(this);
	m_wndRibbonBar.LoadFromResource(IDR_RIBBON);

	//if (!m_wndStatusBar.Create(this))
	//{
	//	LOG_ERROR << __FUNCTION__ << " Failed to create status bar";
	//	return -1;
	//}

	//CString strTitlePane1;
	//CString strTitlePane2;
	//bNameValid = strTitlePane1.LoadString(IDS_STATUS_PANE1);
	//ASSERT(bNameValid);
	//bNameValid = strTitlePane2.LoadString(IDS_STATUS_PANE2);
	//ASSERT(bNameValid);
	//m_wndStatusBar.AddElement(new CMFCRibbonStatusBarPane(ID_STATUSBAR_PANE1, strTitlePane1, TRUE), strTitlePane1);
	//m_wndStatusBar.AddExtendedElement(new CMFCRibbonStatusBarPane(ID_STATUSBAR_PANE2, strTitlePane2, TRUE), strTitlePane2);

	CDockingManager::SetDockingMode(DT_SMART);
	EnableAutoHidePanes(CBRS_ALIGN_ANY);

	if (!CreateDockingWindows())
	{
		LOG_ERROR << __FUNCTION__ << " Failed to create docking windows";
		return -1;
	}

	m_wndOutput.EnableDocking(CBRS_ALIGN_ANY);
	DockPane(&m_wndOutput);

	CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerWindows7));
	m_wndRibbonBar.SetWindows7Look(TRUE);

	m_wndRibbonButton.SetVisible(FALSE);
	CSize c(-10, -10);

	m_wndRibbonBar.SetApplicationButton(&m_wndRibbonButton, c);
	
	return 0;
}

BOOL CMainFrame::PreCreateWindow(CREATESTRUCT& cs)
{
	if( !CFrameWndEx::PreCreateWindow(cs) )
		return FALSE;
	return TRUE;
}

BOOL CMainFrame::CreateDockingWindows()
{
	BOOL		bNameValid;
	CString		strOutputWnd;

	bNameValid = strOutputWnd.LoadString(IDS_OUTPUT_WND);
	ASSERT(bNameValid);
	if (!m_wndOutput.Create(strOutputWnd, this, CRect(0, 0, 100, 100), TRUE, ID_VIEW_OUTPUTWND, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | CBRS_BOTTOM | CBRS_FLOAT_MULTI))
	{
		LOG_ERROR << __FUNCTION__ << " Failed to create Output window";
		return FALSE; // failed to create
	}
	SetDockingWindowIcons(theApp.m_bHiColorIcons);
	return TRUE;
}

void CMainFrame::SetDockingWindowIcons(BOOL bHiColorIcons)
{
	HICON hOutputBarIcon = (HICON) ::LoadImage(::AfxGetResourceHandle(), MAKEINTRESOURCE(bHiColorIcons ? IDI_OUTPUT_WND_HC : IDI_OUTPUT_WND), IMAGE_ICON, ::GetSystemMetrics(SM_CXSMICON), ::GetSystemMetrics(SM_CYSMICON), 0);
	m_wndOutput.SetIcon(hOutputBarIcon, FALSE);

}

#ifdef _DEBUG
void CMainFrame::AssertValid() const
{
	CFrameWndEx::AssertValid();
}

void CMainFrame::Dump(CDumpContext& dc) const
{
	CFrameWndEx::Dump(dc);
}
#endif //_DEBUG


// CMainFrame message handlers


void CMainFrame::OnSettingChange(UINT uFlags, LPCTSTR lpszSection)
{
	CFrameWndEx::OnSettingChange(uFlags, lpszSection);
	m_wndOutput.UpdateFonts();
}

void CMainFrame::OnButtonRefresh()
{
	LOG_DEBUG << __FUNCTION__;

	OfficeAddIn office_add_in;

	BeginWaitCursor();
	if (SUCCEEDED(office_add_in.ReadAddinInformation()))
	{
		if (nullptr != GetActiveView())
		{
			(dynamic_cast<COfficeToolsAddInView*>(GetActiveView()))->ShowAddIns();
		}
	}
	EndWaitCursor();

}

void CMainFrame::OnButtonDisable()
{
	LOG_DEBUG << __FUNCTION__;
	
	BeginWaitCursor();

	OfficeAddIn office_add_in;

	temp_file_export_ = CreateUniqueFile();
	if (SUCCEEDED(office_add_in.DisableAllOfficeAddIn(temp_file_export_)))
	{
		if (nullptr != GetActiveView())
		{
			(dynamic_cast<COfficeToolsAddInView*>(GetActiveView()))->ShowAddIns();
		}
	}	
	Utility::DeleteFile(temp_file_export_);
	EndWaitCursor();
}

void CMainFrame::OnButtonSaveChange()
{
	LOG_DEBUG << __FUNCTION__;

	BeginWaitCursor();

	OfficeAddIn office_add_in;

	temp_file_export_ = CreateUniqueFile();
	if (SUCCEEDED(office_add_in.DisableCurrentOfficeAddIn(temp_file_export_)))
	{
		if (nullptr != GetActiveView())
		{
			(dynamic_cast<COfficeToolsAddInView*>(GetActiveView()))->ShowAddIns();
		}
	}
	Utility::DeleteFile(temp_file_export_);
	EndWaitCursor();
}


fs::path CMainFrame::SaveFile()
{
	LOG_DEBUG << __FUNCTION__;
	char strFilter[] = { "JSON file  (*.json)|*.json|" };
	fs::path p_filename_path;

	CFileDialog FileDlg(FALSE, CString(".json"), NULL, 0, CString(strFilter));
	if (FileDlg.DoModal() == IDOK) // this is the line which gives the errors
	{
		p_filename_path = FileDlg.GetFolderPath().GetBuffer();
		p_filename_path.append(FileDlg.GetFileName().GetBuffer());

		if (fs::exists(p_filename_path))
		{
			try
			{
				fs::remove(p_filename_path);
			}
			catch (std::filesystem::filesystem_error const& ex)
			{
				LOG_ERROR << __FUNCTION__ << "-Unable to remove the file, file=" << p_filename_path << " exception:" << ex.what();
				return fs::path();
			}
		}
	}
	return p_filename_path;
}

fs::path CMainFrame::SaveFileTxt()
{
	LOG_DEBUG << __FUNCTION__;
	char strFilter[] = { "log file  (*.log)|*.log|" };
	fs::path p_filename_path;

	CFileDialog FileDlg(FALSE, CString(".log"), NULL, 0, CString(strFilter));
	if (FileDlg.DoModal() == IDOK) // this is the line which gives the errors
	{
		p_filename_path = FileDlg.GetFolderPath().GetBuffer();
		p_filename_path.append(FileDlg.GetFileName().GetBuffer());

		if (fs::exists(p_filename_path))
		{
			try
			{
				fs::remove(p_filename_path);
			}
			catch (std::filesystem::filesystem_error const& ex)
			{
				LOG_ERROR << __FUNCTION__ << "-Unable to remove the file, file=" << p_filename_path << " exception:" << ex.what();
				return fs::path();
			}
		}
	}
	return p_filename_path;
}

fs::path CMainFrame::OpenFile()
{
	LOG_DEBUG << __FUNCTION__;
	char strFilter[] = { "JSON file  (*.json)|*.json|" };
	fs::path p_filename_path;

	CFileDialog FileDlg(TRUE, CString(".json"), NULL, 0, CString(strFilter));
	if (FileDlg.DoModal() == IDOK) // this is the line which gives the errors
	{
		p_filename_path = FileDlg.GetFolderPath().GetBuffer();
		p_filename_path.append(FileDlg.GetFileName().GetBuffer());
	}
	return p_filename_path;
}


fs::path CMainFrame::GenerateTemporyFile()
{
	wchar_t wUniqueFileName[MAX_PATH + 1];
	_wtmpnam_s(wUniqueFileName);

	if (wUniqueFileName != nullptr)
	{
		fs::path temp_filename = wUniqueFileName;
		return temp_filename;
	}
	return fs::path();
}

std::wstring CMainFrame::CreateUniqueFile()
{
	LOG_DEBUG << __FUNCTION__;

	auto file_export = GenerateTemporyFile();
	if (file_export.empty() == false)
	{
		return file_export;
	}
	else
	{
		LOG_ERROR << __FUNCTION__ << "-Unable to create the tempory file";
	}
	return std::wstring();
}
void CMainFrame::OnButtonExport()
{
	LOG_DEBUG << __FUNCTION__;
	fs::path p_filenamepath;

	OfficeAddIn officeAddIn;

		p_filenamepath = SaveFile();
		if (p_filenamepath.empty() == false)
		{
			officeAddIn.SaveAddinInformationToFile(p_filenamepath);
		}
}
void CMainFrame::OnButtonImport()
{
	LOG_DEBUG << __FUNCTION__;
	fs::path p_filenamepath;

	try
	{
		p_filenamepath = OpenFile();
		if (p_filenamepath.empty() == false)
		{
			pt::wptree root;
			pt::read_json(p_filenamepath.generic_string(), root);
			ProcessInformation processinformation = XLSingleton::getInstance()->Get_Addin_info();

			for (auto addinfo : root.get_child(L"AddinInformation"))
			{
				auto str_addin_load_behavior = addinfo.second.get<std::wstring>(L"LoadBehavior");
				auto str_addin_installed = addinfo.second.get<std::wstring>(L"Installed");
				auto str_prog_id = addinfo.second.get<std::wstring>(L"ProgId");
				auto str_user_name = addinfo.second.get<std::wstring>(L"UserName");
				auto str_user_sid = addinfo.second.get<std::wstring>(L"SID");



				DWORD		dw_loadbehavior = std::stoi(str_addin_load_behavior);

				auto str_key = str_user_sid + L"_" + str_prog_id;

				auto it = processinformation.addininformation_.find(str_key);
				if (it != processinformation.addininformation_.end())
				{
					processinformation.addininformation_[str_key].LoadBehavior_ = dw_loadbehavior;
					processinformation.addininformation_[str_key].Installed_ = str_addin_installed;
				}
			}
			XLSingleton::getInstance()->Set_Addin_info(processinformation);
			OnButtonSaveChange();
		}
	}
	catch (boost::exception const& ex)
	{
		LOG_ERROR << __FUNCTION__ << "-Unable to read the json file, file=" << p_filenamepath << " exception:" << diagnostic_information(ex);	
	}
}


void CMainFrame::OnButtonExportLogs()
{
	LOG_DEBUG << __FUNCTION__;
	fs::path p_filenamepath;

	try
	{
		p_filenamepath = SaveFileTxt();
		if (p_filenamepath.empty() == false)
		{
			SaveLogs(p_filenamepath);

		}
	}
	catch(std::exception ex)
	{
		LOG_ERROR << __FUNCTION__ << "-Unable to save the log file, file=" << p_filenamepath << " exception:" << ex.what();
	}
}

void CMainFrame::SaveLogs(fs::path p_filenamepath)
{
	USES_CONVERSION;

	std::ofstream file_log;

	file_log.open(p_filenamepath, std::ofstream::out);

	auto i_items = m_wndOutput.m_wndOutputDebug.GetCount();
	for (auto i = 0; i < i_items; i++)
	{
		CString s_value;
		m_wndOutput.m_wndOutputDebug.GetText(i, s_value);
		CStringA output = T2A(s_value);

		file_log << output;
		file_log << "\n";
	}
	file_log.close();
}