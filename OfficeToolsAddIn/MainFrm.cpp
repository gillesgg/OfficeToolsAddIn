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
	ON_COMMAND(ID_BUTTON_SAVE_CHANGE, &CMainFrame::OnButtonSaveChange)
END_MESSAGE_MAP()

CMainFrame::CMainFrame() noexcept
{}

CMainFrame::~CMainFrame()
{}


void CMainFrame::OnUpdateFrameTitle(BOOL Nada)
{
	CString csAppName;
	csAppName.Format(AFX_IDS_APP_TITLE);
	SetWindowText(csAppName);
}

int CMainFrame::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	LOG_TRACE << __FUNCTION__;

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

	/*m_wndOutput.EnableDocking(CBRS_ALIGN_ANY);
	DockPane(&m_wndOutput);*/

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
	//BOOL		bNameValid;
	//CString		strOutputWnd;

	//bNameValid = strOutputWnd.LoadString(IDS_OUTPUT_WND);
	//ASSERT(bNameValid);
	//if (!m_wndOutput.Create(strOutputWnd, this, CRect(0, 0, 100, 100), TRUE, ID_VIEW_OUTPUTWND, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | CBRS_BOTTOM | CBRS_FLOAT_MULTI))
	//{
	//	LOG_ERROR << __FUNCTION__ << " Failed to create Output window";
	//	return FALSE; // failed to create
	//}
	//SetDockingWindowIcons(theApp.m_bHiColorIcons);
	return TRUE;
}

void CMainFrame::SetDockingWindowIcons(BOOL bHiColorIcons)
{
	/*HICON hOutputBarIcon = (HICON) ::LoadImage(::AfxGetResourceHandle(), MAKEINTRESOURCE(bHiColorIcons ? IDI_OUTPUT_WND_HC : IDI_OUTPUT_WND), IMAGE_ICON, ::GetSystemMetrics(SM_CXSMICON), ::GetSystemMetrics(SM_CYSMICON), 0);
	m_wndOutput.SetIcon(hOutputBarIcon, FALSE);*/

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
	//m_wndOutput.UpdateFonts();
}

void CMainFrame::OnButtonRefresh()
{
	LOG_TRACE << __FUNCTION__;

	OfficeAddIn office_add_in;

	BeginWaitCursor();
	office_add_in.ReadAddinInformation();
	if (nullptr != GetActiveView())
	{
		(dynamic_cast<COfficeToolsAddInView*>(GetActiveView()))->ShowAddIns();
	}
	EndWaitCursor();

}

void CMainFrame::OnButtonDisable()
{
	LOG_TRACE << __FUNCTION__;
	
	OfficeAddIn office_add_in;

	BeginWaitCursor();
	office_add_in.DisableAllOfficeAddIn();
	if (nullptr != GetActiveView())
	{
		(dynamic_cast<COfficeToolsAddInView*>(GetActiveView()))->ShowAddIns();
	}
	EndWaitCursor();
}

void CMainFrame::OnButtonSaveChange()
{
	LOG_TRACE << __FUNCTION__;

	OfficeAddIn office_add_in;
	office_add_in.SaveAddinInformation();

	BeginWaitCursor();

	if (nullptr != GetActiveView())
	{
		(dynamic_cast<COfficeToolsAddInView*>(GetActiveView()))->ShowAddIns();
	}
	EndWaitCursor();
}


fs::path CMainFrame::SaveFile()
{
	LOG_TRACE << __FUNCTION__;
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
				LOG_ERROR << __FUNCTION__ << "Unable to remote the file, file=" << p_filename_path << "exception:" << ex.what();
				return fs::path();
			}
		}
	}
	return p_filename_path;
}

fs::path CMainFrame::OpenFile()
{
	LOG_TRACE << __FUNCTION__;
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

void CMainFrame::OnButtonExport()
{
	LOG_TRACE << __FUNCTION__;
	fs::path p_filenamepath;

	try
	{
		ProcessInformation processinformation = XLSingleton::getInstance()->Get_Addin_info();
		p_filenamepath = SaveFile();
		if (p_filenamepath.empty() == false)
		{
			pt::wptree root;
			root.put(L"ImageType", processinformation.imagetype_ == ImageType::x64 ? L"X64" : L"X86");
			root.put(L"Name", processinformation.Name_);
			pt::wptree children;
			for (auto& addininfo : processinformation.addininformation_)
			{
				pt::wptree child;

				child.put(L"Description", addininfo.second.Description_);
				child.put(L"Installed", addininfo.second.Installed_);
				child.put(L"key", addininfo.second.key_);
				child.put(L"parent", addininfo.second.parent_ == HKEY_LOCAL_MACHINE ? L"HKEY_LOCAL_MACHINE" : L"HKEY_CURRENT_USER");
				child.put(L"AddInType", addininfo.second.addType_ == AddInType::OFFICE ? L"Office" : L"Excel");
				child.put(L"LoadBehavior", std::to_wstring(addininfo.second.LoadBehavior_));
				child.put(L"ProgId", addininfo.second.ProgId_);
				children.push_back(std::make_pair(L"", child));
			}
			root.add_child(L"AddinInformation", children);
			pt::write_json(p_filenamepath.generic_string(), root);
		}
	}
	catch (boost::exception const& ex)
	{
		LOG_ERROR << __FUNCTION__ << "Unable to write the json file, file=" << p_filenamepath << "exception:" << diagnostic_information(ex);
	}
}


void CMainFrame::OnButtonImport()
{
	LOG_TRACE << __FUNCTION__;
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

				DWORD		dw_loadbehavior = std::stoi(str_addin_load_behavior);

				auto it = processinformation.addininformation_.find(str_prog_id);
				if (it != processinformation.addininformation_.end())
				{
					processinformation.addininformation_[str_prog_id].LoadBehavior_ = dw_loadbehavior;
					processinformation.addininformation_[str_prog_id].Installed_ = str_addin_installed;
				}
			}
			XLSingleton::getInstance()->Set_Addin_info(processinformation);
			OnButtonSaveChange();
		}
	}
	catch (boost::exception const& ex)
	{
		LOG_ERROR << __FUNCTION__ << "Unable to read the json file, file=" << p_filenamepath << "exception:" << diagnostic_information(ex);	
	}
}