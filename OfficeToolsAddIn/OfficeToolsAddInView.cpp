// This MFC Samples source code demonstrates using MFC Microsoft Office Fluent User Interface
// (the "Fluent UI") and is provided only as referential material to supplement the
// Microsoft Foundation Classes Reference and related electronic documentation
// included with the MFC C++ library software.
// License terms to copy, use or distribute the Fluent UI are available separately.
// To learn more about our Fluent UI licensing program, please visit
// https://go.microsoft.com/fwlink/?LinkId=238214.
//
// Copyright (C) Microsoft Corporation
// All rights reserved.

// OfficeToolsAddInView.cpp : implementation of the COfficeToolsAddInView class
//

#include "pch.h"
#include "framework.h"
// SHARED_HANDLERS can be defined in an ATL project implementing preview, thumbnail
// and search filter handlers and allows sharing of document code with that project.
#ifndef SHARED_HANDLERS
#include "OfficeToolsAddIn.h"
#endif

#include "OfficeToolsAddInDoc.h"
#include "OfficeToolsAddInView.h"
#include "ExcelAutomation.h"
#include "XLSingleton.h"
#include "OfficeAddinInformation.h"
#include "resource.h"
#include "Logger.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#define NO_SELECTION -1


IMPLEMENT_DYNCREATE(COfficeToolsAddInView, CListView)

BEGIN_MESSAGE_MAP(COfficeToolsAddInView, CListView)
	ON_WM_CONTEXTMENU()
	ON_WM_RBUTTONUP()
	ON_COMMAND_RANGE(ID_UNLOADED_DO_NOT_LOAD_AUTOMATICALLY, ID_UNKNOWN, &COfficeToolsAddInView::OnChangestartOfficeStartUp)
	ON_COMMAND_RANGE(ID_CHANGECONNECTMODE_NO, ID_CHANGECONNECTMODE_YES, &COfficeToolsAddInView::OnChangestartExcelStartUp)
	ON_UPDATE_COMMAND_UI_RANGE(ID_UNLOADED_DO_NOT_LOAD_AUTOMATICALLY, ID_UNKNOWN, OnUpdateRendererOfficeAddIn)
	ON_UPDATE_COMMAND_UI_RANGE(ID_CHANGECONNECTMODE_NO, ID_CHANGECONNECTMODE_YES, OnUpdateRendererExcelAddIn)
	ON_NOTIFY_REFLECT(NM_CUSTOMDRAW, &COfficeToolsAddInView::OnNMCustomdraw)
END_MESSAGE_MAP()


COfficeToolsAddInView::COfficeToolsAddInView() noexcept
{}
COfficeToolsAddInView::~COfficeToolsAddInView()
{}


#pragma region MenuHandler

/// <summary>
/// Check uncheck Handler for Excel AddIn popup menu
/// </summary>
/// <param name="pCmdUI"></param>
void COfficeToolsAddInView::OnUpdateRendererExcelAddIn(CCmdUI* pCmdUI)
{
	if (GetSelectionAddInType() == AddInType::XL)
	{
		auto i_pos = GetSelectedItem();
		if (i_pos != NO_SELECTION)
		{
			auto str_load = ResolveItemAddInExcel(GetListCtrl().GetItemText(i_pos, ListItem::ProgId).GetBuffer());
			if (str_load == L"True" && pCmdUI->m_nID == ID_CHANGECONNECTMODE_YES)
			{
				pCmdUI->SetCheck(TRUE);
			}
			if (str_load == L"False" && pCmdUI->m_nID == ID_CHANGECONNECTMODE_NO)
			{
				pCmdUI->SetCheck(TRUE);
			}
		}
	}
	if (GetSelectionAddInType() == AddInType::OFFICE)
	{
		pCmdUI->Enable(FALSE);
	}
}
/// <summary>
/// Check uncheck Handler for Office AddIn popup menu
/// </summary>
/// <param name="pCmdUI"></param>
void COfficeToolsAddInView::OnUpdateRendererOfficeAddIn(CCmdUI* pCmdUI)
{
	if (GetSelectionAddInType() == AddInType::OFFICE)
	{
		auto i_pos = GetSelectedItem();

		if (i_pos != NO_SELECTION)
		{
			auto i_value = ResolveItemAddInOffice(GetListCtrl().GetItemText(i_pos, ListItem::ProgId).GetBuffer());
			if (i_value != 200 && pCmdUI->m_nID == i_value)
			{
				pCmdUI->SetCheck(TRUE);
			}
			else
			{
				pCmdUI->SetCheck(FALSE);
			}
		}
	}
	if (GetSelectionAddInType() == AddInType::XL)
	{
		pCmdUI->Enable(FALSE);
	}
}

#pragma endregion


/// <summary>
/// Return the current selected item
/// </summary>
/// <returns></returns>
int COfficeToolsAddInView::GetSelectedItem()
{
	int ipos = NO_SELECTION;
	POSITION pos = this->GetListCtrl().GetFirstSelectedItemPosition();
	if (pos != nullptr)
	{
		ipos = this->GetListCtrl().GetNextSelectedItem(pos);
	}
	return ipos;
}
/// <summary>
/// Return the current addin type of the selected item 
/// </summary>
/// <returns></returns>
AddInType COfficeToolsAddInView::GetSelectionAddInType()
{
	int ipos = GetSelectedItem();

	if (ipos != NO_SELECTION)
	{
		CString addintype = GetListCtrl().GetItemText(ipos, 3);
		CString officeAddin;
		CString xladdin;

		officeAddin.LoadString(IDS_STRING_COM_ADD_IN);
		xladdin.LoadString(IDS_STRING_EXCEL_ADD_IN);

		if (addintype == officeAddin)
		{
			return AddInType::OFFICE;
		}
		else if (addintype == xladdin)
		{
			return AddInType::XL;
		}
	}
	return AddInType::NONE;
}

std::wstring COfficeToolsAddInView::ResolveItemAddInExcel(std::wstring ProgID)
{
	LOG_TRACE << __FUNCTION__;

	ProcessInformation processinformation = XLSingleton::getInstance()->Get_Addin_info();
	if (processinformation.addininformation_.size() > 0 && processinformation.addininformation_.count(ProgID) > 0)
	{
		return processinformation.addininformation_[ProgID].Installed_;
	}
	return std::wstring();
}

UINT COfficeToolsAddInView::ResolveItemAddInOffice(std::wstring ProgID)
{
	LOG_TRACE << __FUNCTION__;

	ProcessInformation processinformation = XLSingleton::getInstance()->Get_Addin_info();
	if (processinformation.addininformation_.size() > 0 && processinformation.addininformation_.count(ProgID) > 0)
	{
		return (ResolveLoadOfficeAddIn(processinformation.addininformation_[ProgID].LoadBehavior_));
	}
	return 200;
}

UINT COfficeToolsAddInView::ResolveLoadOfficeAddIn(DWORD LoadBehavior)
{
	switch (LoadBehavior)
	{
	case 0:
		return ID_UNLOADED_DO_NOT_LOAD_AUTOMATICALLY;
		break;
	case 1:
		return ID_LOADED_DO_NOT_LOAD_AUTOMATICALLY;
		break;
	case 2:
		return ID_UNLOADED_LOAD_AT_STARTUP;
		break;
	case 3:
		return ID_LOADED_LOAD_AT_STARTUP;
		break;
	case 8:
		return ID_UNLOADED_LOAD_ONDEMAND;
		break;
	case 9:
		return ID_LOADED_LOAD_ON_DEMAND;
		break;
	case 16:
		return  ID_LOADED_LOAD_FIRST_TIME_THEN_LOAD_ON_DEMAND;
		break;
	case 100:
		return ID_NOT_APPLICABLE;
		break;

	default:
		return ID_UNKNOWN;
		break;
	}
}

BOOL COfficeToolsAddInView::PreCreateWindow(CREATESTRUCT& cs)
{
	LOG_TRACE << __FUNCTION__;

	cs.style |= LVS_REPORT;
	return CListView::PreCreateWindow(cs);
}

void COfficeToolsAddInView::OnInitialUpdate()
{
	LOG_TRACE << __FUNCTION__;

	CListView::OnInitialUpdate();
	CreateHeaders();
	
}

void COfficeToolsAddInView::OnRButtonUp(UINT /* nFlags */, CPoint point)
{
	ClientToScreen(&point);
	OnContextMenu(this, point);
}

void COfficeToolsAddInView::OnContextMenu(CWnd* /* pWnd */, CPoint point)
{
#ifndef SHARED_HANDLERS
	theApp.GetContextMenuManager()->ShowPopupMenu(IDR_POPUP_EDIT, point.x, point.y, this, TRUE);
#endif
}



std::wstring COfficeToolsAddInView::ResolveStartMode(DWORD startmode)
{
	LOG_TRACE << __FUNCTION__;

	switch (startmode)
	{
	case 0:
		return L"Unloaded-Do not load automatically";
		break;
	case 1:
		return L"Loaded-Do not load automatically";
		break;
	case 2:
		return L"Unloaded-Load at startup";
		break;
	case 3:
		return L"Loaded-Load at startup";
		break;
	case 8:
		return L"Unloaded-Load on demand";
		break;
	case 9:
		return L"Loaded-Load on demand";
		break;
	case 16:
		return L"Loaded-Load first time, then load on demand";
		break;
	case 100:
		return L"Not applicable";
		break;

	default:
		return L"Unknown";
		break;
	}
}

void COfficeToolsAddInView::CreateHeaders()
{
	LOG_TRACE << __FUNCTION__;

	int x = 0;

	this->GetListCtrl().SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);
	this->GetListCtrl().InsertColumn(x++, L"Name", LVCFMT_LEFT, 300);
	this->GetListCtrl().InsertColumn(x++, L"ProgId", LVCFMT_LEFT, 300);	
	this->GetListCtrl().InsertColumn(x++, L"Location", LVCFMT_LEFT, 300);
	this->GetListCtrl().InsertColumn(x++, L"Type", LVCFMT_LEFT, 300);
	this->GetListCtrl().InsertColumn(x++, L"Load", LVCFMT_LEFT, 300);
	this->GetListCtrl().InsertColumn(x++, L"Startup type", LVCFMT_LEFT, 300);
}

void COfficeToolsAddInView::ShowAddIns()
{
	LOG_TRACE << __FUNCTION__;

	int x = 0;

	ProcessInformation processinformation = XLSingleton::getInstance()->Get_Addin_info();
	this->GetListCtrl().DeleteAllItems();

	for (auto item : processinformation.addininformation_)
	{
		auto d = item.second.Description_;
		CString sresource;
		LVITEM lvi;
				
		lvi.mask = LVIF_TEXT /*| LVIF_GROUPID*/;
		lvi.iItem = x;
		lvi.iSubItem = 0;
		lvi.pszText = (wchar_t*)(item.second.Description_.c_str());

		this->GetListCtrl().InsertItem(&lvi);

		lvi.iSubItem = 1;
		lvi.pszText = (wchar_t*)(item.second.ProgId_.c_str());
		this->GetListCtrl().SetItem(&lvi);

		lvi.iSubItem = 2;
		lvi.pszText = (wchar_t*)(item.second.FullName_.c_str());
		this->GetListCtrl().SetItem(&lvi);

		lvi.iSubItem = 3;
		(item.second.addType_ == AddInType::OFFICE) ? sresource.LoadString(IDS_STRING_COM_ADD_IN) : sresource.LoadString(IDS_STRING_EXCEL_ADD_IN);
		lvi.pszText = sresource.GetBuffer();
		this->GetListCtrl().SetItem(&lvi);

		lvi.iSubItem = 4;
		lvi.pszText = (wchar_t*) item.second.Installed_.c_str();
		this->GetListCtrl().SetItem(&lvi);

		auto smode = ResolveStartMode(item.second.LoadBehavior_);

		lvi.iSubItem = 5;
		lvi.pszText = (WCHAR*)smode.c_str();
		this->GetListCtrl().SetItem(&lvi);
	}	
}
#pragma region UpdateAddInParameters

/// <summary>
/// Handle the Menu Office AddIn start mode
/// </summary>
/// <param name="nID"></param>
void COfficeToolsAddInView::OnChangestartOfficeStartUp(UINT nID)
{
	LOG_TRACE << __FUNCTION__;

	POSITION p_pos = this->GetListCtrl().GetFirstSelectedItemPosition();
	if (p_pos != nullptr)
	{
		auto i_pos = this->GetListCtrl().GetNextSelectedItem(p_pos);
		auto str_prog_id = GetListCtrl().GetItemText(i_pos, 1);
		UpdateItemAddInOffice(str_prog_id.GetBuffer(), nID);
	}
}
/// <summary>
/// Handle the Menu Excel AddIn start mode
/// </summary>
/// <param name="nID"></param>
void COfficeToolsAddInView::OnChangestartExcelStartUp(UINT nID)
{
	POSITION p_pos = this->GetListCtrl().GetFirstSelectedItemPosition();
	if (p_pos != nullptr)
	{
		auto i_pos = this->GetListCtrl().GetNextSelectedItem(p_pos);
		auto str_prog_id = GetListCtrl().GetItemText(i_pos, ListItem::ProgId);
		UpdateItemAddInXL(str_prog_id.GetBuffer(), nID);
	}
}
/// <summary>
/// Update the singleton and the listview
/// </summary>
/// <param name="ProgID"></param>
/// <param name="nID"></param>
void COfficeToolsAddInView::UpdateItemAddInXL(std::wstring ProgID, UINT nID)
{
	LOG_TRACE << __FUNCTION__;
	ProcessInformation processinformation = XLSingleton::getInstance()->Get_Addin_info();
	if (processinformation.addininformation_.size() > 0 && processinformation.addininformation_.count(ProgID) > 0)
	{
		//auto item = processinformation.addininformation_[ProgID];
		auto str_value = ResolveMenuAddInXl(nID);
		processinformation.addininformation_[ProgID].Installed_ = str_value;
		XLSingleton::getInstance()->Set_Addin_info(processinformation);
		UpdateListView(ListItem::Load, str_value);

	}
}
/// <summary>
/// Update the singleton and the listview
/// </summary>
/// <param name="ProgID"></param>
/// <param name="nID"></param>
void COfficeToolsAddInView::UpdateItemAddInOffice(std::wstring ProgID,UINT nID )
{
	LOG_TRACE << __FUNCTION__;
	ProcessInformation processinformation = XLSingleton::getInstance()->Get_Addin_info();
	if (processinformation.addininformation_.size() > 0 && processinformation.addininformation_.count(ProgID) > 0)
	{
		//auto item = processinformation.addininformation_[ProgID];
		auto i_value = ResolveMenuOfficeAddIn(nID);
		if (i_value != 100)
		{
			processinformation.addininformation_[ProgID].LoadBehavior_ = i_value;
			XLSingleton::getInstance()->Set_Addin_info(processinformation);			
			UpdateListView(ListItem::StartupType, ResolveStartMode(i_value));
		}
	}
}

void COfficeToolsAddInView::UpdateListView(int i_sub_item, std::wstring str_item_value)
{
	LOG_TRACE << __FUNCTION__;
	POSITION p_pos = this->GetListCtrl().GetFirstSelectedItemPosition();
	if (p_pos != nullptr)
	{
		auto i_pos = this->GetListCtrl().GetNextSelectedItem(p_pos);

		LVITEM lvi;

		lvi.mask = LVIF_TEXT;
		lvi.iItem = i_pos;
		lvi.iSubItem = i_sub_item;
		lvi.pszText = (wchar_t*)(str_item_value.c_str());
		this->GetListCtrl().SetItem(&lvi);
	}
}

#pragma endregion

int COfficeToolsAddInView::ResolveMenuOfficeAddIn(UINT nID)
{
	switch (nID)
	{
	case ID_UNLOADED_DO_NOT_LOAD_AUTOMATICALLY:
		return 0;
		break;
	case ID_LOADED_DO_NOT_LOAD_AUTOMATICALLY:
		return 1;
		break;
	case ID_UNLOADED_LOAD_AT_STARTUP:
		return 2;
		break;
	case ID_LOADED_LOAD_AT_STARTUP:
		return 3;
		break;
	case ID_UNLOADED_LOAD_ONDEMAND:
		return 8;
		break;
	case ID_LOADED_LOAD_ON_DEMAND:
		return 9;
		break;
	case ID_LOADED_LOAD_FIRST_TIME_THEN_LOAD_ON_DEMAND:
		return  16;
		break;
	case ID_NOT_APPLICABLE:
		return 100;
		break;
	default:
		return 100;
		break;
	}
}

std::wstring COfficeToolsAddInView::ResolveMenuAddInXl(UINT nID)
{
	switch (nID)
	{
	case ID_CHANGECONNECTMODE_NO:
		return L"False";
		break;
	case ID_CHANGECONNECTMODE_YES:
		return L"True";
		break;
	default:
		return L"Not applicable";
		break;
	}
}


#ifdef _DEBUG
void COfficeToolsAddInView::AssertValid() const
{
	CListView::AssertValid();
}

void COfficeToolsAddInView::Dump(CDumpContext& dc) const
{
	CListView::Dump(dc);
}

COfficeToolsAddInDoc* COfficeToolsAddInView::GetDocument() const // non-debug version is inline
{
	ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(COfficeToolsAddInDoc)));
	return (COfficeToolsAddInDoc*)m_pDocument;
}
#endif //_DEBUG


void COfficeToolsAddInView::OnNMCustomdraw(NMHDR* pNMHDR, LRESULT* pResult)
{
	NMLVCUSTOMDRAW* pLVCD = reinterpret_cast<NMLVCUSTOMDRAW*>(pNMHDR);
	*pResult = 0;

	if (CDDS_PREPAINT == pLVCD->nmcd.dwDrawStage)
	{
		*pResult = CDRF_NOTIFYITEMDRAW;
	}
	else   if (CDDS_ITEMPREPAINT == pLVCD->nmcd.dwDrawStage)
	{
		*pResult = CDRF_NOTIFYSUBITEMDRAW;
	}
	else   if ((CDDS_ITEMPREPAINT | CDDS_SUBITEM) == pLVCD->nmcd.dwDrawStage)
	{
		auto row_item = pLVCD->nmcd.dwItemSpec;

		

		auto rgb_value = ResolveItemColor(GetListCtrl().GetItemText(row_item, ListItem::ProgId).GetBuffer());

		if (rgb_value != RGB(0,0,0))
			pLVCD->clrText = rgb_value;
		*pResult = CDRF_DODEFAULT;
	}
}

COLORREF COfficeToolsAddInView::ResolveItemColor(std::wstring str_progid)
{
	ProcessInformation processinformation = XLSingleton::getInstance()->Get_Addin_info();
	auto i = processinformation.addininformation_.find(str_progid);
	if (i != processinformation.addininformation_.end())
	{
		if (i->second.addType_ == AddInType::XL && i->second.Installed_ == L"True")
		{
			return RGB(0, 128, 255);
		}

		if (i->second.addType_ == AddInType::OFFICE && (i->second.LoadBehavior_ == 1 || i->second.LoadBehavior_ == 3 || i->second.LoadBehavior_ == 9 || i->second.LoadBehavior_ == 16))
		{
			return RGB(0, 128, 0);
		}
	}
	return RGB(255, 0, 0);
}
