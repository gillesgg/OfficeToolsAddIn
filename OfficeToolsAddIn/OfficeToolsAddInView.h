#pragma once

#include "OfficeAddinInformation.h"

enum ListItem
{
	Name,
	ProgId,
	Location,
	Type,
	Load,
	StartupType
};


class COfficeToolsAddInView : public CListView
{
protected: // create from serialization only
	COfficeToolsAddInView() noexcept;
	DECLARE_DYNCREATE(COfficeToolsAddInView)

// Attributes
public:
	COfficeToolsAddInDoc* GetDocument() const;

// Operations
public:
	void ShowAddIns();
	//void DisableAddins();

// Overrides
public:
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
protected:
	virtual void OnInitialUpdate(); // called first time after construct


private:
	void CreateHeaders();
	std::wstring ResolveStartMode(DWORD startmode);
	UINT ResolveItemAddInOffice(std::wstring ProgID);
	UINT ResolveLoadOfficeAddIn(DWORD LoadBehavior);
	void UpdateItemAddInOffice(std::wstring ProgID, UINT nID);
	int ResolveMenuOfficeAddIn(UINT nID);
	std::wstring ResolveMenuAddInXl(UINT nID);
	int GetSelectedItem();
	AddInType GetSelectionAddInType();
	std::wstring ResolveItemAddInExcel(std::wstring ProgID);
	void UpdateItemAddInXL(std::wstring ProgID, UINT nID);
	void UpdateListView(int i_sub_item, std::wstring str_item_value);
	COLORREF ResolveItemColor(std::wstring str_progid);
private:

public:
	virtual ~COfficeToolsAddInView();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:

// Generated message map functions
protected:
	afx_msg void OnFilePrintPreview();
	afx_msg void OnRButtonUp(UINT nFlags, CPoint point);
	afx_msg void OnContextMenu(CWnd* pWnd, CPoint point);
	afx_msg void OnChangestartOfficeStartUp(UINT nID);
	afx_msg void OnChangestartExcelStartUp(UINT nID);
	afx_msg void OnUpdateRendererOfficeAddIn(CCmdUI* pCmdUI);
	afx_msg void OnUpdateRendererExcelAddIn(CCmdUI* pCmdUI);
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnNMCustomdraw(NMHDR* pNMHDR, LRESULT* pResult);
};

#ifndef _DEBUG  // debug version in OfficeToolsAddInView.cpp
inline COfficeToolsAddInDoc* COfficeToolsAddInView::GetDocument() const
   { return reinterpret_cast<COfficeToolsAddInDoc*>(m_pDocument); }
#endif

