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

// MainFrm.h : interface of the CMainFrame class
//

#pragma once
#include "OfficeAddinInformation.h"
#include "OutputWnd.h"


class CMainFrame : public CFrameWndEx
{
	
protected: // create from serialization only
	CMainFrame() noexcept;
	DECLARE_DYNCREATE(CMainFrame)

private:
	void ExportFile(fs::path str_file_name);
	std::wstring CreateUniqueFile();
	fs::path GenerateTemporyFile();

private:
	fs::path temp_file_export_;
// Attributes
public:

// Operations
public:

// Overrides
public:
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
	void OnUpdateFrameTitle(BOOL Nada);
// Implementation
public:
	virtual ~CMainFrame();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:  // control bar embedded members
	CMFCRibbonBar     m_wndRibbonBar;
	CMFCRibbonApplicationButton m_MainButton;
	CMFCToolBarImages m_PanelImages;
	//CMFCRibbonStatusBar	m_wndStatusBar;
	COutputWnd			m_wndOutput;
	CMFCRibbonApplicationButton m_wndRibbonButton;

private:
	fs::path SaveFile();
	fs::path OpenFile();
	fs::path SaveFileTxt();
	void SaveLogs(fs::path p_filenamepath);
	
// Generated message map functions
protected:
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnSettingChange(UINT uFlags, LPCTSTR lpszSection);
	DECLARE_MESSAGE_MAP()

	BOOL CreateDockingWindows();
	void SetDockingWindowIcons(BOOL bHiColorIcons);
public:
	afx_msg void OnButtonRefresh();
	afx_msg void OnButtonDisable();
	afx_msg void OnButtonExport();
	afx_msg void OnButtonImport();
	afx_msg void OnButtonSaveChange();
	afx_msg void OnButtonExportLogs();
};


