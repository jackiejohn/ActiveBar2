#ifndef __CDesignerPage_H__
#define __CDesignerPage_H__

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "proppage.h"
#include "..\Interfaces.h"
#include "..\PrivateInterfaces.h"
#include "DesignerInterfaces.h"

struct CategoryMgr;
class CCategoryTree;
class CDesignerTree;
class SEException;
class CSplitter;
class CBrowser;

class CDesignerPage : public CPropertyPage
, public IDesignerNotify 
{
public:
	CDesignerPage();
	~CDesignerPage();
	
	void Cleanup();

	enum
	{
		eOk = 1,
		eCancel = 2,
		eApply = 3021,
		eHelp = 9
	};

	void Update();

private:
	int MessageBox(LPTSTR szMsg, UINT nMsgType = MB_OK);
	int MessageBox(UINT nId, UINT nMsgType = MB_OK);
	virtual BOOL DialogProc(HWND,UINT,WPARAM,LPARAM);
    STDMETHOD(TranslateAccelerator) (LPMSG lpMsg);
    STDMETHOD(Close)();
    STDMETHOD(CreateTool)(IDispatch* pBand, long nToolType);
    STDMETHOD(SetDirty)();
	void OnLButtonDown(UINT nFlags, POINT pt);
	void OnMouseMove(UINT nFlags, POINT pt);
	void OnLButtonUp(UINT nFlags, POINT pt);

	BOOL IsButtonDisabled(UINT nId);
	void SaveDisableButtons();
	void RestoreButtons();
	void OnHeader();
	void OnLibrary();
	void OnMenuGrab();
	void OnReport();
	void OnOptions();
	void OnGenSelect();

	BOOL CreateMenu(HMENU hMenu);
	BOOL CreateSubMenu(IBands* pBands, ITool* pTool, BSTR bstrName, BSTR bstrCaption, HMENU hMenu, int nItem);
	HRESULT CreateDesignerToolbar (long nX,  long nY,  long nWidth,  long nHeight);
	HRESULT OnCommand(long wParam, long lParam);

	CDesignerTree* m_pDesignerTree;
	CategoryMgr*   m_pCategoryMgr;
	IActiveBar2*   m_pActiveBar;
	DDToolBar*     m_ptbDesigner;
	CSplitter*     m_pSplitter;
	CBrowser*      m_pBrowser;
	DDString	   m_strMsgTitle;
	VARIANT		   m_vLayoutData;

#pragma pack(push)
#pragma pack (1)
	BOOL m_bDesignerView:1;
	BOOL m_bBrowserActive:1;
	BOOL m_bOk:1;
	BOOL m_bCancel:1;
	BOOL m_bApply:1;
	BOOL m_bHelp:1;
#pragma pack (pop)
};

inline int CDesignerPage::MessageBox(LPTSTR szMsg, UINT nMsgType)
{
	return ::MessageBox (m_hwnd, szMsg, m_strMsgTitle, nMsgType);
}

inline int CDesignerPage::MessageBox(UINT nId, UINT nMsgType)
{
	DDString strMsg;
	strMsg.LoadString(nId);
	return ::MessageBox (m_hwnd, strMsg, m_strMsgTitle, nMsgType);
}

inline void CDesignerPage::Update()
{
	MakeDirty();
}

#endif
