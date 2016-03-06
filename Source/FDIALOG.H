#ifndef __FDIALOG_H__
#define __FDIALOG_H__

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

class FDialog
{
public:
	FDialog(UINT nId);
	virtual ~FDialog();
	HWND m_hWnd;

	int DoModal(HWND hWndParent = NULL);
#ifdef _DDINC_MODELESS_CODE
	void DoModeless(HWND hWndParent);
#endif

	void CenterDialog(HWND HWndRelative);
	
	static int CheckMap();
	static void Init();
	static void CleanUp();

	virtual void OnOK();
	virtual void OnCancel();
	virtual BOOL PreTranslateMessage(MSG* pMsg) {return FALSE;};
	void ScreenToClient(RECT& rc);

protected:
	void SetPreTranslateHook();
	void ClearPreTranslateHook();
	static BOOL CALLBACK FDialogProc(HWND, UINT, WPARAM, LPARAM);
    virtual BOOL DialogProc(UINT, WPARAM, LPARAM);

	BOOL SetTemplate(const DLGTEMPLATE* pTemplate, UINT cb);
	BOOL ExecuteDlgInit();
	BOOL ExecuteDlgInit(LPVOID pResource);
	BOOL SetFont(LPCTSTR lpFaceName, WORD nFontSize);
	BOOL IsSysFont();

	HGLOBAL m_hTemplate ;
	DWORD   m_dwTemplateSize;
	BOOL    m_bSystemFont;
	BOOL    m_bPreTranslate;
	UINT    m_nDialogId;

#ifdef _DDINC_MODELESS_CODE
	BOOL m_bModeless;
#endif
};

inline void FDialog::ScreenToClient(RECT& rc)
{ 
	::ScreenToClient(m_hWnd, (LPPOINT)&rc);
	::ScreenToClient(m_hWnd, ((LPPOINT)&rc)+1); 
}

#endif
