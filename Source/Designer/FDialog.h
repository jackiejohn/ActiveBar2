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

class SEException;

//
// FDialog
//

class FDialog
{
public:
	FDialog(UINT nId);
	virtual ~FDialog();

	int DoModal(HWND hWndParent = NULL);
#ifdef _DDINC_MODELESS_CODE
	void DoModeless(HWND hWndParent);
#endif

	void CenterDialog(HWND hWndRelativeWindow);
	int MessageBox(LPTSTR szMsg, UINT nMsgType = MB_OK);
	int MessageBox(UINT nResourceId, UINT nMsgType = MB_OK);

	virtual void OnOK();
	virtual void OnCancel();
	virtual BOOL PreTranslateMessage(MSG* msg) {return FALSE;};
	HWND hWnd();
protected:
	void SetPreTranslateHook();
	void ClearPreTranslateHook();
	static BOOL CALLBACK FDialogProc(HWND, UINT, WPARAM, LPARAM);
    virtual BOOL DialogProc(UINT, WPARAM, LPARAM);

	BOOL ExecuteDlgInit();
	BOOL ExecuteDlgInit(LPVOID pResource);
	BOOL SetTemplate(const DLGTEMPLATE* pTemplate, UINT cb);
	BOOL SetFont(LPCTSTR szFaceName, WORD nFontSize);
	BOOL IsSysFont();

	HGLOBAL m_hTemplate;
	DDString m_strMsgTitle;
	DWORD   m_dwTemplateSize;
	BOOL    m_bPreTranslate;
	BOOL    m_bSystemFont;
	UINT    m_nDialogId;
	HWND    m_hWnd;
#ifdef _DDINC_MODELESS_CODE
	BOOL m_bModeless;
#endif
};

inline HWND FDialog::hWnd()
{
	return m_hWnd;
}

inline int FDialog::MessageBox(LPTSTR szMsg, UINT nMsgType)
{
	return ::MessageBox (m_hWnd, szMsg, (LPCTSTR)m_strMsgTitle, nMsgType);
}

inline int FDialog::MessageBox(UINT nResourceId, UINT nMsgType)
{
	DDString strMsg;
	strMsg.LoadString(nResourceId);
	return ::MessageBox (m_hWnd, strMsg, m_strMsgTitle, nMsgType);
}

//
// CFileDialog
//

class CFileDialog
{
public:
	CFileDialog(UINT     nTitleId,
				BOOL     bOpenFileDialog = TRUE, 
				LPCTSTR  szDefExt   = NULL, 
				LPCTSTR  szFileName = NULL, 
				DWORD    dwFlags    = OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, 
				LPTSTR   szFilter   = NULL, 
				HWND     hWndParent = NULL,
				UINT     nResourceId = 0,
				LPOFNHOOKPROC pFn = NULL);
	~CFileDialog();
	
	UINT DoModal();
	LPCTSTR GetFileName();
	void SetFileName(LPCTSTR szFileName);
	int GetFilterIndex() const;
	void SetFilterIndex(int nFilterIndex);

private:
	OPENFILENAME m_ofn;
	DDString	 m_strTitle;
	TCHAR	     m_szInitFileName[MAX_PATH];
	TCHAR		 m_szFileTitle[MAX_PATH];
	BOOL		 m_bOpenFileDialog;
};

inline int CFileDialog::GetFilterIndex() const 
{
	return m_ofn.nFilterIndex;
}

inline void CFileDialog::SetFilterIndex(int nFilterIndex)
{
	m_ofn.nFilterIndex = nFilterIndex;
}

inline LPCTSTR CFileDialog::GetFileName() 
{
	return m_szInitFileName;
};

inline void CFileDialog::SetFileName(LPCTSTR szFileName) 
{
	_tcscpy(m_szInitFileName, szFileName);
};

//
// CFontDialog
//

class CFontDialog
{
public:
	CFontDialog(HFONT hFont);

	BOOL DoModal();

	HFONT GetFont();

private:
	CHOOSEFONT m_theChooseFont;
	LOGFONT m_lf;
}; 
 
//
// DoModal
//

inline BOOL CFontDialog::DoModal()
{
	return ChooseFont(&m_theChooseFont);
}

//
// CColorDialog
//

class CColorDialog
{
public:
	CColorDialog(COLORREF crColor);

	BOOL DoModal();

	COLORREF& GetColor();

private:
	CHOOSECOLOR m_theChooseColor;
	COLORREF m_crColor;
	COLORREF m_crUserDefined[16];
}; 
 
//
// DoModal
//

inline BOOL CColorDialog::DoModal()
{
	return ChooseColor(&m_theChooseColor);
}

#endif
