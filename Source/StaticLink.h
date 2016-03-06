#ifndef STATICLINK_INCLUDED
#define STATICLINK_INCLUDED

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

//
// Sample usage of the class contained in this file.  To use these classes you must link with
// version.lib
//
// class AboutDlg: public FDialog
// {
// public:
//	   AboutDlg();
// 	   virtual BOOL DialogProc(UINT message,WPARAM wParam,LPARAM lParam);
// 	   CStaticLink m_theLink;
// 	   CStaticLink m_theLinkIcon;
// 	   CVersionInfo m_viFileVersion;
// };
//
// AboutDlg::AboutDlg()
//	   : FDialog(IDD_ABOUT),
// 	   m_theLink(_T(""), FALSE),
// 	   m_theLinkIcon(_T("http://www.datadynamics.com"), FALSE),
// 	   m_viFileVersion(g_hInstance, _T("actbar2.ocx"))
// {
// }
//
//
// BOOL AboutDlg::DialogProc(UINT message,WPARAM wParam,LPARAM lParam)
// {
//	 switch(message)
//	 {
//	 case WM_INITDIALOG:
//		 CenterDialog(GetParent(m_hWnd));
//		 m_theLink.SubclassDlgItem(IDC_WEBADDRESS, m_hWnd);
//		 m_theLinkIcon.SubclassDlgItem(IDC_ABICON, m_hWnd);
//		 m_viFileVersion.SetLabel(m_hWnd, IDC_VERSION, _T("ActiveBar, Version "), _T("FileVersion"));
//		 break;
//
//	 case WM_CTLCOLORSTATIC:
//		 return SendMessage((HWND)lParam, WM_CTLCOLORSTATIC, wParam, lParam);
//	 }
//	 return FDialog::DialogProc(message,wParam,lParam);
// }
//

//
// CHyperLink 
//

class CHyperLink 
{
public:
	CHyperLink(LPCTSTR szLink = 0)
	{ 
		int nLen = lstrlen(szLink)+sizeof(TCHAR);
		m_szLink = (TCHAR*)malloc(nLen);
		memcpy(m_szLink, szLink, nLen);
	}

	~CHyperLink() 
	{ 
		free (m_szLink);
	}

	const CHyperLink& operator=(LPCTSTR szLink) 
	{
		free (m_szLink);
		int nLen = lstrlen(szLink)+sizeof(TCHAR);
		m_szLink = (TCHAR*)malloc(nLen);
		lstrcpy(m_szLink, szLink);
		return *this;
	}

	operator LPCTSTR() 
	{
		return m_szLink;
	}

	BOOL IsEmpty();

	HINSTANCE Navigate();

	LPTSTR m_szLink;
};

inline BOOL CHyperLink::IsEmpty()
{
	return lstrlen(m_szLink) > 0 ? FALSE : TRUE;
}

//
// CStaticLink
//

class CStaticLink : public FWnd
{
public:
	CStaticLink(LPCTSTR szText, BOOL bDeleteOnDestroy);
	~CStaticLink();

	BOOL SubclassDlgItem(UINT nId, HWND hWndParent);

	virtual LRESULT WindowProc(UINT msg, WPARAM wParam, LPARAM lParam);

private:
	CHyperLink	m_hlLink;
	COLORREF	m_crForeColor;
	HFONT		m_hFont;			// underline font for text control
	BOOL		m_bDeleteOnDestroy;	// delete object when window destroyed?
};

//
// CVersionInfo
//

class CVersionInfo : public VS_FIXEDFILEINFO
{
private:

	struct TRANSLATION
	{
		WORD nLangId;
		WORD nCharSet;
	} m_tLanguage;

	BYTE* m_pVersionInfo;

public:
	CVersionInfo(HINSTANCE hInstance, LPCSTR szName);
	virtual ~CVersionInfo();

	LPCTSTR GetValue(LPCTSTR szValue);
	BOOL SetLabel(HWND hWnd, UINT nDlgItemId, LPCTSTR szPrompt, LPCTSTR szValue);
};

#endif