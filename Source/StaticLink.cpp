//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "StaticLink.h"

//
// CHyperLink
//

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

//
// Navigate link -- ie, execute the file
// Returns instance handle of app run, or error code (just like ShellExecute)
//

HINSTANCE CHyperLink::Navigate()
{
	return ShellExecute(NULL, _T("open"), m_szLink, 0, 0, SW_SHOWNORMAL);
}

static COLORREF g_crUnvisited = RGB(0,0,255); // blue
static COLORREF g_crVisited   = RGB(128,0,128); // purple
static HCURSOR	g_hCursorLink = NULL; 

//
// CStaticLink
//

//
// Constructor sets default colors = blue/purple.
//

CStaticLink::CStaticLink(LPCTSTR szText, BOOL bDeleteOnDestroy)
	:	m_hlLink(szText),
		m_hFont(0)
{
	m_crForeColor = g_crUnvisited;		   // not visited yet
	m_bDeleteOnDestroy = bDeleteOnDestroy;
}

CStaticLink::~CStaticLink()
{
}

BOOL CStaticLink::SubclassDlgItem(UINT nId, HWND hWndParent)
{
	m_hWnd = GetDlgItem(hWndParent, nId);
	SubClassAttach();
	return TRUE;
}

LRESULT CStaticLink::WindowProc(UINT msg, WPARAM wParam, LPARAM lParam)
{
	switch (msg)
	{
	case WM_NCHITTEST:
		return HTCLIENT;

	case WM_CTLCOLORSTATIC:
		{
			DWORD dwStyle = GetWindowLong(m_hWnd, GWL_STYLE);
			HBRUSH hBrush = NULL;
			if ((dwStyle & 0xFF) <= SS_RIGHT)
			{
				HDC hDC = (HDC)wParam;
				if (NULL == m_hFont)
				{
					LOGFONT lf;
					m_hFont = (HFONT)::SendMessage(m_hWnd, WM_GETFONT, 0, 0);
					GetObject(m_hFont, sizeof(lf), &lf);
					lf.lfUnderline = TRUE;
					m_hFont = CreateFontIndirect(&lf);
					::SendMessage(m_hWnd, WM_SETFONT, (WPARAM)m_hFont, MAKELPARAM(TRUE, 0));
				}
				// use underline font and visited/unvisited colors
				SelectFont(hDC, m_hFont);
				SetTextColor(hDC, m_crForeColor);
				SetBkMode(hDC, TRANSPARENT);

				// return hollow brush to preserve parent background color
				hBrush = GetStockBrush(HOLLOW_BRUSH);
				return (long)hBrush;
			}
		}
		break;

	case WM_LBUTTONDOWN:
		{
			if (m_hlLink.IsEmpty()) 
			{	
				// if URL/filename not set get it from window text.
				TCHAR szText[256];
				GetWindowText(m_hWnd, szText, 255);
				m_hlLink = szText;
			}

			if (!m_hlLink.IsEmpty())
			{
				// Call ShellExecute to run the file.
				// For an URL, this means opening it in the browser.
				//
				HINSTANCE hInstance = m_hlLink.Navigate();
				if ((UINT)hInstance > 32) 
				{   
					// success!
					m_crForeColor = g_crVisited; 
					InvalidateRect(NULL, FALSE); 
				} 
				else 
					// unable to execute file!
					MessageBeep(0);		
			}
		}
		break;

	case WM_SETCURSOR:
		{
			if (NULL == g_hCursorLink) 
			{
				static bool bTriedOnce = FALSE;
				if (!bTriedOnce) 
				{
					TCHAR szWinDir[MAX_PATH];
					GetWindowsDirectory(szWinDir, MAX_PATH);
					lstrcat(szWinDir, _T("\\winhlp32.exe"));
					HMODULE hModule = LoadLibrary(szWinDir);
					if (hModule) 
						g_hCursorLink =	CopyCursor(::LoadCursor(hModule, MAKEINTRESOURCE(106)));
					FreeLibrary(hModule);
					bTriedOnce = TRUE;
				}
			}
			if (g_hCursorLink) 
			{
				::SetCursor(g_hCursorLink);
				return TRUE;
			}
			return FALSE;
		}
		break;

	case WM_DESTROY:
		if (NULL != m_hFont)
			DeleteFont(m_hFont);
		UnsubClass();
		break;
	}

	return FWnd::WindowProc(msg, wParam, lParam);
}

//
// Class CVersionInfo
//
typedef DWORD (APIENTRY *PFN_GetFileVersionInfoSizeA)(LPSTR lptstrFilename,LPDWORD lpdwHandle);
typedef BOOL (APIENTRY *PFN_GetFileVersionInfoA)(LPSTR lptstrFilename,DWORD dwHandle,DWORD dwLen,LPVOID lpData);              
typedef BOOL (APIENTRY *PFN_VerQueryValueA)(const LPVOID pBlock,LPSTR lpSubBlock,LPVOID * lplpBuffer,PUINT puLen);


CVersionInfo::CVersionInfo(HINSTANCE hInstance, LPCSTR szName)
{
	const int nLength = 256;
	TCHAR szModuleName[nLength];

	DWORD dwLen = GetModuleFileName(hInstance, szModuleName, nLength-1);
	if (dwLen <= 0)
		return;

	HMODULE hMod = LoadLibrary(_T("version.dll"));
	if (!hMod)
		return;

	PFN_GetFileVersionInfoSizeA pfnGetFileVersionInfoSizeA = (PFN_GetFileVersionInfoSizeA) GetProcAddress(hMod,"GetFileVersionInfoSizeA");
	if (NULL == pfnGetFileVersionInfoSizeA)
	{
		FreeLibrary(hMod);
		return;
	}

	DWORD dwDummy;
	dwLen = pfnGetFileVersionInfoSizeA(szModuleName, &dwDummy);
	if (dwLen <= 0)
	{
		FreeLibrary(hMod);
		return;
	}

	PFN_GetFileVersionInfoA pfnGetFileVersionInfoA = (PFN_GetFileVersionInfoA)GetProcAddress(hMod,"GetFileVersionInfoA");
	if (NULL == pfnGetFileVersionInfoA)
	{
		FreeLibrary(hMod);
		return;
	}

	m_pVersionInfo = new BYTE[dwLen];
	if (!pfnGetFileVersionInfoA(szModuleName, 0, dwLen, m_pVersionInfo))
	{
		FreeLibrary(hMod);
		return;
	}

	LPVOID pVoid;
	UINT nLen;
	PFN_VerQueryValueA pfnVerQueryValueA = (PFN_VerQueryValueA)GetProcAddress(hMod,"VerQueryValueA");
	if (NULL == pfnVerQueryValueA)
	{
		FreeLibrary(hMod);
		return;
	}

	if (!pfnVerQueryValueA(m_pVersionInfo, _T("\\"), &pVoid, &nLen))
	{
		FreeLibrary(hMod);
		return;
	}
	
	*(VS_FIXEDFILEINFO*)this = *(VS_FIXEDFILEINFO*)pVoid;
	
	// Get translation info
	if (pfnVerQueryValueA(m_pVersionInfo, _T("\\VarFileInfo\\Translation"), &pVoid, &nLen) && nLen >= 4)
		m_tLanguage = *(TRANSLATION*)pVoid;

	FreeLibrary(hMod);
}

CVersionInfo::~CVersionInfo()
{
	delete [] m_pVersionInfo;
}

LPCTSTR CVersionInfo::GetValue(LPCTSTR szValue)
{
	if (NULL == m_pVersionInfo)
		return NULL;

	TCHAR szTemp[256];
	wsprintf(szTemp, 
			 _T("\\StringFileInfo\\%04x%04x\\%s"), 
			 m_tLanguage.nLangId,
			 m_tLanguage.nCharSet,
			 szValue);

	LPVOID pVoid = NULL;
	UINT nLen;

	HMODULE hMod = LoadLibrary("version.dll");
	if (!hMod)
		return NULL;

	PFN_VerQueryValueA pfnVerQueryValueA = (PFN_VerQueryValueA)GetProcAddress(hMod,"VerQueryValueA");
	if (NULL == pfnVerQueryValueA)
	{
		FreeLibrary(hMod);
		return NULL;
	}

	if (!pfnVerQueryValueA(m_pVersionInfo, szTemp, &pVoid, &nLen))
		pVoid = NULL;

	FreeLibrary(hMod);
	return (LPTSTR)pVoid;
}

BOOL CVersionInfo::SetLabel(HWND hWnd, UINT nDlgItemId, LPCTSTR szPrompt, LPCTSTR szValue)
{
	LPTSTR szResult = (LPTSTR)GetValue(szValue);
	LPTSTR ptr = szResult;
	while (*ptr)
	{
		if (*ptr==',') 
			*ptr='.';
		else if (*ptr==32)
			lstrcpy(ptr,ptr+1);
		else
			++ptr;
	}
	TCHAR szTemp[256];
	wsprintf(szTemp, 
			 _T("%s %s"), 
			 szPrompt,
			 szResult);
	BOOL bResult = SetDlgItemText(hWnd, nDlgItemId, szTemp);
	return bResult;
}