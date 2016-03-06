//
//  Copyright (c) 1995-1996, Data Dynamics. All rights reserved.
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//

#include "precomp.h"
#include "fwnd.h"
#include "debug.h"
#include "Support.h"
#include "GDIUtil.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

extern HINSTANCE g_hInstance;

static FMap s_mapWindow;
static FWnd* s_pLastCreatedWnd = NULL;

FWnd::FWnd()
	: m_hWnd(NULL),
	  m_szClassName(NULL)
{
	m_bSubClassed = FALSE;
}

FWnd::~FWnd()
{
}

LRESULT CALLBACK FWndWindowProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	try
	{
		FWnd* pWnd;
		if (!s_mapWindow.Lookup(hWnd, (LPVOID&)pWnd))
		{ 
			if (WM_NCCREATE == message)
			{
				// Creating a new window so enter into window map
				s_pLastCreatedWnd->m_hWnd = hWnd;
				s_mapWindow.SetAt(hWnd, s_pLastCreatedWnd);
				pWnd = s_pLastCreatedWnd;
			}
	#ifdef _DEBUG
			else 
			{
				char str[80];
				wsprintf(str, "Fatal : Window Map error %X\n", message);
				TRACE(1, str);
				return 0;
			}
	#else
			else
				return -1;
	#endif		
		}
		return pWnd->WindowProc(message, wParam, lParam);
	}
	catch (...)
	{
		assert(FALSE);
		return DefWindowProc(hWnd,message,wParam,lParam);
	}
}

BOOL FWnd::RegisterWindow(LPCTSTR szClassName,
						  UINT    nStyle,
						  HBRUSH  hbBackground,
						  HCURSOR hCursor,
						  HICON   hIcon,
						  LPCTSTR szMenuName)
{
	m_szClassName = szClassName;
	WNDCLASS wcInfo;
	if (::GetClassInfo(g_hInstance, szClassName, &wcInfo))
        return TRUE;

	wcInfo.style = nStyle;
	wcInfo.lpfnWndProc = FWndWindowProc;
	wcInfo.cbClsExtra = 0;
	wcInfo.cbWndExtra = 0;
	wcInfo.hInstance = g_hInstance;
	wcInfo.hIcon = hIcon;
	wcInfo.hCursor = hCursor;
	wcInfo.hbrBackground = hbBackground;
	wcInfo.lpszMenuName = szMenuName;
	wcInfo.lpszClassName = szClassName;
	if (0 == RegisterClass(&wcInfo))
		return FALSE;
	return TRUE;
}

HWND FWnd::Create(LPCTSTR szWindowName,
				  DWORD   dwStyle,
				  int	  nLeft,
				  int	  nTop,
				  int	  nWidth,	
				  int	  nHeight,
   				  HWND    hWndParent,
				  HMENU   hMenu,
				  LPVOID  lParam)
{
	s_pLastCreatedWnd = this;
	m_hWnd = CreateWindow(m_szClassName, 
						  szWindowName, 
						  dwStyle, 
						  nLeft, 
						  nTop, 
						  nWidth, 
						  nHeight, 
						  hWndParent, 
						  hMenu, 
						  g_hInstance, 
						  lParam);
	if (NULL == m_hWnd)
	{
		LPTSTR szMsgBuf = NULL;
		FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER|FORMAT_MESSAGE_FROM_SYSTEM,
				      NULL,
					  GetLastError(),
					  MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
					  szMsgBuf,    
					  0,    
					  NULL);
		TRACE(1, szMsgBuf);
		LocalFree(szMsgBuf);
	}
	return m_hWnd;
}

HWND FWnd::MDIChildCreate(LPTSTR szWindowName,
						  DWORD  dwStyle,
						  int	 nLeft,
						  int	 nTop,
						  int	 nWidth,	
						  int	 nHeight,
   						  HWND   hWndParent,
						  LPARAM lParam)
{
	s_pLastCreatedWnd = this;
	m_hWnd = CreateMDIWindow((LPTSTR)m_szClassName,
						     szWindowName,
						     dwStyle,
							 nLeft,
							 nTop,
							 nWidth,
							 nHeight,
							 hWndParent,
						     g_hInstance,
						     lParam);
	if (NULL == m_hWnd)
	{
		LPTSTR szMsgBuf = NULL;
		FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER|FORMAT_MESSAGE_FROM_SYSTEM,
				      NULL,
					  GetLastError(),
					  MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
					  szMsgBuf,    
					  0,    
					  NULL);
		TRACE(1, szMsgBuf);
		LocalFree(szMsgBuf);
	}
	return m_hWnd;
}

HWND FWnd::CreateEx(DWORD   dwExStyle,
					LPCTSTR szWindowName,
					DWORD   dwStyle,
					int	    nLeft,
					int	    nTop,
					int	    nWidth,	
					int	    nHeight,
   					HWND    hWndParent,
					HMENU   hMenu,
					LPVOID  lParam)
{
	s_pLastCreatedWnd = this;
	m_hWnd=CreateWindowEx(dwExStyle,
						  m_szClassName,
						  szWindowName,
						  dwStyle,
						  nLeft, 
						  nTop, 
						  nWidth, 
						  nHeight, 
						  hWndParent,
						  hMenu,
						  g_hInstance,
						  lParam);
	if (NULL == m_hWnd)
	{
		LPTSTR szMsgBuf = NULL;
		FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER|FORMAT_MESSAGE_FROM_SYSTEM,
				      NULL,
					  GetLastError(),
					  MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
					  szMsgBuf,    
					  0,    
					  NULL);
		TRACE(1, szMsgBuf);
		LocalFree(szMsgBuf);
	}
	return m_hWnd;
}

#ifdef _AFXDLL
#define CALLWINPROCTYPE WNDPROC
#else
#define CALLWINPROCTYPE FARPROC
#endif

LRESULT FWnd::WindowProc(UINT message,WPARAM wParam,LPARAM lParam)
{ 
	switch(message)
	{
	case WM_CREATE:
		TRACE2(1, "FWnd Created, Window Class: %s, HWND: %X,", m_szClassName, m_hWnd);
		TRACE1(1, " this: %X \n", this);
		return 0;

	case WM_NCDESTROY:
		// Remove from map
		s_mapWindow.RemoveKey(m_hWnd);
		TRACE2(1, "FWnd Destroyed, Window Class: %s, hWnd: %X", m_szClassName, m_hWnd);
		TRACE1(1, " this: %X \n", this);
		return 0;

	default:
		break;
	}
	if (m_bSubClassed)
		return CallWindowProc((WNDPROC)m_wpWindowProc, m_hWnd, message, wParam, lParam);
	else
		return DefWindowProc(m_hWnd,message,wParam,lParam);
}

void FWnd::SubClassAttach(HWND hWnd)
{
	m_bSubClassed = TRUE;
	if (0 != hWnd)
		m_hWnd = hWnd;
	s_mapWindow.SetAt(m_hWnd, this);
	m_wpWindowProc = (WNDPROC)GetWindowLong(m_hWnd, GWL_WNDPROC);
	SetWindowLong(m_hWnd, GWL_WNDPROC, (long)FWndWindowProc);
}

void FWnd::UnsubClass()
{
	m_bSubClassed = FALSE;
	s_mapWindow.RemoveKey(m_hWnd);
	if (IsWindow(m_hWnd))
		SetWindowLong(m_hWnd, GWL_WNDPROC, (long)m_wpWindowProc);
}

//
// Shared small nice font
//

static HFONT s_hFontShared = NULL;
static int s_nFontSharedHeight = 0;

HFONT FGetFont()
{
	if (NULL != s_hFontShared)
		return s_hFontShared;

	LOGFONT lf;
	memset(&lf, 0, sizeof(LOGFONT));
	lstrcpy(lf.lfFaceName, _T("MS Sans Serif"));
	lf.lfHeight = -8;
	s_hFontShared = CreateFontIndirect(&lf);
	if (NULL == s_hFontShared)
		return NULL;

	HDC hDC = GetDC(NULL);
	if (NULL == hDC)
		return NULL;

	HFONT hFontOld = SelectFont(hDC, s_hFontShared);
	if (NULL == s_hFontShared)
	{
		ReleaseDC(NULL, hDC);
		return NULL;
	}

	TEXTMETRIC tm;
	GetTextMetrics(hDC, &tm);
	s_nFontSharedHeight = tm.tmHeight;
	SelectFont(hDC, hFontOld);
	
	ReleaseDC(NULL, hDC);
	return s_hFontShared;
}
