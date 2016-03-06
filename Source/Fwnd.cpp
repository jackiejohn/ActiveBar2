#include "precomp.h"
#include "Debug.h"
#include "FWnd.h"

//
//  Copyright (c) 1995-1996, Data Dynamics. All rights reserved.
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

extern HINSTANCE g_hInstance;
static FWnd* s_pLastCreatedWnd = NULL;

TypedMap<HWND, FWnd*>* g_pmapWindow;

void FWnd::Init()
{
	g_pmapWindow = new TypedMap<HWND, FWnd*>;
}

void FWnd::CleanUp()
{
	delete g_pmapWindow;
}

//
// FWnd
//

FWnd::FWnd()
	: m_hWnd(NULL),
	  m_szClassName(NULL)
{
	m_bSubClassed = FALSE;
}

FWnd::~FWnd()
{
}

int FWnd::CheckMap()
{
	int nResult = g_pmapWindow->GetCount();
#ifdef _DEBUG
	assert(0 == nResult);
	HWND hWnd;
	FWnd* pWnd;
	FPOSITION posMap = g_pmapWindow->GetStartPosition();
	while (posMap)
	{
		try
		{
			g_pmapWindow->GetNextAssoc(posMap, hWnd, pWnd);
			TRACE2(10, _T("hWnd: %X, pWnd: %X, "), hWnd, pWnd);
			TRACE1(10, _T("Class Name: %s\n"), pWnd->m_szClassName);
		}
		catch (...)
		{
			g_pmapWindow->RemoveKey(hWnd);
			assert(FALSE);
		}
	}
#endif
	g_pmapWindow->RemoveAll();
	return nResult;
}

LRESULT CALLBACK FWnd::FWndWindowProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	try
	{
		FWnd* pWnd;
		if (!g_pmapWindow->Lookup(hWnd, pWnd))
		{ 
			if (WM_NCCREATE == message)
			{
				// Creating a new window so enter into window map
				pWnd = s_pLastCreatedWnd;
				pWnd->m_hWnd = hWnd;
				g_pmapWindow->SetAt(hWnd, pWnd);
				s_pLastCreatedWnd = NULL;
			}
	#ifdef _DEBUG
			else 
			{
				TCHAR szError[80];
				wsprintf(szError, _T("FWnd::Map Error, Window Message: %X\n"), message);
				TRACE(1, szError);
				return 0;
			}
	#else
			else
				return -1;
	#endif		
		}
		return pWnd->WindowProc(message, wParam, lParam);
	}
	catch (SEException& e)
	{
		e.ReportException(__FILE__, __LINE__);
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
	return m_hWnd;
}

LRESULT FWnd::WindowProc(UINT message,WPARAM wParam,LPARAM lParam)
{
	try
	{
		switch(message)
		{
		case WM_CREATE:
			TRACE2(4, "FWnd Created, Window Class: %s, HWND: %X,", m_szClassName, m_hWnd);
			TRACE1(4, " this: %X \n", this);
			return 0;

		case WM_NCDESTROY:
			// Remove from map
			g_pmapWindow->RemoveKey(m_hWnd);
			TRACE2(4, "FWnd Destroyed, Window Class: %s, hWnd: %X", m_szClassName, m_hWnd);
			TRACE1(4, " this: %X \n", this);
			return 0;

		default:
			break;
		}

		if (m_bSubClassed)
			return CallWindowProc((WNDPROC)m_wpWindowProc, m_hWnd, message, wParam, lParam);
		else
			return DefWindowProc(m_hWnd,message,wParam,lParam);
	}
	catch (SEException& e)
	{
		e.ReportException(__FILE__, __LINE__);
		return DefWindowProc(m_hWnd,message,wParam,lParam);
	}
}

void FWnd::RemoveFromMap()
{
	g_pmapWindow->RemoveKey(m_hWnd);
}

void FWnd::SubClassAttach(HWND hWnd)
{
	m_bSubClassed = TRUE;
	if (NULL != hWnd)
		m_hWnd = hWnd;
	g_pmapWindow->SetAt(m_hWnd, this);
	m_wpWindowProc = (WNDPROC)GetWindowLong(m_hWnd, GWL_WNDPROC);
	SetWindowLong(m_hWnd, GWL_WNDPROC, (long)FWndWindowProc);
}

void FWnd::UnsubClass()
{
	m_bSubClassed = FALSE;
	SetWindowLong(m_hWnd, GWL_WNDPROC, (long)m_wpWindowProc);
	g_pmapWindow->RemoveKey(m_hWnd);
}

