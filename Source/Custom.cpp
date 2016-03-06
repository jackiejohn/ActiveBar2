//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include <olectl.h>
#include <stddef.h>       // for offsetof()
#include "debug.h"
#include "Globals.h"
#include "FDialog.h"
#include "Support.h"
#include "IpServer.h"
#include "Resource.h"
#include "DropSource.h"
#include "Tool.h"
#include "ChildBands.h"
#include "Band.h"
#include "Bands.h"
#include "Localizer.h"
#include "ShortCuts.h"
#include "Popupwin.h"
#include "CustomizeListbox.h"
#include "Bar.h"
#include "Custom.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif
#define DC_GRADIENT 0x0020

extern BOOL ConnectionAdvise(LPUNKNOWN pUnkSrc, REFIID iid,LPUNKNOWN pUnkSink, BOOL bRefCount, DWORD* pdwCookie);
extern BOOL ConnectionUnadvise(LPUNKNOWN pUnkSrc, REFIID iid,LPUNKNOWN pUnkSink, BOOL bRefCount, DWORD dwCookie);

//
// CenterWindow
//

static void CenterWindow(HWND hWnd)
{
	try
	{
		HWND hWndCenter;

		// determine owner window to center against
		DWORD dwStyle = GetWindowLong(hWnd, GWL_STYLE);
		if (dwStyle & WS_CHILD)
			hWndCenter = ::GetParent(hWnd);
		else
			hWndCenter = ::GetWindow(hWnd, GW_OWNER);

		// get coordinates of the window relative to its parent
		CRect rcDlg;
		GetWindowRect(hWnd, &rcDlg);
		CRect rcArea;
		CRect rcCenter;
		HWND hWndParent;
		if (!(dwStyle & WS_CHILD))
		{
			// don't center against invisible or minimized windows
			if (NULL != hWndCenter)
			{
				DWORD dwStyle = ::GetWindowLong(hWndCenter, GWL_STYLE);
				if (!(dwStyle & WS_VISIBLE) || (dwStyle & WS_MINIMIZE))
					hWndCenter = 0;
			}

			// center within screen coordinates
			SystemParametersInfo(SPI_GETWORKAREA, 0, &rcArea, 0);
			if (0 == hWndCenter)
				rcCenter = rcArea;
			else
				::GetWindowRect(hWndCenter, &rcCenter);
		}
		else
		{
			// center within parent client coordinates
			hWndParent = ::GetParent(hWnd);
			assert(::IsWindow(hWndParent));

			::GetClientRect(hWndParent, &rcArea);
			::GetClientRect(hWndCenter, &rcCenter);
			::MapWindowPoints(hWndCenter, hWndParent, (POINT*)&rcCenter, 2);
		}

		// find dialog's upper left based on rectCenter
		int xLeft = (rcCenter.left + rcCenter.right) / 2 - rcDlg.Width() / 2;
		int yTop = (rcCenter.top + rcCenter.bottom) / 2 - rcDlg.Height() / 2;

		// if the dialog is outside the screen, move it inside
		if (xLeft < rcArea.left)
			xLeft = rcArea.left;
		else if (xLeft + rcDlg.Width() > rcArea.right)
			xLeft = rcArea.right - rcDlg.Width();

		if (yTop < rcArea.top)
			yTop = rcArea.top;
		else if (yTop + rcDlg.Height() > rcArea.bottom)
			yTop = rcArea.bottom - rcDlg.Height();

		// map screen coordinates to child coordinates
		SetWindowPos(hWnd, 
					 0, 
					 xLeft, 
					 yTop, 
					 -1, 
					 -1,
					 SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// CDDPropertyPage
//

TypedMap<HWND, CDDPropertyPage*>* CDDPropertyPage::m_pmapPages = NULL;

void CDDPropertyPage::Init()
{
	m_pmapPages = new TypedMap<HWND, CDDPropertyPage*>;
}

void CDDPropertyPage::Cleanup()
{
	delete m_pmapPages;
}

//
// PropertyPageProc
//

BOOL CDDPropertyPage::PropertyPageProc(HWND hWndPage, UINT nMsg, UINT wParam, LONG lParam)
{
	try
	{
		CDDPropertyPage* pPropPage;
		if (!m_pmapPages->Lookup(hWndPage, pPropPage) && WM_INITDIALOG != nMsg)
			return FALSE;

		switch (nMsg)
		{
		case WM_INITDIALOG:
			{
				CDDPropertyPage* pPropPage = (CDDPropertyPage*)((PROPSHEETPAGE*)lParam)->lParam;
				if (pPropPage)
					m_pmapPages->SetAt(hWndPage, pPropPage);
				pPropPage->hWnd(hWndPage);
				pPropPage->m_pSheet->hWnd(GetParent(hWndPage));
				pPropPage->ExecuteDlgInit();
				return pPropPage->OnPageInit();
			}
			break;

		case WM_NCDESTROY:
			m_pmapPages->RemoveKey(hWndPage);
			break;

		case WM_DESTROY:
			return pPropPage->OnDestroy();

		case WM_NOTIFY:
			{
				if (PSN_SETACTIVE == wParam)
					CenterWindow(GetParent(pPropPage->hWnd()));
				BOOL bResult = TRUE;
				BOOL bResult2 = pPropPage->OnNotify((int)wParam, (LPNMHDR)lParam, bResult);
				SetWindowLong(hWndPage, DWL_MSGRESULT, bResult);
				return bResult2;
			}

		case WM_COMMAND:
			return pPropPage->OnCommand(HIWORD(wParam), LOWORD(wParam), (HWND)lParam);

		default:
			return pPropPage->OnMessage(nMsg, wParam, lParam);
		}
		return TRUE;   
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
}

//
// CDDPropertyPage
//

CDDPropertyPage::CDDPropertyPage(const UINT& nId)
{
	const int nSize = sizeof(PROPSHEETPAGE);
	memset (&m_pspInfo, '\0', nSize);
    m_pspInfo.dwSize      = nSize; 
    m_pspInfo.pszTemplate = MAKEINTRESOURCE(nId); 
    m_pspInfo.hInstance   = g_hInstance;
    m_pspInfo.pfnDlgProc  = (DLGPROC)PropertyPageProc; 
    m_pspInfo.lParam = (LPARAM)this; 
}

//
// ExecuteDlgInit
//

BOOL CDDPropertyPage::ExecuteDlgInit()
{
	try
	{
		// find resource handle
		LPVOID lpResource = NULL;
		HGLOBAL hResource = NULL;
		HRSRC hDlgInit = ::FindResource(g_hInstance, m_pspInfo.pszTemplate, MAKEINTRESOURCE(240));
		if (hDlgInit)
		{
			// load it
			hResource = LoadResource(g_hInstance, hDlgInit);
			if (hResource == NULL)
				return FALSE;
			// lock it
			lpResource = LockResource(hResource);
			assert(lpResource != NULL);
		}

		// execute it
		BOOL bResult = ExecuteDlgInit(lpResource);

		// cleanup
		if (lpResource != NULL)
			UnlockResource(hResource);
		if (hResource != NULL)
			FreeResource(hResource);
		return bResult;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
}

//
// ExecuteDlgInit
//

BOOL CDDPropertyPage::ExecuteDlgInit(LPVOID pResource)
{
	try
	{
		BOOL bSuccess = TRUE;
		if (pResource)
		{
			UNALIGNED WORD* lpnRes = (WORD*)pResource;
			while (bSuccess && *lpnRes)
			{
				WORD nIDC = *lpnRes++;
				WORD nMsg = *lpnRes++;
				DWORD dwLen = *((UNALIGNED DWORD*&)lpnRes)++;
				// In Win32 the WM_ messages have changed.  They have
				// to be translated from the 32-bit values to 16-bit
				// values here.

				#define WIN16_LB_ADDSTRING  0x0401
				#define WIN16_CB_ADDSTRING  0x0403

				if (nMsg == WIN16_LB_ADDSTRING)
					nMsg = LB_ADDSTRING;
				else if (nMsg == WIN16_CB_ADDSTRING)
					nMsg = CB_ADDSTRING;

				// check for invalid/unknown message types
				assert(nMsg == LB_ADDSTRING || nMsg == CB_ADDSTRING);

#ifdef _DEBUG
				// For AddStrings, the count must exactly delimit the
				// string, including the NULL termination.  This check
				// will not catch all mal-formed ADDSTRINGs, but will
				// catch some.
				if (nMsg == LB_ADDSTRING || nMsg == CB_ADDSTRING)
					assert(*((LPBYTE)lpnRes + (UINT)dwLen - 1) == 0);
#endif

				if (nMsg == LB_ADDSTRING || nMsg == CB_ADDSTRING)
				{
					// List/Combobox returns -1 for error
					if (::SendDlgItemMessageA(m_hWnd, nIDC, nMsg, 0, (LONG)lpnRes) == -1)
						bSuccess = FALSE;
				}

				// skip past data
				lpnRes = (WORD*)((LPBYTE)lpnRes + (UINT)dwLen);
			}
		}
		return bSuccess;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
}

//
// CDDPropertySheet
//

//
// MessageProc
//

LRESULT CALLBACK CDDPropertySheet::MessageProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	try
	{
		if (nCode < 0)
			return ::CallNextHookEx(m_hMessageHookProc, nCode, wParam, lParam);

		BOOL bResult;
		int nSize = m_paSheets->GetSize(); 
		for (int nIndex = 0; nIndex < nSize; nIndex++)
		{
			CDDPropertySheet* pSheet = m_paSheets->GetAt(nIndex);
			bResult = pSheet->PreTranslateMessage(wParam, lParam);
			if (bResult)
				return bResult;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return CallNextHookEx(m_hMessageHookProc, nCode, wParam, lParam);
}

//
// CDDPropertySheet
//

CDDPropertySheet::CDDPropertySheet(HWND hWndParent)
{
	try
	{
		const int nSize = sizeof(PROPSHEETHEADER);
		memset (&m_pshInfo, '\0', nSize);
		m_pshInfo.dwSize = nSize; 
		m_pshInfo.hInstance = g_hInstance;
		m_pshInfo.hwndParent = hWndParent; 
		m_pshInfo.pszCaption = NULL;
		m_pshInfo.dwFlags = PSH_DEFAULT;
		m_pshInfo.dwFlags |= PSH_MODELESS;
		m_paSheets->Add(this);
		if (NULL == m_hMessageHookProc)
			m_hMessageHookProc = SetWindowsHookEx(WH_KEYBOARD, (HOOKPROC)MessageProc, 0, GetCurrentThreadId());
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// ~CDDPropertySheet
//

CDDPropertySheet::~CDDPropertySheet()
{
	try
	{
		int nSize = m_paSheets->GetSize(); 
		if (1 == nSize && m_hMessageHookProc)
		{
			UnhookWindowsHookEx(m_hMessageHookProc);
			m_hMessageHookProc = NULL;
		}

		for (int nIndex = 0; nIndex < nSize; nIndex++)
		{
			if (this == m_paSheets->GetAt(nIndex))
			{
				m_paSheets->RemoveAt(nIndex);
				break;
			}
		}
		if (PSH_MODELESS & m_pshInfo.dwFlags)
			delete [] m_pshInfo.phpage;
		m_paSheets->RemoveAll();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

extern HINSTANCE g_hInstance;
BOOL CDDPropertySheet::m_bHookSet = FALSE;
TypedArray<CDDPropertySheet*>* CDDPropertySheet::m_paSheets;
HHOOK  CDDPropertySheet::m_hMessageHookProc = NULL;
WNDPROC CDDPropertySheet::m_wpDialogProc = NULL;
static CDDPropertySheet* s_pLastCreatedSheetd = NULL;
//TypedMap<HWND, CDDPropertySheet*>* g_pmapPropertySheet;

void CDDPropertySheet::Init()
{
	CDDPropertyPage::Init();
	m_paSheets = new TypedArray<CDDPropertySheet*>;
	assert(m_paSheets);
//	g_pmapPropertySheet = new TypedMap<HWND, CDDPropertySheet*>;
}

void CDDPropertySheet::Cleanup()
{
	CDDPropertyPage::Cleanup();
	delete m_paSheets;
//	delete g_pmapPropertySheet;
}

void CDDPropertySheet::OnMinimize(BOOL bActivate)
{
}

void DrawFrameRect(HDC hDC, CRect rc, int nBorderCX, int nBorderCY)
{
	COLORREF crBackground = GetSysColor(COLOR_BTNFACE);
	RECT rcLine = rc;
	rcLine.right = rcLine.left + 1;
	FillSolidRect(hDC, rcLine, crBackground);
	rcLine = rc;
	rcLine.bottom = rcLine.top + 1;
	FillSolidRect(hDC, rcLine, crBackground);
			
	// Right and bottom
	SetBkColor(hDC,GetSysColor(COLOR_WINDOWFRAME));
	rcLine = rc;
	rcLine.left = rcLine.right - 1;
	ExtTextOut(hDC, 0, 0, ETO_OPAQUE, &rcLine, NULL, 0, NULL);
	rcLine = rc;
	rcLine.top = rcLine.bottom - 1;
	ExtTextOut(hDC, 0, 0, ETO_OPAQUE, &rcLine, NULL, 0, NULL);
	
	rc.Inflate(-1, -1);

	// Inner left and top
	COLORREF crHighLight = GetSysColor(COLOR_BTNHIGHLIGHT);
	rcLine = rc;
	rcLine.right = rcLine.left + 1;
	FillSolidRect(hDC, rcLine, crHighLight);
	rcLine = rc;
	rcLine.bottom = rcLine.top + 1;
	FillSolidRect(hDC, rcLine, crHighLight);

	// Right and bottom
	COLORREF crShadow = GetSysColor(COLOR_BTNSHADOW);
	rcLine = rc;
	rcLine.left = rcLine.right - 1;
	FillSolidRect(hDC, rcLine, crShadow);
	rcLine = rc;
	rcLine.top = rcLine.bottom - 1;
	FillSolidRect(hDC, rcLine, crShadow);
	
	rc.Inflate(-1, -1);
	nBorderCX -= 2;	
	
	if (nBorderCX > 0)
	{
		// left and right
		RECT rcLine = rc;
		rcLine.right = rcLine.left + nBorderCX;
		FillSolidRect(hDC, rcLine, crBackground);
		rcLine = rc;
		rcLine.left = rcLine.right - nBorderCX;
		FillSolidRect(hDC, rcLine, crBackground);
	}

	nBorderCY -= 2;
	if (nBorderCY > 0)
	{
		// top and bottom
		rcLine = rc;
		rcLine.bottom = rcLine.top + nBorderCY;
		FillSolidRect(hDC, rcLine, crBackground);
		rcLine = rc;
		rcLine.top = rcLine.bottom - nBorderCY;
		FillSolidRect(hDC, rcLine, crBackground);
	}
}

//
// DialogProc
//

BOOL CALLBACK CDDPropertySheet::DialogProc(HWND hWnd, UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		switch (nMsg)
		{
/*		case WM_CREATE:
			TRACE(5, "WM_CREATE\n");
			break;

		case WM_DESTROY:
			g_pmapPropertySheet->RemoveKey(hWnd);
			break;
*/
		case WM_HELP:
			{
				LPHELPINFO pHelpInfo = (LPHELPINFO)lParam;
				pHelpInfo->hItemHandle = WindowFromPoint(pHelpInfo->MousePos);
				pHelpInfo->iCtrlId = GetWindowLong((HWND)pHelpInfo->hItemHandle, GWL_ID);
				pHelpInfo->dwContextId = 0; 
				int nSize = m_paSheets->GetSize(); 
				for (int nIndex = 0; nIndex < nSize; nIndex++)
				{
					if (m_paSheets->GetAt(nIndex)->hWnd() == hWnd)
						((CCustomize*)m_paSheets->GetAt(nIndex))->Help(pHelpInfo);
				}
			}
			return TRUE;

/*		case WM_NCACTIVATE:
			TRACE(5, "WM_NCACTIVATE\n");
			return FALSE;

		case WM_MOUSEACTIVATE:
			{
				HWND hWndActive = GetForegroundWindow();
				BOOL bResult;
				if (NULL == hWndActive || (hWnd != hWndActive && !IsDescendant(hWndActive, hWnd)))
					bResult = SetForegroundWindow(GetParent(hWnd));
			}
			return MA_NOACTIVATE;

		case WM_NCPAINT:
			{
				//
				// Painting the windows caption area
				//

				CRect rc;
				GetWindowRect(hWnd, &rc);
				HDC hDC = GetWindowDC(hWnd);
				if (hDC)
				{
					SIZE sizeBorder;
					sizeBorder.cx = GetSystemMetrics(SM_CXFRAME);
					sizeBorder.cy = GetSystemMetrics(SM_CYFRAME);
					rc.Offset(-rc.left, -rc.top);

					DrawFrameRect(hDC, rc, sizeBorder.cx, sizeBorder.cy);

					rc.Inflate(-sizeBorder.cx, -sizeBorder.cy);
					
					int nCaptionHeight = GetSystemMetrics(SM_CYSMCAPTION);
					rc.bottom = rc.top + nCaptionHeight - 1;
					UINT uFlags = DC_TEXT | DC_SMALLCAP;
					uFlags |= DC_ACTIVE;
					if (GetGlobals().m_nBitDepth > 8)
						uFlags |= DC_GRADIENT;
					DrawCaption(hWnd, hDC, &rc, uFlags);

					// Space below caption
					CRect rcTemp = rc; 
					rcTemp.top = rcTemp.bottom;
					rcTemp.bottom = rcTemp.top + 1;
					
					FillSolidRect(hDC, rcTemp, GetSysColor(COLOR_BTNFACE));

					//
					// Draw close box
					//

					rc.left = rc.right - nCaptionHeight - 1;
					rc.Inflate(-2, -2);
					
					// Still off a little so adjust it
					rc.left += 2;
					
					// Draw the close button
					DrawFrameControl(hDC, &rc, DFC_CAPTION, DFCS_CAPTIONCLOSE);
					ReleaseDC(hWnd, hDC);
				}
			}
			return 0;
*/
		case WM_ACTIVATEAPP:
			{
				BOOL bActivate = wParam;
				int nSize = m_paSheets->GetSize(); 
				for (int nIndex = 0; nIndex < nSize; nIndex++)
				{
					if (m_paSheets->GetAt(nIndex)->hWnd() == hWnd)
						((CCustomize*)m_paSheets->GetAt(nIndex))->OnMinimize(bActivate);
				}
			}
			break;

		case WM_SYSCOMMAND:
			switch (wParam)
			{
			case SC_CLOSE:
				{
					DestroyWindow(hWnd);
					int nSize = m_paSheets->GetSize(); 
					for (int nIndex = 0; nIndex < nSize; nIndex++)
					{
						if (m_paSheets->GetAt(nIndex)->hWnd() == hWnd)
							((CCustomize*)m_paSheets->GetAt(nIndex))->EndCustomize();
					}
				}
				break;
			}
			break;
		}
		return CallWindowProc((WNDPROC)m_wpDialogProc, hWnd, nMsg, wParam, lParam);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
}

//
// SubClassAttach
//

void CDDPropertySheet::SubClassAttach(HWND hWnd)
{
	m_wpDialogProc = (WNDPROC)GetWindowLong(m_hWnd, GWL_WNDPROC);
	assert(m_wpDialogProc);
	SetWindowLong(m_hWnd, GWL_WNDPROC, (long)DialogProc);
//	g_pmapPropertySheet->SetAt(hWnd, this);
}

//
// ShowSheet
//

HWND CDDPropertySheet::ShowSheet()
{
	try
	{
		m_pshInfo.nPages = m_faPages.GetSize();

		m_pshInfo.phpage = new HPROPSHEETPAGE[m_pshInfo.nPages]; 
		assert(m_pshInfo.phpage);
		if (NULL == m_pshInfo.phpage)
			return NULL;

		CDDPropertyPage* pPropPage;
		for (int nIndex = 0; nIndex < (int)m_pshInfo.nPages; nIndex++)
		{
			pPropPage = m_faPages.GetAt(nIndex);
			if (pPropPage)
				m_pshInfo.phpage[nIndex] = CreatePropertySheetPage(&pPropPage->PageInfo());
		}

		m_hWnd = (HWND)PropertySheet(&m_pshInfo);
		if (!(m_pshInfo.dwFlags & PSH_MODELESS))
			delete [] m_pshInfo.phpage;
		else if (m_hWnd)
			SubClassAttach(m_hWnd);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return NULL;
	}
	return m_hWnd;
}

//
// PreTranslateMessage
//

BOOL CDDPropertySheet::PreTranslateMessage(WPARAM wParam, LPARAM lParam)
{
	try
	{
		HWND hWndFocus = GetFocus();
		if (IsChild(m_hWnd, hWndFocus))
		{
			BOOL bAlt = (GetKeyState(VK_MENU) < 0);
			if (!bAlt)
				return FALSE;

			switch (wParam)
			{
			case 'B':
				PropSheet_SetCurSel(m_hWnd, 0, 0);
				return TRUE;

			case 'C':
				PropSheet_SetCurSel(m_hWnd, 0, 1);
				return TRUE;

			case 'O':
				PropSheet_SetCurSel(m_hWnd, 0, 2);
				return TRUE;

			case 'K':
				PostMessage(m_hWnd, WM_COMMAND, MAKELONG(CCustomize::eKeyboard, BN_CLICKED), (LPARAM)GetDlgItem(m_hWnd, CCustomize::eKeyboard));
				return TRUE;

			case 'N':
			case 'R':
				return FALSE;
			}
			return TRUE;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
	return FALSE;
}

//
// CLabel
//

CLabel::CLabel()
{
}

CLabel::~CLabel()
{
}

HFONT CLabel::m_hFont = NULL;

LRESULT _stdcall CLabel::WindowProc(HWND hWnd, UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		switch (nMsg)
		{
		case WM_SETFONT:
			m_hFont = (HFONT)wParam; 
			break;

		case WM_PAINT:
			{
				PAINTSTRUCT ps;
				HDC hDC = BeginPaint(hWnd, &ps);
				if (hDC)
				{
					CRect rcClient;
					::GetClientRect(hWnd, &rcClient);
				
					HFONT hFontOld = NULL;
					if (m_hFont)
						hFontOld = SelectFont(hDC, m_hFont);

					FillSolidRect(hDC, rcClient, GetSysColor(COLOR_3DFACE));

					TCHAR szBuffer[128];
					GetWindowText(hWnd, szBuffer, 128);

					CRect rcText;
					if (0 != DrawText(hDC, szBuffer, lstrlen(szBuffer), &rcText, DT_VCENTER|DT_LEFT|DT_CALCRECT))
					{
						if (0 != DrawText(hDC, szBuffer, lstrlen(szBuffer), &rcClient, DT_VCENTER|DT_LEFT))
							rcClient.left += rcText.Width() + 4;
					}

					rcClient.top = (rcClient.Height() - 2) / 2;
					rcClient.bottom = rcClient.top + 2;
					DrawEdge(hDC, &rcClient, EDGE_ETCHED, BF_TOP);

					if (m_hFont)
						SelectFont(hDC, hFontOld);
					EndPaint(hWnd, &ps);
					return 0;
				}
			}
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
	return DefWindowProc(hWnd,nMsg,wParam,lParam);
}

void CLabel::RegisterClass()
{
	WNDCLASS wcInfo;
	if (::GetClassInfo(g_hInstance, _T("DDLabel"), &wcInfo))
        return;

	wcInfo.style = CS_SAVEBITS;
	wcInfo.lpfnWndProc = CLabel::WindowProc;
	wcInfo.cbClsExtra = 0;
	wcInfo.cbWndExtra = 0;
	wcInfo.hInstance = g_hInstance;
	wcInfo.hIcon = NULL;
	wcInfo.hCursor = ::LoadCursor(NULL, IDC_ARROW);
	wcInfo.hbrBackground = NULL;
	wcInfo.lpszMenuName = NULL;
	wcInfo.lpszClassName = _T("DDLabel");
	::RegisterClass(&wcInfo);
}

//
// CKeyEdit
//

CKeyEdit::CKeyEdit(CBar* pBar)
{
	m_pShortCutStore = ShortCutStore::CreateInstance(NULL);
	m_pShortCutStore->m_pBar = pBar;
	assert(m_pShortCutStore);
}

CKeyEdit::~CKeyEdit()
{
	if (m_pShortCutStore)
		m_pShortCutStore->Release();
}

//
// Clear
//

void CKeyEdit::Clear()
{
	if (m_pShortCutStore)
		m_pShortCutStore->Clear();
}


//
// WindowProc
//

LRESULT CKeyEdit::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		switch (nMsg)
		{
		case WM_CHAR:
			return 0;
		
		case WM_CUT:
			return 0;

		case WM_COPY:
			return 0;

		case WM_PASTE:
			return 0;

		case WM_SYSKEYDOWN:
		case WM_KEYDOWN:
			{
				if (HIWORD(lParam) & KF_REPEAT)
					break;

				TCHAR szDesc[60];
				UINT nControlState = 0;

				if (GetKeyState(VK_MENU) < 0)
					nControlState |= FALT;

				if (GetKeyState(VK_CONTROL) < 0)
					nControlState |= FCONTROL;

				BOOL bShift = FALSE;
				if (GetKeyState(VK_SHIFT) < 0)
				{
					//nControlState |= FSHIFT;
					bShift = TRUE;
				}

				if (VK_BACK == wParam && 0 == nControlState)
				{
					if (m_pShortCutStore->scpV1.m_nKeyCodes[1])
					{
						m_pShortCutStore->scpV1.m_bShift[1] = FALSE;
						m_pShortCutStore->scpV1.m_nKeyCodes[1] = 0;
						m_pShortCutStore->scpV1.m_nKeyboardCodes[1] = 0;
					}
					else if (m_pShortCutStore->scpV1.m_nKeyCodes[0])
					{
						m_pShortCutStore->scpV1.m_bShift[0] = FALSE;
						m_pShortCutStore->scpV1.m_nKeyCodes[0] = 0;
						m_pShortCutStore->scpV1.m_nKeyboardCodes[0] = 0;
						m_pShortCutStore->scpV1.m_nControlKeys = 0;
					}
					if (m_pShortCutStore->GetShortCutDesc(szDesc))
						SendMessage(WM_SETTEXT, 0, (LPARAM)szDesc);
				}
				else
				{
					if (m_pShortCutStore->scpV1.m_nKeyCodes[1])
						MessageBeep(0xFFFFFFFF);
					else if (m_pShortCutStore->scpV1.m_nKeyCodes[0])
					{
						//
						// Valid Second Keys
						//
						if ((wParam >= VK_NUMPAD0 && wParam <= VK_F24) ||
							(wParam >= VK_INSERT && wParam <= VK_DELETE) ||
							(wParam >= VK_PRIOR && wParam <= VK_DOWN) ||
							(wParam >= 'A' && wParam <= 'Z') ||
							(196 == wParam || 197 == wParam || 214 == wParam) ||
							(wParam == VK_BACK) ||
							(wParam == VK_RETURN) ||
							(wParam == VK_SCROLL)
						   )
						{
							//
							// Add shift to these keys if it is pressed
							//

							if (bShift)
								m_pShortCutStore->scpV1.m_bShift[1] = TRUE;
							
							m_pShortCutStore->scpV1.m_nKeyCodes[1] = wParam;
							m_pShortCutStore->scpV1.m_nKeyboardCodes[1] = lParam;
							if (wParam >= VK_PRIOR && wParam <= VK_DOWN && !(GetKeyState(VK_NUMLOCK) & 0x01))
								m_pShortCutStore->scpV1.m_nKeyboardCodes[1] |= 0x01000000;
							
							if (m_pShortCutStore->GetShortCutDesc(szDesc))
								SendMessage(WM_SETTEXT, 0, (LPARAM)szDesc);
						}
						else if ((wParam >= ' ' && wParam <= '@') ||
							     (wParam >= '[' && wParam <= '`') || 
							     (wParam >= '{' && wParam <= '~')
						        )
						{
							m_pShortCutStore->scpV1.m_nKeyCodes[1] = wParam;
							m_pShortCutStore->scpV1.m_nKeyboardCodes[1] = lParam;
							if (m_pShortCutStore->GetShortCutDesc(szDesc))
								SendMessage(WM_SETTEXT, 0, (LPARAM)szDesc);
						}
					}
					else if (0 == nControlState)
					{
						if ((wParam >= VK_F1 && wParam <= VK_F24) ||
							(wParam >= VK_INSERT && wParam <= VK_DELETE) ||
							(wParam >= VK_PRIOR && wParam <= VK_DOWN)
						   )
						{
							if (bShift)
								m_pShortCutStore->scpV1.m_bShift[0] = TRUE;
							
							m_pShortCutStore->scpV1.m_nKeyCodes[0] = wParam;
							m_pShortCutStore->scpV1.m_nKeyboardCodes[0] = lParam;
							if (wParam >= VK_PRIOR && wParam <= VK_DOWN && !(GetKeyState(VK_NUMLOCK) & 0x01))
								m_pShortCutStore->scpV1.m_nKeyboardCodes[0] |= 0x01000000;
							
							if (m_pShortCutStore->GetShortCutDesc(szDesc))
								SendMessage(WM_SETTEXT, 0, (LPARAM)szDesc);
						}
					}
					else
					{
						//
						// Valid First Keys with control keys being pressed
						//

						if ((wParam >= VK_NUMPAD0 && wParam <= VK_F24) ||
							(wParam >= VK_INSERT && wParam <= VK_DELETE) ||
							(wParam >= VK_PRIOR && wParam <= VK_DOWN) ||
							(wParam >= 'A' && wParam <= 'Z') ||
							(196 == wParam || 197 == wParam || 214 == wParam) ||
							(wParam == VK_BACK) ||
							(wParam == VK_RETURN) ||
							(wParam == VK_SCROLL)
						   )
						{
							//
							// Add shift to these keys if it is pressed
							//
							
							if (bShift)
								m_pShortCutStore->scpV1.m_bShift[0] = TRUE;

							m_pShortCutStore->scpV1.m_nKeyCodes[0] = wParam;
							m_pShortCutStore->scpV1.m_nKeyboardCodes[0] = lParam;
							if (wParam >= VK_PRIOR && wParam <= VK_DOWN && !(GetKeyState(VK_NUMLOCK) & 0x01))
								m_pShortCutStore->scpV1.m_nKeyboardCodes[0] |= 0x01000000;

							m_pShortCutStore->scpV1.m_nControlKeys = nControlState;
							if (m_pShortCutStore->GetShortCutDesc(szDesc))
								SendMessage(WM_SETTEXT, 0, (LPARAM)szDesc);
						}
						else if ((wParam >= ' ' && wParam <= '@') ||
							     (wParam >= '[' && wParam <= '`') || 
							     (wParam >= '{' && wParam <= '~')
								)
						{
							m_pShortCutStore->scpV1.m_nKeyCodes[0] = wParam;
							m_pShortCutStore->scpV1.m_nKeyboardCodes[0] = lParam;
							m_pShortCutStore->scpV1.m_nControlKeys = nControlState;
							if (m_pShortCutStore->GetShortCutDesc(szDesc))
								SendMessage(WM_SETTEXT, 0, (LPARAM)szDesc);
						}
					}
				}
				return 0;
			}
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
	return FWnd::WindowProc(nMsg, wParam, lParam);
}

//
// KeyBoard
//

KeyBoard::KeyBoard(CBar* pBar)
	: FDialog(IDD_KEYBOARD),
	  m_pBar(pBar),
	  m_pSelectedTool(NULL),
	  m_theCustomizeCommands(0),
	  m_theCustomizeTools(0),
	  m_theKeyEdit(pBar)
{
	m_dwCommandsEventCookie = 0;
	m_dwToolsEventCookie = 0;
}

//
// ~KeyBoard
//

KeyBoard::~KeyBoard()
{
	ConnectionUnadvise(reinterpret_cast<IUnknown*>(&m_theCustomizeCommands),
					   IID_IDispatch,
					   this,
					   FALSE,
					   m_dwCommandsEventCookie);
	ConnectionUnadvise(reinterpret_cast<IUnknown*>(&m_theCustomizeTools),
					   IID_IDispatch,
					   this,
					   FALSE,
					   m_dwToolsEventCookie);
}

//
// Customize
//

void KeyBoard::Customize(UINT nId, LocalizationTypes eType)
{
	TCHAR szBuffer[MAX_PATH];
	HWND hWnd = GetDlgItem(m_hWnd, nId);
	if (IsWindow(hWnd))
	{
		LPCTSTR szLocaleString = m_pBar->Localizer()->GetString(eType);
		if (szLocaleString)
		{
			lstrcpyn(szBuffer, szLocaleString, MAX_PATH);
			SetWindowText(hWnd, szBuffer);
		}
	}
}

//
// DialogProc
//

BOOL KeyBoard::DialogProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	try
	{
		switch(message)
		{
		case WM_INITDIALOG:
			{
				CenterDialog(GetParent(m_hWnd));
				LPCTSTR szLocaleString = m_pBar->Localizer()->GetString(ddLTKeyboardCaption);
				if (szLocaleString)
				{
					TCHAR szBuffer[MAX_PATH];
					lstrcpyn(szBuffer, szLocaleString, MAX_PATH);
					SetWindowText(m_hWnd, szBuffer);
				}

				CRect rcTools;
				HWND hWndCmdOld = GetDlgItem(m_hWnd, IDC_LT_COMMANDS);
				GetWindowRect(hWndCmdOld, &rcTools);
				ScreenToClient(rcTools);
				DestroyWindow(hWndCmdOld);		

				m_theCustomizeTools.Recreate(FALSE);
				m_theCustomizeTools.put_Type(ddCTTools);
				m_theCustomizeTools.InternalCreateWindow(m_hWnd, rcTools);
				m_theCustomizeTools.put_ActiveBar((IActiveBar2*)m_pBar);
				ConnectionAdvise(reinterpret_cast<IUnknown*>(&m_theCustomizeTools),
								 IID_IDispatch,
								 this,
								 FALSE,
								 &m_dwToolsEventCookie);

				RECT rcCommands;
				HWND hWndCatOld = GetDlgItem(m_hWnd, IDC_LT_CATEGORY);
				GetWindowRect(hWndCatOld, &rcCommands);
				ScreenToClient(rcCommands);
				DestroyWindow(hWndCatOld);
				m_theCustomizeCommands.Recreate(FALSE);
				m_theCustomizeCommands.put_Type(ddCTCategories);
				m_theCustomizeCommands.InternalCreateWindow(m_hWnd, rcCommands);
				m_theCustomizeCommands.put_ActiveBar((IActiveBar2*)m_pBar);
				ConnectionAdvise(reinterpret_cast<IUnknown*>(&m_theCustomizeCommands),
								 IID_IDispatch,
								 this,
								 FALSE,
								 &m_dwCommandsEventCookie);

				BSTR bstrCategory;
				m_theCustomizeCommands.get_Category(&bstrCategory);
				if (bstrCategory)
					m_theCustomizeTools.put_Category(bstrCategory);
				SysFreeString(bstrCategory);

				m_theCustomizeTools.put_ToolDragDrop(VARIANT_FALSE);

				m_theKeyEdit.SubClassAttach(GetDlgItem(m_hWnd, IDC_ED_SHORTCUT));

				EnableWindow(GetDlgItem(m_hWnd, IDC_ASSIGN), FALSE);

				Customize(IDC_STATIC_CATE, ddLTCategoriesLabel);
				Customize(IDC_STATIC_CMD, ddLTCommandLabel);
				Customize(IDC_STATIC_NEWSHORTCUT, ddLTPressNewShortcutLabel);

				Customize(IDC_STATIC_CURRENTKEY, ddLTCurrentKeysLabel);
				Customize(IDC_STATIC_DESC, ddLTDescription);
				Customize(IDC_CLOSE, ddLTCloseButton);

				Customize(IDC_ASSIGN, ddLTAssignButton);
				Customize(IDC_REMOVE, ddLTRemoveButton);
				Customize(IDC_RESETALL, ddLTResetAllButton);
			}
			break;

		case WM_HELP:
			{
				LPHELPINFO pHelpInfo = (LPHELPINFO)lParam;
				pHelpInfo->hItemHandle = WindowFromPoint(pHelpInfo->MousePos);
				pHelpInfo->iCtrlId = GetWindowLong((HWND)pHelpInfo->hItemHandle, GWL_ID);
				pHelpInfo->dwContextId = 0; 
				((CBar*)m_pBar)->FireCustomizeHelp(pHelpInfo->iCtrlId);
			}
			return TRUE;

		case WM_DESTROY:
			m_theKeyEdit.UnsubClass();
			break;

		case WM_COMMAND:
			{
				WORD nNotifyCode = HIWORD(wParam);
				switch(LOWORD(wParam))
				{
				case IDC_CLOSE:
					EndDialog(m_hWnd, IDC_CLOSE);
					return TRUE;

				case IDC_ASSIGN:
					Assign();
					((CBar*)m_pBar)->m_vbCustomizeModified = VARIANT_TRUE;
					return TRUE;

				case IDC_REMOVE:
					Remove();
					((CBar*)m_pBar)->m_vbCustomizeModified = VARIANT_TRUE;
					return TRUE;

				case IDC_RESETALL:
					ResetAll();
					((CBar*)m_pBar)->m_vbCustomizeModified = VARIANT_TRUE;
					return TRUE;

				case IDC_LB_KEYS:
					return TRUE;

				case IDC_ED_SHORTCUT:
					{
						if (EN_CHANGE == nNotifyCode)
						{
							if (m_pSelectedTool)
							{
								TCHAR szText[64];
								GetDlgItemText(m_hWnd, IDC_ED_SHORTCUT, szText, 64);
								BOOL bEnable = FALSE;
								if (lstrlen(szText) > 0)
									bEnable = TRUE;
								EnableWindow(GetDlgItem(m_hWnd, IDC_ASSIGN), bEnable);
							}
						}
					}
					return TRUE;
				}
				break;
			}
		}
		return FDialog::DialogProc(message,wParam,lParam);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
}

//
// Assign
//

void KeyBoard::Assign()
{
	try
	{
		//
		// Get the shortcut text
		//

		BOOL bFound = FALSE;
		TCHAR szText[64];
		lstrcpy(szText, m_theKeyEdit.GetShortCut());

		//
		// Make sure this shortcut is not already in the listbox
		//

		TCHAR szCompareText[64];
		HWND hWndListbox = GetDlgItem(m_hWnd, IDC_LB_KEYS);
		int nCount = SendMessage(hWndListbox, LB_GETCOUNT, 0, (long)szText);
		for (int nIndex = 0; nIndex < nCount; nIndex++)
		{
			SendMessage(hWndListbox, LB_GETTEXT, nIndex, (long)szCompareText);
			if (0 == lstrcmp(szCompareText, szText))
			{
				bFound = TRUE;
				break;
			}
		}

		//
		// If the short cut was not found add it to the listbox
		//

		if (!bFound)
		{
			SendMessage(hWndListbox, LB_ADDSTRING, 0, (long)szText);

			//
			// Clear the edit control
			//
			
			SetDlgItemText(m_hWnd, IDC_ED_SHORTCUT, _T(""));
			ShortCutStore* pShortCutStore;
			m_theKeyEdit.GetShortCut(pShortCutStore);
			if (pShortCutStore)
			{
				HRESULT hResult = m_pSelectedTool->m_scTool.Add(pShortCutStore);
				if (FAILED(hResult))
				{
					assert(FALSE);
					return;
				}
				m_pBar->SetToolShortCut(m_pSelectedTool->tpV1.m_nToolId, &m_pSelectedTool->m_scTool);
				m_pBar->SetModified();
			}
			if (nCount <= 0)
				EnableWindow(GetDlgItem(m_hWnd, IDC_REMOVE),TRUE);
		}
		m_theKeyEdit.Clear();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// Remove
//

void KeyBoard::Remove()
{
	try
	{
		int nItem = -1;
		HWND hWndListbox = GetDlgItem(m_hWnd, IDC_LB_KEYS);
		if (IsWindow(hWndListbox))
		{
			nItem = SendMessage(hWndListbox, LB_GETCURSEL, 0, 0);
			if (LB_ERR == nItem)
				return;

			int nResult = SendMessage(hWndListbox, LB_DELETESTRING, nItem, 0);

			nResult = SendMessage(hWndListbox, LB_GETCOUNT, 0, 0);
			if (0 == nResult)
				EnableWindow(GetDlgItem(m_hWnd, IDC_REMOVE), FALSE);
			else
				SendMessage(hWndListbox, LB_SETCURSEL, 0, 0);
		}

		//
		// Remove the short cut from the Tool
		//

		if (-1 != nItem)
		{
			ShortCutStore* pShortCutStore = m_pSelectedTool->m_scTool.Get(nItem);
			if (pShortCutStore)
			{
				pShortCutStore->AddRef();
				m_pBar->RemoveShortCutFromAllTools(m_pSelectedTool->tpV1.m_nToolId, pShortCutStore);
				pShortCutStore->Release();
			}
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// FireSelChangeCategory
//

void KeyBoard::FireSelChangeCategory(BSTR strCategory)
{
	m_theCustomizeTools.put_Category(strCategory);
}

//
// FireSelChangeTool
//

void KeyBoard::FireSelChangeTool(ITool* pTool)
{
	try
	{
		m_pSelectedTool = static_cast<CTool*>(pTool);

		//
		// Clear the shortcut listbox
		//

		SendMessage(GetDlgItem(m_hWnd, IDC_LB_KEYS), LB_RESETCONTENT, 0, 0);

		//
		// Set the Description
		//
		
		BSTR bstrDesc;
		HRESULT hResult = pTool->get_Description(&bstrDesc);
		if (SUCCEEDED(hResult))
		{
			MAKE_TCHARPTR_FROMWIDE(szDesc, bstrDesc);
			SetDlgItemText(m_hWnd, IDC_DESCRIPTION, szDesc);
			SysFreeString(bstrDesc);
		}
		else
			SetDlgItemText(m_hWnd, IDC_DESCRIPTION, _T(""));

		//
		// Enable the Assign button if the shortcut edit control has something in it.
		//

		TCHAR szText[64];
		GetDlgItemText(m_hWnd, IDC_ED_SHORTCUT, szText, 64);
		BOOL bEnable = FALSE;
		if (lstrlen(szText) > 0)
			bEnable = TRUE;
		EnableWindow(GetDlgItem(m_hWnd, IDC_ASSIGN), bEnable);

		//
		// Have to fill the short cut listbox
		//

		ShortCutStore* pShortCutStore;
		int nSize = m_pSelectedTool->m_scTool.GetCount();
		for (int nIndex = 0; nIndex < nSize; nIndex++)
		{
			pShortCutStore = m_pSelectedTool->m_scTool.Get(nIndex);
			if (pShortCutStore)
			{
				if (pShortCutStore->GetShortCutDesc(szText))
					SendMessage(GetDlgItem(m_hWnd, IDC_LB_KEYS), LB_ADDSTRING, 0, (LPARAM)szText);
			}
		}
		if (nSize > 0)
		{
			SendMessage(GetDlgItem(m_hWnd, IDC_LB_KEYS), LB_SETCURSEL, 0, 0);
			EnableWindow(GetDlgItem(m_hWnd, IDC_REMOVE), TRUE);
		}
		else
			EnableWindow(GetDlgItem(m_hWnd, IDC_REMOVE), FALSE);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// CToolBar
//

CToolBar::CToolBar(CBar* pBar)
	: CDDPropertyPage(IDD_TOOLBAR),
	  m_theCustomizeListbox(NULL),
	  m_bstrCurrentBandName(NULL),
	  m_pBar(pBar)
{
    m_pspInfo.dwFlags = PSP_USETITLE; 
	
	LPCTSTR szLocaleString = m_pBar->Localizer()->GetString(ddLTToolbarTab);
	if (szLocaleString)
		lstrcpyn(m_szTitle, szLocaleString, MAX_PATH);
    m_pspInfo.pszTitle = m_szTitle;
}

//
// CToolBar
//

CToolBar::~CToolBar()
{
	SysFreeString(m_bstrCurrentBandName);
	ConnectionUnadvise(reinterpret_cast<IUnknown*>(&m_theCustomizeListbox),
					   IID_IDispatch,
					   reinterpret_cast<IUnknown*>(this),
					   FALSE,
					   m_dwEventCookie);
}

//
// OnPageInit 
//

BOOL CToolBar::OnPageInit()
{
	try
	{
		HWND hWndParent = GetParent(m_hWnd);
		if (hWndParent)
		{
			CenterWindow(hWndParent);
			
			HWND hWndCancel = GetDlgItem(hWndParent, IDCANCEL);
			CRect rcCancel;
			GetWindowRect(hWndCancel, &rcCancel);
			DestroyWindow(hWndCancel);

			HWND hWndApply = GetDlgItem(hWndParent, 0x3021);
			CRect rcApply;
			GetWindowRect(hWndApply, &rcApply);
			DestroyWindow(hWndApply);

			HWND hWndOk = GetDlgItem(hWndParent, IDOK);
			DestroyWindow(hWndOk);
			
			HFONT hFont = (HFONT)SendMessage(hWndParent, WM_GETFONT, 0, 0);
			POINT pt;
			pt.x = rcCancel.left;
			pt.y = rcCancel.top;
			ScreenToClient(hWndParent, &pt);

			TCHAR szBuffer[MAX_PATH];
			LPCTSTR szLocaleString = m_pBar->Localizer()->GetString(ddLTKeyboardButton);
			if (szLocaleString)
				lstrcpyn(szBuffer, szLocaleString, MAX_PATH);
			HWND hWnd;

#ifndef _ABDLL			
			hWnd = CreateWindow(_T("BUTTON"), 
									 szBuffer, 
									 BS_PUSHBUTTON|WS_CHILD|WS_VISIBLE, 
									 pt.x,
									 pt.y,
									 rcCancel.Width(), 
									 rcCancel.Height(), 
									 hWndParent,
									 (HMENU)CCustomize::eKeyboard,
									 g_hInstance,
									 0);
			if (hFont)
				SendMessage(hWnd, WM_SETFONT, (WPARAM)hFont, MAKELPARAM(FALSE, 0));
#endif
			pt.x = rcApply.left;
			pt.y = rcApply.top;
			ScreenToClient(hWndParent, &pt);
			
			szLocaleString = m_pBar->Localizer()->GetString(ddLTCloseButton);
			if (szLocaleString)
				lstrcpyn(szBuffer, szLocaleString, MAX_PATH);

			hWnd = CreateWindow(_T("BUTTON"), 
								szBuffer, 
								BS_PUSHBUTTON|WS_CHILD|WS_VISIBLE, 
								pt.x, 
								pt.y, 
								rcApply.Width(), 
								rcApply.Height(), 
								hWndParent,
								(HMENU)CCustomize::eClose,
								g_hInstance,
								0);
			if (hFont)
				SendMessage(hWnd, WM_SETFONT, (WPARAM)hFont, MAKELPARAM(FALSE, 0));
			
			CRect rcToolbar;
			GetClientRect(m_hWnd, &rcToolbar);
			rcToolbar.Inflate(-10, -10);
			
			hWnd = GetDlgItem(m_hWnd, IDC_RESET);
			if (IsWindow(hWnd))
			{
				szLocaleString = m_pBar->Localizer()->GetString(ddLTResetButton);
				if (szLocaleString)
				{
					lstrcpyn(szBuffer, szLocaleString, MAX_PATH);
					SetWindowText(hWnd, szBuffer);
				}
			}
			
			hWnd = GetDlgItem(m_hWnd, IDC_RENAME);
			if (IsWindow(hWnd))
			{
				szLocaleString = m_pBar->Localizer()->GetString(ddLTRenameButton);
				if (szLocaleString)
				{
					lstrcpyn(szBuffer, szLocaleString, MAX_PATH);
					SetWindowText(hWnd, szBuffer);
				}
			}
			
			hWnd = GetDlgItem(m_hWnd, IDC_DELETE);
			if (IsWindow(hWnd))
			{
				szLocaleString = m_pBar->Localizer()->GetString(ddLTDeleteButton);
				if (szLocaleString)
				{
					lstrcpyn(szBuffer, szLocaleString, MAX_PATH);
					SetWindowText(hWnd, szBuffer);
				}
			}
			
			hWnd = GetDlgItem(m_hWnd, IDC_NEW);
			if (IsWindow(hWnd))
			{
				szLocaleString = m_pBar->Localizer()->GetString(ddLTNewButton);
				if (szLocaleString)
				{
					lstrcpyn(szBuffer, szLocaleString, MAX_PATH);
					SetWindowText(hWnd, szBuffer);
				}
			}

			EnableWindow(GetDlgItem(m_hWnd, IDC_RENAME), FALSE);
			EnableWindow(GetDlgItem(m_hWnd, IDC_DELETE), FALSE);
			EnableWindow(GetDlgItem(m_hWnd, IDC_RESET), TRUE);

			CRect rc;
			GetWindowRect(hWnd, &rc);
			pt.x = rc.left;
			pt.y = rc.top;
			ScreenToClient(m_hWnd, &pt);
			rcToolbar.right = pt.x - 10; 
			ConnectionAdvise(reinterpret_cast<IUnknown*>(&m_theCustomizeListbox),
							 IID_IDispatch,
							 reinterpret_cast<IUnknown*>(this),
							 FALSE,
							 &m_dwEventCookie);
			m_theCustomizeListbox.Recreate(FALSE);
			m_theCustomizeListbox.InternalCreateWindow(m_hWnd, rcToolbar);
			m_theCustomizeListbox.put_ActiveBar((IActiveBar2*)m_pBar);
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
	return TRUE;
}

void CToolBar::Clear()
{
	m_theCustomizeListbox.Clear();
}

void CToolBar::Reset()
{
	m_theCustomizeListbox.put_ActiveBar(m_pBar);
}

//
// OnCommand
//

BOOL CToolBar::OnCommand(const WORD& wNotifyCode, const WORD& wID, const HWND hWnd)
{
	try
	{
		switch (wNotifyCode)
		{
		case BN_CLICKED:
			switch (wID)
			{
			case CCustomize::eClose:
				DestroyWindow(GetParent(m_hWnd));
				m_pBar->DoCustomization(FALSE);
				break;

			case CCustomize::eKeyboard:
				{
					KeyBoard theKeyBoard(m_pBar);
					theKeyBoard.DoModal(m_hWnd);
				}
				break;

			case IDC_NEW:
				{
					CReturnString* pReturnString = CReturnString::CreateInstance(NULL);		
					if (pReturnString)
					{
						m_pBar->FireNewToolbar((ReturnString*)pReturnString);
						BSTR bstrBand = NULL;
						HRESULT hResult = pReturnString->get_Value(&bstrBand);
						pReturnString->Release();
						hResult = m_theCustomizeListbox.Add(bstrBand, ddBTNormal);
						if (SUCCEEDED(hResult))
							m_pBar->m_vbCustomizeModified = VARIANT_TRUE;
						SysFreeString(bstrBand);
					}
				}
				break;

			case IDC_RENAME:
				{
					ToolBarName dlgToolBarName(m_pBar, ToolBarName::eRename);
					MAKE_TCHARPTR_FROMWIDE(szBuffer, m_bstrCurrentBandName);
					lstrcpy(dlgToolBarName.Name(), szBuffer);
					if (IDOK == dlgToolBarName.DoModal())
					{
						MAKE_WIDEPTR_FROMTCHAR(wBuffer, dlgToolBarName.Name());
						HRESULT hResult = m_theCustomizeListbox.Rename(wBuffer);
						if (SUCCEEDED(hResult))
							m_pBar->m_vbCustomizeModified = VARIANT_TRUE;
					}
				}
				break;

			case IDC_DELETE:
				{
					HRESULT hResult = m_theCustomizeListbox.Delete();
					if (SUCCEEDED(hResult))
						m_pBar->m_vbCustomizeModified = VARIANT_TRUE;
				}
				break;

			case IDC_RESET:
				{
					HRESULT hResult = m_theCustomizeListbox.Reset();
					if (SUCCEEDED(hResult))
						m_pBar->m_vbCustomizeModified = VARIANT_TRUE;
				}
				break;
			}
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
	return TRUE;
}

//
// OnNotify 
//

BOOL CToolBar::OnNotify(const int&    nControlId, 
					    const LPNMHDR pNMHDR, 
						BOOL&         bResult) 
{ 
	switch (pNMHDR->code)
	{
	case PSN_SETACTIVE:
		SetFocus(m_theCustomizeListbox.m_hWnd);
		bResult = TRUE;
		break;
	}
	return FALSE; 
}

//
// FireSelChangeBand
//

void CToolBar::FireSelChangeBand(IBand* pBand)
{
	try
	{
		assert(pBand);
		if (NULL == pBand)
			return;

		SysFreeString(m_bstrCurrentBandName);

		HRESULT hResult = pBand->get_Name(&m_bstrCurrentBandName);

		CreatedByTypes nCreatedBy;
		hResult = pBand->get_CreatedBy(&nCreatedBy);
		if (SUCCEEDED(hResult))
		{
			BOOL bEndUserCustom = FALSE;
			if (ddCBApplication != nCreatedBy)
				bEndUserCustom = TRUE;

			EnableWindow(GetDlgItem(m_hWnd, IDC_RENAME), bEndUserCustom);
			EnableWindow(GetDlgItem(m_hWnd, IDC_DELETE), bEndUserCustom);
			EnableWindow(GetDlgItem(m_hWnd, IDC_RESET), !bEndUserCustom);
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// CCommands
//

CCommands::CCommands(CBar* pBar)
	: CDDPropertyPage(IDD_COMMANDS),
	  m_theCustomizeCommands(0),
	  m_theCustomizeTools(0),
	  m_pBar(pBar)
{
    m_pspInfo.dwFlags = PSP_USETITLE; 
	LPCTSTR szLocaleString = m_pBar->Localizer()->GetString(ddLTCommandTab);
	if (szLocaleString)
		lstrcpyn(m_szTitle, szLocaleString, MAX_PATH);
    m_pspInfo.pszTitle = m_szTitle;
	m_dwCommandsEventCookie = 0;
	m_dwToolsEventCookie = 0;
}

CCommands::~CCommands()
{
	ConnectionUnadvise(reinterpret_cast<IUnknown*>(&m_theCustomizeCommands),
					   IID_IDispatch,
					   reinterpret_cast<IUnknown*>(this),
					   FALSE,
					   m_dwCommandsEventCookie);

	ConnectionUnadvise(reinterpret_cast<IUnknown*>(&m_theCustomizeTools),
					   IID_IDispatch,
					   reinterpret_cast<IUnknown*>(this),
					   FALSE,
					   m_dwToolsEventCookie);
}

//
// OnPageInit
//

BOOL CCommands::OnPageInit()
{
	try
	{
		CRect rcTemp;
		GetClientRect(GetDlgItem(m_hWnd, IDC_ST_CATEGORIES), &rcTemp);

		TCHAR szBuffer[MAX_PATH];
		LPCTSTR szLocaleString;
		HWND hWnd = GetDlgItem(m_hWnd, IDC_ST_CATEGORIES);
		if (IsWindow(hWnd))
		{
			szLocaleString = m_pBar->Localizer()->GetString(ddLTCategoriesLabel);
			if (szLocaleString)
			{
				lstrcpyn(szBuffer, szLocaleString, MAX_PATH);
				SetWindowText(hWnd, szBuffer);
			}
		}

		hWnd = GetDlgItem(m_hWnd, IDC_ST_COMMANDS);
		if (IsWindow(hWnd))
		{
			szLocaleString = m_pBar->Localizer()->GetString(ddLTCommandLabel);
			if (szLocaleString)
			{
				lstrcpyn(szBuffer, szLocaleString, MAX_PATH);
				SetWindowText(hWnd, szBuffer);
			}
		}

		hWnd = GetDlgItem(m_hWnd, IDC_LABEL_DESC);
		if (IsWindow(hWnd))
		{
			szLocaleString = m_pBar->Localizer()->GetString(ddLTCommandDesc);
			if (szLocaleString)
			{
				lstrcpyn(szBuffer, szLocaleString, MAX_PATH);
				SetWindowText(hWnd, szBuffer);
			}
		}

		hWnd = GetDlgItem(m_hWnd, IDC_MODSELECT);
		if (IsWindow(hWnd))
		{
			szLocaleString = m_pBar->Localizer()->GetString(ddLTModifySelection);
			if (szLocaleString)
			{
				lstrcpyn(szBuffer, szLocaleString, MAX_PATH);
				SetWindowText(hWnd, szBuffer);
			}
		}

		CRect rcTools;
		GetClientRect(m_hWnd, &rcTools);
		rcTools.Inflate(-10, 0);
		rcTools.top += rcTemp.Height() + 4;

		GetClientRect(GetDlgItem(m_hWnd, IDC_LABEL_DESC), &rcTemp);
		rcTools.bottom -= rcTemp.Height() + 8;
		CRect rcCommands = rcTools;
		
		rcCommands.right = (rcCommands.Width() / 2) - 5;
		rcTools.left = rcCommands.right + 10; 
		
		m_theCustomizeCommands.Recreate(FALSE);
		m_theCustomizeCommands.put_Type(ddCTCategories);
		m_theCustomizeCommands.InternalCreateWindow(m_hWnd, rcCommands);
		m_theCustomizeCommands.put_ActiveBar((IActiveBar2*)m_pBar);
		ConnectionAdvise(reinterpret_cast<IUnknown*>(&m_theCustomizeCommands),
						 IID_IDispatch,
						 reinterpret_cast<IUnknown*>(this),
						 FALSE,
						 &m_dwCommandsEventCookie);

		BSTR bstrCategory;
		m_theCustomizeCommands.get_Category(&bstrCategory);
		m_theCustomizeTools.Recreate(FALSE);
		m_theCustomizeTools.put_Type(ddCTTools);
		m_theCustomizeTools.InternalCreateWindow(m_hWnd, rcTools);
		m_theCustomizeTools.put_ActiveBar((IActiveBar2*)m_pBar);
		ConnectionAdvise(reinterpret_cast<IUnknown*>(&m_theCustomizeTools),
						 IID_IDispatch,
						 reinterpret_cast<IUnknown*>(this),
						 FALSE,
						 &m_dwToolsEventCookie);

		m_theCustomizeTools.put_Category(bstrCategory);
		
		m_pBar->m_hWndModifySelection = GetDlgItem(m_hWnd, IDC_MODSELECT);
		if (m_pBar->m_diCustSelection.pTool)
			EnableWindow(m_pBar->m_hWndModifySelection, TRUE);

		SysFreeString(bstrCategory);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
	return TRUE;
}

void CCommands::Clear()
{
	m_theCustomizeCommands.Clear();
	m_theCustomizeTools.Clear();
}

void CCommands::Reset()
{
	m_theCustomizeCommands.put_ActiveBar(m_pBar);
	m_theCustomizeTools.put_ActiveBar(m_pBar);
}

//
// OnDestroy
//

BOOL CCommands::OnDestroy()
{
	m_pBar->m_hWndModifySelection = NULL;
	return TRUE;
}

//
// OnCommand
//

BOOL CCommands::OnCommand(const WORD& wNotifyCode, const WORD& wID, const HWND hWnd)
{
	try
	{
		switch (wNotifyCode)
		{
		case BN_CLICKED:
			switch (wID)
			{
			case CCustomize::eClose:
				DestroyWindow(GetParent(m_hWnd));
				m_pBar->DoCustomization(FALSE);
				break;

			case CCustomize::eKeyboard:
				{
					KeyBoard theKeyBoard(m_pBar);
					theKeyBoard.DoModal(m_hWnd);
				}
				break;

			case IDC_MODSELECT:
				{
					if (BST_UNCHECKED == ::SendMessage(GetDlgItem(m_hWnd, IDC_MODSELECT), BM_GETCHECK, 0, 0))
					{
						if (m_pBar->m_diCustSelection.pTool)
						{
							//
							// change the state of the button and bring up the popup
							//

							SendMessage(GetDlgItem(m_hWnd, IDC_MODSELECT), BM_SETCHECK, BST_CHECKED, 0);
							CRect rcButton;
							GetWindowRect(hWnd, &rcButton);
							SIZE size = {rcButton.left, rcButton.bottom}; 
							PixelToTwips(&size, &size);
							POINT ptScreen = {size.cx, size.cy};
							try
							{
								if (m_pBar->m_diCustSelection.pBand->SetupEditToolPopup(m_pBar->m_diCustSelection.pTool))
								{
									m_pBar->m_mcInfo.pBand = m_pBar->m_diCustSelection.pBand;
									m_pBar->m_mcInfo.pTool = m_pBar->m_diCustSelection.pTool;
									VARIANT x;
									VARIANT y;
									VARIANT vFlags;
									x.vt = VT_I2;
									y.vt = VT_I2;
									vFlags.vt = VT_I2;

									x.iVal = (short)ptScreen.x;
									y.iVal = (short)ptScreen.y;
									vFlags.iVal = (int)ddPopupMenuLeftAlign;
									
									m_pBar->m_pEditToolPopup->PopupMenu(vFlags, x, y);
									m_pBar->m_pEditToolPopup->Release();	
									m_pBar->m_pEditToolPopup = NULL;
									SendMessage(GetDlgItem(m_hWnd, IDC_MODSELECT), BM_SETCHECK, BST_UNCHECKED, 0);
								}
							}
							CATCH
							{
								assert(FALSE);
								REPORTEXCEPTION(__FILE__, __LINE__)
							}			
						}
					}
					else
					{
						::SendMessage(m_pBar->m_hWndModifySelection, BM_SETCHECK, BST_UNCHECKED, 0);
						if (m_pBar->m_pEditToolPopup && m_pBar->m_pEditToolPopup->m_pPopupWin)
							m_pBar->m_pEditToolPopup->m_pPopupWin->DestroyWindow();
					}
				}
				break;
			}
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
	return TRUE;
}

//
// OnNotify 
//

BOOL CCommands::OnNotify(const int&    nControlId, 
						 const LPNMHDR pNMHDR, 
						 BOOL&         bResult) 
{ 
	switch (pNMHDR->code)
	{
	case PSN_SETACTIVE:
		SetFocus(m_theCustomizeCommands.m_hWnd);
		bResult = TRUE;
		break;
	}
	return FALSE; 
}

//
// FireSelChangeCategory
//

void CCommands::FireSelChangeCategory(BSTR bstrCategory)
{
	m_theCustomizeTools.put_Category(bstrCategory);
}

//
// FireSelChangeTool
//

void CCommands::FireSelChangeTool(ITool* pTool)
{
	try
	{
		BSTR bstrDesc;
		HRESULT hResult = pTool->get_Description(&bstrDesc);
		if (SUCCEEDED(hResult))
		{
			MAKE_TCHARPTR_FROMWIDE(szDesc, bstrDesc);
			SetDlgItemText(m_hWnd, IDC_ST_DESC, szDesc);
			SysFreeString(bstrDesc);
		}
		else
			SetDlgItemText(m_hWnd, IDC_ST_DESC, _T(""));
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// COptions
//

COptions::COptions(CBar* pBar)
	: CDDPropertyPage(IDD_OPTIONS),
	  m_pBar(pBar)
{
    m_pspInfo.dwFlags = PSP_USETITLE; 
	LPCTSTR szLocaleString = m_pBar->Localizer()->GetString(ddLTOptionsTab);
	if (szLocaleString)
		lstrcpyn(m_szTitle, szLocaleString, MAX_PATH);
    m_pspInfo.pszTitle = m_szTitle;
}

//
// OnPageInit
//

BOOL COptions::OnPageInit()
{
	try
	{
		static int nMenu[4][2] = 
		{
			{ddMSAnimateUnfold, 2},
			{ddMSAnimateSlide, 3},
			{ddMSAnimateNone, 0},
			{ddMSAnimateRandom, 1}
		};

		TCHAR szBuffer[MAX_PATH];
		LPCTSTR szLocaleString;
		HWND hWnd = GetDlgItem(m_hWnd, IDC_CUSTOM2);
		if (IsWindow(hWnd))
		{
			szLocaleString = m_pBar->Localizer()->GetString(ddLTPersonalizedMenu);
			if (szLocaleString)
			{
				lstrcpyn(szBuffer, szLocaleString, MAX_PATH);
				SetWindowText(hWnd, szBuffer);
			}
		}

		hWnd = GetDlgItem(m_hWnd, IDC_CK_MRUFIRST);
		if (IsWindow(hWnd))
		{
			szLocaleString = m_pBar->Localizer()->GetString(ddLTMenuShowMRUFirst);
			if (szLocaleString)
			{
				lstrcpyn(szBuffer, szLocaleString, MAX_PATH);
				SetWindowText(hWnd, szBuffer);
			}
		}
		hWnd = GetDlgItem(m_hWnd, IDC_CK_FULLMENUAFTERDELAY);
		if (IsWindow(hWnd))
		{
			szLocaleString = m_pBar->Localizer()->GetString(ddLTShowFullMenuAfterDelay);
			if (szLocaleString)
			{
				lstrcpyn(szBuffer, szLocaleString, MAX_PATH);
				SetWindowText(hWnd, szBuffer);
			}
		}
		hWnd = GetDlgItem(m_hWnd, IDC_BN_MENUUSAGE);
		if (IsWindow(hWnd))
		{
			szLocaleString = m_pBar->Localizer()->GetString(ddLTResetUsageData);
			if (szLocaleString)
			{
				lstrcpyn(szBuffer, szLocaleString, MAX_PATH);
				SetWindowText(hWnd, szBuffer);
			}
		}
		hWnd = GetDlgItem(m_hWnd, IDC_CUSTOM1);
		if (IsWindow(hWnd))
		{
			szLocaleString = m_pBar->Localizer()->GetString(ddLTOther);
			if (szLocaleString)
			{
				lstrcpyn(szBuffer, szLocaleString, MAX_PATH);
				SetWindowText(hWnd, szBuffer);
			}
		}
		hWnd = GetDlgItem(m_hWnd, IDC_CK_LARGE);
		if (IsWindow(hWnd))
		{
			szLocaleString = m_pBar->Localizer()->GetString(ddLTLargeIcons);
			if (szLocaleString)
			{
				lstrcpyn(szBuffer, szLocaleString, MAX_PATH);
				SetWindowText(hWnd, szBuffer);
			}
		}
		hWnd = GetDlgItem(m_hWnd, IDC_CK_SCREENTIP);
		if (IsWindow(hWnd))
		{
			szLocaleString = m_pBar->Localizer()->GetString(ddLTScreenTips);
			if (szLocaleString)
			{
				lstrcpyn(szBuffer, szLocaleString, MAX_PATH);
				SetWindowText(hWnd, szBuffer);
			}
		}
		hWnd = GetDlgItem(m_hWnd, IDC_CK_SHORTCUT);
		if (IsWindow(hWnd))
		{
			szLocaleString = m_pBar->Localizer()->GetString(ddLTShortcutKeys);
			if (szLocaleString)
			{
				lstrcpyn(szBuffer, szLocaleString, MAX_PATH);
				SetWindowText(hWnd, szBuffer);
			}
		}
		hWnd = GetDlgItem(m_hWnd, IDC_STATIC_MENU);
		if (IsWindow(hWnd))
		{
			szLocaleString = m_pBar->Localizer()->GetString(ddLTMenuAnimationLabel);
			if (szLocaleString)
			{
				lstrcpyn(szBuffer, szLocaleString, MAX_PATH);
				SetWindowText(hWnd, szBuffer);
			}
		}
		hWnd = GetDlgItem(m_hWnd, IDC_CB_MENUANIMATION);
		if (IsWindow(hWnd))
		{
			int nString = 0;
			szLocaleString = m_pBar->Localizer()->GetString(ddLTMANone + nString);
			int nCount = SendMessage(hWnd, CB_GETCOUNT, 0, 0);
			if (nCount > 0 && szLocaleString)
			{
				SendMessage(hWnd, CB_RESETCONTENT, 0, 0);
				for (; nString < nCount; nString++)
				{
					szLocaleString = m_pBar->Localizer()->GetString(ddLTMANone + nString);
					if (szLocaleString)
					{
						lstrcpyn(szBuffer, szLocaleString, MAX_PATH);
						SendMessage(hWnd, CB_ADDSTRING, 0, (LPARAM)szBuffer);
					}
				}
			}
		}

		PersonalizedMenuTypes pmMenu;
		HRESULT hResult = m_pBar->get_PersonalizedMenus(&pmMenu);
		if (SUCCEEDED(hResult))
		{
			CheckDlgButton(m_hWnd, IDC_CK_MRUFIRST, ddPMDisabled != pmMenu ? BST_CHECKED : BST_UNCHECKED);
			if (ddPMDisabled != pmMenu)
				CheckDlgButton(m_hWnd, IDC_CK_FULLMENUAFTERDELAY, ddPMDisplayOnHover == pmMenu ? BST_CHECKED : BST_UNCHECKED);
			else
				EnableWindow(GetDlgItem(m_hWnd, IDC_CK_FULLMENUAFTERDELAY), FALSE);
		}

		MenuStyles msStyle;
		hResult = m_pBar->get_MenuAnimation(&msStyle);
		if (SUCCEEDED(hResult))
			SendMessage(GetDlgItem(m_hWnd, IDC_CB_MENUANIMATION), CB_SETCURSEL, nMenu[msStyle][1], 0);

		VARIANT_BOOL vbResult;
		hResult = m_pBar->get_LargeIcons(&vbResult);
		if (SUCCEEDED(hResult))
			CheckDlgButton(m_hWnd, IDC_CK_LARGE, vbResult == VARIANT_TRUE ? BST_CHECKED : BST_UNCHECKED);
		
		VARIANT_BOOL vbDisplay;
		hResult = m_pBar->get_DisplayToolTips(&vbDisplay);
		if (SUCCEEDED(hResult))
			CheckDlgButton(m_hWnd, IDC_CK_SCREENTIP, vbDisplay == VARIANT_TRUE ? BST_CHECKED : BST_UNCHECKED);
		
		hResult = m_pBar->get_DisplayKeysInToolTip(&vbResult);
		if (SUCCEEDED(hResult))
			CheckDlgButton(m_hWnd, IDC_CK_SHORTCUT, vbResult == VARIANT_TRUE ? BST_CHECKED : BST_UNCHECKED);

		if (vbDisplay == VARIANT_FALSE)
			EnableWindow(GetDlgItem(m_hWnd, IDC_CK_SHORTCUT), FALSE);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
	return TRUE;
}

void COptions::Clear()
{
}

void COptions::Reset()
{
}

//
// OnCommand
//

BOOL COptions::OnCommand(const WORD& wNotifyCode, const WORD& wID, const HWND hWnd)
{
	try
	{
		HRESULT hResult;
		switch (wNotifyCode)
		{
		case BN_CLICKED:
			switch (wID)
			{
			case CCustomize::eClose:
				DestroyWindow(GetParent(m_hWnd));
				m_pBar->DoCustomization(FALSE);
				break;

			case CCustomize::eKeyboard:
				{
					KeyBoard theKeyBoard(m_pBar);
					theKeyBoard.DoModal(m_hWnd);
				}
				break;

			case IDC_CK_LARGE:
				hResult = m_pBar->put_LargeIcons(IsDlgButtonChecked(m_hWnd, IDC_CK_LARGE) ? VARIANT_TRUE : VARIANT_FALSE);
				if (SUCCEEDED(hResult))
					m_pBar->m_vbCustomizeModified = VARIANT_TRUE;
				break;

			case IDC_CK_SCREENTIP:
				{
					BOOL bChecked = IsDlgButtonChecked(m_hWnd, IDC_CK_SCREENTIP);
					hResult = m_pBar->put_DisplayToolTips(bChecked ? VARIANT_TRUE : VARIANT_FALSE);
					if (SUCCEEDED(hResult))
						m_pBar->m_vbCustomizeModified = VARIANT_TRUE;
					EnableWindow(GetDlgItem(m_hWnd, IDC_CK_SHORTCUT), bChecked);
				}
				break;

			case IDC_CK_SHORTCUT:
				hResult = m_pBar->put_DisplayKeysInToolTip(IsDlgButtonChecked(m_hWnd, IDC_CK_SHORTCUT) ? VARIANT_TRUE : VARIANT_FALSE);
				if (SUCCEEDED(hResult))
					m_pBar->m_vbCustomizeModified = VARIANT_TRUE;
				break;

			case IDC_BN_MENUUSAGE:
				m_pBar->ClearMenuUsageData();
				break;

			case IDC_CK_MRUFIRST:
				{
					if (IsDlgButtonChecked(m_hWnd, IDC_CK_MRUFIRST))
					{
						if (IsDlgButtonChecked(m_hWnd, IDC_CK_FULLMENUAFTERDELAY))
							m_pBar->put_PersonalizedMenus(ddPMDisplayOnHover);
						else
							m_pBar->put_PersonalizedMenus(ddPMDisplayOnClick);
						EnableWindow(GetDlgItem(m_hWnd, IDC_CK_FULLMENUAFTERDELAY), TRUE);
					}
					else
					{
						m_pBar->put_PersonalizedMenus(ddPMDisabled);
						EnableWindow(GetDlgItem(m_hWnd, IDC_CK_FULLMENUAFTERDELAY), FALSE);
					}
					return NOERROR;
				}
				break;

			case IDC_CK_FULLMENUAFTERDELAY:
				if (IsDlgButtonChecked(m_hWnd, IDC_CK_FULLMENUAFTERDELAY))
					m_pBar->put_PersonalizedMenus(ddPMDisplayOnHover);
				else
					m_pBar->put_PersonalizedMenus(ddPMDisplayOnClick);
				break;
			}
			break;

			case CBN_SELCHANGE:
				switch (wID)
				{
				case IDC_CB_MENUANIMATION:
					{
						int nIndex = SendMessage(hWnd, CB_GETCURSEL, 0, 0);
						if (CB_ERR != nIndex)
						{
							static int nMenu[4][2] = 
							{
								{0, ddMSAnimateNone},
								{1, ddMSAnimateRandom},
								{2, ddMSAnimateUnfold},
								{3, ddMSAnimateSlide}
							};
							hResult = m_pBar->put_MenuAnimation((MenuStyles)nMenu[nIndex][1]);
							if (SUCCEEDED(hResult))
								m_pBar->m_vbCustomizeModified = VARIANT_TRUE;
						}
					}
				}
				break;

		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
	return TRUE;
}

//
// OnNotify 
//

BOOL COptions::OnNotify(const int&    nControlId, 
					    const LPNMHDR pNMHDR, 
					    BOOL&         bResult) 
{ 
	switch (pNMHDR->code)
	{
	case PSN_SETACTIVE:
		GetDlgItem(m_hWnd, IDC_CK_MRUFIRST);
		bResult = TRUE;
		break;
	}
	return FALSE; 
}

//
// CCustomize
//

CCustomize::CCustomize(HWND hWndParent, CBar* pBar)
	: CDDPropertySheet(hWndParent),
	  m_pBar(pBar),
	  m_pageToolBar(pBar),
	  m_pageCommands(pBar),
	  m_pageOptions(pBar)
{
	CLabel::RegisterClass();
	LPCTSTR szLocaleString = m_pBar->Localizer()->GetString(ddLTCustomCaption);
	if (szLocaleString)
		lstrcpyn(m_szTitle, szLocaleString, MAX_PATH);
    m_pshInfo.pszCaption = m_szTitle;

	AddPage(&m_pageToolBar);
	
	AddPage(&m_pageCommands);
	
	AddPage(&m_pageOptions);
}

void CCustomize::Clear()
{
	m_pageToolBar.Clear();
	m_pageCommands.Clear();
	m_pageOptions.Clear();
}

void CCustomize::Reset()
{
	m_pageToolBar.Reset();
	m_pageCommands.Reset();
	m_pageOptions.Reset();
}

void CCustomize::EndCustomize()
{
	if (m_pBar)
		m_pBar->DoCustomization(FALSE);
}

BOOL CCustomize::Help(LPHELPINFO pHelpInfo)
{
	if (m_pBar)
		m_pBar->FireCustomizeHelp(pHelpInfo->iCtrlId);
	return TRUE;
}

void CCustomize::OnMinimize(BOOL bActivate)
{
	if (bActivate)
		m_pBar->ShowPopups();
	else
		m_pBar->HidePopups();
}
