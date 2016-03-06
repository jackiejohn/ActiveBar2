//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include <crtdbg.h>
#include "IpServer.h"
#include "Resource.h"
#include "WindowProc.h"
#include "FDialog.h"
#include "Support.h"
#include "Utility.h"
#include "Localizer.h"
#include "BTabs.h"
#include "FWnd.h"
#include "EventLog.h"
#include "Designer\DragDropMgr.h"
#include "ShortCuts.h"
#include "Custom.h"
#include "Map.h"
#include "Bar.h"
#include "Globals.h"

extern HINSTANCE g_hInstance;

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

TCHAR Globals::m_szAlt[25];
TCHAR Globals::m_szCtrl[25];
TCHAR Globals::m_szShift[25];
BOOL g_bDemoDlgDisplayed;

Globals::Globals()
	: m_ilSizer(NULL),
	  m_ilSlidingScrollImages(NULL),
	  m_hBitmapExpandVert(NULL),
	  m_hBitmapExpandHorz(NULL),
	  m_hBitmapSubMenu(NULL),
	  m_hBitmapSMCombo(NULL),
	  m_hBitmapMenuScroll(NULL),
	  m_hBitmapEndMarker(NULL),
	  m_hBitmapMenuCheck(NULL),
	  m_hBitmapMenuExpand(NULL),
	  m_hBitmapMDIButtons(NULL),
	  m_hBrushDither(NULL),
	  m_hBrushPattern(NULL),
	  m_hSplitVCursor(NULL),
	  m_hSplitHCursor(NULL),
	  m_pWhatsThisHelpActiveBar(NULL),
	  m_pCustomizeActiveBar(NULL),
	  m_hSmallCaption(NULL),
	  m_pEventLog(NULL),
	  m_hMMLibrary(NULL),
	  m_hHookAccelator(NULL),
	  m_hHandCursor(NULL),
	  m_hUser32Library(NULL)
{
	try
	{
		m_nSmallCaptionHeight = 0;
		m_seFunction = _set_se_translator( trans_func );

		g_bDemoDlgDisplayed = FALSE;

		FWnd::Init();
		FDialog::Init();
		CDDPropertySheet::Init();
		
		for (int nIndex = 0; nIndex < GrabStyleCount; nIndex++)
			m_hBrushGrab[nIndex] = NULL;
		
		TCHAR* szTemp = LoadStringRes(IDS_ALT);
		if (szTemp)
			lstrcpy(m_szAlt, szTemp);
		
		szTemp = LoadStringRes(IDS_CONTROL);	
		if (szTemp)
			lstrcpy(m_szCtrl, szTemp);
		
		szTemp = LoadStringRes(IDS_SHIFT);
		if (szTemp)
			lstrcpy(m_szShift, szTemp);

		m_bFullDrag = FALSE;
		m_bUseDBCSUI = FALSE;
		m_hHookAccelator = SetWindowsHookEx(WH_KEYBOARD, (HOOKPROC)KeyBoardHook, NULL, GetCurrentThreadId());
		assert(m_hHookAccelator);

		m_pmapAccelator = new FMap;
		assert(m_pmapAccelator);
		
		m_pmapBar = new FMap;
		assert(m_pmapBar);
		
		m_pDragDropMgr = new CDragDropMgr;
		assert(m_pDragDropMgr);
		
		WM_SETTOOLFOCUS = RegisterWindowMessage(_T("DDSETTOOLFOCUS"));
		assert(WM_SETTOOLFOCUS);

		WM_REFRESHMENUBAND = RegisterWindowMessage(_T("DDREFRESHMENU"));
		assert(WM_REFRESHMENUBAND);
		
		WM_TOOLTIPSUPPORT  = RegisterWindowMessage(_T("DDPPTP"));
		assert(WM_TOOLTIPSUPPORT);
		
		WM_ACTIVEBARCLICK = RegisterWindowMessage(_T("DDACTIVEBARCLICK"));
		assert(WM_ACTIVEBARCLICK);
		
		WM_ACTIVEBARTEXTCHANGE = RegisterWindowMessage(_T("DDACTIVETEXTCHANGE"));
		assert(WM_ACTIVEBARTEXTCHANGE);

		WM_RECALCLAYOUT = RegisterWindowMessage(_T("DDPPRECALC"));
		assert(WM_RECALCLAYOUT);
		
		WM_POPUPKEYDOWN = RegisterWindowMessage(_T("DDPPMSGKEYDOWN"));
		assert(WM_POPUPKEYDOWN);
		
		WM_POPUPWINMSG = RegisterWindowMessage(_T("DDPPMSG"));
		assert(WM_POPUPWINMSG);
		
		WM_SIZEPARENT = RegisterWindowMessage(_T("DDWM_SIZEPARENT"));
		assert(WM_SIZEPARENT);
		
		WM_KILLWINDOW = RegisterWindowMessage(_T("DDWM_KILLPOPUP"));
		assert(WM_KILLWINDOW);

		WM_FLOATSTATUS = RegisterWindowMessage(_T("DDWM_FLOATSTATUS"));
		assert(WM_FLOATSTATUS);

		WM_POSTACTIVATE = RegisterWindowMessage(_T("DDWM_POSTACTIVATE"));
		assert(WM_POSTACTIVATE);

		WM_UPDATETABTOOL = RegisterWindowMessage(_T("DDWM_UPDATETABTOOL"));
		assert(WM_UPDATETABTOOL);

		WM_CACHEDOCKAREAS = RegisterWindowMessage(_T("DDWM_CACHEDOCKAREAS"));
		assert(WM_CACHEDOCKAREAS);

		WM_ACTIVEBARCOMBOSELCHANGE = RegisterWindowMessage(_T("DDWM_ACTIVEBARCOMBOSELCHANGE"));
		assert(WM_ACTIVEBARCOMBOSELCHANGE);

		m_nDragDelay = GetProfileInt(_T("Windows"), _T("DragDelay"), DD_DEFDRAGDELAY);
		assert(m_nDragDelay);

		m_nDragDist = GetProfileInt(_T("Windows"), _T("DragMinDist"), DD_DEFDRAGMINDIST);
		assert(m_nDragDist);

		m_nIDClipBandFormat = RegisterClipboardFormat(_T("ActiveBar Band"));
		assert(m_nIDClipBandFormat);
			
		m_nIDClipCategoryFormat = RegisterClipboardFormat(_T("ActiveBar Category"));
		assert(m_nIDClipCategoryFormat);

		m_nIDClipToolFormat = RegisterClipboardFormat(_T("ActiveBar Tool"));
		assert(m_nIDClipToolFormat);

		m_nIDClipToolIdFormat = RegisterClipboardFormat(_T("ActiveBar Tool Id"));
		assert(m_nIDClipToolIdFormat);

		m_nIDClipBandToolIdFormat = RegisterClipboardFormat(_T("ActiveBar Band Tool Id"));
		assert(m_nIDClipBandToolIdFormat);

		m_nIDClipBandChildBandToolIdFormat = RegisterClipboardFormat(_T("ActiveBar Band ChildBand Tool Id"));
		assert(m_nIDClipBandChildBandToolIdFormat);

		m_pControls = new FMap;
		assert(m_pControls);

		m_nBrushPatternWidth = m_nBrushPatternHeight = 0;

		m_bUseDBCSUI = FALSE;
		HFONT hFontSys = GetStockFont(SYSTEM_FONT);
		HDC hDCTest = GetDC(NULL);
		if (hDCTest)
		{
			m_nBitDepth = GetDeviceCaps(hDCTest, BITSPIXEL);

			HFONT hFontOld = SelectFont(hDCTest, hFontSys);
			TEXTMETRIC tm;
			GetTextMetrics(hDCTest, &tm);
			SelectFont(hDCTest, hFontOld);
			switch(tm.tmCharSet)
			{
			case SHIFTJIS_CHARSET:
			case HANGEUL_CHARSET:
			case CHINESEBIG5_CHARSET:
				m_bUseDBCSUI = TRUE;
				break;

			default:
				break;
			}
			ReleaseDC(NULL, hDCTest);
		}

#ifdef JAPBUILD
		_fdDefaultControl.cbSizeofstruct = sizeof(FONTDESC);
		_fdDefaultControl.cySize.Hi = 0;
		_fdDefaultControl.cySize.Lo = 90000;
		_fdDefaultControl.fItalic = FALSE;
		_fdDefaultControl.fStrikethrough = FALSE;
		_fdDefaultControl.fUnderline = FALSE;
		_fdDefaultControl.sCharset = DEFAULT_CHARSET;
		_fdDefaultControl.lpstrName = (LPOLESTR)"‚l‚r ‚oƒSƒVƒbƒN\0\0";
		_fdDefaultControl.sWeight = FW_NORMAL;

		_fdDefault.cbSizeofstruct = sizeof(FONTDESC);
		_fdDefault.cySize.Hi = 0;
		_fdDefault.cySize.Lo = 90000;
		_fdDefault.fItalic = FALSE;
		_fdDefault.fStrikethrough = FALSE;
		_fdDefault.fUnderline = FALSE;
		_fdDefault.sCharset = DEFAULT_CHARSET;
		_fdDefault.lpstrName = (LPOLESTR)"‚l‚r ‚oƒSƒVƒbƒN\0\0";
		_fdDefault.sWeight = FW_NORMAL;
#else
		_fdDefaultControl.cbSizeofstruct = sizeof(FONTDESC);
		_fdDefaultControl.cySize.Hi = 0;
		_fdDefaultControl.cySize.Lo = 80000;
		_fdDefaultControl.fItalic = FALSE;
		_fdDefaultControl.fStrikethrough = FALSE;
		_fdDefaultControl.fUnderline = FALSE;
		_fdDefaultControl.sCharset = DEFAULT_CHARSET;
		_fdDefaultControl.lpstrName = OLESTR("MS Sans Serif");
		_fdDefaultControl.sWeight = FW_NORMAL;

		_fdDefault.cbSizeofstruct = sizeof(FONTDESC);
		_fdDefault.cySize.Hi = 0;
		_fdDefault.cySize.Lo = 80000;
		_fdDefault.fItalic = FALSE;
		_fdDefault.fStrikethrough = FALSE;
		_fdDefault.fUnderline = FALSE;
		_fdDefault.sCharset = DEFAULT_CHARSET;
		_fdDefault.lpstrName = OLESTR("Arial");
		_fdDefault.sWeight = FW_NORMAL;
#endif
		InitializeCriticalSection(&m_csGlobals);
		InitializeCriticalSection(&m_csPaintIcon);
	}
	CATCH
	{
		assert(FALSE);
		m_pEventLog = NULL;
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

Globals::~Globals()
{
	try
	{
		BOOL bResult;
		if (m_hHookAccelator)
		{
			bResult = UnhookWindowsHookEx(m_hHookAccelator);
			assert(bResult);
		}

		if (m_hBitmapExpandVert)
		{
			bResult = DeleteBitmap(m_hBitmapExpandVert);
			assert(bResult);
		}
		
		if (m_hBitmapExpandHorz)
		{
			bResult = DeleteBitmap(m_hBitmapExpandHorz);
			assert(bResult);
		}

		if (m_hBitmapMenuScroll)
		{
			bResult = DeleteBitmap(m_hBitmapMenuScroll);
			assert(bResult);
		}

		if (m_hBitmapEndMarker)
		{
			bResult = DeleteBitmap(m_hBitmapEndMarker);
			assert(bResult);
		}
		
		if (m_hBitmapSubMenu)
		{
			bResult = DeleteBitmap(m_hBitmapSubMenu);
			assert(bResult);
		}

		if (m_hBitmapSMCombo)
		{
			bResult = DeleteBitmap(m_hBitmapSMCombo);
			assert(bResult);
		}

		if (m_hBitmapMenuCheck)
		{
			bResult = DeleteBitmap(m_hBitmapMenuCheck);
			assert(bResult);
		}

		if (m_hBitmapMenuExpand)
		{
			bResult = DeleteBitmap(m_hBitmapMenuExpand);
			assert(bResult);
		}

		if (m_hBitmapMDIButtons)
		{
			bResult = DeleteBitmap(m_hBitmapMDIButtons);
			assert(bResult);
		}

		if (m_hHandCursor)
		{
			bResult = DestroyCursor(m_hHandCursor);
			assert(bResult);
		}

		if (m_hSmallCaption)
		{
			bResult = DeleteFont(m_hSmallCaption);
			assert(bResult);
		}

		if (m_ilSizer)
		{
			bResult = ImageList_Destroy(m_ilSizer);
			assert(bResult);
		}
		
		if (m_ilSlidingScrollImages)
		{
			bResult = ImageList_Destroy(m_ilSlidingScrollImages);
			assert(bResult);
		}

		for (int nIndex = 0; nIndex < GrabStyleCount; nIndex++)
		{
			if (m_hBrushGrab[nIndex])
			{
				bResult = DeleteBrush(m_hBrushGrab[nIndex]);
				assert(bResult);
			}
		}

		if (m_hBrushDither)
		{
			bResult = DeleteBrush(m_hBrushDither);
			assert(bResult);
		}

		if (m_hBrushPattern)
		{
			bResult = DeleteBrush(m_hBrushPattern);
			assert(bResult);
		}

		if (m_pmapBar)
		{
			assert(0 == m_pmapBar->GetCount());
			delete m_pmapBar;
		}

		if (m_pmapAccelator)
		{
			assert(0 == m_pmapAccelator->GetCount());
			delete m_pmapAccelator;
		}

		if (m_hMMLibrary)
			FreeLibrary(m_hMMLibrary);

		if (m_hUser32Library)
			FreeLibrary(m_hUser32Library);

		delete m_pDragDropMgr;

		FDialog::CheckMap();
		FDialog::CleanUp();
		FWnd::CheckMap();
		FWnd::CleanUp();
		CDDPropertySheet::Cleanup();

		delete m_pControls;
		delete m_pEventLog; 

		_set_se_translator(m_seFunction);

		DeleteCriticalSection(&m_csGlobals);
		DeleteCriticalSection(&m_csPaintIcon);
	}
	CATCH
	{
		assert(FALSE);
		m_pEventLog = NULL;
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}
void Globals::Reset()
{
	ClearGrabhandle();
}

void Globals::ClearAccelators()
{
	if (NULL == m_pmapAccelator)
		return;

	CBar* pBar;
	HWND hWnd;
	FPOSITION posMap = m_pmapAccelator->GetStartPosition();
	while (posMap)
	{

		m_pmapAccelator->GetNextAssoc(posMap, (LPVOID&)hWnd, (LPVOID&)pBar);
		try
		{
			pBar->m_pShortCuts->GetShortCutStore()->Clear();
		}
		catch (...)
		{
			assert(FALSE);
		}
	}
}

HCURSOR Globals::GetHandCursor()
{
	TCHAR szWinDir[MAX_PATH];
	EnterCriticalSection(&m_csGlobals);
	if (NULL == m_hHandCursor)
	{
		if (GetWindowsDirectory(szWinDir, MAX_PATH) > 0)
		{
			lstrcat(szWinDir, _T("\\winhlp32.exe"));
			HMODULE hModule = LoadLibrary(szWinDir);
			if (hModule) 
			{
				m_hHandCursor =	CopyCursor(::LoadCursor(hModule, MAKEINTRESOURCE(106)));
				FreeLibrary(hModule);
			}
		}
		if (NULL == m_hHandCursor)
			m_hHandCursor = LoadCursor(g_hInstance, MAKEINTRESOURCE(IDC_POINT));
	}
	LeaveCriticalSection(&m_csGlobals);
	return m_hHandCursor;
}

void Globals::ClearGrabhandle()
{
	EnterCriticalSection(&m_csGlobals);
	BOOL bResult;
	for (int nIndex = 0; nIndex < GrabStyleCount; nIndex++)
	{
		if (m_hBrushGrab[nIndex])
		{
			bResult = DeleteBrush(m_hBrushGrab[nIndex]);
			assert(bResult);
			m_hBrushGrab[nIndex] = NULL;
		}
	}
	LeaveCriticalSection(&m_csGlobals);
}

HFONT Globals::GetSmallCaptionFont()
{
	EnterCriticalSection(&m_csGlobals);
	if (m_hSmallCaption)
	{
		LOGFONT lfSmallCaption;
		int nResult = GetObject(GetStockFont(SYSTEM_FONT), sizeof(LOGFONT), &lfSmallCaption);
		if (nResult > 0)
		{
			m_nSmallCaptionHeight = lfSmallCaption.lfHeight = GetSystemMetrics(SM_CYSMSIZE);
			m_hSmallCaption = CreateFontIndirect(&lfSmallCaption);
		}
	}
	LeaveCriticalSection(&m_csGlobals);
	return m_hSmallCaption;	
}

EventLog* Globals::GetEventLog()
{
	EnterCriticalSection(&m_csGlobals);
	if (NULL == m_pEventLog)
	{
		try 
		{
			m_pEventLog = new EventLog(_T("DataDynamics ActiveBar 2.0"), NULL, g_fSysWinNT, FALSE);
			assert(m_pEventLog);
		}
		CATCH
		{
			assert(FALSE);
			m_pEventLog = NULL;
		}
	}
	LeaveCriticalSection(&m_csGlobals);
	return m_pEventLog;
}

HBRUSH Globals::GetGrabHandleBrush(int nIndex)
{
	EnterCriticalSection(&m_csGlobals);
	if (NULL == m_hBrushGrab[nIndex])
	{
		HBITMAP hBmp = (HBITMAP)LoadImage(g_hInstance, MAKEINTRESOURCE(IDB_GRAB1+nIndex), IMAGE_BITMAP, 0, 0, LR_LOADMAP3DCOLORS);
		if (g_fSysWinNT)
			m_hBrushGrab[nIndex] = CreatePatternBrush(hBmp);
		else
			m_hBrushGrab[nIndex] = (HBRUSH)hBmp;
	}
	LeaveCriticalSection(&m_csGlobals);
	return m_hBrushGrab[nIndex];
}

PLAYSOUNDFUNTION Globals::GetPlaySound()
{
	EnterCriticalSection(&m_csGlobals);
	if (NULL == m_hMMLibrary)
		m_hMMLibrary = LoadLibrary(_T("winmm.dll")); 
	if (m_hMMLibrary)
	{
		PLAYSOUNDFUNTION pFunction;
#ifdef UNICODE
		pFunction = (PLAYSOUNDFUNTION)GetProcAddress(m_hMMLibrary, _T("PlaySoundW"));
#else
		pFunction = (PLAYSOUNDFUNTION)GetProcAddress(m_hMMLibrary, _T("PlaySoundA"));
#endif 
		LeaveCriticalSection(&m_csGlobals);
		return pFunction;
	}
	LeaveCriticalSection(&m_csGlobals);
	return NULL;
}

//User32.lib
//SetLayeredWindowAttributes

SETLAYEREDWINDOWATTRIBUTES Globals::GetSetLayeredWindowAttributes()
{
	EnterCriticalSection(&m_csGlobals);
	if (NULL == m_hUser32Library)
		m_hUser32Library = LoadLibrary(_T("User32.dll")); 
	if (m_hUser32Library)
	{
		SETLAYEREDWINDOWATTRIBUTES pFunction;
		pFunction = (SETLAYEREDWINDOWATTRIBUTES)GetProcAddress(m_hUser32Library, _T("SetLayeredWindowAttributes"));
		LeaveCriticalSection(&m_csGlobals);
		return pFunction;
	}
	LeaveCriticalSection(&m_csGlobals);
	return NULL;
}

HIMAGELIST Globals::GetSizer()
{
	EnterCriticalSection(&m_csGlobals);
	if (NULL == m_ilSizer)
	{
		#define OBM_BTSIZE 32761
		m_ilSizer = ImageList_LoadImage(NULL, 
										MAKEINTRESOURCE(OBM_BTSIZE), 
										16, 
										0, 
										RGB(192, 192, 192), 
										IMAGE_BITMAP, 
										LR_DEFAULTCOLOR);
	}
	LeaveCriticalSection(&m_csGlobals);
	return m_ilSizer;
}
	
HIMAGELIST Globals::GetSlidingScrollImages()
{
	EnterCriticalSection(&m_csGlobals);
	if (NULL == m_ilSlidingScrollImages)
	{
		m_ilSlidingScrollImages = ImageList_LoadBitmap(g_hInstance,
													   MAKEINTRESOURCE(IDB_SLIBINGSCROLL),
													   BTabs::eImageSize,
													   0,
													   RGB(0, 128, 128));
	}
	LeaveCriticalSection(&m_csGlobals);
	return m_ilSlidingScrollImages;
}

HBITMAP Globals::GetBitmapMenuCheck()
{
	EnterCriticalSection(&m_csGlobals);
	if (NULL == m_hBitmapMenuCheck)
		m_hBitmapMenuCheck = LoadBitmap(g_hInstance,MAKEINTRESOURCE(IDB_MENUCHECK));
	LeaveCriticalSection(&m_csGlobals);
	return m_hBitmapMenuCheck;
}

HBITMAP	Globals::GetBitmapSubMenu()
{
	EnterCriticalSection(&m_csGlobals);
	if (NULL == m_hBitmapSubMenu)
		m_hBitmapSubMenu = LoadBitmap(g_hInstance,MAKEINTRESOURCE(IDB_SUBMENUBMP));
	LeaveCriticalSection(&m_csGlobals);
	return m_hBitmapSubMenu;
}

HBITMAP Globals::GetBitmapSMCombo()
{
	EnterCriticalSection(&m_csGlobals);
	if (NULL == m_hBitmapSMCombo)
		m_hBitmapSMCombo = LoadBitmap(g_hInstance,MAKEINTRESOURCE(IDB_SMCOMBO));
	LeaveCriticalSection(&m_csGlobals);
	return m_hBitmapSMCombo;
}

HBITMAP Globals::GetBitmapMenuScroll()
{
	EnterCriticalSection(&m_csGlobals);
	if (NULL == m_hBitmapMenuScroll)
		m_hBitmapMenuScroll = LoadBitmap(g_hInstance,MAKEINTRESOURCE(IDB_MENUSCROLL));
	LeaveCriticalSection(&m_csGlobals);
	return m_hBitmapMenuScroll;
}
HBITMAP Globals::GetBitmapExpandVert()
{
	EnterCriticalSection(&m_csGlobals);
	if (NULL == m_hBitmapExpandVert)
		m_hBitmapExpandVert = LoadBitmap(g_hInstance,MAKEINTRESOURCE(IDB_EXPANDVERT));
	LeaveCriticalSection(&m_csGlobals);
	return m_hBitmapExpandVert;
}
HBITMAP Globals::GetBitmapExpandHorz()
{
	EnterCriticalSection(&m_csGlobals);
	if (NULL == m_hBitmapExpandHorz)
		m_hBitmapExpandHorz = LoadBitmap(g_hInstance,MAKEINTRESOURCE(IDB_EXPANDHORZ));
	LeaveCriticalSection(&m_csGlobals);
	return m_hBitmapExpandHorz;
}
HBITMAP	Globals::GetBitmapEndMarker()
{
	EnterCriticalSection(&m_csGlobals);
	if (NULL == m_hBitmapEndMarker)
		m_hBitmapEndMarker  = LoadBitmap(g_hInstance,MAKEINTRESOURCE(IDB_OVERFLOW));
	LeaveCriticalSection(&m_csGlobals);
	return m_hBitmapEndMarker;
}

HBITMAP Globals::GetBitmapMenuExpand()
{
	EnterCriticalSection(&m_csGlobals);
	if (NULL == m_hBitmapMenuExpand)
		m_hBitmapMenuExpand = LoadBitmap(g_hInstance,MAKEINTRESOURCE(IDB_MENUEXPAND));
	LeaveCriticalSection(&m_csGlobals);
	return m_hBitmapMenuExpand;
}

HBITMAP Globals::GetBitmapMDIButtons()
{
	EnterCriticalSection(&m_csGlobals);
	if (NULL == m_hBitmapMDIButtons)
		m_hBitmapMDIButtons = LoadBitmap(g_hInstance,MAKEINTRESOURCE(IDB_XPMDIBUTTONS));
	LeaveCriticalSection(&m_csGlobals);
	return m_hBitmapMDIButtons;
}

HBRUSH Globals::GetBrushDither()
{
	EnterCriticalSection(&m_csGlobals);
	if (NULL == m_hBrushDither)
		m_hBrushDither = CreateHalftoneBrush(TRUE);
	LeaveCriticalSection(&m_csGlobals);
	return m_hBrushDither;
}

HBITMAP Globals::GetBrushPattern()
{
	EnterCriticalSection(&m_csGlobals);
	HKEY  hKey = NULL;
	DWORD dwType = REG_SZ;
	DWORD dwSize = 120;
	TCHAR* szData[120];
	SIZE size = {8, 8};
	long lResult = RegOpenKeyEx(HKEY_CURRENT_USER, 
								_T("Software\\Microsoft\\Visual Basic\\6.0"),
								0,
								KEY_ALL_ACCESS, 
								&hKey);
	if (ERROR_SUCCESS != lResult) 
	{
		lResult = RegOpenKeyEx(HKEY_CURRENT_USER, 
									_T("Software\\Microsoft\\Visual Basic\\5.0"),
									0,
									KEY_ALL_ACCESS, 
									&hKey);
	}
	if (ERROR_SUCCESS == lResult) 
	{
		lResult = RegQueryValueEx(hKey, 
								_T("GridHeight"), 
								0,
								&dwType,
								(LPBYTE)&szData, 
								&dwSize);

		if (ERROR_SUCCESS == lResult)
		{
			size.cx = atoi((TCHAR*)szData);
			lResult = RegQueryValueEx(hKey, 
									_T("GridWidth"), 
									0,
									&dwType,
									(LPBYTE)&szData, 
									&dwSize);
			if (ERROR_SUCCESS == lResult)
			{
				size.cy = atoi((TCHAR*)szData);
				TwipsToPixel(&size, &size);
			}
			else
				size.cx = 8;
		}
		RegCloseKey(hKey);
	}
	if (NULL == m_hBrushPattern || m_nBrushPatternWidth != (short)size.cx || m_nBrushPatternHeight != (short)size.cy)
	{
		if (m_hBrushPattern)
		{
			BOOL bResult = DeleteBitmap(m_hBrushPattern);
			assert(bResult);
		}
		m_nBrushPatternWidth = (short)size.cx;
		m_nBrushPatternHeight = (short)size.cy;

		int nSize = size.cx * size.cy;
		WORD* pwBits = new WORD[nSize];
		if (pwBits)
		{
			memset(pwBits, 0, nSize);
			pwBits[0] = 0x80;
			m_hBrushPattern = CreateBitmap(size.cx, size.cy, 1, 1, (void*)pwBits);
			delete [] pwBits;
		}
	}
	LeaveCriticalSection(&m_csGlobals);
	return m_hBrushPattern;
}

HCURSOR Globals::GetSplitVCursor()
{
	EnterCriticalSection(&m_csGlobals);
	if (NULL == m_hSplitVCursor)
		m_hSplitVCursor = LoadCursor(g_hInstance, MAKEINTRESOURCE(IDC_HSPLITBAR));
	LeaveCriticalSection(&m_csGlobals);
	return m_hSplitVCursor;
}

HCURSOR Globals::GetSplitHCursor()
{
	EnterCriticalSection(&m_csGlobals);
	if (NULL == m_hSplitHCursor)
		m_hSplitHCursor = LoadCursor(g_hInstance, MAKEINTRESOURCE(IDC_VSPLITBAR));
	LeaveCriticalSection(&m_csGlobals);
	return m_hSplitHCursor;
}

#ifdef _DEBUG
DumpContext::DumpContext()
{
	m_hDebug = CreateFile(_T("Dump.rtf"), 
						  GENERIC_WRITE, 
						  FILE_SHARE_READ, 
						  NULL, 
						  CREATE_ALWAYS, 
						  FILE_ATTRIBUTE_NORMAL, 
						  NULL);

}

DumpContext::~DumpContext()
{
	CloseHandle(m_hDebug);
}

DumpContext& GetDumpContext()
{
	static DumpContext dc;
	return dc;
}
#endif

ITypeInfo *g_bandsTypeLib=NULL;
ITypeInfo *g_childbandsTypeLib=NULL;

Globals* g_pGlobals = NULL;

Globals& GetGlobals()
{
	return *g_pGlobals;
}

#ifdef _DEBUG
int AllocHook(int nAllocType, void* pUserData, size_t size, int nBlockType, long nRequestNumber, const unsigned char* szFilename, int nLineNumber)
{
	switch (nAllocType)
	{
	case _HOOK_ALLOC:
		switch (nRequestNumber)
		{
		case 1609:
			if (szFilename)
			{
				TRACE2(5, "_HOOK_ALLOC Size: %i Request Number: %i\n", size, nRequestNumber);
				TRACE2(5, "FileName: %s LineNumber: %i\n", szFilename, nLineNumber);
			}
			else
			{
				TRACE2(5, "_HOOK_ALLOC Size: %i Request Number: %i\n", size, nRequestNumber);
			}
			break;
		}
		break;
	}
	return TRUE;
}
#endif

void InitializeLibrary()
{
	try
	{
//		_CrtSetReportMode( _CRT_WARN, _CRTDBG_MODE_DEBUG);
//		_CrtSetReportMode( _CRT_ERROR, _CRTDBG_MODE_DEBUG);

//		int iFlags = _CrtSetDbgFlag(_CRTDBG_REPORT_FLAG);
//		iFlags |= _CRTDBG_CHECK_ALWAYS_DF | _CRTDBG_DELAY_FREE_MEM_DF | _CRTDBG_LEAK_CHECK_DF;

//		_CrtSetDbgFlag(iFlags);
//		_CrtSetAllocHook(AllocHook);

		// for the tab control in customization dialog
		InitCommonControls(); 
		g_pGlobals = new Globals;
		assert(g_pGlobals);
		InitHeap();
	}
	catch (...)
	{
	}
}

void UninitializeLibrary()
{
	try
	{
		delete g_pGlobals;

		if (g_bandsTypeLib)
			g_bandsTypeLib->Release();
		
		if (g_childbandsTypeLib)
			g_childbandsTypeLib->Release();

		UnLoadTypeInfo();

		HDC hDCMem = GetMemDC();
		if (hDCMem)
			DeleteDC(hDCMem);

		CleanupHeap();

#ifdef _DEBUG
//		TRACE(5, "---------------- UNINITIALIZE LIBRARY MEMORY CHECK -----------\n");
//		_CrtCheckMemory();
//		_CrtDumpMemoryLeaks();
//		TRACE(5, "---------------- UNINITIALIZE LIBRARY WINDOW MAP CHECK -----------\n");
#endif

	}
	catch (...)
	{
		assert(FALSE);
	}
}

ITypeInfo *GetObjectTypeInfoEx(LCID lcid,REFIID iid)
{
	HRESULT hr;
	ITypeInfo **ppTypeInfo;
	if (iid==IID_IBands)
		ppTypeInfo=&g_bandsTypeLib;
	else if (iid==IID_IChildBands)
		ppTypeInfo=&g_childbandsTypeLib;
	else
		return NULL;

	if ((*ppTypeInfo)==NULL)
	{ //Try to load typelibrary
		ITypeLib *pTypeLib;
        ITypeInfo *pTypeInfoTmp;
        hr = LoadRegTypeLib(LIBID_PROJECT, 2, 0,lcid, &pTypeLib);
		if (FAILED(hr))
		{
			return NULL;
		}
		hr = pTypeLib->GetTypeInfoOfGuid(iid, &pTypeInfoTmp);
        *ppTypeInfo=pTypeInfoTmp;
		pTypeLib->Release();
	}

	if ((*ppTypeInfo)==NULL)
		return NULL;
	(*ppTypeInfo)->AddRef();
	return *ppTypeInfo;
}

