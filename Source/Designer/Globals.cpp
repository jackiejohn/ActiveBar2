//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "Globals.h"
#include "FRegKey.h"
#include "debug.h"
#include "..\EventLog.h"
#include "Support.h"
#include "DragDropMgr.h"
#include "DragDrop.h"
#include "Dialogs.h"
#include "IconEdit.h"
#include "Resource.h"

extern HINSTANCE g_hInstance;

//
// Globals
//

Globals::Globals()
	: m_thePreferences(_T("Data Dynamics"), _T("ActiveBar\\2.0\\Designer")),
	  m_pEventLog(NULL),
	  m_pDragDropMgr(NULL),
	  m_hIconCancel(NULL),
	  m_pFF(NULL),
	  m_hBrushScreen(NULL),
	  m_hBitmapCycle(NULL)
{
	InitializeCriticalSection(&m_csCriticalSection);
	m_bSysWinNT = FALSE;
	m_bSysWin95 = FALSE;
	m_bUseDBCSUI = FALSE;
	m_bSysWin95Shell = FALSE;
	
    DWORD dwVer = GetVersion();
    
	//
	//  swap the two lowest bytes of dwVer so that the major and minor version
    //  numbers are in a usable order.
    //  for dwWinVer: high byte = major version, low byte = minor version
    //     OS               Sys_WinVersion  (as of 5/2/95)
    //     =-------------=  =-------------=
    //     Win95            0x035F   (3.95)
    //     WinNT ProgMan    0x0333   (3.51)
    //     WinNT Win95 UI   0x0400   (4.00)
    //
    //

    DWORD dwWinVer = (UINT)(((dwVer & 0xFF) << 8) | ((dwVer >> 8) & 0xFF));
    if (dwVer < 0x80000000) 
	{
        m_bSysWinNT = TRUE;
        m_bSysWin95Shell = (dwWinVer >= 0x0334);
    } 
	else  
	{
        m_bSysWin95 = TRUE;
        m_bSysWin95Shell = TRUE;
	}

	HFONT hFontSystem = GetStockFont(SYSTEM_FONT);
	HDC hDCTest = GetDC(NULL);
	if (hDCTest)
	{
		m_nBitDepth = GetDeviceCaps(hDCTest, BITSPIXEL);

		HFONT hFontOld = SelectFont(hDCTest, hFontSystem);
		TEXTMETRIC tm;
		GetTextMetrics(hDCTest,&tm);
		SelectFont(hDCTest, hFontOld);
		ReleaseDC(NULL, hDCTest);

		switch(tm.tmCharSet)
		{
			case SHIFTJIS_CHARSET:
//			case HANGEUL_CHARSET:
			case CHINESEBIG5_CHARSET:
				m_bUseDBCSUI = TRUE;
				break;
		}
	}

	WM_BUTTONINACTIVE = RegisterWindowMessage(_T("WM_BUTTONINACTIVE"));
	WM_SETEDITFOCUS = RegisterWindowMessage(_T("WM_SETEDITFOCUS"));
	WM_SIZEPARENT = RegisterWindowMessage(_T("DDWM_SIZEPARENT"));
	WM_BROADCAST = RegisterWindowMessage(_T("DDWM_BROADCAST"));
	WM_ICONEDIT = RegisterWindowMessage(_T("DDWM_ICONEDIT"));
	WM_HIDE = RegisterWindowMessage(_T("WM_HIDE"));

	m_nIDClipBandFormat = RegisterClipboardFormat(_T("ActiveBar Band"));
	m_nIDClipCategoryFormat = RegisterClipboardFormat(_T("ActiveBar Category"));
	m_nIDClipToolFormat = RegisterClipboardFormat(_T("ActiveBar Tool"));
	m_nIDClipToolIdFormat = RegisterClipboardFormat(_T("ActiveBar Tool Id"));
	m_nIDClipBandToolIdFormat = RegisterClipboardFormat(_T("ActiveBar Band Tool Id"));
	m_nIDClipBandChildBandToolIdFormat = RegisterClipboardFormat(_T("ActiveBar Band ChildBand Tool Id"));

	m_nScrollInset = GetProfileInt(_T("windows"), _T("DragScrollInset"), DD_DEFSCROLLINSET);
	m_nScrollDelay = GetProfileInt(_T("windows"), _T("DragScrollDelay"), DD_DEFSCROLLDELAY);
	m_nScrollInterval = GetProfileInt(_T("windows"), _T("DragScrollInterval"), DD_DEFSCROLLINTERVAL);
	
	m_pDefineColor = new CDefineColor();
	assert(m_pDefineColor);
}

static HDC s_hMemDC = NULL;

Globals::~Globals()
{
	DeleteCriticalSection(&m_csCriticalSection);
	delete m_pDragDropMgr;
	delete m_pEventLog;
	delete m_pFF;
	delete m_pDefineColor;

	if (m_hIconCancel)
		DestroyIcon(m_hIconCancel);
	if (m_hBrushScreen)
		DeleteBrush(m_hBrushScreen);
	if (m_hBitmapCycle)
		DeleteBrush(m_hBitmapCycle);
	if (s_hMemDC)
		DeleteDC(s_hMemDC);
}

HDC Globals::GetMemDC()
{
	if (NULL == s_hMemDC)
	{
		HDC hdc = GetDC(GetDesktopWindow());
		assert(hdc);
		if (hdc)
			s_hMemDC = CreateCompatibleDC(hdc);
		ReleaseDC(GetDesktopWindow(),hdc);
	}
	return s_hMemDC;
}

EventLog* Globals::GetEventLog()
{
	if (NULL == m_pEventLog)
	{
		try 
		{
			m_pEventLog = new EventLog(_T("DataDynamics ActiveBar 2.0 Designer"), NULL, m_bSysWinNT, FALSE);
			assert(m_pEventLog);
		}
		catch (SEException& e)
		{
			assert(FALSE);
			e.ReportException(__FILE__, __LINE__);
			m_pEventLog = NULL;
		}
	}
	return m_pEventLog;
}

//
// GetDragDropMgr
//

CDragDropMgr* Globals::GetDragDropMgr()
{
	if (NULL == m_pDragDropMgr)
		m_pDragDropMgr = new CDragDropMgr;
	return m_pDragDropMgr;
}

//
// GetCancelIcon
//

HICON Globals::GetCancelIcon()
{
	if (NULL == m_hIconCancel)
	{
		m_hIconCancel = (HICON)LoadImage(g_hInstance, 
										 MAKEINTRESOURCE(IDI_ICON_CANCEL), 
										 IMAGE_ICON, 
										 15, 
										 15, 
										 LR_DEFAULTCOLOR);
	}
	return m_hIconCancel;
}

CFlickerFree* Globals::FlickerFree()
{
	if (NULL == m_pFF)
		m_pFF = new CFlickerFree;
	assert(m_pFF);
	return m_pFF;
}

HBITMAP Globals::CycleBitmap()
{
	if (NULL == m_hBitmapCycle)
		m_hBitmapCycle = LoadBitmap(g_hInstance, MAKEINTRESOURCE(IDB_CYCLE));
	assert(m_hBitmapCycle);
	return m_hBitmapCycle;
}

HBRUSH Globals::ScreenBrush()
{
	if (NULL == m_hBrushScreen)
	{
		HBITMAP hBitmapHatch = LoadBitmap(g_hInstance, MAKEINTRESOURCE(IDB_HATCH));	
		assert(hBitmapHatch);
		if (hBitmapHatch)
		{
			m_hBrushScreen = CreatePatternBrush(hBitmapHatch);
			assert(m_hBrushScreen);
			DeleteBitmap(hBitmapHatch);
		}
	}
	assert(m_hBrushScreen);
	return m_hBrushScreen;
}
