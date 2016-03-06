#ifndef APPGLOBALS_INCLUDED
#define APPGLOBALS_INCLUDED

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "AppPref.h"
class CDefineColor;
class CFlickerFree;
class CDragDropMgr;
class EventLog;

enum Preferences
{
	eMainWindow = 10,
	eTreeWindow = 20,
	eBrowserWindow = 30
};

struct Globals
{
	Globals();
	~Globals();

	BOOL MachineHasLicense();

	BOOL GetLicenseKey(DWORD dwReserved, BSTR* bstr);
	
	HDC GetMemDC();

	EventLog* GetEventLog();

	CFlickerFree* FlickerFree();
	HBITMAP CycleBitmap();
	HBRUSH ScreenBrush();
	HICON GetCancelIcon();

	CDragDropMgr* GetDragDropMgr();
#ifdef _DEBUG
	void ReportMap();
#endif

    CRITICAL_SECTION m_csCriticalSection;
	CAppPref		 m_thePreferences;

	CDefineColor*    m_pDefineColor;
	CFlickerFree*    m_pFF;
	HBRUSH           m_hBrushScreen;
	HBITMAP          m_hBitmapCycle;
    BOOL			 m_bSysWin95Shell;
	BOOL			 m_bUseDBCSUI;
	BOOL			 m_bSysWinNT;
	BOOL			 m_bSysWin95;

	UINT			 WM_BUTTONINACTIVE;
	UINT			 WM_SIZEPARENT;
	UINT			 WM_ICONEDIT;
	UINT		     WM_HIDE;
	UINT		     WM_SETEDITFOCUS;
	UINT			 WM_BROADCAST;
	WORD			 m_nIDClipBandFormat;
	WORD			 m_nIDClipCategoryFormat;
	WORD			 m_nIDClipToolFormat;
	WORD			 m_nIDClipToolIdFormat;
	WORD			 m_nIDClipBandToolIdFormat;
	WORD			 m_nIDClipBandChildBandToolIdFormat;
	int				 m_nScrollInset;
	int				 m_nScrollDelay;
	int				 m_nScrollInterval;
	int				 m_nBitDepth;

private:
	CDragDropMgr*    m_pDragDropMgr;
	EventLog*        m_pEventLog;
	HICON			 m_hIconCancel;
};

Globals& GetGlobals();

#endif