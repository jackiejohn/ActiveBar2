#ifndef __GLOBALS_H__
#define __GLOBALS_H__

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

class FMap;
class CBar;
class CBand;
class CTool;
class CAppPref;
class CDragDropMgr;
class CLocalizer;
class CDesigner;
class EventLog;

#ifdef JAPBUILD

#define TABS_NORM_FONT L"‚l‚r ‚oƒSƒVƒbƒN\0\0"
#define TABS_ROT_FONT L"‚l‚r ‚oƒSƒVƒbƒN\0\0"
#define TABS_NORM_SIZE -12
#define TABS_ROT_SIZE -10

#else

#define TABS_NORM_FONT L"MS Sans Serif"
#define TABS_ROT_FONT L"Arial"
#define TABS_NORM_SIZE -8
#define TABS_ROT_SIZE -10

#endif

typedef BOOL (_stdcall *PLAYSOUNDFUNTION) (LPCTSTR, HMODULE, DWORD);
typedef BOOL (_stdcall *SETLAYEREDWINDOWATTRIBUTES) (HWND, COLORREF, BYTE, DWORD);

enum 
{  
	FS_SHOW = 0x01, 
	FS_HIDE = 0x02,
	FS_ACTIVATE = 0x04, 
	FS_DEACTIVATE = 0x08,
	FS_ENABLE = 0x10, 
	FS_DISABLE = 0x20,
	FS_SYNCACTIVE = 0x40 
};

#define WF_TOOLTIPS         0x0001  // window is enabled for tooltips
#define WF_TEMPHIDE         0x0002  // window is temporarily hidden
#define WF_STAYDISABLED     0x0004  // window should stay disabled
#define WF_MODALLOOP        0x0008  // currently in modal loop
#define WF_CONTINUEMODAL    0x0010  // modal loop should continue running
#define WF_OLECTLCONTAINER  0x0100  // some descendant is an OLE control
#define WF_TRACKINGTOOLTIPS 0x0400  // window is enabled for tracking tooltips
#define WF_STAYACTIVE       0x0020  // look active even though not active
#define WF_NOPOPMSG         0x0040  // ignore WM_POPMESSAGESTRING calls
#define WF_MODALDISABLE     0x0080  // window is disabled
#define WF_KEEPMINIACTIVE   0x0200  // stay activate even though you are deactivated

struct UpdateTabTool
{
	CBand* pBand;
	CTool* pTool;
	HWND   hWndBand;
	int    nTool;
} g_theUpdateTabTool;

struct Globals
{
	Globals();
	~Globals();

	void Reset();

	enum
	{
		PREFID_MAINWIN  = 1,
		PREFID_TOOLSWIN = 2,
		PREFID_ICONEDIT = 3,
		PREFID_BANDSWIN = 4,
		PREFID_LIBWIN = 5,
		PREFID_DOCWIN = 6,
		GrabStyleCount = 4,
	};

	PLAYSOUNDFUNTION GetPlaySound(); 
	SETLAYEREDWINDOWATTRIBUTES GetSetLayeredWindowAttributes();
	HIMAGELIST GetSizer();
	HIMAGELIST GetSlidingScrollImages();
	EventLog* GetEventLog();
	HBRUSH GetGrabHandleBrush(int nIndex);
	HFONT GetSmallCaptionFont();

	void ClearAccelators();

	HCURSOR GetSplitVCursor();
	HCURSOR GetSplitHCursor();
	HCURSOR GetHandCursor();
	
	HBITMAP GetBitmapMenuCheck();
	HBITMAP	GetBitmapSubMenu();
	HBITMAP GetBitmapSMCombo();
	HBITMAP GetBitmapMenuScroll();
	HBITMAP	GetBitmapEndMarker();
	HBITMAP GetBitmapExpandVert();
	HBITMAP GetBitmapExpandHorz();
	HBITMAP GetBitmapMenuExpand();
	HBITMAP GetBitmapMDIButtons();
	HBITMAP GetBrushPattern();
	
	HBRUSH GetBrushDither();
	
	void ClearGrabhandle();
	void EnterPaintIcon();
	void LeavePaintIcon();

	_se_translator_function m_seFunction;

	HINSTANCE m_hMMLibrary; 
	HINSTANCE m_hUser32Library; 
	
	FONTDESC _fdDefaultControl;
	FONTDESC _fdDefault;

	UINT WM_REFRESHMENUBAND;
	UINT WM_TOOLTIPSUPPORT;
	UINT WM_ACTIVEBARCLICK;
	UINT WM_ACTIVEBARTEXTCHANGE;
	UINT WM_ACTIVEBARCOMBOSELCHANGE;
	UINT WM_RECALCLAYOUT;
	UINT WM_POPUPWINMSG;
	UINT WM_POPUPKEYDOWN;
	UINT WM_SIZEPARENT;
	UINT WM_KILLWINDOW;
	UINT WM_FLOATSTATUS;
	UINT WM_POSTACTIVATE;
	UINT WM_UPDATETABTOOL;
	UINT WM_CACHEDOCKAREAS;
	UINT WM_SETTOOLFOCUS;

	CRITICAL_SECTION m_csPaintIcon;
	CRITICAL_SECTION m_csGlobals;

	CDragDropMgr* m_pDragDropMgr;
	CAppPref*	  m_pAppPref;
	short		  m_nBrushPatternWidth;
	short		  m_nBrushPatternHeight;
	DWORD		  m_nDragDelay;
	HHOOK		  m_hHookAccelator;
	HHOOK		  m_hHookMessage;
	HHOOK		  m_hHookMouse;
	FMap*		  m_pmapAccelator;
	FMap*		  m_pmapCustomize;
	FMap*		  m_pmapBar;
	FMap*		  m_pControls;
	WORD		  m_nIDClipBandFormat;
	WORD		  m_nIDClipCategoryFormat;
	WORD		  m_nIDClipToolFormat;
	WORD		  m_nIDClipToolIdFormat;
	WORD		  m_nIDClipBandToolIdFormat;
	WORD		  m_nIDClipBandChildBandToolIdFormat;
	BOOL		  m_bFullDrag;
	BOOL		  m_bUseDBCSUI;
	CBar*		  m_pCustomizeActiveBar;
	CBar*		  m_pWhatsThisHelpActiveBar;
	HWND		  m_hDlgCurrent;
	int			  m_nDragDist;
	int			  m_nBitDepth;
	int			  m_nSmallCaptionHeight;

	static TCHAR m_szAlt[25];
	static TCHAR m_szCtrl[25];
	static TCHAR m_szShift[25];

private:
	HIMAGELIST m_ilSlidingScrollImages;
	HIMAGELIST m_ilSizer;
	EventLog* m_pEventLog;
	HBITMAP	m_hBitmapMenuScroll;
	HBITMAP m_hBitmapMDIButtons;
	HBITMAP m_hBitmapExpandVert;
	HBITMAP m_hBitmapExpandHorz;
	HBITMAP m_hBitmapMenuExpand;
	HBITMAP	m_hBitmapEndMarker;
	HBITMAP m_hBitmapMenuCheck;
	HBITMAP m_hBitmapSubMenu;
	HBITMAP m_hBitmapSMCombo;
	HBITMAP m_hBrushPattern;
	HCURSOR m_hSplitVCursor;
	HCURSOR m_hSplitHCursor;
	HCURSOR m_hHandCursor;
	HBRUSH m_hBrushGrab[GrabStyleCount];
	HBRUSH m_hBrushDither;
	HFONT  m_hSmallCaption;
}; 

Globals& GetGlobals();

inline void Globals::EnterPaintIcon()
{
	EnterCriticalSection(&m_csPaintIcon);
}

inline void Globals::LeavePaintIcon()
{
	LeaveCriticalSection(&m_csPaintIcon);
}

void InitHeap();
void CleanupHeap();

#endif
