#ifndef DIALOGS_INCLUDED
#define DIALOGS_INCLUDED

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "Resource.h"
#include "..\Interfaces.h"
#include "..\PrivateInterfaces.h"

interface ITool;
struct CategoryMgr;
class CDesignerTree;
class CCategoryTree;
class DDToolBar;
class CSplitter;
class CIconEdit;
class CBrowser;
class CDesigner;

//
// CDlgOptions
//

class CDlgOptions: public FDialog
{
public:
	CDlgOptions();

private:
	virtual BOOL DialogProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	virtual void OnOK();
};

//
// CDlgConfirm
//

class CDlgConfirm : public FDialog
{
public:
	CDlgConfirm::CDlgConfirm(int nStringId);
	UINT Show(HWND hWndParent);	

private:
	virtual BOOL DialogProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	void OnOK();
	void OnCancel();

	BOOL m_bDontShow;
	int  m_nStringId;
};

inline CDlgConfirm::CDlgConfirm(int nStringId)
	:FDialog(IDD_CONFIRM)
{
	m_nStringId = nStringId;
}

//
// CIconEditor
//

class CIconEditor :public FDialog
{
public:
	enum 
	{
		eNumOfTools = 4
	};

	CIconEditor(ITool* pTool);
	~CIconEditor();

	CIconEdit* m_pIconEditor;

private:
	virtual BOOL DialogProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	virtual BOOL PreTranslateMessage(MSG* msg);

	void SetBitmaps();
	void OnSize(int nWidth, int nHeight);
	virtual void OnOK();
	
	ITool* m_pTool;
};

//
// CDlgLibrary
//

class CDlgLibrary :public FDialog
{
public:
	CDlgLibrary();
	~CDlgLibrary();

private:
	virtual BOOL DialogProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	virtual BOOL PreTranslateMessage(MSG* msg);

	void OnSize();
	void OnLButtonDown(UINT nFlags, POINT pt);
	void OnMouseMove(UINT nFlags, POINT pt);
	void OnLButtonUp(UINT nFlags, POINT pt);

	CDesignerTree* m_pDesignerTree;
	CategoryMgr*   m_pCategoryMgr;
	IActiveBar2*   m_pActiveBar;
	CSplitter*     m_pSplitter;
	DDToolBar*	   m_pToolBar;
	CBrowser*	   m_pBrowser;
	BOOL		   m_bDirty;
};

//
// CDlgDesigner
//

class CDlgDesigner : public FDialog, public IDesignerNotify 
{
public:
	CDlgDesigner();
	~CDlgDesigner();

	void Init(IActiveBar2* pActiveBar, BOOL bStandALone);
    STDMETHOD(CreateTool)(IDispatch* pBand, long nToolType);
    STDMETHOD(SetDirty)();
	void Apply();

private:
	virtual BOOL DialogProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	virtual BOOL PreTranslateMessage(MSG* msg);

	void OnSize();
	void OnLButtonDown(UINT nFlags, POINT pt);
	void OnMouseMove(UINT nFlags, POINT pt);
	void OnLButtonUp(UINT nFlags, POINT pt);
	void OnOptions();
	void OnReport();
	void OnHeader();
	void OnLibrary();
	void OnMenuGrab();
	void OnGenSelect();
	void OnVerifyImageReferenceCount();

	CDesignerTree* m_pDesignerTree;
	CategoryMgr*   m_pCategoryMgr;
	IActiveBar2*   m_pActiveBar;
	CSplitter*     m_pSplitter;
	DDToolBar*	   m_pToolBar;
	CBrowser*	   m_pBrowser;
	VARIANT		   m_vLayoutData;
	BOOL		   m_bStandALone;
	BOOL		   m_bFileDirty;
	BOOL		   m_bDirty;
};

//
// CColor
//

class CColor
{
public:
	CColor ();
	CColor (COLORREF crColor);

	enum
	{
		eLSMAX = 240,
		eHUEMAX = 240,
		eRGBMAX = 255,
		eUNDEFINED = (eLSMAX * 2 / 3)
	};

	void GetHLS(short& nHue, short& nLum, short& nSat);
	void SetHLS(short nHue, short nLum, short nSat);
	
	void GetRGB(short& nRed, short& nGreen, short& nBlue);
	void SetRGB(short nRed, short nGreen, short nBlue);

	COLORREF Color();
	void Color(COLORREF crColor);

private:
	void RGBToHLS();
	void HLSToRGB();

	short m_nRed;
	short m_nGreen;
	short m_nBlue;
	short m_nHue;
	short m_nLum;
	short m_nSat;
};

//
// Color
// 

inline COLORREF CColor::Color()
{
	return RGB(m_nRed, m_nGreen, m_nBlue);
}

inline void CColor::Color(COLORREF crColor)
{
	m_nRed = GetRValue(crColor);
	m_nGreen = GetGValue(crColor);
	m_nBlue = GetBValue(crColor);
	RGBToHLS();
}

//
// GetHLS
//

inline void CColor::GetHLS(short& nHue, short& nLum, short& nSat)
{
	nHue = m_nHue;
	nLum = m_nLum;
	nSat = m_nSat;
}

//
// SetHLS
//

inline void CColor::SetHLS(short nHue, short nLum, short nSat)
{
	m_nHue = nHue;
	m_nLum = nLum;
	m_nSat = nSat;
	HLSToRGB();
}

//
// GetRGB
//

inline void CColor::GetRGB(short& nRed, short& nGreen, short& nBlue)
{
	nRed = m_nRed;
	nGreen = m_nGreen;
	nBlue = m_nBlue;
}

//
// SetRGB
//

inline void CColor::SetRGB(short nRed, short nGreen, short nBlue)
{
	m_nRed = nRed;
	m_nGreen = nGreen;
	m_nBlue = nBlue;
	RGBToHLS();
}

//
// CColorDisplay
//

class CColorDisplay : public FWnd
{
public:
	CColorDisplay();

	void Color(const COLORREF& crColor);

private:
	LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	
	COLORREF m_crColor;
};

inline void CColorDisplay::Color(const COLORREF& crColor)
{
	m_crColor = crColor;
	InvalidateRect(NULL, FALSE);
}

//
// CColorSpectrum
//

class CColorSpectrum : public FWnd
{
public:
	CColorSpectrum();
	~CColorSpectrum();

	void Color(const COLORREF& crColor);
	void Initialize();

private:
	LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	void Draw(HDC hDC, CRect& rcBound);

	CFlickerFree m_ffDraw;
	COLORREF     m_crColor;
	BOOL		 m_bCreated;
	int			 m_nWidth;
	int			 m_nHeight;
};

//
// SetColor
//

inline void CColorSpectrum::Color(const COLORREF& crColor)
{
	m_crColor = crColor;
	InvalidateRect(NULL, FALSE);
}

//
// CLuminosity
//

class CLuminosity : public FWnd
{
public:
	CLuminosity();
	~CLuminosity();

	void Color(const COLORREF& crColor);

private:
	LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	void Draw(HDC hDC, CRect& rcBound);

	CFlickerFree m_ffDraw;
	COLORREF     m_crColor;
};

//
// SetColor
//

inline void CLuminosity::Color(const COLORREF& crColor)
{
	m_crColor = crColor;
	InvalidateRect(NULL, FALSE);
}

//
// CDefineColor
//

class CDefineColor :public FDialog
{
public:
	CDefineColor();
	~CDefineColor();

	COLORREF Color();
	void Color(COLORREF& crColor);

private:
	virtual BOOL DialogProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	void OnLButtonDown(WPARAM wParam, const POINT& pt);
	void OnAddColor();
	void OnClose();
	void Draw(HDC hDC, CRect& rcBound);

	void SetRGB();
	void GetRGB();
	void SetHLS();
	void GetHLS();

//	CColorSpectrum m_wndColorSpectrum;
	CColorDisplay  m_wndColorDisplay;
	CLuminosity    m_wndLum;
	CColor		   m_cColor;	
	HCURSOR		   m_hCursorLuminosity;
	POINT		   m_ptLuminosity;
	BOOL		   m_bLock;
};

//
// Color
//

inline COLORREF CDefineColor::Color()
{
	return m_cColor.Color();
}

#endif