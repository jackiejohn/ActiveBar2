#ifndef __TOOLTIP_H__
#define __TOOLTIP_H__

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

class CToolTip : public FWnd
{
public:
	CToolTip(CBar* pBar);
	virtual ~CToolTip();

	BOOL CreateWin();
	void SetFont(HFONT hFont);
	void Show(DWORD dwCookie, BSTR bstrText, const CRect& rcBound, BOOL bTopMost);
	void Hide();

	DWORD m_dwCloseTime;

private:
	enum TOOLTIPTIMER
	{
		eIdle,
		eTimerStarted,
		eTimerCompleted,
		eHiddenByDelay
	} m_tttMode; 

	enum TIMER
	{
		ID_TIMER = 407,
		TIMER_DURATION = 600
	};

	LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	
	void OnDraw(HDC hDC, CRect& rcBound);
	
	void ResetTimer();
	void Open();

	CBar* m_pBar;
	DWORD m_dwOpenTime;
	DWORD m_dwPrevCookie;
	DWORD m_dwWaitTime;
	HFONT m_hFont;
	POINT m_pt;
	CRect m_rcBound;
	BSTR  m_bstrToolTip;
	BOOL  m_bCreateMode;	
	BOOL  m_bAnimated;
	BOOL  m_bTopMost;
	HWND  m_hWndBound;
};

inline void CToolTip::SetFont(HFONT hFont) 
{
	m_hFont = hFont;
};

#endif