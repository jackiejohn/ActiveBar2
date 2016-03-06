#ifndef __BTABS_H__
#define __BTABS_H__

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

struct Tabs
{
	enum ButtonStyle
	{
		eNone,
		eTopOrLeft,
		eBottomOrRight
	};

	Tabs()
	{
		m_eButtonInfo = eNone;
	}
	CRect m_rcTab;
	ButtonStyle m_eButtonInfo;
};

#include "Interfaces.h"

class BTabs
{
public:
	BTabs();
	~BTabs();

	enum SLIDINGHITTEST
	{
		eSlidingNone,
		eSlidingTop,
		eSlidingBottom
	};

	enum AREA 
	{
		eTabTop,
		eTabBottom,
		eTabLeft,
		eTabRight,
		eTabSlidingVert,
		eTabSlidingHorz
	};

	enum SCROLLBUTTONSTYLES
	{
		eRightButton = 0,
		eLeftButton = 1,
		eBottomButton = 2,
		eTopButton = 3,
		eInactiveRightButton = 4,
		eInactiveLeftButton = 5,
		eInactiveBottomButton = 6,
		eInactiveTopButton = 7,
		ePressedRightButton = 8,
		ePressedLeftButton = 9,
		ePressedBottomButton = 10,
		ePressedTopButton = 11,
	};

	enum CONSTANTS
	{
		// between border and text
		TABSTRIP_HSPACE = 9, 
		eScrollButton = 5,
		eImageSize = 9,
		eBorderSpacing = 3,
		eTotalBorderSpacing = 6,
		eTotalImageSize = eImageSize + eTotalBorderSpacing
	};

	virtual int BTGetTabCount() PURE;
	virtual int BTGetVisibleTabCount() PURE;
	virtual void BTSetTabWidth(const int& index, const int& nWidth) PURE;
	virtual int& BTGetTabWidth(const int& index) PURE;
	virtual BSTR& BTGetTabText(const int& index) PURE;
	virtual CPictureHolder& BTGetTabPicture(const int& index) PURE;
	virtual BOOL BTTabVisible(const int& index) PURE;
	virtual int BTCalcEndLen() PURE;

	virtual AREA BTGetArea() PURE;
	virtual HFONT BTGetFont() PURE;
	virtual int BTGetFontHeight() PURE;
	virtual int ScrollButtonHit(POINT pt) PURE;

	void InteractivePaintScrollButton(HDC hDC, SCROLLBUTTONSTYLES sbsStyle);
	void SetTopPaintLock(const BOOL& bLock);
	void SetBottomPaintLock(const BOOL& bLock);

	void PaintTopScrollButton(HDC hDC);
	void SetTopButtonLocation(const int& nX, const int& nY);
	RECT& GetTopButtonLocation();
	void SetTopButtonStyle(SCROLLBUTTONSTYLES sbsTypes);
	SCROLLBUTTONSTYLES& GetTopButtonStyle();

	void PaintBottomScrollButton(HDC hDC);
	void SetBottomButtonLocation(const int& nX, const int& nY);
	RECT& GetBottomButtonLocation();
	void SetBottomButtonStyle(SCROLLBUTTONSTYLES sbsTypes);
	SCROLLBUTTONSTYLES& GetBottomButtonStyle();

	int BTCalcHit(const POINT& pt, const CRect& rcBounds);
	BOOL Draw(CBar* pBar, const HDC& hDC, const CRect& rcBounds, const CRect& rcInside, const POINT& ptPaintOffset);
	void BTSetCurSel(const short& nCurSel);
	short& BTGetCurSel(); 
	void BTGetRect(CRect& rc);
	int BTGetEndLen();

	COLORREF m_crPictureBackgroundMaskColor;
	COLORREF m_crBtnDarkShadow;
	COLORREF m_crBtnHighLight;
	COLORREF m_crBtnShadow;
	COLORREF m_crBackColor;
	COLORREF m_crForeColor;
	COLORREF m_crBtnText;
	COLORREF m_crBtnFace;
	
#pragma pack(push)
#pragma pack (1)
	struct BaseTabProperties
	{
		BaseTabProperties();

		CaptionAlignmentTypes m_caChildBands:4;
		Font3DTypes   m_fsFont3D:4;
		short m_nCurTab;
	} tpV1;
	BOOL m_bTopInteractivePaint:1;
	BOOL m_bBottomInteractivePaint:1;
	BOOL m_bBottomTabPainted:1;
#pragma pack(pop)
	short m_nPrevTab;
	
protected:

	BOOL DrawSlidingTab(CBar* pBar, AREA& nArea, HDC hDC, CRect rcTab, const CRect& rcInside, Tabs::ButtonStyle eStyle, LPTSTR szCaption, const POINT& ptPaintOffset, BSTR bstrCaption = NULL);

	void DrawSlidingTabs(AREA&  nArea,
						 HDC    hDC,
						 int&   nTab, 
						 int&   nSlidingHeight, 
						 CRect& rcPaint,
						 CRect  rcTab, 
						 int&   nNumOfTabs, 
						 BOOL&  bFirstSlidingRight, 
						 LPTSTR szCaption);
	SCROLLBUTTONSTYLES m_sbsTopStyle;
	SCROLLBUTTONSTYLES m_sbsBottomStyle;
	CRect m_rcTopButton;
	CRect m_rcBottomButton;
	Tabs* m_pTabs;
	BOOL m_bOffsetTab;
	int m_nEndLength;
	friend class CBand;
};

//
// BTGetEndLen()
//

inline int BTabs::BTGetEndLen()
{
	return m_nEndLength;
}

inline void BTabs::SetTopPaintLock(const BOOL& bLock) 
{
	m_bTopInteractivePaint = bLock;
}

inline void BTabs::SetBottomPaintLock(const BOOL& bLock) 
{
	m_bBottomInteractivePaint = bLock;
}

inline void BTabs::BTSetCurSel(const short& nCurTab) 
{
	m_nPrevTab = tpV1.m_nCurTab;
	tpV1.m_nCurTab = nCurTab;
}

inline short& BTabs::BTGetCurSel() 
{
	return tpV1.m_nCurTab;
}

inline void BTabs::SetTopButtonLocation(const int& nX, const int& nY)
{
	m_rcTopButton.Set(nX, 
					  nY, 
					  nX + eImageSize + eTotalBorderSpacing, 
					  nY + eImageSize + eTotalBorderSpacing);
}

inline RECT& BTabs::GetTopButtonLocation()
{
	return m_rcTopButton;
}

inline void BTabs::SetTopButtonStyle(SCROLLBUTTONSTYLES sbsTypes)
{
	m_sbsTopStyle = sbsTypes;
}

inline BTabs::SCROLLBUTTONSTYLES& BTabs::GetTopButtonStyle()
{
	return m_sbsTopStyle;
}

inline void BTabs::SetBottomButtonLocation(const int& nX, const int& nY)
{
	m_rcBottomButton.Set(nX, 
						 nY, 
						 nX + eImageSize + eTotalBorderSpacing, 
						 nY + eImageSize + eTotalBorderSpacing);
}

inline RECT& BTabs::GetBottomButtonLocation()
{
	return m_rcBottomButton;
}

inline void BTabs::SetBottomButtonStyle(SCROLLBUTTONSTYLES sbsTypes)
{
	m_sbsBottomStyle = sbsTypes;
}

inline BTabs::SCROLLBUTTONSTYLES& BTabs::GetBottomButtonStyle()
{
	return m_sbsBottomStyle;
}
#endif
