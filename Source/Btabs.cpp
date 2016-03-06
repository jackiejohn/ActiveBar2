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
#include "IPServer.h"
#include "Resource.h"
#include "Support.h"
#include "Globals.h"
#include "Bar.h"
#ifdef _SCRIPTSTRING
#include "ScriptString.h"
#endif
#include "BTabs.h"

extern BOOL g_fSysWinNT;

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

//
// CalcStringEllipsis
//
// Determines whether the specified string is too wide to fit in
// an allotted space, and if not truncates the string and adds some
// points of ellipsis to the end of the string.
//

static BOOL CalcStringEllipsis(HDC& hDC, LPTSTR szString, int nColWidth)
{
    SIZE   sizeEllipsis;
    SIZE   sizeString;
    TCHAR* szEllipsis = _T("...");
    TCHAR* szTemp;
	int	   nString;

    // Adjust the column width to take into account the edges
    nColWidth -= 4;

    // Allocate a string for us to work with.  This way we can mangle the
    // string and still preserve the return value
	nString = lstrlen(szString);
    szTemp = new TCHAR[nString+1];
	if (NULL == szTemp)
		return FALSE;

	lstrcpy(szTemp, szString);

	CRect rcOut;
	int nResult;
	nResult = DrawText(hDC, szTemp, nString, &rcOut, DT_LEFT|DT_CALCRECT);
    if (0 == nResult)
	{
		delete [] szTemp;
		lstrcpy(szString, szEllipsis);
		return FALSE;
	}
	
	sizeString = rcOut.Size();

    // If the width of the string is greater than the column width shave
    // the string and add the ellipsis
    if (sizeString.cx <= nColWidth)
	{
		delete [] szTemp;
		return TRUE;
	}

	if (0 == DrawText(hDC, szEllipsis, lstrlen(szEllipsis), &rcOut, DT_LEFT|DT_CALCRECT))
	{
		delete [] szTemp;
		lstrcpy(szString, szEllipsis);
		return FALSE;
	}
	sizeEllipsis = rcOut.Size();
    
	while (nString > 0)
    {
		nString = lstrlen(szTemp)-1;
		szTemp[nString] = NULL;
        if (0 == DrawText(hDC, szTemp, nString, &rcOut, DT_LEFT|DT_CALCRECT))
		{
			delete [] szTemp;
			lstrcpy(szString, szEllipsis);
			return FALSE;
		}
		sizeString = rcOut.Size();

        if (sizeString.cx + sizeEllipsis.cx <= nColWidth)
        {
			// The string with the ellipsis finally fits, now make sure
            // there is enough room in the string for the ellipsis
            if (lstrlen(szString) >= (nString + lstrlen(szEllipsis)))
            {
				// Concatenate the two strings and break out of the loop
                lstrcat(szTemp, szEllipsis);
				lstrcpy(szString, szTemp);
				delete [] szTemp;
                return TRUE;
            }
        }
	} 
	delete [] szTemp;
	lstrcpy(szString, szEllipsis);
	return TRUE;
}

//
// BTabs
//

BTabs::BTabs()
	: m_pTabs(NULL)
{
	m_nPrevTab = -1;
	m_sbsTopStyle = eInactiveLeftButton;
	m_sbsBottomStyle = eInactiveRightButton;
	m_bTopInteractivePaint = FALSE;
	m_bBottomInteractivePaint = FALSE;
	m_nEndLength = 0;
	m_bOffsetTab = FALSE;
}

BTabs::~BTabs()
{
	delete [] m_pTabs;
}

BTabs::BaseTabProperties::BaseTabProperties()
{
	m_caChildBands = ddCALeft;
	m_fsFont3D = dd3DNone;
	m_nCurTab = -1;
}

//
// BTCalcHit
//

int BTabs::BTCalcHit(const POINT& pt, const CRect& rcBound)
{
	try
	{
		POINT ptTab(pt);
		RECT rcTab(rcBound);
		int nNumOfTabs = BTGetTabCount();
		int nStart=0;
		int nEnd=0;
		int nNonVisible = 0;

		AREA aTab = BTGetArea();
		switch (aTab)
		{
		case eTabSlidingHorz:
		case eTabSlidingVert:
			{
				CRect rcTemp;
				for (int nTab = 0; nTab < nNumOfTabs; nTab++)
				{
					if (!BTTabVisible(nTab))
						continue;

					rcTemp = m_pTabs[nTab].m_rcTab;
					rcTemp.Offset(rcBound.left, rcBound.top);
					if (PtInRect(&rcTemp, pt))
						return nTab;
				}
			}
			break;

		default:
			{
				BOOL bVert = !(eTabTop == aTab || eTabBottom == aTab);
				if (bVert)
				{
					ptTab.y -= rcBound.top;
					rcTab.bottom -= rcTab.top;
				}
				else
				{
					ptTab.x -= rcBound.left;
					rcTab.right -= rcTab.left;
				}

				for (int nTab = 0; nTab < nNumOfTabs; nTab++)
				{
					if (!BTTabVisible(nTab))
						continue;

					nEnd = BTGetTabWidth(nTab) + 2 + 2 + (TABSTRIP_HSPACE*2) + nStart;
					if (bVert)	
					{
						if (ptTab.y >= nStart && ptTab.y <= nEnd)
							return nTab;
					}
					else
					{
						if (ptTab.x >= nStart && ptTab.x <= nEnd)
							return nTab;
					}
					nStart = nEnd;
				}
			}
			break;
		}
	}
	catch (...)
	{
		assert(FALSE);
	}
	return -1;
}

//
// Draw
//

BOOL BTabs::Draw(CBar* pBar, const HDC& hDC, const CRect& rcPaint, const CRect& rcInside, const POINT& ptPaintOffset)
{
	BOOL bResult = TRUE;
	try
	{
		SetBkMode(hDC, TRANSPARENT);
		SetTextColor(hDC, m_crBtnText);
		
		HFONT hFontOld = SelectFont(hDC, BTGetFont());
		int nNumOfTabs = BTGetTabCount();
		
		CRect rcTab;
		BSTR bstrCaption;
		AREA nArea = BTGetArea();
		switch (nArea)
		{
		case eTabSlidingHorz:
		case eTabSlidingVert:
			{
				if (NULL == m_pTabs)
					break;

				m_bBottomTabPainted = FALSE;
				CRect rcTemp;
				COLORREF crGap = GetSysColor(COLOR_3DDKSHADOW);
				for (int nTab = 0; nTab < nNumOfTabs; ++nTab)
				{
					if (!BTTabVisible(nTab))
						continue;

					bstrCaption = BTGetTabText(nTab);
					MAKE_TCHARPTR_FROMWIDE(szCaption, bstrCaption);
					rcTab = m_pTabs[nTab].m_rcTab;
					rcTab.Offset(rcPaint.left, rcPaint.top);
					DrawSlidingTab(pBar, nArea, hDC, rcTab, rcInside, m_pTabs[nTab].m_eButtonInfo, szCaption, ptPaintOffset, bstrCaption);
					rcTemp = rcTab;
					if (eTabSlidingHorz == nArea)
					{
						rcTemp.left = rcTemp.right;
						rcTemp.right++;
					}
					else
					{
						rcTemp.top = rcTemp.bottom;
						rcTemp.bottom++;
					}
					FillSolidRect(hDC, rcTemp, crGap);
				}
				if (!m_bBottomTabPainted)
				{
					int nLineHeight = BTGetFontHeight();
					m_rcBottomButton = rcPaint;
					m_rcBottomButton.Offset(ptPaintOffset.x, ptPaintOffset.y);
					switch (nArea)
					{
					case eTabSlidingHorz:
						m_rcBottomButton.top = m_rcBottomButton.bottom - eTotalImageSize;
						m_rcBottomButton.left = m_rcBottomButton.right - eTotalImageSize;
						m_rcBottomButton.Offset(-eScrollButton, -eScrollButton);
						break;
					case eTabSlidingVert:
						m_rcBottomButton.top = m_rcBottomButton.bottom - eTotalImageSize;
						m_rcBottomButton.left = m_rcBottomButton.right - eTotalImageSize;
						m_rcBottomButton.Offset(-eScrollButton, -eScrollButton);
						break;
					}
					PaintBottomScrollButton(hDC);
				}
			}
			break;

		default:
			{
				CRect rc = rcPaint;
				CRect rcRender;
				int nItemSize;
				int nX = rcPaint.left;
				int nY = rcPaint.top;
				int nHeight = rcPaint.Height();
				int nWidth = rcPaint.Width();
				int nSlidingHeight = BTGetFontHeight();
				for (int nTab = 0; nTab < nNumOfTabs; ++nTab)
				{
					if (!BTTabVisible(nTab))
						continue;

					CPictureHolder& thePicture = BTGetTabPicture(nTab);

					nHeight = rcPaint.Height();
					nWidth = rcPaint.Width();
					nItemSize = BTGetTabWidth(nTab) + 2 + 2 + (TABSTRIP_HSPACE*2);
					if (nTab != tpV1.m_nCurTab)
					{
						nWidth -= 2;
						nHeight -= 2;
					}
					
					bstrCaption = BTGetTabText(nTab);
					MAKE_TCHARPTR_FROMWIDE(szCaption, bstrCaption);
					switch (nArea)
					{
					case eTabBottom: // bottom tabs
						FillSolidRect(hDC, nX + nItemSize - 2, nY, 1, nHeight, m_crBtnShadow);	// right shadow
						FillSolidRect(hDC, nX + nItemSize - 1, nY, 1, nHeight - 1, m_crBtnDarkShadow);
						FillSolidRect(hDC, nX, nY + nHeight - 2, nItemSize - 2, 1, m_crBtnShadow);  // bottom shadow
						FillSolidRect(hDC, nX+2, nY + nHeight - 1, nItemSize - 3, 1, m_crBtnDarkShadow);
						FillSolidRect(hDC, nX, nY, 1, nHeight - 2, m_crBtnHighLight); // left highlight
						
						if (thePicture.m_pPict)
						{
							rcRender.Set(nX + 2, nY + 2, nX + 2 + 16, nY + 2 + 16);
							thePicture.Render(hDC, rcRender, rcRender);
						}

#ifdef _SCRIPTSTRING
						if (pBar && VARIANT_TRUE == pBar->bpV1.m_vbUseUnicode)
							ScriptTextOut(hDC,
										  nX + TABSTRIP_HSPACE + (thePicture.m_pPict ? 20 : 0),
										  nY + 1 + 2 * (nTab == tpV1.m_nCurTab),
										  bstrCaption,
										  wcslen(bstrCaption));
						else
#endif
							TextOut(hDC,
									nX + TABSTRIP_HSPACE + (thePicture.m_pPict ? 20 : 0),
									nY + 1 + 2 * (nTab == tpV1.m_nCurTab),
									szCaption,
									lstrlen(szCaption));

						if (nTab != tpV1.m_nCurTab)
							FillSolidRect(hDC, nX, nY, nItemSize, 1, m_crBtnShadow); // top line
						else
							FillSolidRect(hDC, nX + 1, nY, nItemSize - 3, 1, m_crBtnFace); // top line

						nX += nItemSize;
						break;

					case eTabTop:
						if (nTab != tpV1.m_nCurTab)
							nY += 2;

						FillSolidRect(hDC, nX, nY + 1, 1, nHeight - 1, m_crBtnHighLight); // left highlight
						FillSolidRect(hDC, nX, nY, nItemSize - 1, 1, m_crBtnHighLight); // top highlight
						FillSolidRect(hDC, nX + nItemSize - 2, nY+1, 1, nHeight - 1, m_crBtnShadow); // right shadow
						FillSolidRect(hDC, nX + nItemSize - 1, nY+2, 1, nHeight - 2, m_crBtnDarkShadow);

						if (thePicture.m_pPict)
						{
							rcRender.Set(nX + 2, nY + 2, nX + 2 + 16, nY + 2 + 16);
							thePicture.Render(hDC, rcRender, rcRender);
						}

#ifdef _SCRIPTSTRING
						if (pBar && VARIANT_TRUE == pBar->bpV1.m_vbUseUnicode)
							ScriptTextOut(hDC,
										  nX + TABSTRIP_HSPACE + (thePicture.m_pPict ? 20 : 0),
										  nY + 1 + 2 * (nTab == tpV1.m_nCurTab),
										  bstrCaption,
										  wcslen(bstrCaption));
						else
#endif
							TextOut(hDC,
									nX + TABSTRIP_HSPACE + (thePicture.m_pPict ? 20 : 0),
									nY + 1 + 2 * (nTab == tpV1.m_nCurTab),
									szCaption,
									lstrlen(szCaption));

						if (nTab != tpV1.m_nCurTab)
						{
							// Bottom Line
							FillSolidRect(hDC, nX, nY + nHeight, nItemSize, 1, m_crBtnHighLight); 
							nY -= 2;
						}
						else
							// Bottom Line
							FillSolidRect(hDC, nX+1, nY + nHeight, nItemSize - 2, 1, m_crBtnFace); 

						nX += nItemSize;
						break;

					case eTabLeft:
						if (nTab != tpV1.m_nCurTab)
							nX += 2;

						FillSolidRect(hDC, nX+1, nY, nWidth-1, 1, m_crBtnHighLight);			// left highlight
						FillSolidRect(hDC, nX,nY, 1, nItemSize-1, m_crBtnHighLight);			// top highlight
						FillSolidRect(hDC, nX+1, nY+nItemSize-2, nWidth-1, 1, m_crBtnShadow);	// right shadow
						FillSolidRect(hDC, nX+2, nY+nItemSize-1, nWidth-2, 1, m_crBtnDarkShadow);

						if (thePicture.m_pPict)
						{
							rcRender.Set(nX + 2, nY + 2, nX + 2 + 16, nY + 2 + 16);
							thePicture.Render(hDC, rcRender, rcRender);
						}

#ifdef _SCRIPTSTRING
						if (pBar && VARIANT_TRUE == pBar->bpV1.m_vbUseUnicode)
							ScriptTextOut(hDC,
										  nX - 2 * (nTab == tpV1.m_nCurTab) + nWidth - 1,
										  nY + TABSTRIP_HSPACE + (thePicture.m_pPict ? 20 : 0),
										  bstrCaption,
										  wcslen(bstrCaption));
						else
#endif
							TextOut(hDC,
									nX - 2 * (nTab == tpV1.m_nCurTab) + nWidth - 1,
									nY + TABSTRIP_HSPACE + (thePicture.m_pPict ? 20 : 0),
									szCaption,
									lstrlen(szCaption));
						
						if (nTab != tpV1.m_nCurTab)
							FillSolidRect(hDC, nX+nWidth, nY, 1, nItemSize, m_crBtnHighLight); // bottom line
						else
							FillSolidRect(hDC, nX+nWidth, nY+1, 1, nItemSize-2, m_crBtnFace); // bottom line

						if (nTab != tpV1.m_nCurTab)
							nX -= 2;
						nY += nItemSize;
						break;

					case eTabRight:
						if (nTab != tpV1.m_nCurTab)
							nWidth -= 2;

						FillSolidRect(hDC, nX, nY+nItemSize - 2, nWidth, 1, m_crBtnShadow);	// right shadow
						FillSolidRect(hDC, nX, nY+nItemSize - 1, nWidth-1, 1, m_crBtnDarkShadow);
						FillSolidRect(hDC, nX + nWidth - 2, nY, 1, nItemSize-2, m_crBtnShadow);  // bottom shadow
						FillSolidRect(hDC, nX + nWidth-1, nY + 2, 1, nItemSize-3, m_crBtnDarkShadow);
						FillSolidRect(hDC, nX, nY, nWidth-2, 1, m_crBtnHighLight);			// left highlight

						if (thePicture.m_pPict)
						{
							rcRender.Set(nX + 2, nY + 2, nX + 2 + 16, nY + 2 + 16);
							thePicture.Render(hDC, rcRender, rcRender);
						}

#ifdef _SCRIPTSTRING
						if (pBar && VARIANT_TRUE == pBar->bpV1.m_vbUseUnicode)
							ScriptTextOut(hDC,
										  nX + nWidth - 1 - 2 * (nTab == tpV1.m_nCurTab),
										  nY + TABSTRIP_HSPACE + (thePicture.m_pPict ? 20 : 0),
										  bstrCaption,
										  wcslen(bstrCaption));
						else
#endif
							TextOut(hDC,
									nX + nWidth - 1 - 2 * (nTab == tpV1.m_nCurTab),
									nY + TABSTRIP_HSPACE + (thePicture.m_pPict ? 20 : 0),
									szCaption,
									lstrlen(szCaption));

						if (nTab != tpV1.m_nCurTab)
							FillSolidRect(hDC, nX, nY, 1, nItemSize, m_crBtnShadow); // top line
						else
							FillSolidRect(hDC, nX, nY+1, 1, nItemSize - 3, m_crBtnFace); // top line

						nY += nItemSize;
						break;
					}
				}
				
				switch (nArea)
				{
				case eTabBottom:
					FillSolidRect(hDC,
								  nX,
								  nY,
								  rcPaint.right-nX,
								  1,
								  m_crBtnShadow);
					break;

				case eTabLeft:
					FillSolidRect(hDC,
								  nX+rcPaint.Width(),
								  nY,
								  1,
								  rcPaint.bottom-nY,
								  m_crBtnHighLight); 
					break;

				case eTabRight:
					FillSolidRect(hDC,
								  nX,
								  nY,
								  1,
								  rcPaint.bottom-nY,
								  m_crBtnShadow);
					break;
				}
			}
			break;
		}
		SelectFont(hDC, hFontOld);
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
// BTGetRect
//

void BTabs::BTGetRect(CRect& rc)
{
	rc.SetEmpty();
	AREA nArea = BTGetArea();
	int nLineHeight = BTGetFontHeight();
	int nNumOfTabs = 0;

	//
	// Now return size
	//

	switch (nArea)
	{
	case eTabSlidingVert:
		nNumOfTabs = BTGetVisibleTabCount();
		rc.top = (nLineHeight + 1) * nNumOfTabs;
		return;

	case eTabSlidingHorz:
		nNumOfTabs = BTGetVisibleTabCount();
		rc.left = (nLineHeight + 1) * nNumOfTabs;
		return;
	}

	nNumOfTabs = BTGetTabCount();
	HDC hDC = GetDC(NULL);
	if (hDC)
	{
		HFONT hFontOld = SelectFont(hDC, BTGetFont());
		BSTR bstr;
		SIZE sizeTab;
		for (int nTab = 0; nTab < nNumOfTabs; nTab++)
		{
			if (!BTTabVisible(nTab))
				continue;

			CPictureHolder& thePicture = BTGetTabPicture(nTab);
			if (thePicture.m_pPict)
				nLineHeight = max(nLineHeight, 20);

			bstr = BTGetTabText(nTab);
			if (bstr && GetTextExtentPoint32W(hDC, bstr, lstrlenW(bstr), &sizeTab))
				BTSetTabWidth(nTab, sizeTab.cx + (thePicture.m_pPict ? 20 : 0));
			else
				BTSetTabWidth(nTab, thePicture.m_pPict ? 20 : 0);
		}
		SelectFont(hDC, hFontOld);
		ReleaseDC(NULL, hDC);
	}

	//
	// Now return size
	//

	switch (nArea)
	{
	case eTabTop:
		rc.top = nLineHeight;
		break;

	case eTabBottom:
		rc.bottom = nLineHeight;
		break;

	case eTabLeft:
		rc.left = nLineHeight;
		break;

	case eTabRight:
		rc.right = nLineHeight;
		break;
	}
}

//
// DrawSlidingTab
//

BOOL BTabs::DrawSlidingTab(CBar* pBar, AREA& nArea, HDC hDC, CRect rcTab, const CRect& rcInside, Tabs::ButtonStyle eStyle, LPTSTR szCaption, const POINT& ptPaintOffset, BSTR bstrCaption)
{
	CRect rcPaint = rcTab;
	POINT ptText;
	int nWidth; 
	CRect rcText;
	switch (nArea)
	{
	case eTabSlidingVert:
		switch (eStyle)
		{
		case Tabs::eTopOrLeft:
			m_rcTopButton = rcTab;
			m_rcTopButton.left = m_rcTopButton.right - eTotalImageSize;
			m_rcTopButton.top = m_rcTopButton.bottom - eTotalImageSize;
			if (eTabSlidingHorz == nArea)
				m_rcTopButton.Offset(rcTab.Width() + eScrollButton, -eScrollButton);
			else
				m_rcTopButton.Offset(-eScrollButton, rcTab.Height() + eScrollButton);
			if (m_rcTopButton.left >= rcInside.left &&  m_rcTopButton.right <= rcInside.right &&  
				m_rcTopButton.top >= rcInside.top &&  m_rcTopButton.bottom <= rcInside.bottom)
				PaintTopScrollButton(hDC);
			break;

		case Tabs::eBottomOrRight:
			m_rcBottomButton = rcTab;
			m_rcBottomButton.Offset(ptPaintOffset.x, ptPaintOffset.y);
			m_rcBottomButton.left = m_rcBottomButton.right - eTotalImageSize;
			m_rcBottomButton.top = m_rcBottomButton.bottom - eTotalImageSize;
			if (eTabSlidingHorz == nArea)
				m_rcBottomButton.Offset(-rcTab.Width()-eScrollButton, -eScrollButton);
			else
				m_rcBottomButton.Offset(-eScrollButton, -rcTab.Height()-eScrollButton);
			if (m_rcBottomButton.left >= rcInside.left &&  m_rcBottomButton.right <= rcInside.right &&  
				m_rcBottomButton.top >= rcInside.top &&  m_rcBottomButton.bottom <= rcInside.bottom)
			PaintBottomScrollButton(hDC);
			m_bBottomTabPainted = TRUE;
			break;
		}
		rcText = rcTab;
		rcText.Inflate(-2, -2);
		ptText.x = rcText.left;
		ptText.y = rcText.top;
		nWidth = rcText.Width();
		break;
	
	case eTabSlidingHorz:
		switch (eStyle)
		{
		case Tabs::eTopOrLeft:
			m_rcTopButton = rcTab;
			m_rcTopButton.top = m_rcTopButton.bottom - eTotalImageSize;
			m_rcTopButton.left = m_rcTopButton.right - eTotalImageSize;
			if (eTabSlidingHorz == nArea)
				m_rcTopButton.Offset(rcTab.Width()+eScrollButton, -eScrollButton);
			else
				m_rcTopButton.Offset(-eScrollButton, rcTab.Height()+eScrollButton);
			PaintTopScrollButton(hDC);
			break;

		case Tabs::eBottomOrRight:
			m_rcBottomButton = rcTab;
			m_rcBottomButton.top = m_rcBottomButton.bottom - eTotalImageSize;
			m_rcBottomButton.left = m_rcBottomButton.right - eTotalImageSize;
			if (eTabSlidingHorz == nArea)
				m_rcBottomButton.Offset(-rcTab.Width() - eScrollButton, -eScrollButton);
			else
				m_rcBottomButton.Offset(-eScrollButton, -rcTab.Height() - eScrollButton);
			PaintBottomScrollButton(hDC);
			m_bBottomTabPainted = TRUE;
			break;
		}
		rcText = rcTab;
		rcText.Inflate(-2, -2);
		ptText.x = rcText.right + 2;
		ptText.y = rcText.top + 1;
		nWidth = rcText.Height();
		break;
	}

	FillSolidRect(hDC, rcTab, m_crBtnFace);
	pBar->DrawEdge(hDC, rcTab, BDR_RAISEDINNER, BF_RECT);

	SetTextColor(hDC, m_crForeColor);
	SetBkColor(hDC, m_crBtnFace);

	BOOL bResult = CalcStringEllipsis(hDC, szCaption, nWidth);
	int nCaptionLen = lstrlen(szCaption);

	SIZE sizeString;
	int nDrawStrLen = 0;
	if (GetTextExtentPoint32(hDC, szCaption, nCaptionLen, &sizeString))
		nDrawStrLen = sizeString.cx;

	switch (tpV1.m_caChildBands)
	{
	case ddCACenter:
		if (nArea == eTabSlidingVert)
			ptText.x = rcText.left + ((rcPaint.Width() - nDrawStrLen - 4) / 2);
		else
			ptText.y = rcText.top + ((rcPaint.Height() - nDrawStrLen - 4) / 2);
		break;

	case ddCARight:
		if (nArea == eTabSlidingVert)
			ptText.x = rcText.right - nDrawStrLen - 4;
		else
			ptText.y = rcText.bottom - nDrawStrLen - 4;
		break;
	}

	rcTab;
	if (eTabSlidingVert == nArea)
		ptText.y = ptText.y + ((rcTab.Height() - sizeString.cy) / 2);
	else
		ptText.x = ptText.x - ((rcTab.Width() - sizeString.cy) / 2);

	switch (tpV1.m_fsFont3D)
	{
	case dd3DNone:
		if (eTabSlidingVert == nArea)
			ptText.y -= 1;
		else
		{
			ptText.y += 2;
			ptText.x -= 2;
		}

		SetTextColor(hDC, m_crForeColor);
#ifdef _SCRIPTSTRING
		if (pBar && VARIANT_TRUE == pBar->bpV1.m_vbUseUnicode)
			ScriptExtTextOut(hDC, 
						     ptText.x, 
						     ptText.y, 
						     ETO_CLIPPED|ETO_OPAQUE, 
						     &rcText, 
						     bstrCaption, 
						     nCaptionLen,
						     0);
		else
#endif
			ExtTextOut(hDC, 
					   ptText.x, 
					   ptText.y, 
					   ETO_CLIPPED|ETO_OPAQUE, 
					   &rcText, 
					   szCaption, 
					   nCaptionLen,
					   0);
		break;

	case dd3DRaisedLight:
	case dd3DInsetLight:
		if (eTabSlidingVert == nArea)
		{
			ptText.x += 2;
			ptText.y -= 2;
		}
		else
		{
			ptText.x -= 2;
			ptText.y += 2;
		}
		SetTextColor(hDC, RGB(255, 255, 255));
#ifdef _SCRIPTSTRING
		if (pBar && VARIANT_TRUE == pBar->bpV1.m_vbUseUnicode)
			ScriptExtTextOut(hDC, 
						     ptText.x, 
						     ptText.y, 
						     ETO_CLIPPED|ETO_OPAQUE, 
						     &rcText, 
						     bstrCaption, 
						     nCaptionLen,
						     0);
		else
#endif
			ExtTextOut(hDC, 
					   ptText.x, 
					   ptText.y, 
					   ETO_CLIPPED|ETO_OPAQUE, 
					   &rcText, 
					   szCaption, 
					   nCaptionLen,
					   0);

		if (dd3DRaisedLight == tpV1.m_fsFont3D)
		{
			ptText.x += 1;
			ptText.y += 1;
		}
		else
		{
			ptText.x -= 1;
			ptText.y -= 1;
		}

		SetBkMode(hDC, TRANSPARENT);
		SetTextColor(hDC, m_crForeColor);
		ExtTextOut(hDC, 
				   ptText.x, 
				   ptText.y, 
				   ETO_CLIPPED, 
				   &rcText, 
				   szCaption, 
				   nCaptionLen,
				   0);
		break;

	case dd3DRaisedHeavy:
	case dd3DInsetHeavy:
		if (dd3DRaisedHeavy == tpV1.m_fsFont3D)
			SetTextColor(hDC, RGB(128, 128, 128));
		else
			SetTextColor(hDC, RGB(255, 255, 255));

#ifdef _SCRIPTSTRING
		if (pBar && VARIANT_TRUE == pBar->bpV1.m_vbUseUnicode)
			ScriptExtTextOut(hDC, 
						     ptText.x, 
							 ptText.y, 
							 ETO_CLIPPED|ETO_OPAQUE, 
							 &rcText, 
							 bstrCaption, 
							 nCaptionLen,
							 0);
		else
#endif
			ExtTextOut(hDC, 
					   ptText.x, 
					   ptText.y, 
					   ETO_CLIPPED|ETO_OPAQUE, 
					   &rcText, 
					   szCaption, 
					   nCaptionLen,
					   0);

		SetBkMode(hDC, TRANSPARENT);
		if (eTabSlidingVert == nArea)
			ptText.y -= 2; 
		else
			ptText.x -= 2; 

		if (dd3DRaisedHeavy == tpV1.m_fsFont3D)
			SetTextColor(hDC, RGB(255, 255, 255));
		else
			SetTextColor(hDC, RGB(128, 128, 128));

#ifdef _SCRIPTSTRING
		if (pBar && VARIANT_TRUE == pBar->bpV1.m_vbUseUnicode)
			ScriptExtTextOut(hDC, 
						     ptText.x, 
						     ptText.y, 
						     ETO_CLIPPED, 
						     &rcText, 
						     bstrCaption, 
						     nCaptionLen,
						     0);
		else
#endif
			ExtTextOut(hDC, 
					   ptText.x, 
					   ptText.y, 
					   ETO_CLIPPED, 
					   &rcText, 
					   szCaption, 
					   nCaptionLen,
					   0);

		if (eTabSlidingVert == nArea)
			ptText.y += 1; 
		else
			ptText.x += 1; 

		SetTextColor(hDC, m_crForeColor);
#ifdef _SCRIPTSTRING
		if (pBar && VARIANT_TRUE == pBar->bpV1.m_vbUseUnicode)
			ScriptExtTextOut(hDC, 
						     ptText.x, 
						     ptText.y, 
						     ETO_CLIPPED, 
						     &rcText, 
						     bstrCaption, 
						     nCaptionLen,
						     NULL);
		else
#endif
			ExtTextOut(hDC, 
					   ptText.x, 
					   ptText.y, 
					   ETO_CLIPPED, 
					   &rcText, 
					   szCaption, 
					   nCaptionLen,
					   NULL);
		break;
	}
	return TRUE;
}

//
// PaintTopScrollButton
//

void BTabs::PaintTopScrollButton(HDC hDC)
{
	UINT nStyle = 0;
	if (eInactiveLeftButton == m_sbsTopStyle || 
		eInactiveTopButton == m_sbsTopStyle)
	{
		return;
	}
	if (ePressedLeftButton == m_sbsTopStyle || 
		ePressedTopButton == m_sbsTopStyle ||
		m_bTopInteractivePaint)
	{
		nStyle |= DFCS_PUSHED;
	}
	switch (m_sbsTopStyle)
	{
	case eLeftButton:
	case eInactiveLeftButton:
	case ePressedLeftButton:
		nStyle |= DFCS_SCROLLLEFT;
		break;

	case eTopButton:
	case eInactiveTopButton:
	case ePressedTopButton:
		nStyle |= DFCS_SCROLLUP;
		break;
	}
	DrawFrameControl(hDC, 
					 &m_rcTopButton, 
					 DFC_SCROLL, 
					 nStyle);
}

//
// PaintBottomScrollButton
//

void BTabs::PaintBottomScrollButton(HDC hDC)
{
	UINT nStyle = 0;
	if (eInactiveRightButton == m_sbsBottomStyle || 
		eInactiveBottomButton == m_sbsBottomStyle)
	{
		return;
	}
	if (ePressedRightButton == m_sbsBottomStyle || 
		ePressedBottomButton == m_sbsBottomStyle ||
		m_bBottomInteractivePaint)
	{
		nStyle |= DFCS_PUSHED;
	}
	switch (m_sbsBottomStyle)
	{
	case eRightButton:
	case eInactiveRightButton:
	case ePressedRightButton:
		nStyle |= DFCS_SCROLLRIGHT;
		break;
	
	case eBottomButton:
	case eInactiveBottomButton:
	case ePressedBottomButton:
		nStyle |= DFCS_SCROLLDOWN;
		break;
	}
	DrawFrameControl(hDC, 
					 &m_rcBottomButton, 
					 DFC_SCROLL, 
					 nStyle);
}

//
// InteractivePaintScrollButton
//

void BTabs::InteractivePaintScrollButton(HDC hDC, SCROLLBUTTONSTYLES sbsStyle)
{
	TRACE(5, _T("InteractivePaintScrollButton\n"));
	UINT nStyle = 0;
	CRect rc; 
	switch (sbsStyle)
	{
	case eLeftButton:
	case eInactiveLeftButton:
	case ePressedLeftButton:
		rc = m_rcTopButton;
		nStyle |= DFCS_SCROLLLEFT;
		break;

	case eTopButton:
	case eInactiveTopButton:
	case ePressedTopButton:
		rc = m_rcTopButton;
		nStyle |= DFCS_SCROLLUP;
		break;

	case eRightButton:
	case eInactiveRightButton:
	case ePressedRightButton:
		rc = m_rcBottomButton;
		nStyle |= DFCS_SCROLLRIGHT;
		break;

	case eBottomButton:
	case eInactiveBottomButton:
	case ePressedBottomButton:
		rc = m_rcBottomButton;
		nStyle |= DFCS_SCROLLDOWN;
		break;
	}
	switch (sbsStyle)
	{
	case ePressedRightButton:
	case ePressedBottomButton:
	case ePressedTopButton:
	case ePressedLeftButton:
		nStyle |= DFCS_PUSHED;
	}
	DrawFrameControl(hDC, 
					 &rc, 
					 DFC_SCROLL, 
					 nStyle);
}

