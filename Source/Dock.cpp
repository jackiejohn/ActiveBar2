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
#include <stdio.h>
#include <crtdbg.h>
#include "Debug.h"
#include "Globals.h"
#include "IpServer.h"
#include "FDialog.h"
#include "Support.h"
#include "Resource.h"
#include "DropSource.h"
#include "MiniWin.h"
#include "PopupWin.h"
//#include "Accessible.h"
#include "FRegKey.h"
#include <MultiMon.h>
#include "Designer\DesignerInterfaces.h"
#include "Designer\DragDrop.h"
#include "Tool.h"
#include "ChildBands.h"
#include "Band.h"
#include "Bands.h"
#include "Dock.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

extern BOOL g_fMachineHasLicense;
extern BOOL g_bDemoDlgDisplayed;

//
// CSplitter
//

CSplitter::CSplitter(HWND hWndParent, CDock* pDock)
 : m_hPrimaryCursor(NULL),
   m_hWndParent(hWndParent),
   m_pDock(pDock)
{
	m_bTracking = FALSE;
}

CSplitter::~CSplitter()
{
}

//
// FindNextBand
//

CBand* CSplitter::FindNextBand(CBand* pBandIn, BOOL bNext)
{
	CBand* pBand;
	int nSplitterCount = m_aInsideBands.GetSize();
	for (int nBand = 0; nBand < nSplitterCount; nBand++)
	{
		pBand = m_aInsideBands.GetAt(nBand);
		assert(pBand);
		if (NULL == pBand)
			continue;
		if (pBand == pBandIn)
		{
			if (bNext ? nBand < nSplitterCount - 1 : nBand - 1 >= 0)
				return m_aInsideBands.GetAt(bNext ? nBand + 1 : nBand-1);
			return NULL;
		}
	}
	return NULL;
}

//
// SortSplitterBandsByLastDocked
//

static int SortSplitterBandsByLastDockedFunction(const void *arg1, const void *arg2)
{
	CBand* pBand1 = *((CBand**)arg1);
	CBand* pBand2 = *((CBand**)arg2);

	if (pBand2->m_dwLastDocked != pBand1->m_dwLastDocked)
		return pBand2->m_dwLastDocked - pBand1->m_dwLastDocked;
	return 0;
}

static void SortSplitterBandsByLastDocked(ArrayOfBandPointers* pfaBands)
{
	qsort(pfaBands->GetData(), pfaBands->GetSize(), sizeof(void*), SortSplitterBandsByLastDockedFunction);
}

//
// SetLastDocked 
//

void CSplitter::SetToTheHighestLastDocked()
{
	TypedArray<CBand*> aSortedBands;
	aSortedBands.Copy(m_aInsideBands);
	SortSplitterBandsByLastDocked(&aSortedBands);
	DWORD dwLastDocked = 0;
	CBand* pBand;
	BOOL bFirst = TRUE;
	int nSplitterCount = aSortedBands.GetSize();
	for (int nBand = 0; nBand < nSplitterCount; nBand++)
	{
		pBand = aSortedBands.GetAt(nBand);
		assert(pBand);
		if (NULL == pBand)
			continue;

		if (bFirst)
		{
			dwLastDocked = pBand->m_dwLastDocked;
			bFirst = FALSE;
		}
		else
			pBand->m_dwLastDocked = dwLastDocked;
	}
}

//
// OnInvertTracker
//

void CSplitter::OnInvertTracker()
{
	HDC hDC = GetDC(NULL);
	if (hDC)
	{
		//
		// Invert the brush pattern (Looks just like frame window sizing)
		//

		CRect rcSplitter;
		GetSplitterRect(rcSplitter);

		HBRUSH hBrush = GetGlobals().GetBrushDither();
		assert(hBrush);
		if (NULL == hBrush)
			return;

		HBRUSH hBrushOld = SelectBrush(hDC, hBrush);
		
		PatBlt(hDC, 
			   rcSplitter.left, 
			   rcSplitter.top, 
			   rcSplitter.Width(), 
			   rcSplitter.Height(), 
			   PATINVERT);

		SelectBrush(hDC, hBrushOld);
		ReleaseDC(NULL, hDC);
	}
}

//
// SetSplitCursor
//

void CSplitter::SetSplitCursor(const BOOL& bSetSplit)
{
	if (bSetSplit)
	{
		switch (m_pDock->GetDockingArea())
		{
		case ddDATop:
		case ddDABottom:
			m_hPrimaryCursor = ::SetCursor(GetGlobals().GetSplitHCursor());
			break;

		case ddDALeft:
		case ddDARight:
			m_hPrimaryCursor = ::SetCursor(GetGlobals().GetSplitVCursor());
			break;
		}
	}
	else
		::SetCursor(m_hPrimaryCursor);
}

//
// GetSplitterRect
//

void CSplitter::GetSplitterRect(CRect& rcSplitter)
{
	CRect rcDockWindow;
	m_pDock->GetWindowRect(rcDockWindow);
	rcSplitter = m_rcSplitter;
	rcSplitter.Offset(rcDockWindow.left, rcDockWindow.top);
	switch (m_pDock->GetDockingArea())
	{
	case ddDATop:
	case ddDABottom:
		rcSplitter.left = rcDockWindow.left;
		rcSplitter.right = rcDockWindow.right;
		break;

	case ddDALeft:
	case ddDARight:
		rcSplitter.top = rcDockWindow.top;
		rcSplitter.bottom = rcDockWindow.bottom;
		break;
	}
}

//
// SetSplitterRect
//

void CSplitter::SetSplitterRect(const CRect& rcSplitter)
{
	try
	{
		m_rcSplitter = rcSplitter;
		switch (m_pDock->GetDockingArea())
		{
		case ddDATop:
			m_rcSplitter.top = m_rcSplitter.bottom - eSplitterThickness;
			break;
		case ddDABottom:
			m_rcSplitter.bottom = m_rcSplitter.top + eSplitterThickness;
			break;
		case ddDALeft:
			m_rcSplitter.left = m_rcSplitter.right - eSplitterThickness;
			break;
		case ddDARight:
			m_rcSplitter.right = m_rcSplitter.left + eSplitterThickness;
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}	
}

//
// UpdateSplitterRect
//

void CSplitter::UpdateSplitterRect(POINT& ptScreen)
{
	try
	{
		//
		// Check the boundries of the parent window and the boundries of the bands
		//

		int nTemp;
		CRect rcParent;
		GetWindowRect(m_hWndParent, &rcParent);
		
		if (m_aInsideBands.GetSize() < 1)
			return;

		CBand* pBand = m_aInsideBands.GetAt(0);
		if (NULL == pBand)
			return;

		CRect rcDock;
		CRect rcBand = pBand->m_rcDock;
		m_pDock->ClientToScreen(rcBand);
		switch (m_pDock->GetDockingArea())
		{
		case ddDATop:
			rcDock = m_pDock->DockMgr().GetDockRect(ddDABottom);
			nTemp = rcParent.bottom - rcDock.Height() - eBottomLimit;
			if (ptScreen.y > nTemp)
				ptScreen.y = nTemp;

			nTemp = pBand->bpV1.m_nDockedHorzMinWidth;
			if (ptScreen.y < rcBand.top + nTemp)
				ptScreen.y = rcBand.top + nTemp;
			break;

		case ddDABottom:
			rcDock = m_pDock->DockMgr().GetDockRect(ddDATop);
			nTemp = rcParent.top + rcDock.Height() + eTopLimit;
			if (ptScreen.y < nTemp)
				ptScreen.y = nTemp;

			nTemp = pBand->bpV1.m_nDockedHorzMinWidth;
			if (ptScreen.y > rcBand.bottom - nTemp)
				ptScreen.y = rcBand.bottom - nTemp;
			break;

		case ddDALeft:
			rcDock = m_pDock->DockMgr().GetDockRect(ddDARight);
			nTemp = rcParent.right - rcDock.Width() - eRightLimit;
			if (ptScreen.x > nTemp)
				ptScreen.x = nTemp;

			nTemp = pBand->bpV1.m_nDockedVertMinWidth;
			if (ptScreen.x < rcBand.left + nTemp)
				ptScreen.x = rcBand.left + nTemp;
			break;

		case ddDARight:
			rcDock = m_pDock->DockMgr().GetDockRect(ddDALeft);
			nTemp = rcParent.left + rcDock.Width() + eLeftLimit;
			if (ptScreen.x < nTemp)
				ptScreen.x = nTemp;

			nTemp = pBand->bpV1.m_nDockedVertMinWidth;
			if (ptScreen.x > rcBand.right - nTemp)
				ptScreen.x = rcBand.right - nTemp;
			break;
		}
		
		//
		// Move the Splitter Rect
		//
		
		CRect rcSplitter;
		GetSplitterRect(rcSplitter);
		switch (m_pDock->GetDockingArea())
		{
		case ddDATop:
		case ddDABottom:
			m_rcSplitter.Offset(0, ptScreen.y - rcSplitter.top);
			break;

		case ddDALeft:
		case ddDARight:
			m_rcSplitter.Offset(ptScreen.x - rcSplitter.left, 0);
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}	
}

//
// StartTracking
//

void CSplitter::StartTracking()
{
	try
	{
		m_bTracking = TRUE;
		SetCapture(m_hWndParent);

		OnInvertTracker();

		SetSplitCursor(TRUE);
		
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}	
	try
	{
		MSG msg;
		while (GetCapture() == m_hWndParent && m_bTracking)
		{
			GetMessage(&msg, NULL, 0, 0);

			if (GetCapture() != m_hWndParent)
				goto ExitLoop;

			switch (msg.message)
			{
			case WM_CANCELMODE:
			case WM_LBUTTONUP:
				goto ExitLoop;
				break;

			case WM_MOUSEMOVE:
				DoTracking(msg.pt);
				break;

			default:
				DispatchMessage(&msg);
			}
		}
ExitLoop:
		StopTracking(msg.pt);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}	
}

//
// DoTracking
//

void CSplitter::DoTracking(POINT ptScreen)
{
	try
	{
		OnInvertTracker();
		UpdateSplitterRect(ptScreen);
		OnInvertTracker();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}	
}

//
// StopTracking
//

BOOL CSplitter::StopTracking(POINT ptScreen)
{
	try
	{
		if (!m_bTracking)
			return FALSE;
		m_bTracking = FALSE;

		ReleaseCapture();
		SetSplitCursor(FALSE);

		// Erase tracker rectangle
		OnInvertTracker();

		UpdateSplitterRect(ptScreen);

		CBand* pBand;
		CRect  rcBand;
		CRect  rcSplitter;
		GetSplitterRect(rcSplitter);

		int nTemp = -1;
		int nCount = m_aInsideBands.GetSize();
		for (int nBand = 0; nBand < nCount; nBand++)
		{
			pBand = m_aInsideBands.GetAt(nBand);
			assert(pBand);
			if (NULL == pBand)
				continue;

			if (-1 == nTemp)
			{
				rcBand = pBand->m_rcDock;
				m_pDock->ClientToScreen(rcBand);
				switch (m_pDock->GetDockingArea())
				{
				case ddDATop:
					nTemp = rcSplitter.bottom - rcBand.top;
					break;
				case ddDABottom:
					nTemp = rcBand.bottom - rcSplitter.top;
					break;
				case ddDALeft:
					nTemp = rcSplitter.right - rcBand.left;
					break;
				case ddDARight:
					nTemp = rcBand.right - rcSplitter.left;
					break;
				}
			}
			switch (m_pDock->GetDockingArea())
			{
			case ddDATop:
			case ddDABottom:
				pBand->m_rcDock.bottom = pBand->m_rcDock.top + nTemp;
				pBand->bpV1.m_rcHorzDockForm.bottom = pBand->bpV1.m_rcHorzDockForm.top + nTemp - pBand->BandEdge().cy;
				break;

			case ddDALeft:
			case ddDARight:
				pBand->m_rcDock.right = pBand->m_rcDock.left + nTemp;
				pBand->bpV1.m_rcVertDockForm.right = pBand->bpV1.m_rcVertDockForm.left + nTemp - pBand->BandEdge().cx;
				break;
			}
		}
		if (m_pDock->m_pBar->m_pDesigner)
			m_pDock->m_pBar->m_pDesigner->SetDirty();
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
// CBetweenBandsSplitter
//

CBetweenBandsSplitter::CBetweenBandsSplitter(CSplitter* pLineSplitter, HWND hWndParent, CDock* pDock, ArrayOfBandPointers& aBands)
	: CSplitter(hWndParent, pDock),
	  m_pLineSplitter(pLineSplitter),
	  m_aBands(aBands)
{
}

//
// AdjustSplitter
//

void CBetweenBandsSplitter::AdjustSplitter()
{
	CBand* pBand = m_aInsideBands.GetAt(0);
	if (NULL == pBand)
		return;
	SetSplitterRect(pBand->m_rcDock);
}

//
// GetSplitterRect
//

void CBetweenBandsSplitter::GetSplitterRect(CRect& rcSplitter)
{
	CRect rcDockWindow;
	m_pDock->GetWindowRect(rcDockWindow);
	rcSplitter = m_rcSplitter;
	rcSplitter.Offset(rcDockWindow.left, rcDockWindow.top);
}

//
// SetSplitterRect
//

void CBetweenBandsSplitter::SetSplitterRect(const CRect& rcSplitter)
{
	try
	{
		m_rcSplitter = rcSplitter;
		switch (m_pDock->GetDockingArea())
		{
		case ddDATop:
		case ddDABottom:
			m_rcSplitter.left = m_rcSplitter.right - eSplitterThickness;
			break;

		case ddDALeft:
		case ddDARight:
			m_rcSplitter.top = m_rcSplitter.bottom - eSplitterThickness;
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}	
}

//
// SetSplitCursor
//

void CBetweenBandsSplitter::SetSplitCursor(const BOOL& bSetSplit)
{
	try
	{
		if (bSetSplit)
		{
			switch (m_pDock->GetDockingArea())
			{
			case ddDATop:
			case ddDABottom:
				m_hPrimaryCursor = ::SetCursor(GetGlobals().GetSplitVCursor());
				break;

			case ddDALeft:
			case ddDARight:
				m_hPrimaryCursor = ::SetCursor(GetGlobals().GetSplitHCursor());
				break;
			}
		}
		else
			::SetCursor(m_hPrimaryCursor);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}	
}

//
// UpdateSplitterRect
//

void CBetweenBandsSplitter::UpdateSplitterRect(POINT& ptScreen)
{
	try
	{
		//
		// Check the boundries of the parent window and the boundries of the bands
		//

		BOOL bHitBand = FALSE;
		CBand* pBand;
		CRect rcBand;
		int nLeftAdjustment = 0;
		int nRightAdjustment = 0;
		int nCount  = m_aBands.GetSize();
		for (int nBand = 0; nBand < nCount; nBand++)
		{
			pBand = m_aBands.GetAt(nBand);
			rcBand = pBand->m_rcDock;

			if (m_pBandOutSide == pBand)
				bHitBand = TRUE;

			switch (m_pDock->GetDockingArea())
			{
			case ddDATop:
			case ddDABottom:
				if (bHitBand)
				{
					if (!(ddBFSizer & pBand->bpV1.m_dwFlags))
						nRightAdjustment += rcBand.Width();
					else
						nRightAdjustment += pBand->m_nAdjustForGrabBar + eRightLimit;
				}
				else
				{
					if (!(ddBFSizer & pBand->bpV1.m_dwFlags))
						nLeftAdjustment += rcBand.Width();
					else
						nLeftAdjustment += pBand->m_nAdjustForGrabBar + eLeftLimit;
				}
				break;

			case ddDALeft:
			case ddDARight:
				if (bHitBand)
				{
					if (!(ddBFSizer & pBand->bpV1.m_dwFlags))
						nRightAdjustment += rcBand.Height();
					else
						nRightAdjustment += pBand->m_nAdjustForGrabBar + eBottomLimit;
				}
				else
				{
					if (!(ddBFSizer & pBand->bpV1.m_dwFlags))
						nLeftAdjustment += rcBand.Height();
					else
						nLeftAdjustment += pBand->m_nAdjustForGrabBar + eTopLimit;
				}
				break;
			}
		}

		CRect rcDock;
		m_pDock->GetWindowRect(rcDock);

		switch (m_pDock->GetDockingArea())
		{
		case ddDATop:
		case ddDABottom:
			if (ptScreen.x < rcDock.left + nLeftAdjustment)
				ptScreen.x = rcDock.left + nLeftAdjustment;

			if (ptScreen.x > rcDock.right - nRightAdjustment)
				ptScreen.x = rcDock.right - nRightAdjustment;
			break;

		case ddDALeft:
		case ddDARight:
			if (ptScreen.y < rcDock.top + nLeftAdjustment)
				ptScreen.y = rcDock.top + nLeftAdjustment;

			if (ptScreen.y > rcDock.bottom - nRightAdjustment)
				ptScreen.y = rcDock.bottom - nRightAdjustment;
			break;
		}
		
		//
		// Move the Splitter Rect
		//
		
		CRect rcSplitter;
		GetSplitterRect(rcSplitter);
		switch (m_pDock->GetDockingArea())
		{
		case ddDATop:
		case ddDABottom:
			m_rcSplitter.Offset(ptScreen.x - rcSplitter.left, 0);
			break;

		case ddDALeft:
		case ddDARight:
			m_rcSplitter.Offset(0, ptScreen.y - rcSplitter.top);
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}	
}

//
// StopTracking
//

BOOL CBetweenBandsSplitter::StopTracking(POINT ptScreen)
{
	try
	{
		if (!m_bTracking)
			return FALSE;
		m_bTracking = FALSE;

		ReleaseCapture();
		SetSplitCursor(FALSE);

		// Erase tracker rectangle
		OnInvertTracker();

		UpdateSplitterRect(ptScreen);

		int nCount = m_aInsideBands.GetSize();
		if (nCount < 1)
			return FALSE;

		CBand* pBand = m_aInsideBands.GetAt(0);
		assert(pBand);
		if (NULL == pBand)
			return FALSE;

		int nDiff;
		CRect rcBand = pBand->m_rcDock;
		m_pDock->ClientToScreen(rcBand);
		
		CRect rcSplitter;
		GetSplitterRect(rcSplitter);

		CBand* pBandSplitter;
		CBand* pNextBand;
		if (!(ddBFSizer & pBand->bpV1.m_dwFlags))
		{
			pNextBand = m_pBandOutSide;
			pBandSplitter = m_pLineSplitter->FindNextBand(pNextBand, FALSE);
		}
		else
		{
			pBandSplitter = pBand;
			pNextBand = m_pLineSplitter->FindNextBand(pBand);
		}

		switch (m_pDock->GetDockingArea())
		{
		case ddDATop:
		case ddDABottom:
			nDiff = rcSplitter.right - rcBand.right - 2;
			
			pBandSplitter->bpV1.m_rcHorzDockForm.right += nDiff;
			if (pNextBand)
				pNextBand->bpV1.m_rcHorzDockForm.right -= nDiff;
			break;

		case ddDALeft:
		case ddDARight:
			nDiff = rcSplitter.bottom - rcBand.bottom - 2;

			pBandSplitter->bpV1.m_rcVertDockForm.bottom += nDiff;
			if (pNextBand)
				pNextBand->bpV1.m_rcVertDockForm.bottom -= nDiff;
			break;
		}
		m_pDock->MakeLineContracted(pBand->m_nDockLineIndex);
		if (m_pDock->Bar()->m_pDesigner)
			m_pDock->Bar()->m_pDesigner->SetDirty();
		m_pDock->InvalidateRect(NULL, FALSE);
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
// CSplitterMgr
//

CSplitterMgr::CSplitterMgr()
	: m_pWnd(NULL)
{
}

CSplitterMgr::~CSplitterMgr()
{
	Cleanup();
}

//
// CleanUp
//

void CSplitterMgr::Cleanup()
{
	int nCount = m_aSplitters.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
		delete m_aSplitters.GetAt(nIndex);
	if (nCount > 0)
		m_aSplitters.RemoveAll();
}

//
// HitTest
//

CSplitter* CSplitterMgr::HitTest(POINT pt)
{
	try
	{
		m_pWnd->ClientToScreen(pt);
		CRect rcSplitter;
		CSplitter* pSplitter;
		int nCount = m_aSplitters.GetSize();
		for (int nIndex = 0; nIndex < nCount; nIndex++)
		{
			pSplitter = m_aSplitters.GetAt(nIndex);
			assert(pSplitter);
			if (NULL == pSplitter)
				continue;

			pSplitter->GetSplitterRect(rcSplitter);
			if (PtInRect(&rcSplitter, pt))
				return pSplitter;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}	
	return NULL;
}

//
// RemoveSplitter
//

BOOL CSplitterMgr::RemoveSplitter(CSplitter* pSplitter)
{
	try
	{
		int nCount = m_aSplitters.GetSize();
		for (int nIndex = 0; nIndex < nCount; nIndex++)
		{
			if (pSplitter = m_aSplitters.GetAt(nIndex))
			{
				m_aSplitters.RemoveAt(nIndex);
				return TRUE;
			}
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}	
	return FALSE;
}

//
// Draw
//

BOOL CSplitterMgr::Draw(HDC hDC)
{
	return TRUE;
}


//
// CDockDragDrop
//

CBand* CDockDragDrop::GetBand(POINT pt)
{
	assert(m_pWnd);
	m_pWnd->ScreenToClient(pt);
	CBand* pBand = ((CDock*)m_pWnd)->HitTest(pt);
	if (pBand && ddCBSNone != pBand->bpV1.m_cbsChildStyle)
		pBand = pBand->m_pChildBands->GetCurrentChildBand();
	return pBand;
}

void CDockDragDrop::OffsetPoint(CBand* pBand, POINT& pt)
{
	assert(pBand);
	if (pBand)
	{
		m_pWnd->ScreenToClient(pt);
		pt.x -= pBand->m_pRootBand->m_rcDock.left;
		pt.y -= pBand->m_pRootBand->m_rcDock.top;
	}
}
				
//
// CLine
//

CLine::CLine()
{
	m_bSplitterLine = FALSE;
	m_nThickness = 0;
	m_nOffset = 0;
	m_eSplitterLineExpandedState = CBand::eGrayed;
}

#ifdef _DEBUG
void CLine::Dump(DumpContext& dc)
{
	_stprintf(dc.m_szBuffer,
			  _T("Thickness: %i Offset: %i\n"), 
			  m_nThickness,
			  m_nOffset);
	dc.Write();
	int nBandCount = m_aBands.GetSize();
	for (int nBand = 0; nBand < nBandCount; nBand++)
		m_aBands.GetAt(nBand)->DumpLayoutInfo(dc);
}
#endif

//
// CLines
//

CLines::~CLines()
{
	Cleanup();
}

void CLines::Cleanup()
{
	int nSize = m_aLines.GetSize();
	for (int nLine = 0; nLine < nSize; nLine++)
		delete m_aLines.GetAt(nLine);
	m_aLines.RemoveAll();
}

#ifdef _DEBUG
void CLines::Dump(LPTSTR szLabel, DumpContext& dc)
{
	_stprintf(dc.m_szBuffer, _T("\nLines - %s\n"), szLabel);
	dc.Write();
	int nSize = m_aLines.GetSize();
	for (int nLine = 0; nLine < nSize; nLine++)
	{
		_stprintf(dc.m_szBuffer, _T("Line: %i\n"), nLine);
		dc.Write();
		m_aLines.GetAt(nLine)->Dump(dc);
	}
}
#endif

//
// CDock
//

DWORD CDock::m_dwDockFlags[] = {ddBFDockTop, ddBFDockBottom, ddBFDockLeft, ddBFDockRight};

CDock::CDock(CDockMgr* pDockMgr, const DockingAreaTypes& daDockArea)
	:	m_pDockMgr(pDockMgr),
		m_pBar(pDockMgr->Bar()),
		m_daDockingArea(daDockArea),
		m_pDesignerBand(NULL),
		m_pLines(NULL),
		m_pCacheLines(NULL),
//		m_pSplitter(NULL),
		m_theDockDropTarget(CToolDataObject::eActiveBarDragDropId)//,
//		m_pAccessible(NULL)
{
	BOOL bResult = RegisterWindow(DD_WNDCLASS("ActiveBarDockWnd"),
								  CS_DBLCLKS,
								  NULL,
								  LoadCursor(NULL, IDC_ARROW));
	assert(bResult);
 	m_theSplitters.Wnd(this);
	m_nLineCount = 0;
	m_theBandDragDrop.m_pWnd = this;
	m_theDockDropTarget.Init(&m_theBandDragDrop, m_pBar, CToolDataObject::eActiveBarDragDropId);
	m_nCacheSize = 2;
}

//
// ~CDock()
//

CDock::~CDock()
{
	delete m_pLines;
	if (IsWindow())
		DestroyWindow();
//	if (m_pAccessible)
//		m_pAccessible->Release();
}

//
// Create
//

BOOL CDock::Create(HWND hWndParent, CRect& rcWindow)
{
	try
	{
		static TCHAR* s_szDockTitles[CDockMgr::eNumOfDockAreas] = 
		{	_T("DockTop"), 
			_T("DockBottom"), 
			_T("DockLeft"), 
			_T("DockRight")
		};
		FWnd::Create(s_szDockTitles[m_daDockingArea],
					 WS_VISIBLE | WS_CHILDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
					 rcWindow.left,
					 rcWindow.top,
					 rcWindow.Width(),
					 rcWindow.Height(),
   					 hWndParent);
		
		assert(IsWindow());
		if (IsWindow())
		{
			assert(m_pBar);
			if (m_pBar->AmbientUserMode())
				DragAcceptFiles(m_hWnd, TRUE);

			if (GetGlobals().m_pDragDropMgr->RegisterDragDrop((OLE_HANDLE)m_hWnd, &m_theDockDropTarget))
			{
				assert(FALSE);
				TRACE(1, "Failed to register drag and drop target\n");
			}
			return TRUE;
		}
		return FALSE;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}	
}

// {96B554FD-70DD-419D-B450-F019B06847EF}
DEFINE_GUID(IID_License, 0x96b554fd, 0x70dd, 0x419d, 0xb4, 0x50, 0xf0, 0x19, 0xb0, 0x68, 0x47, 0xef);
// {1ADF77E2-5E77-4E0D-B913-19992D7D36C6}
DEFINE_GUID(IID_LicenseKey, 0x1adf77e2, 0x5e77, 0x4e0d, 0xb9, 0x13, 0x19, 0x99, 0x2d, 0x7d, 0x36, 0xc6);

//
// WindowProc
//

LRESULT CDock::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		if (GetGlobals().WM_SIZEPARENT == nMsg)
		{
			OnSizeParent((SIZEPARENTPARAMS*)wParam);
			return TRUE;
		}

		switch (nMsg)
		{
		case WM_NCHITTEST:
			{
				POINT ptBand = {LOWORD(lParam), HIWORD(lParam)};
				if (ScreenToClient(ptBand))
				{
					CBand* pBand = HitTest(ptBand);
					if (pBand && ddBTStatusBar == pBand->bpV1.m_btBands)
					{
						CRect rc = pBand->m_rcSizer;
						POINT pt[] = {{rc.left, rc.bottom}, {rc.right, rc.top}, {rc.right, rc.bottom}};
						HRGN hTriangle = CreatePolygonRgn(pt, 3, WINDING);
						if (hTriangle)
						{
							BOOL bInRegion = PtInRegion(hTriangle, ptBand.x, ptBand.y);
							DeleteRgn(hTriangle);
							if (bInRegion)
								return HTBOTTOMRIGHT;
						}
					}
				}
			}
			break;

		case WM_DROPFILES:
			try
			{
				// handle of internal drop structure 
				HDROP hDrop = (HDROP)wParam;
				POINT pt;
				BOOL bResult = DragQueryPoint(hDrop, &pt); 
				CBand* pBand = HitTest(pt);

				UINT nCount = DragQueryFile(hDrop, 0xFFFFFFFF, NULL, 0); 
				if (nCount < 1)
					return FALSE;

				int nResult;
				TCHAR szBuffer[_MAX_PATH];
				for (UINT nIndex = 0; nIndex < nCount; nIndex++)
				{
					nResult = DragQueryFile(hDrop, nIndex, szBuffer, _MAX_PATH); 
					if (nResult > _MAX_PATH)
						continue;

					MAKE_WIDEPTR_FROMTCHAR(wBuffer, szBuffer);
					BSTR bstrBuffer = SysAllocString(wBuffer);
					if (bstrBuffer)
					{
						m_pBar->FireFileDrop((Band*)pBand, bstrBuffer);
						SysFreeString(bstrBuffer);
					}
				}
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
			}	
			break;

		case WM_DESTROY:
			try
			{
				assert(m_pBar);
				if (m_pBar->AmbientUserMode() && IsWindow())
					DragAcceptFiles(m_hWnd, FALSE);

				assert(GetGlobals().m_pDragDropMgr);
				GetGlobals().m_pDragDropMgr->RevokeDragDrop((OLE_HANDLE)m_hWnd);
				int nCount = m_aBands.GetSize();
				for (int nBand = 0; nBand < nCount; nBand++)
					m_aBands.GetAt(nBand)->HideWindowedTools();
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
			}	
			break;

		case WM_MOUSEACTIVATE:
			{
				// Added for CR 1671
				// Added the check for the splitter for CR 2123
				// Added a check for CR 2390
				// Added a check for CR 2362
				POINT pt;
				GetCursorPos(&pt);
				ScreenToClient(pt);
				if (m_pBar->AmbientUserMode()&& NULL == m_theSplitters.HitTest(pt))
				{
					switch (HIWORD(lParam))
					{
					case WM_LBUTTONDOWN:
					case WM_RBUTTONDOWN:
					case WM_MBUTTONDOWN:
						{
							GetCursorPos(&pt);
							HWND hWndPoint = ::WindowFromPoint(pt);
							if (hWndPoint == GetFocus())
								return MA_ACTIVATE;

							HWND hWnd = m_pBar->m_theFormsAndControls.FindChildFromPoint(pt);
							if (::IsWindow(hWnd))
								return MA_ACTIVATE;
						}
						break;
					}
				}
			}
			break;

			//
			// This code is this way to help debug CRs
			//

	    case WM_NOTIFY:
			{
				HWND hWnd = (HWND)lParam;
				CTool* pTool = NULL;
				if (m_pBar->m_theFormsAndControls.FindControl(hWnd, pTool))
				{
					if (::IsWindow(hWnd))
						return ::SendMessage(hWnd, nMsg, wParam, lParam);
				}
			}
			break;

			//
			// This code is this way to help debug CRs
			//

		case WM_PARENTNOTIFY:
			{
				switch (LOWORD(wParam))
				{
				case WM_CREATE:
				case WM_DESTROY:
					{
						HWND hWnd = (HWND)lParam;
						CTool* pTool = NULL;
						if (m_pBar->m_theFormsAndControls.FindControl(hWnd, pTool))
						{
							if (::IsWindow(hWnd))
								return ::SendMessage(hWnd, nMsg, wParam, lParam);
						}
					}
					break;

				case WM_LBUTTONDOWN:
				case WM_MBUTTONDOWN:
				case WM_RBUTTONDOWN:
					{
						POINT pt;
						GetCursorPos(&pt);
						HWND hWnd = m_pBar->m_theFormsAndControls.FindChildFromPoint(pt);
						if (::IsWindow(hWnd))
							return ::SendMessage(hWnd, nMsg, wParam, lParam);
					}
					break;
				}
			}
			break;

			//
			// This code is this way to help debug CRs
			//

		case WM_NOTIFYFORMAT:
			{
				HWND hWnd = (HWND)lParam;
				CTool* pTool = NULL;
				if (m_pBar->m_theFormsAndControls.FindControl(hWnd, pTool))
				{
					if (::IsWindow(hWnd))
						return ::SendMessage(hWnd, nMsg, wParam, lParam);
				}
			}
			break;

			//
			// This code is this way to help debug CRs
			//

        case WM_CTLCOLORBTN:
        case WM_CTLCOLORDLG:
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORLISTBOX:
        case WM_CTLCOLORMSGBOX:
        case WM_CTLCOLORSCROLLBAR:
        case WM_CTLCOLORSTATIC:
			{
				HWND hWnd = (HWND)lParam;
				CTool* pTool = NULL;
				if (m_pBar->m_theFormsAndControls.FindControl(hWnd, pTool))
				{
					if (::IsWindow(hWnd))
						return ::SendMessage(hWnd, nMsg, wParam, lParam);
				}
			}
			break;

			//
			// This code is this way to help debug CRs
			//

        case WM_DRAWITEM:
			{
				HWND hWnd = (HWND)lParam;
				CTool* pTool = NULL;
				if (m_pBar->m_theFormsAndControls.FindControl(hWnd, pTool))
				{
					if (::IsWindow(hWnd))
						return ::SendMessage(hWnd, nMsg, wParam, lParam);
				}
			}
			break;
        
			//
			// This code is this way to help debug CRs
			//

		case WM_MEASUREITEM:
        case WM_DELETEITEM:
        case WM_VKEYTOITEM:
        case WM_CHARTOITEM:
        case WM_COMPAREITEM:
			{
				HWND hWnd = (HWND)lParam;
				CTool* pTool = NULL;
				if (m_pBar->m_theFormsAndControls.FindControl(hWnd, pTool))
				{
					if (::IsWindow(hWnd))
						return ::SendMessage(hWnd, nMsg, wParam, lParam);
				}
			}
			break;

			//
			// This code is this way to help debug CRs
			//

        case WM_HSCROLL:
        case WM_VSCROLL:
			{
				HWND hWnd = (HWND)lParam;
				CTool* pTool = NULL;
				if (m_pBar->m_theFormsAndControls.FindControl(hWnd, pTool))
				{
					if (::IsWindow(hWnd))
						return ::SendMessage(hWnd, nMsg, wParam, lParam);
				}
			}
			break;

			//
			// This code is this way to help debug CRs
			//

        case WM_SIZE:
			{
				HWND hWnd = (HWND)lParam;
				CTool* pTool = NULL;
				if (m_pBar->m_theFormsAndControls.FindControl(hWnd, pTool))
				{
					if (::IsWindow(hWnd))
						return ::SendMessage(hWnd, nMsg, wParam, lParam);
				}
			}
			break;

        case WM_SETFOCUS:
			if (NULL != m_pBar->m_hWndMDIClient)
			{
				BOOL bMaximized;
				HWND hWndActiveMDIChild = (HWND)::SendMessage(m_pBar->m_hWndMDIClient, WM_MDIGETACTIVE, 0, (LPARAM)&bMaximized);
				if (::IsWindow(hWndActiveMDIChild))
					::SetFocus(hWndActiveMDIChild);
			}
			TRACE1(5,"Focus is on %X\n",GetFocus());
			return 0;

		case WM_COMMAND:
			{
				HWND hWnd = (HWND)lParam;
				CTool* pTool = NULL;
				if (m_pBar->m_theFormsAndControls.FindControl(hWnd, pTool))
				{
					if (::IsWindow(hWnd))
						return ::SendMessage(hWnd, nMsg, wParam, lParam);
				}
				else
				{
					switch (LOWORD(wParam))
					{
					case CTool::eEdit:
					case CTool::eCombobox:
						//
						// Message reflection for the built in Combobox and Edit Tool.
						//
						switch (HIWORD(wParam))
						{
						case EN_SETFOCUS:
						case EN_KILLFOCUS:
						case EN_CHANGE:
						case CBN_CLOSEUP:
						case CBN_KILLFOCUS:
						case CBN_DROPDOWN:
						case CBN_SELCHANGE:
						case CBN_EDITCHANGE:
							::SendMessage((HWND)lParam, WM_COMMAND, wParam, lParam);
							break;
						}
						break;

					default:
						OnCommand(HIWORD(wParam), LOWORD(wParam), (HWND)lParam);
						break;
					}
				}
			}
			break;

		case WM_LBUTTONDOWN:
			{
				m_pBar->ActiveBand(NULL);
				POINT pt = {LOWORD(lParam), HIWORD(lParam)};
				CSplitter* pSplitter = m_theSplitters.HitTest(pt);
				if (pSplitter)
				{
					pSplitter->StartTracking();
					::PostMessage(Bar()->m_hWnd, GetGlobals().WM_RECALCLAYOUT, 0, 0);
					return TRUE;
				}

				CBand* pBand = HitTest(pt);
				if (NULL == pBand)
				{
					if (m_pBar->m_bCustomization || m_pBar->m_bWhatsThisHelp)
						m_pBar->ClearCustomization();
					return FALSE;
				}

				//
				// Make the point relative to the Band 
				//

				pt.x -= pBand->m_rcDock.left;
				pt.y -= pBand->m_rcDock.top;
				m_pBar->AddRef();
				m_pBar->ActiveBand(pBand);
				try
				{
					pBand->OnLButtonDown(wParam, pt);
				}
				catch (...)
				{
					assert(FALSE);
				}
				m_pBar->Refresh();
				m_pBar->Release();
			}
			return TRUE;
		
		case WM_LBUTTONDBLCLK:
			{
				if (m_pBar->m_bMenuLoop)
					return TRUE;

				POINT pt = {LOWORD(lParam), HIWORD(lParam)};
				CBand* pBand = HitTest(pt);
				if (NULL == pBand)
				{
					if (!m_pBar->m_bCustomization)
						m_pBar->StartCustomization();
					return TRUE;
				}

				//
				// Make the point relative to the Band 
				//

				pt.x -= pBand->m_rcDock.left;
				pt.y -= pBand->m_rcDock.top;
				pBand->OnLButtonDblClk(wParam, pt);
			}
			return TRUE;

		case WM_LBUTTONUP:
			{
				POINT pt = {LOWORD(lParam), HIWORD(lParam)};
				CBand* pBand = HitTest(pt);
				if (NULL == pBand)
					return FALSE;

				pt.x -= pBand->m_rcDock.left;
				pt.y -= pBand->m_rcDock.top;
				BOOL bToolChanged = pBand->OnLButtonUp(wParam, pt);
			}
			return TRUE;

		case WM_RBUTTONDOWN:
			{
				POINT pt = {LOWORD(lParam), HIWORD(lParam)};
				CBand* pBand = HitTest(pt);
				if (NULL == pBand)
					return FALSE;

				pt.x -= pBand->m_rcDock.left;
				pt.y -= pBand->m_rcDock.top;
				pBand->OnRButtonDown(wParam, pt);
			}
			return TRUE;

		case WM_RBUTTONUP:
			{
				POINT pt = {LOWORD(lParam), HIWORD(lParam)};
				if (!m_pBar->m_bCustomization)
				{
					BOOL bResult = m_pBar->OnBarContextMenu();
					return TRUE;
				}
				else if (m_pBar->m_pDesigner)
				{
					m_pDesignerBand = HitTest(pt);
					if (NULL == m_pDesignerBand)
						return FALSE;

					//
					// Could bring up a menu to create a new tool
					//
					HMENU hMenu = LoadMenu(g_hInstance, MAKEINTRESOURCE(IDR_MENUBAND)); 
					if (NULL == hMenu)
						return FALSE;

					HMENU hMenuSub = GetSubMenu(hMenu, 0);
					if (NULL == hMenuSub)
					{
						DestroyMenu(hMenu);
						return FALSE;
					}
								
					ClientToScreen(pt);
					BOOL bResult = TrackPopupMenu(hMenuSub, 
												  TPM_LEFTALIGN, 
												  pt.x, 
												  pt.y, 
												  0, 
												  m_hWnd, 
												  0);
					DestroyMenu(hMenu);
				}
			}
			return FALSE;

		case WM_MOUSEMOVE:
			{

/*	

  Took out because of bug report 1691
			
				HWND hWndActive = GetActiveWindow();
				HWND hWndTemp = m_pBar->GetDockWindow();
				if (hWndTemp != hWndActive)
				{
					while (hWndTemp)
					{
						hWndTemp = GetParent(hWndTemp);
						if (hWndTemp == hWndActive)
							break;
					}
					if (NULL == hWndTemp)
						break;
				}
*/
				POINT pt = {LOWORD(lParam), HIWORD(lParam)};
				CSplitter* pSplitter = m_theSplitters.HitTest(pt);
				if (pSplitter)
				{
					pSplitter->SetSplitCursor(TRUE);
					return TRUE;
				}

				CBand* pBand = HitTest(pt);
				if (NULL == pBand)
					return FALSE;

				//
				// Make the point relative to the Band 
				//

				pt.x -= pBand->m_rcDock.left;
				pt.y -= pBand->m_rcDock.top;
				if (pBand->OnMouseMove(wParam, pt))
					return TRUE;

				if (m_pBar->m_pActiveTool)
				{
					// Undo the ActiveTool
					KillTimer(m_pBar->m_hWnd, CBar::eFlyByTimer);
					CTool* pTempTool = m_pBar->m_pActiveTool;
					m_pBar->SetActiveTool(NULL);
					if (pTempTool)
						m_pBar->ToolNotification(pTempTool, TNF_VIEWCHANGED);	
				}
			}
			return FALSE;

		case WM_ERASEBKGND:
			return 0;

		case WM_PAINT:
			try
			{
				if (!g_fMachineHasLicense) 
				{
					FRegKey rkLicenses;
					if (rkLicenses.Open(HKEY_CLASSES_ROOT,_T("Licenses")))
					{
						long nBuffer = 255;
						TCHAR szBuffer[255];
						if (ERROR_SUCCESS == RegQueryValue(rkLicenses, "96B554FD-70DD-419D-B450-F019B06847EF", szBuffer, &nBuffer))
						{
							WCHAR wValue[40];
							if (StringFromGUID2(IID_LicenseKey, (WCHAR*)wValue, 40) > 0)
							{
								MAKE_WIDEPTR_FROMTCHAR(wValue2, szBuffer);
								if (0 == wcscmp(wValue2, wValue))
								{
									g_fMachineHasLicense = TRUE;
									g_bDemoDlgDisplayed = TRUE;
								}
							}
						}
					}
					if (!g_bDemoDlgDisplayed)
						DemoDlg();
				}

#ifdef PERFORMACE_TEST
				++g_dockPaintCount;
				DWORD start=GetTickCount();
#endif
				BOOL bResult;
				PAINTSTRUCT thePaint;
				HDC hDC = BeginPaint(m_hWnd, &thePaint);
				if (hDC)
				{
					POINT ptPaintOffset = {0, 0};
					CRect rcBoundOffset;
					CRect rcClipBox;
					CRect rcBound;
					HDC hDCPaint;
					
					GetClientRect(rcBound);

					int nResult = GetClipBox(hDC, &rcClipBox); 
					if (ERROR == nResult)
						rcClipBox = rcBound;

					BOOL bOptimizePaint = FALSE;
					if (rcClipBox.left > rcBound.left || rcClipBox.right < rcBound.right ||
						rcClipBox.top > rcBound.top || rcClipBox.bottom < rcBound.bottom)
					{
						bOptimizePaint = TRUE;
					}
					
					BOOL bTexture = m_pBar->HasTexture() && ddBODockAreas & m_pBar->bpV1.m_dwBackgroundOptions;
					if (bOptimizePaint && !bTexture)
					{
						hDCPaint = m_pBar->m_ffObj.RequestDC(hDC, rcClipBox.Width(), rcClipBox.Height());
						if (NULL == hDCPaint) 
							hDCPaint = hDC;
						else
						{
							ptPaintOffset.x = rcClipBox.left - rcBound.left;
							ptPaintOffset.y = rcClipBox.top - rcBound.top;
						}
					}
					else
					{
						hDCPaint = m_pBar->m_ffObj.RequestDC(hDC, rcBound.Width(), rcBound.Height());
						if (NULL == hDCPaint) 
							hDCPaint = hDC;
					}

					rcBoundOffset = rcBound;
					rcBoundOffset.Offset(-ptPaintOffset.x,-ptPaintOffset.y);

					
					HPALETTE hPalOld = NULL;
					HPALETTE hPal = m_pBar->Palette();
					if (hPal)
					{
						hPalOld = SelectPalette(hDCPaint, hPal, FALSE);
						RealizePalette(hDCPaint);
					}

					if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook)
					{
						FillSolidRect(hDCPaint, rcBoundOffset, m_pBar->m_crXPBackground);
					}
					else
					{
						if (bTexture)
							m_pBar->FillTexture(hDCPaint, rcBoundOffset);
						else
							FillSolidRect(hDCPaint, rcBoundOffset, m_pBar->m_crBackground);
					}

					CBand* pBand;
					CRect  rcBand;
					int    nBandsInLine;
					int    nBand;
					try
					{
						for (int nLine = 0; nLine < m_nLineCount; nLine++)
						{
							nBandsInLine = m_pLines->GetLine(nLine)->BandCount();
							for (nBand = 0; nBand < nBandsInLine; nBand++)
							{
								try
								{
									pBand = m_pLines->GetLine(nLine)->GetBand(nBand);
									assert(pBand);
									if (NULL == pBand)
										continue;

									rcBand = pBand->m_rcDock;
								
									try
									{
										if (RectVisible(hDC, &rcBand))
										{
											rcBand.Offset(-ptPaintOffset.x, -ptPaintOffset.y);

											if (VARIANT_TRUE == m_pBar->bpV1.m_vbXPLook && !(ddBTChildMenuBar == pBand->bpV1.m_btBands || ddBTMenuBar == pBand->bpV1.m_btBands || ddBTStatusBar == pBand->bpV1.m_btBands))
											{
												HPEN penBack = CreatePen(PS_NULL, 1, m_pBar->m_crXPBandBackground);
												if (penBack)
												{
													HPEN penOld = SelectPen(hDCPaint, penBack);
													HBRUSH hBrush = CreateSolidBrush(m_pBar->m_crXPBandBackground);
													if (hBrush)
													{
														HBRUSH brushOld = SelectBrush(hDCPaint, hBrush);
														
														RoundRect(hDCPaint, rcBand.left, rcBand.top, rcBand.right, rcBand.bottom, 3, 3);
													
														SelectBrush(hDCPaint, brushOld);
														bResult = DeleteBrush(hBrush);
														assert(bResult);
													}
													SelectPen(hDCPaint, penOld);
													bResult = DeletePen(penBack);
													assert(bResult);
												}
											}

											pBand->Draw(hDCPaint, rcBand, pBand->IsWrappable(), ptPaintOffset);
											if (pBand->m_bDrawAnimated)
												pBand->DrawAnimatedSlidingTabs(m_pBar->m_ffObj, hDC, rcBand);
										}
									}
									CATCH
									{
										assert(FALSE);
										REPORTEXCEPTION(__FILE__, __LINE__)
									}	
								}
								CATCH
								{
									assert(FALSE);
									REPORTEXCEPTION(__FILE__, __LINE__)
								}	
							}
						}
					}
					CATCH
					{
						assert(FALSE);
						REPORTEXCEPTION(__FILE__, __LINE__)
					}	
					if (hDCPaint != hDC)
						m_pBar->m_ffObj.Paint(hDC, ptPaintOffset.x, ptPaintOffset.y);

					if (hPal)
						SelectPalette(hDCPaint, hPalOld, FALSE);

					EndPaint(m_hWnd, &thePaint);
#ifdef PERFORMACE_TEST
					DWORD end=GetTickCount();

					totaldockPaintTime += end-start;

					char str[200];
					wsprintf(str,"Dock paint time=%d , totalpaintcount=%d , totaltime=%d\n",end-start,g_dockPaintCount,totaldockPaintTime);
					OutputDebugString(str);
#endif
					return 0;
				}
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
			}	
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}	
	return FWnd::WindowProc(nMsg, wParam, lParam);
}

//
// OnCommand
//

void CDock::OnCommand(int nNotifyCode, int nId, HWND hWndControl)
{
	try
	{
		switch (nId)
		{
		case ID_INSERTTOOL_BUTTON:
			if (m_pBar->m_pDesigner)
				m_pBar->m_pDesigner->CreateTool((IBand*)m_pDesignerBand, ddTTButton);
			break;

		case ID_INSERTTOOL_DROPDOWNBUTTON:
			if (m_pBar->m_pDesigner)
				m_pBar->m_pDesigner->CreateTool((IBand*)m_pDesignerBand, ddTTButtonDropDown);
			break;

		case ID_INSERTTOOL_TEXTBOX:
			if (m_pBar->m_pDesigner)
				m_pBar->m_pDesigner->CreateTool((IBand*)m_pDesignerBand, ddTTEdit);
			break;

		case ID_INSERTTOOL_COMBOBOX:
			if (m_pBar->m_pDesigner)
				m_pBar->m_pDesigner->CreateTool((IBand*)m_pDesignerBand, ddTTCombobox);
			break;

		case ID_INSERTTOOL_LABEL:
			if (m_pBar->m_pDesigner)
				m_pBar->m_pDesigner->CreateTool((IBand*)m_pDesignerBand, ddTTLabel);
			break;

		case ID_INSERTTOOL_SEPARATOR:
			if (m_pBar->m_pDesigner)
				m_pBar->m_pDesigner->CreateTool((IBand*)m_pDesignerBand, ddTTSeparator);
			break;

		case ID_INSERTTOOL_WINDOWLIST:
			if (m_pBar->m_pDesigner)
				m_pBar->m_pDesigner->CreateTool((IBand*)m_pDesignerBand, ddTTWindowList);
			break;

		case ID_INSERTTOOL_ACTIVEXCONTROL:
			if (m_pBar->m_pDesigner)
				m_pBar->m_pDesigner->CreateTool((IBand*)m_pDesignerBand, ddTTControl);
			break;

		case ID_INSERTTOOL_ACTIVEXFORM:
			if (m_pBar->m_pDesigner)
				m_pBar->m_pDesigner->CreateTool((IBand*)m_pDesignerBand, ddTTForm);
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}	
}

//
// ResizeDockableForms
//

void CDock::ResizeDockableForms()
{
	if (NULL == m_pLines)
		return;

	int nBand;
	int nBandCount;
	CLine* pLine;
	CBand* pBand;
	int nNumOfLines = m_pLines->LineCount();
	for (int nLine = 0; nLine < nNumOfLines; nLine++)
	{
		pLine = m_pLines->m_aLines.GetAt(nLine);
		nBandCount = pLine->m_aBands.GetSize();
		for (nBand = 0; nBand < nBandCount; nBand++)
		{
			pBand = pLine->m_aBands.GetAt(nBand);
			if (ddBFSizer & pBand->bpV1.m_dwFlags)
				pBand->m_bSizerResized = TRUE;
		}
	}
}

//
// AdjustSizableBands
//

BOOL CDock::AdjustSizableBands(CBand* pBand, int nDiff, BOOL bHorz)
{
	CBand* pNextBand;
	CBand* pBandSplitter;
	
	if (NULL == m_pLines)
		return false;

	int nCount = m_pLines->GetLine(pBand->m_nDockLineIndex)->SizableBandCount();
	if (nCount <= 1)
		return FALSE;

	int nBand = m_pLines->GetLine(pBand->m_nDockLineIndex)->Find(pBand);
	if (!(ddBFSizer & pBand->bpV1.m_dwFlags))
	{
		pBandSplitter = m_pLines->GetLine(pBand->m_nDockLineIndex)->FindNextSplitter(nBand, FALSE);
		pNextBand = m_pLines->GetLine(pBand->m_nDockLineIndex)->FindNextSplitter(nBand, TRUE);
	}
	else
	{
		pBandSplitter = pBand;
		pNextBand = m_pLines->GetLine(pBand->m_nDockLineIndex)->FindNextSplitter(nBand, TRUE);
		if (NULL == pNextBand)
			pNextBand = m_pLines->GetLine(pBand->m_nDockLineIndex)->FindNextSplitter(nBand, FALSE);
	}
	if (bHorz)
	{
		pBandSplitter->bpV1.m_rcHorzDockForm.right += nDiff;
		if (pNextBand)
			pNextBand->bpV1.m_rcHorzDockForm.right -= nDiff;
	}
	else
	{
		pBandSplitter->bpV1.m_rcVertDockForm.bottom += nDiff;
		if (pNextBand)
			pNextBand->bpV1.m_rcVertDockForm.bottom -= nDiff;
	}
	MakeLineContracted(pBand->m_nDockLineIndex);
	InvalidateRect(NULL, TRUE);
	return TRUE;
}

//
// CacheBandsOffsets
//

BOOL CDock::CacheBandsOffsets()
{
	CBand* pBand;
	int nBand;
	int nBandsInLine;
	for (int nLine = 0; nLine < m_nLineCount; nLine++)
	{
		nBandsInLine = m_pLines->GetLine(nLine)->BandCount();
		for (nBand = 0; nBand < nBandsInLine; nBand++)
		{
			pBand = m_pLines->GetLine(nLine)->GetBand(nBand);
			assert(pBand);
			if (NULL == pBand)
				continue;

			pBand->m_nCachedDockOffset = pBand->bpV1.m_nDockOffset;
		}
	}
	return TRUE;
}

//
// RestoreBandsOffsets
//

BOOL CDock::RestoreBandsOffsets(CBand* pBandSkip)
{
	CBand* pBand;
	int nBand;
	int nBandsInLine;
	for (int nLine = 0; nLine < m_nLineCount; nLine++)
	{
		nBandsInLine = m_pLines->GetLine(nLine)->BandCount();
		for (nBand = 0; nBand < nBandsInLine; nBand++)
		{
			pBand = m_pLines->GetLine(nLine)->GetBand(nBand);
			assert(pBand);
			if (NULL == pBand || pBandSkip == pBand)
				continue;

			pBand->bpV1.m_nDockOffset = pBand->m_nCachedDockOffset;
		}
	}
	return TRUE;
}

//
// OnSizeParent
//

void CDock::OnSizeParent(SIZEPARENTPARAMS* pSPPParent)
{
	// maximum size available
	CRect rc(pSPPParent->rc);
	SIZE sizeAvail = rc.Size();
	
	BOOL bVert = ddDALeft == m_daDockingArea || ddDARight == m_daDockingArea;
	SIZE size = GetSize();
	if (bVert)
		size.cy = 32767;
	else
	{
		size.cx = 32767;
		// Changed 3/30/2000 was sizeAvail.cy = rc.Width(); which can't be right
		sizeAvail.cy = rc.Height();

		if (m_pBar->AmbientUserMode())
		{
			switch (m_pBar->m_vAlign.iVal)
			{
			case 1: // Top
			case 2: // Bottom
				{
					CRect rcParent;
					::GetClientRect(GetParent(m_hWnd), &rcParent);
					rc.left = rcParent.left;
					rc.right = rcParent.right;
				}
				break;
			}
		}
	}

	size.cx = min(size.cx, sizeAvail.cx);
	size.cy = min(size.cy, sizeAvail.cy);

	if (bVert)
	{
		pSPPParent->sizeTotal.cx += size.cx;
		pSPPParent->sizeTotal.cy = max(pSPPParent->sizeTotal.cy, size.cy);
		if (ddDALeft == m_daDockingArea)
			pSPPParent->rc.left += size.cx;
		else
		{
			rc.left = rc.right - size.cx;
			pSPPParent->rc.right -= size.cx;
		}
	}
	else
	{
		pSPPParent->sizeTotal.cy += size.cy ;
		pSPPParent->sizeTotal.cx = max(pSPPParent->sizeTotal.cx, size.cx);
		if (ddDATop == m_daDockingArea)
			pSPPParent->rc.top += size.cy;
		else
		{
			rc.top = rc.bottom - size.cy;
			pSPPParent->rc.bottom -= size.cy;
		}
	}

	rc.right = rc.left + size.cx;
	rc.bottom = rc.top + size.cy;

	m_rcClientPosition = rc;

	RepositionWindow(pSPPParent, m_hWnd, rc);
}

//
// BandSortFunc
//

static int BandSortFunc(const void *arg1, const void *arg2)
{
	CBand* pBand1 = *((CBand**)arg1);
	CBand* pBand2 = *((CBand**)arg2);

	if (ddBTStatusBar == pBand1->bpV1.m_btBands)
		return 1;

	if (pBand1->bpV1.m_nDockLine != pBand2->bpV1.m_nDockLine)
		return pBand1->bpV1.m_nDockLine - pBand2->bpV1.m_nDockLine;

	if (pBand1->bpV1.m_nDockOffset != pBand2->bpV1.m_nDockOffset)
		return pBand1->bpV1.m_nDockOffset - pBand2->bpV1.m_nDockOffset;

	return pBand1->m_nCollectionIndex - pBand2->m_nCollectionIndex;
}

//
// SortBandsByPosition
//

static void SortBandsByPosition(ArrayOfBandPointers* pfaBands)
{
	qsort(pfaBands->GetData(), pfaBands->GetSize(), sizeof(void*), BandSortFunc);
}

//
// MakeLineContracted
//

BOOL CDock::MakeLineContracted(int nLine)
{
	if (nLine < 0 && nLine >= m_nLineCount)
		return FALSE;
	int nCount = m_pLines->m_aLines.GetAt(nLine)->m_aBands.GetSize();
	for (int nBand = 0; nBand < nCount; nBand++)
		m_pLines->GetLine(nLine)->GetBand(nBand)->m_eExpandButtonState = CBand::eContracted;
	return TRUE;
}

//
// RecalcDock 
//
// nMaxSize is Vertical or Horizontal depending on DockingArea
//

SIZE CDock::RecalcDock(int nMaxSize, BOOL bCommit) 
{
	ArrayOfBandPointers aVisibleStretchBands;
	ArrayOfBandPointers aVisibleNonStretchBands;
	CBetweenBandsSplitter* pBetweenBandsSplitter = NULL;
	CSplitter* pLineSplitter = NULL;

	HRESULT hResult;
	CLines* pNewLines = NULL;
	CLine* pLine = NULL;
	CLine* pNewLine = NULL;
	CBand* pStatusBar = NULL;
	CBand* pBandPrev = NULL;
	CBand* pStretchBand = NULL;
	CBand* pBand = NULL;
	CBand* pBandLastDocked = NULL;
	short  nPrevLine;
	DWORD  dwLastDocked;
	CRect  rcThickestBandInLine;
	CRect  rcReturn;
	CRect  rcBound;
	CRect  rcNew;
	CRect  rcTemp;
	BOOL   bHorz = ddDATop == m_daDockingArea || ddDABottom == m_daDockingArea;
	BOOL   bSameLastDocked;
	BOOL   bResult;
	SIZE   size;
	int*   pnSlack = NULL;
	int	   nNewLineCount = 0;
	int	   nBandsInLine;
	int	   nSlackTotal;
	int	   nLineHeight;
	int	   nBandIndex;
	int	   nLineIndex;
	int    nBandCount;
	int	   nCurWidth;
	int    nRemainder;
	int	   nContractedBandsWidth;
	int	   nSplitterBandWidth;
	int	   nPrevSplitterBandWidth;
	int	   nNonSplitterBandWidth;
	int    nNumOfSplitterBands;
	int	   nLine;
	int    nBand;
	int	   nX;
	int	   nY;
   
//	try
//	{
//		SortBandsByPosition(&m_aBands);
//		
//		//
//		// Safeguard against -1 DockLine value
//		//
//		
//		nPrevLine = -100;
//
//		//
//		//  Calculate the number of lines needed 
//		//	One Band might span several screen lines since Bands might wrap
//		//
//		
//		try
//		{
//			nBandCount = m_aBands.GetSize();
//			for (nBand = 0; nBand < nBandCount; nBand++)
//			{
//				pBand = m_aBands.GetAt(nBand);
//				if (VARIANT_TRUE == pBand->bpV1.m_vbVisible && ddBTPopup != pBand->bpV1.m_btBands)
//				{
//					if (!pBand->NeedsFullStretch()) //ddBTStatusBar != pBand->bpV1.m_btBands) 
//					{
//						hResult = aVisibleNonStretchBands.Add(pBand);
//						if (FAILED(hResult))
//							continue;
//
//						pBand->m_nTempDockOffset = -1;
//						
//						//
//						// Bands on the same DockLine go on the same line
//						// 
//						//
//						
//						if (pBand->bpV1.m_nDockLine != nPrevLine)
//						{
//							nPrevLine = pBand->bpV1.m_nDockLine;
//							nNewLineCount++;
//						}
//					}
//					else
//					{
//						hResult = aVisibleStretchBands.Add(pBand);
//						if (FAILED(hResult))
//							continue;
//						nNewLineCount++;
//					}
//				}
//			}
//		}
//		CATCH
//		{
//			assert(FALSE);
//			REPORTEXCEPTION(__FILE__, __LINE__)
//			goto Error;
//		}	
//			
//		try
//		{
//			//
//			// Allocate an array for each line
//			//
//
//			if (nNewLineCount > 0)
//			{
//				pNewLines = new CLines;
//				if (NULL == pNewLines)
//					goto Error;
//			}
//
//			//
//			// Add bands to lines
//			//
//
//			nLineIndex = -1;
//			nPrevLine = -100;
//
//			//
//			// Add all Non Stretch Bands
//			//
//			
//			nBandCount = aVisibleNonStretchBands.GetSize();
//			for (nBand = 0; nBand < nBandCount; nBand++)
//			{
//				pBand = aVisibleNonStretchBands.GetAt(nBand);
//				assert(pBand);
//				if (NULL == pBand)
//					goto Error;
//
//				if (pBand->bpV1.m_nDockLine != nPrevLine)
//				{
//					nLineIndex++;
//					nPrevLine = pBand->bpV1.m_nDockLine;
//
//					pNewLine = new CLine;
//					if (NULL == pNewLine)
//						goto Error;
//
//					hResult = pNewLines->m_aLines.Add(pNewLine);
//					if (FAILED(hResult))
//						goto Error;
//				}
//				hResult = pNewLine->m_aBands.Add(pBand);	
//				if (FAILED(hResult))
//					goto Error;
//			}
//			
//			//
//			// Now nLineIndex shows how many non stretch lines were filled
//			//
//			
//			nLineIndex++; 
//			int nAnchor;
//			
//			//
//			// Now insert stretch bands at or above MenuBar DockLine
//			//
//			
//			nBandCount = aVisibleStretchBands.GetSize();
//			for (nBand = 0; nBand < nBandCount; nBand++)
//			{
//				//
//				// For each stretch band
//				//
//
//				pStretchBand = aVisibleStretchBands.GetAt(nBand);
//				if (NULL == pStretchBand)
//				{
//					assert(pStretchBand);
//					continue;
//				}
//
//				if (ddBTStatusBar == pStretchBand->bpV1.m_btBands)
//				{
//					pStatusBar = pStretchBand;
//					continue;
//				}
//
//				for (nLine = 0; nLine < nLineIndex; nLine++)
//				{ 
//					//
//					// For each line 
//					//
//					
//					//
//					// Check against non stretch lines
//					//
//					
//					CBand* pCmpBand = pNewLines->m_aLines.GetAt(nLine)->m_aBands.GetAt(0);
//					assert(pCmpBand);
//					if (NULL == pCmpBand)
//						continue;
//
//					if (!pCmpBand->NeedsFullStretch())
//					{
//						nAnchor = pCmpBand->bpV1.m_nDockLine;
//
//						if (pStretchBand->bpV1.m_nDockLine <= nAnchor)
//						{
//							//
//							// This line goes on top of the other lines so we need to shift all of the
//							// lines down one. The array in the last position is empty so grab it and put
//							// it in nLine position.
//							//
//
//							nLineIndex++;
//
//							pNewLine = new CLine;
//							if (NULL == pNewLine)
//								goto Error;
//
//							hResult = pNewLines->m_aLines.InsertAt(nLine, pNewLine);
//							if (FAILED(hResult))
//								goto Error;
//
//							hResult = pNewLine->m_aBands.Add(pStretchBand);
//							if (FAILED(hResult))
//								goto Error;
//
//							break;
//						}
//					}
//				}
//				if (nLine == nLineIndex)
//				{
//					//
//					// We did not shift any lines down so put this stretch band on this line
//					//
//
//					pNewLine = new CLine;
//					if (NULL == pNewLine)
//						goto Error;
//
//					hResult = pNewLines->m_aLines.Add(pNewLine);
//					if (FAILED(hResult))
//						goto Error;
//
//					hResult = pNewLine->m_aBands.Add(pStretchBand);
//					if (FAILED(hResult))
//						goto Error;
//					nLineIndex++;
//				}
//			}
//			if (pStatusBar)
//			{
//				pNewLine = new CLine;
//				if (NULL == pNewLine)
//					goto Error;
//
//				hResult = pNewLines->m_aLines.Add(pNewLine);
//				if (FAILED(hResult))
//					goto Error;
//
//				hResult = pNewLine->m_aBands.Add(pStatusBar);
//				if (FAILED(hResult))
//					goto Error;
//			}
//		}
//		CATCH
//		{
//			assert(FALSE);
//			REPORTEXCEPTION(__FILE__, __LINE__)
//			goto Error;
//		}	
//
//		try
//		{
//			// 
//			// Now calc exact positions of bands inside the line and update size info
//			// 
//
//			m_theSplitters.Cleanup();
//
//			size.cx = size.cy = 0;
//			nX = nY = 0;
//			
//			SIZE sizeSqueezed;
//			int nTotalRequiredSize;	
//			for (nLine = 0; nLine < nNewLineCount; nLine++) 
//			{
//				try
//				{
//StartLine:
//					pLine = pNewLines->m_aLines.GetAt(nLine);
//
//					rcThickestBandInLine.SetEmpty();
//					dwLastDocked = 0;
//					bSameLastDocked = TRUE;
//					nTotalRequiredSize = 0;
//					nBandsInLine = pLine->m_aBands.GetSize();
//
//					pnSlack = new int[nBandsInLine];
//					assert(pnSlack);
//					if (NULL == pnSlack)
//						goto Error;
//
//					//
//					// Since we'll try to squeeze max number of bands in the same line to eliminate wrapping, 
//					// calculate slack for each band on this line
//					//
//
//					pBandLastDocked = NULL;
//					nPrevSplitterBandWidth = 0;
//					nSplitterBandWidth = 0;
//					nNonSplitterBandWidth = 0;
//					nNumOfSplitterBands = 0;
//					nContractedBandsWidth = 0;
//					nSlackTotal = 0;
//					BOOL bFirst = TRUE;
//					for (nBandIndex = 0; nBandIndex < nBandsInLine; nBandIndex++)
//					{
//						try
//						{
//							pBand = pLine->m_aBands.GetAt(nBandIndex);
//							assert(pBand);
//							if (NULL == pBand)
//								goto Error;
//
//							//
//							// Finding out how much space is required for all the bands to fit into a line and how much space 
//							// can be squeezed out of each band to make them fit on one line.
//							//
//
//							rcBound.SetEmpty();
//							if (bFirst)
//							{
//								dwLastDocked = pBand->m_dwLastDocked;
//								bFirst = FALSE;
//							}
//							else if (dwLastDocked < pBand->m_dwLastDocked)
//							{
//								dwLastDocked = pBand->m_dwLastDocked;
//								pBandLastDocked = pBand;
//								if (bSameLastDocked)
//									bSameLastDocked = FALSE;
//							}
//							if (bHorz)
//							{
//								//
//								// Horizontal Docking Area
//								//
//
//								rcBound.right = rcBound.left + nMaxSize;
//								rcBound.bottom = rcBound.top + 32767;
//								
//								// this is optimal
//								bResult = pBand->CalcLayout(rcBound, CBand::eLayoutHorz, rcReturn, TRUE); 
//								
//								nTotalRequiredSize += rcReturn.Width();
//
//								if (ddBFSizer & pBand->bpV1.m_dwFlags)
//								{
//									if (CBand::eExpanded == pBand->m_eExpandButtonState || CBand::eExpandedToContracted == pBand->m_eExpandButtonState)
//										pLine->m_eSplitterLineExpandedState = pBand->m_eExpandButtonState;
//
//									if (CBand::eContracted == pBand->m_eExpandButtonState)
//										nContractedBandsWidth += pBand->m_nAdjustForGrabBar + CBand::eBevelBorder2*2;
//
//									nNumOfSplitterBands++;
//									if (rcThickestBandInLine.Height() < rcReturn.Height())
//									{
//										rcThickestBandInLine.top = rcReturn.top;
//										rcThickestBandInLine.bottom = rcReturn.bottom;
//									}
//									if (!pLine->m_bSplitterLine)
//										pLine->m_bSplitterLine = TRUE;
//
//									pnSlack[nBandIndex] = 0;
//									nSplitterBandWidth += rcReturn.Width();
//									nPrevSplitterBandWidth += pBand->m_rcDock.Width();
//								}
//								else if (pBand->IsWrappable())
//								{
//									pnSlack[nBandIndex] = 0;
//									nNonSplitterBandWidth += rcReturn.Width();
//								}
//								else
//								{
//									sizeSqueezed = pBand->CalcSqueezedSize();
//									pnSlack[nBandIndex] = rcReturn.Width() - sizeSqueezed.cx;
//									nNonSplitterBandWidth += rcReturn.Width();
//								}
//							}
//							else
//							{
//								//
//								// Veritical Docking Area
//								//
//
//								rcBound.right = rcBound.left + 32767;
//								rcBound.bottom = rcBound.top + nMaxSize;
//
//								bResult = pBand->CalcLayout(rcBound, CBand::eLayoutVert, rcReturn, TRUE);
//								
//								nTotalRequiredSize += rcReturn.Height();
//
//								if (ddBFSizer & pBand->bpV1.m_dwFlags)
//								{
//									if (CBand::eExpanded == pBand->m_eExpandButtonState || CBand::eExpandedToContracted == pBand->m_eExpandButtonState)
//										pLine->m_eSplitterLineExpandedState = pBand->m_eExpandButtonState;
//
//									if (CBand::eContracted == pBand->m_eExpandButtonState)
//										nContractedBandsWidth += pBand->m_nAdjustForGrabBar + CBand::eBevelBorder2 * 2;
//
//									nNumOfSplitterBands++;
//									if (rcThickestBandInLine.Width() < rcReturn.Width())
//									{
//										rcThickestBandInLine.left = rcReturn.left;
//										rcThickestBandInLine.right = rcReturn.right;
//									}
//									if (!pLine->m_bSplitterLine)
//										pLine->m_bSplitterLine = TRUE;
//
//									pnSlack[nBandIndex] = 0;
//									nSplitterBandWidth += rcReturn.Height();
//									nPrevSplitterBandWidth += pBand->m_rcDock.Height();
//								}
//								else if (pBand->IsWrappable())
//								{
//									pnSlack[nBandIndex] = 0;
//									nNonSplitterBandWidth += rcReturn.Height();
//								}
//								else
//								{
//									sizeSqueezed = pBand->CalcSqueezedSize();
//									pnSlack[nBandIndex] = rcReturn.Height() - sizeSqueezed.cy;
//									nNonSplitterBandWidth += rcReturn.Height();
//								}
//							}
//							nSlackTotal += pnSlack[nBandIndex];
//						}
//						CATCH
//						{
//							assert(FALSE);
//							REPORTEXCEPTION(__FILE__, __LINE__)
//							goto Error;
//						}	
//					}
//				}
//				CATCH
//				{
//					assert(FALSE);
//					REPORTEXCEPTION(__FILE__, __LINE__)
//					goto Error;
//				}
//
//				//
//				// This part will calculate the offsets for each band in a line.
//				//
//				//
//
//				if (pLine->m_bSplitterLine)
//				{
//					try
//					{
//						//
//						// This line contains a Splitter type band which is special from the other 
//						// types of lines
//						//
//
//						nLineHeight = 0;
//						int nPrevWhatIsLeft = 0;
//						int nWhatIsLeft = 0;
//						int nWidthPerSplitterBand = 0;
//						int nTotalSpace = nMaxSize;
//						int nAvailSpace = nMaxSize;
//						nCurWidth = 0;
//						pLineSplitter = NULL;
//						pBandPrev = NULL;
//						if (!bSameLastDocked && ddBFSizer & pBandLastDocked->bpV1.m_dwFlags)
//						{
//							nNumOfSplitterBands--;
//							if (bHorz)
//							{
//								rcBound.right = rcBound.left + nMaxSize;
//								rcBound.bottom = rcBound.top + 32767;
//								bResult = pBandLastDocked->CalcLayout(rcBound, CBand::eLayoutHorz, rcReturn, FALSE);
//								nWhatIsLeft = nMaxSize - nNonSplitterBandWidth - rcReturn.Width();
//							}
//							else
//							{
//								rcBound.right = rcBound.left + 32767;
//								rcBound.bottom = rcBound.top + nMaxSize;
//								bResult = pBandLastDocked->CalcLayout(rcBound, CBand::eLayoutVert, rcReturn, FALSE);
//								nWhatIsLeft = nMaxSize - nNonSplitterBandWidth - rcReturn.Height();
//							}
//							//
//							// This could be negative
//							//
//
//							if (nWhatIsLeft < 50)
//							{
//								//
//								// Give all splitter bands equal weight 
//								//
//
//								nWhatIsLeft = nMaxSize - nNonSplitterBandWidth;
//								nNumOfSplitterBands++;
//								pBandLastDocked = NULL;
//							}
//						}
//						else
//						{
//							nWhatIsLeft = nMaxSize - nNonSplitterBandWidth;
//							if (nWhatIsLeft < 50 && nNonSplitterBandWidth != 0)
//							{
//								//
//								// We want to create a new line insert and insert the non form bands and restart the line
//								//
//								try 
//								{
//									//
//									// Now insert a new line
//									//
//									
//									pNewLine = new CLine;
//									if (NULL == pNewLine)
//										goto Error;
//
//									nBandsInLine--;
//									pBand = pLine->m_aBands.GetAt(nBandsInLine);
//									assert(pBand);
//									if (NULL == pBand)
//										goto Error;
//
//									pBand->bpV1.m_nDockOffset = 0;
//
//									pLine->m_aBands.RemoveAt(nBandsInLine);
//
//									//
//									// Insert the new line
//									//
//
//									if (nLine == nNewLineCount-1)
//									{
//
//										//
//										// At the end of the array
//										//
//
//										hResult = pNewLines->m_aLines.Add(pNewLine);
//										if (FAILED(hResult))
//											goto Error;
//									}
//									else
//									{
//										hResult = pNewLines->m_aLines.InsertAt(nLine + 1, pNewLine);
//										if (FAILED(hResult))
//											goto Error;
//									}
//
//									pNewLine->m_aBands.Add(pBand);
//									nNewLineCount++;
//									if (ddBFSizer & pBand->bpV1.m_dwFlags)
//										pNewLine->m_bSplitterLine = TRUE;
//									pLine->m_bSplitterLine = FALSE;
//									for (nBandIndex = 0; nBandIndex < nBandsInLine; nBandIndex++)
//									{
//										if (ddBFSizer & pLine->m_aBands.GetAt(nBandIndex)->bpV1.m_dwFlags)
//										{
//											pLine->m_bSplitterLine = TRUE;
//											break;
//										}
//									}
//									delete [] pnSlack;
//									pnSlack = NULL;
//									goto StartLine;
//								}
//								CATCH
//								{
//									assert(FALSE);
//									REPORTEXCEPTION(__FILE__, __LINE__)
//									goto Error;
//								}
//							}
//						}
//
//						if (nNumOfSplitterBands > 0)
//							nWidthPerSplitterBand = nWhatIsLeft / nNumOfSplitterBands;
//						else
//							nWidthPerSplitterBand = nWhatIsLeft;
//
//						nRemainder = nWhatIsLeft % nNumOfSplitterBands;
//
//						//
//						// For each band in this line
//						//
//
//						for (nBandIndex = 0; nBandIndex < nBandsInLine; nBandIndex++)
//						{
//							try
//							{
//								pBand->m_dwLastDocked = dwLastDocked;
//
//								pBand = pLine->m_aBands.GetAt(nBandIndex);
//								assert (pBand);
//								if (NULL == pBand)
//									goto Error;
//
//								pBand->m_nDockLineIndex = nLine;
//
//								rcBound.SetEmpty();
//								if (bHorz)
//								{
//									//
//									// Horizontial Docking Area
//									//
//
//									rcBound.right = rcBound.left + nMaxSize;
//									rcBound.bottom = rcBound.top + 32767;
//									
//									bResult = pBand->CalcLayout(rcBound, CBand::eLayoutHorz, rcReturn, FALSE);
//
//									//
//									// Horz Adjustments
//									//
//
//									if (ddBFSizer & pBand->bpV1.m_dwFlags)
//									{
//										pBand->m_rcDock.bottom = nY + rcThickestBandInLine.Height();
//										pBand->bpV1.m_rcHorzDockForm.top = 0;
//										pBand->bpV1.m_rcHorzDockForm.bottom = rcThickestBandInLine.Height() - pBand->BandEdge().cy;
//										if (1 == nBandsInLine)
//										{
//											pBand->m_rcDock.left = 0;
//											pBand->m_rcDock.right = nMaxSize;
//											pBand->bpV1.m_rcHorzDockForm.left = 0;
//											pBand->bpV1.m_rcHorzDockForm.right = nMaxSize - pBand->BandEdge().cx;
//											pBand->m_eExpandButtonState = CBand::eGrayed;
//										}
//										else if (bSameLastDocked)
//										{
//											if (pBand->m_bSizerResized)
//											{
//												// The form was resized
//												
//												int nNewWidth;
//												if (nPrevSplitterBandWidth > 0)
//													nNewWidth = MulDiv(pBand->m_rcDock.Width(), nWhatIsLeft, nPrevSplitterBandWidth);
//												else
//													nNewWidth = nWidthPerSplitterBand;
//												pBand->m_rcDock.left = nCurWidth;
//												pBand->m_rcDock.right = nCurWidth + nNewWidth;
//												pBand->bpV1.m_rcHorzDockForm.left = 0;
//												pBand->bpV1.m_rcHorzDockForm.right = nNewWidth - pBand->BandEdge().cx;
//												pBand->m_bSizerResized = FALSE;
//											}
//											else if (CBand::eExpanded == pLine->m_eSplitterLineExpandedState)
//											{
//												// For the expand and collapse buttons
//
//												if (pBand->m_eExpandButtonState == CBand::eExpanded)
//												{
//													int nTemp = nMaxSize - nNonSplitterBandWidth - nContractedBandsWidth;
//													pBand->m_rcDock.left = nCurWidth;
//													pBand->m_rcDock.right = nCurWidth + nTemp;
//													pBand->bpV1.m_rcHorzDockForm.left = 0;
//													pBand->bpV1.m_rcHorzDockForm.right = nTemp - pBand->BandEdge().cx;
//												}
//												else
//												{
//													pBand->m_rcDock.left = nCurWidth;
//													pBand->m_rcDock.right = nCurWidth + pBand->m_nAdjustForGrabBar + CBand::eBevelBorder2*2;
//													pBand->bpV1.m_rcHorzDockForm.left = 0;
//													pBand->bpV1.m_rcHorzDockForm.right = pBand->m_nAdjustForGrabBar + CBand::eBevelBorder2*2 - pBand->BandEdge().cx;
//												}
//											}
//											else if (CBand::eExpandedToContracted == pLine->m_eSplitterLineExpandedState || nMaxSize != nNonSplitterBandWidth + nSplitterBandWidth + nRemainder)
//											{
//												// Initial sizing of the forms in this line.
//												int nNewWidth = nWidthPerSplitterBand;
//												pBand->m_rcDock.left = nCurWidth;
//												pBand->m_rcDock.right = nCurWidth + nNewWidth;
//												pBand->bpV1.m_rcHorzDockForm.left = 0;
//												pBand->bpV1.m_rcHorzDockForm.right = nNewWidth - pBand->BandEdge().cx;
//											}
//											else
//											{
//												//
//												// Bands retain their current width
//												//
//
//												pBand->m_rcDock.left = nCurWidth;
//												pBand->m_rcDock.right = nCurWidth + rcReturn.Width();
//											}
//											if (CBand::eExpanded != pBand->m_eExpandButtonState && nNumOfSplitterBands > 1)
//												pBand->m_eExpandButtonState = CBand::eContracted;
//										}
//										else
//										{
//											//
//											// The last docked gets it's requested width
//											//
//
//											if (pBand != pBandLastDocked)
//											{
//												//
//												// Adjust the width of the last docked
//												//
//
//												rcReturn.right = rcReturn.left + nWidthPerSplitterBand;
//												pBand->bpV1.m_rcHorzDockForm.left = 0;
//												pBand->bpV1.m_rcHorzDockForm.right = nWidthPerSplitterBand - pBand->BandEdge().cx;
//											}
//											pBand->m_rcDock.left = nCurWidth;
//											pBand->m_rcDock.right = nCurWidth + rcReturn.Width();
//
//											if (CBand::eExpanded != pBand->m_eExpandButtonState && nNumOfSplitterBands > 1)
//												pBand->m_eExpandButtonState = CBand::eContracted;
//										}
//									}
//									else
//									{
//										//
//										// Non Splitter Band
//										//
//
//										pBand->m_rcDock.bottom = nY + rcReturn.Height();
//										pBand->m_rcDock.left = nCurWidth;
//										pBand->m_rcDock.right = nCurWidth + rcReturn.Width();
//									}
//									
//									if (pBand->NeedsFullStretch() || !(ddBFSizer & pBand->bpV1.m_dwFlags))
//										rcBound.right = rcBound.left + nMaxSize;
//									else
//										rcBound.right = rcBound.left + pBand->bpV1.m_rcHorzDockForm.Width();
//									rcBound.bottom = rcBound.top + 32767;
//									
//									//
//									// Commit the layout this time
//									//
//
//									bResult = pBand->CalcLayout(rcBound, CBand::eLayoutHorz, rcReturn, TRUE);
//
//									size = pBand->bpV1.m_rcHorzDockForm.Size();
//
//									pBand->SetCustomToolSizes(size);
//
//									//
//									// Vertical Adjustments
//									//
//									
//									pBand->m_rcDock.top = nY;
//
//									if (pBand->m_rcDock.right > nMaxSize)
//										pBand->m_rcDock.right = nMaxSize;
//									
//									// Set new DockOffset
//									if (-1 == pBand->m_nTempDockOffset)
//										pBand->bpV1.m_nDockOffset = nCurWidth; 
//									
//									if (rcReturn.Height() > nLineHeight)
//										nLineHeight = rcReturn.Height();
//
//									nCurWidth += rcReturn.Width();
//									nAvailSpace -= rcReturn.Width();
//								}
//								else
//								{
//									//
//									// Vertical Docking Area
//									//
//
//									rcBound.right = rcBound.left + 32767;
//									rcBound.bottom = rcBound.top + nMaxSize;
//
//									bResult = pBand->CalcLayout(rcBound, CBand::eLayoutVert, rcReturn, FALSE);
//									
//									pBand->m_rcDock.left = nX;
//
//									if (ddBFSizer & pBand->bpV1.m_dwFlags)
//									{
//										pBand->m_rcDock.right = nX + rcThickestBandInLine.Width();
//										pBand->bpV1.m_rcVertDockForm.left = 0;
//										pBand->bpV1.m_rcVertDockForm.right = rcThickestBandInLine.Width() - pBand->BandEdge().cx;
//
//										if (1 == nBandsInLine)
//										{
//											pBand->m_rcDock.top = 0;
//											pBand->m_rcDock.bottom = nMaxSize;
//											pBand->bpV1.m_rcVertDockForm.top = 0;
//											pBand->bpV1.m_rcVertDockForm.bottom = nMaxSize - pBand->BandEdge().cy;
//											pBand->m_eExpandButtonState = CBand::eGrayed;
//										}
//										else if (bSameLastDocked)
//										{
//											if (pBand->m_bSizerResized)
//											{
//												int nNewWidth;
//												if (nPrevSplitterBandWidth > 0)
//													nNewWidth = MulDiv(pBand->m_rcDock.Height(), nWhatIsLeft, nPrevSplitterBandWidth);
//												else
//													nNewWidth = nWidthPerSplitterBand;
//
//												pBand->m_rcDock.top = nCurWidth;
//												pBand->m_rcDock.bottom = nCurWidth + nNewWidth;
//												pBand->bpV1.m_rcVertDockForm.top = 0;
//												pBand->bpV1.m_rcVertDockForm.bottom = nNewWidth - pBand->BandEdge().cy;
//												pBand->m_bSizerResized = FALSE;
//											}
//											else if (CBand::eExpanded == pLine->m_eSplitterLineExpandedState)
//											{
//												if (pBand->m_eExpandButtonState == CBand::eExpanded)
//												{
//													int nTemp = nMaxSize - nNonSplitterBandWidth - nContractedBandsWidth;
//													pBand->m_rcDock.top = nCurWidth;
//													pBand->m_rcDock.bottom = nCurWidth + nTemp;
//													pBand->bpV1.m_rcVertDockForm.top = 0;
//													pBand->bpV1.m_rcVertDockForm.bottom = nTemp - pBand->BandEdge().cy;
//												}
//												else
//												{
//													pBand->m_rcDock.top = nCurWidth;
//													pBand->m_rcDock.bottom = nCurWidth + pBand->m_nAdjustForGrabBar + CBand::eBevelBorder2*2;
//													pBand->bpV1.m_rcVertDockForm.top = 0;
//													pBand->bpV1.m_rcVertDockForm.bottom = pBand->m_nAdjustForGrabBar + CBand::eBevelBorder2*2 - pBand->BandEdge().cy;
//												}
//											}
//											else if (CBand::eExpandedToContracted == pLine->m_eSplitterLineExpandedState || nMaxSize != nNonSplitterBandWidth + nSplitterBandWidth + nRemainder)
//											{
//												//
//												// Bands all get the same width
//												//
//
//												int nNewWidth = nWidthPerSplitterBand;
//												pBand->m_rcDock.top = nCurWidth;
//												pBand->m_rcDock.bottom = nCurWidth + nNewWidth;
//												pBand->bpV1.m_rcVertDockForm.top = 0;
//												pBand->bpV1.m_rcVertDockForm.bottom = nNewWidth - pBand->BandEdge().cy;
//											}
//											else
//											{
//												//
//												// Bands retain their current width
//												//
//
//												pBand->m_rcDock.top = nCurWidth;
//												pBand->m_rcDock.bottom = nCurWidth + rcReturn.Height();
//											}
//											if (CBand::eExpanded != pBand->m_eExpandButtonState && nNumOfSplitterBands > 1)
//												pBand->m_eExpandButtonState = CBand::eContracted;
//										}
//										else
//										{
//											//
//											// The last docked gets it's requested width
//											//
//
//											if (pBand != pBandLastDocked)
//											{
//												//
//												// Adjust the width of the last docked
//												//
//
//												rcReturn.bottom = rcReturn.top + nWidthPerSplitterBand;
//												pBand->bpV1.m_rcVertDockForm.top = 0;
//												pBand->bpV1.m_rcVertDockForm.bottom = nWidthPerSplitterBand - pBand->BandEdge().cy;
//											}
//											pBand->m_rcDock.top = nCurWidth;
//											pBand->m_rcDock.bottom = nCurWidth + rcReturn.Height();
//											if (CBand::eExpanded != pBand->m_eExpandButtonState && nNumOfSplitterBands > 1)
//												pBand->m_eExpandButtonState = CBand::eContracted;
//										}
//									}
//									else
//									{
//										//
//										// Non Splitter Band
//										//
//
//										pBand->m_rcDock.right = nX + rcReturn.Width();
//										pBand->m_rcDock.top = nCurWidth;
//										pBand->m_rcDock.bottom = nCurWidth + rcReturn.Height();
//									}
//
//									rcBound.right = rcBound.left + 32767;
//									if (pBand->NeedsFullStretch() || !(ddBFSizer & pBand->bpV1.m_dwFlags))
//										rcBound.bottom = rcBound.top + nMaxSize;
//									else
//										rcBound.bottom = rcBound.top + pBand->bpV1.m_rcVertDockForm.Height();
//
//									//
//									// Commit the layout this time
//									//
//
//									bResult = pBand->CalcLayout(rcBound, CBand::eLayoutVert, rcReturn, TRUE);
//
//									//
//									// Set custom tool size
//									//
//									
//									size = pBand->bpV1.m_rcVertDockForm.Size();
//
//									pBand->SetCustomToolSizes(size);
//
//									if (pBand->m_rcDock.bottom > nMaxSize)
//										pBand->m_rcDock.bottom = nMaxSize;
//
//									//
//									// Set new dock offset
//									//
//
//									if (-1 == pBand->m_nTempDockOffset)
//										pBand->bpV1.m_nDockOffset = nCurWidth; 
//									
//									if (rcReturn.Width() > nLineHeight)
//										nLineHeight = rcReturn.Width();
//									
//									nCurWidth += rcReturn.Height();
//									nAvailSpace -= rcReturn.Height();
//								}
//
//								//
//								// Calculate the splitters
//								//
//
//								if (ddBFSizer & pBand->bpV1.m_dwFlags)
//								{
//									// 
//									// This Band needs a splitter, This is the splitter on the 
//									// outside edge of a line
//									//
//
//									if (NULL == pLineSplitter)
//									{
//										//
//										// Create a new Splitter, one Splitter for the whole line 
//										//
//
//										pLineSplitter = new CSplitter(m_pBar->GetParentWindow(), this);
//										assert(pLineSplitter);
//										if (pLineSplitter)
//											m_theSplitters.AddSplitter(pLineSplitter, pBand->m_rcDock);
//									}
//
//									//
//									// Add the band to the splitter
//									//
//									
//									if (pLineSplitter)
//										pLineSplitter->AddInsideBand(pBand);
//
//									//
//									// If not the last band and if not the second to last band and the last band is a non
//									// splitter band
//									//
//
//									if (nBandIndex != nBandsInLine - 1 && !(nBandIndex == nBandsInLine - 2 && !(ddBFSizer & pLine->m_aBands.GetAt(nBandsInLine-1)->bpV1.m_dwFlags)))
//									{
//										// 
//										// This Band needs a between band splitter
//										//
//									
//										pBetweenBandsSplitter = new CBetweenBandsSplitter(pLineSplitter, m_pBar->GetParentWindow(), this, pLine->m_aBands);
//										if (pBetweenBandsSplitter)
//										{
//											m_theSplitters.AddSplitter(pBetweenBandsSplitter, pBand->m_rcDock);
//											pLineSplitter->AddBetweenBandSplitters(pBetweenBandsSplitter);
//											pBetweenBandsSplitter->AddInsideBand(pBand);
//											pBetweenBandsSplitter->AddOutsideBand(pLine->m_aBands.GetAt(nBandIndex+1));
//										}
//									}
//									if (pBandPrev && (!(ddBFSizer & pBandPrev->bpV1.m_dwFlags) && 1 != nBandIndex))
//									{
//										pBetweenBandsSplitter = new CBetweenBandsSplitter(pLineSplitter, m_pBar->GetParentWindow(), this, pLine->m_aBands);
//										if (pBetweenBandsSplitter)
//										{
//											CRect rcSplitter = pBandPrev->m_rcDock;
//											if (bHorz)
//											{
//												rcSplitter.top = pBand->m_rcDock.top;
//												rcSplitter.bottom = pBand->m_rcDock.bottom;
//											}
//											else
//											{
//												rcSplitter.left = pBand->m_rcDock.left;
//												rcSplitter.right = pBand->m_rcDock.right;
//											}
//
//											m_theSplitters.AddSplitter(pBetweenBandsSplitter, rcSplitter);
//											pLineSplitter->AddBetweenBandSplitters(pBetweenBandsSplitter);
//											pBetweenBandsSplitter->AddInsideBand(pBandPrev);
//											pBetweenBandsSplitter->AddOutsideBand(pBand);
//										}
//									}
//								}
//								pBandPrev = pBand;
//							}
//							CATCH
//							{
//								assert(FALSE);
//								REPORTEXCEPTION(__FILE__, __LINE__)
//								goto Error;
//							}
//						}
//					}
//					CATCH
//					{
//						assert(FALSE);
//						REPORTEXCEPTION(__FILE__, __LINE__)
//						goto Error;
//					}
//				}
//				else if (nTotalRequiredSize <= nMaxSize || 1 == nBandsInLine)
//				{
//					//
//					// All the bands will fit on this line without any problems
//					//
//
//					try
//					{
//						nLineHeight = 0;
//						int nTotalSpace = nMaxSize;
//						int nAvailSpace;
//						int nSpace;
//Retry:
//						nAvailSpace = nTotalSpace;
//						nCurWidth = 0;
//						for (nBandIndex = 0; nBandIndex < nBandsInLine; nBandIndex++)
//						{
//							try
//							{
//								pBand = pLine->m_aBands.GetAt(nBandIndex);
//								assert (pBand);
//								if (NULL == pBand)
//									goto Error;
//
//								if ((-1 == pBand->m_nTempDockOffset ? pBand->bpV1.m_nDockOffset : pBand->m_nTempDockOffset) > nCurWidth)
//								{
//									nSpace = (-1 == pBand->m_nTempDockOffset ? pBand->bpV1.m_nDockOffset : pBand->m_nTempDockOffset) - nCurWidth;
//									if (nSpace < nAvailSpace)
//									{
//										nCurWidth += nSpace;
//										nAvailSpace -= nSpace;
//									}
//									else
//									{
//										if (nAvailSpace > 0)
//											nCurWidth += nAvailSpace;
//										nAvailSpace = 0;
//									}
//								}
//					
//								rcBound.SetEmpty();
//								if (bHorz)
//								{
//									//
//									// Horizontial Docking Area
//									//
//
//									rcBound.right = rcBound.left + nMaxSize;
//									rcBound.bottom = rcBound.top + 32767;
//									
//									bResult = pBand->CalcLayout(rcBound, CBand::eLayoutHorz, rcReturn, TRUE);
//
//									if (pBand->NeedsFullStretch())
//									{
//										pBand->m_rcDock.left = 0;
//										pBand->m_rcDock.right = nMaxSize;
//									}
//									else
//									{
//										pBand->m_rcDock.left = nCurWidth;
//										pBand->m_rcDock.right = nCurWidth + rcReturn.Width();
//									}
//									
//									if (pBand->m_rcDock.right > nMaxSize)
//										pBand->m_rcDock.right = nMaxSize;
//									
//									pBand->m_rcDock.top = nY;
//									pBand->m_rcDock.bottom = nY + rcReturn.Height();
//
//									// Set new DockOffset
//									if (-1 == pBand->m_nTempDockOffset)
//										pBand->bpV1.m_nDockOffset = nCurWidth; 
//									
//									if (rcReturn.Height() > nLineHeight)
//										nLineHeight = rcReturn.Height();
//
//									nCurWidth += rcReturn.Width();
//					
//									if (nCurWidth > nMaxSize && nTotalSpace > 0)
//									{ 
//										// Ops we ran out of space here
//										nTotalSpace -= rcReturn.Width();
//										goto Retry;
//									}
//									nAvailSpace -= rcReturn.Width();
//								}
//								else
//								{
//									//
//									// Vertical Docking Area
//									//
//
//									rcBound.right = rcBound.left + 32767;
//									rcBound.bottom = rcBound.top + nMaxSize;
//
//									bResult = pBand->CalcLayout(rcBound, CBand::eLayoutVert, rcReturn, TRUE);
//									if (pBand->NeedsFullStretch())
//									{
//										pBand->m_rcDock.top = 0;
//										pBand->m_rcDock.bottom = nMaxSize;
//									}
//									else
//									{
//										pBand->m_rcDock.top = nCurWidth;
//										pBand->m_rcDock.bottom = nCurWidth + rcReturn.Height();
//									}
//
//									if (pBand->m_rcDock.bottom > nMaxSize)
//										pBand->m_rcDock.bottom = nMaxSize;
//
//									pBand->m_rcDock.left = nX;
//									pBand->m_rcDock.right = nX + rcReturn.Width();
//									
//									// set new dock offset
//									if (-1 == pBand->m_nTempDockOffset)
//										pBand->bpV1.m_nDockOffset = nCurWidth; 
//									
//									if (rcReturn.Width() > nLineHeight)
//										nLineHeight = rcReturn.Width();
//									
//									nCurWidth += rcReturn.Height();
//									
//									if (nCurWidth > nMaxSize && nTotalSpace > 0)
//									{ 
//										//
//										// Oops we ran out of space here
//										//
//
//										nTotalSpace -= rcReturn.Height();
//										goto Retry;
//									}
//									nAvailSpace -= rcReturn.Height();
//								}
//								pBand->m_nDockLineIndex = nLine;
//							}
//							CATCH
//							{
//								assert(FALSE);
//								REPORTEXCEPTION(__FILE__, __LINE__)
//								goto Error;
//							}	
//						}
//					}
//					CATCH
//					{
//						assert(FALSE);
//						REPORTEXCEPTION(__FILE__, __LINE__)
//						goto Error;
//					}
//				}
//				else
//				{
//					if (nTotalRequiredSize - nSlackTotal > nMaxSize)			
//					{ 
//						try 
//						{
//							//
//							// Since nTotalRequiredSize > nMaxSize 
//							// 
//							// We have to squeeze the bands so they will fit everything together, if bands need to be 
//							// wrapped we will insert a new line and attach this guy.
//							//
//
//							//
//							// Oh boy doesn't fit in a line no matter how much we try. Wrap last tool to 
//							// next line
//							//
//
//							int nLastBandInLine = pLine->m_aBands.GetSize()-1;
//
//							CBand* pWrapBand = pLine->m_aBands.GetAt(nLastBandInLine);
//							assert (pWrapBand);
//							if (NULL == pWrapBand)
//								goto Error;
//
//							pLine->m_aBands.RemoveAt(nLastBandInLine);
//							
//							//
//							// Now insert a new line
//							//
//							
//							try 
//							{
//								pNewLine = new CLine;
//								if (NULL == pNewLine)
//									goto Error;
//
//								//
//								// Insert the new line
//								//
//
//								if (nLine == nNewLineCount-1)
//								{
//
//									//
//									// At the end of the array
//									//
//
//									hResult = pNewLines->m_aLines.Add(pNewLine);
//									if (FAILED(hResult))
//										goto Error;
//								}
//								else
//								{
//									hResult = pNewLines->m_aLines.InsertAt(nLine + 1, pNewLine);
//									if (FAILED(hResult))
//										goto Error;
//								}
//								hResult = pNewLine->m_aBands.Add(pWrapBand);
//								if (FAILED(hResult))
//									goto Error;
//								pWrapBand->bpV1.m_nDockOffset = 0;
//								++nNewLineCount;
//								delete [] pnSlack;
//								pnSlack = NULL;
//								goto StartLine;
//							}
//							CATCH
//							{
//								assert(FALSE);
//								REPORTEXCEPTION(__FILE__, __LINE__)
//								goto Error;
//							}
//						}
//						CATCH
//						{
//							assert(FALSE);
//							REPORTEXCEPTION(__FILE__, __LINE__)
//						}
//					}
//					else
//					{	
//						try
//						{
//							//
//							// If we squeeze the bands enough they will fit on this line so do it
//							// first reduce the slack values uniformly
//							// 
//
//							int nTotalReduction = nTotalRequiredSize - nMaxSize;
//							int nEvenDistributionAmount = nTotalReduction / nBandsInLine;
//							int i;
//							int* pnSizeReduction = new int[nBandsInLine];
//							assert(pnSizeReduction);
//							if (NULL == pnSizeReduction)
//								goto Error;
//
//							memset(pnSizeReduction, 0, nBandsInLine * sizeof(int));
//							// Now distribute rest of TotalReduction
//							for (i = nBandsInLine-1; i >= 0; --i)
//							{
//								if (0 != pnSlack[i])
//								{
//									if (pnSlack[i] > nTotalReduction)
//									{
//										pnSlack[i] -= nTotalReduction;
//										pnSizeReduction[i] += nTotalReduction;
//										// we are done
//										break; 
//									}
//									else
//									{
//										pnSizeReduction[i] += pnSlack[i];
//										nTotalReduction -= pnSlack[i];
//										pnSlack[i] = 0;
//									}
//								}
//								if (nTotalReduction <= 0)
//									break;
//							}
//
//							//
//							// Distribution complete. Now layout the bands using 
//							// their sizeReduction values
//							//
//							
//							int dx;
//							nCurWidth = 0;
//							nLineHeight = 0;
//							for (nBandIndex = 0; nBandIndex < nBandsInLine; nBandIndex++)
//							{
//								pBand = pLine->m_aBands.GetAt(nBandIndex);
//								assert(pBand);
//								if (NULL == pBand)
//									goto Error;
//
//								rcBound.SetEmpty();
//								if (bHorz)
//								{
//									//
//									// Horizontial Docking Area
//									//
//
//									rcBound.right = rcBound.left + nMaxSize;
//									rcBound.bottom = rcBound.top + 32767;
//									
//									bResult = pBand->CalcLayout(rcBound, CBand::eLayoutHorz, rcReturn, TRUE);
//									
//									dx = rcReturn.right - rcReturn.left - pnSizeReduction[nBandIndex];
//
//									if (pBand->NeedsFullStretch())
//									{
//										pBand->m_rcDock.left = 0;
//										pBand->m_rcDock.right = nMaxSize;
//									}
//									else
//									{
//										pBand->m_rcDock.left = nCurWidth;
//										pBand->m_rcDock.right = nCurWidth + dx;
//									}
//
//									pBand->m_rcDock.top = nY;
//									pBand->m_rcDock.bottom = nY + rcReturn.Height();
//									if (-1 == pBand->m_nTempDockOffset)
//										pBand->bpV1.m_nDockOffset = nCurWidth; 
//
//									//
//									// Set new DockOffset
//									//
//									
//									if (rcReturn.bottom - rcReturn.top > nLineHeight)
//										nLineHeight = rcReturn.Height();
//								}
//								else
//								{
//									//
//									// Vertical Docking Area
//									//
//
//									rcBound.right = rcBound.left + 32767;
//									rcBound.bottom = rcBound.top + nMaxSize;
//
//									bResult = pBand->CalcLayout(rcBound, CBand::eLayoutVert, rcReturn, TRUE);
//									
//									dx = rcReturn.Height();
//									dx -= pnSizeReduction[nBandIndex];
//
//									if (pBand->NeedsFullStretch())
//									{
//										pBand->m_rcDock.top = 0;
//										pBand->m_rcDock.bottom = nMaxSize;
//									}
//									else
//									{
//										pBand->m_rcDock.top = nCurWidth;
//										pBand->m_rcDock.bottom = nCurWidth + dx;
//									}
//									
//									pBand->m_rcDock.left = nX;
//									pBand->m_rcDock.right = nX + rcReturn.Width();
//									
//									if (-1 == pBand->m_nTempDockOffset)
//										pBand->bpV1.m_nDockOffset = nCurWidth; 
//
//									// Set new dock offset
//									if (rcReturn.Width() > nLineHeight)
//										nLineHeight = rcReturn.Width();
//								}
//								nCurWidth += dx;
//								pBand->m_nDockLineIndex = nLine;
//							}
//							// end layout of this line
//							delete [] pnSizeReduction;
//						}
//						CATCH
//						{
//							assert(FALSE);
//							REPORTEXCEPTION(__FILE__, __LINE__)
//							goto Error;
//						}
//					}
//				}
//				try
//				{
//					delete [] pnSlack;
//					pnSlack = NULL;
//					pLine->m_nThickness = nLineHeight;
//					if (ddDATop == m_daDockingArea || ddDABottom == m_daDockingArea)
//					{	
//						rcNew.bottom += nLineHeight;
//						nY += nLineHeight;
//					}
//					else
//					{
//						rcNew.right += nLineHeight;
//						nX += nLineHeight;
//					}
//					rcThickestBandInLine.SetEmpty();
//				}
//				CATCH
//				{
//					assert(FALSE);
//					REPORTEXCEPTION(__FILE__, __LINE__)
//					goto Error;
//				}	
//			} // End For Loop on Lines
//		}
//		CATCH
//		{
//			assert(FALSE);
//			REPORTEXCEPTION(__FILE__, __LINE__)
//			goto Error;
//		}	
//
//		try
//		{
//			delete m_pLines;
//			m_pLines = NULL;
//			m_pLines = pNewLines;
//			m_nLineCount = nNewLineCount;
//		}
//		CATCH
//		{
//			assert(FALSE);
//			REPORTEXCEPTION(__FILE__, __LINE__)
//			goto Error;
//		}
//	
//		//
//		// Why is this here
//		//
//		
//		if (bCommit)
//			m_rcDock = rcNew;
//
//		size.cx = rcNew.Width();
//		size.cy = rcNew.Height();
//	}
//	CATCH
//	{
//		assert(FALSE);
//		REPORTEXCEPTION(__FILE__, __LINE__)
//		goto Error;
//	}
//	return size;
//
//Error:
//	assert(FALSE);
//	delete [] pnSlack;
//	delete pBetweenBandsSplitter;
//	size.cx = size.cy = 0;
//	if (GetGlobals().GetEventLog())
//		GetGlobals().GetEventLog()->WriteCustom(_T("Error occurred in CDock::ReCalcDock"), EVENTLOG_ERROR_TYPE);
	return size;
}

//
// HitTest
//

CBand* CDock::HitTest(const POINT& pt)
{
	try
	{
		CBand* pBand;
		int nBandCount = m_aBands.GetSize();
		for (int nBand = 0; nBand < nBandCount; nBand++)
		{
			pBand = m_aBands.GetAt(nBand);
			if (VARIANT_TRUE != pBand->bpV1.m_vbVisible)
				continue;
			if (pBand->bpV1.m_btBands == ddBTPopup)
				continue;
			if (PtInRect(&(pBand->m_rcDock), pt))
				return pBand;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return NULL;
}

void CDock::Cache(CBand* pRemoveBand)
{
	if (m_pCacheLines)
		m_pCacheLines->Cleanup();
	else
		m_pCacheLines = new CLines;

	m_nCacheLineCount = m_nLineCount;

	if (0 == m_nLineCount)
		return;

	HRESULT hResult;
	CBand* pBand;
	CLine* pNewLine;
	int nDestLineIndex = 0;
	int nBandCount = 0;
	int nMaxSize = 0;
	int nNewSize = 0;
	int nBand = 0;
	m_nCacheSize = 0;
	for (int nLine = 0; nLine < m_nLineCount; nLine++)
	{
		nBandCount = m_pLines->GetLine(nLine)->BandCount();

		pNewLine = new CLine;
		if (NULL == pNewLine)
			continue;

		hResult = m_pCacheLines->m_aLines.Add(pNewLine);
		if (FAILED(hResult))
		{
			delete pNewLine;
			continue;
		}

		nBandCount = m_pLines->GetLine(nLine)->BandCount();
		m_pCacheLines->GetLine(nLine)->m_nThickness = m_pLines->GetLine(nLine)->m_nThickness;
		nMaxSize = 0;
		for (nBand = 0; nBand < nBandCount; nBand++)
		{
			pBand = m_pLines->GetLine(nLine)->GetBand(nBand);
			if (pBand == pRemoveBand)
			{
				if (1 == m_pLines->GetLine(nLine)->BandCount())
				{
					--nDestLineIndex;
					--m_nCacheLineCount;
					break;
				}
			}
			else
			{
				if (ddDABottom == m_daDockingArea || ddDATop == m_daDockingArea)
				{
					nNewSize = pBand->m_rcDock.Height();
					if (nNewSize > nMaxSize)
						nMaxSize = nNewSize;
				}
				if (ddDARight == m_daDockingArea || ddDALeft == m_daDockingArea)
				{
					nNewSize = pBand->m_rcDock.Width();
					if (nNewSize > nMaxSize)
						nMaxSize = nNewSize;
				}
				m_pCacheLines->GetLine(nDestLineIndex)->m_aBands.Add(pBand);
			}
		}
		++nDestLineIndex;
		m_nCacheSize += nMaxSize; // used for adjusting dockRect on bottom and right areas
	}
	
}

// 
// GetDockLine
//

void CDock::GetDockLine(CBand* pBand, const POINT& ptBand)
{
	assert(m_pLines);
	CBand* pTmpBand;
	CLine* pLine;
	CRect rcDock;
	GetWindowRect(rcDock);

	int nCurrentMousePos;
	int nDockPos;
	int nDockMargin;
	int nThickness;
	int nTool = 0;
	int nLine = 0; 

	switch (m_daDockingArea)
	{
	case ddDATop:
		nCurrentMousePos = ptBand.y;
		nDockPos = rcDock.top;
		nDockMargin = eLineGravity;
		nThickness = pBand->m_rcDock.Height();
		break;
	
	case ddDABottom:
		nCurrentMousePos = ptBand.y;
		nDockPos = rcDock.top;
		nDockMargin = eLineGravity;
		nThickness = pBand->m_rcDock.Height();
		break;
	
	case ddDALeft:
		nCurrentMousePos = ptBand.x;
		nDockPos = rcDock.left;
		nDockMargin = eLineGravity;
		nThickness = pBand->m_rcDock.Width();
		break;

	case ddDARight:
		nCurrentMousePos = ptBand.x;
		nDockPos = rcDock.left;
		nDockMargin = eLineGravity;
		nThickness = pBand->m_rcDock.Width();
		break;
	}

	if (m_daDockingArea != pBand->bpV1.m_daDockingArea)
	{
		try
		{
			if (nCurrentMousePos < nDockPos)
			{
				//
				// Top
				//

				pBand->bpV1.m_nDockLine = 0;
				
				//
				// Move the other lines down one dock line
				//

				for (nLine = 0; nLine < m_nLineCount; nLine++)
				{
					int nBandCount = m_pLines->GetLine(nLine)->m_aBands.GetSize();
					for (int nBand = 0; nBand < nBandCount; nBand++)
						m_pLines->GetLine(nLine)->GetBand(nBand)->bpV1.m_nDockLine = nLine + 1;
				}
			}
			else
			{
				//
				// Bottom
				//
				pBand->bpV1.m_nDockLine = m_nLineCount + 1;
				for (int nLine = 0; nLine < m_nLineCount; nLine++)
				{
					int nBandCount = m_pLines->GetLine(nLine)->m_aBands.GetSize();
					for (int nBand = 0; nBand < nBandCount; nBand++)
						m_pLines->GetLine(nLine)->GetBand(nBand)->bpV1.m_nDockLine = nLine;
				}
			}
			pBand->bpV1.m_daDockingArea = m_daDockingArea;
			pBand->m_dwLastDocked = m_pBar->IncrementLastDocked();
			if (m_pBar->m_pDesigner)
				m_pBar->m_pDesigner->SetDirty();
		}
		CATCH
		{
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
		}
	}
	else
	{
		try
		{
			//
			// Find the current dock line position
			//

			int nBandCurrentDockLinePos = nDockPos;
			for (int nLine = 0; nLine < pBand->m_nDockLineIndex; nLine++)
				nBandCurrentDockLinePos += m_pLines->GetLine(nLine)->m_nThickness;
			
			if (m_pLines->GetLine(pBand->m_nDockLineIndex)->BandCount() > 1)
			{
				// If there is more than one band in this line.

				try
				{
					if (nCurrentMousePos < nBandCurrentDockLinePos - nDockMargin)
					{
						//
						// We are moving up the DockArea
						//
						// This band stays on the same DockLine. The Bands below this one moves
						// down 1 DockLine. 
						//

						for (nLine = 0; nLine < pBand->m_nDockLineIndex; nLine++)
						{
							pLine = m_pLines->GetLine(nLine);
							int nBandCount = pLine->BandCount();
							for (int nBand = 0; nBand < nBandCount; nBand++)
							{
								pTmpBand = pLine->GetBand(nBand);
								if (pTmpBand == pBand)
									continue;

								pTmpBand->bpV1.m_nDockLine = nLine;
							}
						}
						for (nLine = pBand->m_nDockLineIndex; nLine < m_nLineCount; nLine++)
						{
							pLine = m_pLines->GetLine(nLine);
							int nBandCount = pLine->BandCount();
							for (int nBand = 0; nBand < nBandCount; nBand++)
							{
								pTmpBand = pLine->GetBand(nBand);
								if (pTmpBand == pBand)
								{
									pTmpBand->bpV1.m_nDockLine = nLine;
									continue;
								}

								pTmpBand->bpV1.m_nDockLine = nLine + 1;
							}
						}
					}
					else if (nCurrentMousePos > nBandCurrentDockLinePos + m_pLines->GetLine(nLine)->m_nThickness + (2 * nDockMargin))
					{
						//
						// We are moving down the DockArea
						//
						// This band moves down one DockLine creating a new DockLine for it's self. 
						// The Bands below this one move down one DockLine. 
						//

						pBand->bpV1.m_nDockLine++;

						for (nLine = pBand->m_nDockLineIndex + 1; nLine < m_nLineCount; nLine++)
						{
							int nBandCount = m_pLines->GetLine(nLine)->BandCount();
							for (int nBand = 0; nBand < nBandCount; nBand++)
							{
								if (m_pLines->GetLine(nLine)->GetBand(nBand) == pBand)
									continue;

								m_pLines->GetLine(nLine)->GetBand(nBand)->bpV1.m_nDockLine++;
							}
						}
					}
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
			}
			else
			{
				try
				{
					if (nLine != 0 && nCurrentMousePos < nBandCurrentDockLinePos - (m_pLines->GetLine(nLine - 1)->m_nThickness / 2))
					{
						//
						// We are moving up the DockArea
						//

						if (0 == pBand->m_nDockLineIndex)
						{
							// We are a the top of the dock area.  
							if (pBand->bpV1.m_nDockLine >= 0)
								pBand->bpV1.m_nDockLine--;
						}
						else
						{
							pTmpBand = m_pLines->GetLine(pBand->m_nDockLineIndex - 1)->GetBand(0);
							if (pBand->NeedsFullStretch() || pTmpBand->NeedsFullStretch())
							{
								int nTemp = pTmpBand->bpV1.m_nDockLine;
								
								int nBandCount = m_pLines->GetLine(pBand->m_nDockLineIndex - 1)->m_aBands.GetSize();
								for (int nBand = 0; nBand < nBandCount; nBand++)
									m_pLines->GetLine(pBand->m_nDockLineIndex - 1)->GetBand(nBand)->bpV1.m_nDockLine = pBand->bpV1.m_nDockLine;

								pBand->bpV1.m_nDockLine = nTemp;
							}
							else
								pBand->bpV1.m_nDockLine--;
						}
					}
					else if (pBand->m_nDockLineIndex < m_pLines->LineCount() - 1 && nCurrentMousePos > nBandCurrentDockLinePos + m_pLines->GetLine(nLine)->m_nThickness + (m_pLines->GetLine(nLine + 1)->m_nThickness / 2))
					{
						//
						// We are moving Down the DockArea
						//

						if (pBand->m_nDockLineIndex == m_nLineCount - 1)
						{
							// We are at the bottom of the dockarea.
							pBand->bpV1.m_nDockLine++;
						}
						else
						{
							pTmpBand = m_pLines->GetLine(pBand->m_nDockLineIndex + 1)->GetBand(0);
							if (pBand->NeedsFullStretch() || pTmpBand->NeedsFullStretch())
							{
								int nTemp = pTmpBand->bpV1.m_nDockLine;
								
								int nBandCount = m_pLines->GetLine(pBand->m_nDockLineIndex + 1)->m_aBands.GetSize();
								for (int nBand = 0; nBand < nBandCount; nBand++)
									m_pLines->GetLine(pBand->m_nDockLineIndex + 1)->GetBand(nBand)->bpV1.m_nDockLine = pBand->bpV1.m_nDockLine;

								pBand->bpV1.m_nDockLine = nTemp;
							}
							else
							{
								if (pBand->m_nDockLineIndex == pTmpBand->m_nDockLineIndex)
									pBand->bpV1.m_nDockLine++;
								else
								{
									int nBandCount = m_pLines->GetLine(pBand->m_nDockLineIndex + 1)->m_aBands.GetSize();
									for (int nBand = 0; nBand < nBandCount; nBand++)
										m_pLines->GetLine(pBand->m_nDockLineIndex + 1)->GetBand(nBand)->bpV1.m_nDockLine = pBand->bpV1.m_nDockLine;
								}
							}
						}
					}
					else if (pBand->bpV1.m_nDockLine > 0)
					{
						pBand->bpV1.m_nDockLine = m_pLines->GetLine(nLine)->GetBand(0)->bpV1.m_nDockLine;
					}
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
			}
		}
		CATCH
		{
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
		}
	}
}

//
// Reset
//

void CDock::Reset()
{
///	delete m_pLines;
//	m_pLines = NULL;
	delete m_pCacheLines;
	m_pCacheLines = 0;
//	m_nLineCount = 0;
	m_nCacheLineCount = 0;
	m_aBands.RemoveAll();
	m_rcDock.SetEmpty();
}

//
// CDockMgr
//

CDockMgr::CDockMgr(CBar* pBar)
	: m_pBar(pBar),
	  m_prcLast(NULL),
	  m_pBand(NULL),
	  m_hWndTemp(NULL),
	  m_pPrevDockArea(NULL)
{
	m_bHasMoved = FALSE;
	m_bDragging = FALSE;
	
	m_pDockObject[ddDATop] = new CDock(this, ddDATop);
	assert(m_pDockObject[ddDATop]);

	m_pDockObject[ddDABottom] = new CDock(this, ddDABottom);
	assert(m_pDockObject[ddDABottom]);

	m_pDockObject[ddDALeft] = new CDock(this, ddDALeft);
	assert(m_pDockObject[ddDALeft]);

	m_pDockObject[ddDARight] = new CDock(this, ddDARight);
	assert(m_pDockObject[ddDARight]);

	m_nPrevDragSize = eDockBorder;
}

//
// ~CDockMgr
//

CDockMgr::~CDockMgr()
{
	for (int nIndex = 0; nIndex < eNumOfDockAreas; nIndex++)
	{
		delete m_pDockObject[nIndex];
		m_pDockObject[nIndex] = NULL;
	}
}

//
// ResizeDockableForms
//

void CDockMgr::ResizeDockableForms()
{
	for (int nIndex = 0; nIndex < eNumOfDockAreas; nIndex++)
		m_pDockObject[nIndex]->ResizeDockableForms();
}

//
// Invalidate
//

void CDockMgr::Invalidate(DockingAreaTypes daDockingArea, CBand* pBand)
{
	if (daDockingArea >= 0 && daDockingArea < ddDAFloat)
	{
		if (NULL == pBand)
			m_pDockObject[daDockingArea]->InvalidateRect(NULL, FALSE); 
		else
		{
			CRect rcTemp = pBand->m_rcDock;
			m_pDockObject[daDockingArea]->InvalidateRect(&rcTemp, FALSE); 
		}
	}
}

//
// HitTest
//

CDock* CDockMgr::HitTest(const POINT& pt)
{
	CRect rcTemp;
	HWND hWndParent = m_pBar->GetParentWindow();

	for (int nDock = 0; nDock < eNumOfDockAreas; nDock++)
	{
		rcTemp = m_pDockObject[nDock]->m_rcClientPosition;
		ClientToScreen(hWndParent, rcTemp);
		
		//
		// rcTemp is now in screen coordinates so pt must be
		// in screen coordinates
		//

		rcTemp.Inflate(eBandDockGravity, eBandDockGravity);

		if (PtInRect(&rcTemp, pt))
			return m_pDockObject[nDock]; 
	}
	return NULL;
}

//
// RecalcLayout
//

BOOL CDockMgr::RecalcLayout(SIZE sizeWin, BOOL bJustRecalc)
{
	try
	{
		//
		// This is used to detect disappearing dock areas
		//
		
		// 
		// We blow everything away because the user may have changed a band(s) docking area 
		// and then called RecalcLayout to change a band(s) docking area
		//

		int nBand = 0;
		int nDock = 0;

		for (nDock = 0; nDock < eNumOfDockAreas; nDock++)
			m_pDockObject[nDock]->Reset(); 

		CBand*  pBand;
		CBands* pBands = m_pBar->GetBands();
		assert(pBands);
		int nBandCount = pBands->GetBandCount();
		for (nBand = 0; nBand < nBandCount; ++nBand)
		{
			try
			{
				pBand = pBands->GetBand(nBand);
				assert(pBand);
				if (NULL == pBand)
					continue;
				
				//
				// Make sure user hasn't detached at wrong time. 
				//
				
				pBand->ResetToolPressedState(); 

				if (pBand->bpV1.m_daDockingArea >= ddDATop && pBand->bpV1.m_daDockingArea <= ddDARight)
				{
					m_pDockObject[pBand->bpV1.m_daDockingArea]->m_aBands.Add(pBand);
					pBand->m_pDock = m_pDockObject[pBand->bpV1.m_daDockingArea];
				}

				pBand->m_nCollectionIndex = nBand;

				//
				// Now check floating bands
				//
				
				if (ddDAFloat == pBand->bpV1.m_daDockingArea && VARIANT_TRUE == pBand->bpV1.m_vbVisible && !pBand->m_bMiniWinShowLock)
				{
					if (NULL == pBand->GetFloat() && !bJustRecalc)
						pBand->DoFloat(TRUE);
				}
				
//				if (ddBTPopup == pBand->bpV1.m_btBands && VARIANT_FALSE == pBand->bpV1.m_vbVisible && pBand->m_pPopupWin && pBand->m_pPopupWin->IsWindow())
//				{
//					pBand->m_pPopupWin->DestroyWindow();
//					pBand->m_pPopupWin = NULL;
//				}

				// 
				// If the Band was floating and now the band is being docked we need to destroy the
				// MiniWin because the band is now in a dock area.
				//
				
				if (pBand->GetFloat() && (ddDAFloat != pBand->bpV1.m_daDockingArea || VARIANT_FALSE == pBand->bpV1.m_vbVisible))
				{
					//
					// CBand DoFloat sets the band as not visible so we need to retain the visible state
					// 

					VARIANT_BOOL vbPrevVisible = pBand->bpV1.m_vbVisible;

					//
					// Destroy the MiniWin
					//

					if  (!bJustRecalc)
						pBand->DoFloat(FALSE);

					//
					// It's already false so only set it if it was true
					//

					if (VARIANT_TRUE == vbPrevVisible)
						pBand->bpV1.m_vbVisible = vbPrevVisible;
				}
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
				return FALSE;
			}	
		}

		SIZE  sizeDock;
		BOOL  bHorz;
		for (nDock = 0; nDock < eNumOfDockAreas; nDock++)
		{
			try
			{
				bHorz = ddDATop == m_pDockObject[nDock]->m_daDockingArea || ddDABottom == m_pDockObject[nDock]->m_daDockingArea;
				if (0 != m_pDockObject[nDock]->GetBandCount())
				{
					if (bHorz)
					{
						sizeDock = m_pDockObject[nDock]->RecalcDock(sizeWin.cx, !bJustRecalc);
						sizeWin.cy -= sizeDock.cy;
					}
					else
					{
						sizeDock = m_pDockObject[nDock]->RecalcDock(sizeWin.cy, !bJustRecalc);
						sizeWin.cx -= sizeDock.cx;
					}
				}
				else
				{
					//
					// The dock area is empty
					//

					m_pDockObject[nDock]->m_rcDock.SetEmpty(); 
					
					//
					// This stuff is later used for band dragging rectBound calc
					//

					if (bHorz)
						m_pDockObject[nDock]->m_rcDock.right = sizeWin.cx;
					else
						m_pDockObject[nDock]->m_rcDock.bottom = sizeWin.cy;
				}
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
				return FALSE;
			}	
		}
		
		//
		// This was used for intelligent repainting right I don't have intelligent repainting.  Some else 
		// I need to work on.
		//

		nBandCount = pBands->GetBandCount();
		for (nBand = 0; nBand < nBandCount; ++nBand)
		{
			pBand = pBands->GetBand(nBand);
			if (pBand)
				pBand->m_vbCachedVisible = pBand->bpV1.m_vbVisible;
		}

		//
		// Resize all Popup and Floating Bands if their size changed.
		// 
		// Resize all bands with page style set to none, the other page styles can be resize freely
		//
		
		CMiniWin* pFloatWin;
		CRect	  rcNew;
		CRect	  rcWin;
		GetWindowRect(m_pBar->GetDockWindow(), &rcWin);
		for (nBand = 0; nBand < nBandCount; nBand++)
		{
			try
			{
				pBand = pBands->GetBand(nBand);
				assert(pBand);
				if (NULL == pBand)
					continue;

				if (VARIANT_TRUE == pBand->bpV1.m_vbVisible)
				{
					SIZE size;
					switch (pBand->bpV1.m_daDockingArea)
					{
					case ddDAPopup:
					case ddDAFloat:
						pFloatWin = pBand->GetFloat();
						if (pFloatWin && !pBand->m_bMiniWinShowLock)
						{
							//
							// Resize Floating Windows
							//

							if (ddCBSNone == pBand->bpV1.m_cbsChildStyle && !(ddBFSizer & pBand->bpV1.m_dwFlags))
							{
								//
								// No pages 
								//

								rcNew = pBand->GetOptimalFloatRect(TRUE);

								size = rcNew.Size();

								pFloatWin->GetWindowRect(rcWin);

								if (size.cx != rcWin.Width() || size.cy != rcWin.Height())
								{
									pFloatWin->SetWindowPos(NULL, 
															0, 
															0,
															size.cx, 
															size.cy,
															SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOREPOSITION);
								}
							}
							else
							{
								//
								// We have pages 
								//
								
								pBand->GetOptimalFloatRect(TRUE);
								if (pFloatWin->IsWindow())
									pFloatWin->InvalidateRect(NULL, FALSE);
								else
								{
									pFloatWin->RemoveFromMap();
									delete pFloatWin;
									pBand->m_pFloat = NULL;
								}
							}
							if (ddBFSizer & pBand->bpV1.m_dwFlags)
							{
								if (pFloatWin->IsWindow())
								{
									pFloatWin->GetClientRect(rcWin);

									size = rcWin.Size();
									size.cx = size.cx - pBand->BandEdge().cx;
									size.cy = size.cy - pBand->BandEdge().cy;

									pBand->SetCustomToolSizes(size);
								}
							}
						}
						else if (pBand->GetPopup() && pBand->GetPopup()->IsWindow())
						{
							pBand->GetPopup()->GetOptimalRect(rcWin);
							size = rcWin.Size();
							pBand->GetPopup()->SetWindowPos(NULL, 
															0, 
															0,
															size.cx, 
															size.cy,
															SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOMOVE);
							pBand->Refresh();
						}
						break;
					}
				}
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
				return FALSE;
			}	
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
// RegisterDragDrop
//

void CDockMgr::RegisterDragDrop(IDragDropManager* pDragDropManager)
{
	for (int nIndex = 0; nIndex < eNumOfDockAreas; nIndex++)
	{
		pDragDropManager->RegisterDragDrop((OLE_HANDLE)m_pDockObject[nIndex]->hWnd(), 
										   (LPUNKNOWN)&m_pDockObject[nIndex]->m_theDockDropTarget);
	}
}

//
// RevokeDragDrop
//

void CDockMgr::RevokeDragDrop(IDragDropManager* pDragDropManager)
{
	for (int nIndex = 0; nIndex < eNumOfDockAreas; nIndex++)
		pDragDropManager->RevokeDragDrop((OLE_HANDLE)m_pDockObject[nIndex]->hWnd());
}

//
// CalcNewDockPosition
// 

void CDockMgr::CalcNewDockPosition(POINT  pt, CDock* pDock, CBand* pBand) 
{
	assert(m_pBar);
	assert(pDock);
	assert(pBand);

	//
	// Get the Dock Area Window Rect
	//
	
	//
	// Do the offset stuff, 
	//  
	//		The docking offset is relative to the Dock Area's rectangle
	// the point is relative to the screen
	//

	CRect rcBand;
	int nHeight;
	int nWidth;
	if (GetGlobals().m_bFullDrag && ddDAFloat == m_pBand->bpV1.m_daDockingArea)
	{
		pBand->GetFloat()->GetWindowRect(rcBand); 
		nWidth = rcBand.Width();
		nHeight = rcBand.Height();
	}
	else
	{
		nWidth = m_pBand->m_rcDock.Width();
		nHeight = m_pBand->m_rcDock.Height();
	}

	if (pDock->m_daDockingArea != m_pBand->bpV1.m_daDockingArea)
	{
		m_pBar->m_vbCustomizeModified = VARIANT_TRUE;
		if (ddDAFloat == m_pBand->bpV1.m_daDockingArea)
			m_pBar->FireBandDock((Band*)m_pBand);
	}

	int nOffsetX = (nWidth * m_pBand->m_ptPercentOffset.x) / eGranularity;
	int nOffsetY = (nHeight * m_pBand->m_ptPercentOffset.y) / eGranularity;
	
	nOffsetX += CBand::eBevelBorder;
	nOffsetY += CBand::eBevelBorder;
	
	POINT ptOffset = pt;
	ScreenToClient(pDock->m_hWnd, &ptOffset);
	switch (pDock->m_daDockingArea)
	{
	case ddDATop:
	case ddDABottom:
		pBand->bpV1.m_nDockOffset = ptOffset.x - nOffsetX;
		break;

	case ddDALeft:
	case ddDARight:
		pBand->bpV1.m_nDockOffset = ptOffset.y - nOffsetY;
		break;
	}

	if (0 == pDock->GetBandCount() || 0 == pDock->m_nLineCount)
	{
		//
		// There are no Bands in this dock area yet, so this one is 
		// easy put it in the first dockline which is zero
		//
		
		pBand->bpV1.m_nDockLine = 0;
		if (pBand->bpV1.m_daDockingArea != pDock->m_daDockingArea)
		{
			pBand->bpV1.m_daDockingArea = pDock->m_daDockingArea;
			pBand->m_dwLastDocked = m_pBar->IncrementLastDocked();
			if (m_pBar->m_pDesigner)
				m_pBar->m_pDesigner->SetDirty();
		}
		return;
	}

	//
	// Find the correct dock line
	// 
	// pt is relative to the screen so leave it that way
	//

	try
	{

		pDock->GetDockLine(pBand, pt);
		pDock->RestoreBandsOffsets(pBand);

	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// StartDrag
//

void CDockMgr::StartDrag(CBand* pBand, const POINT& ptBand)
{
	try
	{
		m_pBand = pBand;
		m_daInit = m_pBand->bpV1.m_daDockingArea;
		MSG msg;
		while (::PeekMessage(&msg, NULL, WM_PAINT, WM_PAINT, PM_NOREMOVE))
		{
			if (!GetMessage(&msg, NULL, WM_PAINT, WM_PAINT))
				break;

			DispatchMessage(&msg);
		}
		m_bDragging = TRUE;
		
		for (int nIndex = 0; nIndex < eNumOfDockAreas; nIndex++)
			m_pDockObject[nIndex]->Cache(pBand);
	
		DockingAreaTypes daCurrent = m_pBand->bpV1.m_daDockingArea;
		m_pBand->bpV1.m_daDockingArea = ddDATop;
		m_sizeSqueezed = m_pBand->CalcSqueezedSize();
		m_pBand->bpV1.m_daDockingArea = daCurrent;

		int nWidth;
		int nHeight;
		POINT pt2 = ptBand;
		if (ddDAFloat == pBand->bpV1.m_daDockingArea)
		{
			try
			{
				CMiniWin* pFloat = pBand->GetFloat();
				if (pFloat)
				{
					CRect rcFloatWindow;
					pFloat->GetWindowRect(rcFloatWindow);
					
					nWidth = rcFloatWindow.Width();
					nHeight = rcFloatWindow.Height();

					CRect rcMiniWin;
					CMiniWin::AdjustWindowRectEx(rcMiniWin);
					pt2.x -= rcMiniWin.left - 1;
					pt2.y -= rcMiniWin.top - 1;
				}
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
			}
		}
		else
		{
			nWidth = pBand->m_rcDock.Width();
			nHeight = pBand->m_rcDock.Height();
		}

		SystemParametersInfo(SPI_GETDRAGFULLWINDOWS, 0, &GetGlobals().m_bFullDrag, 0);

		pBand->m_ptPercentOffset.x = (pt2.x * eGranularity) / nWidth;
		pBand->m_ptPercentOffset.y = (pt2.y * eGranularity) / nHeight;

		try
		{
			try
			{
				POINT pt = ptBand;
				GetCursorPos(&pt);
				for (int nIndex = 0; nIndex < eNumOfDockAreas; nIndex++)
					m_pDockObject[nIndex]->CacheBandsOffsets();
				TrackDrag(pt);
				m_pBar->FireBandMove((Band*)pBand);
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
			}
		}
		CATCH
		{
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// CalcDockPosition
//

BOOL CDockMgr::CalcDockPosition(const POINT& pt)
{
	CDock* pDock = HitTest(pt);
	
	//
	// Find the dock area and make sure this band can dock there
	//

	if (pDock && pDock->GetDockFlag() & m_pBand->bpV1.m_dwFlags)
	{
		try
		{
			CalcNewDockPosition(pt, pDock, m_pBand);
			if (GetGlobals().m_bFullDrag)
			{
				m_pBar->InternalRecalcLayout();
				Invalidate();
			}
			else
			{
				m_pBar->InternalRecalcLayout(TRUE);

				CRect rc(0, m_pBand->m_rcDock.Width(), 0, m_pBand->m_rcDock.Height());
				POINT pt2 = m_pBand->OffsetDragPoint(pt, rc);
				rc.Offset(pt2.x, pt2.y);
				CMiniWin::DrawDragRect(rc, eDockBorder, m_prcLast, m_nPrevDragSize);
				
				if (NULL == m_prcLast)
					m_prcLast = new CRect;
				
				if (m_prcLast)
				{
					*m_prcLast = rc;
					m_nPrevDragSize = eDockBorder;
				}
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
	return FALSE;
}

//
// UpdateFloatPosition
//

void CDockMgr::UpdateFloatPosition(const POINT& pt)
{
	try
	{
		POINT pt2 = m_pBand->OffsetDragPoint(pt);
		if (GetGlobals().m_bFullDrag)
		{
			CMiniWin* pFloat = m_pBand->GetFloat();
			assert(pFloat);
			if (pFloat)
			{
				m_pBand->AdjustFloatRect(pt2);
				pFloat->SetWindowPos(NULL,
									 pt2.x,
									 pt2.y,
									 0,
									 0,
									 SWP_NOZORDER|SWP_NOACTIVATE|SWP_NOSIZE);
			}
		}
		else
		{
			m_pBand->AdjustFloatRect(pt2);
			CMiniWin::DrawDragRect(m_pBand->bpV1.m_rcFloat, 
								   m_nCxFloatDrawBorder, 
								   m_prcLast, 
								   m_nPrevDragSize);
			if (NULL == m_prcLast)
				m_prcLast = new CRect;
			if (m_prcLast)
			{
				*m_prcLast = m_pBand->bpV1.m_rcFloat;
				m_nPrevDragSize = m_nCxFloatDrawBorder;
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
// DockToFloat
//

void CDockMgr::DockToFloat(const POINT& pt)
{
	try
	{
		try
		{
			if (!(ddBFFloat & m_pBand->bpV1.m_dwFlags))
			{
				if (!GetGlobals().m_bFullDrag)
				{
					if (m_prcLast)
					{
						CMiniWin::DrawDragRect(*m_prcLast, m_nPrevDragSize);
						delete m_prcLast;
						m_prcLast = NULL;
					}
				}
				return;
			}

			m_pBar->FireBandUndock((Band*)m_pBand);
		}
		CATCH
		{
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
		}

		m_pBand->m_daPrevDockingArea = m_pBand->bpV1.m_daDockingArea;
		m_pBand->m_nPrevDockOffset = m_pBand->bpV1.m_nDockOffset;
		m_pBand->m_nPrevDockLine = m_pBand->bpV1.m_nDockLine;		
		
		POINT pt2 = m_pBand->OffsetDragPoint(pt);
		m_pBand->AdjustFloatRect(pt2);
		if (GetGlobals().m_bFullDrag)
		{
			m_pBand->bpV1.m_daDockingArea = ddDAFloat;
			m_pBand->bpV1.m_nDockLine = 0;
			m_pBand->bpV1.m_nDockOffset = 0;
			m_pBar->RecalcLayout();
		}
		else
		{
			m_pBand->bpV1.m_daDockingArea = ddDAFloat;
			m_pBar->InternalRecalcLayout(TRUE);
			CMiniWin::DrawDragRect(m_pBand->bpV1.m_rcFloat, 
								   m_nCxFloatDrawBorder, 
								   m_prcLast, 
								   m_nPrevDragSize);
			if (NULL == m_prcLast)
				m_prcLast = new CRect;
			if (m_prcLast)
			{
				*m_prcLast = m_pBand->bpV1.m_rcFloat;
				m_nPrevDragSize = m_nCxFloatDrawBorder;
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
// Move
//

void CDockMgr::Move(const POINT& pt, BOOL bControl)
{
	if (!bControl && CalcDockPosition(pt))
		return;
	else if (ddDAFloat == m_pBand->bpV1.m_daDockingArea)
		UpdateFloatPosition(pt);
	else
		DockToFloat(pt);
}

//
// CancelLoop
//

void CDockMgr::CancelLoop()
{
	TRACE(1, "CancelLoop - ReleaseCapture\n");
	ReleaseCapture();
	m_bDragging = FALSE;
	m_pBand = NULL;
}

//
// TrackDrag
//

void CDockMgr::TrackDrag(const POINT& pt)
{
	static CRect rcPrevDrag;
	static int nPrevDragSize = 1;
	m_nCxFloatDrawBorder = GetSystemMetrics(SM_CXFRAME);
	m_nCyFloatDrawBorder = GetSystemMetrics(SM_CYFRAME);

	// Don't handle if capture already set
	if (GetCapture())
	{
		TRACE(1, _T("Don't handle if capture already set\n"))
		return;
	}

	if (m_pBar->m_pDesigner)
	{
		// Set capture
		CRect rc;
		m_pBand->GetBandRect(m_hWndCapture, rc);
	}
	else 
		m_hWndCapture = m_pBar->m_hWnd;

	CRect rcBase;
	HMONITOR hMon = MonitorFromRect(&rcBase, MONITOR_DEFAULTTONEAREST);
	MONITORINFO mi;
	mi.cbSize = sizeof(MONITORINFO);
	BOOL bResult = GetMonitorInfo(hMon, &mi);
	CRect rcWin;
	
	if (GetSystemMetrics(SM_CMONITORS) > 1)
	{
		rcWin.right = GetSystemMetrics(SM_CXVIRTUALSCREEN);
		rcWin.bottom = GetSystemMetrics(SM_CYVIRTUALSCREEN);
	}
	else
	{
		HMONITOR hMon = MonitorFromRect(&rcBase, MONITOR_DEFAULTTONEAREST);
		MONITORINFO mi;
		mi.cbSize = sizeof(MONITORINFO);
		BOOL bResult = GetMonitorInfo(hMon, &mi);
		rcWin = mi.rcWork;
	}

	rcWin.Inflate(-5, -5);

	SetCapture(m_hWndCapture);

	// Get messages until capture lost or cancelled/accepted
	BOOL bControl = 0x8000 & GetKeyState(VK_CONTROL);
	if (m_bDragging)
		Move(pt, bControl);

	MSG  msg;
	BOOL bProcessing = TRUE;
	DWORD nDragStartTick = GetTickCount();
	while (bProcessing && m_hWndCapture == GetCapture())
	{
		try
		{
			if (!::GetMessage(&msg, NULL, 0, 0))
				break;

			bControl = 0x8000 & GetKeyState(VK_CONTROL);
			if ((GetKeyState(VK_LBUTTON) & 0x8000) == 0)
			{
				// Lost BUTTONUP Message Recover
				bProcessing = FALSE;
				break;
			}

			switch (msg.message)
			{
			case WM_PAINT:
			case WM_NCPAINT:
				DispatchMessage(&msg);
				break;

			case WM_RBUTTONUP:
			case WM_RBUTTONDOWN:
				break;

			case WM_NCLBUTTONUP:
			case WM_LBUTTONUP:
				bProcessing = FALSE;
				break;

			case WM_NCMOUSEMOVE:
			case WM_MOUSEMOVE:
				try
				{
					if (m_bDragging)
					{
						if (PtInRect(&rcWin, msg.pt))
							Move(msg.pt, bControl);
					}
					else
						Stretch(msg.pt);
				}
				CATCH
				{
					bProcessing = FALSE;
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
				break;

			case WM_KEYDOWN:
				switch (msg.wParam)
				{
				case VK_ESCAPE:
					bProcessing = FALSE;
					break;

				case VK_CONTROL:
					bControl = TRUE;
					if (ddDAFloat != m_pBand->bpV1.m_daDockingArea)
						DockToFloat(msg.pt);
					break;
				}
				break;

			case WM_KEYUP:
				switch (msg.wParam)
				{
				case VK_CONTROL:
					if (m_bDragging && ddDAFloat == m_pBand->bpV1.m_daDockingArea)
						CalcDockPosition(msg.pt);
					if (!GetGlobals().m_bFullDrag)
						Move(msg.pt, bControl);
					break;
				}
				break;

			default:
				// Just dispatch rest of the messages
				TranslateMessage(&msg);
				DispatchMessage(&msg);
				break;
			}
		}
		CATCH
		{
			bProcessing = FALSE;
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
		}
	}
	
	try
	{
		if (m_bDragging)
			EndDrag(msg.pt, bControl);
		else
			EndResize();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// StartResize
//

void CDockMgr::StartResize(CBand* pBand, int nHitTest, POINT pt)
{
	m_pBand = pBand;
	MSG msg;
	while (::PeekMessage(&msg, NULL, WM_PAINT, WM_PAINT, PM_NOREMOVE))
	{
		if (!GetMessage(&msg, NULL, WM_PAINT, WM_PAINT))
			break;
		DispatchMessage(&msg);
	}
	SystemParametersInfo(SPI_GETDRAGFULLWINDOWS, 0, &GetGlobals().m_bFullDrag, 0);
	m_ptLast = pt;
	m_nHitTest = nHitTest;
	TrackDrag(pt);
}

//
// CalcSizeRect
//

static void CalcSizeRect(UINT nHit, CRect& rc, int nDeltaX, int nDeltaY)
{
	switch (nHit)
	{
	case HTLEFT:
	case HTTOPLEFT:
	case HTBOTTOMLEFT:
		rc.left += nDeltaX;
		break;
	
	case HTRIGHT:
	case HTTOPRIGHT:
	case HTBOTTOMRIGHT:
		rc.right += nDeltaX;
		break;
	}
	
	switch (nHit)
	{
	case HTTOP:
	case HTTOPLEFT:
	case HTTOPRIGHT:
		rc.top += nDeltaY;
		break;

	case HTBOTTOM:
	case HTBOTTOMLEFT:
	case HTBOTTOMRIGHT:
		rc.bottom += nDeltaY;
		break;
	}
}

//
// Stretch
//

void CDockMgr::Stretch(POINT pt)
{
	CMiniWin* pFloat = m_pBand->GetFloat();
	assert(pFloat);
	if (NULL == pFloat)
		return;

	CRect rcWin;
	pFloat->GetWindowRect(rcWin);

	POINT ptOffset;
	ptOffset.y = ptOffset.x = 0;
	switch(m_nHitTest)
	{
	case HTLEFT:
		ptOffset.x = pt.x - rcWin.left;
		break;

	case HTRIGHT:
		ptOffset.x = pt.x - rcWin.right;
		break;

	case HTTOP:
		ptOffset.y = pt.y - rcWin.top;
		break;
	
	case HTBOTTOM:
		ptOffset.y = pt.y - rcWin.bottom;
		break;

	case HTTOPLEFT:
		ptOffset.x = pt.x - rcWin.left;
		ptOffset.y = pt.y - rcWin.top;
		break;

	case HTTOPRIGHT:
		ptOffset.x = pt.x - rcWin.right;
		ptOffset.y = pt.y - rcWin.top;
		break;

	case HTBOTTOMLEFT:
		ptOffset.x = pt.x - rcWin.left;
		ptOffset.y = pt.y - rcWin.bottom;
		break;

	case HTBOTTOMRIGHT:
		ptOffset.x = pt.x - rcWin.right;
		ptOffset.y = pt.y - rcWin.bottom;
		break;
	}

	// Did we move?
	if (0 == ptOffset.x && 0 == ptOffset.y)
		return;
	m_bHasMoved = TRUE;

	CRect rcBound = rcWin;
	CalcSizeRect(m_nHitTest, rcBound, ptOffset.x, ptOffset.y);
	
	CRect rcResult;
	CRect rcAdjust;
	if (pFloat->FreeResize())
	{
		if (rcBound.Width() < 32)
		{
			switch (m_nHitTest)
			{
			case HTLEFT:
			case HTTOPLEFT:
			case HTBOTTOMLEFT:
				rcBound.left = rcBound.right - 32;
				break;
			default:
				rcBound.right = rcBound.left + 32;
				break;
			}
		}
		else if (rcBound.Height() < 32)
		{
			switch (m_nHitTest)
			{
			case HTTOP:
			case HTTOPLEFT:
			case HTTOPRIGHT:
				rcBound.top = rcBound.bottom - 32;
				break;
			default:
				rcBound.bottom = rcBound.top + 32;
				break;
			}
		}
		CMiniWin::AdjustWindowRectEx(rcAdjust, FALSE);
		rcBound.left -= rcAdjust.left;
		rcBound.top -= rcAdjust.top;
		rcBound.right -= rcAdjust.right;
		rcBound.bottom -= rcAdjust.bottom;

		if (ddCBSSlidingTabs == m_pBand->bpV1.m_cbsChildStyle)
			m_pBand->CalcLayout(rcBound, CBand::eLayoutFloat | CBand::eLayoutVert, rcResult, TRUE);
		else
			m_pBand->CalcLayout(rcBound, CBand::eLayoutFloat | CBand::eLayoutHorz, rcResult, TRUE);
		
		CMiniWin::AdjustWindowRectEx(rcBound, FALSE);
		
		rcResult = rcBound;
	}
	else
	{
		if (m_nHitTest == HTLEFT || m_nHitTest == HTRIGHT)
			m_pBand->CalcLayout(rcBound, CBand::eLayoutFloat|CBand::eLayoutHorz, rcResult, TRUE);
		else
			m_pBand->CalcLayout(rcBound, CBand::eLayoutFloat|CBand::eLayoutVert, rcResult, TRUE);

		rcAdjust = rcResult;

		CMiniWin::AdjustWindowRectEx(rcAdjust, FALSE);

		rcResult = rcBound;

		switch(m_nHitTest)
		{
		case HTLEFT:
			rcResult.left = rcResult.right - rcAdjust.Width();
			rcResult.bottom = rcResult.top + rcAdjust.Height();
			break;

		case HTTOP:
			rcResult.top = rcResult.bottom - rcAdjust.Height();
			rcResult.right = rcResult.left + rcAdjust.Width();
			break;

		case HTRIGHT:
			rcResult.right = rcResult.left + rcAdjust.Width();
			rcResult.bottom = rcResult.top + rcAdjust.Height();
			break;

		case HTBOTTOM:
			rcResult.bottom = rcResult.top + rcAdjust.Height();
			rcResult.right = rcResult.left + rcAdjust.Width();
			break;
		}
	}
	if (NULL == m_prcLast || *m_prcLast != rcResult)
	{
		TRACE(3, "rcWin != rcResult\n");
		if (GetGlobals().m_bFullDrag)
		{
			CRect rcAdjust;
			CMiniWin::AdjustWindowRectEx(rcAdjust, FALSE);
			rcAdjust.left = rcResult.left - rcAdjust.left;
			rcAdjust.top = rcResult.top - rcAdjust.top;
			rcAdjust.right = rcResult.right - rcAdjust.right;
			rcAdjust.bottom = rcResult.bottom - rcAdjust.bottom;

			SIZE size = rcAdjust.Size();
			size.cx = size.cx - m_pBand->BandEdge().cx;
			size.cy = size.cy - m_pBand->BandEdge().cy;
			m_pBand->SetCustomToolSizes(size);
			pFloat->DragWindow(rcResult);
		}
		else
			CMiniWin::DrawDragRect(rcResult, m_nCxFloatDrawBorder, m_prcLast, m_nPrevDragSize);

		if (NULL == m_prcLast)
			m_prcLast = new CRect;
		if (m_prcLast)
			*m_prcLast = rcResult;

		if (m_pBar->m_pDesigner)
			m_pBar->m_pDesigner->SetDirty();
		
		m_ptLast = pt;
	}
}

//
// EndResize
//

void CDockMgr::EndResize()
{
	if (m_prcLast)
	{
		if (!GetGlobals().m_bFullDrag)
		{
			CMiniWin::DrawDragRect(*m_prcLast, m_nPrevDragSize);
			CRect rcAdjust;
			CMiniWin::AdjustWindowRectEx(rcAdjust, FALSE);
			rcAdjust.left = m_prcLast->left - rcAdjust.left;
			rcAdjust.top = m_prcLast->top - rcAdjust.top;
			rcAdjust.right = m_prcLast->right - rcAdjust.right;
			rcAdjust.bottom = m_prcLast->bottom - rcAdjust.bottom;
			SIZE size = rcAdjust.Size();
			size.cx = size.cx - m_pBand->BandEdge().cx;
			size.cy = size.cy - m_pBand->BandEdge().cy;
			m_pBand->SetCustomToolSizes(size);
		}
		m_pBand->GetFloat()->DragWindow(*m_prcLast);
		m_pBand->bpV1.m_rcFloat = *m_prcLast;
		delete m_prcLast;
		m_prcLast = NULL;
	}
	if (m_bHasMoved)
		m_bHasMoved = FALSE;
	CancelLoop();
}

//
// EndDrag
//

void CDockMgr::EndDrag(const POINT& pt, BOOL bControl)
{
	try
	{
		if (!GetGlobals().m_bFullDrag)
		{
			if (m_prcLast)
				CMiniWin::DrawDragRect(*m_prcLast, m_nPrevDragSize);

			CDock* pDock = HitTest(pt);
			if (!bControl && pDock && pDock->GetDockFlag() & m_pBand->bpV1.m_dwFlags)
			{
				CalcNewDockPosition(pt, pDock, m_pBand);
				m_pBar->RecalcLayout();
			}
			else
			{
				POINT pt2 = m_pBand->OffsetDragPoint(pt);
				if (ddDAFloat == m_pBand->bpV1.m_daDockingArea)
				{
					CMiniWin* pFloat = m_pBand->GetFloat();
					if (pFloat)
					{
						pFloat->SetWindowPos(NULL,
											 pt2.x,
											 pt2.y,
											 0,
											 0,
											 SWP_NOZORDER|SWP_NOACTIVATE|SWP_NOSIZE);
					}
					else
					{
						m_pBand->AdjustFloatRect(pt2);
						m_pBand->bpV1.m_nDockLine = 0;
						m_pBand->bpV1.m_nDockOffset = 0;
						m_pBar->RecalcLayout();
					}
				}
				else
				{
					m_pBand->bpV1.m_daDockingArea = m_daInit;
					m_pBar->RecalcLayout();
				}
			}
			delete m_prcLast;
			m_prcLast = NULL;
			m_pPrevDockArea = NULL;
		}
		else
		{
			CDock* pDock = HitTest(pt);
			if (!bControl && pDock && pDock->GetDockFlag() & m_pBand->bpV1.m_dwFlags)
			{
				// Do nothing 
			}
			else if (ddDAFloat != m_pBand->bpV1.m_daDockingArea)
			{
				m_pBand->bpV1.m_daDockingArea = m_daInit;
				m_pBar->RecalcLayout();
			}
		}
		CancelLoop();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}	
}

//
// CreateDockAreasWnds
//

void CDockMgr::CreateDockAreasWnds(HWND hWndFrame)
{
	CRect rcWindow;
	BOOL bResult = m_pDockObject[ddDATop]->Create(hWndFrame, rcWindow);
	assert(bResult);
	bResult = m_pDockObject[ddDABottom]->Create(hWndFrame, rcWindow);
	assert(bResult);
	bResult = m_pDockObject[ddDALeft]->Create(hWndFrame, rcWindow);
	assert(bResult);
	bResult = m_pDockObject[ddDARight]->Create(hWndFrame, rcWindow);
	assert(bResult);
}

//
// DestroyDockAreasWnds
//

void CDockMgr::DestroyDockAreasWnds()
{
	for (int nIndex = 0; nIndex < eNumOfDockAreas; nIndex++)
	{
		if (IsWindow(m_pDockObject[nIndex]->m_hWnd))
		{
			m_pDockObject[nIndex]->m_aBands.RemoveAll();
			DestroyWindow(m_pDockObject[nIndex]->m_hWnd);
			m_pDockObject[nIndex]->m_hWnd = NULL;
		}
	}
}

//
// CDemoDlg
//

class CDemoDlg: public FDialog
{
public:
	CDemoDlg();
	virtual BOOL DialogProc(UINT message, WPARAM wParam, LPARAM lParam);
};

//
// CDemoDlg
//

CDemoDlg::CDemoDlg()
	: FDialog(IDD_DEMOSCR)
{
}

//
// DialogProc
//

BOOL CDemoDlg::DialogProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	switch(message)
	{
	case WM_INITDIALOG:
		CenterDialog(GetParent(m_hWnd));
		break;
	}
	return FDialog::DialogProc(message, wParam, lParam);
}

//
// DemoDlg
//

void DemoDlg()
{
	if (!g_bDemoDlgDisplayed)
	{
		g_bDemoDlgDisplayed = TRUE;
		CDemoDlg dlg;
		dlg.DoModal();
	}
}

