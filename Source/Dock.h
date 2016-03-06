#ifndef __DOCK_H__
#define __DOCK_H__

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "Interfaces.h"
#include "Band.h"
#include "Designer\DragDropMgr.h"
#include "DropSource.h"

interface IDesignerObjects;
class CDockMgr;

typedef TypedArray<CBand*> ArrayOfBandPointers;

//
// CSplitter
//

class CSplitter
{
public:
	CSplitter(HWND hWndParent, CDock* pDock);
	virtual ~CSplitter();
	
	enum
	{
		eSplitterThickness = 4,
		eLeftLimit = 6,
		eRightLimit = 6,
		eTopLimit = 6,
		eBottomLimit = 6
	};

	virtual void SetSplitCursor(const BOOL& bSetSplit);

	void AddBetweenBandSplitters(CSplitter* pSplitter);
	void AddInsideBand(CBand* pBand);
	CBand* FindNextBand(CBand* pBandIn, BOOL bNext = TRUE);
	void SetToTheHighestLastDocked();
	void StartTracking();

protected:
	virtual void GetSplitterRect(CRect& rcSplitter);
	virtual void SetSplitterRect(const CRect& rcSplitter);
	virtual BOOL StopTracking(POINT pt);
	virtual void UpdateSplitterRect(POINT& pt);

	void DoTracking(POINT pt);
	void OnInvertTracker();
	
	TypedArray<CSplitter*> m_aBetweenBandSplitters;
	TypedArray<CBand*> m_aInsideBands;
	HCURSOR	m_hPrimaryCursor;
	CDock* m_pDock;
	CRect m_rcSplitter; // In Screen Coordinates
	HWND m_hWndParent;
	BOOL m_bTracking;
	
	friend class CSplitterMgr;
};

inline void CSplitter::AddInsideBand(CBand* pBand)
{
	assert(pBand);
	HRESULT hResult = m_aInsideBands.Add(pBand);
}

inline void CSplitter::AddBetweenBandSplitters(CSplitter* pSplitter)
{
	assert(pSplitter);
	HRESULT hResult = m_aBetweenBandSplitters.Add(pSplitter);
}

//
// CSplitter
//

class CBetweenBandsSplitter : public CSplitter
{
public:
	CBetweenBandsSplitter(CSplitter* pLineSplitter, HWND hWndParent, CDock* pDock, ArrayOfBandPointers& aBands);
	
	virtual void SetSplitCursor(const BOOL& bSetSplit);

	void AddOutsideBand(CBand* pBand);
	void AdjustSplitter();

private:
	virtual void GetSplitterRect(CRect& rcSplitter);
	virtual void SetSplitterRect(const CRect& rcSplitter);
	virtual BOOL StopTracking(POINT pt);
	virtual void UpdateSplitterRect(POINT& pt);

	ArrayOfBandPointers& m_aBands;
	CSplitter* m_pLineSplitter;
	CBand*     m_pBandOutSide;
	int        m_nLine;
};

inline void CBetweenBandsSplitter::AddOutsideBand(CBand* pBand)
{
	assert(pBand);
	m_pBandOutSide = pBand;
}

//
// CSplitterMgr
//

class CSplitterMgr
{
public:
	CSplitterMgr();
	~CSplitterMgr();
	void Cleanup();

	CSplitter* HitTest(POINT pt);
	BOOL Draw(HDC hDC);

	int AddSplitter(CSplitter* pSplitter, const CRect& rcSplitter);
	BOOL RemoveSplitter(CSplitter* pSplitter);

	void Wnd(FWnd* pWnd);

private:
	TypedArray<CSplitter*> m_aSplitters;
	CSplitter*			   m_pCurrentSplitter;	
	FWnd*				   m_pWnd;
};

inline int CSplitterMgr::AddSplitter(CSplitter* pSplitter, const CRect& rcSplitter)
{
	int nIndex;
	HRESULT hResult = m_aSplitters.Add(pSplitter, &nIndex);
	pSplitter->SetSplitterRect(rcSplitter);
	return nIndex;
}

inline void CSplitterMgr::Wnd(FWnd* pWnd)
{
	m_pWnd = pWnd;
}

//
// Structure used by CBar::InternalRecalcLayout to determine the size of the client Windows and the
// size of each dock area.
//

struct SIZEPARENTPARAMS
{
	HDWP  hDWP;      // handle for DeferWindowPos
	CRect rc;        // parent client rectangle (trim as appropriate)
	SIZE  sizeTotal; // total size on each side as layout proceeds
};

//
// RepositionWindow
//

inline void RepositionWindow(SIZEPARENTPARAMS* pSPPParent, HWND hWnd, const CRect& rcWindow)
{
	pSPPParent->hDWP = ::DeferWindowPos(pSPPParent->hDWP, 
									    hWnd, 
									    0, 
									    rcWindow.left, 
									    rcWindow.top, 
									    rcWindow.Width(),
									    rcWindow.Height(),
									    SWP_NOACTIVATE|SWP_NOZORDER);
	assert(pSPPParent->hDWP);

//	SetWindowPos(hWnd,NULL,rcWindow.left,rcWindow.top,rcWindow.Width(),rcWindow.Height(),SWP_NOZORDER|SWP_NOREDRAW|SWP_NOACTIVATE);
}

//
// CDockDragDrop
//

struct CDockDragDrop : public BandDragDrop
{
	virtual CBand* GetBand(POINT pt);
	virtual void OffsetPoint(CBand* pBand, POINT& pt);
};

//
// CLine
//

struct CLine
{
	CLine();

	int BandCount();
	int SizableBandCount();
	CBand* GetBand(int nIndex);
	int Find(CBand* pBand);
	CBand* FindNextSplitter(int nBand, BOOL bNext);
	ArrayOfBandPointers      m_aBands;
	CBand::ExpandButtonState m_eSplitterLineExpandedState;

#ifdef _DEBUG
	void Dump(DumpContext& dc);
#endif

#pragma pack(push)
#pragma pack (1)
	BOOL m_bSplitterLine:1;
	int	 m_nThickness;
	int	 m_nOffset;
#pragma pack(pop)
};

inline int CLine::BandCount()
{
	return m_aBands.GetSize();
}

inline int CLine::SizableBandCount()
{
	int nCount = 0;
	int nSize = m_aBands.GetSize();
	for (int nBand = 0; nBand < nSize; nBand++)
	{
		if (ddBFSizer & m_aBands.GetAt(nBand)->bpV1.m_dwFlags)
			nCount++;
	}
	return nCount;
}

inline CBand* CLine::GetBand(int nIndex)
{
	return m_aBands.GetAt(nIndex);
}

inline int CLine::Find(CBand* pBand)
{
	int nSize = m_aBands.GetSize();
	for (int nBand = 0; nBand < nSize; nBand++)
	{
		if (pBand == m_aBands.GetAt(nBand))
			return nBand;
	}
	return -1;
}

inline CBand* CLine::FindNextSplitter(int nBandIn, BOOL bNext)
{
	int nSize = m_aBands.GetSize();
	if (bNext)
	{
		for (int nBand = nBandIn + 1; nBand < nSize; nBand++)
		{
			if (ddBFSizer & m_aBands.GetAt(nBand)->bpV1.m_dwFlags)
				return m_aBands.GetAt(nBand);
		}
	}
	else
	{
		for (int nBand = nBandIn - 1; nBand >= 0; nBand--)
		{
			if (ddBFSizer & m_aBands.GetAt(nBand)->bpV1.m_dwFlags)
				return m_aBands.GetAt(nBand);
		}
	}
	return NULL;
}

typedef TypedArray<CLine*> ArrayOfLinesPointers;

//
// CLines
//

class CLines
{
public:
	~CLines();
	void Cleanup();

	int LineCount();
	CLine* GetLine(int nIndex);

#ifdef _DEBUG
	void Dump(LPTSTR szLabel, DumpContext& dc);
#endif

	ArrayOfLinesPointers m_aLines;
};

inline int CLines::LineCount()
{
	return m_aLines.GetSize();
}

inline CLine* CLines::GetLine(int nIndex)
{
	return m_aLines.GetAt(nIndex);
}

//
// CDock
//

class CDock : public FWnd
{
public:
	enum Constants
	{
		eLineGravity = 4,
	};

	CDock(CDockMgr* pDockMgr, const DockingAreaTypes& daArea);
	~CDock();
	
	CBar* Bar() const;
	
	BOOL Create(HWND hWndParent, CRect& rcWindow);
	
	CDockMgr& DockMgr()  const;
	
	DockingAreaTypes& GetDockingArea();
	
	DWORD GetDockFlag();
	
	CBand* HitTest(const POINT& pt);
	
	CRect& DockRect();

	BOOL MakeLineContracted(int nLine);

	BOOL CacheBandsOffsets();
	BOOL RestoreBandsOffsets(CBand* pBandSkip);

	BOOL AdjustSizableBands(CBand* pBand, int nDiff, BOOL bHorz);

	void ResizeDockableForms();

private:
	virtual LRESULT WindowProc(UINT nMessage, WPARAM wParam, LPARAM lParam);

	CRect& ClientRect();

	void CheckForBetweenBandSplitter(CBand* pBandPrev, CBand* pBand);
	
	int GetBandCount();
	SIZE GetSize();
	void GetDockLine(CBand* pBand, const POINT& pt);
	void GetDockLine2(CBand* pBand, const POINT& ptBand);
	int m_nCacheSize;
	void Cache(CBand* pRemoveBand);

	void OnSizeParent(SIZEPARENTPARAMS* pSPPParent);
	SIZE RecalcDock(int nMaxSize, BOOL bCommit); 
	void Reset();

	void OnCommand(int nNotifyCode, int nId, HWND hWndControl);

	ArrayOfBandPointers m_aBands;
	int m_nLineCount;
	int m_nCacheLineCount;

	CBandDropTarget  m_theDockDropTarget;
	CDockDragDrop    m_theBandDragDrop;
	CSplitterMgr     m_theSplitters;
	DockingAreaTypes m_daDockingArea;
//	CAccessible*     m_pAccessible;
	CDockMgr*		 m_pDockMgr;
	CLines*			 m_pLines;
	CLines*			 m_pCacheLines;
	CBand*		     m_pDesignerBand;
	CRect			 m_rcClientPosition;
	CRect		     m_rcDock;
	CRect			 m_rcDockableFormsSize;
	CBar*		     m_pBar;
	
	static DWORD m_dwDockFlags[4];
	friend class CSplitter;
	friend class CDockMgr;
	friend class CBand;
};

inline CBar* CDock::Bar() const 
{
	return m_pBar;
}

inline CDockMgr& CDock::DockMgr() const
{
	return *m_pDockMgr;
}

inline int CDock::GetBandCount() 
{
	return m_aBands.GetSize();
};

inline SIZE CDock::GetSize() 
{
	SIZE s = {m_rcDock.Width(), m_rcDock.Height()}; 
	return s;
};

inline DWORD CDock::GetDockFlag()
{
	return m_dwDockFlags[m_daDockingArea];
}

inline DockingAreaTypes& CDock::GetDockingArea()
{
	return m_daDockingArea;
}

inline CRect& CDock::DockRect()
{
	return m_rcDock;
}

inline CRect& CDock::ClientRect()
{
	return m_rcClientPosition;
}

//
// CDockMgr
//

class CDockMgr
{
public:
	enum Constants
	{
		eNumOfDockAreas = 4,
		eBandDockGravity = 8,
		eDockingGravityDx = 14,
		eDockingGravityDy = 14,
		eGranularity = 128,
		eDockBorder = 1
	};
	
	CDockMgr(CBar* pBar);
	~CDockMgr();

	void CreateDockAreasWnds(HWND hWndFrame);
	
	void DestroyDockAreasWnds();

	const CRect& GetDockRect(DockingAreaTypes daDockingArea);
	
	CDock* HitTest(const POINT& pt);

	const HWND hWnd(DockingAreaTypes daDockingArea);

	void Invalidate(BOOL bBackground = FALSE);
	void Invalidate(DockingAreaTypes daDockingArea, CBand* pBand = NULL);

	BOOL RecalcLayout(SIZE sizeWin, BOOL bJustRecalc = FALSE);

	BOOL RegisterOleDragDrop();
	BOOL RevokeOleDragDrop();

	BOOL RegisterCustomDragDrop();
	BOOL RevokeCustomDragDrop();

	void RegisterDragDrop(IDragDropManager* m_pDragDropManager);
	void RevokeDragDrop(IDragDropManager* m_pDragDropManager);

	void StartDrag(CBand* pBand, const POINT& pt);
	void StartResize(CBand* pBand, int nHitTest, POINT pt);
	BOOL& Dragging();

	void ResizeDockableForms();

	void CleanupLines();

private:
	CBar* Bar() const;
	BOOL CalcDockPosition(const POINT& pt);
	void CalcNewDockPosition(POINT pt, CDock* pDock, CBand* pBand);
	void CancelLoop();
	void DockToFloat(const POINT& pt);
	void EndDrag(const POINT& pt, BOOL bControl);
	void EndResize();
	void Move(const POINT& pt, BOOL bControl);
	void Stretch(POINT pt);
	void TrackDrag(const POINT& pt);
	void UpdateFloatPosition(const POINT& pt);
	HWND GetTempWin();

	DockingAreaTypes m_daInit;
	CDock* m_pDockObject[eNumOfDockAreas]; 
	CDock* m_pPrevDockArea; 
	CRect* m_prcLast;
	CBand* m_pBand;
	CBar*  m_pBar;
	POINT  m_ptLast;
	BOOL   m_bDragging;
	BOOL   m_bHasMoved;
	SIZE   m_sizeSqueezed;
	HWND   m_hWndCapture;
	HWND   m_hWndTemp;
	int    m_nCxFloatDrawBorder;
	int    m_nCyFloatDrawBorder;
	int    m_nPrevDragSize;
	int    m_nHitTest;

	friend CDock;
};

inline BOOL& CDockMgr::Dragging()
{
	return m_bDragging;
}

inline const CRect& CDockMgr::GetDockRect(DockingAreaTypes daDockingArea)
{
	return m_pDockObject[daDockingArea]->m_rcDock;
}

inline const HWND CDockMgr::hWnd(DockingAreaTypes daDockingArea)
{
	return m_pDockObject[daDockingArea]->hWnd();
}

inline CBar* CDockMgr::Bar() const 
{
	return m_pBar;
}

inline void CDockMgr::Invalidate(BOOL bBackground)
{
	for (int nIndex = 0; nIndex < eNumOfDockAreas; nIndex++)
	{
		if (m_pDockObject[nIndex]->IsWindow())
			m_pDockObject[nIndex]->InvalidateRect(NULL, bBackground); 
	}
}

inline BOOL CDockMgr::RegisterOleDragDrop()
{
	HRESULT nResult;
	for (int nIndex = 0; nIndex < eNumOfDockAreas; nIndex++)
	{
		nResult = ::RegisterDragDrop(m_pDockObject[nIndex]->hWnd(), &m_pDockObject[nIndex]->m_theDockDropTarget);
		assert(S_OK == nResult);
	}
}

inline BOOL CDockMgr::RegisterCustomDragDrop()
{
	HRESULT hResult;
	for (int nIndex = 0; nIndex < eNumOfDockAreas; nIndex++)
	{
		hResult = GetGlobals().m_pDragDropMgr->RegisterDragDrop((OLE_HANDLE)m_pDockObject[nIndex]->hWnd(), 
																&m_pDockObject[nIndex]->m_theDockDropTarget);
		assert(S_OK == hResult);
	}
}

inline BOOL CDockMgr::RevokeOleDragDrop()
{
	HRESULT hResult;
	for (int nIndex = 0; nIndex < eNumOfDockAreas; nIndex++)
	{
		hResult = ::RevokeDragDrop(m_pDockObject[nIndex]->hWnd());
		assert(S_OK == hResult);
	}
}

inline BOOL CDockMgr::RevokeCustomDragDrop()
{
	HRESULT hResult;
	for (int nIndex = 0; nIndex < eNumOfDockAreas; nIndex++)
	{
		hResult = GetGlobals().m_pDragDropMgr->RevokeDragDrop((OLE_HANDLE)m_pDockObject[nIndex]->hWnd());
		assert(S_OK == hResult);
	}
}

inline void CDockMgr::CleanupLines()
{
	for (int nIndex = 0; nIndex < eNumOfDockAreas; nIndex++)
	{
		if (m_pDockObject[nIndex] && m_pDockObject[nIndex]->m_pLines)
		{
			m_pDockObject[nIndex]->m_pLines->Cleanup();
			m_pDockObject[nIndex]->m_nLineCount = 0;
		}
	}
}
//
// DemoDlg
//

void DemoDlg();

#endif