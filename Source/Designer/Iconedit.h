#ifndef __ICONEDIT_H__
#define __ICONEDIT_H__

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "debug.h"

class     CIconEdit;
class     Painter;
interface IPicture;
interface IBand;

//
// ColorWell 
//

class ColorWell : public FWnd
{
public:

	enum 
	{
		eNX = 4,
		eNY = 4,
		eBoxSize = 18,
		eNunOfColors = 16
	};

	ColorWell();

	BOOL CreateWin(HWND hWndParent, CRect rc, UINT nId);
	
	virtual LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);

	void Draw(HDC hDC, CRect& rcClient);
	void OnKeyDown(int nChar, LPARAM nKeyData);
	void OnLButtonDown(UINT nFlags, POINT pt);
	void OnRButtonDown(UINT nFlags, POINT pt);
	
	void InvalidateItem(int nIndex);
	int GetSelIndex(COLORREF crColor);
	void SetCurSel(int nNewIndex);
	COLORREF GetCurColor();

	BOOL m_bSupportSelection;
	int  m_nCurSel;
};

//
// Preview
//

class Preview : public FWnd
{
public:
	enum PREVIEWTYPES
	{
		ePreview = 0,
		eColor = 1,
		eScreenSize = 20,
		eCycleBitmapSize = 12
	};

	BOOL CreateWin(HWND hWndParent, PREVIEWTYPES eType, UINT nId); 

	// Extenally called to update thumbnail image
	void UpdateDibSection();
	void PaintDibSection(HDC hDCOff, const CRect& rc);

private:
	virtual LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	void Draw(HDC hDC, CRect& rcClient);
	void OnLButtonDblClk(UINT nFlags, POINT pt);

	PREVIEWTYPES m_nType; 
	CIconEdit*   m_pIconEdit;
	CRect        m_rcBackColor;
	CRect		 m_rcForeColor;
	CRect        m_rcScreen;
	CRect	     m_rcCycle;

	friend class CIconEdit;
};

//
// Painter
//

class Painter : public FWnd
{
public:
	enum CONSTANTS
	{
		eDefaultPainterSize = 16,
		eMoveBarHeight = 30,
		eShiftBarHeight = 30,
		eNumOfImages = 4,
		eMinBoxSize = 5
	};

	Painter();
	~Painter();

	enum IMAGEBORDERS
	{
		eBorder,
		eBorderFill,
		eFill
	};

	enum MOUSEBUTTONS
	{
		eLeftButton = 0,
		eRightButton = 1
	};

	BOOL CreateWin(HWND hWndParent); 
	void ClearCanvas();
	void SetSize(int cx, int cy);
	void Resize(SIZE sizeNew);
	void SetTool(UINT nTool);
	void SetBoxSize(int nWidthHeight);
	void RefreshImage();

#ifdef _DEBUG
	void Dump(DumpContext& dc);
#endif
	
private:
	virtual LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	void Draw(HDC hDC, CRect& rcClient);
	void OnButtonDown(MOUSEBUTTONS nButton);
	void PutPixel(int x, int y, COLORREF crColor);
	void PutPixelRaw(int x, int y, COLORREF crColor);
	
	// some action:

	void OnPenTrack(POINT pt, COLORREF crColor);
	void OnPenTrack(int x, int y, COLORREF crColor);
	void OnPatternTrack(POINT pt, COLORREF crColor);
	void OnPipetTrack(POINT pt, MOUSEBUTTONS nButton);

	static void CALLBACK PainterLineDDAProc(int X, int Y, LPARAM pData);
	static void CALLBACK CircleProc(int nX, int nY, LPARAM pData);
	static void CALLBACK CircleLineProc(int nX, int nY, int nLen, LPARAM pData);

	void StdInit(POINT  pt, COLORREF crColor);

	void OnLineInit(POINT pt, COLORREF crColor);
	void OnLineTrack(POINT pt, COLORREF crColor);
	void OnLineTrackPutPoint(POINT pt);
	void OnLineCleanup(POINT pt, COLORREF crColor, BOOL);
	
	void OnRectInit(POINT pt, IMAGEBORDERS nMode);
	void OnRectTrack(POINT pt, IMAGEBORDERS nMode);
	void OnRectCleanup(POINT pt, IMAGEBORDERS nMode, BOOL bCancelled);

	void OnCircleInit(POINT pt, IMAGEBORDERS nMode, COLORREF  crColor);
	void OnCircleTrack(POINT pt, IMAGEBORDERS nMode, COLORREF  crColor);
	void OnCircleCleanup(POINT pt, IMAGEBORDERS nMode, BOOL bCancelled, COLORREF crColor);

	void OnSelectInit(POINT pt);
	void OnSelectTrack(POINT pt);

	void OnMoverInit(POINT pt);
	void OnMoverTrack(POINT pt);
	void OnMoverCleanup(POINT pt, BOOL bCancelled);

	void OnHScroll(int nScrollCode, int nPos);
	void OnVScroll(int nScrollCode, int nPos);

	// Fill
	void FloodFill(int nInitX, int nInitY, COLORREF crColor);

	// Selection
	void PaintSelectionRect();
	void ShiftIcon(UINT msg);
	void SetImageIndex(int nNewIndex);

	void SetScrollInfo();
	void DoTimerScrolling();
private:
	CIconEdit* m_pIconEdit;
	BITMAPINFO m_bmInfo;
	COLORREF*  m_pcrMainData[eNumOfImages];
	COLORREF*  m_pcrBackup;
	COLORREF*  m_pcrImageData;
	COLORREF   m_crTrack;
	HCURSOR    m_hCursor;
	POINT      m_ptInit;
	CRect      m_rcSelection;
	CRect      m_rcPrevSelection;
	BOOL	   m_bRedraw;	
	UINT       m_nTool;
	SIZE       m_size;
	int		   m_nMaxVScrollPos;
	int		   m_nVScrollPos;
	int		   m_nVDisplayBoxes;
	int		   m_nMaxHScrollPos;
	int		   m_nHScrollPos;
	int		   m_nHDisplayBoxes;
	int        m_nBoxSize;

	friend class Preview;
	friend class CIconEdit;
};

//
// SetBoxSize
//

inline void Painter::SetBoxSize(int nWidthHeight) 
{
	m_nBoxSize = nWidthHeight;
};

//
// PutPixelRaw
//

inline void Painter::PutPixelRaw(int x, int y, COLORREF crColor)
{
	assert(m_pcrImageData);
	m_pcrImageData[m_size.cx * y + x] = crColor;
}

//
// SetImageIndex
//

inline void Painter::SetImageIndex(int nNewIndex)
{
	m_pcrImageData = m_pcrMainData[nNewIndex];
	assert(m_pcrImageData);
}

//
// OnSelectInit
//

inline void Painter::OnSelectInit(POINT pt)
{
	m_ptInit = pt;
	OnSelectTrack(pt);
}

//
// OnMoverInit
//

inline void Painter::OnMoverInit(POINT pt)
{
	m_rcPrevSelection = m_rcSelection;
	int nSize = m_size.cx * m_size.cy;
	m_pcrBackup = new COLORREF[nSize];
	if (NULL == m_pcrBackup)
		return;

	memcpy(m_pcrBackup, m_pcrImageData, nSize * sizeof(COLORREF));
	m_ptInit = pt;
}

//
// OnCircleInit
//

inline void Painter::OnCircleInit(POINT pt, IMAGEBORDERS eMode, COLORREF crColor)
{
	StdInit(pt, 0);
	OnCircleTrack(pt, eMode, crColor);
}

//
// UndoItem
//

class UndoItem
{
public:
	UndoItem(LPBYTE pData,
			 DWORD  dwSize,
			 int    cx,
			 int    cy) 
	{
		m_dwSize = dwSize;
		m_pData = new BYTE[dwSize];
		memcpy(m_pData, pData, dwSize);
		m_cx = cx;
		m_cy = cy;
	};

	~UndoItem() 
	{
		delete [] m_pData;
	};

	void GetSize(int& nCx, int& nCy);
	void GetData(LPBYTE& pData);

private:
	LPBYTE m_pData;
	DWORD  m_dwSize;
	int    m_cx;
	int    m_cy;
};

inline void UndoItem::GetSize(int& nCx, int& nCy)
{
	nCx = m_cx;
	nCy = m_cy;
}

inline void UndoItem::GetData(LPBYTE& pData)
{
	memcpy (pData, m_pData, m_dwSize);
}

//
// UndoManager
//

class UndoManager
{
public:
	UndoManager() {};
	~UndoManager();

	void Clear();
	void Push(int cx, int cy, LPBYTE pData, DWORD dwSize);
	BOOL PopQuery(int& nCx, int& nCy);
	void Pop(LPBYTE pDestData);
	BOOL Dirty();

private:
	FArray m_faUndoList;
};

//
// ~UndoManager
//

inline UndoManager::~UndoManager()
{
	Clear();
}

//
// Dirty
//

inline BOOL UndoManager::Dirty()
{
	return m_faUndoList.GetSize() == 0 ? FALSE : TRUE;
}


//
// CIconEdit
//

class CIconEdit : public FWnd
{
	enum CONSTANTS
	{
		eBevel = 4,
		eGridSpacing = 2,
		eNumOfStyles = 4,
		ePreviewWidth = 76,
		ePreviewHeight = 68,
		ePaintToolCount = 12,
		eStateAndShiftCount = 9,
		eColorPreviewHeight = 62,
		eFileDimensionHeight = 52,
		eImageMax = 200,
		eImageMin = 1,
		eMono = 1,
		e16Color = 2,
		e256Color = 3,
		e24Bit = 4
	};

public:
	CIconEdit();

	BOOL CreateWin(HWND hWndParent, HWND hWndSnapShotParent = NULL, int nOffsetButtons = 0);

	// Interface
	BOOL IsDirty(int nImageIndex);
	void SetDirty();
	void SetImageId(int nId, BOOL bUpdateView = TRUE);
	void SetImageIndex(int nId);
	BOOL SetBitmaps(HBITMAP hBitmap, HBITMAP hBitmapMask, COLORREF crMask); 
	BOOL GetBitmaps(HBITMAP& hBitmap, HBITMAP& hBitmapMask);
	SIZE& GetSize();

	void OnSize();
	void OnUndo();
	void Cut();
	BOOL Copy();
	void Paste();
	BOOL ClearPicture();

	HBITMAP GetBitmap();
	HBITMAP GetMaskBitmap();

private:

	void LoadPicture();
	void SavePicture();

	BOOL SaveBitmap(HANDLE hFile, int nBitCount);
	BOOL SaveIcon(HANDLE hFile);

	//
	// Snapshot stuff
	//

	void UndoPush();
	BOOL DrawSnapShotZoom(HWND hWnd, HDC hDC, int nX, int nY);
	static LRESULT CALLBACK SnapShotWndProc(HWND hWnd, UINT nMsg, WPARAM wParam, LPARAM lParam);

	void Init();
	BOOL CreateChildren();

	virtual LRESULT WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam);
	BOOL OnCommand(WPARAM wParam, LPARAM lParam);
	void OnNotify(int nControlId, LPNMHDR pNMhdr);
	void OnBandOpen(IBand* pBand);
	void OnSnapShot(HWND hWndParent = NULL);

	void UpdateColorPreview();
	BOOL UpdateSnapShotZoom(HWND hWnd, int nX, int nY);

	BOOL CreateImageBitmap(HBITMAP& hBmp, HBITMAP& hDib, BITMAPINFO& bmi, LPVOID& ppvBits, int nBitCount = -1);
	void PastePicture(IPicture* pPict, BOOL bCreateMask = FALSE);
	void PasteBitmap(HBITMAP hBitmap, BOOL bMask, int& nOption);
	void PasteIcon(HICON hIcon);

	void SetToolBarImageDimemsions();

	friend class Painter;
	friend class Preview;
	friend class CToolDoc;

	DDToolBar m_tbPaintTool;
	DDToolBar m_tbShiftBar;
	DDToolBar m_tbFileDimensions;
	ColorWell m_cwPallette;
	Preview   m_pvWin;
	Preview   m_pvColor;
	Painter   m_pntEdit;
	 
	UndoManager m_umManager;
	COLORREF    m_crBackground;
	COLORREF    m_crForeground;
	POINT       m_ptZCaretPos;
	BOOL        m_bDirty[eNumOfStyles];
	HWND        m_hWndSnap;
	HWND        m_hWndSnapShotParent;
	UINT		m_nPrevId;
	int         m_nCurrentTool;
	int		    m_nImageIndex;
	int         m_nZArea;
};

//
//
//

inline SIZE& CIconEdit::GetSize()
{
	return m_pntEdit.m_size;
}

//
// SetDirty
//

inline void CIconEdit::SetDirty()
{
	m_bDirty[m_nImageIndex] = TRUE;
}

//
// IsDirty
//

inline BOOL CIconEdit::IsDirty(int nImageIndex)
{
	if (nImageIndex < 0 || nImageIndex > 3)
		return FALSE;

	return m_bDirty[nImageIndex];
}

//
// UpdateColorPreview
//

inline void CIconEdit::UpdateColorPreview()
{
	m_pvColor.InvalidateRect(NULL, FALSE);
}

//
// SetImageIndex
//

inline void CIconEdit::SetImageIndex(int nId)
{
	m_nImageIndex = nId;
	m_pntEdit.SetImageIndex(m_nImageIndex);
}

//
// Cut
//

inline void CIconEdit::Cut()
{
	if (Copy())
	{
		m_bDirty[m_nImageIndex] = TRUE;
		ClearPicture();
	}
}

#endif
