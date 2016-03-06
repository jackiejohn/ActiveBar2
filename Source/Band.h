#ifndef __CBand_H__
#define __CBand_H__

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "Interfaces.h"
#include "Bar.h"
#include "Tools.h"
#include "Tool.h"

class CChildBands;
class CPopupWin;
class CDockMgr;
class CMiniWin;

class CBand : public IBand,public ISupportErrorInfo,public IDDPerPropertyBrowsing,public ICategorizeProperties
{
public:  
	CBand();
	ULONG m_refCount;
	~CBand();

	void SetOwner(CBar* pBar, BOOL bSetRoot = FALSE);
	CBar*  m_pBar;
	void SetParent(CBand* pBand);
	CBand* m_pParentBand;
	CBand* m_pRootBand;

	int m_objectIndex;

	//{BEGIN INTERFACEDEFS}
	static void *objectDef;
	static CBand *CreateInstance(IUnknown *pUnkOuter);
	// IUnknown members
	STDMETHOD(QueryInterface)(REFIID riid, void **ppvObj);
	STDMETHOD_(ULONG, AddRef)(void);
	STDMETHOD_(ULONG, Release)(void);

	// IDispatch members

	STDMETHOD(GetTypeInfoCount)( UINT *pctinfo);
	STDMETHOD(GetTypeInfo)( UINT itinfo, LCID lcid, ITypeInfo **pptinfo);
	STDMETHOD(GetIDsOfNames)( REFIID riid, OLECHAR **rgszNames, UINT cNames, ULONG lcid, long *rgdispid);
	STDMETHOD(Invoke)( long dispidMember, REFIID riid, ULONG lcid, USHORT wFlags, DISPPARAMS *pdispparams, VARIANT *pvarResult, EXCEPINFO *pexcepinfo, UINT *puArgErr);
	// DEFS for IDispatch
	
	

	// IBand members

	STDMETHOD(get_Caption)(BSTR *retval);
	STDMETHOD(put_Caption)(BSTR val);
	STDMETHOD(get_Flags)(BandFlags *retval);
	STDMETHOD(put_Flags)(BandFlags val);
	STDMETHOD(get_DockingArea)(DockingAreaTypes *retval);
	STDMETHOD(put_DockingArea)(DockingAreaTypes val);
	STDMETHOD(get_ToolsHPadding)(long *retval);
	STDMETHOD(put_ToolsHPadding)(long val);
	STDMETHOD(get_ToolsVPadding)(long *retval);
	STDMETHOD(put_ToolsVPadding)(long val);
	STDMETHOD(get_ToolsHSpacing)(long *retval);
	STDMETHOD(put_ToolsHSpacing)(long val);
	STDMETHOD(get_ToolsVSpacing)(long *retval);
	STDMETHOD(put_ToolsVSpacing)(long val);
	STDMETHOD(get_MouseTracking)(TrackingStyles *retval);
	STDMETHOD(put_MouseTracking)(TrackingStyles val);
	STDMETHOD(get_Height)(long *retval);
	STDMETHOD(put_Height)(long val);
	STDMETHOD(get_Left)(long *retval);
	STDMETHOD(put_Left)(long val);
	STDMETHOD(get_Top)(long *retval);
	STDMETHOD(put_Top)(long val);
	STDMETHOD(get_Type)(BandTypes *retval);
	STDMETHOD(put_Type)(BandTypes val);
	STDMETHOD(get_Visible)(VARIANT_BOOL *retval);
	STDMETHOD(put_Visible)(VARIANT_BOOL val);
	STDMETHOD(get_Width)(long *retval);
	STDMETHOD(put_Width)(long val);
	STDMETHOD(get_WrapTools)(VARIANT_BOOL *retval);
	STDMETHOD(put_WrapTools)(VARIANT_BOOL val);
	STDMETHOD(get_ActiveTool)(Tool* *retval);
	STDMETHOD(putref_ActiveTool)(Tool* *retval);
	STDMETHOD(put_ActiveTool)(Tool* val);
	STDMETHOD(get_DockingOffset)(long *retval);
	STDMETHOD(put_DockingOffset)(long val);
	STDMETHOD(get_Tools)(Tools * *retval);
	STDMETHOD(get_DockLine)(short *retval);
	STDMETHOD(put_DockLine)(short val);
	STDMETHOD(Clone)( IBand **pBand);
	STDMETHOD(CopyTo)( IBand *pBand);
	STDMETHOD(DSGGetSize)( int dx, int dy, int bandType, int *w,  int *h);
	STDMETHOD(DSGDraw)( OLE_HANDLE hdc, int x, int y, int w, int h);
	STDMETHOD(DSGCalcDropIndex)( int x, int y,  int *dropIndex, int *direction);
	STDMETHOD(DSGInsertTool)( int index, ITool *pTool);
	STDMETHOD(DSGCalcHit)( int x, int y,  int *index);
	STDMETHOD(DSGDrawSelection)( OLE_HANDLE hdc, int x,  int y, int width,  int index);
	STDMETHOD(PopupMenu)( VARIANT flags,  VARIANT x,  VARIANT y);
	STDMETHOD(get_CreatedBy)(CreatedByTypes *retval);
	STDMETHOD(put_CreatedBy)(CreatedByTypes val);
	STDMETHOD(get_ChildBandStyle)(ChildBandStyles *retval);
	STDMETHOD(put_ChildBandStyle)(ChildBandStyles val);
	STDMETHOD(get_ChildBands)(ChildBands* *retval);
	STDMETHOD(DSGSetDropLoc)( OLE_HANDLE hdc, int x, int y, int w, int h, int dropIndex, int direction);
	STDMETHOD(Refresh)();
	STDMETHOD(DSGCalcHitEx)( int x, int y, int *pHit);
	STDMETHOD(get_IsDetached)(VARIANT_BOOL *retval);
	STDMETHOD(put_IsDetached)(VARIANT_BOOL val);
	STDMETHOD(get_GrabHandleStyle)(GrabHandleStyles *retval);
	STDMETHOD(put_GrabHandleStyle)(GrabHandleStyles val);
	STDMETHOD(get_Name)(BSTR *retval);
	STDMETHOD(put_Name)(BSTR val);
	STDMETHOD(PopupMenuEx)( int flags,  int x, int y, int left, int top, int right, int bottom, VARIANT_BOOL *bDoubleClicked);
	STDMETHOD(get_TagVariant)(VARIANT *retval);
	STDMETHOD(put_TagVariant)(VARIANT val);
	STDMETHOD(get_DockedHorzWidth)(long *retval);
	STDMETHOD(put_DockedHorzWidth)(long val);
	STDMETHOD(get_DockedHorzHeight)(long *retval);
	STDMETHOD(put_DockedHorzHeight)(long val);
	STDMETHOD(get_DockedVertWidth)(long *retval);
	STDMETHOD(put_DockedVertWidth)(long val);
	STDMETHOD(get_DockedVertHeight)(long *retval);
	STDMETHOD(put_DockedVertHeight)(long val);
	STDMETHOD(IsChild)( VARIANT_BOOL*IsChild);
	STDMETHOD(get_AutoSizeForms)(VARIANT_BOOL *retval);
	STDMETHOD(put_AutoSizeForms)(VARIANT_BOOL val);
	STDMETHOD(get_DisplayMoreToolsButton)(VARIANT_BOOL *retval);
	STDMETHOD(put_DisplayMoreToolsButton)(VARIANT_BOOL val);
	STDMETHOD(get_Picture)(LPPICTUREDISP *retval);
	STDMETHOD(put_Picture)(LPPICTUREDISP val);
	STDMETHOD(putref_Picture)(LPPICTUREDISP *val);
	STDMETHOD(get_PopupBannerBackgroundStyle)(PopupBannerBackgroundStyles *retval);
	STDMETHOD(put_PopupBannerBackgroundStyle)(PopupBannerBackgroundStyles val);
	STDMETHOD(get_PopupBannerBackgroundColor)(OLE_COLOR *retval);
	STDMETHOD(put_PopupBannerBackgroundColor)(OLE_COLOR val);
	STDMETHOD(DragDropExchange)( IUnknown*pStream,  VARIANT_BOOL vbSave);
	STDMETHOD(GetToolIndex)( VARIANT *tool,  short *index);
	STDMETHOD(get_DockedVertMinWidth)(short *retval);
	STDMETHOD(put_DockedVertMinWidth)(short val);
	STDMETHOD(get_DockedHorzMinWidth)(short *retval);
	STDMETHOD(put_DockedHorzMinWidth)(short val);
	STDMETHOD(get_ScrollingSpeed)(short *retval);
	STDMETHOD(put_ScrollingSpeed)(short val);
	// DEFS for IBand
	
	

	// ISupportErrorInfo members

	STDMETHOD(InterfaceSupportsErrorInfo)( REFIID riid);
	// DEFS for ISupportErrorInfo
	
	

	// IPerPropertyBrowsing members

	STDMETHOD(GetDisplayString)( DISPID dispID, BSTR __RPC_FAR *pBstr);
	STDMETHOD(MapPropertyToPage)( DISPID dispID, CLSID __RPC_FAR *pClsid);
	STDMETHOD(GetPredefinedStrings)(  DISPID dispID,  CALPOLESTR __RPC_FAR *pCaStringsOut,  CADWORD __RPC_FAR *pCaCookiesOut);
	STDMETHOD(GetPredefinedValue)( DISPID dispID, DWORD dwCookie, VARIANT __RPC_FAR *pVarOut);
	// DEFS for IPerPropertyBrowsing
	
	

	// IDDPerPropertyBrowsing members

	STDMETHOD(GetType)( DISPID dispID,   long *pnType);
	// DEFS for IDDPerPropertyBrowsing
	
	

	// ICategorizeProperties members

	STDMETHOD(MapPropertyToCategory)( DISPID dispid,  PROPCAT *ppropcat);
	STDMETHOD(GetCategoryName)( PROPCAT propcat,  LCID lcid, BSTR *pbstrName);
	// DEFS for ICategorizeProperties
	
	
	// Events 
	//{END INTERFACEDEFS}
	
#ifdef _DEBUG
	void Dump(DumpContext& dc);
	void DumpLayoutInfo(DumpContext& dc);
#endif

	enum DropDirections
	{
		eDropLeft,
		eDropRight,
		eDropTop,
		eDropBottom
	};

	enum LayoutFlagConstants
	{
		eLayoutHorz = 1,
		eLayoutVert = 2,
		eLayoutFloat = 4
	};

	enum GrabHandleConstants
	{
		eGrabHandleX = 4,
		eGrabHandleY = 2,
		eGrabHandleWidth = 3,
		eGrabTotalWidth = 7,
		eGrabTotalHeight = 8
	};

	enum ExpandButtonState
	{
		eGrayed,
		eContracted,
		eExpanded,
		eExpandedToContracted,
		eFormResize,
	};

	enum ClipMarkerStates
	{
		eHover,
		ePressed,
		eNormal
	};

	enum MiscConstants
	{
		eMdiOffset = 16,
		eDetachBarHeight = 9,
		eHToolPadding = 2,
		eVToolPadding = 2,
		eBevelBorder = 1,
		eBevelBorder2 = 2,
		eDropSignSize = 4,
		eGranularity = 128,
		eCloseButton = 10,
		eCloseButtonSpace = 2,
		eExpandButton = 10,
		eExpandButtonSpace = 2,
		eClipMarkerWidth = 4,
		eClipMarkerHeight = 8,
		eEmptyStatusBarWidth = 22,
		eEmptyStatusBarHeight = 17,
		eEmptyBarWidth = 22,
		eEmptyBarHeight = 22,
		eEmptyDockWidth = 50,
		eEmptyDockHeight = 50,
		eMenuMouseUpDelay = 300
	};

	//
	// Band Properties
	//

	// Variable Sized Properties
	CPictureHolder m_phPicture;
	VARIANT m_vTag;
	BSTR	m_bstrCaption;
	BSTR	m_bstrName;

	// Fixed Sized Properties
#pragma pack(push)
#pragma pack (1)
	struct BandPropV1
	{
		BandPropV1();

		GrabHandleStyles m_ghsGrabHandleStyle:4;
		TrackingStyles    m_tsMouseTracking:4;
		DockingAreaTypes  m_daDockingArea:4;
		ChildBandStyles	  m_cbsChildStyle:4;
		BandTypes		  m_btBands:8;
		DWORD			  m_dwFlags;
		int				  m_nDockOffset;
		int				  m_nToolsHPadding;
		int				  m_nToolsVPadding;
		int				  m_nToolsHSpacing;
		int				  m_nToolsVSpacing;
		short			  m_nDockLine;
		short			  m_nCreatedBy;
		VARIANT_BOOL	  m_vbDetached:1;
		VARIANT_BOOL	  m_vbVisible:1;
		VARIANT_BOOL	  m_vbWrapTools:1;
		VARIANT_BOOL	  m_vbDesignerCreated:1;
		VARIANT_BOOL	  m_vbAutoSizeForms:1;
		CRect			  m_rcDimension; 
		CRect			  m_rcFloat; 
		CRect			  m_rcHorzDockForm;
		CRect			  m_rcVertDockForm;
		VARIANT_BOOL	  m_vbDisplayMoreToolsButton:1;
		PopupBannerBackgroundStyles m_pbsPicture:4;
		OLE_COLOR		  m_ocPictureBackground;
		short			  m_nScrollingSpeed;
		short			  m_nDockedVertMinWidth;
		short			  m_nDockedHorzMinWidth;
	} bpV1;

	VARIANT_BOOL m_vbCachedVisible:1;

	BOOL m_bMiniWinShowLock:1;
	BOOL m_bVerticalMenu:1;
	BOOL m_bDrawAnimated:1;
	BOOL m_bPopupWinLock:1;
	BOOL m_bClipMarker:1;
	BOOL m_bCycleMark:1; 
	BOOL m_bExpandButtonVisible:1;
	BOOL m_bCloseButtonVisible:1; 
	BOOL m_bToolRemoved:1;
	BOOL m_bDontExpand:1;
	BOOL m_bAllTools:1;
	BOOL m_bStatusBarSizerPresent:1;
	BOOL m_bSizerResized:1;
#pragma pack(pop)

	COLORREF m_crPictureBackground;
	TypedArray<int> m_aLineHeight;
	TypedArray<int> m_aLineWidth;
	CChildBands* m_pChildBands;
	CRect   m_rcChildBandArea; // Cached in CalcHorzLayout of ChildBands
	CRect   m_rcClipMarker;
	short   m_nLastPopupExpandedArea;
	SIZE	m_sizeLargestTool;
	int		m_nToolMouseOver;
	int		m_nCurrentTool;
	short	m_nDockLineIndex;
	CRect   m_rcGrabHandle;

	HRESULT Exchange(IStream* pStream, VARIANT_BOOL vbSave);
	HRESULT ExchangeConfig(IStream* pStream, VARIANT_BOOL vbSave);
	HRESULT ExchangeMenuUsageData(IStream* pStream, VARIANT_BOOL vbSave);

	HRESULT ClearMenuUsageData();
	BOOL GetCustomToolSizes(VARIANT& vWidth, VARIANT& vHeight);
	BOOL SetCustomToolSizes(SIZE size);
	
	BOOL IsWrappable();
	BOOL IsMenuBand();
	BOOL IsGrabHandle();
	BOOL IsVertical();
	BOOL IsMoreTools();

	int Height();
	
	// Used for sorting Dock Lines and bands in RecalcDock 
	// when lines are equal and DockOffset are equal
	int m_nCollectionIndex; 
	int m_nCachedDockOffset;

	void ResetFirstInLine();
	void SetFirstInLine();

	BOOL GetOptimalRect(CRect& rcPopup);

	void CalcDropLoc(POINT pt, CBar::DropInfo* pDropInfo);

	void SetDropLoc(HDC			 hDC, 
				    const CRect& rcLocation,
				    int			 nDropIndex, 
				    int			 nDirection);
	int m_nDropLoc;
	int m_nDropLocDir;

	void OnCustomMouseDown(UINT nMsg, POINT pt);

	void ShutdownFloatWin();
	BOOL GetToolRect(int nIndex, CRect& rc);

	BOOL GetAccelerators(LPACCEL& pAccel, int& nCount);

	CMiniWin* m_pFloat;
	CTools*   m_pTools;

	// index to tool which has an open popup
	int m_nPopupIndex;
	int m_nAdjustForGrabBar;

	// Internal functions used for impl.
	void SetPopupIndex(int				   nNewIndex, 
					   TypedArray<CTool*>* pToolsToBeDropped = NULL, 
					   BOOL				   bSelectFirstItem = FALSE);
	
	BOOL GetBandRect(HWND& hWnd, CRect& rcBand);
	HWND GetDockHandle();

	void TrackSlidingTabScrollButton(const POINT& pt, int nHit);

	void OnChildBandChanged(short nNewChildBand);

	POINT OffsetDragPoint(const POINT& pt, const CRect& rc);
	POINT OffsetDragPoint(const POINT& pt);
	POINT m_ptPercentOffset;

	void OnCustomCleanUp();
	
	//
	// Used for opening sub popups
	//

	BOOL GetToolScreenRect(int nToolIndex, CRect& rcTool); 
	BOOL GetToolBandRect(int nToolIndex, CRect& rcTool, HWND& hWnd);
	int GetIndexOfTool(CTool* pTool);

	BOOL CalcPopupLayout(int& nWidth, int& nHeight, BOOL bCommit, BOOL bPopupMenu = FALSE);
	
	BOOL CalcLayout(const CRect& rcBound, 
					DWORD        dwLayoutFlag, 
					CRect&		 rcReturn, 
					BOOL		 bCommit);
	
	BOOL CalcLayoutEx(const CRect& rcBound,
					  DWORD		   dwLayoutFlag,
					  BOOL		   bLeftHandle,
					  BOOL		   bTopHandle,
					  CRect&	   rcReturn,
					  BOOL		   bCommit,
					  BOOL         bWrapFlag);

	BOOL CalcHorzLayout(HDC        hDC,
						const CRect& rcBound,
						BOOL       bCommit,
						const int& nTotalWidth,
						const int& nTotalHeight,
						BOOL       bWrapFlag,
						BOOL       bLeftHandle,
						BOOL       bTopHandle,
						CRect&     rcReturn,
						int&       nReturnLineCount,
						DWORD	   dwLayoutFlag);

	int  m_nFormControlCount;
	int  m_nHostControlCount;
	int* m_pnCustomControlIndexes;
	int	 m_nFirstTool;
	int	 m_nWidth; // cached width GetTextExtentPoint32

	void ParentWindowedTools(HWND hWndParent);
	void HideWindowedTools();
	void TabbedWindowedTools(BOOL bHide);

	BOOL InsideCalcHorzLayout(HDC    hDC,
							  const CRect& rcBound,
							  BOOL   bCommit,
							  int    nTotalWidth,
							  BOOL   bWrapFlag,
							  BOOL   bLeftHandle,
							  BOOL   bTopHandle,
							  CRect& rcReturn,
							  int&   nReturnLineCount,
							  DWORD  dwLayoutFlag);

	BOOL InsideDraw(HDC   hDC, 
				    CRect rcPaint, 
				    BOOL  bWrapFlag, 
				    BOOL  bVertical, 
				    BOOL  bVerticalPaint, 
				    BOOL  bLeftHandle, 
				    BOOL  bRightHandle,
					const POINT& ptPaintOffset,
				    BOOL  bDrawControlsOrForms = TRUE);
	// returns minimum squeezed size for a docked band
	SIZE CalcSqueezedSize(); 

	void Draw(HDC hDC, CRect rcPaint, BOOL bWrapFlag, const POINT& ptPaintOffset);

	BOOL DoFloat(BOOL bDoFloat);
	CMiniWin* GetFloat();
	CRect GetOptimalFloatRect(BOOL bCommit);

	CTool* HitTestTools(CBand*& pBandHitTest, POINT pt, int& nToolIndex);
	CTool* GetTool(POINT pt, int& nToolIndex);
	void ToolNotification(CTool* pTool, TOOLNF tnf);
	void ToolNotificationEx(CTool* pTool, int nToolIndex, TOOLNF tnf);
	
	void DrawAnimatedSlidingTabs(CFlickerFree& ffObj, const HDC& hDC, CRect rcBand);

	void IncrementFirstTool();
	void DecrementFirstTool();
	int& GetFirstTool();

	//
	// Menu Stuff
	//

	enum 
	{
		ENTERMENULOOP_KEY=0,
		ENTERMENULOOP_CLICK=1,
		ENTERMENULOOP_FIRSTITEM=2,
		ENTERMENULOOP_LASTITEM=3
	};

	CPopupWin* GetPopup();
	void QueryPopupWin(HWND& hWnd, POINT& ptReturn);
	void EnterMenuLoop(int nToolIndex, CTool* pTool, int nEntryMode, UINT nInitKey); 
	int CheckMenuKey(WORD nVirtKey);
	
	CPopupWin* m_pPopupWin;
	CPopupWin* m_pParentPopupWin;
	CPopupWin* m_pChildPopupWin;
	
	friend class CCustomizeListbox;
	friend class CChildBands;
	friend class CInsideWnd;
	friend class CCustomize;
	friend class CPopupWin;
	friend class CMiniWin;
	friend class CDock;
	friend class CTool;
	friend class CBar;

	BOOL AddSubBandsToArray(TypedArray<CBand*>* pArray, 
							CBand**			    ppParentOfProblemBand = NULL, 
							BOOL			    bAddIfVisible = FALSE);

	void AdjustChildBandsRect(CRect& rcChildBand, BOOL* pbVertical = NULL, BOOL* pbVerticalPaint = NULL, BOOL* pbLeftHandle = NULL, BOOL* pbTopHandle = NULL);
	void AdjustFloatRect(const POINT& pt);

	BOOL CheckPopupOpen();
	BOOL ChildBandChanging();

	STDMETHOD(DrawDropSign)(OLE_HANDLE hdcScr,int x,int y,int w,int h,int dropIndex,int direction);
	STDMETHOD(TrackPopupEx)(int nFlags, int x, int y, CRect* prcBound, BOOL* pDblClick = NULL);

	void GetActiveTools(CTools** ppTools);

	BOOL ContainsSubBandOfToolTree(CTool* pTool,CBand** ppParentOfProblemBand);

	int GetVisibleToolCount();
	int GetNextVisibleToolIndex(int nVisibleIndex);

	BOOL NeedsFullStretch();
	void DelayedFloatResize();

	// RecalcLayout
	void ResetToolPressedState();
	
	void InvalidateRect(CRect* prcUpdate, BOOL bBackground);

	CDock* m_pDock;
	SIZE   m_sizeDragOffset;

	int  m_nPopupShortCutStringOffset;

	// Used when a MiniWin is double clicked in the caption area 
	DockingAreaTypes m_daPrevDockingArea;
	short m_nPrevDockLine;
	int m_nPrevDockOffset;

	SIZE& BandEdge();

	// Used in RecalcDock
	CRect m_rcDock;
	int m_nTempDockOffset; 

	// To eliminate ToolNotification during Startup
	int m_nInBandOpen;
	
	DWORD m_dwLastDocked;
	SIZE  m_sizePage;
	CRect m_rcInsideBand;
	CRect m_rcSizer;
	void CustomEditTools(const POINT& ptScreen, const CTool* pTool);
	BOOL SetupEditToolPopup(const CTool* pToolIn);

	CTool* PopupToolSelected();
	void PopupToolSelected(CTool* pTool);

	BOOL HasFormsOrHostControls();

private:
	BOOL OnLButtonDown(UINT nFlags, POINT pt);
	BOOL OnLButtonUp(UINT nFlags, POINT pt);
	BOOL OnLButtonDblClk(UINT nFlags, POINT pt);
	BOOL OnRButtonDown(UINT nFlags, POINT pt);
	BOOL OnRButtonUp(UINT nFlags, POINT pt);
	BOOL OnMouseMove(UINT nFlags, POINT pt);

	void OnMouseMove(CTool* pTool, POINT pt);
	void OnLButtonDown(CTool* pTool, int nToolIndex, POINT pt);


	CFlickerFree m_ffOldBand; 
	SIZE   m_sizeBandEdge;
	SIZE   m_sizeEdgeOffset;
	CRect* m_prcTools;
	CRect* m_prcPopupExpandArea;
	CRect  m_rcCurrentPaint;
	CRect  m_rcCurrentPage;

	CRect  m_rcCloseButton;
	void DrawCloseButton(HDC hDC, BOOL bPressed = FALSE, BOOL bHover = FALSE);
	BOOL TrackCloseButton(const POINT& pt);

	ExpandButtonState m_eExpandButtonState;
	CRect m_rcExpandButton;
	void DrawExpandButton(HDC hDC, BOOL bPressed = FALSE, BOOL bHover = FALSE);
	BOOL TrackExpandButton(const POINT& pt);
	
	BOOL DrawFlatEdge(HDC hDC, const CRect& rcEdge, BOOL bVertical);
	void DrawGrabHandles(HDC hDC, BOOL bVertical, CRect rcPaint, BOOL& bLeftHandle, BOOL& bTopHandle);
	void DrawTool(HDC hDC, CTool* pTool, CRect& rcTool, BOOL bVertical, const POINT& ptPaintOffset, BOOL bDrawControlsOrForms = TRUE);
	BOOL DrawSizer(HDC hDC, const CRect& rcSizer);

public:
	void InvalidateToolRect(int nToolIndex);
	void InvalidateToolRect(CTool* pTool, const CRect& rcTool);
	void XPLookPopupWinInvalidateToolRect(HWND hWnd, CTool* pTool, const CRect& rcTool);

	CFlickerFree m_ffGradient;

	CDockMgr* m_pDockMgr;

	CTool* m_pToolMenuExpand;
	
	CTool* m_pPopupToolSelected;

	SIZE m_sizeMaxIcon;
#pragma pack(push)
#pragma pack (1)
	BOOL m_bChildBandChanging:1;
	BOOL m_bDesignTime:1;
	BOOL m_bLoading:1;
#pragma pack(pop)
};

//
// BandEdge 
//
// This is the size of the band's outer edge 
//

inline SIZE& CBand::BandEdge()
{
	return m_sizeBandEdge;
}

//
// OffsetDragPoint
//

inline POINT CBand::OffsetDragPoint(const POINT& pt, const CRect& rc)
{
	POINT pt2 = pt;
	pt2.x -= (rc.Width() * m_ptPercentOffset.x) / eGranularity;
	pt2.y -= (rc.Height() * m_ptPercentOffset.y) / eGranularity;
	return pt2;
}

//
// OffsetDragPoint
//

inline POINT CBand::OffsetDragPoint(const POINT& pt)
{
	return OffsetDragPoint(pt, bpV1.m_rcFloat);
}

//
// GetFloat
//

inline CMiniWin* CBand::GetFloat()
{
	return m_pFloat;
}

//
// GetPopup
//

inline CPopupWin* CBand::GetPopup()
{
	return m_pPopupWin;
}

//
// IsWrappable
//

inline BOOL CBand::IsWrappable()
{
	return ((m_pBar && m_pBar->m_bCustomization) || VARIANT_TRUE == bpV1.m_vbWrapTools || ddBTMenuBar == bpV1.m_btBands || ddBTChildMenuBar == bpV1.m_btBands);
}

//
// IsMenuBand
//

inline BOOL CBand::IsMenuBand()
{
	return (ddBTMenuBar == bpV1.m_btBands || ddBTChildMenuBar == bpV1.m_btBands);
}

//
// GetToolRect
//

inline BOOL CBand::GetToolRect(int nIndex, CRect& rc)
{
	if (NULL == m_prcTools)
		return FALSE;
	rc = m_prcTools[nIndex];
	return TRUE;
}

inline BOOL CBand::HasFormsOrHostControls()
{
	return (m_nFormControlCount > 0 || m_nHostControlCount > 0);
}

//
// NeedsFullStretch
//

inline BOOL CBand::NeedsFullStretch()
{
#ifdef AR20
	return (0 != (bpV1.m_dwFlags & ddBFStretch));
#else
	return (((ddBTMenuBar == bpV1.m_btBands || ddBTChildMenuBar == bpV1.m_btBands) && (ddDATop == bpV1.m_daDockingArea || ddDABottom == bpV1.m_daDockingArea)) || ddCBSNone != bpV1.m_cbsChildStyle || 0 != (bpV1.m_dwFlags & ddBFStretch));
#endif
}

//
// SetParent
//

inline void CBand::SetParent(CBand* pParentBand)
{
	m_pParentBand = pParentBand;
	if (m_pParentBand)
		m_pRootBand = m_pParentBand->m_pRootBand;
}

//
// DecrementFirstTool
//

inline void CBand::DecrementFirstTool()
{
	if (0 != m_nFirstTool)
		m_nFirstTool--;
}

//
// GetFirstTool
//

inline int& CBand::GetFirstTool()
{
	return m_nFirstTool;
}

//
// ChildBandChanging
//

inline BOOL CBand::ChildBandChanging()
{
	return m_bChildBandChanging;
}

inline int CBand::GetIndexOfTool(CTool* pTool)
{
	return m_pTools->GetToolIndex(pTool);
}

inline BOOL CBand::IsGrabHandle()
{
	if (NULL == m_pRootBand)
		return FALSE;
	return ((ddGSNone != m_pRootBand->bpV1.m_ghsGrabHandleStyle || ddBFClose & m_pRootBand->bpV1.m_dwFlags || ddBFExpand & m_pRootBand->bpV1.m_dwFlags) && NULL == m_pRootBand->m_pFloat);
}

inline BOOL CBand::IsVertical()
{
	if (NULL == m_pRootBand)
		return FALSE;
	BOOL bVertical = (ddDALeft == m_pRootBand->bpV1.m_daDockingArea || ddDARight == m_pRootBand->bpV1.m_daDockingArea || (ddDAFloat == m_pRootBand->bpV1.m_daDockingArea && ddCBSSlidingTabs == m_pRootBand->bpV1.m_cbsChildStyle));
	return bVertical;
}

inline BOOL CBand::IsMoreTools()
{
	return (ddBTNormal == bpV1.m_btBands && 
		    VARIANT_TRUE == bpV1.m_vbDisplayMoreToolsButton && 
		    !(m_pParentBand && ddCBSSlidingTabs == m_pParentBand->bpV1.m_cbsChildStyle) && 
			ddDAPopup != bpV1.m_daDockingArea);
}

inline CTool* CBand::PopupToolSelected()
{
	return m_pPopupToolSelected;
}

inline void CBand::PopupToolSelected(CTool* pTool)
{
	m_pPopupToolSelected = pTool;
}

extern ITypeInfo *GetObjectTypeInfo(LCID lcid,void *objectDef);

#endif
