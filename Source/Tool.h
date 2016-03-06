#ifndef __CTool_H__
#define __CTool_H__
#include "Interfaces.h"
#include "PrivateInterfaces.h"

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

class ActiveCombobox;
class ShortCutStore;
class ActiveEdit;
class CPopupWin;
class CTools;
class CBList;
class CBand;
class CTool;
class CBar;

//
// CToolShortCut
//

class CToolShortCut
{
public:
	CToolShortCut()
		: m_pBar(NULL),
		  m_pTool(NULL)
	{}
	
	~CToolShortCut();

	void Initialize(CBar* pBar, CTool* pTool);
	int GetCount();
	ShortCutStore* Get(int nIndex);
	HRESULT Add(ShortCutStore* pShortCutStore);
	HRESULT CopyTo(CToolShortCut& theToolShortCutIn);
	HRESULT Remove(ShortCutStore* pShortCutStore);
	HRESULT RemoveAll();
	HRESULT Exchange(IStream* pStream, VARIANT_BOOL vbSave);

private:
	CBar*  m_pBar;
	CTool* m_pTool;
	TypedArray<ShortCutStore*> m_aShortCutStores;
};

inline void CToolShortCut::Initialize(CBar* pBar, CTool* pTool)
{
	m_pBar = pBar;
	m_pTool = pTool;
}

inline int CToolShortCut::GetCount()
{
	return m_aShortCutStores.GetSize();
}

inline ShortCutStore* CToolShortCut::Get(int nIndex)
{
	return m_aShortCutStores.GetAt(nIndex);
}

//
// CTool
//

class CTool : public ITool,public ISupportErrorInfo,public ICustomHost,public IDDPerPropertyBrowsing,public ICategorizeProperties
{
public:
	enum
	{
		eEdit = 500,
		eCombobox = 501,
		eComboDropWidth = 9,
		eMenuCheckWidth = 8,
		eMenuCheckHeight = 7,
		eDefaultComboWidth = 64,
		eNumOfImageTypes = 4,
		eButtonDropDownWidth = 9,
		eButtonDropDownHeight = 9,
		eToolWithSubBandWidth = 9,
		eSeparatorThickness = 6,
		eMenuSeparatorThickness = 2,
		eDefualtImageSize = 16,
		ddTTMDIButtons = 100,
		ddTTChildSysMenu = 101,
		ddTTMenuExpandTool = 102,
		ddTTMoreTools = 103,
		ddTTAllTools = 104,
		eAutoRepeat = 2222, 
		eAutoRepeatTime = 500,
		eClipMarkerWidth = 8,
		eClipMarkerHeight = 5,
		eMoreToolsHeight = 11,
		eMoreToolsWidth = 11,
		eExpandedMoreToolsHeight = 20,
		eExpandedMoreToolsWidth = 20
	};

	enum MDIButtons
	{
		eNone,
		eMinimize,
		eRestore,
		eClose
	};

	enum HorzToolAlignment
	{
		eHTALeft,
		eHTACenter,
		eHTARight
	};

	enum VertToolAlignment
	{
		eVTAAbove,
		eVTACenter,
		eVTABelow
	};

	CTool();
	ULONG m_refCount;
	int m_objectIndex;
	~CTool();

	void SetBar(CBar* pBar);
	CBar*  m_pBar;
	void SetBand(CBand* pBand);
	CBand* m_pBand;

	void AddRefImages();
	void ReleaseImages();

	int m_nShortCutOffset;
	int m_nComboNameOffset;

	void SetParent(HWND hWndParent);
	HWND m_hWndActive; 
	HWND m_hWndParent; 
	HWND m_hWndPrevParent; 

	long m_nImageLoad;

	//{BEGIN INTERFACEDEFS}
	static void *objectDef;
	static CTool *CreateInstance(IUnknown *pUnkOuter);
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
	
	

	// ITool members

	STDMETHOD(get_ID)(long *retval);
	STDMETHOD(put_ID)(long val);
	STDMETHOD(get_Name)(BSTR *retval);
	STDMETHOD(put_Name)(BSTR val);
	STDMETHOD(get_HelpContextID)(long *retval);
	STDMETHOD(put_HelpContextID)(long val);
	STDMETHOD(get_TooltipText)(BSTR *retval);
	STDMETHOD(put_TooltipText)(BSTR val);
	STDMETHOD(get_Enabled)(VARIANT_BOOL *retval);
	STDMETHOD(put_Enabled)(VARIANT_BOOL val);
	STDMETHOD(get_Checked)(VARIANT_BOOL *retval);
	STDMETHOD(put_Checked)(VARIANT_BOOL val);
	STDMETHOD(get_Caption)(BSTR *retval);
	STDMETHOD(put_Caption)(BSTR val);
	STDMETHOD(get_Style)(ToolStyles *retval);
	STDMETHOD(put_Style)(ToolStyles val);
	STDMETHOD(get_CaptionPosition)(CaptionPositionTypes *retval);
	STDMETHOD(put_CaptionPosition)(CaptionPositionTypes val);
	STDMETHOD(get_ShortCuts)(VARIANT *retval);
	STDMETHOD(put_ShortCuts)(VARIANT val);
	STDMETHOD(get_Description)(BSTR *retval);
	STDMETHOD(put_Description)(BSTR val);
	STDMETHOD(get_ControlType)(ToolTypes *retval);
	STDMETHOD(put_ControlType)(ToolTypes val);
	STDMETHOD(get_Custom)(IDispatch * *retval);
	STDMETHOD(putref_Custom)(IDispatch * *retval);
	STDMETHOD(put_Custom)(IDispatch * val);
	STDMETHOD(get_Width)(long *retval);
	STDMETHOD(put_Width)(long val);
	STDMETHOD(get_Height)(long *retval);
	STDMETHOD(put_Height)(long val);
	STDMETHOD(get_Alignment)(ToolAlignmentTypes *retval);
	STDMETHOD(put_Alignment)(ToolAlignmentTypes val);
	STDMETHOD(DrawPict)( OLE_HANDLE hdc, int x, int y, int w, int h,  VARIANT_BOOL enabled);
	STDMETHOD(Clone)( ITool **pTool);
	STDMETHOD(CopyTo)( ITool **pDest);
	STDMETHOD(GetSize)( short bandType, int *cx, int *cy);
	STDMETHOD(ExtDraw)( OLE_HANDLE hdc, int x, int y, int w, int h, short bandType, VARIANT_BOOL fSel);
	STDMETHOD(get_Category)(BSTR *retval);
	STDMETHOD(put_Category)(BSTR val);
	STDMETHOD(SetPicture)( ImageTypes index, IPictureDisp *picture, VARIANT color);
	STDMETHOD(SetPictureMask)( ImageTypes index,  IPictureDisp *mask);
	STDMETHOD(get_Bitmap)( ImageTypes index,  OLE_HANDLE *hBitmap);
	STDMETHOD(put_Bitmap)( ImageTypes index,  OLE_HANDLE hBitmap);
	STDMETHOD(GetMaskColor)( OLE_COLOR *mcolor);
	STDMETHOD(get_SubBand)(BSTR *retval);
	STDMETHOD(put_SubBand)(BSTR val);
	STDMETHOD(get_Text)(BSTR *retval);
	STDMETHOD(put_Text)(BSTR val);
	STDMETHOD(get_CBListCount)(short *retval);
	STDMETHOD(get_CBWidth)(short *retval);
	STDMETHOD(put_CBWidth)(short val);
	STDMETHOD(get_CBListIndex)(short *retval);
	STDMETHOD(put_CBListIndex)(short val);
	STDMETHOD(get_CBStyle)(ComboStyles *retval);
	STDMETHOD(put_CBStyle)(ComboStyles val);
	STDMETHOD(get_CBList)(ComboList * *retval);
	STDMETHOD(put_CBList)(ComboList * val);
	STDMETHOD(GetPicture)( ImageTypes index, IPictureDisp **retval);
	STDMETHOD(get_Visible)(VARIANT_BOOL *retval);
	STDMETHOD(put_Visible)(VARIANT_BOOL val);
	STDMETHOD(SetFocus)();
	STDMETHOD(get_Default)(VARIANT_BOOL *retval);
	STDMETHOD(put_Default)(VARIANT_BOOL val);
	STDMETHOD(get_hWnd)(OLE_HANDLE *retval);
	STDMETHOD(put_hWnd)(OLE_HANDLE val);
	STDMETHOD(get_TagVariant)(VARIANT *retval);
	STDMETHOD(put_TagVariant)(VARIANT val);
	STDMETHOD(get_LabelStyle)(LabelStyles *retval);
	STDMETHOD(put_LabelStyle)(LabelStyles val);
	STDMETHOD(DragDropExchange)( IUnknown*pStream, VARIANT_BOOL vSave);
	STDMETHOD(CBClear)();
	STDMETHOD(_CBAddItem)(BSTR Item, long Index);
	STDMETHOD(CBRemoveItem)(long Index);
	STDMETHOD(get_List)(long Index, BSTR*Item);
	STDMETHOD(put_List)( long Index, BSTR Item);
	STDMETHOD(get_AutoRepeat)(VARIANT_BOOL *retval);
	STDMETHOD(put_AutoRepeat)(VARIANT_BOOL val);
	STDMETHOD(get_AutoRepeatInterval)(short *retval);
	STDMETHOD(put_AutoRepeatInterval)(short val);
	STDMETHOD(get_MaskBitmap)( ImageTypes index,  OLE_HANDLE *hBitmap);
	STDMETHOD(put_MaskBitmap)( ImageTypes index,  OLE_HANDLE hBitmap);
	STDMETHOD(get_MenuVisibility)(MenuVisibilityStyles *retval);
	STDMETHOD(put_MenuVisibility)(MenuVisibilityStyles val);
	STDMETHOD(put_CBItemData)( int Index,  long Data);
	STDMETHOD(get_CBItemData)( int Index,  long *Data);
	STDMETHOD(get_Left)(long *retval);
	STDMETHOD(get_Top)(long *retval);
	STDMETHOD(get_ImageWidth)(short *retval);
	STDMETHOD(put_ImageWidth)(short val);
	STDMETHOD(get_ImageHeight)(short *retval);
	STDMETHOD(put_ImageHeight)(short val);
	STDMETHOD(get_LabelBevel)(LabelBevelStyles *retval);
	STDMETHOD(put_LabelBevel)(LabelBevelStyles val);
	STDMETHOD(get_AutoSize)(ToolAutoSizeTypes *retval);
	STDMETHOD(put_AutoSize)(ToolAutoSizeTypes val);
	STDMETHOD(get_SelStart)(short *retval);
	STDMETHOD(put_SelStart)(short val);
	STDMETHOD(get_SelLength)(short *retval);
	STDMETHOD(put_SelLength)(short val);
	STDMETHOD(CBAddItem)(BSTR Item, VARIANT Index);
	STDMETHOD(get_CBLines)(short *retval);
	STDMETHOD(put_CBLines)(short val);
	STDMETHOD(get_CBNewIndex)(long *retval);
	// DEFS for ITool
	
	

	// ISupportErrorInfo members

	STDMETHOD(InterfaceSupportsErrorInfo)( REFIID riid);
	// DEFS for ISupportErrorInfo
	
	

	// ICustomHost members

	STDMETHOD(get_hWin)(OLE_HANDLE *retval);
	STDMETHOD(put_hWin)(OLE_HANDLE val);
	STDMETHOD(get_State)(short *retval);
	STDMETHOD(put_State)(short val);
	STDMETHOD(Refresh)();
	STDMETHOD(Close)();
	// DEFS for ICustomHost
	
	

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
#endif

	//
	// Properties
	//

	BSTR	m_bstrDescription;
	BSTR	m_bstrToolTipText;
	BSTR	m_bstrCategory;
	BSTR	m_bstrSubBand;
	BSTR	m_bstrCaption;
	BSTR	m_bstrName;
	BSTR	m_bstrText;
	LPTSTR  m_szCleanCaptionNormal;
	LPTSTR  m_szCleanCaptionPopup;
	VARIANT	m_vTag;

#pragma pack(push)
#pragma pack (1)
	struct ToolPropV1
	{
		ToolPropV1();

		VARIANT_BOOL m_vbEnabled:1;
		VARIANT_BOOL m_vbChecked:1;
		VARIANT_BOOL m_vbVisible:1;
		VARIANT_BOOL m_vbDefault:1;
		VARIANT_BOOL m_vbDesignerCreated:1;
		VARIANT_BOOL m_vbAutoRepeat:1;

		int   m_nWidth;
		int   m_nHeight;
		short m_nAutoRepeatInterval;
		ULONG m_nToolId;
		long  m_nHelpContextId;
		long  m_nUsage;
		long  m_nImageIds[4];

		ToolStyles       m_tsStyle:4;
		CaptionPositionTypes m_cpTools:4;
		ToolTypes          m_ttTools:10;
		ToolAlignmentTypes m_taTools:4;
		LabelStyles	       m_lsLabelStyle:4;
		MenuVisibilityStyles m_mvMenuVisibility:4;
		short m_nImageWidth;
		short m_nImageHeight;
		LabelBevelStyles   m_lsLabelBevel:4;
		short m_nSelStart;
		short m_nSelLength;
		ToolAutoSizeTypes m_asAutoSize;
		DWORD m_dwIdentity;
	} tpV1;

	BOOL m_bVerticalSeparator:1;
	BOOL m_bDropDownPressed:1;
	BOOL m_bGrayImage:1; // used for mousetracking
	BOOL m_bPressed:1;// used by band during mouse track, and tool during draw
	BOOL m_bClone:1;
	BOOL m_bShowTool:1;
	BOOL m_bAdjusted:1;
	BOOL m_bMFU:1;
	BOOL m_bCreatedInternally:1;
	BOOL m_bMoreTools:1;
#pragma pack(pop)

	SYSTEMTIME m_theSystemTime;
	IDispatch* m_pDispCustom;
	DWORD m_dwCustomFlag;
	SIZE m_sizeTool;

	CRect m_rcTemp; // used by band recalc algorithm to speed up recalc
	int   m_nTempMaxWidth; // running maxWidth used during horzLayout calc
	int   m_nBandLine; // running maxWidth used during horzLayout calc

	HRESULT Exchange (IStream* pStream, VARIANT_BOOL vbSave);
	HRESULT ExchangeConfig(IStream* pStream, VARIANT_BOOL vbSave);

	void TrackToolLButtonDown();
	BOOL TrackMDIButton(HDC hDC, HWND hWndBand, const CRect& rcButton, int nButton);
	BOOL HasSubBand();

	void Draw(HDC hDC, const CRect& rcPaint, BandTypes btType, BOOL bVertical, BOOL bVerticalPaint);
	void DrawEdit(HDC hDC, CRect rcPaint, BandTypes btType, BOOL bVertical);
	void DrawText(HDC hDC, const CRect& rcBound, BandTypes btType, BOOL bVertical, LPCTSTR szCaption, BSTR bstrCaption, BOOL bImage = FALSE, BOOL bChecked = FALSE, int nImageAdjustment = 0, BOOL bShortCut = FALSE);
	void DrawImage(HDC hDC, const CRect& rcBound);
	void DrawBorder(HDC hDC, const CRect& rcPaint, CRect rcImage, BandTypes btType, BOOL bVertical, BOOL bPressed);
	inline void DrawStatic(HDC hDC, const CRect& rcPaint, BandTypes btType, BOOL bVertical);
	inline void DrawButton(HDC hDC, const CRect& rcPaint, BandTypes btType, BOOL bVertical);
	void DrawChecked(HDC hDC, const CRect& rcPaint);
	void DrawCombobox(HDC hDC, CRect rcPaint, BandTypes btType, BOOL bVertical);
	void DrawSeparator(HDC hDC, const CRect& rcPaint, BandTypes btType, BOOL bVertical);
	void DrawMenuExpand(HDC hDC, CRect rcPaint);
	void DrawMDIButtons(HDC hDC, const CRect& rcPaint, BandTypes btType, BOOL bVertical);
	void DrawChildSysMenu(HDC hDC, const CRect& rcPaint, BandTypes btType, BOOL bVertical);
	void DrawDropDownImage(HDC hDC, const CRect& rcBound, BOOL bDouble = FALSE);
	void DrawButtonDropDown(HDC hDC, const CRect& rcPaint, BandTypes btType, BOOL bVertical);
	void DrawSubBandIndicator(HDC hDC, const CRect& rcBound);
	static BOOL DrawClipMarker(HDC hDC, CBar* pBar, CRect rcClipMarker, BOOL bVertical, BOOL bLargeIcons);
	void DrawMoreTools(HDC hDC, const CRect& rcPaint, BandTypes	btType, BOOL bVertical);
	void DrawAllTools(HDC hDC, const CRect& rcPaint, BandTypes	btType, BOOL bVertical);

	BOOL HandleEdit(int nFlags, POINT& pt);
	BOOL HandleCombobox(int nFlags, POINT& pt);

	void CalcImageTextPos(const CRect& rcPaint, 
						  BandTypes	   btType, 
						  BOOL		   bVertical, 
						  SIZE	       sizeTotal, 
						  SIZE		   sizeIcon, 
						  SIZE		   sizeText, 
						  BOOL		   bIcon, 
						  BOOL		   bText,
						  CRect&	   rcText,
						  CRect&	   rcImage);

	void PopupBand(HDC		    hDC, 
				   LPCTSTR      szCaption,
				   BandTypes    btType, 
				   const CRect& rcPaint,
				   const SIZE&  sizeTotal, 
				   const SIZE&  sizeChecked, 
				   const SIZE&  sizeIcon, 
				   const SIZE&  sizeText, 
				   BOOL		    bVertical, 
				   BOOL		    bChecked, 
				   BOOL		    bIcon, 
				   BOOL		    bText);

	void DrawCombo(HDC hDC, const CRect& rcPaint, BOOL bPressed, BOOL bDrawText = TRUE);
	void DrawXPCombo(HDC hDC, const CRect& rcBound, BOOL bPressed, BOOL bDrawText = TRUE);
	void DrawEdit(HDC hDC, const CRect& rcPaint);
	void DrawTextEdit(HDC hDC, const CRect& rcPaint);
	void DrawSlidingEdge(HDC hDC, const CRect& rcEdge, BOOL bSunken);
	inline void CleanCaption(TCHAR* szCaption, BOOL bPopup);

	SIZE CalcSize(HDC hDC, BandTypes btType, BOOL bVertical, SIZE* psizeIcon = NULL);

	void SetVisibleOnPaint(BOOL val);
	BOOL IsVisibleOnPaint();

	void OnLButtonDown();
	void OnLButtonDown(UINT nFlags, POINT pt);
	void OnLButtonUp();
	BOOL OnLButtonUp(UINT nFlags, POINT pt);

	void OnMenuLButtonDown(UINT nFlags, POINT pt);
	void OnMenuLButtonUp(BOOL bClicked = FALSE, BOOL bDelayed = FALSE);

	LPTSTR GetShortCutString(); // process \t and place into resStrBuffer
	void CalcImageSize();
	
	BOOL CheckForAlt(WCHAR& nKey);
	BOOL GetMDIButton(int& nButton, CRect& rcButton, POINT pt, BOOL bVertical);

	void HidehWnd();

	ActiveCombobox* GetCombo() const;
	CBList* GetComboList();

	void SetCombo(const ActiveCombobox* pActiveCombo);
	void SetEdit(const ActiveEdit* pActiveEdit);

	BOOL IsComboReadOnly();

	void GenerateMoreTools();

	CToolShortCut m_scTool;
	SIZE		  m_sizeImage; // This is set when put_picture is called or when serialization is done

	BOOL EnableMoreTools();

	void SubBand(CPopupWin* pSubBandWindow);
	CPopupWin* SubBand();

	BOOL m_bMDICloseEnabled;

	MDIButtons m_mdibActive;

private:

	HBITMAP XPShiftUpShadow(HBITMAP bmSource, HBITMAP bmSourceMask, BOOL bShadow);
	HBITMAP XPDisableImage(HBITMAP bmSource, HBITMAP bmSourceMask, BOOL bShadow);

	BOOL Customization();

	void CalcSizeEx(HDC       hDC,
					BandTypes btBand,
					BOOL      bVertical,
					SIZE&	  sizeTotal,
					SIZE&	  sizeChecked,
					SIZE&	  sizeIcon,
					SIZE&	  sizeSize,
					BOOL&     bChecked,
					BOOL&     bIcon,
					BOOL&     bText);

	HFONT GetMyFont(BandTypes btType, BOOL bVertical);
	WCHAR* GetShortCutStr();


	ActiveCombobox* m_pActiveCombo;
	ActiveEdit*		m_pActiveEdit;
	CTools*			m_pTools;
	CBList*			m_pComboList;
	short			m_nCustomWidth;
	DWORD			m_dwStyleOld;
	DWORD			m_dwExStyleOld;
	BOOL			m_bPaintVisible; // set by ::Draw method and used during hit calc	
	BOOL		    m_bFocus;
	int				m_nCachedHeight;
	int				m_nCachedWidth;

	CPopupWin* m_pSubBandWindow;

	friend CTools;
};

inline ActiveCombobox* CTool::GetCombo() const 
{
	return m_pActiveCombo;
}

inline void CTool::SetVisibleOnPaint(BOOL val) 
{
	if (m_bPaintVisible != val)
	{
		m_bPaintVisible=val;
		if (!m_bPaintVisible)
		{
			switch (tpV1.m_ttTools)
			{
			case ddTTForm:
			case ddTTControl:
				HidehWnd();
				break;
			}
		}
	}
};

inline void CTool::SubBand(CPopupWin* pSubBandWindow)
{
	m_pSubBandWindow = pSubBandWindow;
}

inline CPopupWin* CTool::SubBand()
{
	return m_pSubBandWindow;
}

inline BOOL CTool::IsVisibleOnPaint() 
{
	return m_bPaintVisible;
};

inline void CTool::OnLButtonUp()
{
	POINT pt = {0,0};
	OnLButtonUp(0, pt);
}

//
// SetCombo
//

inline void CTool::SetCombo(const ActiveCombobox* pActiveCombo)
{
	m_pActiveCombo = const_cast<ActiveCombobox*>(pActiveCombo);
}

//
// SetEdit
//

inline void CTool::SetEdit(const ActiveEdit* pActiveEdit)
{
	m_pActiveEdit = const_cast<ActiveEdit*>(pActiveEdit);
}

extern ITypeInfo *GetObjectTypeInfo(LCID lcid,void *objectDef);

#endif
