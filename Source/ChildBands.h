#ifndef __CChildBands_H__
#define __CChildBands_H__
#include "Interfaces.h"
#include "PrivateInterfaces.h"
#include "Flicker.h"
#include "BTabs.h"

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

class CBand;

class CChildBands : public IChildBands,public ISupportErrorInfo,public ICategorizeProperties
	, public BTabs
{
public:
	CChildBands();
	ULONG m_refCount;
	~CChildBands();
	void Cleanup();

	//{BEGIN INTERFACEDEFS}
	static void *objectDef;
	static CChildBands *CreateInstance(IUnknown *pUnkOuter);
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
	
	

	// IChildBands members

	STDMETHOD(Item)( VARIANT *Index, Band **retval);
	STDMETHOD(Count)( short *retval);
	STDMETHOD(Add)( BSTR name, Band **retval);
	STDMETHOD(Remove)( VARIANT *Index);
	STDMETHOD(_NewEnum)( IUnknown **retval);
	STDMETHOD(get_ChildBandFont3D)(Font3DTypes *retval);
	STDMETHOD(put_ChildBandFont3D)(Font3DTypes val);
	STDMETHOD(get_PictureBackground)(LPPICTUREDISP *retval);
	STDMETHOD(put_PictureBackground)(LPPICTUREDISP val);
	STDMETHOD(putref_PictureBackground)(LPPICTUREDISP *val);
	STDMETHOD(get_PictureBackgroundMaskColor)(OLE_COLOR *retval);
	STDMETHOD(put_PictureBackgroundMaskColor)(OLE_COLOR val);
	STDMETHOD(get_PictureBackgroundStyle)(PictureBackgroundStyles *retval);
	STDMETHOD(put_PictureBackgroundStyle)(PictureBackgroundStyles val);
	STDMETHOD(get_PictureBackgroundUseMask)(VARIANT_BOOL *retval);
	STDMETHOD(put_PictureBackgroundUseMask)(VARIANT_BOOL val);
	STDMETHOD(get_BackColor)(OLE_COLOR *retval);
	STDMETHOD(put_BackColor)(OLE_COLOR val);
	STDMETHOD(get_ChildBandBackgroundStyle)(ChildBandBackgroundStyles *retval);
	STDMETHOD(put_ChildBandBackgroundStyle)(ChildBandBackgroundStyles val);
	STDMETHOD(get_ForeColor)(OLE_COLOR *retval);
	STDMETHOD(put_ForeColor)(OLE_COLOR val);
	STDMETHOD(get_ChildBandCaptionAlignment)(CaptionAlignmentTypes *retval);
	STDMETHOD(put_ChildBandCaptionAlignment)(CaptionAlignmentTypes val);
	STDMETHOD(get_HighLightColor)(OLE_COLOR *retval);
	STDMETHOD(put_HighLightColor)(OLE_COLOR val);
	STDMETHOD(get_ShadowColor)(OLE_COLOR *retval);
	STDMETHOD(put_ShadowColor)(OLE_COLOR val);
	STDMETHOD(get_ToolForeColor)(OLE_COLOR *retval);
	STDMETHOD(put_ToolForeColor)(OLE_COLOR val);
	STDMETHOD(get_CurrentChildBand)(Band * *retval);
	STDMETHOD(putref_CurrentChildBand)(Band * *retval);
	STDMETHOD(put_CurrentChildBand)(Band * val);
	STDMETHOD(RemoveAll)();
	STDMETHOD(get_GradientEndColor)(OLE_COLOR *retval);
	STDMETHOD(put_GradientEndColor)(OLE_COLOR val);
	// DEFS for IChildBands
	
	

	// ISupportErrorInfo members

	STDMETHOD(InterfaceSupportsErrorInfo)( REFIID riid);
	// DEFS for ISupportErrorInfo
	
	

	// ICategorizeProperties members

	STDMETHOD(MapPropertyToCategory)( DISPID dispid,  PROPCAT *ppropcat);
	STDMETHOD(GetCategoryName)( PROPCAT propcat,  LCID lcid, BSTR *pbstrName);
	// DEFS for ICategorizeProperties
	
	
	// Events 
	//{END INTERFACEDEFS}
#ifdef _DEBUG
	void Dump(DumpContext& dc);
#endif

	void SetOwner(const CBand* pBand);

	CBand* GetChildBand(int nIndex);
	int GetChildBandCount();
	int GetPosOfItem(VARIANT* pIndex);
	void SetCurrentChildBand(const short& nNewBand);
	CBand* GetCurrentChildBand();
	short& GetCurrentChildBandIndex();

	CPictureHolder m_phPictureBackground;
	CPictureHolder m_phPictureTransparentBackground;

#pragma pack(push)
#pragma pack (1)
	struct ChildBandsPropV1
	{
		ChildBandsPropV1();

		PictureBackgroundStyles m_PictureBackgroundStyle:4;
		ChildBandBackgroundStyles m_ChildBandBackgroundStyle:4;
		VARIANT_BOOL		   m_vbPictureBackgroundUseMask:1;
		OLE_COLOR			   m_ocBackColor;
		OLE_COLOR			   m_ocForeColor;
		OLE_COLOR			   m_ocToolForeColor;
		OLE_COLOR			   m_ocPictureBackgroundMaskColor;
		OLE_COLOR			   m_ocHighLightColor;
		OLE_COLOR			   m_ocShadowColor;
		OLE_COLOR			   m_ocGradientEndColor;
	} ppV1;

	BOOL m_bBackgroundDirty:1;
#pragma pack(pop)
	
	COLORREF m_crToolForeColor;
	COLORREF m_crHighLightColor;
	COLORREF m_crShadowColor;
	COLORREF m_crGradientEndColor;

	STDMETHOD(CopyTo)(CChildBands* pDest);

	HRESULT Exchange(IStream* pStream, VARIANT_BOOL bSave);
	HRESULT ExchangeConfig(IStream* pStream, VARIANT_BOOL vbSave);
	HRESULT DragDropExchange(IStream* pStream, VARIANT_BOOL vbSave);

	BOOL CalcHorzLayout(HDC    hDC,
						const CRect& rcBound,
						BOOL   bCommit,
						int    nTotalWidth,
						BOOL   bWrapFlag,
						BOOL   bLeftHandle,
						BOOL   bTopHandle,
						CRect& rcReturn,
						int&   nReturnLineCount,
						DWORD  dwLayoutFlag);

	BOOL Draw(HDC	hDC, 
		      CRect rcPaint, 
			  BOOL	bWrapFlag, 
			  BOOL	bVertical, 
			  BOOL	bVerticalPaint, 
			  BOOL	bLeftHandle, 
			  BOOL	bRightHandle,
			  const POINT& ptPaintOffset,
			  BOOL	bClipRgn = TRUE,
			  BOOL	bDrawControlsOrForms = TRUE);

	void ParentWindowedTools(HWND hWndParent);
	void HideWindowedTools();
	void ShowWindowedTools();
	void TabbedWindowedTools(BOOL bHide);

	// Tab Public Virtual functions
	virtual AREA BTGetArea();
	virtual int BTGetFontHeight();
	virtual int ScrollButtonHit(POINT pt);

	void DrawToolSlidingTabBackground(HDC		   hDC, 
									  const CRect& rcTotal, 
									  CRect		   rcTool, 
									  BOOL		   bVertical);

	BOOL LayoutSlidingTabs(BOOL bVertical, CRect rcArea);
	int m_nVertSlidingOffset;
	int m_nHorzSlidingOffset;

private:
	// Tab Virtual functions
	virtual HFONT BTGetFont();
	virtual int BTGetTabCount();
	virtual int BTGetVisibleTabCount();
	virtual BOOL BTTabVisible(const int& index);
	virtual BSTR& BTGetTabText(const int& nIndex);
	virtual CPictureHolder& BTGetTabPicture(const int& index);
	virtual int& BTGetTabWidth(const int& nIndex);
	virtual void BTSetTabWidth(const int& nIndex, const int& nWidth);
	virtual int BTCalcEndLen();

	int GetVisibleBands(TypedArray<CBand*>& aBands);
	BOOL CommitVisibleBands(TypedArray<CBand*>& aBands);
	
	BOOL MakeTransparentBackground();
	HRESULT TranslateColors();

	TypedArray<CBand*> m_aChildBands;
	TypedArray<CBand*> m_aVisibleChildBands;
	CFlickerFree m_ffPageBackground;
	CBand* m_pCurChildBand;
	CBand* m_pBand;
	friend class CBand;
};

//
// GetCurrentChildBand
//

inline CBand* CChildBands::GetCurrentChildBand()
{
	return m_pCurChildBand;
}

//
// GetCurrentChildBandIndex
//

inline short& CChildBands::GetCurrentChildBandIndex()
{
	return tpV1.m_nCurTab;
}

//
// GetChildBand
//

inline CBand* CChildBands::GetChildBand(int index)
{
	return m_aChildBands.GetAt(index);
}

//
// GetChildBandCount
//

inline int CChildBands::GetChildBandCount()
{
	return m_aChildBands.GetSize();
}

//
// BTGetTabCount
//

inline int CChildBands::BTGetTabCount()
{
	return m_aChildBands.GetSize();
}

extern ITypeInfo *GetObjectTypeInfo(LCID lcid,void *objectDef);

#endif
