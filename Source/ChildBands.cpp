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
#include "ipserver.h"
#include <stddef.h>       // for offsetof()
#include "Debug.h"
#include "Errors.h"
#include "Support.h"
#include "Resource.h"
#include "MiniWin.h"
#include "Dib.h"
#include "Band.h"
#include "ChildBands.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

//{OLE VERBS}
//{OLE VERBS}
//{OBJECT CREATEFN}
IUnknown *CreateFN_CChildBands(IUnknown *pUnkOuter)
{
	return (IUnknown *)(IChildBands *)new CChildBands();
}

DEFINE_AUTOMATIONOBJECT(&CLSID_ChildBands,ChildBands,_T("ChildBands"),CreateFN_CChildBands,2,0,&IID_IChildBands,_T(""));
void *CChildBands::objectDef=&ChildBandsObject;
CChildBands *CChildBands::CreateInstance(IUnknown *pUnkOuter)
{
	return new CChildBands();
}
//{OBJECT CREATEFN}
CChildBands::CChildBands()
	: m_refCount(1)
	, m_pBand(NULL),
	  m_pCurChildBand(NULL)
{
	InterlockedIncrement(&g_cLocks);
// {BEGIN INIT}
// {END INIT}
	TranslateColors();
	m_nHorzSlidingOffset = m_nVertSlidingOffset = 0;
	m_bBackgroundDirty = TRUE;
}

CChildBands::~CChildBands()
{
	InterlockedDecrement(&g_cLocks);
// {BEGIN CLEANUP}
// {END CLEANUP}
	Cleanup();
}
#ifdef _DEBUG
void CChildBands::Dump(DumpContext& dc)
{
	int nElem = m_aChildBands.GetSize();
	for (int cnt=0; cnt < nElem; ++cnt)
		m_aChildBands.GetAt(cnt)->Dump(dc);
}
#endif

CChildBands::ChildBandsPropV1::ChildBandsPropV1()
{
	m_ocBackColor = 0x80000000 + COLOR_APPWORKSPACE;
	m_ocForeColor = 0x80000000 + COLOR_WINDOWTEXT;
	m_ocToolForeColor = 0x80000000 + COLOR_WINDOWTEXT;
	m_ocHighLightColor = 0x80000000 + COLOR_3DLIGHT;
	m_ocShadowColor = 0x80000000 + COLOR_3DDKSHADOW;
	m_ocPictureBackgroundMaskColor = 0x80000000 + COLOR_3DFACE;
	m_ocGradientEndColor = 0x80000000;
	m_ChildBandBackgroundStyle = ddBSNormal;
	m_PictureBackgroundStyle = ddPBSCentered;
	m_vbPictureBackgroundUseMask = VARIANT_FALSE;
}

void CChildBands::Cleanup()
{
	int nElem = m_aChildBands.GetSize();
	for (int nIndex = 0; nIndex < nElem; ++nIndex)
		m_aChildBands.GetAt(nIndex)->Release();
	m_aChildBands.RemoveAll();
	m_pCurChildBand = NULL;
}

//{DEF IUNKNOWN MEMBERS}
STDMETHODIMP CChildBands::QueryInterface(REFIID riid, void **ppvObjOut)
{
	if (DO_GUIDS_MATCH(riid,IID_IUnknown))
	{
		AddRef();
		*ppvObjOut=this;
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IDispatch))
	{
		*ppvObjOut=(void *)(IDispatch *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_IChildBands))
	{
		*ppvObjOut=(void *)(IChildBands *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_ISupportErrorInfo))
	{
		*ppvObjOut=(void *)(ISupportErrorInfo *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_ICategorizeProperties))
	{
		*ppvObjOut=(void *)(ICategorizeProperties *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	//{CUSTOM QI}
	//{ENDCUSTOM QI}
	*ppvObjOut = NULL;
	return E_NOINTERFACE;
}
STDMETHODIMP_(ULONG) CChildBands::AddRef()
{
	return ++m_refCount;
}
STDMETHODIMP_(ULONG) CChildBands::Release()
{
	if (0!=--m_refCount)
		return m_refCount;
	delete this;
	return 0;
}
//{ENDDEF IUNKNOWN MEMBERS}
STDMETHODIMP CChildBands::get_GradientEndColor(OLE_COLOR *retval)
{
	*retval = ppV1.m_ocGradientEndColor;
	return NOERROR;
}
STDMETHODIMP CChildBands::put_GradientEndColor(OLE_COLOR val)
{
	if (ppV1.m_ocGradientEndColor != val)
	{
		ppV1.m_ocGradientEndColor = val;
		HRESULT hResult = OleTranslateColor(ppV1.m_ocGradientEndColor, NULL, &m_crGradientEndColor);
		if (FAILED(hResult))
			return hResult;
	}
	return NOERROR;
}
STDMETHODIMP CChildBands::MapPropertyToCategory( DISPID dispid,  PROPCAT*ppropcat)
{
	return E_FAIL;
}
STDMETHODIMP CChildBands::GetCategoryName( PROPCAT propcat,  LCID lcid,  BSTR*pbstrName)
{
	return E_FAIL;
}
// IDispatch members
STDMETHODIMP CChildBands::get_CurrentChildBand(Band **retval)
{
	if (NULL == retval)
		return E_INVALIDARG;

	int nBandIndex = GetCurrentChildBandIndex();
	if (nBandIndex > -1 && nBandIndex < m_aChildBands.GetSize())
	{
		*retval = (Band*)m_aChildBands.GetAt(nBandIndex);
		reinterpret_cast<CBand*>(*retval)->AddRef();
		return NOERROR;
	}
	return E_FAIL;
}
STDMETHODIMP CChildBands::put_CurrentChildBand(Band *val)
{
	if (NULL == val)
		return E_INVALIDARG;

	int nCount = m_aChildBands.GetSize();
	for (int nBand = 0; nBand < nCount; nBand++)
	{
		if ((Band*)m_aChildBands.GetAt(nBand) == val)
		{
			m_pBand->OnChildBandChanged(nBand);
			return NOERROR;
		}
	}
	return E_FAIL;
}
STDMETHODIMP CChildBands::putref_CurrentChildBand(Band * *val)
{
	if (NULL == val || NULL == *val)
		return E_INVALIDARG;

	return put_CurrentChildBand(*val);
}

STDMETHODIMP CChildBands::GetTypeInfoCount( UINT *pctinfo)
{
	*pctinfo = 1; 
	return NOERROR; 
}
STDMETHODIMP CChildBands::GetTypeInfo( UINT itinfo, LCID lcid, ITypeInfo **pptinfo)
{
	*pptinfo=GetObjectTypeInfo(lcid,objectDef);
	if ((*pptinfo)==NULL)
		return E_FAIL;
	return NOERROR;
}
STDMETHODIMP CChildBands::GetIDsOfNames( REFIID riid, OLECHAR **rgszNames, UINT cNames, ULONG lcid, long *rgdispid)
{
	HRESULT hr;
	ITypeInfo *pTypeInfo;
	if (!(riid==IID_NULL))
        return E_INVALIDARG;
	pTypeInfo=GetObjectTypeInfo(lcid,objectDef);
	if (pTypeInfo==NULL)
		return E_FAIL;
	hr=pTypeInfo->GetIDsOfNames(rgszNames, cNames, rgdispid);
	pTypeInfo->Release();
	return hr;
}
STDMETHODIMP CChildBands::Invoke( long dispidMember, REFIID riid, ULONG lcid, USHORT wFlags, DISPPARAMS *pdispparams, VARIANT *pvarResult, EXCEPINFO *pexcepinfo, UINT *puArgErr)
{
#ifdef _DEBUG
	if (dispidMember<0)
		TRACE1(1, "IDispatch::Invoke -%X\n",-dispidMember)
	else
		TRACE1(1, "IDispatch::Invoke %X\n",dispidMember)
#endif
	HRESULT hr;
	ITypeInfo *pTypeInfo;
	if (!(riid==IID_NULL))
        return E_INVALIDARG;
	pTypeInfo=GetObjectTypeInfo(lcid,objectDef);
	if (pTypeInfo==NULL)
		return E_FAIL;
	#ifdef ERRORINFO_SUPPORT
	SetErrorInfo(0L,NULL); // should be called if ISupportErrorInfo is used
	#endif
	hr = pTypeInfo->Invoke((IDispatch *)this, dispidMember, wFlags,
		pdispparams, pvarResult,
        pexcepinfo, puArgErr);
    pTypeInfo->Release();
	return hr;
}

// IChildBands members

STDMETHODIMP CChildBands::Item(VARIANT *Index, Band **retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	int index=GetPosOfItem(Index);
	if (index < 0 || index >= m_aChildBands.GetSize())
	{
		m_pBand->m_pBar->m_theErrorObject.SendError(IDERR_BADCOLLINDEX,0);
		return CUSTOM_CTL_SCODE(IDERR_BADCOLLINDEX);
	}
	CBand* pChildBand = m_aChildBands.GetAt(index);
	pChildBand->AddRef();
	*retval = (Band*)pChildBand;
	return NOERROR;
}

STDMETHODIMP CChildBands::Count( short *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval=m_aChildBands.GetSize();
	return NOERROR;
}

STDMETHODIMP CChildBands::Add(BSTR name, Band **retval)
{
	try
	{
		if (NULL == retval)
			return E_INVALIDARG;

		CBand* pChildBand = CBand::CreateInstance(NULL);
		if (NULL == pChildBand)
			return E_OUTOFMEMORY;
		
		if (m_pBand)
		{
			pChildBand->SetOwner(m_pBand->m_pBar);
			pChildBand->SetParent(m_pBand);

			if (m_pBand->m_pBar && m_pBand->m_pBar->m_pDesigner)
				static_cast<CBand*>(pChildBand)->bpV1.m_vbDesignerCreated = TRUE;
		}

		HRESULT hResult = m_aChildBands.Add(pChildBand);
		if (FAILED(hResult))
		{
			pChildBand->Release();
			return hResult;
		}

		*retval = (Band*)pChildBand;
		pChildBand->AddRef();
		hResult = pChildBand->put_Name(name);
		hResult = pChildBand->put_Caption(name);
		if (1 == m_aChildBands.GetSize())
		{
			tpV1.m_nCurTab = 0;
			m_pCurChildBand = pChildBand;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
	return NOERROR;	
}

STDMETHODIMP CChildBands::Remove( VARIANT *Index)
{
	int nIndex = GetPosOfItem(Index);
	if (nIndex == -1)
		return DISP_E_BADINDEX;

	if (nIndex < 0 || nIndex >= m_aChildBands.GetSize())
		return DISP_E_BADINDEX;

	CBand* pChildBand = m_aChildBands.GetAt(nIndex);
	if (pChildBand)
	{
		pChildBand->Release();
		m_aChildBands.RemoveAt(nIndex);
		if (pChildBand == m_pCurChildBand)
		{
			if (m_aChildBands.GetSize() > 0)
				m_pCurChildBand = m_aChildBands.GetAt(0);
			else
				m_pCurChildBand = NULL;
		}
	}
	return NOERROR;
}

STDMETHODIMP CChildBands::RemoveAll()
{
	Cleanup();
	return NOERROR;
}

int CChildBands::GetPosOfItem(VARIANT* pvIndex)
{
	int nPos = -1;
	switch (pvIndex->vt)
	{
	case VT_I2:
		nPos = pvIndex->iVal;
		break;

	case VT_I2|VT_BYREF:
		nPos = *(pvIndex->piVal);
		break;

	case VT_I4:
		nPos = pvIndex->lVal;
		break;

	case VT_I4|VT_BYREF:
		nPos = *(pvIndex->plVal);
		break;

	default:
		{
			BSTR bstrName;
			VARIANT vStr;
			VariantInit(&vStr);
			if (VT_BSTR == pvIndex->vt)
				bstrName = pvIndex->bstrVal;
			else
			{
				if (FAILED(VariantChangeType(&vStr, pvIndex, VARIANT_NOVALUEPROP, VT_BSTR)))
					return -1;
				bstrName = vStr.bstrVal;
			}

			CBand* pChildBand;
			int nElem = m_aChildBands.GetSize();
			for (int nBand = 0; nBand < nElem; nBand++)
			{
				pChildBand = m_aChildBands.GetAt(nBand);
				if (pChildBand->m_bstrName && bstrName)
				{
					if (0 == wcscmp(pChildBand->m_bstrName, bstrName))
					{
						VariantClear(&vStr);
						return nBand;
					}
				}
			}
			VariantClear(&vStr);
		}
		break;
	}

	VARIANT v;
	VariantInit(&v);
	HRESULT hResult = VariantChangeType(&v, pvIndex, VARIANT_NOVALUEPROP, VT_I4);
	if (FAILED(hResult))
		return hResult;

	nPos = v.lVal;
	return nPos;
}

STDMETHODIMP CChildBands::_NewEnum( IUnknown **retval)
{
	int cItems = m_aChildBands.GetSize();
	*retval=0;
	
	VARIANT* pColl;
	if (cItems!=0 && !(pColl = (VARIANT*)HeapAlloc(g_hHeap, 0, cItems * sizeof(VARIANT))))
		return E_OUTOFMEMORY;

	for (int nIndex=0; nIndex < cItems; ++nIndex)
	{
		(pColl+nIndex)->vt = VT_DISPATCH;
		(pColl+nIndex)->pdispVal = (IBand*)m_aChildBands.GetAt(nIndex);
	}

	IEnumX* pEnum = (IEnumX*)new CEnumX(IID_IEnumVARIANT, 
							            cItems, 
										sizeof(VARIANT), 
										pColl, 
										CopyDISPVARIANT);
	if (!pEnum)
        return E_OUTOFMEMORY;
	
	*retval=pEnum;
	return NOERROR;
}
STDMETHODIMP CChildBands::get_ToolForeColor(OLE_COLOR *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval = ppV1.m_ocToolForeColor;
	return NOERROR;
}
STDMETHODIMP CChildBands::put_ToolForeColor(OLE_COLOR val)
{
	if (ppV1.m_ocToolForeColor != val)
	{
		ppV1.m_ocToolForeColor = val;
		HRESULT hResult = OleTranslateColor(ppV1.m_ocToolForeColor, NULL, &m_crToolForeColor);
		if (FAILED(hResult))
			return hResult;
	}
	return NOERROR;
}
STDMETHODIMP CChildBands::get_ChildBandFont3D(Font3DTypes *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval = tpV1.m_fsFont3D; 
	return NOERROR;
}
STDMETHODIMP CChildBands::put_ChildBandFont3D(Font3DTypes val)
{
	switch (val)
	{
	case dd3DNone:
	case dd3DRaisedLight:
	case dd3DRaisedHeavy:
	case dd3DInsetLight:
	case dd3DInsetHeavy:
		if (tpV1.m_fsFont3D != val)
		{
			tpV1.m_fsFont3D = val;
			m_pBand->Refresh();
		}
		break;

	default:
		return E_INVALIDARG;
	}
	return NOERROR;
}
STDMETHODIMP CChildBands::get_HighLightColor(OLE_COLOR *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval = ppV1.m_ocHighLightColor;
	return NOERROR;
}
STDMETHODIMP CChildBands::put_HighLightColor(OLE_COLOR val)
{
	if (ppV1.m_ocHighLightColor != val)
	{
		ppV1.m_ocHighLightColor = val;
		HRESULT hResult = OleTranslateColor(ppV1.m_ocHighLightColor, NULL, &m_crHighLightColor);
		if (FAILED(hResult))
			return hResult;
	}
	return NOERROR;
}
STDMETHODIMP CChildBands::get_ShadowColor(OLE_COLOR *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval = ppV1.m_ocShadowColor;
	return NOERROR;
}
STDMETHODIMP CChildBands::put_ShadowColor(OLE_COLOR val)
{
	if (ppV1.m_ocShadowColor != val)
	{
		ppV1.m_ocShadowColor = val;
		HRESULT hResult = OleTranslateColor(ppV1.m_ocShadowColor, NULL, &m_crShadowColor);
		if (FAILED(hResult))
			return hResult;
	}
	return NOERROR;
}
STDMETHODIMP CChildBands::get_ChildBandCaptionAlignment(CaptionAlignmentTypes *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval = tpV1.m_caChildBands;
	return NOERROR;
}
STDMETHODIMP CChildBands::put_ChildBandCaptionAlignment(CaptionAlignmentTypes val)
{
	switch (val)
	{
    case ddCALeft:
	case ddCACenter:
	case ddCARight:
		if (tpV1.m_caChildBands != val)
		{
			tpV1.m_caChildBands = val;
			m_pBand->Refresh();
		}
		break;

	default:
		return E_INVALIDARG;
	}
	return NOERROR;
}
STDMETHODIMP CChildBands::get_PictureBackground(LPPICTUREDISP *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval = m_phPictureBackground.GetPictureDispatch();
	return NOERROR;
}
STDMETHODIMP CChildBands::put_PictureBackground(LPPICTUREDISP val)
{
	m_phPictureBackground.SetPictureDispatch(val);
	m_bBackgroundDirty = TRUE;
	if (ddBSPicture == ppV1.m_ChildBandBackgroundStyle && VARIANT_TRUE == ppV1.m_vbPictureBackgroundUseMask)
	{
		BOOL bResult = MakeTransparentBackground();
		assert(bResult);
		if (!bResult)
			return E_FAIL;
	}
	return NOERROR;
}
STDMETHODIMP CChildBands::putref_PictureBackground(LPPICTUREDISP* val)
{
	m_phPictureBackground.SetPictureDispatch(*val);
	m_bBackgroundDirty = TRUE;
	if (ddBSPicture == ppV1.m_ChildBandBackgroundStyle && VARIANT_TRUE == ppV1.m_vbPictureBackgroundUseMask)
	{
		BOOL bResult = MakeTransparentBackground();
		assert(bResult);
		if (!bResult)
			return E_FAIL;
	}
	return NOERROR;
}
STDMETHODIMP CChildBands::get_PictureBackgroundMaskColor(OLE_COLOR *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval = ppV1.m_ocPictureBackgroundMaskColor;
	return NOERROR;
}
STDMETHODIMP CChildBands::put_PictureBackgroundMaskColor(OLE_COLOR val)
{
	if (ppV1.m_ocPictureBackgroundMaskColor != val)
	{
		m_bBackgroundDirty = TRUE;
		ppV1.m_ocPictureBackgroundMaskColor = val;
		HRESULT hResult = OleTranslateColor(ppV1.m_ocPictureBackgroundMaskColor, NULL, &m_crPictureBackgroundMaskColor);
		if (FAILED(hResult))
			return hResult;
		if (ddBSPicture == ppV1.m_ChildBandBackgroundStyle && VARIANT_TRUE == ppV1.m_vbPictureBackgroundUseMask)
		{
			m_bBackgroundDirty = TRUE;
			BOOL bResult = MakeTransparentBackground();
			assert(bResult);
			if (!bResult)
				return E_FAIL;
		}
	}
	return NOERROR;
}
STDMETHODIMP CChildBands::get_PictureBackgroundStyle(PictureBackgroundStyles *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval = ppV1.m_PictureBackgroundStyle;
	return NOERROR;
}
STDMETHODIMP CChildBands::put_PictureBackgroundStyle(PictureBackgroundStyles val)
{
	switch (val)
	{
    case ddPBSCentered:
	case ddPBSStretched:
	case ddPBSTiled:
		if (ppV1.m_PictureBackgroundStyle != val)
		{
			m_bBackgroundDirty = TRUE;
			ppV1.m_PictureBackgroundStyle = val;
			m_pBand->Refresh();
		}
		break;

	default:
		return E_INVALIDARG;
	}
	return NOERROR;
}
STDMETHODIMP CChildBands::get_PictureBackgroundUseMask(VARIANT_BOOL *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval = ppV1.m_vbPictureBackgroundUseMask;
	return NOERROR;
}
STDMETHODIMP CChildBands::put_PictureBackgroundUseMask(VARIANT_BOOL val)
{
	if (ppV1.m_vbPictureBackgroundUseMask != val)
	{
		m_bBackgroundDirty = TRUE;
		ppV1.m_vbPictureBackgroundUseMask = val;
		if (ddBSPicture == ppV1.m_ChildBandBackgroundStyle && VARIANT_TRUE == ppV1.m_vbPictureBackgroundUseMask)
		{
			BOOL bResult = MakeTransparentBackground();
			assert(bResult);
			if (!bResult)
				return E_FAIL;
		}
	}
	return NOERROR;
}
STDMETHODIMP CChildBands::get_BackColor(OLE_COLOR *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval = ppV1.m_ocBackColor;
	return NOERROR;
}
STDMETHODIMP CChildBands::put_BackColor(OLE_COLOR val)
{
	if (ppV1.m_ocBackColor != val)
	{
		m_bBackgroundDirty = TRUE;
		ppV1.m_ocBackColor = val;
		HRESULT hResult = OleTranslateColor(ppV1.m_ocBackColor, NULL, &m_crBackColor);
		if (FAILED(hResult))
			return hResult;

		if (ddBSPicture == ppV1.m_ChildBandBackgroundStyle && VARIANT_TRUE == ppV1.m_vbPictureBackgroundUseMask)
		{
			BOOL bResult = MakeTransparentBackground();
			assert(bResult);
			if (!bResult)
				return E_FAIL;
		}
	}
	return NOERROR;
}
STDMETHODIMP CChildBands::get_ChildBandBackgroundStyle(ChildBandBackgroundStyles *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval = ppV1.m_ChildBandBackgroundStyle;
	return NOERROR;
}
STDMETHODIMP CChildBands::put_ChildBandBackgroundStyle(ChildBandBackgroundStyles val)
{
	switch (val)
	{
    case ddBSNormal:
	case ddBSGradient:
	case ddBSPicture:
		if (ppV1.m_ChildBandBackgroundStyle != val)
		{
			m_bBackgroundDirty = TRUE;
			ppV1.m_ChildBandBackgroundStyle = val;
			if (ddBSPicture == val && VARIANT_TRUE == ppV1.m_vbPictureBackgroundUseMask)
			{
				BOOL bResult = MakeTransparentBackground();
				assert(bResult);
				if (!bResult)
					return E_FAIL;
			}
		}
		break;

	default:
		return E_INVALIDARG;
	}
	return NOERROR;
}
STDMETHODIMP CChildBands::get_ForeColor(OLE_COLOR *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval = ppV1.m_ocForeColor;
	return NOERROR;
}
STDMETHODIMP CChildBands::put_ForeColor(OLE_COLOR val)
{
	if (ppV1.m_ocForeColor != val)
	{
		ppV1.m_ocForeColor = val;
		HRESULT hResult = OleTranslateColor(ppV1.m_ocForeColor, 0, &m_crForeColor);
		if (FAILED(hResult))
			return hResult;
	}
	return NOERROR;
}
// ISupportErrorInfo members

STDMETHODIMP CChildBands::InterfaceSupportsErrorInfo( REFIID riid)
{
    if (riid == (REFIID)INTERFACEOFOBJECT(6))
        return S_OK;

    return S_FALSE;
}
//{EVENT WRAPPERS}
//{EVENT WRAPPERS}
//
// GetVisibleBands
//

int CChildBands::GetVisibleBands(TypedArray<CBand*>& aBands)
{
	HRESULT hResult;
	CBand* pBand;
	int nCount = m_aChildBands.GetSize();
	for (int nBand = 0; nBand < nCount; nBand++)
	{
		pBand = m_aChildBands.GetAt(nBand);

		if (VARIANT_TRUE == pBand->bpV1.m_vbVisible)
			hResult = aBands.Add(pBand);
	}
	return aBands.GetSize();
}

//
// BTTabVisible
//

BOOL CChildBands::BTTabVisible(const int& index)
{
	return (VARIANT_TRUE == m_aChildBands.GetAt(index)->bpV1.m_vbVisible);
}


//
// BTGetTabCount
//

inline int CChildBands::BTGetVisibleTabCount()
{
	int nCount = m_aChildBands.GetSize();
	int nReturnCount = 0;
	for (int nIndex = 0; nIndex < nCount; nIndex++)
		if (VARIANT_TRUE == m_aChildBands.GetAt(nIndex)->bpV1.m_vbVisible)
			nReturnCount++;
	return nReturnCount;
}


//
// CommitVisibleBands
//

BOOL CChildBands::CommitVisibleBands(TypedArray<CBand*>& aBands)
{
	return SUCCEEDED(m_aVisibleChildBands.Copy(aBands));
}

//
// ScrollButtonHit
//

int CChildBands::ScrollButtonHit(POINT pt)
{
	pt.x += m_pBand->m_rcCurrentPaint.left;
	pt.y += m_pBand->m_rcCurrentPaint.top;

	if (eInactiveBottomButton != m_sbsBottomStyle && PtInRect(&m_rcBottomButton, pt))
		return eSlidingBottom;

	if (eInactiveTopButton != m_sbsTopStyle && PtInRect(&m_rcTopButton, pt))
		return eSlidingTop;
	return eSlidingNone;
}

//
// MakeTransparentBackground
//

BOOL CChildBands::MakeTransparentBackground()
{
	BOOL bResult = TRUE;
	try
	{
		BOOL bResultResouce = TRUE;

		if (NULL == m_phPictureBackground.m_pPict ||
			PICTYPE_BITMAP != m_phPictureBackground.GetType() ||
			VARIANT_FALSE == ppV1.m_vbPictureBackgroundUseMask)
		{
			return FALSE;
		}

		SIZEL sizeHiMetric;
		m_phPictureBackground.m_pPict->get_Width(&sizeHiMetric.cx);
		m_phPictureBackground.m_pPict->get_Height(&sizeHiMetric.cy);
		
		SIZEL sizePixel;
		HiMetricToPixel(&sizeHiMetric, &sizePixel);
		if (0 == sizePixel.cx || 0 == sizePixel.cy)
			return FALSE;

		CRect rcTemp(0, sizePixel.cx, 0, sizePixel.cy);

		HBITMAP hBitmapSrc;
		m_phPictureBackground.m_pPict->get_Handle((OLE_HANDLE*)&hBitmapSrc);
		if (NULL == hBitmapSrc)
			return FALSE;

		HDC hDC = GetDC(NULL);
		if (NULL == hDC)
			return FALSE;

		HBITMAP hBitmapTransOld = NULL;
		HBITMAP hBitmapMaskOld = NULL;
		HBITMAP hBitmapSrcOld = NULL;
		HBITMAP hBitmapTrans = NULL;
		HBITMAP hBitmapMask = NULL;
		HDC hDCTrans = NULL;
		HDC hDCMask = NULL;

		HDC hDCSrc = CreateCompatibleDC(hDC);
		if (NULL == hDCSrc)
		{
			bResult = FALSE;
			goto Cleanup;
		}

		hBitmapSrcOld = SelectBitmap(hDCSrc, hBitmapSrc);

		//
		// Creating the Mask Bitmap
		//

		hBitmapMask = CreateBitmap(sizePixel.cx, sizePixel.cy, 1, 1, NULL);
		if (NULL == hBitmapMask)
		{
			bResult = FALSE;
			goto Cleanup;
		}

		hDCMask = CreateCompatibleDC(hDC);
		if (NULL == hDCMask)
		{
			bResult = FALSE;
			goto Cleanup;
		}

		hBitmapMaskOld = SelectBitmap(hDCMask, hBitmapMask);

		SetBkColor(hDCSrc, m_crPictureBackgroundMaskColor);
		bResult = BitBlt(hDCMask, 0, 0, sizePixel.cx, sizePixel.cy, hDCSrc, 0, 0, SRCCOPY);
		if (!bResult)
			goto Cleanup;

		//
		// Create the Transparent bitmap
		//

		hBitmapTrans = CreateCompatibleBitmap(hDC, sizePixel.cx, sizePixel.cy);
		if (NULL == hBitmapTrans)
		{
			bResult = FALSE;
			goto Cleanup;
		}

		hDCTrans = CreateCompatibleDC(hDC);
		if (NULL == hDCTrans)
		{
			bResult = FALSE;
			goto Cleanup;
		}

		hBitmapTransOld = SelectBitmap(hDCTrans, hBitmapTrans);

		//
		// Initialize the background
		//

		FillSolidRect(hDCTrans, rcTemp, m_crBackColor);

		//
		// Make it transparent
		//

		SetBkColor(hDCTrans, RGB(255, 255, 255));
		SetTextColor(hDCTrans, RGB(0, 0, 0));
		
		bResult = BitBlt(hDCTrans, 0, 0, sizePixel.cx, sizePixel.cy, hDCSrc, 0, 0, SRCINVERT);
		if (!bResult)
			goto Cleanup;
		
		bResult = BitBlt(hDCTrans, 0, 0, sizePixel.cx, sizePixel.cy, hDCMask, 0, 0, SRCAND);
		if (!bResult)
			goto Cleanup;

		bResult = BitBlt(hDCTrans, 0, 0, sizePixel.cx, sizePixel.cy, hDCSrc, 0, 0, SRCINVERT);
		if (!bResult)
			goto Cleanup;

		//
		// Create the Transparent Picture Object
		//
		// It now owns the Transparent Bitmap
		//

		bResult = m_phPictureTransparentBackground.CreateFromBitmap(hBitmapTrans, NULL, TRUE);

Cleanup:
		if (hDCTrans)
		{
			SelectBitmap(hDCTrans, hBitmapTransOld);
			bResultResouce = DeleteDC(hDCTrans);
			assert(bResultResouce);
		}

		if (hDCMask)
		{
			SelectBitmap(hDCMask, hBitmapMaskOld);
			bResultResouce = DeleteDC(hDCMask);
			assert(bResultResouce);
		}

		if (hBitmapMask)
		{
			bResultResouce = DeleteBitmap(hBitmapMask);
			assert(bResultResouce);
		}

		if (hDCSrc)
		{
			SelectBitmap(hDCSrc, hBitmapSrcOld);
			bResultResouce = DeleteDC(hDCSrc);
		}

		ReleaseDC(NULL, hDC);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
	return bResult;
}

//
// BTGetArea
//

BTabs::AREA CChildBands::BTGetArea()
{
	try
	{
		switch (m_pBand->bpV1.m_daDockingArea)
		{
		case ddDALeft:
		case ddDARight:
			switch (m_pBand->bpV1.m_cbsChildStyle)
			{
			case ddCBSToolbarTopTabs:
				return eTabLeft;

			case ddCBSSlidingTabs:
				return eTabSlidingVert;

			case ddCBSTopTabs:
				return eTabTop;

			case ddCBSBottomTabs:
				return eTabBottom;

			case ddCBSLeftTabs:
				return eTabLeft;

			case ddCBSRightTabs:
				return eTabRight;

			default:
				return eTabRight;
			}
			break;

		case ddDAFloat:
			switch (m_pBand->bpV1.m_cbsChildStyle)
			{
			case ddCBSToolbarTopTabs:
				return eTabTop;

			case ddCBSSlidingTabs:
				return eTabSlidingVert;

			case ddCBSTopTabs:
				return eTabTop;

			case ddCBSBottomTabs:
				return eTabBottom;

			case ddCBSLeftTabs:
				return eTabLeft;

			case ddCBSRightTabs:
				return eTabRight;

			default:
				return eTabBottom;
			}
			break;

		case ddDATop:
		case ddDABottom:
			switch (m_pBand->bpV1.m_cbsChildStyle)
			{
			case ddCBSToolbarTopTabs:
				return eTabTop;

			case ddCBSSlidingTabs:
				return eTabSlidingHorz;

			case ddCBSTopTabs:
				return eTabTop;

			case ddCBSBottomTabs:
				return eTabBottom;

			case ddCBSLeftTabs:
				return eTabLeft;

			case ddCBSRightTabs:
				return eTabRight;

			default:
				return eTabBottom;
			}
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return eTabTop;
}

//
// BTGetFont
//

HFONT CChildBands::BTGetFont()
{
	try
	{
		switch (BTGetArea())
		{
		case eTabSlidingVert:
			return m_pBand->m_pBar->GetChildBandFont(FALSE);
		case eTabSlidingHorz:
			return m_pBand->m_pBar->GetChildBandFont(TRUE);

		case eTabTop:
		case eTabBottom:
			return m_pBand->m_pBar->GetChildBandFont(FALSE);

		case eTabLeft:
		case eTabRight:
			return m_pBand->m_pBar->GetChildBandFont(TRUE);
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
// BTGetFontHeight
//

int CChildBands::BTGetFontHeight()
{
	try
	{
		switch (BTGetArea())
		{
		case eTabSlidingVert:
			return m_pBand->m_pBar->GetChildBandFontHeight(FALSE) + BTabs::TABSTRIP_HSPACE;
		
		case eTabSlidingHorz:
			return m_pBand->m_pBar->GetChildBandFontHeight(TRUE) + BTabs::TABSTRIP_HSPACE;

		case eTabTop:
		case eTabBottom:
			return m_pBand->m_pBar->GetChildBandFontHeight(FALSE) + BTabs::TABSTRIP_HSPACE;

		case eTabLeft:
		case eTabRight:
			return m_pBand->m_pBar->GetChildBandFontHeight(TRUE) + BTabs::TABSTRIP_HSPACE;

		default:
			return m_pBand->m_pBar->GetChildBandFontHeight(ddDALeft == m_pBand->bpV1.m_daDockingArea || ddDARight == m_pBand->bpV1.m_daDockingArea) + BTabs::TABSTRIP_HSPACE;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return 0;
}

//
// SetOwner
//

void CChildBands::SetOwner(const CBand* pBand)
{
	try
	{
		m_pBand = const_cast<CBand*>(pBand);
		int nCount = m_aChildBands.GetSize();
		for (int nBand = 0; nBand < nCount; nBand++)
		{
			m_aChildBands.GetAt(nBand)->SetOwner(m_pBand->m_pBar);
			m_aChildBands.GetAt(nBand)->SetParent(m_pBand);
		}

	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// BTGetTabText
//

BSTR& CChildBands::BTGetTabText(const int& nIndex)
{
	assert(m_aChildBands.GetAt(nIndex));
	return m_aChildBands.GetAt(nIndex)->m_bstrCaption;
}

//
// BTGetTabPictue
//

CPictureHolder& CChildBands::BTGetTabPicture(const int& nIndex)
{
	assert(m_aChildBands.GetAt(nIndex));
	return m_aChildBands.GetAt(nIndex)->m_phPicture;
}

//
// BTGetTabWidth
//

int& CChildBands::BTGetTabWidth(const int& nIndex)
{
	assert(m_aChildBands.GetAt(nIndex));
	return m_aChildBands.GetAt(nIndex)->m_nWidth;
}

//
// BTSetTabWidth
//

void CChildBands::BTSetTabWidth(const int& nIndex, const int& nWidth)
{
	assert(m_aChildBands.GetAt(nIndex));
	m_aChildBands.GetAt(nIndex)->m_nWidth = nWidth;
}

//
// Exchange
//

HRESULT CChildBands::Exchange(IStream* pStream, VARIANT_BOOL vbSave)
{
	long nStreamSize;
	short nSize;
	short nSize2;
	short nBandCount;
	HRESULT hResult;
	if (VARIANT_TRUE == vbSave)
	{
		//
		// Saving
		//

		try
		{
			if (m_pBand && m_pBand->m_pBar && m_pBand->m_pBar->m_bSaveImages)
			{
				//
				// Saving just images
				//

				nBandCount = m_aChildBands.GetSize();
				for (int nBand = 0; nBand < nBandCount; nBand++)
				{
					hResult = m_aChildBands.GetAt(nBand)->Exchange(pStream, vbSave);
					if (FAILED(hResult))
						return hResult;
				}
			}
			else
			{
				//
				// Saving everything
				//

				nStreamSize = sizeof(nSize) + sizeof(ppV1) + sizeof(nSize) + sizeof(tpV1) + m_phPictureBackground.GetSize();
				hResult = pStream->Write(&nStreamSize, sizeof(nStreamSize), NULL);
				if (FAILED(hResult))
					return hResult;

				nSize = sizeof(ppV1);
				hResult = pStream->Write(&nSize, sizeof(nSize), 0);
				if (FAILED(hResult))
					return hResult;

				hResult = pStream->Write(&ppV1, nSize, 0);
				if (FAILED(hResult))
					return hResult;
				
				nSize = sizeof(tpV1);
				hResult = pStream->Write(&nSize, sizeof(nSize), 0);
				if (FAILED(hResult))
					return hResult;

				hResult = pStream->Write(&tpV1, sizeof(tpV1), 0);
				if (FAILED(hResult))
					return hResult;
				
				hResult = PersistPicture(pStream, &m_phPictureBackground, vbSave);
				if (FAILED(hResult))
					return hResult;

				nBandCount = m_aChildBands.GetSize();
				hResult = pStream->Write(&nBandCount, sizeof(nBandCount), 0);
				if (FAILED(hResult))
					return hResult;
				
				for (int nBand = 0; nBand < nBandCount; nBand++)
				{
					hResult = m_aChildBands.GetAt(nBand)->Exchange(pStream, vbSave);
					if (FAILED(hResult))
						return hResult;
				}
			}
		}
		CATCH
		{
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
			return E_FAIL;
		}
	}
	else
	{
		//
		// Loading
		//

		try
		{
			ULARGE_INTEGER nBeginPosition;
			LARGE_INTEGER  nOffset;

			hResult = pStream->Read(&nStreamSize, sizeof(nStreamSize), NULL);
			if (FAILED(hResult))
				return hResult;

			hResult = pStream->Read(&nSize, sizeof(nSize), 0);
			if (FAILED(hResult))
				return hResult;

			nStreamSize -= sizeof(nSize);
			if (nStreamSize <= 0)
				goto FinishedReading;

			nSize2 = sizeof(ppV1);
			hResult = pStream->Read(&ppV1,  nSize < nSize2 ? nSize : nSize2, 0);
			if (FAILED(hResult))
				return hResult;

			if (nSize2 < nSize)
			{
				ULARGE_INTEGER nBeginPosition;
				LARGE_INTEGER  nOffset;
				nOffset.HighPart = 0;
				nOffset.LowPart = nSize - nSize2;
				hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
				if (FAILED(hResult))
					return hResult;
			}

			nStreamSize -= nSize;
			if (nStreamSize <= 0)
				goto FinishedReading;

			hResult = pStream->Read(&nSize, sizeof(nSize), 0);
			if (FAILED(hResult))
				return hResult;

			nStreamSize -= sizeof(nSize);
			if (nStreamSize <= 0)
				goto FinishedReading;

			nSize2 = sizeof(tpV1);
			hResult = pStream->Read(&tpV1,  nSize < nSize2 ? nSize : nSize2, 0);
			if (FAILED(hResult))
				return hResult;

			if (nSize2 < nSize)
			{
				nOffset.HighPart = 0;
				nOffset.LowPart = nSize - nSize2;
				hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
				if (FAILED(hResult))
					return hResult;
			}

			nStreamSize -= nSize;
			if (nStreamSize <= 0)
				goto FinishedReading;

			hResult = PersistPicture(pStream, &m_phPictureBackground, vbSave);
			if (FAILED(hResult))
				return hResult;

			nStreamSize -= m_phPictureBackground.GetSize();
			if (nStreamSize <= 0)
				goto FinishedReading;

			nOffset.HighPart = 0;
			nOffset.LowPart = nStreamSize;
			hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
			if (FAILED(hResult))
				return hResult;
FinishedReading:
			hResult = pStream->Read(&nBandCount, sizeof(nBandCount),0);
			if (FAILED(hResult))
				return hResult;

			Cleanup();
			if (nBandCount > 0)
			{
				for (int nBand = 0; nBand < nBandCount; nBand++)
				{
					CBand* pNewBand = CBand::CreateInstance(NULL);
					if (NULL == pNewBand)
						return E_OUTOFMEMORY;

					pNewBand->SetOwner(m_pBand->m_pBar);
					pNewBand->SetParent(m_pBand);

					hResult = pNewBand->Exchange(pStream, vbSave);
					if (FAILED(hResult))
					{
						pNewBand->Release();
						return hResult;
					}
					
					hResult = m_aChildBands.Add(pNewBand);
					if (FAILED(hResult))
					{
						pNewBand->Release();
						return hResult;
					}
				}
				SetCurrentChildBand(tpV1.m_nCurTab);
			}

			TranslateColors();
			if (8 == GetGlobals().m_nBitDepth)
			{
				m_pBand->m_pBar->m_pColorQuantizer->AddColor(GetRValue(m_crBtnDarkShadow), 
															 GetGValue(m_crBtnDarkShadow), 
															 GetBValue(m_crBtnDarkShadow));
				m_pBand->m_pBar->m_pColorQuantizer->AddColor(GetRValue(m_crBtnHighLight), 
															 GetGValue(m_crBtnHighLight), 
															 GetBValue(m_crBtnHighLight));
				m_pBand->m_pBar->m_pColorQuantizer->AddColor(GetRValue(m_crBtnShadow), 
															 GetGValue(m_crBtnShadow), 
															 GetBValue(m_crBtnShadow));
				m_pBand->m_pBar->m_pColorQuantizer->AddColor(GetRValue(m_crBtnText), 
															 GetGValue(m_crBtnText), 
															 GetBValue(m_crBtnText));
				m_pBand->m_pBar->m_pColorQuantizer->AddColor(GetRValue(m_crBtnFace), 
															 GetGValue(m_crBtnFace), 
															 GetBValue(m_crBtnFace));

				m_pBand->m_pBar->m_pColorQuantizer->AddColor(GetRValue(m_crBackColor), 
															 GetGValue(m_crBackColor), 
															 GetBValue(m_crBackColor));
				m_pBand->m_pBar->m_pColorQuantizer->AddColor(GetRValue(m_crForeColor), 
															 GetGValue(m_crForeColor), 
															 GetBValue(m_crForeColor));
				m_pBand->m_pBar->m_pColorQuantizer->AddColor(GetRValue(m_crToolForeColor), 
															 GetGValue(m_crToolForeColor), 
															 GetBValue(m_crToolForeColor));
				m_pBand->m_pBar->m_pColorQuantizer->AddColor(GetRValue(m_crHighLightColor), 
															 GetGValue(m_crHighLightColor), 
															 GetBValue(m_crHighLightColor));
				m_pBand->m_pBar->m_pColorQuantizer->AddColor(GetRValue(m_crShadowColor), 
															 GetGValue(m_crShadowColor), 
															 GetBValue(m_crShadowColor));
			}

			if (ddBSPicture == ppV1.m_ChildBandBackgroundStyle && VARIANT_TRUE == ppV1.m_vbPictureBackgroundUseMask)
			{
				BOOL bResult = MakeTransparentBackground();
				if (!bResult)
					return E_FAIL;
			}
		}
		CATCH
		{
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
			return E_FAIL;
		}
	}
	return NOERROR;
}

//
// ExchangeConfig
//

HRESULT CChildBands::ExchangeConfig(IStream* pStream, VARIANT_BOOL vbSave)
{
	short nSize;
	short nSize2;
	int nBand;
	int nCount;
	short nBandCount;
	HRESULT hResult;
	if (VARIANT_TRUE == vbSave)
	{
		//
		// Saving
		//

		try
		{
			nSize = sizeof(ppV1);
			hResult = pStream->Write(&nSize, sizeof(nSize), 0);
			if (FAILED(hResult))
				return hResult;

			hResult = pStream->Write(&ppV1, nSize, 0);
			if (FAILED(hResult))
				return hResult;
			
			nSize = sizeof(tpV1);
			hResult = pStream->Write(&nSize, sizeof(nSize), 0);
			if (FAILED(hResult))
				return hResult;

			hResult = pStream->Write(&tpV1, sizeof(tpV1), 0);
			if (FAILED(hResult))
				return hResult;
			
			hResult = PersistPicture(pStream,&m_phPictureBackground, vbSave);
			if (FAILED(hResult))
				return hResult;

			nBandCount = m_aChildBands.GetSize();
			hResult = pStream->Write(&nBandCount, sizeof(nBandCount), 0);
			if (FAILED(hResult))
				return hResult;
			
			for (nBand = 0; nBand < nBandCount; nBand++)
			{
				//
				// Saving a list of the current child bands
				//
				
				hResult = StWriteBSTR(pStream, m_aChildBands.GetAt(nBand)->m_bstrName);
				if (FAILED(hResult))
					return hResult;
			}

			for (nBand = 0; nBand < nBandCount; nBand++)
			{
				//
				// Saving the child band properties
				//

				hResult = m_aChildBands.GetAt(nBand)->ExchangeConfig(pStream, vbSave);
				if (FAILED(hResult))
					return hResult;
			}
		}
		CATCH
		{
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
			return E_FAIL;
		}
	}
	else
	{
		try
		{
			ULARGE_INTEGER nBeginPosition;
			LARGE_INTEGER  nOffset;

			hResult = pStream->Read(&nSize, sizeof(nSize), 0);
			if (FAILED(hResult))
				return hResult;

			nSize2 = sizeof(ppV1);
			hResult = pStream->Read(&ppV1,  nSize < nSize2 ? nSize : nSize2, 0);
			if (FAILED(hResult))
				return hResult;

			if (nSize2 < nSize)
			{
				nOffset.HighPart = 0;
				nOffset.LowPart = nSize - nSize2;
				hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
				if (FAILED(hResult))
					return hResult;
			}

			hResult = pStream->Read(&nSize, sizeof(nSize), 0);
			if (FAILED(hResult))
				return hResult;

			nSize2 = sizeof(tpV1);
			hResult = pStream->Read(&tpV1,  nSize < nSize2 ? nSize : nSize2, 0);
			if (FAILED(hResult))
				return hResult;

			if (nSize2 < nSize)
			{
				nOffset.HighPart = 0;
				nOffset.LowPart = nSize - nSize2;
				hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
				if (FAILED(hResult))
					return hResult;
			}

			hResult = PersistPicture(pStream, &m_phPictureBackground, vbSave);
			if (FAILED(hResult))
				return hResult;
			
			hResult = pStream->Read(&nBandCount, sizeof(nBandCount), 0);
			if (FAILED(hResult))
				return hResult;

			if (nBandCount > 0)
			{
				BSTR* pbstrChildBandName = new BSTR[nBandCount];
				if (NULL == pbstrChildBandName)
					return E_OUTOFMEMORY;

				for (nBand = 0; nBand < nBandCount; nBand++)
				{
					pbstrChildBandName[nBand] = NULL;
					hResult = StReadBSTR(pStream, pbstrChildBandName[nBand]);
					if (FAILED(hResult))
						return hResult;
				}

				//
				// Removing the old ChildBands that are not in this configuration and marking the old 
				// ChildBands so we can tell which ones are new
				//

				BOOL* pbChildBandOld = new BOOL[nBandCount];
				if (NULL == pbChildBandOld)
					return E_OUTOFMEMORY;

				pbChildBandOld[0] = TRUE;
				for (nBand = 0; nBand < nBandCount; nBand++)
					pbChildBandOld[nBand] = FALSE;

				BSTR bstrChildBandName;
				BOOL bFound;
				nCount = m_aChildBands.GetSize();
				for (nBand = 0; nBand < nCount; nBand++)
				{
					bFound = FALSE;
					bstrChildBandName = m_aChildBands.GetAt(nBand)->m_bstrName;
					for (int nBand2 = 0; nBand2 < nBandCount; nBand2++)
					{
						if (0 == wcscmp(pbstrChildBandName[nBand2], bstrChildBandName))
						{
							pbChildBandOld[nBand2] = TRUE;
							bFound = TRUE;
							break;
						}
					}
					if (!bFound)
					{
						VARIANT vChildBand;
						vChildBand.vt = VT_I4;
						vChildBand.lVal = nBand;
						Remove(&vChildBand);
					}
				}

				//
				// Reading in the ChildBand information
				//

				CBand* pChildBand;
				VARIANT vBand;
				vBand.vt = VT_BSTR;
				for (nBand = 0; nBand < nBandCount; nBand++)
				{
					if (!pbChildBandOld[nBand])
					{
						//
						// Adding the New Child Bands
						//

						pChildBand = CBand::CreateInstance(NULL);
						if (NULL == pChildBand)
							return E_OUTOFMEMORY;
						
						pChildBand->SetOwner(m_pBand->m_pBar);
						pChildBand->SetParent(m_pBand);
						hResult = m_aChildBands.Add(pChildBand);
						if (FAILED(hResult))
						{
							pChildBand->Release();
							return hResult;
						}

						hResult = pChildBand->ExchangeConfig(pStream, vbSave);
						if (FAILED(hResult))
						{
							pChildBand->Release();
							return hResult;
						}
					}
					else
					{
						//
						// Updating the Old ChildBands
						//

						vBand.bstrVal = pbstrChildBandName[nBand];
						int nPos = GetPosOfItem(&vBand);
						if (-1 != nPos)
						{
							//
							// Found the old band 
							//

							hResult = m_aChildBands.GetAt(nPos)->ExchangeConfig(pStream, vbSave);
							if (FAILED(hResult))
								return hResult;
						}
					}
				}
				
				for (nBand = 0; nBand < nBandCount; nBand++)
					SysFreeString(pbstrChildBandName[nBand]);

				delete [] pbstrChildBandName;
				delete [] pbChildBandOld;

				SetCurrentChildBand(tpV1.m_nCurTab);
			}
			TranslateColors();
			if (ddBSPicture == ppV1.m_ChildBandBackgroundStyle && VARIANT_TRUE == ppV1.m_vbPictureBackgroundUseMask)
			{
				BOOL bResult = MakeTransparentBackground();
				if (!bResult)
					return E_FAIL;
			}
		}
		CATCH
		{
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
			return E_FAIL;
		}
	}
	return NOERROR;
}

HRESULT CChildBands::DragDropExchange(IStream* pStream, VARIANT_BOOL vbSave)
{
	HRESULT hResult;
	short nBandCount;
	if (VARIANT_TRUE == vbSave)
	{
		//
		// Saving
		//

		try
		{
			//
			// Saving everything
			//

			hResult = pStream->Write(&ppV1, sizeof(ppV1), NULL);
			if (FAILED(hResult))
				return hResult;
			
			hResult = pStream->Write(&tpV1, sizeof(tpV1), NULL);
			if (FAILED(hResult))
				return hResult;
			
			hResult = PersistPicture(pStream, &m_phPictureBackground, vbSave);
			if (FAILED(hResult))
				return hResult;

			nBandCount = m_aChildBands.GetSize();
			hResult = pStream->Write(&nBandCount, sizeof(nBandCount), 0);
			if (FAILED(hResult))
				return hResult;
			
			for (int nBand = 0; nBand < nBandCount; nBand++)
			{
				hResult = m_aChildBands.GetAt(nBand)->DragDropExchange(pStream, vbSave);
				if (FAILED(hResult))
					return hResult;
			}
		}
		CATCH
		{
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
			return E_FAIL;
		}
	}
	else
	{
		//
		// Loading
		//

		try
		{
			hResult = pStream->Read(&ppV1, sizeof(ppV1), NULL);
			if (FAILED(hResult))
				return hResult;

			hResult = pStream->Read(&tpV1,  sizeof(tpV1), NULL);
			if (FAILED(hResult))
				return hResult;

			hResult = PersistPicture(pStream, &m_phPictureBackground, vbSave);
			if (FAILED(hResult))
				return hResult;

			hResult = pStream->Read(&nBandCount, sizeof(nBandCount),0);
			if (FAILED(hResult))
				return hResult;

			Cleanup();
			if (nBandCount > 0)
			{
				for (int nBand = 0; nBand < nBandCount; nBand++)
				{
					CBand* pNewBand = CBand::CreateInstance(NULL);
					if (NULL == pNewBand)
						return E_OUTOFMEMORY;

					pNewBand->SetOwner(m_pBand->m_pBar);
					pNewBand->SetParent(m_pBand);

					hResult = pNewBand->DragDropExchange(pStream, vbSave);
					if (FAILED(hResult))
					{
						pNewBand->Release();
						return hResult;
					}
					
					hResult = m_aChildBands.Add(pNewBand);
					if (FAILED(hResult))
					{
						pNewBand->Release();
						return hResult;
					}
				}
				SetCurrentChildBand(tpV1.m_nCurTab);
			}

			TranslateColors();

			if (ddBSPicture == ppV1.m_ChildBandBackgroundStyle && VARIANT_TRUE == ppV1.m_vbPictureBackgroundUseMask)
			{
				BOOL bResult = MakeTransparentBackground();
				if (!bResult)
					return E_FAIL;
			}
		}
		CATCH
		{
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
			return E_FAIL;
		}
	}
	return NOERROR;
}

//
// CopyTo
//

STDMETHODIMP CChildBands::CopyTo(CChildBands* pDest)
{
	if (pDest == this)
		return NOERROR;
	try
	{
		memcpy (&pDest->ppV1, &ppV1, sizeof(ppV1));

		pDest->m_crToolForeColor = m_crToolForeColor;
		pDest->m_crHighLightColor = m_crHighLightColor;
		pDest->m_crShadowColor = m_crShadowColor;
		pDest->m_crPictureBackgroundMaskColor = m_crPictureBackgroundMaskColor;
		pDest->m_crBtnDarkShadow = m_crBtnDarkShadow;
		pDest->m_crBtnHighLight = m_crBtnHighLight;
		pDest->m_crBtnShadow = m_crBtnShadow;
		pDest->m_crBackColor = m_crBackColor;
		pDest->m_crForeColor = m_crForeColor;
		pDest->m_crBtnText = m_crBtnText;
		pDest->m_crBtnFace = m_crBtnFace;

		//
		// Copying the ChildBands
		//

		pDest->Cleanup();

		CBand* pDestChildBand;
		CBand* pSrcChildBand;
		HRESULT hResult;
		int nCount = m_aChildBands.GetSize();
		for (int nBand = 0; nBand < nCount; nBand++)
		{
			pSrcChildBand = m_aChildBands.GetAt(nBand);
			assert(pSrcChildBand);
			if (NULL == pSrcChildBand)
				continue;

			//
			// Creating and adding a Band to the ChildBand collection
			//

			hResult = pDest->Add(pSrcChildBand->m_bstrCaption, (Band**)&pDestChildBand);
			if (FAILED(hResult))
				continue;

			//
			// Copying the ChildBand's Tools
			//

			hResult = pSrcChildBand->CopyTo(pDestChildBand);
			if (FAILED(hResult))
				return hResult;

			pDestChildBand->Release();
		}
		pDest->SetCurrentChildBand(tpV1.m_nCurTab);
		return NOERROR;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
}

//
// SetCurrentChildBand
//

void CChildBands::SetCurrentChildBand(const short& nNewChildBand)
{
	BOOL bResult = TRUE;
	try
	{
		if (!m_pBand->m_bLoading && nNewChildBand == tpV1.m_nCurTab)
			return;

		int nBandCount = m_aChildBands.GetSize();
		if (0 == nBandCount)
		{
			m_pCurChildBand = NULL;
			m_nPrevTab = -1;
			tpV1.m_nCurTab = -1;
			return;
		}
		else if (1 == nBandCount)
		{
			m_pCurChildBand = m_aChildBands.GetAt(0);
			m_nPrevTab = -1;
			tpV1.m_nCurTab = 0;
		}

		assert(m_pBand);
		if (m_pBand && ddCBSSlidingTabs == m_pBand->bpV1.m_cbsChildStyle)
		{
			//
			// Save the Current Band for Animation of the Sliding Band
			//

			HWND  hWndDock;
			CRect rcChildBands;
			if (m_pBand->GetBandRect(hWndDock, rcChildBands))
			{
				m_pBand->AdjustChildBandsRect(rcChildBands);

				HDC hDC = GetDC(hWndDock);
				if (hDC)
				{
					HDC hDCOld = m_pBand->m_ffOldBand.RequestDC(hDC, rcChildBands.Width(), rcChildBands.Height());
					if (NULL == hDCOld)
						hDCOld = hDC;
					else
						rcChildBands.Offset(-rcChildBands.left, -rcChildBands.top);

					POINT pt = {0, 0};
					bResult = Draw(hDCOld, 
								   rcChildBands, 
								   m_pBand->IsWrappable(), 
								   ddDALeft == m_pBand->m_pRootBand->bpV1.m_daDockingArea || ddDARight == m_pBand->m_pRootBand->bpV1.m_daDockingArea, 
								   FALSE,
								   TRUE, 
								   FALSE,
								   pt,
								   FALSE,
								   FALSE);
					assert(bResult);
					ReleaseDC(hWndDock, hDC);
				}
			}
		}
		m_nPrevTab = tpV1.m_nCurTab;
		tpV1.m_nCurTab = nNewChildBand;

		if (tpV1.m_nCurTab >= nBandCount)
			tpV1.m_nCurTab = nBandCount - 1;

		if (tpV1.m_nCurTab < 0)
			tpV1.m_nCurTab = 0;

		if (m_pBand &&  m_pCurChildBand)
			m_pCurChildBand->TabbedWindowedTools(FALSE);

		m_pCurChildBand = m_aChildBands.GetAt(tpV1.m_nCurTab);
		if (m_pCurChildBand)
			m_pCurChildBand->TabbedWindowedTools(TRUE);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// CalcHorzLayout
//

BOOL CChildBands::CalcHorzLayout(HDC    hDC,
							     const CRect& rcBound,
								 BOOL   bCommit,
								 int    nTotalWidth,
								 BOOL   bWrapFlag,
								 BOOL   bLeftHandle,
								 BOOL   bTopHandle,
								 CRect& rcReturn,
								 int&   nReturnLineCount,
								 DWORD  dwLayoutFlag)
{
	try
	{
		BOOL bResult = FALSE;
		int nSlidingEndLength = 0;
		BOOL bVertical = (dwLayoutFlag & CBand::eLayoutVert);

		if (bCommit)
			m_bBackgroundDirty = TRUE;

		switch (m_pBand->bpV1.m_cbsChildStyle)
		{
		case ddCBSTopTabs:
			BTGetRect(m_pBand->m_rcChildBandArea);
			nTotalWidth -= CBand::eBevelBorder2;
			rcReturn.right += CBand::eBevelBorder2;
			rcReturn.bottom += m_pBand->m_rcChildBandArea.top + CBand::eBevelBorder2;
			break;

		case ddCBSBottomTabs:
			BTGetRect(m_pBand->m_rcChildBandArea);
			nTotalWidth -= CBand::eBevelBorder2;
			rcReturn.right += CBand::eBevelBorder2;
			rcReturn.bottom += CBand::eBevelBorder2;
			break;

		case ddCBSLeftTabs:
			BTGetRect(m_pBand->m_rcChildBandArea);
			nTotalWidth -= CBand::eBevelBorder2;
			rcReturn.right += m_pBand->m_rcChildBandArea.left + CBand::eBevelBorder2;
			rcReturn.bottom += CBand::eBevelBorder2;
			break;

		case ddCBSRightTabs:
			BTGetRect(m_pBand->m_rcChildBandArea);
			nTotalWidth -= CBand::eBevelBorder2;
			rcReturn.right += CBand::eBevelBorder2;
			rcReturn.bottom += CBand::eBevelBorder2;
			break;

		case ddCBSToolbarTopTabs:
			BTGetRect(m_pBand->m_rcChildBandArea);
			nTotalWidth -= CBand::eBevelBorder2;
			if (bVertical)
			{
				rcReturn.right += m_pBand->m_rcChildBandArea.left + CBand::eBevelBorder2;
				rcReturn.bottom += CBand::eBevelBorder2;
			}
			else
			{
				rcReturn.right += CBand::eBevelBorder2;
				rcReturn.bottom += m_pBand->m_rcChildBandArea.top + CBand::eBevelBorder2;
			}
			break;
		
		case ddCBSToolbarBottomTabs:
			BTGetRect(m_pBand->m_rcChildBandArea);
			nTotalWidth -= CBand::eBevelBorder2;
			rcReturn.right += CBand::eBevelBorder2;
			rcReturn.bottom += CBand::eBevelBorder2;
			break;
		
		case ddCBSSlidingTabs:
			//
			// Space for the inside edge for Sliding Tabs
			//

			if (!bVertical)
				bVertical = (dwLayoutFlag & CBand::eLayoutFloat);

			nTotalWidth -= 2 * CBand::eBevelBorder; 
			rcReturn.right += CBand::eBevelBorder;
			rcReturn.bottom += CBand::eBevelBorder;

			//
			// Space for the tabs
			//

			BTGetRect(m_pBand->m_rcChildBandArea);
			nSlidingEndLength = BTCalcEndLen();
			if (bVertical)
			{
				rcReturn.right += CBand::eBevelBorder;
				rcReturn.bottom += m_pBand->m_rcChildBandArea.top - nSlidingEndLength + CBand::eBevelBorder;
			}
			else
			{
				nTotalWidth -= m_pBand->m_rcChildBandArea.left + CBand::eBevelBorder;
				rcReturn.right += m_pBand->m_rcChildBandArea.left - nSlidingEndLength + CBand::eBevelBorder;
				rcReturn.bottom += CBand::eBevelBorder;
			}
			break;
		}

		m_pBand->m_sizeBandEdge = rcReturn.Size();
		m_pBand->m_sizeEdgeOffset = m_pBand->m_sizeBandEdge;

		if (m_pCurChildBand)
		{
			bResult = m_pCurChildBand->InsideCalcHorzLayout(hDC, rcBound, bCommit, nTotalWidth, bWrapFlag, bLeftHandle, bTopHandle, rcReturn, nReturnLineCount, dwLayoutFlag);
			assert(bResult);
		}

		switch (m_pBand->bpV1.m_cbsChildStyle)
		{
		case ddCBSTopTabs:
			rcReturn.right += CBand::eBevelBorder2;
			rcReturn.bottom += CBand::eBevelBorder2;
			m_pBand->m_sizeBandEdge.cx += CBand::eBevelBorder2;
			m_pBand->m_sizeBandEdge.cy += CBand::eBevelBorder2;
			break;

		case ddCBSBottomTabs:
			rcReturn.right += CBand::eBevelBorder2;
			rcReturn.bottom += m_pBand->m_rcChildBandArea.bottom + CBand::eBevelBorder2;

			m_pBand->m_sizeBandEdge.cx += CBand::eBevelBorder2;
			m_pBand->m_sizeBandEdge.cy += m_pBand->m_rcChildBandArea.bottom + CBand::eBevelBorder2;
			break;

		case ddCBSLeftTabs:
			rcReturn.right += CBand::eBevelBorder2;
			rcReturn.bottom += CBand::eBevelBorder2;
			m_pBand->m_sizeBandEdge.cx += CBand::eBevelBorder2;
			m_pBand->m_sizeBandEdge.cy += CBand::eBevelBorder2;
			break;

		case ddCBSRightTabs:
			rcReturn.right += m_pBand->m_rcChildBandArea.right + CBand::eBevelBorder2;
			rcReturn.bottom += CBand::eBevelBorder2;

			m_pBand->m_sizeBandEdge.cx += m_pBand->m_rcChildBandArea.right + CBand::eBevelBorder2;
			m_pBand->m_sizeBandEdge.cy += CBand::eBevelBorder2;
			break;

		case ddCBSToolbarTopTabs:
			rcReturn.right += CBand::eBevelBorder2;
			rcReturn.bottom += CBand::eBevelBorder2;
			m_pBand->m_sizeBandEdge.cx += CBand::eBevelBorder2;
			m_pBand->m_sizeBandEdge.cy += CBand::eBevelBorder2;
			break;

		case ddCBSToolbarBottomTabs:
			if (bVertical)
			{
				rcReturn.right += m_pBand->m_rcChildBandArea.right + CBand::eBevelBorder2;
				rcReturn.bottom += CBand::eBevelBorder2;

				m_pBand->m_sizeBandEdge.cx += m_pBand->m_rcChildBandArea.right + CBand::eBevelBorder2;
				m_pBand->m_sizeBandEdge.cy += CBand::eBevelBorder2;
			}
			else
			{
				rcReturn.right += CBand::eBevelBorder2;
				rcReturn.bottom += m_pBand->m_rcChildBandArea.bottom + CBand::eBevelBorder2;

				m_pBand->m_sizeBandEdge.cx += CBand::eBevelBorder2;
				m_pBand->m_sizeBandEdge.cy += m_pBand->m_rcChildBandArea.bottom + CBand::eBevelBorder2;
			}
			break;
		
		case ddCBSSlidingTabs:
		
			//
			// Space for the tabs
			//

			if (bVertical)
			{
				m_pBand->m_sizeBandEdge.cx += CBand::eBevelBorder;
				m_pBand->m_sizeBandEdge.cy += CBand::eBevelBorder + nSlidingEndLength;
			}
			else
			{
				m_pBand->m_sizeBandEdge.cx += CBand::eBevelBorder + nSlidingEndLength;
				m_pBand->m_sizeBandEdge.cy += CBand::eBevelBorder;
			}

			//
			// More space for the inside edge for Sliding Tabs
			//

			rcReturn.right += CBand::eBevelBorder;
			rcReturn.bottom += CBand::eBevelBorder;
			break;
		}
		return bResult;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
}

//
// Draw
//

BOOL CChildBands::Draw(HDC	 hDC, 
					   CRect rcPaint, 
					   BOOL	 bWrapFlag, 
					   BOOL	 bVertical, 
					   BOOL	 bVerticalPaint, 
					   BOOL	 bLeftHandle, 
					   BOOL	 bRightHandle,
					   const POINT& ptPaintOffset,
					   BOOL	 bClipRgn,
					   BOOL	 bDrawControlsOrForms)
{
	BOOL bResult = TRUE;
	try
	{
		m_nHorzSlidingOffset = 0;
		m_nVertSlidingOffset = 0;
		CRect rcEdge = rcPaint;
		CRect rcTabs = rcPaint;

		HPALETTE hPal = m_pBand->m_pBar->Palette();
		switch (m_pBand->bpV1.m_cbsChildStyle)
		{

		case ddCBSTopTabs:
			rcEdge.top += m_pBand->m_rcChildBandArea.top;
			rcTabs.bottom = rcTabs.top + m_pBand->m_rcChildBandArea.top;

			m_pBand->m_pBar->DrawEdge(hDC, rcEdge, BDR_RAISEDINNER, BF_RECT);

			m_crBtnDarkShadow = GetSysColor(COLOR_3DDKSHADOW);
			m_crBtnHighLight = VARIANT_FALSE == m_pBand->m_pBar->bpV1.m_vbXPLook ? m_pBand->m_pBar->m_crHighLight : m_pBand->m_pBar->m_crXPHighLight;
			m_crBtnShadow = VARIANT_FALSE == m_pBand->m_pBar->bpV1.m_vbXPLook ? m_pBand->m_pBar->m_crShadow : m_pBand->m_pBar->m_crXPShadow;
			m_crBtnText = m_pBand->m_pBar->m_crForeground;
			m_crBtnFace = VARIANT_FALSE == m_pBand->m_pBar->bpV1.m_vbXPLook ? m_pBand->m_pBar->m_crBackground : m_pBand->m_pBar->m_crXPBackground;
			
			bResult = BTabs::Draw(m_pBand->m_pBar, hDC, rcTabs, rcPaint, ptPaintOffset);
			assert(bResult);
			if (!bResult)
				return bResult;
			
			rcPaint.top += m_pBand->m_rcChildBandArea.top + CBand::eBevelBorder2;
			rcPaint.left += m_pBand->m_rcChildBandArea.left + CBand::eBevelBorder2;
			rcPaint.right -= m_pBand->m_rcChildBandArea.right;
			rcPaint.bottom -= m_pBand->m_rcChildBandArea.bottom;
			break;

		case ddCBSBottomTabs:
			rcEdge.bottom -= m_pBand->m_rcChildBandArea.bottom;
			rcTabs.top = rcTabs.bottom - m_pBand->m_rcChildBandArea.bottom;
			rcTabs.top--;

			bResult = m_pBand->m_pBar->DrawEdge(hDC, rcEdge, BDR_RAISEDINNER, BF_RECT);
			assert(bResult);
			if (!bResult)
				return bResult;

			m_crBtnDarkShadow = GetSysColor(COLOR_3DDKSHADOW);
			m_crBtnHighLight = VARIANT_FALSE == m_pBand->m_pBar->bpV1.m_vbXPLook ? m_pBand->m_pBar->m_crHighLight : m_pBand->m_pBar->m_crXPHighLight;
			m_crBtnShadow = VARIANT_FALSE == m_pBand->m_pBar->bpV1.m_vbXPLook ? m_pBand->m_pBar->m_crShadow : m_pBand->m_pBar->m_crXPShadow;
			m_crBtnText = m_pBand->m_pBar->m_crForeground;
			m_crBtnFace = VARIANT_FALSE == m_pBand->m_pBar->bpV1.m_vbXPLook ? m_pBand->m_pBar->m_crBackground : m_pBand->m_pBar->m_crXPBackground;
			
			bResult = BTabs::Draw(m_pBand->m_pBar, hDC, rcTabs, rcPaint, ptPaintOffset);
			assert(bResult);
			if (!bResult)
				return bResult;
			
			rcPaint.top += m_pBand->m_rcChildBandArea.top + CBand::eBevelBorder2;
			rcPaint.left += m_pBand->m_rcChildBandArea.left + CBand::eBevelBorder2;
			rcPaint.right -= m_pBand->m_rcChildBandArea.right;
			rcPaint.bottom -= m_pBand->m_rcChildBandArea.bottom;
			break;

		case ddCBSLeftTabs:
			rcEdge.left += m_pBand->m_rcChildBandArea.left;
			rcTabs.right = rcTabs.left + m_pBand->m_rcChildBandArea.left;

			m_pBand->m_pBar->DrawEdge(hDC, rcEdge, BDR_RAISEDINNER, BF_RECT);

			m_crBtnDarkShadow = GetSysColor(COLOR_3DDKSHADOW);
			m_crBtnHighLight = VARIANT_FALSE == m_pBand->m_pBar->bpV1.m_vbXPLook ? m_pBand->m_pBar->m_crHighLight : m_pBand->m_pBar->m_crXPHighLight;
			m_crBtnShadow = VARIANT_FALSE == m_pBand->m_pBar->bpV1.m_vbXPLook ? m_pBand->m_pBar->m_crShadow : m_pBand->m_pBar->m_crXPShadow;
			m_crBtnText = m_pBand->m_pBar->m_crForeground;
			m_crBtnFace = VARIANT_FALSE == m_pBand->m_pBar->bpV1.m_vbXPLook ? m_pBand->m_pBar->m_crBackground : m_pBand->m_pBar->m_crXPBackground;
			
			bResult = BTabs::Draw(m_pBand->m_pBar, hDC, rcTabs, rcPaint, ptPaintOffset);
			assert(bResult);
			if (!bResult)
				return bResult;
			
			rcPaint.top += m_pBand->m_rcChildBandArea.top + CBand::eBevelBorder2;
			rcPaint.left += m_pBand->m_rcChildBandArea.left + CBand::eBevelBorder2;
			rcPaint.right -= m_pBand->m_rcChildBandArea.right;
			rcPaint.bottom -= m_pBand->m_rcChildBandArea.bottom;
			break;

		case ddCBSRightTabs:
			rcEdge.right -= m_pBand->m_rcChildBandArea.right;
			rcTabs.left = rcTabs.right - m_pBand->m_rcChildBandArea.right;
			rcTabs.left--;

			bResult = m_pBand->m_pBar->DrawEdge(hDC, rcEdge, BDR_RAISEDINNER, BF_RECT);
			assert(bResult);
			if (!bResult)
				return bResult;

			m_crBtnDarkShadow = GetSysColor(COLOR_3DDKSHADOW);
			m_crBtnHighLight = VARIANT_FALSE == m_pBand->m_pBar->bpV1.m_vbXPLook ? m_pBand->m_pBar->m_crHighLight : m_pBand->m_pBar->m_crXPHighLight;
			m_crBtnShadow = VARIANT_FALSE == m_pBand->m_pBar->bpV1.m_vbXPLook ? m_pBand->m_pBar->m_crShadow : m_pBand->m_pBar->m_crXPShadow;
			m_crBtnText = m_pBand->m_pBar->m_crForeground;
			m_crBtnFace = VARIANT_FALSE == m_pBand->m_pBar->bpV1.m_vbXPLook ? m_pBand->m_pBar->m_crBackground : m_pBand->m_pBar->m_crXPBackground;
			
			bResult = BTabs::Draw(m_pBand->m_pBar, hDC, rcTabs, rcPaint, ptPaintOffset);
			assert(bResult);
			if (!bResult)
				return bResult;
			
			rcPaint.top += m_pBand->m_rcChildBandArea.top + CBand::eBevelBorder2;
			rcPaint.left += m_pBand->m_rcChildBandArea.left + CBand::eBevelBorder2;
			rcPaint.right -= m_pBand->m_rcChildBandArea.right;
			rcPaint.bottom -= m_pBand->m_rcChildBandArea.bottom;
			break;

		case ddCBSToolbarTopTabs:
			if (bVertical)
			{
				rcEdge.left += m_pBand->m_rcChildBandArea.left;
				rcTabs.right = rcTabs.left + m_pBand->m_rcChildBandArea.left;
			}
			else
			{
				rcEdge.top += m_pBand->m_rcChildBandArea.top;
				rcTabs.bottom = rcTabs.top + m_pBand->m_rcChildBandArea.top;
			}

			m_pBand->m_pBar->DrawEdge(hDC, rcEdge, BDR_RAISEDINNER, BF_RECT);

			m_crBtnDarkShadow = GetSysColor(COLOR_3DDKSHADOW);
			m_crBtnHighLight = VARIANT_FALSE == m_pBand->m_pBar->bpV1.m_vbXPLook ? m_pBand->m_pBar->m_crHighLight : m_pBand->m_pBar->m_crXPHighLight;
			m_crBtnShadow = VARIANT_FALSE == m_pBand->m_pBar->bpV1.m_vbXPLook ? m_pBand->m_pBar->m_crShadow : m_pBand->m_pBar->m_crXPShadow;
			m_crBtnText = m_pBand->m_pBar->m_crForeground;
			m_crBtnFace = VARIANT_FALSE == m_pBand->m_pBar->bpV1.m_vbXPLook ? m_pBand->m_pBar->m_crBackground : m_pBand->m_pBar->m_crXPBackground;
			
			bResult = BTabs::Draw(m_pBand->m_pBar, hDC, rcTabs, rcPaint, ptPaintOffset);
			assert(bResult);
			if (!bResult)
				return bResult;
			
			rcPaint.top += m_pBand->m_rcChildBandArea.top + CBand::eBevelBorder2;
			rcPaint.left += m_pBand->m_rcChildBandArea.left + CBand::eBevelBorder2;
			rcPaint.right -= m_pBand->m_rcChildBandArea.right;
			rcPaint.bottom -= m_pBand->m_rcChildBandArea.bottom;
			break;

		case ddCBSToolbarBottomTabs:
			if (bVertical)
			{
				rcEdge.right -= m_pBand->m_rcChildBandArea.right;
				rcTabs.left = rcTabs.right - m_pBand->m_rcChildBandArea.right;
				rcTabs.left--;
			}
			else
			{
				rcEdge.bottom -= m_pBand->m_rcChildBandArea.bottom;
				rcTabs.top = rcTabs.bottom - m_pBand->m_rcChildBandArea.bottom;
				rcTabs.top--;
			}

			bResult = m_pBand->m_pBar->DrawEdge(hDC, rcEdge, BDR_RAISEDINNER, BF_RECT);
			assert(bResult);
			if (!bResult)
				return bResult;

			m_crBtnDarkShadow = GetSysColor(COLOR_3DDKSHADOW);
			m_crBtnHighLight = VARIANT_FALSE == m_pBand->m_pBar->bpV1.m_vbXPLook ? m_pBand->m_pBar->m_crHighLight : m_pBand->m_pBar->m_crXPHighLight;
			m_crBtnShadow = VARIANT_FALSE == m_pBand->m_pBar->bpV1.m_vbXPLook ? m_pBand->m_pBar->m_crShadow : m_pBand->m_pBar->m_crXPShadow;
			m_crBtnText = m_pBand->m_pBar->m_crForeground;
			m_crBtnFace = VARIANT_FALSE == m_pBand->m_pBar->bpV1.m_vbXPLook ? m_pBand->m_pBar->m_crBackground : m_pBand->m_pBar->m_crXPBackground;
			
			bResult = BTabs::Draw(m_pBand->m_pBar, hDC, rcTabs, rcPaint, ptPaintOffset);
			assert(bResult);
			if (!bResult)
				return bResult;
			
			rcPaint.top += m_pBand->m_rcChildBandArea.top + CBand::eBevelBorder2;
			rcPaint.left += m_pBand->m_rcChildBandArea.left + CBand::eBevelBorder2;
			rcPaint.right -= m_pBand->m_rcChildBandArea.right;
			rcPaint.bottom -= m_pBand->m_rcChildBandArea.bottom;
			break;

		case ddCBSSlidingTabs:
			{
				//
				// Painting the Sliding Tabs background
				//

				switch (ppV1.m_ChildBandBackgroundStyle)
				{
				case ddBSNormal:
					FillSolidRect(hDC, rcPaint, m_crBackColor);
					break;

				case ddBSGradient:
					{
						if (m_bBackgroundDirty)
						{
							long nWidth = rcPaint.Width();
							long nHeight = rcPaint.Height();
							bResult = CreateGradient(hDC, hPal, m_ffPageBackground, m_crBackColor, m_crGradientEndColor, bVertical, nWidth, nHeight);
							assert(bResult);
							if (!bResult)
								return FALSE;
							m_bBackgroundDirty = FALSE;
						}
						m_ffPageBackground.Paint(hDC, rcPaint.left, rcPaint.top);
					}
					break;

				case ddBSPicture:
					{
						long nWidth = rcPaint.Width();
						long nHeight = rcPaint.Height();
						CRect rcBounds(0, nWidth, 0, nHeight);

						HDC hDCNoFlicker = m_ffPageBackground.RequestDC(hDC, nWidth, nHeight);
						assert(hDCNoFlicker);
						if (NULL == hDCNoFlicker)
							hDCNoFlicker = hDC;

						if (hDCNoFlicker)
						{
							HPALETTE hPalOld;
							if (hPal)
							{
								hPalOld = SelectPalette(hDCNoFlicker, hPal, FALSE);
								RealizePalette(hDCNoFlicker);
							}
							if (m_phPictureBackground.m_pPict)
							{
								if (m_bBackgroundDirty)
								{
									switch (ppV1.m_PictureBackgroundStyle)
									{
									case ddPBSCentered:
										{
											FillSolidRect(hDCNoFlicker, rcBounds, m_crBackColor);

											SIZEL sizeHiMetric;
											m_phPictureBackground.m_pPict->get_Width(&sizeHiMetric.cx);
											m_phPictureBackground.m_pPict->get_Height(&sizeHiMetric.cy);

											SIZEL sizePixel;
											HiMetricToPixel(&sizeHiMetric, &sizePixel);

											CRect rcCenter;
											rcCenter.left = (rcBounds.Width() - sizePixel.cx) / 2;
											rcCenter.top = (rcBounds.Height() - sizePixel.cy) / 2;
											rcCenter.right = rcCenter.left + sizePixel.cx;
											rcCenter.bottom = rcCenter.top + sizePixel.cy;
											
											IntersectRect(&rcCenter, &rcCenter, &rcBounds);
											if (VARIANT_TRUE == ppV1.m_vbPictureBackgroundUseMask)
												m_phPictureTransparentBackground.Render(hDCNoFlicker, rcCenter, rcBounds);
											else
												m_phPictureBackground.Render(hDCNoFlicker, rcCenter, rcBounds);
										}
										break;

									case ddPBSStretched:
										{
											FillSolidRect(hDC, rcBounds, m_crBackColor);

											if (VARIANT_TRUE == ppV1.m_vbPictureBackgroundUseMask)
												m_phPictureTransparentBackground.Render(hDCNoFlicker, rcBounds, rcBounds);
											else
												m_phPictureBackground.Render(hDCNoFlicker, rcBounds, rcBounds);
										}
										break;

									case ddPBSTiled:
										{
											HBITMAP hBitmapTiled;
											HRESULT hResult;
											if (VARIANT_TRUE == ppV1.m_vbPictureBackgroundUseMask)
												hResult = m_phPictureTransparentBackground.m_pPict->get_Handle((OLE_HANDLE*)&hBitmapTiled);
											else
												hResult = m_phPictureBackground.m_pPict->get_Handle((OLE_HANDLE*)&hBitmapTiled);

											if (SUCCEEDED(hResult) && hBitmapTiled)
											{
												HDC hDCNoFlicker = m_ffPageBackground.RequestDC(hDC, nWidth, nHeight);
												if (hDCNoFlicker)
												{
													HPALETTE hPalOld;
													if (hPal)
													{
														hPalOld = SelectPalette(hDCNoFlicker, hPal, FALSE);
														RealizePalette(hDCNoFlicker);
													}

													if (g_fSysWinNT)
													{
														HBRUSH hBrushTexture = CreatePatternBrush(hBitmapTiled);
														assert(hBrushTexture);
														if (hBrushTexture)
														{
															FillRect(hDCNoFlicker, &rcBounds, hBrushTexture);
															bResult = DeleteBrush(hBrushTexture);
															assert(bResult);
														}
													}
													else
													{
														BITMAP bmInfo;
														GetObject(hBitmapTiled,sizeof(BITMAP),&bmInfo);
														FillRgnBitmap(hDCNoFlicker, 0, hBitmapTiled, rcBounds, bmInfo.bmWidth, bmInfo.bmHeight);
													}
													if (hPal)
														SelectPalette(hDCNoFlicker, hPalOld, FALSE);
												}
											}
										}
										break;
									}
									m_bBackgroundDirty = FALSE;
								}
							}
							else
								FillSolidRect(hDCNoFlicker, rcBounds, m_crBackColor);
						}
						m_ffPageBackground.Paint(hDC, rcPaint.left, rcPaint.top);
					}
					break;
				}
	
				//
				// This is for scrolling
				//

				int nTool = 0;
				if (m_pCurChildBand)
					nTool = m_pCurChildBand->GetFirstTool();
				if (nTool)
				{
					//
					// If the first tool position is not equal to zero then 
					//		scrolling is active
					//

					if (bVertical)
					{
						m_sbsTopStyle = BTabs::eTopButton;
						m_sbsBottomStyle = BTabs::eInactiveBottomButton;
						m_nVertSlidingOffset = m_pCurChildBand->m_prcTools[nTool-1].bottom - m_pCurChildBand->m_prcTools[0].top;
					}
					else
					{
						m_sbsTopStyle = BTabs::eLeftButton;
						m_sbsBottomStyle = BTabs::eInactiveRightButton;
						m_nHorzSlidingOffset = m_pCurChildBand->m_prcTools[nTool-1].right - m_pCurChildBand->m_prcTools[0].left;
					}
				}
				else
				{
					if (bVertical)
					{
						m_sbsTopStyle = BTabs::eInactiveTopButton;
						m_sbsBottomStyle = BTabs::eInactiveBottomButton;
					}
					else
					{
						m_sbsTopStyle = BTabs::eInactiveLeftButton;
						m_sbsBottomStyle = BTabs::eInactiveRightButton;
					}
				}

				//
				// Normal Drawing stuff
				//

				m_crBtnDarkShadow = GetSysColor(COLOR_3DDKSHADOW);
				m_crBtnHighLight = VARIANT_FALSE == m_pBand->m_pBar->bpV1.m_vbXPLook ? m_pBand->m_pBar->m_crHighLight : m_pBand->m_pBar->m_crXPHighLight;
				m_crBtnShadow = VARIANT_FALSE == m_pBand->m_pBar->bpV1.m_vbXPLook ? m_pBand->m_pBar->m_crShadow : m_pBand->m_pBar->m_crXPShadow;
				m_crBtnText = m_pBand->m_pBar->m_crForeground;
				m_crBtnFace = VARIANT_FALSE == m_pBand->m_pBar->bpV1.m_vbXPLook ? m_pBand->m_pBar->m_crBackground : m_pBand->m_pBar->m_crXPBackground;

				rcTabs = rcPaint;
				BTGetRect(m_pBand->m_rcChildBandArea);
				int nSlidingEndLength = BTGetEndLen();
				if (bVertical)
				{
					rcPaint.top += (m_pBand->m_rcChildBandArea.top - nSlidingEndLength);
					rcPaint.bottom -= nSlidingEndLength;
				}
				else
				{
					rcPaint.left += (m_pBand->m_rcChildBandArea.left - nSlidingEndLength);
					rcPaint.right -= nSlidingEndLength;
				}
				CRect rc = rcPaint; 
				if (bVertical)
				{
					rc.right = rc.left + 1;
					FillSolidRect(hDC, rc, RGB(0,0,0));
					rc = rcPaint; 
					rc.bottom = rc.top;
					rc.top--;
					FillSolidRect(hDC, rc, RGB(0,0,0));
				}
				else
				{
					rc.right = rc.left;
					rc.left--;
					FillSolidRect(hDC, rc, RGB(0,0,0));
					rc = rcPaint; 
					rc.bottom = rc.top + 1;
					FillSolidRect(hDC, rc, RGB(0,0,0));
				}
			}
			break;
		}
		
		if (m_pCurChildBand)
			bResult = m_pCurChildBand->InsideDraw(hDC, rcPaint, bWrapFlag, bVertical, bVerticalPaint, bLeftHandle, bRightHandle, ptPaintOffset, bDrawControlsOrForms);
		assert(bResult);

		switch (m_pBand->bpV1.m_cbsChildStyle)
		{
		case ddCBSSlidingTabs:
			bResult = BTabs::Draw(m_pBand->m_pBar, hDC, rcTabs, rcPaint, ptPaintOffset);
			break;
		}
		return bResult;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
}

//
// DrawToolSlidingTabBackground
// 
// Draw the different types of backgrounds for sliding tabs
//

void CChildBands::DrawToolSlidingTabBackground(HDC hDC, const CRect& rcTotal, CRect	rcTool, BOOL bVertical)
{
	try
	{
		switch (ppV1.m_ChildBandBackgroundStyle)
		{
		case ddBSNormal:
			FillSolidRect(hDC, rcTotal, m_crBackColor);
 			break;

		case ddBSPicture:
		case ddBSGradient:
			{
				int nEndOffset = BTGetEndLen();
				if (bVertical)
					rcTool.Offset(-m_nHorzSlidingOffset - CBand::eBevelBorder, m_pBand->m_rcChildBandArea.top - nEndOffset - m_nVertSlidingOffset - CBand::eBevelBorder); 
				else
					rcTool.Offset(m_pBand->m_rcChildBandArea.left - nEndOffset - m_nHorzSlidingOffset - CBand::eBevelBorder, -m_nVertSlidingOffset - CBand::eBevelBorder); 
				m_ffPageBackground.Paint(hDC, rcTotal.left, rcTotal.top, rcTotal.Width(), rcTotal.Height(), rcTool.left, rcTool.top);
			}
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
// TranslateColors
//

HRESULT CChildBands::TranslateColors()
{
	m_bBackgroundDirty = TRUE;

	HRESULT hResult = OleTranslateColor(ppV1.m_ocPictureBackgroundMaskColor, NULL, &m_crPictureBackgroundMaskColor);
	if (FAILED(hResult))
		return hResult;

	hResult = OleTranslateColor(ppV1.m_ocShadowColor, NULL, &m_crShadowColor);
	if (FAILED(hResult))
		return hResult;

	hResult = OleTranslateColor(ppV1.m_ocHighLightColor, NULL, &m_crHighLightColor);
	if (FAILED(hResult))
		return hResult;

	hResult = OleTranslateColor(ppV1.m_ocToolForeColor, NULL, &m_crToolForeColor);
	if (FAILED(hResult))
		return hResult;

	hResult = OleTranslateColor(ppV1.m_ocForeColor, NULL, &m_crForeColor);
	if (FAILED(hResult))
		return hResult;

	hResult = OleTranslateColor(ppV1.m_ocBackColor, NULL, &m_crBackColor);
	if (FAILED(hResult))
		return hResult;

	hResult = OleTranslateColor(ppV1.m_ocGradientEndColor, NULL, &m_crGradientEndColor);
	if (FAILED(hResult))
		return hResult;
	return NOERROR;
}

//
// ParentWindowedTools
//
// Do not check for IsWindow here it is check in the Tool::ReparenthWndTools function
//

void CChildBands::ParentWindowedTools(HWND hWndParent)
{
	int nBandCount = m_aChildBands.GetSize();
	for (int nBand = 0; nBand < nBandCount; nBand++)
		m_aChildBands.GetAt(nBand)->ParentWindowedTools(hWndParent);
}

//
// HideWindowedTools
//

void CChildBands::HideWindowedTools()
{
	int nBandCount = m_aChildBands.GetSize();
	for (int nBand = 0; nBand < nBandCount; nBand++)
		m_aChildBands.GetAt(nBand)->HideWindowedTools();
}

//
// TabbedWindowedTools
//

void CChildBands::TabbedWindowedTools(BOOL bHide)
{
	int nBandCount = m_aChildBands.GetSize();
	for (int nBand = 0; nBand < nBandCount; nBand++)
		m_aChildBands.GetAt(nBand)->TabbedWindowedTools(bHide);
}

//
// LayoutSlidingTabs 
//

BOOL CChildBands::LayoutSlidingTabs(BOOL bVertical, CRect rcArea)
{
	int nSlidingHeight = BTGetFontHeight();
	
	CRect rcInside = rcArea;
	rcInside.Offset(-rcInside.left, -rcInside.top);
	CRect rcTab = rcInside;

	int nCount = BTGetTabCount();
	delete [] m_pTabs;
	if (nCount > 0)
		m_pTabs = new Tabs[nCount];
	else
		m_pTabs = NULL;

	m_nEndLength = 0;
	BOOL bFirstEndTab = TRUE;
	int nEnd;
	CTool* pTool;
	int nOffset;
	int nTool;
	int nTab = 0;
	int nToolCount = 0;


	if (m_pCurChildBand)
		nToolCount = m_pCurChildBand->m_pTools->GetVisibleToolCount();

	if (bVertical)
	{
		//
		// LayoutTabs
		//

		rcTab.bottom = rcTab.top + nSlidingHeight;
		for (nTab = 0; nTab < nCount; nTab++)
		{
			if (VARIANT_FALSE == m_aChildBands.GetAt(nTab)->bpV1.m_vbVisible)
				continue;

			if (nTab <= tpV1.m_nCurTab)
			{
				m_pTabs[nTab].m_rcTab = rcTab;
				rcTab.Offset(0, nSlidingHeight+1);
				if (nTab == tpV1.m_nCurTab)
					m_pTabs[nTab].m_eButtonInfo = Tabs::eTopOrLeft;
			}
			else
			{
				m_pTabs[nTab].m_eButtonInfo = Tabs::eBottomOrRight;
				break;
			}
		}
		nEnd = nTab;
		rcTab = rcInside;
		rcTab.top = rcTab.bottom - nSlidingHeight;
		for (nTab = nCount - 1; nTab >= nEnd; nTab--)
		{
			if (VARIANT_FALSE == m_aChildBands.GetAt(nTab)->bpV1.m_vbVisible)
				continue;

			m_nEndLength += (nSlidingHeight + 1);
			m_pTabs[nTab].m_rcTab = rcTab;
			rcTab.Offset(0, -(nSlidingHeight+1));
		}

		//
		// Adjusting tools Width
		//
		
		for (nTool = 0; nTool < nToolCount; nTool++)
		{
			pTool = m_pCurChildBand->m_pTools->GetVisibleTool(nTool);

			if (ddTTForm == pTool->tpV1.m_ttTools)
				continue;

			switch(abs(pTool->tpV1.m_taTools % 3))
			{
			case CTool::eHTALeft:
				nOffset = 0;
				break;

			case CTool::eHTACenter:
				nOffset = (rcInside.Width() - pTool->m_sizeTool.cx) / 2;
				break;

			case CTool::eHTARight:
				nOffset = rcInside.Width() - pTool->m_sizeTool.cx;
				break;

			default:
				assert(FALSE);
				break;
			}
			if (VARIANT_FALSE == m_pCurChildBand->bpV1.m_vbWrapTools)
				m_pCurChildBand->m_prcTools[nTool].Offset(nOffset, 0);
		}
	}
	else
	{
		//
		// LayoutTabs
		//

		rcTab.right = rcTab.left + nSlidingHeight;
		for (nTab = 0; nTab < nCount; nTab++)
		{
			if (VARIANT_FALSE == m_aChildBands.GetAt(nTab)->bpV1.m_vbVisible)
				continue;

			if (nTab <= tpV1.m_nCurTab)
			{
				m_pTabs[nTab].m_rcTab = rcTab;
				rcTab.Offset(nSlidingHeight+1, 0);
				if (nTab == tpV1.m_nCurTab)
					m_pTabs[nTab].m_eButtonInfo = Tabs::eTopOrLeft;
			}
			else
			{
				m_pTabs[nTab].m_eButtonInfo = Tabs::eBottomOrRight;
				break;
			}
		}
		nEnd = nTab;
		rcTab = rcInside;
		rcTab.left = rcTab.right - nSlidingHeight;
		for (nTab = nCount - 1; nTab >= nEnd; nTab--)
		{
			if (VARIANT_FALSE == m_aChildBands.GetAt(nTab)->bpV1.m_vbVisible)
				continue;

			m_nEndLength += (nSlidingHeight + 1);
			m_pTabs[nTab].m_rcTab = rcTab;
			rcTab.Offset(-(nSlidingHeight+1), 0);
		}
		
		//
		// Adjusting tools Height
		//
		
		for (nTool = 0; nTool < nToolCount; nTool++)
		{
			pTool = m_pCurChildBand->m_pTools->GetVisibleTool(nTool);
			
			if (ddTTForm == pTool->tpV1.m_ttTools)
				continue;

			switch(abs(pTool->tpV1.m_taTools % 3))
			{
			case CTool::eHTALeft:
				nOffset = 0;
				break;

			case CTool::eHTACenter:
				nOffset = (rcInside.Height() - pTool->m_sizeTool.cy) / 2;
				break;

			case CTool::eHTARight:
				nOffset = rcInside.Height() - pTool->m_sizeTool.cy;
				break;

			default:
				assert(FALSE);
				break;
			}
			if (VARIANT_FALSE == m_pCurChildBand->bpV1.m_vbWrapTools)
				m_pCurChildBand->m_prcTools[nTool].Offset(0, nOffset);
		}
	}
	return TRUE;
}

//
// BTCalcEndLen
//

int CChildBands::BTCalcEndLen()
{
	m_nEndLength = 0;
	int nSlidingHeight = BTGetFontHeight();
	int nCount = BTGetTabCount();
	for (int nTab = 0; nTab < nCount; nTab++)
	{
		if (VARIANT_FALSE == m_aChildBands.GetAt(nTab)->bpV1.m_vbVisible)
			continue;

		if (nTab > tpV1.m_nCurTab)
			m_nEndLength += (nSlidingHeight + 1);
	}
	return m_nEndLength;
}

