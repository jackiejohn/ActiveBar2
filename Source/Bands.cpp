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
#include <stddef.h>       // for offsetof()
#include "Support.h"
#include "Resource.h"
#include "Debug.h"     
#include "Bar.h"
#include "ChildBands.h"
#include "Band.h"
#include "Errors.h"
#include "MiniWin.h"
#include "Bands.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

//{OLE VERBS}
//{OLE VERBS}
//{OBJECT CREATEFN}
IUnknown *CreateFN_CBands(IUnknown *pUnkOuter)
{
	return (IUnknown *)(IBands *)new CBands();
}

DEFINE_AUTOMATIONOBJECT(&CLSID_Bands,Bands,_T("Bands Class"),CreateFN_CBands,2,0,&IID_IBands,_T(""));
void *CBands::objectDef=&BandsObject;
CBands *CBands::CreateInstance(IUnknown *pUnkOuter)
{
	return new CBands();
}
//{OBJECT CREATEFN}
CBands::CBands()
	: m_refCount(1)
{
	InterlockedIncrement(&g_cLocks);
// {BEGIN INIT}
// {END INIT}
}

CBands::~CBands()
{
	InterlockedDecrement(&g_cLocks);
	RemoveAll();
// {BEGIN CLEANUP}
// {END CLEANUP}
}

//{DEF IUNKNOWN MEMBERS}
STDMETHODIMP CBands::QueryInterface(REFIID riid, void **ppvObjOut)
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
	if (DO_GUIDS_MATCH(riid,IID_IBands))
	{
		*ppvObjOut=(void *)(IBands *)this;
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
STDMETHODIMP_(ULONG) CBands::AddRef()
{
	return ++m_refCount;
}
STDMETHODIMP_(ULONG) CBands::Release()
{
	if (0!=--m_refCount)
		return m_refCount;
	delete this;
	return 0;
}
//{ENDDEF IUNKNOWN MEMBERS}
STDMETHODIMP CBands::MapPropertyToCategory( DISPID dispid,  PROPCAT*ppropcat)
{
	return E_FAIL;
}
STDMETHODIMP CBands::GetCategoryName( PROPCAT propcat,  LCID lcid,  BSTR*pbstrName)
{
	return E_FAIL;
}
STDMETHODIMP CBands::_NewEnum( IUnknown **retval)
{
	*retval = 0;
	VARIANT* pColl;
	int cItems = m_aBands.GetSize();
	if (NULL != cItems)
	{
		if (! (pColl = (VARIANT*)HeapAlloc(g_hHeap, 0, cItems * sizeof(VARIANT))))
			return E_OUTOFMEMORY;
	}
	else
		pColl = NULL;

	for (int cnt=0;cnt<cItems;++cnt)
	{
		(pColl+cnt)->vt=VT_DISPATCH;
		(pColl+cnt)->pdispVal=(IBand*)m_aBands.GetAt(cnt);
	}
	IEnumX* pEnum=(IEnumX*)new CEnumX(IID_IEnumVARIANT,
									  cItems, 
									  sizeof(VARIANT), 
									  pColl, 
									  CopyDISPVARIANT);
	if (!pEnum)
        return E_OUTOFMEMORY;
	*retval=pEnum;
	return NOERROR;
}

STDMETHODIMP CBands::InterfaceSupportsErrorInfo( REFIID riid)
{
    return S_OK;
}
// IDispatch members

STDMETHODIMP CBands::GetTypeInfoCount( UINT *pctinfo)
{
	*pctinfo = 1; 
	return NOERROR; 
}
STDMETHODIMP CBands::GetTypeInfo( UINT itinfo, LCID lcid, ITypeInfo **pptinfo)
{
	*pptinfo=GetObjectTypeInfoEx(lcid,IID_IBands);
	if ((*pptinfo)==NULL)
		return E_FAIL;
	return NOERROR;
}
STDMETHODIMP CBands::GetIDsOfNames( REFIID riid, OLECHAR **rgszNames, UINT cNames, ULONG lcid, long *rgdispid)
{
	HRESULT hr;
	ITypeInfo *pTypeInfo;
	if (!(riid==IID_NULL))
        return E_INVALIDARG;
	pTypeInfo=GetObjectTypeInfoEx(lcid,IID_IBands);
	if (pTypeInfo==NULL)
		return E_FAIL;
	hr=pTypeInfo->GetIDsOfNames(rgszNames, cNames, rgdispid);
	pTypeInfo->Release();
	return hr;
}
STDMETHODIMP CBands::Invoke( long dispidMember, REFIID riid, ULONG lcid, USHORT wFlags, DISPPARAMS *pdispparams, VARIANT *pvarResult, EXCEPINFO *pexcepinfo, UINT *puArgErr)
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
	pTypeInfo=GetObjectTypeInfoEx(lcid,IID_IBands);
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

// IBands members
int CBands::GetPosOfItem(VARIANT *Index)
{
	int pos=-1;
	switch(Index->vt)
	{
	case VT_I2:
		pos = Index->iVal;
		break;
	case VT_I2|VT_BYREF:
		pos = *(Index->piVal);
		break;
	case VT_I4:
		pos=Index->lVal;
		break;
	case VT_I4|VT_BYREF:
		pos=*(Index->plVal);
		break;
	default:
		{
			CBand *pBand;
			BSTR searchName;
			VARIANT vstr;
			VariantInit(&vstr);
			if (Index->vt==VT_BSTR)
				searchName=Index->bstrVal;
			else
			{
				if (FAILED(VariantChangeType(&vstr,Index,VARIANT_NOVALUEPROP,VT_BSTR)))
					return -1;
				searchName=vstr.bstrVal;
			}

			int nElem=m_aBands.GetSize();
			for (int cnt = 0; cnt < nElem; ++cnt)
			{
				pBand = m_aBands.GetAt(cnt);
				if (NULL != pBand->m_bstrName && NULL != searchName)
				{
					if (0 == wcscmp(pBand->m_bstrName, searchName))
					{
						VariantClear(&vstr);
						return cnt;
					}
				}
			}
			VariantClear(&vstr);
		}
	}

	VARIANT v;
	VariantInit(&v);
	HRESULT hRes=VariantChangeType(&v,Index,VARIANT_NOVALUEPROP,VT_I4);
	if (FAILED(hRes))
		return -1;

	pos=v.lVal;
	return pos;
}

STDMETHODIMP CBands::Item( VARIANT *Index, Band **retval)
{
	int pos=GetPosOfItem(Index);
	if (pos < 0 || pos > m_aBands.GetSize()-1)
	{
		m_pBar->m_theErrorObject.SendError(IDERR_BADBANDSINDEX, 0);
		return CUSTOM_CTL_SCODE(IDERR_BADBANDSINDEX);
	}

	*retval=(Band *)m_aBands.GetAt(pos);
	((IUnknown *)(*retval))->AddRef();

	return NOERROR;
}

STDMETHODIMP CBands::Count( short *retval)
{
	*retval=m_aBands.GetSize();
	return NOERROR;
}

STDMETHODIMP CBands::Add(BSTR name, Band** retval)
{
	*retval = NULL;
	VARIANT vName;
	vName.vt = VT_BSTR;
	vName.bstrVal = name;
	int nPos = GetPosOfItem(&vName);
	if (nPos > -1)
	{
		m_pBar->m_theErrorObject.SendError(IDERR_DUPLICATEBANDNAME, name);
		return CUSTOM_CTL_SCODE(IDERR_DUPLICATEBANDNAME);
	}

	CBand* pNewBand = CBand::CreateInstance(NULL);
	if (NULL == pNewBand)
		return E_OUTOFMEMORY;

	if (m_pBar && m_pBar->m_pDesigner)
		pNewBand->bpV1.m_vbDesignerCreated = TRUE;

	pNewBand->SetOwner(m_pBar, TRUE);
	HRESULT hResult = pNewBand->put_Caption(name);
	
	// Have to give it a name before you add it to the band collection.
	hResult = pNewBand->put_Name(name);

	hResult = m_aBands.Add(pNewBand);

	if (FAILED(hResult))
	{
		pNewBand->Release();
		return hResult;
	}

	pNewBand->AddRef();
	pNewBand->OnChildBandChanged(0);
	*retval = (Band*)pNewBand;
	if (m_pBar->m_pDesignerNotify)
		m_pBar->m_pDesignerNotify->BandChanged((LPDISPATCH)(IBand*)pNewBand, (LPDISPATCH)NULL, ddBandAdded);
	return NOERROR;
}

STDMETHODIMP CBands::Remove( VARIANT *Index)
{
	int nItemPosition = GetPosOfItem(Index);
	if (nItemPosition < 0 || nItemPosition > m_aBands.GetSize() - 1)
	{
		m_pBar->m_theErrorObject.SendError(IDERR_BADBANDSINDEX,NULL);
		return CUSTOM_CTL_SCODE(IDERR_BADBANDSINDEX);
	}
	CBand* pBand = m_aBands.GetAt(nItemPosition);
	pBand->ShutdownFloatWin();

	if (m_pBar->m_pDesignerNotify)
		m_pBar->m_pDesignerNotify->BandChanged((LPDISPATCH)(IBand*)pBand, NULL, ddBandDeleted);

	pBand->Release();
	m_aBands.RemoveAt(nItemPosition);
	m_pBar->InternalRecalcLayout();
	return NOERROR;
}

HRESULT CBands::RemoveEx(CBand* pBand)
{
	int nBand = 0;
	int nElem = m_aBands.GetSize();
	for (nBand = 0; nBand < nElem; nBand++)
	{
		if (m_aBands.GetAt(nBand) == pBand)
			break;
	}
	if (nBand == nElem)
		return S_FALSE;

	if (m_pBar->m_pDesignerNotify)
		m_pBar->m_pDesignerNotify->BandChanged((LPDISPATCH)(IBand*)pBand, NULL, ddBandDeleted);

	pBand->Release();
	m_aBands.RemoveAt(nBand);

	return S_OK;
}

STDMETHODIMP CBands::RemoveAll()
{
	long nRefCnt;
	CBand* pBand;
	if (m_pBar)
		m_pBar->StatusBand(NULL);
	int nCount = m_aBands.GetSize();
	for (int nBand = 0; nBand < nCount; nBand++)
	{
		pBand = m_aBands.GetAt(nBand);
		if (m_pBar->m_pDesignerNotify)
			m_pBar->m_pDesignerNotify->BandChanged((LPDISPATCH)(IBand*)pBand, NULL, ddBandAdded);
		nRefCnt = pBand->Release();
	}
	m_aBands.RemoveAll();
	return NOERROR;
}

STDMETHODIMP CBands::InsertBand( int nIndex, IBand* pBand)
{
	((CBand*)pBand)->SetOwner(m_pBar, TRUE);

	if (m_pBar && m_pBar->m_pDesigner)
		static_cast<CBand*>(pBand)->bpV1.m_vbDesignerCreated = VARIANT_TRUE;

	HRESULT hResult;
	if (-1 == nIndex)
		hResult = m_aBands.Add(static_cast<CBand*>(pBand));
	else
		hResult = m_aBands.InsertAt(nIndex, pBand);

	if (FAILED(hResult))
		return hResult;

	if (m_pBar->m_pDesignerNotify)
		m_pBar->m_pDesignerNotify->BandChanged((LPDISPATCH)pBand, NULL, ddBandAdded);
	pBand->AddRef();
	return NOERROR;
}

//
// CheckForStatusBand
//

BOOL CBands::CheckForStatusBand(BandTypes btType)
{
	CBand* pBand;
	int nCount = m_aBands.GetSize();
	for (int nBand = 0; nBand < nCount; nBand++)
	{
		pBand = m_aBands.GetAt(nBand);
		switch (m_aBands.GetAt(nBand)->bpV1.m_btBands)
		{
		case ddBTStatusBar:
			if (ddBTStatusBar == pBand->bpV1.m_btBands && ddBTStatusBar == btType)
				return TRUE;
			break;
		}
	}
	return FALSE;
}

// 
// Exchange
//

HRESULT CBands::Exchange(IStream* pStream, VARIANT_BOOL vbSave, BSTR bstrBandName, short nBandCount)
{
	try
	{
		int nBand = 0;
		BSTR bstrBandNameStream = NULL;
		ULARGE_INTEGER nBeginPosition;
		nBeginPosition.LowPart = nBeginPosition.HighPart = 0;
		
		LARGE_INTEGER  nEndOffset;
		nEndOffset.LowPart = nEndOffset.HighPart = 0;

		LARGE_INTEGER  nOffset;
		nOffset.LowPart = nOffset.HighPart = 0;
		
		LARGE_INTEGER  nTempOffset;
		nTempOffset.LowPart = nTempOffset.HighPart = 0;

		HRESULT hResult = S_OK;

		if (bstrBandName && *bstrBandName)
		{
			//
			// Band Specific
			//

			CBand* pBand = NULL;
			if (VARIANT_TRUE == vbSave)
			{
				//
				// Saving
				//

				pBand = m_pBar->FindBand(bstrBandName);
				if (NULL == pBand)
					return E_FAIL;

				//
				// Save the count of bands
				//
				
				nBandCount = 1;
				hResult = pStream->Write(&nBandCount, sizeof(nBandCount), NULL);
				if (FAILED(hResult))
					return hResult;

				//
				// Save a place for the end of the bands layout to go
				//

				hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
				if (FAILED(hResult))
					return hResult;

				nEndOffset.LowPart = nBeginPosition.LowPart;

				hResult = pStream->Write(&nBeginPosition, sizeof(LARGE_INTEGER), NULL);
				if (FAILED(hResult))
					return hResult;

				//
				// Save the band name
				//

				hResult = StWriteBSTR(pStream, pBand->m_bstrName);
				if (FAILED(hResult))
					return hResult;

				//
				// Get and save the offset to this band
				//

				hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
				
				nTempOffset.LowPart = nBeginPosition.LowPart + sizeof(LARGE_INTEGER);

				hResult = pStream->Write(&nTempOffset, sizeof(LARGE_INTEGER), NULL);
				if (FAILED(hResult))
					return hResult;

				//
				// Save the band's information
				//

				hResult = pBand->Exchange(pStream, vbSave);
				if (FAILED(hResult))
					return hResult;

				//
				// Update the end of bands offset
				//
				
				hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
				if (FAILED(hResult))
					return hResult;

				nTempOffset.LowPart = nBeginPosition.LowPart;

				hResult = pStream->Seek(nEndOffset, STREAM_SEEK_SET, &nBeginPosition);
				if (FAILED(hResult))
					return hResult;

				hResult = pStream->Write(&nTempOffset, sizeof(LARGE_INTEGER), NULL);
				if (FAILED(hResult))
					return hResult;

				//
				// Moving to the end of the bands
				//
				
				hResult = pStream->Seek(nTempOffset, STREAM_SEEK_SET, &nBeginPosition);
			}
			else
			{
				//
				// Loading
				//

				//
				// Get the band count
				//

				short nCount;
				hResult = pStream->Read(&nCount, sizeof(nCount), NULL);
				if (FAILED(hResult))
					return hResult;

				//
				// Get the end of the band's offset
				//

				hResult = pStream->Read(&nEndOffset, sizeof(LARGE_INTEGER), NULL);
				if (FAILED(hResult))
					return hResult;

				//
				// Read the band names and their offsets
				//

				for (int nBand = 0; nBand < nCount; nBand++)
				{
					hResult = StReadBSTR(pStream, bstrBandNameStream);
					if (FAILED(hResult))
						return hResult;

					hResult = pStream->Read(&nTempOffset, sizeof(LARGE_INTEGER), NULL);
					if (FAILED(hResult))
						return hResult;

					if (0 == wcscmp(bstrBandNameStream, bstrBandName))
					{
						SysFreeString(bstrBandNameStream);

						//
						// We found the band so move to it's offset
						//

						hResult = pStream->Seek(nTempOffset, STREAM_SEEK_SET, &nBeginPosition);
						if (FAILED(hResult))
							return hResult;

						pBand = CBand::CreateInstance(NULL);
						if (NULL == pBand)
							return E_OUTOFMEMORY;

						//
						// Read the bands information
						//
						
						pBand->SetOwner(m_pBar, TRUE);

						hResult = pBand->Exchange(pStream, vbSave);
						if (FAILED(hResult))
							return hResult;

						CBand* pBandOld = m_pBar->FindBand(bstrBandName);
						if (NULL == pBand)
							return E_FAIL;

						hResult = RemoveEx(pBandOld);

						hResult = m_aBands.Add(pBand);
						if (FAILED(hResult))
						{
							pBand->Release();
							return hResult;
						}

						CTool* pTool;
						CTool* pMainTool;
						int nToolCount = pBand->m_pTools->GetToolCount();
						for (int nTool = 0; nTool < nToolCount; nTool++)
						{
							pTool = pBand->m_pTools->GetTool(nTool);
							pMainTool = m_pBar->m_pTools->FindToolId(pTool->tpV1.m_nToolId);
							if (pMainTool)
							{
								// Set Shared Properties
								pTool->tpV1.m_vbEnabled = pMainTool->tpV1.m_vbEnabled;
								pTool->tpV1.m_vbChecked = pMainTool->tpV1.m_vbChecked;
								pTool->tpV1.m_vbVisible = pMainTool->tpV1.m_vbVisible;
							}
						}

						//
						// Move to the end of the bands
						//

						hResult = pStream->Seek(nEndOffset, STREAM_SEEK_SET, &nBeginPosition);
						if (FAILED(hResult))
							return hResult;
						return NOERROR;
					}
					SysFreeString(bstrBandNameStream);
				}
				return E_FAIL;
			}
		}
		else
		{
			//
			// All Bands
			//

			if (VARIANT_TRUE == vbSave)
			{
				//
				// Saving
				//

				//
				// Read the count of bands
				//

				hResult = pStream->Write(&nBandCount, sizeof(nBandCount), NULL);
				if (FAILED(hResult))
					return hResult;
				
				//
				// Create a place of the end of bands offset to go
				//

				hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
				if (FAILED(hResult))
					return hResult;

				nEndOffset.LowPart = nBeginPosition.LowPart;

				hResult = pStream->Write(&nBeginPosition, sizeof(LARGE_INTEGER), NULL);
				if (FAILED(hResult))
					return hResult;

				//
				// Save the band names and a place for their offsets
				//

				LARGE_INTEGER* pBandNames = new LARGE_INTEGER[nBandCount];
				for (nBand = 0; nBand < nBandCount; nBand++)
				{
					hResult = StWriteBSTR(pStream, m_aBands.GetAt(nBand)->m_bstrName);
					if (FAILED(hResult))
						return hResult;

					hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
					
					pBandNames[nBand].LowPart = nBeginPosition.LowPart;
					pBandNames[nBand].HighPart = 0;

					hResult = pStream->Write(&nBeginPosition, sizeof(LARGE_INTEGER), NULL);
					if (FAILED(hResult))
						return hResult;
				}

				//
				// Saving the bands information and putting their offsets into an array
				//

				LARGE_INTEGER* pBandOffsets = new LARGE_INTEGER[nBandCount];
				for (nBand = 0; nBand < nBandCount; nBand++)
				{
					hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
					
					pBandOffsets[nBand].LowPart = nBeginPosition.LowPart;
					pBandOffsets[nBand].HighPart = 0;

					hResult = m_aBands.GetAt(nBand)->Exchange(pStream, vbSave);
					if (FAILED(hResult))
						return hResult;
				}

				//
				// Moving back to the place where we made room to the end of the bands offsets
				//

				hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
				if (FAILED(hResult))
					return hResult;

				nTempOffset.LowPart = nBeginPosition.LowPart;

				hResult = pStream->Seek(nEndOffset, STREAM_SEEK_SET, &nBeginPosition);
				if (FAILED(hResult))
					return hResult;
				
				hResult = pStream->Write(&nTempOffset, sizeof(LARGE_INTEGER), NULL);
				if (FAILED(hResult))
					return hResult;

				//
				// Updatings the bands offsets
				//

				for (nBand = 0; nBand < nBandCount; nBand++)
				{
					hResult = pStream->Seek(pBandNames[nBand], STREAM_SEEK_SET, &nBeginPosition);
					if (FAILED(hResult))
						return hResult;
					
					hResult = pStream->Write(&pBandOffsets[nBand], sizeof(LARGE_INTEGER), NULL);
					if (FAILED(hResult))
						return hResult;
				}
				
				//
				// Moving to the end of the bands
				//
				
				hResult = pStream->Seek(nTempOffset, STREAM_SEEK_SET, &nBeginPosition);

				delete [] pBandNames;
				delete [] pBandOffsets;
			}
			else
			{
				//
				// Loading
				//

				CBand* pBand;
				short nCount;

				//
				// Read the band count
				//
				
				hResult = pStream->Read(&nCount, sizeof(nCount), NULL);
				if (FAILED(hResult))
					return hResult;

				//
				// If the read in Band count differs form the one read earlier we have a problem
				//

				if (nCount != nBandCount)
				{
					if (!m_pBar->AmbientUserMode())
					{
						DDString strMsg;
						strMsg.Format(IDS_ERR_BANDLIMIT, nBandCount);
						int nResult = m_pBar->MessageBox(strMsg, MB_YESNO);
						if (IDNO == nResult)
							return E_FAIL;
					}
					else
						return E_FAIL;
				}

				//
				// Reading the end of the bands
				//

				hResult = pStream->Read(&nBeginPosition, sizeof(LARGE_INTEGER), NULL);
				if (FAILED(hResult))
					return hResult;

				//
				// Reading the band names and their offsets
				//

				for (nBand = 0; nBand < nCount; nBand++)
				{
					hResult = StReadBSTR(pStream, bstrBandNameStream);
					if (FAILED(hResult))
						return hResult;

					hResult = pStream->Read(&nBeginPosition, sizeof(LARGE_INTEGER), NULL);
					if (FAILED(hResult))
						return hResult;
				}
				SysFreeString(bstrBandNameStream);

				//
				// Reading the bands
				//

				for (nBand = 0; nBand < nCount; nBand++)
				{
					pBand = CBand::CreateInstance(NULL);
					if (NULL == pBand)
						return E_OUTOFMEMORY;
					
					pBand->SetOwner(m_pBar, TRUE);
					hResult = m_aBands.Add(pBand);
					if (FAILED(hResult))
					{
						pBand->Release();
						return hResult;
					}

					hResult = pBand->Exchange(pStream, vbSave);
					if (FAILED(hResult))
						return hResult;
				}
			}
		}
		return hResult;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
}

//
// ExchangeConfig
//

HRESULT CBands::ExchangeConfig(IStream* pStream, VARIANT_BOOL vbSave, short nBandCount)
{
	try
	{
		int nBand = 0;
		HRESULT hResult = S_OK;
		//
		// All Bands
		//

		if (VARIANT_TRUE == vbSave)
		{
			//
			// Saving
			//

			hResult = pStream->Write(&nBandCount, sizeof(nBandCount), NULL);
			if (FAILED(hResult))
				return hResult;
			
			//
			// Saving all of the current band names
			//

			for (int nBand = 0; nBand < nBandCount; nBand++)
			{
				hResult = StWriteBSTR(pStream, m_aBands.GetAt(nBand)->m_bstrName);
				if (FAILED(hResult))
					return hResult;
			}

			//
			// Saving the information for all of the bands
			//

			for (nBand = 0; nBand < nBandCount; nBand++)
			{
				hResult = m_aBands.GetAt(nBand)->ExchangeConfig(pStream, vbSave);
				if (FAILED(hResult))
					return hResult;
			}
		}
		else
		{
			//
			// Loading
			//

			CBand* pBand;
			short nCount;
			hResult = pStream->Read(&nCount, sizeof(nCount), NULL);
			if (FAILED(hResult))
				return hResult;

			if (nCount != nBandCount)
			{
				if (!m_pBar->AmbientUserMode())
				{
					DDString strMsg;
					strMsg.Format(IDS_ERR_BANDLIMIT, nBandCount);
					int nResult = m_pBar->MessageBox(strMsg, MB_YESNO);
					if (IDNO == nResult)
						return E_FAIL;
				}
				else
					return E_FAIL;
			}

			//
			// Loading all of the Band Names
			//

			BSTR* pbstrBandName = new BSTR[nCount];
			for (nBand = 0; nBand < nCount; nBand++)
			{
				pbstrBandName[nBand] = NULL;
				hResult = StReadBSTR(pStream, pbstrBandName[nBand]);
				if (FAILED(hResult))
					return hResult;
			}

			//
			// Removing the old bands that are not in this configuration and marking the old bands so
			// we can tell which ones are new
			//

			BOOL* pbBandOld = new BOOL[nCount];
			for (nBand = 0; nBand < nCount; nBand++)
				pbBandOld[nBand] = FALSE;

			BSTR bstrBandName;
			BOOL bFound;
			int nBandCount = m_aBands.GetSize();
			for (nBand = 0; nBand < nBandCount; nBand++)
			{
				bFound = FALSE;
				bstrBandName = m_aBands.GetAt(nBand)->m_bstrName;
				for (int nBand2 = 0; nBand2 < nCount; nBand2++)
				{
					if (0 == wcscmp(pbstrBandName[nBand2], bstrBandName))
					{
						pbBandOld[nBand2] = TRUE;
						bFound = TRUE;
						break;
					}
				}
				if (!bFound)
					RemoveEx(m_aBands.GetAt(nBand));
			}

			//
			// Updating the information
			//

			for (nBand = 0; nBand < nCount; nBand++)
			{
				if (!pbBandOld[nBand])
				{
					//
					// Adding the new bands
					//

					pBand = CBand::CreateInstance(NULL);
					if (NULL == pBand)
						return E_OUTOFMEMORY;
					
					pBand->SetOwner(m_pBar, TRUE);

					hResult = m_aBands.Add(pBand);
					if (FAILED(hResult))
					{
						pBand->Release();
						return hResult;
					}

					hResult = pBand->ExchangeConfig(pStream, vbSave);
					if (FAILED(hResult))
						return hResult;
				}
				else
				{
					//
					// Updating the Old Bands
					//

					pBand = m_pBar->FindBand(pbstrBandName[nBand]);

					hResult = pBand->ExchangeConfig(pStream, vbSave);
					if (FAILED(hResult))
						return hResult;
				}
			}
			for (nBand = 0; nBand < nCount; nBand++)
				SysFreeString(pbstrBandName[nBand]);
			delete [] pbstrBandName;
			delete [] pbBandOld;
		}
		return hResult;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
}

//
// ExchangeMenuUsageData
//

HRESULT CBands::ExchangeMenuUsageData(IStream* pStream, VARIANT_BOOL vbSave)
{
	HRESULT hResult;
	CBand* pBand;
	int nCount = m_aBands.GetSize();
	for (int nBand = 0; nBand < nCount; nBand++)
	{
		pBand = m_aBands.GetAt(nBand);
		hResult = pBand->ExchangeMenuUsageData(pStream, vbSave);
		if (FAILED(hResult))
			return hResult;
	}
	return NOERROR;
}

HRESULT CBands::ClearMenuUsageData()
{
	HRESULT hResult;
	CBand* pBand;
	int nCount = m_aBands.GetSize();
	for (int nBand = 0; nBand < nCount; nBand++)
	{
		pBand = m_aBands.GetAt(nBand);
		hResult = pBand->ClearMenuUsageData();
		if (FAILED(hResult))
			return hResult;
	}
	return NOERROR;
}

CBand* CBands::Find(BSTR bstrName)
{
	CBand* pBand;
	int nCount = m_aBands.GetSize();
	for (int nBand = 0; nBand < nCount; nBand++)
	{
		pBand = m_aBands.GetAt(nBand);
		
		if (NULL == pBand->m_bstrName && NULL == bstrName)
			return pBand;

		if (NULL == pBand->m_bstrName || NULL == bstrName)
			return NULL;

		if (0 == wcscmp(pBand->m_bstrName, bstrName))
			return pBand;
	}
	return NULL;
}

void CBands::RefreshFloating()
{
	CBand* pBand;
	int nCount = m_aBands.GetSize();
	for (int nBand = 0; nBand < nCount; nBand++)
	{
		pBand = m_aBands.GetAt(nBand);
		if (ddDAFloat == pBand->bpV1.m_daDockingArea && pBand->m_pFloat && pBand->m_pFloat->IsWindow())
			pBand->m_pFloat->InvalidateRect(NULL, FALSE);
	}
}

//{EVENT WRAPPERS}
//{EVENT WRAPPERS}
