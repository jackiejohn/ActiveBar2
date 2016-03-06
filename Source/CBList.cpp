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
#include "IpServer.h"
#include <stddef.h>       // for offsetof()
#include "Support.h"
#include "Resource.h"
#include "Debug.h"
#include "Utility.h"
#include "Tool.h"
#include "Bar.h"
#include "CBList.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

//
// ListItem
//

struct ListItem
{
	ListItem(BSTR bstrName)
		: m_bstrString(SysAllocString(bstrName))
	{
		m_dwItemData = 0;
	}

	ListItem()
		: m_bstrString(NULL)
	{
		m_dwItemData = 0;
	}

	~ListItem()
	{
		SysFreeString(m_bstrString);
	}

	HRESULT CopyTo(ListItem* pListItem);
	HRESULT Exchange(IStream* pStream, VARIANT_BOOL vbSave);

	DWORD m_dwItemData;
	BSTR  m_bstrString;
};

HRESULT ListItem::CopyTo(ListItem* pListItem)
{
	if (pListItem == this)
		return NOERROR;
	pListItem->m_dwItemData = m_dwItemData;
	SysFreeString(pListItem->m_bstrString);
	pListItem->m_bstrString = SysAllocString(m_bstrString);
	return NOERROR;
}

HRESULT ListItem::Exchange(IStream* pStream, VARIANT_BOOL vbSave)
{
	try
	{
		long nStreamSize;
		int nCount = 0;
		HRESULT hResult;
		if (VARIANT_TRUE == vbSave)
		{
			nStreamSize = GetStringSize(m_bstrString) + sizeof(m_dwItemData);
			hResult = pStream->Write(&nStreamSize, sizeof(nStreamSize), NULL);
			if (FAILED(hResult))
				return hResult;

			hResult = StWriteBSTR(pStream, m_bstrString);
			if (FAILED(hResult))
				return hResult;
					
			hResult = pStream->Write(&m_dwItemData, sizeof(m_dwItemData), NULL);
			if (FAILED(hResult))
				return hResult;
		}
		else
		{
			ULARGE_INTEGER nBeginPosition;
			LARGE_INTEGER  nOffset;

			hResult = pStream->Read(&nStreamSize, sizeof(nStreamSize), NULL);
			if (FAILED(hResult))
				return hResult;

			hResult = StReadBSTR(pStream, m_bstrString);
			if (FAILED(hResult))
				return hResult;

			nStreamSize -= GetStringSize(m_bstrString);
			if (nStreamSize <= 0)
				goto FinishedReading;

			hResult = pStream->Read(&m_dwItemData, sizeof(m_dwItemData), NULL);
			if (FAILED(hResult))
				return hResult;

			nStreamSize -= sizeof(m_dwItemData);
			if (nStreamSize <= 0)
				goto FinishedReading;

			nOffset.HighPart = 0;
			nOffset.LowPart = nStreamSize;
			hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
			if (FAILED(hResult))
				return hResult;
		}
FinishedReading:
		return NOERROR;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
}

//{OLE VERBS}
//{OLE VERBS}
//{OBJECT CREATEFN}
IUnknown *CreateFN_CBList(IUnknown *pUnkOuter)
{
	return (IUnknown *)(IComboList *)new CBList();
}

DEFINE_AUTOMATIONOBJECT(&CLSID_ComboList,ComboList,_T("ComboList Class"),CreateFN_CBList,2,0,&IID_IComboList,_T(""));
void *CBList::objectDef=&ComboListObject;
CBList *CBList::CreateInstance(IUnknown *pUnkOuter)
{
	return new CBList();
}
//{OBJECT CREATEFN}
CBList::CBList()
	: m_refCount(1)
	, m_hWndActive(NULL)
{
	InterlockedIncrement(&g_cLocks);
// {BEGIN INIT}
// {END INIT}
	m_bNeedsSort=FALSE;
	m_nNewIndex = -1;
}

CBList::~CBList()
{
	InterlockedDecrement(&g_cLocks);
// {BEGIN CLEANUP}
// {END CLEANUP}
	Cleanup();
}

CBList::CBListPropV1::CBListPropV1()
{
	m_nLines=0;
	m_nWidth=0;
	m_nStyle=ddCBSNormal;
	m_nListIndex=-1;
}

void CBList::Cleanup()
{
	int nElem = m_aItems.GetSize();
	for (int cnt = 0; cnt < nElem; ++cnt)
		delete m_aItems.GetAt(cnt);
	m_aItems.RemoveAll();
	m_bNeedsSort = FALSE;
}

//{DEF IUNKNOWN MEMBERS}
STDMETHODIMP CBList::QueryInterface(REFIID riid, void **ppvObjOut)
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
	if (DO_GUIDS_MATCH(riid,IID_IComboList))
	{
		*ppvObjOut=(void *)(IComboList *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	//{CUSTOM QI}
	//{ENDCUSTOM QI}
	*ppvObjOut = NULL;
	return E_NOINTERFACE;
}
STDMETHODIMP_(ULONG) CBList::AddRef()
{
	return ++m_refCount;
}
STDMETHODIMP_(ULONG) CBList::Release()
{
	if (0!=--m_refCount)
		return m_refCount;
	delete this;
	return 0;
}
//{ENDDEF IUNKNOWN MEMBERS}
STDMETHODIMP CBList::get_NewIndex(long *retval)
{
	*retval = m_nNewIndex;
	return NOERROR;
}
STDMETHODIMP CBList::CopyTo(IComboList **pComboList)
{
	CBList* pList = (CBList*)(*pComboList);

	if (pList == this)
		return NOERROR;

	pList->lpV1 = lpV1;
	
	pList->Cleanup();

	HRESULT hResult;
	ListItem* pListItem;
	int nCount = m_aItems.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		pListItem = new ListItem;
		if (NULL == pListItem)
			return E_OUTOFMEMORY;

		hResult = m_aItems.GetAt(nIndex)->CopyTo(pListItem);
		if (FAILED(hResult))
		{
			delete pListItem;
			return hResult;
		}
		hResult = pList->m_aItems.Add(pListItem);
		if (FAILED(hResult))
		{
			delete pListItem;
			return hResult;
		}
	}
	return NOERROR;
}
// IDispatch members

STDMETHODIMP CBList::GetTypeInfoCount( UINT *pctinfo)
{
	*pctinfo = 1; 
	return NOERROR; 
}
STDMETHODIMP CBList::GetTypeInfo( UINT itinfo, LCID lcid, ITypeInfo **pptinfo)
{
	*pptinfo=GetObjectTypeInfo(lcid,objectDef);
	if ((*pptinfo)==NULL)
		return E_FAIL;
	return NOERROR;
}
STDMETHODIMP CBList::GetIDsOfNames( REFIID riid, OLECHAR **rgszNames, UINT cNames, ULONG lcid, long *rgdispid)
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
STDMETHODIMP CBList::Invoke( long dispidMember, REFIID riid, ULONG lcid, USHORT wFlags, DISPPARAMS *pdispparams, VARIANT *pvarResult, EXCEPINFO *pexcepinfo, UINT *puArgErr)
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

//
// BandSortFunc
//

static int ListSortFunc(const void *arg1, const void *arg2 )
{
	ListItem* p1 = *((ListItem**)arg1);
	ListItem* p2 = *((ListItem**)arg2);

	if (NULL == p1->m_bstrString && NULL != p2->m_bstrString)
		return -1;

	if (NULL != p1->m_bstrString && NULL == p2->m_bstrString)
		return 1;

	if (NULL == p1->m_bstrString && NULL == p2->m_bstrString)
		return 0;

	return wcscmp(p1->m_bstrString, p2->m_bstrString);
}

//
// SortList
//

static void SortList(TypedArray<ListItem*>& faItems)
{
	qsort(faItems.GetData(), faItems.GetSize(), sizeof(void*), ListSortFunc);
}

// IComboList members

STDMETHODIMP CBList::AddItem( BSTR name)
{
	ListItem* pListItem = new ListItem(name);
	if (NULL == pListItem)
		return E_OUTOFMEMORY;

	HRESULT hResult = m_aItems.Add(pListItem, (int*)&m_nNewIndex);
	if (FAILED(hResult))
	{
		delete pListItem;
		return hResult;
	}

	if (ddCBSSortedReadOnly == lpV1.m_nStyle || ddCBSSorted == lpV1.m_nStyle)
		m_bNeedsSort=TRUE;

	if (IsWindow(m_hWndActive))
	{
		MAKE_TCHARPTR_FROMWIDE(szName, name);
		m_nNewIndex = SendMessage(m_hWndActive, CB_ADDSTRING, 0, (LPARAM)szName);
	}
	if (m_pTool && m_pTool->m_pBar)
		m_pTool->m_pBar->SetToolComboList(m_pTool->tpV1.m_nToolId, this);
	return NOERROR;
}
STDMETHODIMP CBList::Remove( VARIANT *Index)
{
	if (NULL == Index)
		return E_INVALIDARG;

	int i = ConvertVariantToIndex(Index);
	if (i == -1)
		return DISP_E_BADINDEX;

	delete m_aItems.GetAt(i);
	m_aItems.RemoveAt(i);
	
	if (IsWindow(m_hWndActive))
		SendMessage(m_hWndActive, CB_DELETESTRING, i, 0);
	if (m_pTool && m_pTool->m_pBar)
		m_pTool->m_pBar->SetToolComboList(m_pTool->tpV1.m_nToolId, this);
	return NOERROR;
}
STDMETHODIMP CBList::Item( VARIANT *Index, BSTR *retval)
{
	if (NULL == retval || NULL == Index)
		return E_INVALIDARG;
	if (m_bNeedsSort)
	{
		SortList(m_aItems);
		m_bNeedsSort = FALSE;
	}
	int nIndex = ConvertVariantToIndex(Index);
	if (nIndex == -1)
		return DISP_E_BADINDEX;

	*retval = SysAllocString(m_aItems.GetAt(nIndex)->m_bstrString);
	if (m_pTool && m_pTool->m_pBar)
		m_pTool->m_pBar->SetToolComboList(m_pTool->tpV1.m_nToolId, this);
	return NOERROR;
}
STDMETHODIMP CBList::Clear()
{
	Cleanup();
	if (IsWindow(m_hWndActive))
		SendMessage(m_hWndActive, CB_RESETCONTENT, 0, 0);
	if (m_pTool && m_pTool->m_pBar)
		m_pTool->m_pBar->SetToolComboList(m_pTool->tpV1.m_nToolId, this);
	m_pTool->put_Text(L"");
	lpV1.m_nListIndex = -1;
	return NOERROR;
}
STDMETHODIMP CBList::InsertItem(int index, BSTR name)
{
	if (index < 0 || index > m_aItems.GetSize())
		return DISP_E_BADINDEX;

	ListItem* pListItem = new ListItem(name);
	if (NULL == pListItem)
		return E_OUTOFMEMORY;

	m_aItems.InsertAt(index, pListItem);
	if (ddCBSSortedReadOnly == lpV1.m_nStyle || ddCBSSorted == lpV1.m_nStyle)
		m_bNeedsSort = TRUE;

	if (IsWindow(m_hWndActive))
	{
		if (index == m_aItems.GetSize())
			index = -1;

		MAKE_TCHARPTR_FROMWIDE(szName, name);
		m_nNewIndex = SendMessage(m_hWndActive, CB_INSERTSTRING, (WPARAM)index, (LPARAM)szName);
	}
	if (m_pTool && m_pTool->m_pBar)
		m_pTool->m_pBar->SetToolComboList(m_pTool->tpV1.m_nToolId, this);
	return NOERROR;
}
STDMETHODIMP CBList::Count( short *retval)
{
	if (NULL == retval)
		return E_INVALIDARG;
	*retval = m_aItems.GetSize();
	return NOERROR;
}
STDMETHODIMP CBList::put_ItemData( VARIANT Index,  long Data)
{
	int i = ConvertVariantToIndex(&Index);
	if (-1 == i)
		return DISP_E_BADINDEX;
	reinterpret_cast<ListItem*>(m_aItems.GetAt(i))->m_dwItemData = Data;
	if (m_pTool && m_pTool->m_pBar)
		m_pTool->m_pBar->SetToolComboList(m_pTool->tpV1.m_nToolId, this);
	return NOERROR;
}
STDMETHODIMP CBList::get_ItemData( VARIANT Index,  long*Data)
{
	if (NULL == Data)
		return E_INVALIDARG;

	int i=ConvertVariantToIndex(&Index);
	if (-1 == i)
		return DISP_E_BADINDEX;
	*Data = reinterpret_cast<ListItem*>(m_aItems.GetAt(i))->m_dwItemData;
	return NOERROR;
}

//
// Exchange
//

HRESULT CBList::Exchange(IStream* pStream, VARIANT_BOOL vbSave)
{
	try
	{
		long nStreamSize;
		short nSize, nSize2;
		int nCount = 0;
		HRESULT hResult;
		if (VARIANT_TRUE == vbSave)
		{
			nStreamSize = sizeof(nSize) + sizeof(lpV1);
			hResult = pStream->Write(&nStreamSize, sizeof(nStreamSize), NULL);
			if (FAILED(hResult))
				return hResult;

			nSize = sizeof(lpV1);
			hResult = pStream->Write(&nSize, sizeof(nSize), NULL);
			if (FAILED(hResult))
				return hResult;

			hResult = pStream->Write(&lpV1, nSize, NULL);
			if (FAILED(hResult))
				return hResult;

			nCount = m_aItems.GetSize();
			
			hResult = pStream->Write(&nCount, sizeof(nCount), NULL);
			if (FAILED(hResult))
				return hResult;
			
			for (int nIndex = 0; nIndex < nCount; nIndex++)
			{
				hResult = m_aItems.GetAt(nIndex)->Exchange(pStream, vbSave);
				if (FAILED(hResult))
					return hResult;
			}
		}
		else
		{
			ULARGE_INTEGER nBeginPosition;
			LARGE_INTEGER  nOffset;

			hResult = pStream->Read(&nStreamSize, sizeof(nStreamSize), NULL);
			if (FAILED(hResult))
				return hResult;

			hResult = pStream->Read(&nSize, sizeof(nSize), NULL);
			if (FAILED(hResult))
				return hResult;

			nStreamSize -= sizeof(nSize);
			if (nStreamSize <= 0)
				goto FinishedReading;

			nSize2 = sizeof(lpV1);
			hResult = pStream->Read(&lpV1, nSize < nSize2 ? nSize : nSize2, NULL);
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

			nOffset.HighPart = 0;
			nOffset.LowPart = nStreamSize;
			hResult = pStream->Seek(nOffset, STREAM_SEEK_CUR, &nBeginPosition);
			if (FAILED(hResult))
				return hResult;

FinishedReading:
			ListItem* pListItem;
			hResult = pStream->Read(&nCount, sizeof(nCount), NULL);
			for (int nIndex = 0; nIndex < nCount; nIndex++)
			{
				pListItem = new ListItem;
				if (NULL == pListItem)
					return E_OUTOFMEMORY;
				
				hResult = pListItem->Exchange(pStream, vbSave);
				if (FAILED(hResult))
				{
					delete pListItem;
					return hResult;
				}
			
				hResult = m_aItems.Add(pListItem);
				if (FAILED(hResult))
				{
					delete pListItem;
					return hResult;
				}
			}
			if (ddCBSSortedReadOnly == lpV1.m_nStyle || ddCBSSorted == lpV1.m_nStyle)
				m_bNeedsSort = TRUE;
		}
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
// ConvertVariantToIndex
//

int CBList::ConvertVariantToIndex(VARIANT *Index)
{
	int nPos = -1;
	switch (Index->vt)
	{
	case VT_I2:
		nPos = Index->iVal;
		break;

	case VT_I4:
		nPos = Index->lVal;
		break;

	default:
		{
			VARIANT v;
			VariantInit(&v);
			HRESULT hRes = VariantChangeType(&v, Index, VARIANT_NOVALUEPROP, VT_I4);
			if (FAILED(hRes))
				return nPos;
			nPos = v.lVal;
		}
		break;
	}
	if (nPos < 0 || nPos >= m_aItems.GetSize())
		return -1;
	return nPos;
}

//
// GetPosOfItem
//

int CBList::GetPosOfItem(BSTR strItem)
{
	if (NULL == strItem || NULL == *strItem)
		return -1;

	int nCount = m_aItems.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		if (0 == wcscmp(m_aItems.GetAt(nIndex)->m_bstrString, strItem))
			return nIndex;
	}
	return -1;
}

int CBList::Count()
{
	return m_aItems.GetSize();
}

BSTR CBList::GetName(int nIndex)
{
	if (m_bNeedsSort)
	{
		SortList(m_aItems);
		m_bNeedsSort = FALSE;
	}
	return m_aItems.GetAt(nIndex)->m_bstrString;
}

DWORD CBList::GetData(int nIndex)
{
	if (m_bNeedsSort)
	{
		SortList(m_aItems);
		m_bNeedsSort = FALSE;
	}
	return m_aItems.GetAt(nIndex)->m_dwItemData;
}
//{EVENT WRAPPERS}
//{EVENT WRAPPERS}
