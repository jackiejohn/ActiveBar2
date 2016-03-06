//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include <Math.h>
#include "Debug.h"
#include "Support.h"
#include "..\EventLog.h"
#include "GdiUtil.h"
#include "FDialog.h"
#include "Resource.h"
#include "DesignerPage.h"
#include "Globals.h"
#include "Dialogs.h"
#include "Browser.h"

extern HINSTANCE g_hInstance;

#define WM_SHOWEDIT WM_USER + 1000

//
// COleDispatchDriver
//

COleDispatchDriver::COleDispatchDriver(LPDISPATCH pDispatch)
	: m_pDispatch(pDispatch)
{
}

//
// InvokeHelperV
//

HRESULT COleDispatchDriver::InvokeHelperV(DISPID      dwDispID, 
										  WORD        wFlags,
										  VARTYPE     vtRet, 
										  VARIANT*    pvRet, 
										  const BYTE* pbParamInfo, 
										  va_list     argList)
{
	if (NULL == m_pDispatch)
	{
		TRACE(1, "Warning: attempt to call Invoke with NULL m_pDispatch!\n");
		return E_FAIL;
	}

	DISPPARAMS dispparams;
	memset(&dispparams, 0, sizeof dispparams);

	// determine number of arguments
	if (pbParamInfo)
		dispparams.cArgs = lstrlenA((LPCSTR)pbParamInfo);

	DISPID dispidNamed = DISPID_PROPERTYPUT;
	if (wFlags & (DISPATCH_PROPERTYPUT|DISPATCH_PROPERTYPUTREF))
	{
		assert(dispparams.cArgs > 0);
		dispparams.cNamedArgs = 1;
		dispparams.rgdispidNamedArgs = &dispidNamed;
	}

	if (0 != dispparams.cArgs)
	{
		// allocate memory for all VARIANT parameters
		VARIANT* pArg = new VARIANT[dispparams.cArgs];
		assert(pArg);   // should have thrown exception
		dispparams.rgvarg = pArg;
		memset(pArg, 0, sizeof(VARIANT) * dispparams.cArgs);

		// get ready to walk vararg list
		const BYTE* pb = pbParamInfo;
		pArg += dispparams.cArgs - 1;   // params go in opposite order

		while (*pb != 0)
		{
			assert(pArg >= dispparams.rgvarg);
			pArg->vt = *pb; // set the variant type
			switch (pArg->vt)
			{
			case VT_UI1:
			case VT_I2:
			case VT_I4:
			case VT_INT:
			case VT_UINT:
			case VT_R4:
			case VT_R8:
			case VT_DATE:
			case VT_CY:
			case VT_BSTR:
			case VT_DISPATCH:
			case VT_ERROR:
			case VT_BOOL:
			case VT_VARIANT:
			case VT_UNKNOWN:
			case VT_I2|VT_BYREF:
			case VT_UI1|VT_BYREF:
			case VT_I4|VT_BYREF:
			case VT_R4|VT_BYREF:
			case VT_R8|VT_BYREF:
			case VT_DATE|VT_BYREF:
			case VT_CY|VT_BYREF:
			case VT_BSTR|VT_BYREF:
			case VT_DISPATCH|VT_BYREF:
			case VT_ERROR|VT_BYREF:
			case VT_BOOL|VT_BYREF:
			case VT_VARIANT|VT_BYREF:
			case VT_BSTR|VT_ARRAY:
			case VT_DISPATCH|VT_ARRAY:
				VariantCopy(pArg, &(VARIANT)va_arg(argList, VARIANT));
				break;

			default:
				assert(FALSE);  // unknown type!
				break;
			}

			// get ready to fill next argument
			--pArg; 
			++pb;
		}
	}

	// initialize return value
	VARIANT* pvarResult = NULL;
	VARIANT vaResult;
	VariantInit(&vaResult);
	if (vtRet != VT_EMPTY)
		pvarResult = &vaResult;

	// initialize EXCEPINFO struct
	EXCEPINFO excepInfo;
	memset(&excepInfo, 0, sizeof(excepInfo));

	UINT nArgErr = (UINT)-1;  // initialize to invalid arg

	// make the call
	SCODE sc = m_pDispatch->Invoke(dwDispID, IID_NULL, 0, wFlags, &dispparams, pvarResult, &excepInfo, &nArgErr);

	// cleanup any arguments that need cleanup
	if (0 != dispparams.cArgs)
	{
		VARIANT* pArg = dispparams.rgvarg + dispparams.cArgs - 1;
		const BYTE* pb = pbParamInfo;
		while (*pb != 0)
		{
			switch ((VARTYPE)*pb)
			{
			case VT_BSTR:
				VariantClear(pArg);
				break;
			}
			--pArg;
			++pb;
		}
	}
	delete[] dispparams.rgvarg;

	if (FAILED(sc))
	{
		VariantClear(&vaResult);
		if (sc != DISP_E_EXCEPTION)
		{
			// Non - Exception error code
			TRACE1(1, "Warning: DISP_E_EXCEPTION: Dispid: %X.\n", dwDispID);
			return sc;
		}

		// make sure excepInfo is filled in
		if (excepInfo.pfnDeferredFillIn)
			excepInfo.pfnDeferredFillIn(&excepInfo);

		if (excepInfo.bstrSource)
		{
			MAKE_TCHARPTR_FROMWIDE(szSource, excepInfo.bstrSource);
			SysFreeString(excepInfo.bstrSource);
		}
		if (excepInfo.bstrDescription)
		{
			MAKE_TCHARPTR_FROMWIDE(szDesc, excepInfo.bstrDescription);
			m_strError = szDesc;
			SysFreeString(excepInfo.bstrDescription);
		}
		if (excepInfo.bstrHelpFile)
			SysFreeString(excepInfo.bstrHelpFile);

		TRACE1(1, "Failed to get property: Dispid: %X.\n", dwDispID);
		return excepInfo.scode;
	}

	if (vtRet != VT_EMPTY)
	{
		// convert return value
		if (vtRet != VT_VARIANT && vtRet != vaResult.vt)
		{
			SCODE sc = VariantChangeType(&vaResult, &vaResult, 0, vtRet);
			if (FAILED(sc))
			{
				TRACE(1, "Warning: automation return value coercion failed.\n");
				VariantClear(&vaResult);
				return E_FAIL;
			}
			assert(vtRet == vaResult.vt);
		}

		// copy return value into return spot!
		switch (vtRet)
		{
		case VT_UI1:
		case VT_I2:
		case VT_I4:
		case VT_INT:
		case VT_UINT:
		case VT_R4:
		case VT_R8:
		case VT_DATE:
		case VT_CY:
		case VT_BSTR:
		case VT_DISPATCH:
		case VT_ERROR:
		case VT_BOOL:
		case VT_VARIANT:
		case VT_UNKNOWN:
		case VT_BSTR|VT_ARRAY:
		case VT_DISPATCH|VT_ARRAY:
			*pvRet = vaResult;
			break;
		default:
			assert(FALSE);  // invalid return type specified
		}
	}
	return NOERROR;
}

//
// InvokeHelper
//

HRESULT __cdecl COleDispatchDriver::InvokeHelper(DISPID      dwDispID, 
												 WORD        wFlags,
												 VARTYPE     vtRet, 
												 VARIANT*    pvRet, 
												 const BYTE* pbParamInfo, 
												 ...)
{
	HRESULT hResult;
	va_list argList;
	va_start(argList, pbParamInfo);
	hResult = InvokeHelperV(dwDispID, wFlags, vtRet, pvRet, pbParamInfo, argList);
	va_end(argList);
	return hResult;
}

//
// GetProperty
//

HRESULT COleDispatchDriver::GetProperty(DISPID dwDispID, VARTYPE vtProp, VARIANT* pvProp) const
{
	return ((COleDispatchDriver*)this)->InvokeHelper(dwDispID, DISPATCH_PROPERTYGET, vtProp, pvProp, NULL);
}

//
// SetProperty
//

HRESULT __cdecl COleDispatchDriver::SetProperty(DISPID dwDispID, VARTYPE vtProp, ...)
{
	va_list argList;    // really only one arg, but...
	va_start(argList, vtProp);

	BYTE rgbParams[2];
	if (vtProp & VT_BYREF)
		vtProp &= ~VT_BYREF;
	rgbParams[0] = (BYTE)vtProp;
	rgbParams[1] = NULL;
	WORD wFlags = (WORD)(vtProp == (VT_BYREF|VT_DISPATCH) ? DISPATCH_PROPERTYPUTREF : DISPATCH_PROPERTYPUT);
	HRESULT hResult = InvokeHelperV(dwDispID, wFlags, VT_EMPTY, NULL, rgbParams, argList);
	va_end(argList);
	return hResult;
}

//
// EnumItem
//
// Dump
//

void EnumItem::Dump()
{
	TRACE1(1, (LPCTSTR)m_strDisplay, _T("Name: %s"));
	TRACE(1, _T("\n"));
}


//
// CEnum
//

CEnum::~CEnum()
{
	int nCount = m_aEnumItems.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
		delete m_aEnumItems.GetAt(nIndex);
	SysFreeString(m_bstrEnumName);
}

//
// Add
//

int CEnum::Add(EnumItem* pEnumItem)
{
	return m_aEnumItems.Add(pEnumItem);
}

//
// Dump
//

void CEnum::Dump()
{
	MAKE_TCHARPTR_FROMWIDE(szEnumName, m_bstrEnumName);
	TRACE1(1, szEnumName, _T("Enum Name: %s"));
	TRACE(1, _T("\n"));
	int nCount = m_aEnumItems.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
		m_aEnumItems.GetAt(nIndex)->Dump();
}

//
// ByValue
//

EnumItem* CEnum::ByValue(long nValue, int& nIndex)
{
	int nCount = m_aEnumItems.GetSize();
	for (nIndex = 0; nIndex < nCount; nIndex++)
	{
		if (m_aEnumItems.GetAt(nIndex)->m_nValue == nValue)
			return m_aEnumItems.GetAt(nIndex);
	}
	nIndex = -1;
	return NULL;
}

//
// ByIndex
//

EnumItem* CEnum::ByIndex(int nIndex)
{
	int nCount = m_aEnumItems.GetSize();
	if (nIndex > -1 && nIndex < nCount)
		return m_aEnumItems.GetAt(nIndex);
	return NULL;
}

//
// IndexOfValue
//

int CEnum::IndexOfValue(long nValue)
{
	int nCount = m_aEnumItems.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		if (m_aEnumItems.GetAt(nIndex)->m_nValue == nValue)
			return nIndex;
	}
	return -1;
}

BOOL IsEqual (const VARIANT& v1, const VARIANT& v2)
{
	if (v1.vt != v2.vt)
		return FALSE;

	switch (v1.vt)
	{
	case VT_EMPTY:
	case VT_NULL:
		return TRUE;

	case VT_UI1:
		return v1.bVal == v2.bVal;

	case VT_I2:
		return v1.iVal == v2.iVal;

	case VT_I4:
		return v1.lVal == v2.lVal;

	case VT_R4:
		return v1.fltVal == v2.fltVal;

	case VT_R8:
		return v1.dblVal == v2.dblVal;

	case VT_BOOL:
		return v1.boolVal == v2.boolVal;

	case VT_ERROR:
		return v1.scode == v2.scode;

	case VT_CY:
		return memcmp(&v2.cyVal, &v1.cyVal, sizeof(CY)) == 0;

	case VT_DATE:
		return v1.date == v2.date;

	case VT_BSTR:
		return (::SysStringByteLen(v1.bstrVal) == ::SysStringByteLen(v2.bstrVal)) &&
				(memcmp(v1.bstrVal, v2.bstrVal, ::SysStringByteLen(v1.bstrVal)) == 0);

	case VT_UNKNOWN:
		return v1.punkVal == v2.punkVal;

	case VT_DISPATCH:
		return v1.pdispVal == v2.pdispVal;

	case VT_DECIMAL:
		return memcmp(&v1.decVal, &v2.decVal, sizeof(DECIMAL)) == 0;

	case VT_BYREF:
        return v1.byref == v2.byref;

	case VT_I1:
		return v1.cVal == v2.cVal;

	case VT_UI2:
        return v1.uiVal == v2.uiVal;

	case VT_UI4:
        return v1.ulVal == v2.ulVal;

	case VT_INT:
        return v1.intVal == v2.intVal;

	case VT_UINT:
        return v1.uintVal == v2.uintVal;

	case VT_BSTR|VT_ARRAY:
		return FALSE;

	case VT_DISPATCH|VT_ARRAY:
		return FALSE;

	default:
		assert(FALSE);
	}
	return FALSE;
}
//
// CProperty
//

CProperty::CProperty(LPDISPATCH pDispatch, 
					 MEMBERID   nMemId, 
					 VARTYPE    vt, 
					 EditType   etType, 
					 short      nParams, 
					 BSTR       bstrDocString,
					 DWORD      dwContextId,
					 CEnum*     pEnum)
	: m_szName(NULL),
	  m_nMemId(nMemId),
	  m_nParams(nParams),
	  m_theDriver(pDispatch),
	  m_vt(vt),
	  m_etType(etType),
	  m_pEnum(pEnum),
	  m_bstrDocString(bstrDocString)
{
	m_dwContextId = dwContextId;
	m_bPutRef = FALSE;
	m_bGet = FALSE;
	m_bPut = FALSE;
	VariantInit(&m_vProperty);
	VariantInit(&m_vPropertyInitial);
}

CProperty::~CProperty()
{
	delete [] m_szName;
	VariantClear(&m_vProperty);
	VariantClear(&m_vPropertyInitial);
	SysFreeString(m_bstrDocString);
}

//
// GetValue
//

HRESULT CProperty::GetValue()
{
	try
	{
		VariantClear(&m_vProperty);
		HRESULT hResult = m_theDriver.GetProperty(m_nMemId, m_vt, &m_vProperty);
		if (SUCCEEDED(hResult))
			VariantCopy(&m_vPropertyInitial, &m_vProperty);
		return hResult;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
	return S_OK;
}

BOOL CProperty::IsDirty()
{
	return !IsEqual (m_vPropertyInitial, m_vProperty);
}

//
// SetValue
//

HRESULT CProperty::SetValue()
{
	try
	{
		if (IsEqual(m_vPropertyInitial, m_vProperty))
			return S_OK;

		HRESULT hResult = m_theDriver.SetProperty(m_nMemId, m_vt, m_vProperty);
		if (SUCCEEDED(hResult))
			VariantCopy(&m_vPropertyInitial, &m_vProperty);

		return hResult;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
	return S_OK;
}

//
// SetValueNoCheck
//

HRESULT CProperty::SetValueNoCheck()
{
	try
	{
		HRESULT hResult = m_theDriver.SetProperty(m_nMemId, m_vt, m_vProperty);
		if (SUCCEEDED(hResult))
			VariantCopy(&m_vPropertyInitial, &m_vProperty);

		return hResult;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
	return S_OK;
}

//
// GetFont
//

HFONT CProperty::GetFont()
{
	HFONT  hFont = NULL;
	IFont* pFont;
	HRESULT hResult = GetData().pdispVal->QueryInterface(IID_IFont, (void**)&pFont);
	if (FAILED(hResult))
		return hFont;

	hResult = pFont->get_hFont(&hFont);
	pFont->Release();
	return hFont;
}

//
// SetFont
//

HRESULT CProperty::SetFont(HFONT hFont)
{
	IFont* pFont;
	HRESULT hResult = GetData().pdispVal->QueryInterface(IID_IFont, (void**)&pFont);
	if (FAILED(hResult))
		return hResult;

	LOGFONT lf;
	GetObject(hFont, sizeof(LOGFONT), &lf);
	
	MAKE_WIDEPTR_FROMTCHAR(wFaceName, lf.lfFaceName);
	BSTR bstrFaceName = SysAllocString(wFaceName);
	if (NULL == bstrFaceName)
	{
		pFont->Release();
		return E_OUTOFMEMORY;
	}

	hResult = pFont->put_Name(bstrFaceName);
	SysFreeString(bstrFaceName);
	
	CY size;
	size.Hi = 0L;
	size.Lo = 0L;
	HDC hDC = GetDC(NULL);
	if (hDC)
	{
		size.Hi = 0L;
		size.Lo = -MulDiv(lf.lfHeight * 100, 7200, GetDeviceCaps(hDC, LOGPIXELSY));
		ReleaseDC(NULL, hDC);
	}
	hResult = pFont->put_Size(size);
	if (FAILED(hResult))
	{
		pFont->Release();
		return hResult;
	}

	hResult = pFont->put_Bold((lf.lfWeight > (FW_NORMAL + FW_BOLD) / 2) ? VARIANT_TRUE : VARIANT_FALSE);
	if (FAILED(hResult))
	{
		pFont->Release();
		return hResult;
	}

	hResult = pFont->put_Italic(lf.lfItalic ? VARIANT_TRUE : VARIANT_FALSE);
	if (FAILED(hResult))
	{
		pFont->Release();
		return hResult;
	}

	hResult = pFont->put_Underline(lf.lfUnderline ? VARIANT_TRUE : VARIANT_FALSE);
	if (FAILED(hResult))
	{
		pFont->Release();
		return hResult;
	}

	hResult = pFont->put_Strikethrough(lf.lfStrikeOut ? VARIANT_TRUE : VARIANT_FALSE);
	if (FAILED(hResult))
	{
		pFont->Release();
		return hResult;
	}

	hResult = pFont->put_Weight((short)lf.lfWeight);
	if (FAILED(hResult))
	{
		pFont->Release();
		return hResult;
	}

	hResult = pFont->put_Charset(lf.lfCharSet);
	if (FAILED(hResult))
	{
		pFont->Release();
		return hResult;
	}

	pFont->Release();
	return hResult;
}

//
// GetPictureType
//

short CProperty::GetPictureType()
{
	LPDISPATCH pDispatch = GetData().pdispVal;
	if (NULL == pDispatch)
		return -1;

	IPicture* pPicture;
	HRESULT hResult = pDispatch->QueryInterface(IID_IPicture, (void**)&pPicture);
	if (FAILED(hResult))
		return -1;

	short nType;
	hResult = pPicture->get_Type(&nType);
	if (FAILED(hResult))
	{
		pPicture->Release();
		return -1;
	}

	pPicture->Release();
	return nType;
}

//
// ParentWndMonitorThread
//

ParentWndMonitorThread::ParentWndMonitorThread()
{
	m_bPaused = TRUE;
	m_dwWaitTimer = 100;
	DWORD dwThreadId;
	m_hThread = CreateThread(NULL, 
							 NULL, 
							 Process, 
							 (void*)this, 
							 CREATE_SUSPENDED, 
							 &dwThreadId);
}

ParentWndMonitorThread::~ParentWndMonitorThread()
{
}

//
// Reset
//

void ParentWndMonitorThread::Reset()
{
	Pause();
	m_bPaused = TRUE;
}

//
// SetWindow
//

void ParentWndMonitorThread::SetWindow(CPropertyPopup* pPropertyPopup)
{
	m_pPropertyPopup = pPropertyPopup;
}

//
// Process
//

DWORD WINAPI ParentWndMonitorThread::Process(void* pData)
{
	ParentWndMonitorThread* pMonitorThread = reinterpret_cast<ParentWndMonitorThread*>(pData);
	CRect rcParent;
	while (0 != pMonitorThread->m_dwWaitTimer)
	{
		if (IsWindow(pMonitorThread->m_pPropertyPopup->hWndParent()))
		{
			GetWindowRect(pMonitorThread->m_pPropertyPopup->hWndParent(), &rcParent);
			if (rcParent != pMonitorThread->m_rcParent)
			{
				pMonitorThread->m_rcParent = rcParent;
				if (IsWindow(pMonitorThread->m_pPropertyPopup->hWnd()))
					pMonitorThread->m_pPropertyPopup->PostMessage(GetGlobals().WM_HIDE, CBrowser::eHide);
			}
		}
		Sleep(pMonitorThread->m_dwWaitTimer);
	}
	return 0;
}

//
// Run
//

BOOL ParentWndMonitorThread::Run()
{
	if (!m_bPaused)
		return TRUE;
	if (m_pPropertyPopup && IsWindow(m_pPropertyPopup->hWndParent()))
		GetWindowRect(m_pPropertyPopup->hWndParent(), &m_rcParent);
	if (0xFFFFFFFF == ResumeThread(m_hThread))
		return FALSE;
	m_bPaused = FALSE;
	return TRUE;
}

//
// Stop
// 

BOOL ParentWndMonitorThread::Pause()
{
	if (m_bPaused)
		return TRUE;
	if (0xFFFFFFFF == SuspendThread(m_hThread))
		return FALSE;
	m_bPaused = TRUE;
	return TRUE;
}

//
// WaitUntilDead
//

BOOL ParentWndMonitorThread::WaitUntilDead()
{
	m_dwWaitTimer = 0;
	Run();
	return TRUE;
}

//
// CPropertyPopup
//

CPropertyPopup::CPropertyPopup()
	: m_pProperty(NULL)
{
	m_theMonitor.Reset();
}

CPropertyPopup::~CPropertyPopup()
{
	if (IsWindow(m_hWnd))
		DestroyWindow();
	m_theMonitor.WaitUntilDead();
}

ParentWndMonitorThread CPropertyPopup::m_theMonitor;
CPropertyPopup*		   CPropertyPopup::m_pCurrentPopup = NULL;
HHOOK				   CPropertyPopup::m_hHook = NULL;

//
// RegisterWindow
//

void CPropertyPopup::RegisterWindow()
{
	BOOL bResult = FWnd::RegisterWindow(DD_WNDCLASS("ABPropertyPopup"),
										CS_SAVEBITS | CS_DBLCLKS,
										(HBRUSH)(COLOR_WINDOW + 1),
										LoadCursor(NULL, MAKEINTRESOURCE(IDC_ARROW)),
										(HICON)NULL);
	assert(bResult);
}

//
// KeyboardProc
//

LRESULT CALLBACK CPropertyPopup::KeyboardProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	try
	{
		if (nCode < 0)
			return ::CallNextHookEx(m_hHook, nCode, wParam, lParam);

		if (!(lParam & 0x80000000))
		{
			if (m_pCurrentPopup && m_pCurrentPopup->PreTranslateKeyBoard (WM_KEYDOWN, wParam, lParam))
				return -1;
		}
		else
		{
			if (m_pCurrentPopup && m_pCurrentPopup->PreTranslateKeyBoard (WM_KEYUP, wParam, lParam))
				return -1;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return CallNextHookEx(m_hHook, nCode, wParam, lParam);
}

//
// Show
//

BOOL CPropertyPopup::Create(HWND hWndParent)
{
	m_hWndParent = hWndParent;
	FWnd::Create(_T(""),
				   WS_POPUPWINDOW,
				   0,
				   0,
				   0,
				   0,
   				   m_hWndParent,
				   (HMENU)NULL);
	return m_hWnd ? TRUE : FALSE;
}

//
// Show
//

void CPropertyPopup::Show(CRect rcPosition,	CProperty* pProperty)
{
	assert(NULL == m_hHook);
	if (NULL == m_hHook)
	{
		m_hHook = SetWindowsHookEx(WH_KEYBOARD,
								   (HOOKPROC)KeyboardProc, 
								   0, 
								   GetCurrentThreadId());
		assert(m_hHook);
	}
	m_pProperty = pProperty;
	m_pCurrentPopup = this;
	m_theMonitor.SetWindow(this);
	m_theMonitor.Run();
}

//
// Hide
//

void CPropertyPopup::Hide()
{
	if (m_hHook)
	{
		BOOL bResult = UnhookWindowsHookEx(m_hHook);
		m_hHook = NULL;
	}
	m_pCurrentPopup = NULL;
	m_theMonitor.Pause();
	DWORD dwStyle = GetWindowLong(m_hWnd, GWL_STYLE);
	if (WS_VISIBLE & dwStyle)
		ShowWindow(SW_HIDE);
}

//
// CEdit
//

CEdit::CEdit()
{
	RegisterWindow();
}

CEdit::~CEdit()
{
}

//
// Fill
//

void CEdit::Fill(void* pData)
{
	try
	{
		Clear();
		m_pProperty->GetValue();
		VARIANT& vProperty = m_pProperty->GetData();

		long nLBound;
		HRESULT hResult = SafeArrayGetLBound(vProperty.parray, 1,  &nLBound);
		if (FAILED(hResult))
			return;
		
		long nUBound;
		hResult = SafeArrayGetUBound(vProperty.parray, 1,  &nUBound);
		if (FAILED(hResult))
			return;

		BSTR* pbstr;
		hResult = SafeArrayAccessData(vProperty.parray, (void**)&pbstr);
		if (FAILED(hResult))
			return;

		int nLength = 1;
		int nCount = nUBound - nLBound + 1;
		for (int nIndex = nLBound; nIndex < nCount; nIndex++)
		{
			MAKE_TCHARPTR_FROMWIDE(szShortCut, pbstr[nIndex]);
		}
		SafeArrayUnaccessData(vProperty.parray);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// Create
//

BOOL CEdit::Create(HWND hWnd)
{
	if (!CPropertyPopup::Create(hWnd))
		return FALSE;
	m_hWndEdit = CreateWindow(_T("EDIT"), 
							  _T(""), 
							  WS_CHILD|WS_VSCROLL|WS_VISIBLE|ES_MULTILINE|ES_AUTOHSCROLL,
							  0,
							  0,
							  0,
							  0,
							  m_hWnd,
							  (HMENU)100,
							  g_hInstance,
							  NULL);
	HFONT hFont = (HFONT)::SendMessage(m_hWndParent, WM_GETFONT, 0, 0);
	::SendMessage(m_hWndEdit, WM_SETFONT, (WPARAM)hFont, TRUE);
	SetFont(hFont);
	return m_hWnd ? TRUE : FALSE;
}

//
// WindowProc
//

LRESULT CEdit::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
			DWORD dwStyle = GetWindowLong(m_hWnd, GWL_STYLE);
			if (WS_VISIBLE & dwStyle)
				ShowWindow(SW_HIDE);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return FWnd::WindowProc(nMsg, wParam, lParam);
}

//
// PreTranslateKeyBoard
//

BOOL CEdit::PreTranslateKeyBoard (UINT nMsg, UINT wParam, UINT lParam)
{
	switch (nMsg)
	{
	case WM_KEYUP:
		switch (wParam)
		{
		case VK_ESCAPE:
			return TRUE;
		}
		break;

	case WM_KEYDOWN:
		switch (wParam)
		{
		case VK_ESCAPE:
			return TRUE;

		case VK_DELETE:
			return TRUE;
		}
		break;
	}
	return FALSE;
}

//
// Show
//

void CEdit::Show(CRect rcBound, CProperty* pProperty)
{
	CPropertyPopup::Show(rcBound, pProperty);
	rcBound.Offset(0, rcBound.Height());

	::ClientToScreen(GetParent(m_hWndParent), rcBound);
	SetWindowPos(HWND_TOPMOST,
				 rcBound.left,
				 rcBound.top,
				 rcBound.Width(),
				 100,
				 SWP_NOZORDER|SWP_SHOWWINDOW|SWP_NOACTIVATE);
	::MoveWindow(m_hWndEdit, 0, 0, rcBound.Width(), 100, TRUE);
	UpdateWindow();
}

//
// CTabs
//

CComboListEdit::CComboListEdit()
{
	m_szClassName = _T("EDIT");
}

CComboListEdit::~CComboListEdit()
{
	if (IsWindow(m_hWnd))
		DestroyWindow();
}

BOOL CComboListEdit::Create(CComboList* pComboList, HFONT hFont)
{
	m_pComboList = pComboList;
	FWnd::CreateEx(0,
				   _T(""),
				   WS_CHILD|WS_VSCROLL|WS_VISIBLE|ES_MULTILINE|ES_AUTOHSCROLL,
				   0,
				   0,
				   0,
				   0,
				   m_pComboList->hWnd(),
				   (HMENU)100);
	SubClassAttach();
	SetFont(hFont);
	return m_hWnd ? TRUE : FALSE;
}

LRESULT CComboListEdit::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		switch (nMsg)
		{
		case WM_KEYDOWN:
			switch (wParam)
			{
			case VK_RETURN:
				if (!(GetKeyState(VK_CONTROL) < 0))
				{
					m_pComboList->Save();
					return 0;
				}
			}
			break;

		case WM_NCDESTROY:
			UnsubClass();
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
// CComboList
//

CComboList::CComboList()
{
	RegisterWindow();
}

CComboList::~CComboList()
{
}

//
// Save
//

void CComboList::Save()
{
	IComboList* pComboList = (IComboList*)m_pProperty->GetData().pdispVal;
	if (NULL == pComboList)
	{
		SendMessage(GetGlobals().WM_HIDE, CBrowser::eUpdate);
		return;
	}

	HRESULT hResult = pComboList->Clear();
	if (FAILED(hResult))
	{
		SendMessage(GetGlobals().WM_HIDE, CBrowser::eUpdate);
		return;
	}

	const WORD nLen = 128;
	TCHAR szBuffer[nLen];
	int nLineLength;
	int nLineCount = m_theEdit.SendMessage(EM_GETLINECOUNT);
	for (int nIndex = 0; nIndex < nLineCount; nIndex++)
	{
		memcpy(szBuffer, &nLen, sizeof(nLen));
		nLineLength = m_theEdit.SendMessage(EM_GETLINE, nIndex, (LPARAM)szBuffer);
		if (nLineLength > 0)
		{
			szBuffer[nLineLength] = NULL;
			MAKE_WIDEPTR_FROMTCHAR(wBuffer, szBuffer);
			pComboList->AddItem(wBuffer);
		}
	}

	hResult = m_pProperty->SetValueNoCheck();
	SendMessage(GetGlobals().WM_HIDE, CBrowser::eUpdate);
}

//
// Fill
//

void CComboList::Fill(void* pData)
{
	try
	{
		Clear();
		HRESULT hResult = m_pProperty->GetValue();
		if (FAILED(hResult))
			return;

		IComboList* pComboList = (IComboList*)m_pProperty->GetData().pdispVal;
		if (NULL == pComboList)
			return;

		short nCount;
		hResult = pComboList->Count(&nCount);
		if (FAILED(hResult) && nCount > 0)
			return;

		int nRequiredLength = 0;
		BSTR bstrText;
		VARIANT vIndex;
		vIndex.vt = VT_I2;
		for (vIndex.iVal = 0; vIndex.iVal < nCount; vIndex.iVal++)
		{
			hResult = pComboList->Item(&vIndex, &bstrText);
			if (FAILED(hResult))
				continue;

			MAKE_TCHARPTR_FROMWIDE(szText, bstrText);
			nRequiredLength += lstrlen(szText) + 2;
			SysFreeString(bstrText);
		}
		
		TCHAR* szBuffer = new TCHAR[nRequiredLength+1];
		if (NULL == szBuffer)
			return;

		TCHAR* szCurrent = szBuffer;
		for (vIndex.iVal = 0; vIndex.iVal < nCount; vIndex.iVal++)
		{
			hResult = pComboList->Item(&vIndex, &bstrText);
			if (FAILED(hResult))
				continue;

			MAKE_TCHARPTR_FROMWIDE(szText, bstrText);
			_tcscpy(szCurrent, szText);
			szCurrent += lstrlen(szText);
			*(szCurrent++) = '\r';
			*(szCurrent++) = '\n';
			SysFreeString(bstrText);
		}
		*szCurrent = NULL;
		m_theEdit.SendMessage(WM_SETTEXT, 0, (LRESULT)szBuffer);
		delete [] szBuffer;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// Create
//

BOOL CComboList::Create(HWND hWnd)
{
	if (!CPropertyPopup::Create(hWnd))
		return FALSE;

	HFONT hFont = (HFONT)::SendMessage(m_hWndParent, WM_GETFONT, 0, 0);
	SetFont(hFont);

	m_theEdit.Create(this, hFont);
	return m_hWnd ? TRUE : FALSE;
}

//
// PreTranslateKeyBoard
//

BOOL CComboList::PreTranslateKeyBoard (UINT nMsg, UINT wParam, UINT lParam)
{
	switch (nMsg)
	{
	case WM_KEYUP:
		switch (wParam)
		{
		case VK_ESCAPE:
			return TRUE;
		}
		break;

	case WM_KEYDOWN:
		switch (wParam)
		{
		case VK_ESCAPE:
			PostMessage(GetGlobals().WM_HIDE, CBrowser::eHide);
			return TRUE;

		case VK_DELETE:
			return TRUE;
		}
		break;
	}
	return FALSE;
}

//
// WindowProc
//

LRESULT CComboList::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		if (GetGlobals().WM_HIDE == nMsg)
		{
			Hide();
			::SendMessage(m_hWndParent, GetGlobals().WM_BUTTONINACTIVE, wParam, 0);
			return 0;
		}
		switch (nMsg)
		{
		case WM_KILLFOCUS:
			if (!IsChild(m_hWnd, (HWND)wParam))
				CPropertyPopup::Hide();
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
// Show
//

void CComboList::Show(CRect rcBound, CProperty* pProperty)
{
	CPropertyPopup::Show(rcBound, pProperty);
	rcBound.Offset(0, rcBound.Height());

	Fill();
	::ClientToScreen(GetParent(m_hWndParent), rcBound);
	SetWindowPos(HWND_TOPMOST,
				 rcBound.left,
				 rcBound.top,
				 rcBound.Width(),
				 100,
				 SWP_NOZORDER|SWP_SHOWWINDOW|SWP_NOACTIVATE);
	::MoveWindow(m_theEdit.hWnd(), 0, 0, rcBound.Width(), 100, TRUE);
}

//
// Keys
//

DDString Keys::m_strAlt;
DDString Keys::m_strCtrl;
DDString Keys::m_strShift;

static Keys* s_pKeys = NULL;

Keys::Keys()
{
	m_strAlt.LoadString(IDS_ALT);
	m_strCtrl.LoadString(IDS_CONTROL);	
	m_strShift.LoadString(IDS_SHIFT);
}

Keys::~Keys()
{
	delete s_pKeys;
}

Keys* Keys::GetKeys()
{
	if (NULL == s_pKeys)
		s_pKeys = new Keys;
	return s_pKeys;
}

//
// CShortCutEdit
//

CShortCutEdit::CShortCutEdit(CEditListBox& theEditListBox)
	: m_theEditListBox(theEditListBox)
{
	m_szClassName = _T("EDIT");
	Keys::GetKeys();	
	HRESULT hResult = CoCreateInstance(CLSID_ShortCut, 
									   NULL, 
									   CLSCTX_INPROC_SERVER, 
									   IID_IShortCut, 
									   (LPVOID*)&m_pShortCutStore);
	if (FAILED(hResult))
	{
		DDString strError;
		strError.LoadString(IDS_ERR_FAILEDTOCREATESHORTCUTOBJECT);
		MessageBox(strError);
	}
}

//
// ~CShortCutEdit
//

CShortCutEdit::~CShortCutEdit()
{
	int nCount  = m_faShortCut.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
		delete [] m_faShortCut.GetAt(nIndex);
	m_faShortCut.RemoveAll();
	if (IsWindow(m_hWnd))
		DestroyWindow();
	if (m_pShortCutStore)
		m_pShortCutStore->Release();
}

//
// Create
//

BOOL CShortCutEdit::Create(HWND hWndParent)
{
	m_szClassName = _T("EDIT");
	FWnd::Create(_T(""), 
			     WS_CHILD|WS_VISIBLE|WS_BORDER|ES_AUTOHSCROLL|ES_WANTRETURN|ES_WANTRETURN,
			     0,
			     0,
			     0,
			     0,
			     hWndParent,
				 (HMENU)CEditListBox::eEditControl);
	HFONT hFont = (HFONT)::SendMessage(hWndParent, WM_GETFONT, 0, 0);
	SetFont(hFont);
	SubClassAttach();
	return m_hWnd ? TRUE : FALSE;
}

//
// StartEditing
//

void CShortCutEdit::StartEditing(TCHAR* szText, CRect& rcItem, int& nIndex, HFONT hFont)
{
	try
	{
		if (szText)
		{
			SetWindowText(szText);
			TCHAR* szChar = _tcschr(szText, ',');
			if (szChar)
			{
				TCHAR* szCharTmp = szChar;
				*szCharTmp = NULL;
				szChar = new TCHAR[lstrlen(szText) + 1];
				if (NULL == szChar)
					return;
				_tcscpy(szChar, szText);
				m_faShortCut.Add(szChar);
				*szCharTmp = ',';
				szChar = new TCHAR[lstrlen(szText) + 1];
				if (NULL == szChar)
					return;
				_tcscpy(szChar, szText);
				m_faShortCut.Add(szChar);
			}
			else
			{
				szChar = new TCHAR[lstrlen(szText)+1];
				if (NULL == szChar)
					return;
				_tcscpy(szChar, szText);
				m_faShortCut.Add(szChar);
			}
		}
		else
		{
			SetWindowText(_T(""));
			int nCount  = m_faShortCut.GetSize();
			for (int nIndex = 0; nIndex < nCount; nIndex++)
				delete [] m_faShortCut.GetAt(nIndex);
			m_faShortCut.RemoveAll();
		}

		SetFont(hFont);
		MoveWindow(rcItem);
		ShowWindow(SW_SHOW);
		SetFocus();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// ProcessKeyboard
//

BOOL CShortCutEdit::ProcessKeyboard(WPARAM wParam, LPARAM lParam)
{
	try
	{
		if (HIWORD(lParam) & KF_REPEAT)
			return FALSE;

		HRESULT hResult;
		UINT nControlState = 0;
		long nKeyCode;
		BSTR bstrDesc;

		if (GetKeyState(VK_MENU) < 0)
			nControlState |= FALT;

		if (GetKeyState(VK_CONTROL) < 0)
			nControlState |= FCONTROL;

		BOOL bShift = FALSE;
		if (GetKeyState(VK_SHIFT) < 0)
			bShift = TRUE;

		if (VK_BACK == wParam && 0 == nControlState)
		{
			hResult = m_pShortCutStore->GetKeyCode(1, &nKeyCode);
			if (SUCCEEDED(hResult) && nKeyCode)
				m_pShortCutStore->Set(1, 0, 0, VARIANT_FALSE);
			else
			{
				hResult = m_pShortCutStore->GetKeyCode(0, &nKeyCode);
				if (SUCCEEDED(hResult) && nKeyCode)
				{
					m_pShortCutStore->Set(0, 0, 0, VARIANT_FALSE);
					m_pShortCutStore->SetControlCode(0);
				}
			}
			hResult = m_pShortCutStore->get_Value(&bstrDesc);
			if (SUCCEEDED(hResult))
			{
				MAKE_TCHARPTR_FROMWIDE(szDesc, bstrDesc);
				SendMessage(WM_SETTEXT, 0, (LPARAM)szDesc);
				SysFreeString(bstrDesc);
			}
		}
		else
		{
			hResult = m_pShortCutStore->GetKeyCode(1, &nKeyCode);
			if (SUCCEEDED(hResult) && nKeyCode)
				MessageBeep(0xFFFFFFFF);
			else
			{
				hResult = m_pShortCutStore->GetKeyCode(0, &nKeyCode);
				if (SUCCEEDED(hResult) && nKeyCode)
				{
					//
					// Valid Second Keys
					//
					if ((wParam >= VK_NUMPAD0 && wParam <= VK_F24) ||
						(wParam >= VK_INSERT && wParam <= VK_DELETE) ||
						(wParam >= VK_PRIOR && wParam <= VK_DOWN) ||
						(wParam >= 'A' && wParam <= 'Z') ||
						(wParam >= 128 && wParam <= 254) ||
						(wParam == VK_BACK) ||
						(wParam == VK_RETURN) ||
						(wParam == VK_SCROLL)
					   )
					{

						if (wParam >= VK_PRIOR && wParam <= VK_DOWN && !(GetKeyState(VK_NUMLOCK) & 0x01))
							lParam |= 0x01000000;

						//
						// Add shift to these keys if it is pressed
						//
						m_pShortCutStore->Set(1, wParam, lParam, bShift ? VARIANT_TRUE : VARIANT_FALSE);
						hResult = m_pShortCutStore->get_Value(&bstrDesc);
						if (SUCCEEDED(hResult))
						{
							MAKE_TCHARPTR_FROMWIDE(szDesc, bstrDesc);
							if (!CheckForDuplicates(m_pShortCutStore))
								SendMessage(WM_SETTEXT, 0, (LPARAM)szDesc);
							else
							{
								m_pShortCutStore->Set(1, 0, 0, VARIANT_FALSE);
								MessageBeep(0xFFFFFFFF);
							}
							SysFreeString(bstrDesc);
						}
					}
					else if ((wParam >= ' ' && wParam <= '@') ||
						     (wParam >= '[' && wParam <= '`') || 
						     (wParam >= '{' && wParam <= '~')
					        )
					{
						m_pShortCutStore->Set(1, wParam, lParam, VARIANT_FALSE);
						hResult = m_pShortCutStore->get_Value(&bstrDesc);
						if (SUCCEEDED(hResult))
						{
							MAKE_TCHARPTR_FROMWIDE(szDesc, bstrDesc);
							if (!CheckForDuplicates(m_pShortCutStore))
								SendMessage(WM_SETTEXT, 0, (LPARAM)szDesc);
							else
							{
								m_pShortCutStore->Set(1, 0, 0, VARIANT_FALSE);
								MessageBeep(0xFFFFFFFF);
							}
							SysFreeString(bstrDesc);
						}
					}
				}
				else if (0 == nControlState)
				{
					if ((wParam >= VK_F1 && wParam <= VK_F24) ||
						(wParam >= VK_INSERT && wParam <= VK_DELETE) ||
						(wParam >= VK_PRIOR && wParam <= VK_DOWN)
					   )
					{
						if (wParam >= VK_PRIOR && wParam <= VK_DOWN && !(GetKeyState(VK_NUMLOCK) & 0x01))
							lParam |= 0x01000000;
						m_pShortCutStore->Set(0, wParam, lParam, bShift ? VARIANT_TRUE : VARIANT_FALSE);
						hResult = m_pShortCutStore->get_Value(&bstrDesc);
						if (SUCCEEDED(hResult))
						{
							MAKE_TCHARPTR_FROMWIDE(szDesc, bstrDesc);
							SendMessage(WM_SETTEXT, 0, (LPARAM)szDesc);
							SysFreeString(bstrDesc);
						}
					}
				}
				else
				{
					//
					// Valid First Keys with control keys being pressed
					//

					if ((wParam >= VK_NUMPAD0 && wParam <= VK_F24) ||
						(wParam >= VK_INSERT && wParam <= VK_DELETE) ||
						(wParam >= VK_PRIOR && wParam <= VK_DOWN) ||
						(wParam >= 'A' && wParam <= 'Z') ||
						(wParam >= 128 && wParam <= 254) ||
						(wParam == VK_BACK) ||
						(wParam == VK_RETURN) ||
						(wParam == VK_SCROLL)
					   )
					{
						//
						// Add shift to these keys if it is pressed
						//
						if (wParam >= VK_PRIOR && wParam <= VK_DOWN && !(GetKeyState(VK_NUMLOCK) & 0x01))
							lParam |= 0x01000000;
						m_pShortCutStore->Set(0, wParam, lParam,  bShift ? VARIANT_TRUE : VARIANT_FALSE);
						m_pShortCutStore->SetControlCode(nControlState);
						hResult = m_pShortCutStore->get_Value(&bstrDesc);
						if (SUCCEEDED(hResult))
						{
							MAKE_TCHARPTR_FROMWIDE(szDesc, bstrDesc);
							SendMessage(WM_SETTEXT, 0, (LPARAM)szDesc);
							SysFreeString(bstrDesc);
						}
					}
					else if ((wParam >= ' ' && wParam <= '@') ||
						     (wParam >= '[' && wParam <= '`') || 
						     (wParam >= '{' && wParam <= '~')
					        )
					{
						m_pShortCutStore->Set(0, wParam, lParam, VARIANT_FALSE);
						m_pShortCutStore->SetControlCode(nControlState);
						hResult = m_pShortCutStore->get_Value(&bstrDesc);
						if (SUCCEEDED(hResult))
						{
							MAKE_TCHARPTR_FROMWIDE(szDesc, bstrDesc);
							SendMessage(WM_SETTEXT, 0, (LPARAM)szDesc);
							SysFreeString(bstrDesc);
						}
					}
				}
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

//
// CheckForDuplicates
//

BOOL CShortCutEdit::CheckForDuplicates(IShortCut* pShortCut)
{
	try
	{
		HRESULT hResult;
		LRESULT lResult;
		int nCount = m_theEditListBox.SendMessage(LB_GETCOUNT) - 1;
		for (int nIndex = 0; nIndex < nCount; nIndex++)
		{
			lResult = m_theEditListBox.SendMessage(LB_GETITEMDATA, nIndex);
			if (LB_ERR != lResult)
			{
				hResult = pShortCut->IsEqual((ShortCut*)lResult);
				if (SUCCEEDED(hResult))
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
// UpdateString
//

void CShortCutEdit::UpdateString()
{
	try
	{
		IShortCut* pShortCut = NULL;
		if (LB_ERR == m_theEditListBox.CurrentItem())
			return;

		if (CheckForDuplicates(m_pShortCutStore))
		{
			SetWindowText(_T(""));
			MessageBeep(0xFFFFFFFF);
			m_pShortCutStore->Clear();
			return;
		}

		TCHAR szBuffer[256];
		LRESULT lResult = GetWindowText(m_hWnd, szBuffer, 256);
		if (lResult > 0)
		{
			HRESULT hResult;

			pShortCut = reinterpret_cast<IShortCut*>(m_theEditListBox.SendMessage(LB_GETITEMDATA, m_theEditListBox.CurrentItem()));
			if (NULL == pShortCut)
			{
				//
				// Insert at the end of the list box
				//

				if (lstrlen(szBuffer) > 0)
				{
					lResult = m_theEditListBox.SendMessage(LB_GETCOUNT);
					if (LB_ERR != lResult)
					{
						int nIndex = m_theEditListBox.SendMessage(LB_INSERTSTRING, (UINT)lResult-1, (LPARAM)szBuffer);
						if (LB_ERR != lResult)
						{
							hResult = m_pShortCutStore->Clone(&pShortCut);
							if (SUCCEEDED(hResult))
							{
								lResult = m_theEditListBox.SendMessage(LB_SETITEMDATA, (UINT)nIndex, (LPARAM)pShortCut);
								if (LB_ERR != lResult)
								{
									lResult = m_theEditListBox.SendMessage(LB_SETCURSEL, (UINT)nIndex+1);
									m_theEditListBox.CurrentItem(nIndex+1);
								}
							}
						}
					}
				}
			}
			else
			{
				//
				// Delete item and reinsert it
				//

				lResult = m_theEditListBox.SendMessage(LB_DELETESTRING, m_theEditListBox.CurrentItem());
				if (LB_ERR != lResult)
				{
					long nIndex = m_theEditListBox.SendMessage(LB_INSERTSTRING, (UINT)m_theEditListBox.CurrentItem(), (LPARAM)szBuffer);
					if (LB_ERR != nIndex)
					{
						hResult = m_pShortCutStore->Clone(&pShortCut);
						if (SUCCEEDED(hResult))
						{
							lResult = m_theEditListBox.SendMessage(LB_SETITEMDATA, (UINT)nIndex, (LPARAM)pShortCut);
							if (LB_ERR != lResult)
							{
								lResult = m_theEditListBox.SendMessage(LB_SETCURSEL, (UINT)nIndex);
								m_theEditListBox.CurrentItem(nIndex);
							}
						}
					}
				}
			}
		}
		else
		{
			pShortCut = reinterpret_cast<IShortCut*>(m_theEditListBox.SendMessage(LB_GETITEMDATA, m_theEditListBox.CurrentItem()));
			if (NULL != pShortCut)
				m_theEditListBox.SendMessage(LB_DELETESTRING, m_theEditListBox.CurrentItem());
		}
		m_pShortCutStore->Clear();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// CleanUp
//

void CShortCutEdit::CleanUp()
{
	try
	{
		int nCount  = m_faShortCut.GetSize();
		for (int nIndex = 0; nIndex < nCount; nIndex++)
			delete [] m_faShortCut.GetAt(nIndex);
		m_faShortCut.RemoveAll();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// WindowProc
//

LRESULT CShortCutEdit::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
//	switch (nMsg)
//	{
//	}
	return FWnd::WindowProc(nMsg, wParam, lParam);
}

//
// CEditListBox
//

CEditListBox::CEditListBox(CShortCut* pShortCut)
	: m_pShortCut(pShortCut)
{
	m_szClassName = _T("LISTBOX");
	m_pShortCutEdit = new  CShortCutEdit(*this);
	assert(m_pShortCutEdit);
	if (NULL == m_pShortCutEdit)
		throw E_OUTOFMEMORY;
	m_nCurrentEditItem = 0;
	m_bEmptyEditItem = TRUE;
	m_bEditMode = FALSE;
	m_bSorted = FALSE;
	m_nCount = 0;
}

//
// ~CEditListBox
//

CEditListBox::~CEditListBox()
{
	delete m_pShortCutEdit;
	if (IsWindow(m_hWnd))
		DestroyWindow();
}

//
// Create
//

BOOL CEditListBox::Create(HWND hWndParent)
{
	m_hWndParent = hWndParent;
	FWnd::Create(_T(""),
				 WS_CHILD|WS_VSCROLL|LBS_OWNERDRAWFIXED|LBS_HASSTRINGS|LBS_NOTIFY,
			     0,
			     0,
			     0,
			     0,
   			     hWndParent,
			     (HMENU)120);
	m_pShortCutEdit->Create(m_hWnd);
	SubClassAttach();
	if (m_bEmptyEditItem)
		SendMessage(LB_ADDSTRING, 0, (LPARAM)_T(""));
	return m_hWnd ? TRUE : FALSE;
}

//
// OnDelete
//

void CEditListBox::OnDelete()
{
	int nLast = SendMessage(LB_GETCOUNT);
	if (LB_ERR == nLast)
		return;
	
	int nCurSel = SendMessage(LB_GETCURSEL);
	if (LB_ERR == nCurSel)
		return;

	nLast--;
	if (m_bEmptyEditItem && nCurSel == nLast)
	{
		MessageBeep(0xFFFFFFFF);
		return;
	}
	
	if (LB_ERR != SendMessage(LB_DELETESTRING, nCurSel))
	{
		nLast--;
		if (nCurSel >= 0 && nCurSel < nLast)
		{
			SendMessage(LB_SETCURSEL, nCurSel);
			m_nCurrentEditItem = nCurSel;
		}
		else
		{
			m_nCurrentEditItem = nLast - 1;
			if (m_nCurrentEditItem < 0)
				m_nCurrentEditItem = 0;
			SendMessage(LB_SETCURSEL, m_nCurrentEditItem);
		}
	}
}

//
// GetCount
//

int CEditListBox::GetCount()
{
	int nCount = SendMessage(LB_GETCOUNT);
	if (m_bEmptyEditItem)
		nCount--;
	return nCount;
}

//
// GetText
//

void CEditListBox::GetText(int nIndex, DDString& strText)
{
	DDString* pstrItem = (DDString*)SendMessage(LB_GETITEMDATA, nIndex);
	if ((DDString*)-1 == pstrItem || NULL == pstrItem)
		strText = _T("");
	else
		strText = *pstrItem;
}

//
// StartEditing
//

void CEditListBox::StartEditing()
{
	try
	{
		int nIndex = SendMessage(LB_GETCURSEL);
		if (LB_ERR == nIndex)
			return;

		if (m_bEditMode)
		{
			if (nIndex == m_nCurrentEditItem)
				return;
			StopEditing();
		}

		SendMessage(WM_SETREDRAW, FALSE);

		TCHAR szText[120];
		int nResult = SendMessage(LB_GETTEXT, nIndex, (LRESULT)szText);
		if (LB_ERR != nResult)
		{
			CRect rcItem;
			nResult = SendMessage(LB_GETITEMRECT, nIndex, (LRESULT)&rcItem);
			if (LB_ERR == nResult)
			{
				SendMessage(WM_SETREDRAW, TRUE);
				return;
			}

			m_pShortCutEdit->StartEditing(szText, rcItem, nIndex, GetFont());
			SendMessage(LB_SETSEL, TRUE, ++nIndex);
			m_bEditMode = TRUE;
		}
		else
			SendMessage(WM_SETREDRAW, TRUE);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// StopEditing
//

inline void CEditListBox::StopEditing()
{
	m_pShortCutEdit->UpdateString();
	m_pShortCutEdit->ShowWindow(SW_HIDE);
	m_bEditMode = FALSE;
	SendMessage(WM_SETREDRAW, TRUE);
}

//
// ReportError
//

void CEditListBox::ReportError (LRESULT lResult)
{
	if (LB_ERR == lResult || LB_ERRSPACE == lResult)
		MessageBeep(MB_ICONEXCLAMATION);
}

//
// GetTextMetrics
//

BOOL CEditListBox::GetTextMetrics (HDC hDC, TEXTMETRIC& tm)
{
	HFONT hFontOld = NULL;
	HFONT hFontCurrent = NULL;
	if (IsWindow(m_hWnd))
	{
		hFontCurrent = GetFont();
		if (hFontCurrent)
			hFontOld = SelectFont(hDC, hFontCurrent);
	}
	::GetTextMetrics(hDC, &tm);
	if (hFontCurrent)
		SelectFont(hDC, hFontOld);
	return TRUE;
}

//
// FindIndex
//
// This for searching for the correct postion in the list box
//

int CEditListBox::FindIndex (int nStart, int nEnd, UINT nKey)
{
	int nIndex = -1;
	try
	{
		for (nStart; nStart < nEnd; nStart++)
		{
			DDString strItem;
			GetText (nStart, strItem);
			TCHAR cChar = *((LPCTSTR)strItem);
			if (_totupper(cChar) == _totupper((TCHAR)nKey))
			{
				nIndex = nStart;
				break;
			}
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return nIndex;
}

//
// DestroyItem
//

void CEditListBox::DestroyItem(const DWORD& dwItemData)
{
	try
	{
		DDString* pcsString = (DDString*)dwItemData;
		delete pcsString;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// WindowProc
//

LRESULT CEditListBox::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		switch (nMsg)
		{
		case LB_RESETCONTENT:
			return OnResetContent();

		case WM_KEYDOWN:
			if (0 == OnKeyDown (wParam, LOWORD(lParam), HIWORD(lParam)))
				return 0;
			break;

		case WM_LBUTTONDOWN:
			{
				LRESULT lResult = OnLButtonDown(wParam, lParam);
				if (0 == lResult)
					return lResult;
			}
			break;

		case WM_LBUTTONDBLCLK:
			{
				POINT pt = {LOWORD(lParam), HIWORD(lParam)};
				OnLButtonDblClk(wParam, pt);
			}
			break;

		case WM_CHAR:
			OnChar(wParam, LOWORD(lParam), HIWORD(lParam));
			break;

		case WM_COMMAND:
			switch (HIWORD(wParam))
			{
			case LBN_SELCHANGE:
				LRESULT lResult = FWnd::WindowProc(nMsg, wParam, lParam);
				m_nCurrentEditItem = SendMessage(LB_GETCURSEL);
				return lResult;
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
// OnResetContent
//

LRESULT CEditListBox::OnResetContent()
{
	try
	{
		LRESULT lResult = FWnd::WindowProc(LB_RESETCONTENT, 0, 0);
		if (m_bEmptyEditItem)
			SendMessage(LB_ADDSTRING, 0, (LPARAM)_T(""));
		return lResult;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return 0;
	}
}

//
// DrawItem
//

void CEditListBox::DrawItem(LPDRAWITEMSTRUCT pDrawItemStruct)
{
	try
	{
		TEXTMETRIC tm;
		COLORREF   crForeground;
		HBRUSH	   hBrushBackground;
		int nCount = SendMessage(LB_GETCOUNT);
		BOOL       bLastEntry = (pDrawItemStruct->itemID == (ULONG)nCount-1);
		HDC        hDC = pDrawItemStruct->hDC;
		CRect      rcItem = pDrawItemStruct->rcItem;

		if (pDrawItemStruct->itemState & ODS_SELECTED)
		{
			hBrushBackground = (HBRUSH)(1+COLOR_HIGHLIGHT);
			crForeground = GetSysColor(COLOR_HIGHLIGHTTEXT);
		}
		else
		{
			hBrushBackground = (HBRUSH)(1+COLOR_WINDOW);
			crForeground = GetSysColor(COLOR_WINDOWTEXT);
		}

		SetBkMode(hDC, TRANSPARENT);
		FillRect(hDC, &rcItem, hBrushBackground);
		SetTextColor(hDC, crForeground);
			
		TCHAR szText[120];
		int nResult = SendMessage(LB_GETTEXT, pDrawItemStruct->itemID, (LPARAM)szText);
		if (LB_ERR != nResult)
			nResult = DrawText(hDC, szText, lstrlen(szText), &rcItem, DT_LEFT|DT_NOCLIP|DT_NOPREFIX);		

		if ((ODS_FOCUS & pDrawItemStruct->itemState) || 0xFFFFFFFF == pDrawItemStruct->itemID)
			DrawFocusRect(hDC, &rcItem);

		if (bLastEntry)
		{
			GetTextMetrics(hDC, tm);
			rcItem.left += tm.tmAveCharWidth; 
			rcItem.top += tm.tmInternalLeading; 
			rcItem.right = rcItem.left + eLastItemWidth * tm.tmAveCharWidth;
			rcItem.bottom -= tm.tmInternalLeading;
			DrawFocusRect(hDC, &rcItem);
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// MeasureItem
//

void CEditListBox::MeasureItem(LPMEASUREITEMSTRUCT pMeasureItemStruct)
{
	HDC hDC = GetDC(m_hWnd);
	if (hDC)
	{
		TEXTMETRIC tmData;
		if (GetTextMetrics(hDC, tmData))
			pMeasureItemStruct->itemHeight = tmData.tmHeight;
		ReleaseDC(m_hWnd, hDC);
	}
}

//
// DeleteItem
//

void CEditListBox::DeleteItem(LPDELETEITEMSTRUCT pDeleteItemStruct)
{
	try
	{
		if (pDeleteItemStruct->itemData)
		{
			IShortCut* pShortCut = reinterpret_cast<IShortCut*>(pDeleteItemStruct->itemData);
			if (pShortCut)
				pShortCut->Release();
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnKeyDown
//

LRESULT CEditListBox::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	switch (nChar)
	{
	case VK_F2: // Edit mode
		StartEditing();
		return 0;

	case VK_RETURN:
		if (m_pShortCut)
			m_pShortCut->SaveAndClose();
		break;
	}
	return 1;
}

//
// OnLButtonDown
//

BOOL CEditListBox::OnLButtonDown(UINT nFlags, LPARAM lParam)
{
	int nResult = SendMessage(LB_ITEMFROMPOINT, 0, lParam);
	if (0 == HIWORD(nResult))
	{
		int nItem = LOWORD(nResult);
		CRect rcItem;
		nResult = SendMessage(LB_GETITEMRECT, nItem, (LPARAM)&rcItem);
		POINT pt;
		pt.x = LOWORD(lParam);
		pt.y = HIWORD(lParam);
		if (PtInRect(&rcItem, pt))
		{
			if (m_bEditMode)
			{
				StopEditing();
				return 0;
			}
			else if (nItem == SendMessage(LB_GETCURSEL))
			{
				StartEditing();
				return 0;
			}
			return -1;
		}
		else
		{
			if (m_bEditMode)
				StopEditing();
			::SendMessage(GetParent(m_hWnd), 
						  WM_COMMAND, 
						  MAKELONG(120, CBN_CLOSEUP), 
						  (LPARAM)NULL);
			return 0;
		}
	}
	else
	{
		if (m_bEditMode)
			StopEditing();
		::SendMessage(GetParent(m_hWnd), 
					  WM_COMMAND, 
					  MAKELONG(120, CBN_CLOSEUP), 
					  (LPARAM)NULL);
		return 0;
	}
	return -1;
}

//
// OnLButtonDblClk
//

void CEditListBox::OnLButtonDblClk(UINT nFlags, POINT pt)
{
	int nIndex = SendMessage(LB_GETCURSEL);
	if (LB_ERR != nIndex)
	{
		CRect rcItem;
		SendMessage(LB_GETITEMRECT, nIndex, (LRESULT)&rcItem);
		if (PtInRect(&rcItem, pt))
		{
			StartEditing();
			return;
		}
		else 
			StopEditing();
	}
}

//
// OnChar
//

void CEditListBox::OnChar(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	int nCount = SendMessage(LB_GETCOUNT);
	if (nCount > 0)
	{
		int nCurIndex = SendMessage(LB_GETCURSEL);
		if (LB_ERR == nCurIndex)
		{
			SendMessage(LB_SETCURSEL, 0);
			nCurIndex = 0;
		}
		StartEditing();
		LRESULT lResult = m_pShortCutEdit->SendMessage(WM_CHAR, 
													   nChar, 
												       MAKELONG(nRepCnt, nFlags));
		return;
	}
}

//
// SaveAndClose
//

void CEditListBox::SaveAndClose()
{
	m_pShortCut->SaveAndClose();
}

//
// Close
//

void CEditListBox::Close()
{
	m_pShortCut->Close();
}

//
// CShortCut
//

CShortCut::CShortCut()
{
	RegisterWindow();
	m_pEditListBox = new CEditListBox(this);
	assert(m_pEditListBox);
}

CShortCut::~CShortCut()
{
	delete m_pEditListBox;
}

//
// Fill
//

void CShortCut::Fill(void* pData)
{
	try
	{
		Clear();
		m_pProperty->GetValue();
		VARIANT& vProperty = m_pProperty->GetData();

		long nLBound;
		HRESULT hResult = SafeArrayGetLBound(vProperty.parray, 1,  &nLBound);
		if (FAILED(hResult))
			return;
		
		long nUBound;
		hResult = SafeArrayGetUBound(vProperty.parray, 1,  &nUBound);
		if (FAILED(hResult))
			return;

		IShortCut** ppShortCut;
		hResult = SafeArrayAccessData(vProperty.parray, (void**)&ppShortCut);
		if (FAILED(hResult))
			return;

		BSTR bstrDesc;
		int nIndex2;
		int nLength = 1;
		int nCount = nUBound - nLBound + 1;
		for (int nIndex = nLBound; nIndex < nCount; nIndex++)
		{
			hResult = ppShortCut[nIndex]->get_Value(&bstrDesc);
			if (SUCCEEDED(hResult))
			{
				MAKE_TCHARPTR_FROMWIDE(szShortCut, bstrDesc);
				SysFreeString(bstrDesc);
				nIndex2 = m_pEditListBox->SendMessage(LB_INSERTSTRING, nIndex, (LRESULT)szShortCut);
				if (LB_ERR != nIndex2)
				{
					nIndex2 = m_pEditListBox->SendMessage(LB_SETITEMDATA, nIndex2, (LRESULT)ppShortCut[nIndex]);
					if (LB_ERR != nIndex2)
						ppShortCut[nIndex]->AddRef();
				}
			}
		}
		SafeArrayUnaccessData(vProperty.parray);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// Create
//

BOOL CShortCut::Create(HWND hWndParent)
{
	if (!CPropertyPopup::Create(hWndParent))
		return FALSE;

	if (!m_pEditListBox->Create(m_hWnd))
		return FALSE;

	HFONT hFont = (HFONT)::SendMessage(m_hWndParent, WM_GETFONT, 0, 0);
	if (hFont)
	{
		SetFont(hFont);
		m_pEditListBox->SetFont(hFont);
	}
	return m_hWnd ? TRUE : FALSE;
}

//
// PreTranslateKeyBoard
//

BOOL CShortCut::PreTranslateKeyBoard (UINT nMsg, UINT wParam, UINT lParam)
{
	try
	{
		switch (nMsg)
		{
		case WM_KEYUP:
			switch (wParam)
			{
			case VK_ESCAPE:
				return TRUE;
			}
			break;

		case WM_KEYDOWN:
			switch (wParam)
			{
			case VK_DELETE:
				if (m_pEditListBox->EditMode())
					m_pEditListBox->ProcessKeyboard(wParam, lParam);
				else
					m_pEditListBox->OnDelete();
				return TRUE;

			case VK_LEFT:
			case VK_RIGHT:
			case VK_UP:
			case VK_DOWN:
				if (m_pEditListBox->EditMode())
					m_pEditListBox->ProcessKeyboard(wParam, lParam);
				else
				{
					m_pEditListBox->SendMessage(WM_KEYDOWN, wParam, lParam);
					return TRUE;
				}
				break;

			case VK_RETURN:
				if (m_pEditListBox->EditMode())
					m_pEditListBox->StopEditing();
				else
				{
					m_pEditListBox->SaveAndClose();
				}
				return TRUE;

			case VK_ESCAPE:
				if (m_pEditListBox->EditMode())
					m_pEditListBox->StopEditing();
				else
				{
					m_pEditListBox->Close();
					m_pEditListBox->SendMessage(LB_RESETCONTENT);
				}
				return TRUE;

			default:
				if (m_pEditListBox->EditMode())
					m_pEditListBox->ProcessKeyboard(wParam, lParam);
				else
				{
					m_pEditListBox->StartEditing();
					m_pEditListBox->ProcessKeyboard(wParam, lParam);
				}
				return TRUE;
			}
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return 0;
	}
	return FALSE;
}

//
// WindowProc
//

LRESULT CShortCut::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		if (GetGlobals().WM_HIDE == nMsg)
		{
			Close();
			::SendMessage(m_hWndParent, GetGlobals().WM_BUTTONINACTIVE, wParam, 0);
			return 0;
		}
		switch (nMsg)
		{
		case WM_SHOWWINDOW:
			if (0 == wParam)
			{
				SaveAndClose();
			}
			break;

		case WM_KEYDOWN:
			switch (wParam)
			{
			case VK_DELETE:
				break;
			}
			break;

		case WM_KILLFOCUS:
			if (wParam && !IsChild(m_hWnd, (HWND)wParam))
				Close();
			break;

		case WM_MEASUREITEM:
			if (120 == wParam)
				m_pEditListBox->MeasureItem((LPMEASUREITEMSTRUCT)lParam);
			break;

		case WM_DRAWITEM:
			if (120 == wParam)
				m_pEditListBox->DrawItem((LPDRAWITEMSTRUCT)lParam);
			break;

		case WM_DELETEITEM:
			if (120 == wParam)
				m_pEditListBox->DeleteItem((LPDELETEITEMSTRUCT)lParam);
			break;

		case WM_COMMAND:
			switch (HIWORD(wParam))
			{
			case LBN_SELCHANGE:
				m_pEditListBox->SendMessage(nMsg, wParam, lParam);
				break;

			case CBN_CLOSEUP:
				SaveAndClose();
				break;
			}
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return 0;
	}
	return FWnd::WindowProc(nMsg, wParam, lParam);
}

//
// SaveShortCuts
//

BOOL CShortCut::SaveShortCuts()
{
	try
	{
		//
		// Create an safearray of bstrs and put the shortcuts into it.
		//

		int nCount = m_pEditListBox->SendMessage(LB_GETCOUNT) - 1;
		SAFEARRAY* psaShortCuts = NULL;
		if (nCount < 0)
			return FALSE;

		SAFEARRAYBOUND rgsabound[1];    
		rgsabound[0].lLbound = 0;
		rgsabound[0].cElements = nCount;
		psaShortCuts = SafeArrayCreate(VT_DISPATCH, 1, rgsabound);
		if (NULL == psaShortCuts)
			return FALSE;

		IShortCut** ppShortCut;

		// Get a pointer to the elements of the array.
		HRESULT hResult = SafeArrayAccessData(psaShortCuts, (void**)&ppShortCut);
		if (FAILED(hResult))
		{
			SafeArrayDestroy(psaShortCuts);
			return FALSE;
		}

		for (int nIndex = 0; nIndex < nCount; nIndex++)
		{
			ppShortCut[nIndex] = (IShortCut*)m_pEditListBox->SendMessage(LB_GETITEMDATA, nIndex);
			if (LB_ERR == (int)ppShortCut[nIndex])
				continue;
		}

		SafeArrayUnaccessData(psaShortCuts);

		//
		// Update the property
		//
		
		m_pProperty->GetData().vt = VT_ARRAY|VT_DISPATCH;
		SAFEARRAY* pArrayOld = m_pProperty->GetData().parray;
		m_pProperty->GetData().parray = psaShortCuts;

		if (SUCCEEDED(m_pProperty->SetValue()))
			return TRUE;
		m_pProperty->GetData().parray = pArrayOld;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return FALSE;
}

//
// Show
//

void CShortCut::Show(CRect rcBound, CProperty* pProperty)
{
	CPropertyPopup::Show(rcBound, pProperty);
	rcBound.Offset(0, rcBound.Height());

	Fill();
	
	::ClientToScreen(GetParent(m_hWndParent), rcBound);
	
	SetWindowPos(HWND_TOPMOST,
				 rcBound.left,
				 rcBound.top,
				 rcBound.Width(),
				 100,
				 SWP_NOZORDER|SWP_SHOWWINDOW|SWP_NOACTIVATE);
	
	if (m_pEditListBox)
	{
		m_pEditListBox->SetWindowPos(NULL,
									 0,
									 0,
									 rcBound.Width(),
									 100,
									 SWP_NOZORDER|SWP_SHOWWINDOW|SWP_NOACTIVATE);
		m_pEditListBox->SendMessage(LB_SETCURSEL, 0); 
	}
}

//
// SaveAndClose
//

void CShortCut::SaveAndClose()
{
	if (SaveShortCuts())
		::SendMessage(m_hWndParent, GetGlobals().WM_BUTTONINACTIVE, CBrowser::eUpdate, 0);
	else
		::SendMessage(m_hWndParent, GetGlobals().WM_BUTTONINACTIVE, CBrowser::eError, 0);
	Close();
}

//
// Close
//

void CShortCut::Close()
{
	CPropertyPopup::Hide();
	if (IsWindow(m_hWndParent))
		::SendMessage(m_hWndParent, GetGlobals().WM_BUTTONINACTIVE, CBrowser::eHide, 0);
	if (m_pEditListBox && IsWindow(m_pEditListBox->hWnd()))
	{
		m_pEditListBox->SendMessage(LB_RESETCONTENT);
		m_pEditListBox->StopEditing();
	}
}

//
// Hide
//

void CShortCut::Hide()
{
	CPropertyPopup::Hide();
	if (m_pEditListBox && IsWindow(m_pEditListBox->hWnd()))
		m_pEditListBox->SendMessage(LB_RESETCONTENT);
}

//
// DDCheckListBox
//

class CCLItemData
{
public:
	CCLItemData();
	~CCLItemData();

	LPARAM m_lParam;
	BSTR   m_bstrName;
	BOOL   m_bChecked;
};

CCLItemData::CCLItemData() 
	: m_bstrName(NULL)
{
	m_bChecked = FALSE;
	m_lParam = 0;
}

CCLItemData::~CCLItemData()
{
	SysFreeString(m_bstrName);
};

DDCheckListBox::DDCheckListBox()
	: m_hFont(NULL)
{
	m_bDirty = FALSE;
}

DDCheckListBox::~DDCheckListBox()
{
}

//
// Create
//

BOOL DDCheckListBox::Create(DWORD dwStyle, const CRect& rc, HWND hWndParent, UINT nId)
{
	SetClassName(_T("LISTBOX"));

	if (!CreateEx(0, _T(""), dwStyle, rc.left, rc.top, rc.Width(), rc.Height(), hWndParent, (HMENU)nId, NULL))
		return FALSE;

	SubClassAttach();
	return TRUE;
}

//
// WindowProc
//

LRESULT DDCheckListBox::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	switch(nMsg)
	{
	case WM_LBUTTONDOWN:
		{
			POINT pt;
			pt.x = LOWORD(lParam);
			pt.y = HIWORD(lParam);
			FWnd::WindowProc(nMsg, wParam, lParam);
			OnLButtonDown(wParam, pt);
		}
		return 0;

	case WM_KEYDOWN:
		if (VK_SPACE == wParam)
		{
			int nIndex = SendMessage(LB_GETCARETINDEX);
			if (LB_ERR != nIndex)
			{
				// toggle check
				CCLItemData* pData = (CCLItemData*)GetItemData(nIndex);
				pData->m_bChecked ^= 1;
				InvalidateItem(nIndex);
				::SendMessage(GetParent(m_hWnd), 
					          WM_COMMAND, 
					          MAKELONG((WORD)GetWindowLong(m_hWnd,GWL_ID), 55), 
					          (LPARAM)m_hWnd);
			}
			return 0;
		}
		break;

	case WM_KEYUP:
		if (VK_SPACE == wParam)
			return 0;
		break;

	case WM_CHAR:
		return 0;
	}
	return FWnd::WindowProc(nMsg, wParam, lParam);
}

//
// MeasureItem
//

void DDCheckListBox::MeasureItem(LPMEASUREITEMSTRUCT pMeasureItem)
{
	HDC hDC = GetDC(m_hWnd);
	if (hDC)
	{
		assert(m_hFont);
		HFONT hFontOld = SelectFont(hDC, m_hFont);
		TEXTMETRIC tm;
		GetTextMetrics(hDC, &tm);
		m_nTextHeight = m_nItemHeight = tm.tmHeight;

		if ((eCheckBoxHeight + 4) > m_nItemHeight)
			m_nItemHeight = eCheckBoxHeight + 4;

		SelectFont(hDC, hFontOld);
		::ReleaseDC(NULL, hDC);
	}
	if (pMeasureItem)
		pMeasureItem->itemHeight = m_nItemHeight;
}

//
// DrawItem
//

void DDCheckListBox::DrawItem(LPDRAWITEMSTRUCT pDrawItem)
{
	HDC hDC = pDrawItem->hDC;
	assert(m_hFont);
	HFONT hFontOld = SelectFont(hDC, m_hFont);
	BOOL bSelected;
	BSTR bstr = NULL;

	if (((LONG)(pDrawItem->itemID) >= 0) && (pDrawItem->itemAction & ODA_DRAWENTIRE || pDrawItem->itemAction & ODA_SELECT))
	{
		BOOL bDisabled = !IsWindowEnabled(m_hWnd);

		COLORREF crBackOld;
		COLORREF crTextOld;

		if (!bDisabled && ODS_SELECTED & pDrawItem->itemState)
		{
			crBackOld = SetTextColor(hDC, GetSysColor(COLOR_HIGHLIGHTTEXT));
			crTextOld = SetBkColor(hDC, GetSysColor(COLOR_HIGHLIGHT));
		}
		else
		{
			crBackOld = SetBkColor(hDC, GetSysColor(COLOR_WINDOW));
			crTextOld = SetTextColor(hDC, BLACKNESS);
		}

		if (pDrawItem->itemData)
		{
			bstr = ((CCLItemData*)(pDrawItem->itemData))->m_bstrName;
			bSelected = ((CCLItemData*)(pDrawItem->itemData))->m_bChecked;
		}

		CRect rcText = pDrawItem->rcItem;
		rcText.left += eCheckBoxWidth + 4;
		MAKE_TCHARPTR_FROMWIDE(szText, bstr);

		ExtTextOut(hDC, 
			       rcText.left, 
				   rcText.top + (m_nItemHeight - m_nTextHeight + 1) / 2, 
				   ETO_OPAQUE, 
				   &pDrawItem->rcItem, 
				   szText, 
				   lstrlen(szText), 
				   NULL);
		
		CRect rc = pDrawItem->rcItem;
		rc.left += 1;
		rc.right = rc.left + eCheckBoxWidth;
		rc.top += (rc.Height() - eCheckBoxWidth) / 2;
		rc.bottom = rc.top + eCheckBoxWidth;

		DrawFrameControl(hDC, &rc, DFC_BUTTON, bSelected ? DFCS_BUTTONCHECK|DFCS_CHECKED : DFCS_BUTTONCHECK);
		
		SetTextColor(hDC,crTextOld);
		SetBkColor(hDC, crBackOld);
	}

	if (ODA_FOCUS & pDrawItem->itemState)
		DrawFocusRect(hDC, &pDrawItem->rcItem);

	SelectFont(hDC, hFontOld);
}

//
// AddItem
//

void DDCheckListBox::AddItem(BSTR bstrName, BOOL bChecked, LPARAM lParam)
{
	CCLItemData* pData = new CCLItemData;
	if (pData)
	{
		pData->m_bstrName = SysAllocString(bstrName);
		pData->m_bChecked = bChecked;
		pData->m_lParam = lParam;
		int nIndex = AddString((LPCTSTR)pData);
		SetItemData(nIndex, (LPARAM)pData);
	}
}

//
// GetlParam
//

LPARAM DDCheckListBox::GetlParam(int nIndex)
{
	return ((CCLItemData*)GetItemData(nIndex))->m_lParam;
}

//
// OnLButtonDown
//

void DDCheckListBox::OnLButtonDown(UINT nFlags, POINT pt) 
{
	CCLItemData* pData;
	CRect rcItem;
	BOOL bCtrlPressed = GetKeyState(VK_CONTROL) & 0x8000;
	if (pt.x < eCheckBoxWidth)
	{
		int nIndex = GetCurSel();
		if (LB_ERR != nIndex)
		{
			GetItemRect(nIndex, rcItem);
			rcItem.left++;
			rcItem.right = rcItem.left + eCheckBoxWidth ;
			rcItem.top += (m_nItemHeight - eCheckBoxHeight)/2;
			rcItem.bottom = rcItem.top + eCheckBoxHeight;
			if (PtInRect(&rcItem, pt))
			{
				if (bCtrlPressed)
				{
					BOOL bDoCheck = ((CCLItemData*)GetItemData(nIndex))->m_bChecked ^ 1;
					int nCount = GetCount();
					for (nIndex = 0; nIndex < nCount; nIndex++)
						((CCLItemData*)GetItemData(nIndex))->m_bChecked = bDoCheck;
					InvalidateRect(NULL, FALSE);
				}
				else
				{
					pData = ((CCLItemData *)GetItemData(nIndex));
					pData->m_bChecked ^= 1;
					InvalidateRect(&rcItem, FALSE);
				}

				::SendMessage(GetParent(m_hWnd),
							  WM_COMMAND,
							  MAKELONG((WORD)GetWindowLong(m_hWnd,GWL_ID),55),
							  (LPARAM)m_hWnd);
				m_bDirty = TRUE;
			}
		}
	}
}

//
// CompareItem
//

int DDCheckListBox::CompareItem(LPCOMPAREITEMSTRUCT pCompareItem)
{
	return 0; 
}

//
// OnChildNotify
//

BOOL DDCheckListBox::OnChildNotify(UINT nMsg, WPARAM wParam, LPARAM lParam, LRESULT* plResult)
{
	switch (nMsg)
	{
	case WM_DRAWITEM:
		DrawItem((LPDRAWITEMSTRUCT)lParam);
		return FALSE;

	case WM_MEASUREITEM:
		MeasureItem((LPMEASUREITEMSTRUCT)lParam);
		*plResult = TRUE;
		break;

	case WM_COMPAREITEM:
		*plResult = CompareItem((LPCOMPAREITEMSTRUCT)lParam);
		break;

	case WM_DELETEITEM:
		OnDeleteItem((LPDELETEITEMSTRUCT)lParam);
		break;

	default:
		return FALSE;
	}
	return TRUE;
}

void DDCheckListBox::OnDeleteItem(LPDELETEITEMSTRUCT pDeleteItem)
{
	CCLItemData* pData = (CCLItemData*)pDeleteItem->itemData;
	delete pData;
}

void DDCheckListBox::InvalidateItem(int nIndex)
{
	CRect rc;
	SendMessage(LB_GETITEMRECT, nIndex,(LPARAM)&rc);
	::InvalidateRect(m_hWnd, &rc, FALSE);
}

int DDCheckListBox::GetCheck(int nIndex)
{
	CCLItemData* pData = (CCLItemData*)GetItemData(nIndex);
	return pData->m_bChecked;
}

void DDCheckListBox::SetCheck(int nIndex, BOOL bValue)
{
	CCLItemData* pData = (CCLItemData*)GetItemData(nIndex);
	pData->m_bChecked = bValue;
}

//
// CFlags
//

CFlags::CFlags()
	: m_hWndOk(NULL)
{
	RegisterWindow();
}

CFlags::~CFlags()
{
}

//
// Fill
//

void CFlags::Fill(void* pData)
{
	CEnum* pEnum = (CEnum*)pData;
	m_theCheckList.SendMessage(LB_RESETCONTENT);
	EnumItem* pEnumItem;
	int nCount = pEnum->NumOfItems();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		pEnumItem = pEnum->Item(nIndex);
		m_theCheckList.AddItem(pEnumItem->m_strDisplay.AllocSysString(), 0 != (m_nValue & m_pProperty->GetEnum()->Item(nIndex)->m_nValue));
	}
}

//
// Create
//

BOOL CFlags::Create(HWND hWnd)
{
	if (!CPropertyPopup::Create(hWnd))
		return FALSE;

	HFONT hFont = (HFONT)::SendMessage(hWnd, WM_GETFONT, 0, 0);
	assert(hFont);

	m_theCheckList.m_hFont = hFont;
	CRect rcWindow;
	if (!m_theCheckList.Create(WS_VSCROLL|WS_CHILD|LBS_OWNERDRAWFIXED, rcWindow, m_hWnd, eListbox))
		return FALSE;

	m_hWndOk = CreateWindow(_T("BUTTON"), 
							_T("Ok"), 
							WS_CHILD|BS_PUSHBUTTON|BS_TEXT,
							0,
							0,
							0,
							0,
							m_hWnd,
							(HMENU)eOk,
							g_hInstance,
							NULL);
	if (hFont)
		::SendMessage(m_hWndOk, WM_SETFONT, (WPARAM)hFont, MAKELONG(TRUE, 0));
	return m_hWnd ? TRUE : FALSE;
}

//
// Show
//

void CFlags::Show(CRect rcBound, CProperty* pProperty)
{
	CPropertyPopup::Show(rcBound, pProperty);
	rcBound.Offset(0, rcBound.Height());

	//
	// If no width for the drop down has been set set it to the width of
	// the combobox
	//

	Fill(m_pProperty->GetEnum());
	m_theCheckList.SetCurSel(0);

	int nItems = m_theCheckList.GetCount();
	if (nItems > 8)
		nItems = 8;
	rcBound.bottom = rcBound.top + nItems * m_theCheckList.m_nItemHeight + 2 * GetSystemMetrics(SM_CYBORDER) + 18;
	::ClientToScreen(GetParent(m_hWndParent), rcBound);

	CRect rcChildren = rcBound;
	rcChildren.Offset(-rcChildren.left, -rcChildren.top);
	rcChildren.bottom -= 18;
	
	m_theCheckList.SetWindowPos(NULL,
							    rcChildren.left, 
							    rcChildren.top, 
							    rcChildren.Width(), 
							    rcChildren.Height(),
								SWP_NOZORDER|SWP_SHOWWINDOW);
	rcChildren.top = rcChildren.bottom - 4;
	rcChildren.bottom += 17;
	rcChildren.Inflate(-1, -1);
	
	::SetWindowPos(m_hWndOk,
				   NULL,
				   rcChildren.left, 
				   rcChildren.top, 
				   rcChildren.Width(), 
				   rcChildren.Height(),
				   SWP_NOZORDER|SWP_SHOWWINDOW);

	SetWindowPos(HWND_TOPMOST,
				 rcBound.left,
				 rcBound.top,
				 rcBound.Width(),
				 rcBound.Height(),
				 SWP_NOZORDER|SWP_SHOWWINDOW);
}

//
// GetFlags
//

UINT CFlags::GetFlags()
{
	UINT nFlags = 0;
	int nSize = m_theCheckList.GetCount();
	for (int nIndex= 0; nIndex < nSize; nIndex++)
	{
		if (m_theCheckList.GetCheck(nIndex))
			nFlags |= m_pProperty->GetEnum()->Item(nIndex)->m_nValue;
	}
	return nFlags;
}

//
// PreTranslateKeyBoard
//

BOOL CFlags::PreTranslateKeyBoard (UINT nMsg, UINT wParam, UINT lParam)
{
	try
	{
		switch (nMsg)
		{
		case WM_KEYUP:
			switch (wParam)
			{
			case VK_ESCAPE:
				return TRUE;
			}
			break;

		case WM_KEYDOWN:
			switch (wParam)
			{
			case VK_LEFT:
			case VK_RIGHT:
			case VK_UP:
			case VK_DOWN:
				m_theCheckList.SendMessage(WM_KEYDOWN, wParam, lParam);
				return TRUE;

			case VK_TAB:
				{
					HWND hWnd = GetFocus();
					if (hWnd == m_theCheckList.hWnd())
						::SetFocus(m_hWndOk);
					else 
						m_theCheckList.SetFocus();
				}
				return TRUE;

			case VK_ESCAPE:
				PostMessage(GetGlobals().WM_HIDE, CBrowser::eHide); 
				return TRUE;

			case VK_DELETE:
				return TRUE;
			}
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return 0;
	}
	return FALSE;
}

//
// WindowProc
//

LRESULT CFlags::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		if (GetGlobals().WM_HIDE == nMsg)
		{
			CPropertyPopup::Hide();
			::SendMessage(m_hWndParent, GetGlobals().WM_BUTTONINACTIVE, wParam, 0);
			return 0;
		}
		switch (nMsg)
		{
		case WM_DESTROY:
			CPropertyPopup::Hide();
			::SendMessage(m_hWndParent, GetGlobals().WM_BUTTONINACTIVE, CBrowser::eHide, 0);
			break;

		case WM_COMMAND:
			switch (HIWORD(wParam))
			{
			case BN_CLICKED:
				switch (LOWORD(wParam))
				{
				case eOk:
					{
						long nFlagsOld = m_pProperty->GetData().lVal;
						m_pProperty->GetData().lVal = GetFlags();
						if (SUCCEEDED(m_pProperty->SetValue()))
							::SendMessage(m_hWndParent, GetGlobals().WM_BUTTONINACTIVE, CBrowser::eUpdate, 0);
						else
						{
							m_pProperty->GetData().lVal = nFlagsOld;
							::SendMessage(m_hWndParent, GetGlobals().WM_BUTTONINACTIVE, CBrowser::eError, 0);
						}
						Hide();
					}
					break;
				}
				break;
			}
			break;

		case WM_KILLFOCUS:
			if (!IsChild(m_hWnd, (HWND)wParam))
				CPropertyPopup::Hide();
			break;

		case WM_DRAWITEM:
		case WM_MEASUREITEM:
		case WM_COMPAREITEM:
		case WM_DELETEITEM:
			if (eListbox == wParam)
			{
				LRESULT lResult = 0;
				if (m_theCheckList.OnChildNotify(nMsg, wParam, lParam, &lResult))
					return lResult;
			}
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return 0;
	}
	return FWnd::WindowProc(nMsg, wParam, lParam);
}

//
// DrawBox
//

static void DrawBox(HDC hDC, const CRect& rcBound, HPEN hPen)
{
	HPEN hPenOld = SelectPen(hDC, hPen);
	MoveToEx(hDC, rcBound.left, rcBound.top, 0); 
	LineTo(hDC, rcBound.right, rcBound.top);
	MoveToEx(hDC, rcBound.right, rcBound.top, 0); 
	LineTo(hDC, rcBound.right, rcBound.bottom);
	MoveToEx(hDC, rcBound.right, rcBound.bottom, 0); 
	LineTo(hDC, rcBound.left, rcBound.bottom);
	MoveToEx(hDC, rcBound.left, rcBound.bottom, 0); 
	LineTo(hDC, rcBound.left, rcBound.top);
	SelectPen(hDC, hPenOld);
}

//
// CTabs
//

CTabs::CTabs(CFlickerFree* pffDraw)
{
    BOOL bResult = RegisterWindow(DD_WNDCLASS("DDTabs"),
								  CS_SAVEBITS|CS_DBLCLKS,
								  (HBRUSH)(COLOR_3DFACE+1),
								  LoadCursor(NULL, MAKEINTRESOURCE(IDC_ARROW)),
								  (HICON)NULL);
	assert(bResult);
	m_nCurrentTab = -1;
}

CTabs::~CTabs()
{
	int nCount = m_aTabs.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
		delete m_aTabs.GetAt(nIndex);
	m_aTabs.RemoveAll();
	if (IsWindow(m_hWnd))
		DestroyWindow();
}

BOOL CTabs::Create(HWND hWnd, HFONT hFont)
{
	try
	{
		m_hWndParent = hWnd;
		FWnd::CreateEx(0,
					   _T(""),
					   WS_VISIBLE|WS_CHILD,
					   0,
					   0,
					   0,
					   0,
					   m_hWndParent,
					   (HMENU)108);
		m_hFont = hFont;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return m_hWnd ? TRUE : FALSE;
}

//
// Add
//

void CTabs::Add(const DDString& strCaption)
{
	TabInfo* pTabInfo = new TabInfo;
	assert(pTabInfo);
	if (pTabInfo)
	{
		pTabInfo->m_strCaption = strCaption;
		m_aTabs.Add(pTabInfo);
	}
}

//
// TabPosition
//

void CTabs::TabPosition(int nTab)
{
	if (m_nCurrentTab != nTab)
	{
		m_nCurrentTab = nTab;
		NMHDR notify;
		notify.hwndFrom = m_hWnd; 
		notify.idFrom = m_nCurrentTab; 
		notify.code = TCN_SELCHANGE; 
		::SendMessage(m_hWndParent, WM_NOTIFY, 108, (LPARAM)&notify); 
		InvalidateRect(NULL, TRUE);
	}
}

//
// WindowProc
//

LRESULT CTabs::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		switch (nMsg)
		{
		case WM_SHOWWINDOW:
			{
				if (wParam && -1 != m_nCurrentTab)
				{
					NMHDR notify;
					notify.hwndFrom = m_hWnd; 
					notify.idFrom = m_nCurrentTab; 
					notify.code = TCN_SELCHANGE; 
					::SendMessage(m_hWndParent, WM_NOTIFY, 108, (LPARAM)&notify); 
				}
			}
			break;

		case WM_LBUTTONDOWN:
			{
				POINT pt;
				pt.x = LOWORD(lParam);
				pt.y = HIWORD(lParam);
				OnLButtonDown(wParam, pt);
			}
			break;

		case WM_PAINT:
			{
				PAINTSTRUCT ps;
				HDC hDC = BeginPaint(m_hWnd, &ps);
				if (hDC)
				{
					CRect rcWindow;
					GetClientRect(rcWindow);
					Draw(hDC, rcWindow);
					EndPaint(m_hWnd, &ps);
					return 0;
				}
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
// CalcLayout
//

void CTabs::CalcLayout()
{
	TabInfo* pTabInfo;
	HDC hDC = GetDC(m_hWnd);
	assert(hDC);
	if (hDC)
	{
		HFONT hFontOld = SelectFont(hDC, m_hFont);

		int nOffset = 0;
		int nCount = m_aTabs.GetSize();
		for (int nIndex = 0; nIndex < nCount; nIndex++)
		{
			pTabInfo = m_aTabs.GetAt(nIndex);
			DrawText(hDC, 
					 pTabInfo->m_strCaption, 
					 lstrlen(pTabInfo->m_strCaption), 
					 &pTabInfo->m_rcTab, 
					 DT_CALCRECT|DT_VCENTER|DT_SINGLELINE);
			pTabInfo->m_rcTab.bottom += 20 - pTabInfo->m_rcTab.Height() + 1;
			pTabInfo->m_rcTab.right += 10;
			pTabInfo->m_rcTab.Offset(nOffset + 1, 3);
			nOffset += pTabInfo->m_rcTab.Width();
		}

		SelectFont(hDC, hFontOld);
		ReleaseDC(m_hWnd, hDC);
	}
}

//
// OnDraw
//

void CTabs::Draw(HDC hDC, const CRect& rcWindow)
{
	try
	{
		SetTextColor(hDC, GetSysColor(COLOR_BTNTEXT));
		SetBkMode(hDC, TRANSPARENT);
		HFONT hFontOld = SelectFont(hDC, m_hFont);
		int nX = 0;
		int nY = 0;
		int nWidth = 0;
		int nHeight = 0;
		CRect rcTab = rcWindow;
		rcTab.top = rcTab.bottom - 1;
		UIUtilities::FillSolidRect(hDC, rcTab, GetSysColor(COLOR_BTNHIGHLIGHT));
		TabInfo* pTabInfo;
		CRect rcTemp = rcTab;
		int nCount = m_aTabs.GetSize();
		for (int nIndex = 0; nIndex < nCount; nIndex++)
		{
			pTabInfo = m_aTabs.GetAt(nIndex);
			rcTab = pTabInfo->m_rcTab;
			if (nIndex == m_nCurrentTab)
			{
				rcTemp = rcTab;
				rcTab.top -= 1;
				rcTab.left -= 1;
				rcTemp.top = rcWindow.bottom - 1;
				rcTemp.bottom = rcWindow.bottom;
			}
			
			nX = rcTab.left;
			nY = rcTab.top;
			nWidth = rcTab.Width();
			nHeight = rcTab.Height();
			UIUtilities::FillSolidRect(hDC, nX, nY + 1, 1, nHeight - 1, GetSysColor(COLOR_BTNHIGHLIGHT));
			UIUtilities::FillSolidRect(hDC, nX, nY, nWidth - 1, 1, GetSysColor(COLOR_BTNHIGHLIGHT));
			UIUtilities::FillSolidRect(hDC, nX + nWidth - 2, nY+1, 1, nHeight - 1, GetSysColor(COLOR_BTNSHADOW));
			UIUtilities::FillSolidRect(hDC, nX + nWidth - 1, nY+2, 1, nHeight - 2, GetSysColor(COLOR_3DDKSHADOW));
			
			if (nIndex == m_nCurrentTab)
				UIUtilities::FillSolidRect(hDC, rcTemp, GetSysColor(COLOR_3DFACE));

			DrawText(hDC, 
					 pTabInfo->m_strCaption, 
					 lstrlen(pTabInfo->m_strCaption), 
					 &rcTab, 
					 DT_CENTER|DT_VCENTER|DT_SINGLELINE);
			if (nIndex == m_nCurrentTab)
			{
				rcTab.Inflate(-3, -3);
				DrawFocusRect(hDC, &rcTab);
			}
		}
		SelectFont(hDC, hFontOld);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnLButtonDown
//

void CTabs::OnLButtonDown(UINT mFlags, POINT pt)
{
	TabInfo* pTabInfo;
	int nCount = m_aTabs.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		pTabInfo = m_aTabs.GetAt(nIndex);
		if (PtInRect(&pTabInfo->m_rcTab, pt))
		{
			TabPosition(nIndex);
			return;
		}
	}
	::SendMessage(m_hWndParent, 
				  WM_COMMAND, 
				  MAKELONG(108, CBN_CLOSEUP), 
				  (LONG)-1);
}

//
// CSystemColors
//

CSystemColors::CSystemColors(CFlickerFree* pffDraw)
	: m_pffDraw(pffDraw)
{
	m_szClassName = _T("LISTBOX");
	m_nCurrentIndex = -1;
}

CSystemColors::~CSystemColors()
{
	if (IsWindow(m_hWnd))
		DestroyWindow();
}

struct SystemColors 
{
	UINT nLabelId;
	UINT nColor;
};

static SystemColors s_SystemColors [] = 
{
	{IDS_SCROLLBAR, COLOR_SCROLLBAR},
	{IDS_DESKTOP, COLOR_DESKTOP},
	{IDS_ACTIVECAPTION, COLOR_ACTIVECAPTION},
	{IDS_INACTIVECAPTION, COLOR_INACTIVECAPTION},
	{IDS_MENU, COLOR_MENU},
	{IDS_WINDOW, COLOR_WINDOW},
	{IDS_WINDOWFRAME, COLOR_WINDOWFRAME},
	{IDS_MENUTEXT, COLOR_MENUTEXT},
	{IDS_WINDOWTEXT, COLOR_WINDOWTEXT},
	{IDS_CAPTIONTEXT, COLOR_CAPTIONTEXT},
	{IDS_ACTIVEBORDER, COLOR_ACTIVEBORDER},
	{IDS_INACTIVEBORDER, COLOR_INACTIVEBORDER},
	{IDS_APPWORKSPACE, COLOR_APPWORKSPACE},
	{IDS_HIGHLIGHT, COLOR_HIGHLIGHT},
	{IDS_HIGHLIGHTTEXT, COLOR_HIGHLIGHTTEXT},
	{IDS_BTNFACE, COLOR_BTNFACE},
	{IDS_BTNSHADOW, COLOR_BTNSHADOW},
	{IDS_GRAYTEXT, COLOR_GRAYTEXT},
	{IDS_BTNTEXT, COLOR_BTNTEXT},
	{IDS_INACTIVECAPTIONTEXT, COLOR_INACTIVECAPTIONTEXT},
	{IDS_BTNHIGHLIGHT, COLOR_BTNHIGHLIGHT},
	{IDS_3DDKSHADOW, COLOR_3DDKSHADOW},
	{IDS_3DLIGHT, COLOR_3DLIGHT},
	{IDS_INFOTEXT, COLOR_INFOTEXT},
	{IDS_INFOBK, COLOR_INFOBK}
};

//
// SetSelection
//

BOOL CSystemColors::SetSelection(COLORREF crColor)
{
	int nCount = sizeof(s_SystemColors) / sizeof(SystemColors);
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		if (GetSysColor(s_SystemColors[nIndex].nColor) == crColor)
			break;
	}
	if (nIndex == nCount)
		return FALSE;
	
	m_nCurrentIndex = nIndex;

	if (IsWindow(m_hWnd))
		SendMessage(LB_SETCURSEL, m_nCurrentIndex);
	return TRUE;
}

//
// GetSelection
//

BOOL CSystemColors::GetSelection(COLORREF& crColor)
{
	if (-1 == m_nCurrentIndex)
	{
		crColor = 0x80000000;
		return FALSE;
	}
	crColor = (0x80000000 + s_SystemColors[m_nCurrentIndex].nColor);
	return TRUE;
}

//
// Create
//

BOOL CSystemColors::Create(HWND hWnd, HFONT hFont)
{
	try
	{
		m_hWndParent = hWnd;
		FWnd::CreateEx(WS_EX_NOPARENTNOTIFY,
					   _T(""),
					   WS_OVERLAPPED|WS_BORDER|WS_VISIBLE|WS_CHILD|WS_VSCROLL|LBS_HASSTRINGS|LBS_OWNERDRAWFIXED|LBS_NOINTEGRALHEIGHT,
					   0,
					   0,
					   0,
					   0,
					   m_hWndParent,
					   (HMENU)107);
		SubClassAttach(m_hWnd);

		m_hFont = hFont;
		HDC hDC = GetDC(m_hWnd);
		if (hDC)
		{
			HFONT hFontOld = SelectFont(hDC, m_hFont);
			TEXTMETRIC tm;
			GetTextMetrics(hDC, &tm);
			m_nItemHeight = tm.tmHeight + tm.tmExternalLeading + 2;
			SelectFont(hDC, hFontOld);
			ReleaseDC(m_hWnd, hDC);
		}
		
		LRESULT lResult = SendMessage(LB_SETITEMHEIGHT, 0, MAKELPARAM(m_nItemHeight, 0));
		
		DDString strLabel;
		int nCount = sizeof(s_SystemColors) / sizeof(SystemColors);
		for (int nIndex = 0; nIndex < nCount; nIndex++)
		{
			strLabel.LoadString(s_SystemColors[nIndex].nLabelId);
			SendMessage(LB_ADDSTRING, 0, (LPARAM)(LPCTSTR)strLabel);
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return m_hWnd ? TRUE : FALSE;
}

//
// WindowProc
//

LRESULT CSystemColors::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		switch (nMsg)
		{
		case WM_KEYDOWN:
			switch (wParam)
			{
			case VK_RETURN:
				 m_nCurrentIndex = SendMessage(LB_GETCURSEL);
				::SendMessage(m_hWndParent, 
							  WM_COMMAND, 
							  MAKELONG(107, CBN_CLOSEUP), 
							  (LONG)(0x80000000 + s_SystemColors[m_nCurrentIndex].nColor));
				break;
			}
			break;

		case WM_LBUTTONUP:
			{
				 m_nCurrentIndex = SendMessage(LB_GETCURSEL);
				::SendMessage(m_hWndParent, 
							  WM_COMMAND, 
							  MAKELONG(107, CBN_CLOSEUP), 
							  (LONG)(0x80000000 + s_SystemColors[m_nCurrentIndex].nColor));
			}
			break;

		case WM_DRAWITEM:
			OnDrawItem((LPDRAWITEMSTRUCT)lParam);
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return 0;
	}
	return FWnd::WindowProc(nMsg, wParam, lParam);
}

//
// OnDrawItem
//

void CSystemColors::OnDrawItem(LPDRAWITEMSTRUCT pDrawItem)
{
	try
	{
		HDC	hDC = pDrawItem->hDC;
		CRect rcBound = pDrawItem->rcItem;
		CRect rc = rcBound;

		HDC hDCOff = m_pffDraw->RequestDC(hDC, 
										  rcBound.Width(), 
										  rcBound.Height());
		if (NULL == hDCOff)
			hDCOff = hDC;
		else
			rcBound.Offset(-rcBound.left, -rcBound.top);

		FillRect(hDCOff, &rcBound, (HBRUSH)(1+COLOR_WINDOW));
		if (pDrawItem->itemID >= 0 && (pDrawItem->itemAction & (ODA_DRAWENTIRE | ODA_SELECT | ODA_FOCUS)))
		{
			SetBkMode(hDCOff, TRANSPARENT);

			BOOL bDisabled = !IsWindowEnabled(m_hWnd);
			if (!bDisabled && 0 != (pDrawItem->itemState & ODS_SELECTED))
			{
				FillRect(hDCOff, &rcBound, (HBRUSH)(1+COLOR_HIGHLIGHT));
				SetTextColor(hDCOff, GetSysColor(COLOR_HIGHLIGHTTEXT));
			}
			else
				SetTextColor(hDCOff, GetSysColor(COLOR_BTNTEXT));

			//
			// Drawing the Color
			//

			CRect rcColor = rcBound;
			rcColor.right = rcColor.left + rcColor.Height();
			rcColor.Inflate(-2, -2);

			DrawBox(hDCOff, rcColor, GetStockPen(BLACK_PEN));

			rcColor.left += 1;
			rcColor.top += 1;
			HBRUSH hColor = CreateSolidBrush(GetSysColor(s_SystemColors[pDrawItem->itemID].nColor));
			if (hColor)
			{
				FillRect(hDCOff, &rcColor, hColor);
				DeleteBrush(hColor);
			}

			//
			// Drawing the Text
			//

			CRect rcText = rcBound;
			rcText.left = rcText.left + rcBound.Height() + 2;

			TCHAR szBuffer[MAX_PATH];
			if (LB_ERR != SendMessage(LB_GETTEXT, pDrawItem->itemID, (LPARAM)szBuffer))
			{
				HFONT hFontOld = SelectFont(hDCOff, m_hFont);
		
				DrawText(hDCOff, 
						 szBuffer, 
						 lstrlen(szBuffer), 
						 &rcText, 
						 DT_VCENTER|DT_SINGLELINE);

				SelectFont(hDCOff, hFontOld);
			}
		}
		
		if (0 != (pDrawItem->itemAction & ODA_FOCUS))
			DrawFocusRect(hDCOff, &rc);

		if (hDCOff != hDC)
			m_pffDraw->Paint(hDC, pDrawItem->rcItem.left, pDrawItem->rcItem.top);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// CPaleteColors
//


CPaleteColors::CPaleteColors(CFlickerFree* pffDraw)
	: m_pffDraw(pffDraw)
{
    BOOL bResult = RegisterWindow(DD_WNDCLASS("PaleteColors"),
								  CS_SAVEBITS|CS_DBLCLKS,
								  (HBRUSH)(COLOR_APPWORKSPACE+1),
								  LoadCursor(NULL, MAKEINTRESOURCE(IDC_ARROW)),
								  (HICON)NULL);
	assert(bResult);
	m_nCell = 0;
	m_nCurrentIndex = -1;
	m_nPrevIndex = -1;
}

CPaleteColors::~CPaleteColors()
{
	if (IsWindow(m_hWnd))
		DestroyWindow();
}
	
static COLORREF s_crPalette [] = 
{
	{RGB(255,255,255)},{RGB(255,192,192)},{RGB(255,224,192)},{RGB(255,255,192)},{RGB(192,255,192)},{RGB(192,255,255)},{RGB(192,192,255)},{RGB(255,192,255)},
	{RGB(224,224,224)},{RGB(255,128,128)},{RGB(255,192,128)},{RGB(255,255,128)},{RGB(128,255,128)},{RGB(128,255,255)},{RGB(128,128,255)},{RGB(255,128,255)},
	{RGB(192,192,192)},{RGB(255,000,000)},{RGB(255,128,000)},{RGB(255,255,000)},{RGB(000,255,000)},{RGB(000,255,255)},{RGB(000,000,255)},{RGB(255,000,255)},
	{RGB(128,128,128)},{RGB(192,000,000)},{RGB(255,064,000)},{RGB(192,192,000)},{RGB(000,192,000)},{RGB(000,192,192)},{RGB(000,000,192)},{RGB(192,000,192)},
	{RGB(064,064,064)},{RGB(128,000,000)},{RGB(128,064,000)},{RGB(128,128,000)},{RGB(000,128,000)},{RGB(000,128,128)},{RGB(000,000,128)},{RGB(128,000,128)},
	{RGB(000,000,000)},{RGB(064,000,000)},{RGB(128,064,064)},{RGB(064,064,000)},{RGB(000,064,000)},{RGB(000,064,064)},{RGB(000,000,064)},{RGB(064,000,064)},
	{RGB(255,255,255)},{RGB(255,255,255)},{RGB(255,255,255)},{RGB(255,255,255)},{RGB(255,255,255)},{RGB(255,255,255)},{RGB(255,255,255)},{RGB(255,255,255)},
	{RGB(255,255,255)},{RGB(255,255,255)},{RGB(255,255,255)},{RGB(255,255,255)},{RGB(255,255,255)},{RGB(255,255,255)},{RGB(255,255,255)},{RGB(255,255,255)},
};

//
// SetSelection
//

BOOL CPaleteColors::SetSelection(COLORREF crColor)
{
	int nCount = sizeof(s_crPalette) / sizeof(COLORREF);
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		if (s_crPalette[nIndex] == crColor)
			break;
	}
	m_nPrevIndex = m_nCurrentIndex;
	if (nIndex == nCount)
	{
		m_nCurrentIndex = -1;
		return FALSE;
	}
 	m_nCurrentIndex = nIndex;
	CRect rcItem;
	LRESULT lResult = SendMessage(LB_GETITEMRECT, m_nCurrentIndex, (LPARAM)&rcItem);
	if (LB_ERR != lResult)
		InvalidateRect(&rcItem, TRUE);
	return TRUE;
}

//
// GetSelection
//

BOOL CPaleteColors::GetSelection(COLORREF& crColor)
{
	if (-1 == m_nCurrentIndex)
	{
		crColor = RGB(255, 255, 255);
		return FALSE;
	}
	crColor = s_crPalette[m_nCurrentIndex];
	return TRUE;
}

//
// Create
//

BOOL CPaleteColors::Create(HWND hWnd, HFONT hFont)
{
	try
	{
		m_hWndParent = hWnd;
		FWnd::CreateEx(WS_EX_NOPARENTNOTIFY,
					   _T(""),
					   WS_CHILD|WS_OVERLAPPED|WS_BORDER,
					   0,
					   0,
					   0,
					   0,
					   m_hWndParent,
					   (HMENU)108);
		m_hFont = hFont;

		//
		// I need to read the registry for the custom colors
		//

		for (int nColor = 48; nColor < 64; nColor++)
			s_crPalette[nColor] = GetGlobals().m_thePreferences.GetColor(nColor);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return -10;
	}
	return m_hWnd ? TRUE : FALSE;
}

//
// UpdatePrev
//

void CPaleteColors::UpdatePrev()
{
	HDC hDC = GetDC(m_hWnd);
	if (hDC)
	{
		CRect rcOldCell;
		rcOldCell.top = (m_nPrevIndex / eVertCells) * m_nCell;
		rcOldCell.left = (m_nPrevIndex % eVertCells) * m_nCell;
		rcOldCell.bottom = rcOldCell.top + m_nCell;
		rcOldCell.right = rcOldCell.left + m_nCell;
		HPEN hPen = CreatePen(PS_SOLID, 0, GetSysColor(COLOR_APPWORKSPACE));
		if (hPen)
		{
			DrawBox(hDC, rcOldCell, hPen);
			BOOL bResult = DeletePen(hPen);
			assert(bResult);
		}
		ReleaseDC(m_hWnd, hDC);
	}
}

//
// CalcLayout
//

void CPaleteColors::CalcLayout(const CRect& rcBound, CRect& rcWindow)
{
	rcWindow.SetEmpty();
	int nWidth = rcBound.Width();
	int nHeight = rcBound.Height();
	m_nCell = min (nWidth / eHorzCells, nHeight / eVertCells);
	rcWindow.right = (m_nCell * eHorzCells) + 4*GetSystemMetrics(SM_CXEDGE) + 1;
	rcWindow.bottom = (m_nCell * eVertCells) + 4*GetSystemMetrics(SM_CYEDGE) + 1;
}

//
// HitTest
//

int CPaleteColors::HitTest(const POINT& pt, BOOL bSelect)
{
	try
	{
		CRect rcWindow;
		GetClientRect(rcWindow);
		CRect rcCell;
		BOOL bFound = FALSE;
		int nVert;
		int nHorz;
		int nIndex = -1;
		for (nVert = 0; nVert < eVertCells && !bFound; nVert++)
		{
			rcCell = rcWindow;
			rcCell.right = rcCell.left + m_nCell;
			rcCell.bottom = rcCell.top + m_nCell;
			rcCell.Offset(m_nCell * nVert, 0);
			for (nHorz = 0; nHorz < eHorzCells && !bFound; nHorz++)
			{
				if (PtInRect(&rcCell, pt))
				{
					if (bSelect)
					{
						HDC hDC = GetDC(m_hWnd);
						if (hDC)
						{
							if (-1 != m_nCurrentIndex)
							{
								CRect rcOldCell;
								rcOldCell.top = (m_nCurrentIndex/eVertCells) * m_nCell;
								rcOldCell.left = (m_nCurrentIndex%eVertCells) * m_nCell;
								rcOldCell.bottom = rcOldCell.top + m_nCell;
								rcOldCell.right = rcOldCell.left + m_nCell;
								HPEN hPen = CreatePen(PS_SOLID, 0, GetSysColor(COLOR_APPWORKSPACE));
								if (hPen)
								{
									DrawBox(hDC, rcOldCell, hPen);
									DeletePen(hPen);
								}
							}
							DrawBox(hDC, rcCell, GetStockPen(WHITE_PEN));
							ReleaseDC(m_hWnd, hDC);
						}
					}
					nIndex = nVert+(nHorz*eVertCells);
					bFound = TRUE;
				}
				else
					rcCell.Offset(0, m_nCell);
			}
		}
		if (bSelect)
		{
			m_nPrevIndex = m_nCurrentIndex;
			m_nCurrentIndex = nIndex;
		}
		return nIndex;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return -1;
	}
}

//
// WindowProc
//

LRESULT CPaleteColors::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		switch (nMsg)
		{
		case WM_KEYDOWN:
			switch (wParam)
			{
			case VK_RETURN:
				::SendMessage(m_hWndParent, 
							  WM_COMMAND, 
							  MAKELONG(108, CBN_CLOSEUP), 
							  (LONG)s_crPalette[m_nCurrentIndex]);
				break;

			case VK_DOWN:
				if (m_nCurrentIndex < 56)
				{
					m_nPrevIndex = m_nCurrentIndex;
					m_nCurrentIndex += 8;
					InvalidateRect(NULL, FALSE);
					UpdatePrev();
				}
				break;
			
			case VK_UP:
				if (m_nCurrentIndex > 7)
				{
					m_nPrevIndex = m_nCurrentIndex;
					m_nCurrentIndex -= 8;
					InvalidateRect(NULL, FALSE);
					UpdatePrev();
				}
				break;
			
			case VK_LEFT:
				if (m_nCurrentIndex > 0)
				{
					m_nPrevIndex = m_nCurrentIndex;
					m_nCurrentIndex--;
					InvalidateRect(NULL, FALSE);
					UpdatePrev();
				}
				break;

			case VK_RIGHT:
				if (m_nCurrentIndex < 63)
				{
					m_nPrevIndex = m_nCurrentIndex;
					m_nCurrentIndex++;
					InvalidateRect(NULL, FALSE);
					UpdatePrev();
				}
				break;
			}
			break;

		case WM_LBUTTONUP:
			{
				::SendMessage(m_hWndParent, 
							  WM_COMMAND, 
							  MAKELONG(108, CBN_CLOSEUP), 
							  (LONG)s_crPalette[m_nCurrentIndex]);
			}
			break;

		case WM_LBUTTONDOWN:
			{
				POINT pt;
				pt.x = LOWORD(lParam);
				pt.y = HIWORD(lParam);
				HitTest(pt);
			}
			break;

		case WM_RBUTTONUP:
			{
				POINT pt;
				pt.x = LOWORD(lParam);
				pt.y = HIWORD(lParam);
				int nHit = HitTest(pt, FALSE);
				if (-1 != nHit && nHit > 47)
				{
					try
					{
						//
						// Bring up customize color Dialog
						//

						LRESULT lResult = FWnd::WindowProc(nMsg, wParam, lParam);
						::ShowWindow(GetParent(m_hWnd), SW_HIDE);
						GetGlobals().m_pDefineColor->Color(s_crPalette[nHit]);
						LONG nColor = -1;
						if (ID_ADDCOLOR == GetGlobals().m_pDefineColor->DoModal(m_hWnd))
						{
							m_nPrevIndex = m_nCurrentIndex;
							m_nCurrentIndex = nHit;
							s_crPalette[m_nCurrentIndex] = GetGlobals().m_pDefineColor->Color(); 
							nColor = (LONG)s_crPalette[m_nCurrentIndex];
							
							GetGlobals().m_thePreferences.SetColor(m_nCurrentIndex, nColor);
							::SendMessage(m_hWndParent, WM_COMMAND, MAKELONG(108, CBN_CLOSEUP), nColor);
						}
						else
							::SendMessage(m_hWndParent, WM_COMMAND, MAKELONG(108, CBN_CLOSEUP), -1);

						return lResult;
					}
					CATCH
					{
						assert(FALSE);
						REPORTEXCEPTION(__FILE__, __LINE__)
						return 0;
					}
				}
			}
			break;

		case WM_PAINT:
			{
				PAINTSTRUCT ps;
				HDC hDC = BeginPaint(m_hWnd, &ps);
				if (hDC)
				{
					CRect rcWindow;
					GetClientRect(rcWindow);
					Draw(hDC, rcWindow);
					EndPaint(m_hWnd, &ps);
					return 0;
				}
			}
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return 0;
	}
	return FWnd::WindowProc(nMsg, wParam, lParam);
}

//
// OnDraw
//

void CPaleteColors::Draw(HDC hDC, const CRect& rcWindow)
{
	try
	{
		int nIndex;
		CRect rcColor;
		CRect rcColor2;
		CRect rcFill;
		BOOL bResult;
		int  nHorz;
		for (int nVert = 0; nVert < eVertCells; nVert++)
		{
			rcColor = rcWindow;
			rcColor.right = rcColor.left + m_nCell;
			rcColor.bottom = rcColor.top + m_nCell;
			rcColor.Offset(m_nCell * nVert, 0);
			for (nHorz = 0; nHorz < eHorzCells; nHorz++)
			{
				nIndex = nVert + (nHorz * eVertCells);
				if (m_nCurrentIndex == nIndex)
					DrawBox(hDC, rcColor, GetStockPen(WHITE_PEN));

				rcColor2 = rcColor;
				rcColor2.Inflate(-1, -1);
				DrawBox(hDC, rcColor2, GetStockPen(BLACK_PEN));

				rcFill = rcColor;
				rcFill.left += 1;
				rcFill.top += 1;
				rcFill.Inflate(-1, -1);
				HBRUSH hColor = CreateSolidBrush(s_crPalette[nIndex]);
				FillRect(hDC, &rcFill, hColor);
				bResult = DeleteBrush(hColor);
				assert(bResult);

				rcColor.Offset(0, m_nCell);
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
// CColor
//

CColorWnd::CColorWnd()
{
	RegisterWindow();
	
	m_pffDraw = new CFlickerFree;
	assert(m_pffDraw);
	if (NULL == m_pffDraw)
		throw E_OUTOFMEMORY;
	
	m_pSystemColors = new CSystemColors(m_pffDraw);
	assert(m_pSystemColors);
	if (NULL == m_pSystemColors)
		throw E_OUTOFMEMORY;
	
	m_pPaleteColors = new CPaleteColors(m_pffDraw);
	assert(m_pPaleteColors);
	if (NULL == m_pPaleteColors)
		throw E_OUTOFMEMORY;
	
	m_pTabs = new CTabs(m_pffDraw);
	assert(m_pTabs);
	if (NULL == m_pTabs)
		throw E_OUTOFMEMORY;

	DDString strCaption;
	strCaption.LoadString(IDS_PALETTE);
	m_pTabs->Add(strCaption);
	strCaption.LoadString(IDS_SYSTEM);
	m_pTabs->Add(strCaption);
}

CColorWnd::~CColorWnd()
{
	delete m_pSystemColors;
	delete m_pPaleteColors;
	delete m_pTabs;
	delete m_pffDraw;
	if (IsWindow(m_hWnd))
		DestroyWindow();
}

//
// RegisterWindow
//

void CColorWnd::RegisterWindow()
{
    BOOL bResult = FWnd::RegisterWindow(DD_WNDCLASS("DDColorPopup"),
										 CS_SAVEBITS | CS_DBLCLKS,
										 (HBRUSH)(COLOR_3DFACE + 1),
										 LoadCursor(NULL, MAKEINTRESOURCE(IDC_ARROW)),
										 (HICON)NULL);
	assert(bResult);
}

//
// SetSelection
// 

BOOL CColorWnd::SetSelection(COLORREF crColor)
{
	if (m_pSystemColors->SetSelection(crColor))
	{
		m_pTabs->TabPosition(1);
		return TRUE;
	}

	if (m_pPaleteColors->SetSelection(crColor))
	{
		m_pTabs->TabPosition(0);
		return TRUE;
	}

	m_pTabs->TabPosition(1);
	return FALSE;
}

//
// GetSelection
//

BOOL CColorWnd::GetSelection(COLORREF& crColor)
{
	if (m_pSystemColors->GetSelection(crColor))
		return TRUE;

	if (m_pPaleteColors->GetSelection(crColor))
		return TRUE;

	return FALSE;
}

//
// Create
//

BOOL CColorWnd::Create(HWND hWndParent)
{
	try
	{
		if (!CPropertyPopup::Create(hWndParent))
			return FALSE;

		m_hFont = (HFONT)::SendMessage(hWndParent, WM_GETFONT, 0, 0);
		if (NULL == m_hFont)
			return FALSE;

		m_pSystemColors->Create(m_hWnd, m_hFont);
		
		m_pPaleteColors->Create(m_hWnd, m_hFont);
		
		m_pTabs->Create(m_hWnd, m_hFont);
		
		m_pTabs->CalcLayout();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return 0;
	}
	return m_hWnd ? TRUE : FALSE;
}

//
// Show
//

void CColorWnd::Show(CRect rcBound, CProperty* pProperty)
{
	CPropertyPopup::Show(rcBound, pProperty);
	
	OLE_COLOR ocColor = m_pProperty->GetData().lVal;
	
	COLORREF crColor;
	OleTranslateColor(ocColor, NULL, &crColor);
	
	SetSelection(crColor);

	//
	// Setting up the bound for the color popup
	//

	rcBound.Offset(-3, rcBound.Height());
	::ClientToScreen(GetParent(m_hWndParent), rcBound);
	rcBound.top -= 1;
	rcBound.left = rcBound.right - 150;
	rcBound.right = rcBound.left + 150;
	rcBound.bottom = rcBound.top + 150;
	CRect rcWindow;
	m_pPaleteColors->CalcLayout(rcBound, rcWindow);
	rcWindow.bottom += 25;
	rcWindow.Offset(rcBound.left, rcBound.top);
	
	SetWindowPos(HWND_TOPMOST,
				 rcWindow.left,
				 rcWindow.top,
				 rcWindow.Width(),
				 rcWindow.Height(),
				 SWP_NOZORDER|SWP_SHOWWINDOW|SWP_NOACTIVATE);
}

//
// PreTranslateKeyBoard
//

BOOL CColorWnd::PreTranslateKeyBoard (UINT nMsg, UINT wParam, UINT lParam)
{
	try
	{
		switch (nMsg)
		{
		case WM_CHAR:
			switch (wParam)
			{
			case VK_TAB:
				if (0 == m_pTabs->TabPosition())
					m_pTabs->TabPosition(1);
				else
					m_pTabs->TabPosition(0);
				return TRUE;

			case VK_DELETE:
				return TRUE;

			case VK_UP:
				switch (m_pTabs->TabPosition())
				{
				case 0:
					m_pPaleteColors->SendMessage(nMsg, wParam, lParam);
					return TRUE;

				case 1:
					m_pSystemColors->SendMessage(nMsg, wParam, lParam);
					return TRUE;
				}
				break;

			case VK_DOWN:
				switch (m_pTabs->TabPosition())
				{
				case 0:
					m_pPaleteColors->SendMessage(nMsg, wParam, lParam);
					return TRUE;

				case 1:
					m_pSystemColors->SendMessage(nMsg, wParam, lParam);
					return TRUE;
				}
				break;

			case VK_RETURN:
				switch (m_pTabs->TabPosition())
				{
				case 0:
					m_pPaleteColors->SendMessage(nMsg, wParam, lParam);
					return TRUE;

				case 1:
					m_pSystemColors->SendMessage(nMsg, wParam, lParam);
					return TRUE;
				}
				break;
			}
			break;

		case WM_KEYUP:
			switch (wParam)
			{
			case VK_ESCAPE:
				return TRUE;
			}
			break;

		case WM_KEYDOWN:
			switch (wParam)
			{
			case VK_ESCAPE:
				PostMessage(GetGlobals().WM_HIDE, CBrowser::eHide); 
				return TRUE;

			case VK_DELETE:
				return TRUE;

			case VK_TAB:
				if (0 == m_pTabs->TabPosition())
					m_pTabs->TabPosition(1);
				else
					m_pTabs->TabPosition(0);
				InvalidateRect(NULL, FALSE);
				return TRUE;

			case VK_RETURN:
			case VK_DOWN:
			case VK_UP:
			case VK_LEFT:
			case VK_RIGHT:
				switch (m_pTabs->TabPosition())
				{
				case 0:
					m_pPaleteColors->SendMessage(nMsg, wParam, lParam);
					return TRUE;

				case 1:
					m_pSystemColors->SendMessage(nMsg, wParam, lParam);
					return TRUE;
				}
				break;
			}
			break;
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
// WindowProc
//

LRESULT CColorWnd::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		if (GetGlobals().WM_HIDE == nMsg)
		{
			Hide();
			::SendMessage(m_hWndParent, GetGlobals().WM_BUTTONINACTIVE, wParam, 0);
			return 0;
		}

		switch (nMsg)
		{
		case WM_KILLFOCUS:
			if (IsWindow(m_hWnd) && !IsChild(m_hWnd, (HWND)wParam))
				CPropertyPopup::Hide();
			break;

		case WM_KEYDOWN:
			{
				switch (wParam)
				{
				case VK_UP:
					if (m_pSystemColors->hWnd() == GetFocus())
						m_pSystemColors->SendMessage(nMsg, wParam, lParam);
					else if (m_pPaleteColors->hWnd() == GetFocus())
						m_pSystemColors->SendMessage(nMsg, wParam, lParam);
					break;

				case VK_DOWN:
					if (m_pSystemColors->hWnd() == GetFocus())
						m_pSystemColors->SendMessage(nMsg, wParam, lParam);
					else if (m_pPaleteColors->hWnd() == GetFocus())
						m_pSystemColors->SendMessage(nMsg, wParam, lParam);
					break;

				case VK_RETURN:
					if (m_pSystemColors->hWnd() == GetFocus())
						m_pSystemColors->SendMessage(nMsg, wParam, lParam);
					else if (m_pPaleteColors->hWnd() == GetFocus())
						m_pSystemColors->SendMessage(nMsg, wParam, lParam);
					break;
				}
			}
			break;

		case WM_DRAWITEM:
			::SendMessage(GetDlgItem(m_hWnd, wParam), nMsg, wParam, lParam);
			break;

		case WM_COMMAND:
			switch (HIWORD(wParam))
			{
			case CBN_CLOSEUP:
				{
					COLORREF cr = lParam;
					if (-1 != cr)
					{
						OLE_COLOR ocColorOld = m_pProperty->GetData().lVal;
						m_pProperty->GetData().lVal = cr;
						if (SUCCEEDED(m_pProperty->SetValue()))
							::SendMessage(m_hWndParent, GetGlobals().WM_BUTTONINACTIVE, CBrowser::eUpdate, 1);
						else
						{
							m_pProperty->GetData().lVal = ocColorOld;
							::SendMessage(m_hWndParent, GetGlobals().WM_BUTTONINACTIVE, CBrowser::eError, 0);
						}
					}
					else if (IsWindow(m_hWndParent))
						::SendMessage(m_hWndParent, GetGlobals().WM_BUTTONINACTIVE, CBrowser::eHide, 0);
				}
				break;
			}
			break;

		case WM_NOTIFY:
			{
				NMHDR* pNotify = (NMHDR*)lParam;
				switch (pNotify->code)
				{
				case TCN_SELCHANGE:
					{
						switch (pNotify->idFrom)
						{
						case 0:
							m_pSystemColors->ShowWindow(SW_HIDE);
							m_pPaleteColors->ShowWindow(SW_SHOW);
							break;

						case 1:
							m_pPaleteColors->ShowWindow(SW_HIDE);
							m_pSystemColors->ShowWindow(SW_SHOW);
							break;
						}
					}
				}
			}
			break;

		case WM_SIZE:
			if (m_pTabs->hWnd())
			{
				CRect rcClient;
				GetClientRect(rcClient);
				CRect rcTab = rcClient;
				rcTab.bottom = rcTab.top + 25;
				m_pTabs->MoveWindow(rcTab);
				rcClient.top = rcTab.bottom;
				rcClient.Inflate(-2, -2);
				m_pSystemColors->MoveWindow(rcClient);
				m_pPaleteColors->MoveWindow(rcClient);
			}
			break;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return 0;
	}
	return FWnd::WindowProc(nMsg, wParam, lParam);
}

CListbox::CListbox()
{
	m_szClassName = _T("LISTBOX");
}

CListbox::~CListbox()
{
}

BOOL CListbox::Create(HWND hWndParent)
{
	FWnd::Create(_T(""),
			    WS_CHILD|WS_VISIBLE|LBS_NOTIFY|WS_VSCROLL,
			    0,
			    0,
			    0,
			    0,
   			    hWndParent,
			    (HMENU)105);
	SubClassAttach();
	return (m_hWnd ? TRUE : FALSE);
}

LRESULT CListbox::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	switch (nMsg)
	{
	case WM_MOUSEACTIVATE:
		return MA_NOACTIVATE;

	case WM_LBUTTONUP:
		::PostMessage(GetParent(m_hWnd), WM_LBUTTONUP, wParam, lParam);
		break;
	}
	return FWnd::WindowProc(nMsg, wParam, lParam);
}

//
// CBrowserCombo
//

CBrowserCombo::CBrowserCombo()
{
	RegisterWindow();
}

CBrowserCombo::~CBrowserCombo()
{
	Clear();
}

//
// Fill
//

void CBrowserCombo::Fill(void* pData)
{
	Clear();
	switch (m_pProperty->GetEditType())
	{
	case CProperty::eBool:
		{
			DDString strTemp;
			strTemp.LoadString(IDS_TRUE);
			m_theListBox.SendMessage(LB_ADDSTRING,
									 0,
									 (LPARAM)(LPCTSTR)strTemp);
			strTemp.LoadString(IDS_FALSE);
			m_theListBox.SendMessage(LB_ADDSTRING,
									 0,
									 (LPARAM)(LPCTSTR)strTemp);
			SetSelection(m_pProperty->GetData().boolVal == VARIANT_TRUE ? 0 : 1);
		}
		break;

	case CProperty::eEnum:
		{
			CEnum* pEnum = m_pProperty->GetEnum();
			if (pEnum)
			{
				EnumItem* pEnumItem;
				int nCount = pEnum->NumOfItems();
				for (int nIndex = 0; nIndex < nCount; nIndex++)
				{
					pEnumItem = pEnum->Item(nIndex);
					if (pEnumItem)
					{
						m_theListBox.SendMessage(LB_ADDSTRING,
												 0,
												 (LPARAM)(LPCTSTR)pEnumItem->m_strDisplay);
					}
				}
				SetSelection(pEnum->IndexOfValue(m_pProperty->GetData().lVal));
			}
		}
		break;
	}
}

//
// Create
//

BOOL CBrowserCombo::Create(HWND hWnd)
{
	if (!CPropertyPopup::Create(hWnd))
		return FALSE;
	m_theListBox.Create(m_hWnd);
	HFONT hFont = (HFONT)::SendMessage(m_hWndParent, WM_GETFONT, 0, 0);
	if (hFont)
	{
		SetFont(hFont);
		m_theListBox.SetFont(hFont);
	}
	return m_hWnd ? TRUE : FALSE;
}

//
// Show
//

void CBrowserCombo::Show(CRect rcPosition, CProperty* pProperty)
{
	CPropertyPopup::Show(rcPosition, pProperty);
	Fill();

	//
	// Max number of items in the combobox
	//

	// Determine the maximum vertical scrolling position. 
	// The two is added for extra space below the lines         
	// of text. 
	
	int nCount = max (0, m_theListBox.SendMessage(LB_GETCOUNT));
	int nWidthDropDown = 0;
	int nHeightDropDown = 0;

	//
	// Calculate Width
	//
	
	HDC hDC = GetDC(m_hWnd);
	if (hDC)
	{
		HFONT hFontOld = SelectFont(hDC, m_hFont);
		TEXTMETRIC tm;
		GetTextMetrics(hDC, &tm);
		int nItemHeight = tm.tmHeight + tm.tmExternalLeading;

		//
		// Calculate the height of the drop down if no height was set it to the 
		// Default number of lines
		//

		if (nCount > eDefaultNumOfLines)
			nHeightDropDown = eDefaultNumOfLines * nItemHeight + 2;
		else
			nHeightDropDown = nCount * nItemHeight + 2;

		//
		// If a font was passed in set the m_hFont member and select this font
		// into the device context
		// 

		nWidthDropDown = rcPosition.Width();

		TCHAR szItem[256];
		SIZE sizeItem;
		for (int nItem = 0; nItem < nCount; nItem++)
		{
			m_theListBox.SendMessage(LB_GETTEXT, nItem, (LRESULT)szItem);
			if (GetTextExtentPoint32(hDC, szItem, lstrlen(szItem), &sizeItem))
			{
				sizeItem.cx =+ 4;
				if (sizeItem.cx > nWidthDropDown)
					nWidthDropDown = sizeItem.cx;
			}
		}
		SelectFont(hDC, hFontOld);
		ReleaseDC(m_hWnd, hDC);
	}
	CRect rcBound = rcPosition;

	rcBound.Offset(0, rcBound.Height()-1);
	rcBound.right = rcBound.left + nWidthDropDown;
	rcBound.bottom = rcBound.top + nHeightDropDown;

	//
	// If no width for the drop down has been set set it to the width of
	// the combobox
	//

	::ClientToScreen(GetParent(m_hWndParent), rcBound);

	SetWindowPos(HWND_TOPMOST,
				 rcBound.left,
				 rcBound.top,
				 rcBound.Width(),
				 rcBound.Height(),
				 SWP_NOZORDER|SWP_SHOWWINDOW|SWP_NOACTIVATE);
	rcBound.Offset(-rcBound.left, -rcBound.top);
	m_theListBox.MoveWindow(rcBound, TRUE);
}

//
// PreTranslateKeyBoard
//

BOOL CBrowserCombo::PreTranslateKeyBoard (UINT nMsg, UINT wParam, UINT lParam)
{
	switch (nMsg)
	{
	case WM_KEYUP:
		switch (wParam)
		{
		case VK_ESCAPE:
			return TRUE;
		}
		break;

	case WM_KEYDOWN:
		switch (wParam)
		{
		case VK_ESCAPE:
			PostMessage(GetGlobals().WM_HIDE, CBrowser::eHide); 
			return TRUE;

		case VK_DELETE:
			return TRUE;

		case VK_LEFT:
		case VK_RIGHT:
		case VK_UP:
		case VK_DOWN:
			m_theListBox.SendMessage(WM_KEYDOWN, wParam, lParam);
			return TRUE;

		case VK_RETURN:
			Update();
			return TRUE;
		}
		break;

	}
	return FALSE;
}

//
// Update
//

void CBrowserCombo::Update()
{
	switch (m_pProperty->GetEditType())
	{
	case CProperty::eBool:
		{
			VARIANT_BOOL vbTemp = m_pProperty->GetData().boolVal;
			VARIANT_BOOL vbBrowser = (GetSelection() == 0 ? VARIANT_TRUE : VARIANT_FALSE);
			if (vbBrowser != vbTemp)
			{
				m_pProperty->GetData().boolVal = (GetSelection() == 0 ? VARIANT_TRUE : VARIANT_FALSE);
				if (SUCCEEDED(m_pProperty->SetValue()))
					::SendMessage(m_hWndParent, GetGlobals().WM_BUTTONINACTIVE, CBrowser::eUpdate, 0);
				else
				{
					m_pProperty->GetData().boolVal = vbTemp;
					::SendMessage(m_hWndParent, GetGlobals().WM_BUTTONINACTIVE, CBrowser::eError, 0);
				}
			}
		}
		break;

	case CProperty::eEnum:
		{
			CEnum* pEnum = m_pProperty->GetEnum();
			if (pEnum)
			{
				EnumItem* pEnumItem = pEnum->ByIndex(GetSelection());
				if (pEnumItem)
				{
					switch (m_pProperty->VarType())
					{
					case VT_I4:
						{
							long nValOld = m_pProperty->GetData().lVal;
							if (nValOld != pEnumItem->m_nValue)
							{
								m_pProperty->GetData().lVal = pEnumItem->m_nValue;
								if (SUCCEEDED(m_pProperty->SetValue()))
									::PostMessage(m_hWndParent, GetGlobals().WM_BUTTONINACTIVE, CBrowser::eUpdate, 0);
								else
								{
									m_pProperty->GetData().lVal = nValOld;
									::SendMessage(m_hWndParent, GetGlobals().WM_BUTTONINACTIVE, CBrowser::eError, 0);
								}
							}
						}
						break;

					case VT_BSTR:
						{
							BSTR bstrOld = m_pProperty->GetData().bstrVal;
							BSTR bstrNew = pEnumItem->m_strDisplay.AllocSysString();
							if (NULL == bstrOld || 0 != wcscmp(bstrOld, bstrNew))
							{
								DDString strNone;
								strNone.LoadString(IDS_NONE);
								if (0 == _tcscmp(pEnumItem->m_strDisplay, strNone))
									m_pProperty->GetData().bstrVal = NULL;
								else
									m_pProperty->GetData().bstrVal = bstrNew;

								if (SUCCEEDED(m_pProperty->SetValue()))
								{
									SysFreeString(bstrOld);
									::SendMessage(m_hWndParent, GetGlobals().WM_BUTTONINACTIVE, CBrowser::eUpdate, 0);
								}
								else
								{	
									m_pProperty->GetData().bstrVal = bstrOld;
									::SendMessage(m_hWndParent, GetGlobals().WM_BUTTONINACTIVE, CBrowser::eError, 0);
								}
							}
							else 
								SysFreeString(bstrNew);
						}
						break;
					}
				}
			}
		}
	}
	Hide();
}

//
// WindowProc
//

LRESULT CBrowserCombo::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	if (GetGlobals().WM_HIDE == nMsg)
	{
		Hide();
		::SendMessage(m_hWndParent, GetGlobals().WM_BUTTONINACTIVE, wParam, 0);
		return 0;
	}
	switch (nMsg)
	{
	case WM_MOUSEACTIVATE:
		return MA_NOACTIVATE;

	case WM_KILLFOCUS:
		if (IsWindow(m_hWnd) && !IsChild(m_hWnd, (HWND)wParam))
			CPropertyPopup::Hide();
		break;

	case WM_LBUTTONUP:
		Update();
		break;

	case WM_SETFONT:
		if (0 != wParam)
			m_hFont = (HFONT)wParam;
		break;
	}
	return FWnd::WindowProc(nMsg, wParam, lParam);
}

//
// GetString
//

LPTSTR CBrowserCombo::GetString(int& nValue)
{
	static TCHAR szBuffer[256];
	if (LB_ERR == m_theListBox.SendMessage(LB_GETTEXT, nValue, (LRESULT)szBuffer))
		return NULL;
	return szBuffer;
}

//
// SetSelection
//

BOOL CBrowserCombo::SetSelection(int nIndex)
{
	m_theListBox.SendMessage(LB_SETCURSEL, nIndex);
	return TRUE;
}

//
// GetSelection
//

int CBrowserCombo::GetSelection()
{
	return m_theListBox.SendMessage(LB_GETCURSEL);
}


//
// CBrowserEdit
//

HHOOK CBrowserEdit::m_hHook = NULL;
CBrowserEdit* CBrowserEdit::m_pCurrentEdit = NULL;

CBrowserEdit::CBrowserEdit()
	: m_pBrowser(NULL),
	  m_pProperty(NULL),
	  m_pBrowserCombo(NULL),
	  m_pComboList(NULL),
	  m_pShortCut(NULL),
	  m_pColor(NULL),
	  m_pFlags(NULL),
	  m_pEdit(NULL),
	  m_szPriorText(NULL)
{
	m_szClassName = _T("EDIT");
	m_bChildActive = FALSE;
	
	//
	// Because it takes a long time to create
	//

	m_pColor = new CColorWnd;
	assert(m_pColor);
	if (NULL == m_pColor)
		throw E_OUTOFMEMORY;
}

CBrowserEdit::~CBrowserEdit()
{
	delete m_pBrowserCombo;
	delete m_pComboList;
	delete m_pShortCut;
	delete m_pColor;
	delete m_pFlags;
	delete m_pEdit;
	delete [] m_szPriorText;
}

//
// KeyboardProc
//

LRESULT CALLBACK CBrowserEdit::KeyboardProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	try
	{
		if (nCode < 0)
			return ::CallNextHookEx(m_hHook, nCode, wParam, lParam);

		if (!(lParam & 0x80000000))
		{
			if (m_pCurrentEdit && m_pCurrentEdit->PreTranslateKeyBoard (WM_KEYDOWN, wParam, lParam))
				return -1;
		}
		else
		{
			if (m_pCurrentEdit && m_pCurrentEdit->PreTranslateKeyBoard (WM_KEYUP, wParam, lParam))
				return -1;
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return CallNextHookEx(m_hHook, nCode, wParam, lParam);
}

//
// PreTranslateKeyBoard
//

BOOL CBrowserEdit::PreTranslateKeyBoard (UINT nMsg, UINT wParam, UINT lParam)
{
	try
	{
		switch (nMsg)
		{
		case WM_KEYDOWN:
			switch (wParam)
			{
			case VK_ESCAPE:
				return OnEscape(nMsg, wParam, lParam);
			case VK_RETURN:
				return OnReturn(nMsg, wParam, lParam);
			}
			break;
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
// BrowserCombo
//

CBrowserCombo* CBrowserEdit::BrowserCombo()
{
	if (NULL == m_pBrowserCombo)
	{
		m_pBrowserCombo = new CBrowserCombo;
		assert(m_pBrowserCombo);
		if (NULL == m_pBrowserCombo)
			throw E_OUTOFMEMORY;

		m_pBrowserCombo->Create(m_hWnd);
	}
	return m_pBrowserCombo;
}
	
//
// ComboList
//

CComboList* CBrowserEdit::ComboList()
{
	if (NULL == m_pComboList)
	{
		m_pComboList = new CComboList;
		assert(m_pComboList);
		if (NULL == m_pComboList)
			throw E_OUTOFMEMORY;
		m_pComboList->Create(m_hWnd);
	}
	return m_pComboList;
}	

CShortCut* CBrowserEdit::ShortCut()
{
	if (NULL == m_pShortCut)
	{
		m_pShortCut = new CShortCut;
		assert(m_pShortCut);
		if (NULL == m_pShortCut)
			throw E_OUTOFMEMORY;
		m_pShortCut->Create(m_hWnd);
	}
	return m_pShortCut;
}

CColorWnd* CBrowserEdit::ColorWnd()
{
	if (NULL == m_pColor->hWnd())
	{
		if (m_pColor)
			m_pColor->Create(m_hWnd);
	}
	return m_pColor;
}	

CFlags* CBrowserEdit::Flags()
{
	if (NULL == m_pFlags)
	{
		m_pFlags = new CFlags;
		assert(m_pFlags);
		if (NULL == m_pFlags)
			throw E_OUTOFMEMORY;

		m_pFlags->Create(m_hWnd);
	}
	return m_pFlags;
}	

CEdit* CBrowserEdit::Edit()
{
	if (NULL == m_pEdit)
	{
		m_pEdit = new CEdit;
		assert(m_pEdit);
		if (NULL == m_pEdit)
			throw E_OUTOFMEMORY;

		m_pEdit->Create(m_hWnd);
	}
	return m_pEdit;
}

//
// Create
//

BOOL CBrowserEdit::Create(CBrowser* pBrowser)
{
	m_pBrowser = pBrowser;
	FWnd::Create(_T(""),
			    WS_CHILD|ES_NOHIDESEL|ES_AUTOHSCROLL|ES_WANTRETURN,
			    0,
			    0,
			    0,
			    0,
   			    m_pBrowser->hWnd(),
			    (HMENU)105);
	HFONT hFont = m_pBrowser->GetFont();
	SetFont(hFont);
	SubClassAttach(m_hWnd);
	return m_hWnd ? TRUE : FALSE;
}

//
// Show
//

void CBrowserEdit::Show(CRect rcPosition, CProperty* pProperty, BOOL bShowChild, int nVirtKey)
{
	assert(pProperty);
	try
	{
		DDString strTemp;
		BOOL bDontShow = TRUE;
		m_pProperty = pProperty;
		switch (m_pProperty->GetEditType())
		{
		case CProperty::eComboList:
			if (bShowChild)
			{
				m_bChildActive = TRUE;
				ComboList()->Show(rcPosition, m_pProperty);
			}
			else
			{
				rcPosition.right = rcPosition.right - rcPosition.Height() - 2;
				SendMessage(EM_SETREADONLY, TRUE); 
			}
			break;

		case CProperty::eShortCut:
			if (bShowChild)
			{
				m_bChildActive = TRUE;
				ShortCut()->Show(rcPosition, m_pProperty);
			}
			else
			{
				rcPosition.right = rcPosition.right - rcPosition.Height() - 2;
				SendMessage(EM_SETREADONLY, TRUE); 
			}
			break;

		case CProperty::eEdit:
			if (bShowChild)
			{
				m_bChildActive = TRUE;
				Edit()->Show(rcPosition, m_pProperty);
			}
			else
			{
				rcPosition.right = rcPosition.right - rcPosition.Height() - 2;
				SendMessage(EM_SETREADONLY, TRUE); 
			}
			break;

		case CProperty::eBool:
		case CProperty::eEnum:
			if (bShowChild)
			{
				m_bChildActive = TRUE;
				BrowserCombo()->Show(rcPosition, m_pProperty);
			}
			else
			{
				rcPosition.right = rcPosition.right - rcPosition.Height() - 2;
				SendMessage(EM_SETREADONLY, TRUE); 
			}
			break;

		case CProperty::eColor:
			strTemp.Format(_T("&H%08X&"), m_pProperty->GetData().lVal);
			SetText(strTemp);
			if (bShowChild)
			{
				m_bChildActive = TRUE;
				ColorWnd()->Show(rcPosition, m_pProperty);
			}
			else
			{
				rcPosition.right = rcPosition.right - rcPosition.Height() - 2;
				rcPosition.left = rcPosition.left + rcPosition.Height() + 3;
				SendMessage(EM_SETREADONLY, FALSE); 
				bDontShow = FALSE;
			}
			break;

		case CProperty::eFlags:
			if (bShowChild)
			{
				m_bChildActive = TRUE;
				Flags()->Value(m_pProperty->GetData().lVal);
				Flags()->Show(rcPosition, m_pProperty);
			}
			else
			{
				rcPosition.right = rcPosition.right - rcPosition.Height() - 2;
				SendMessage(EM_SETREADONLY, FALSE); 
			}
			break;

		case CProperty::ePicture:
			{
				UINT nId = IDS_PICNONE;
				switch (pProperty->GetPictureType())
				{
				case PICTYPE_BITMAP:
					nId = IDS_PICBITMAP;
					break;

				case PICTYPE_ICON:
					nId = IDS_PICICON;
					break;

				case PICTYPE_METAFILE:
					nId = IDS_PICMETAFILE;
					break;

				case PICTYPE_ENHMETAFILE:
					nId = IDS_PICENHMETAFILE;
					break;
				}
				strTemp.LoadString(nId);
				SetText(strTemp);
				bDontShow = FALSE;
				rcPosition.right = rcPosition.right - rcPosition.Height() - 2;
				SendMessage(EM_SETREADONLY, FALSE); 
			}
			break;

		case CProperty::eNumber:
			{
				if (-1 != nVirtKey)
					strTemp.Format(_T("%c"), nVirtKey);
				else
				{
					long nNumber = 0;
					switch (m_pProperty->GetData().vt)
					{
					case VT_I2:
						nNumber = m_pProperty->GetData().iVal;
						break;

					case VT_INT:
					case VT_UINT:
					case VT_I4:
						nNumber = m_pProperty->GetData().lVal;
						break;
					}
					strTemp.Format(_T("%li"), nNumber);
				}
				SetText(strTemp);
				if (m_pProperty->Put())
				{
					bDontShow = FALSE;
					SendMessage(EM_SETREADONLY, FALSE); 
				}
			}
			break;

		case CProperty::eVariant:
			{
				switch (m_pProperty->VarType())
				{
				case VT_I2:
					strTemp.Format(_T("%li"), m_pProperty->GetData().iVal);
					break;

				case VT_INT:
				case VT_UINT:
				case VT_I4:
					strTemp.Format(_T("%li"), m_pProperty->GetData().lVal);
					break;

				case VT_BSTR:
					if (m_pProperty->GetData().bstrVal && *m_pProperty->GetData().bstrVal)
					{
						MAKE_TCHARPTR_FROMWIDE(szValue, m_pProperty->GetData().bstrVal);
						strTemp = szValue;
					}
					break;

				case VT_EMPTY:
					break;
				}
				SetText(strTemp);
				if (m_pProperty->Put())
					bDontShow = FALSE;
				SendMessage(EM_SETREADONLY, FALSE); 
			}
			break;

		case CProperty::eString:
			if (-1 != nVirtKey)
			{
				TCHAR szTemp[2];
				szTemp[0] = nVirtKey;
				szTemp[1] = NULL;
				SetText(szTemp);
			}
			else if (m_pProperty->GetData().bstrVal && *m_pProperty->GetData().bstrVal)
			{
				MAKE_TCHARPTR_FROMWIDE(szValue, m_pProperty->GetData().bstrVal);
				SetText(szValue);
			}
			else
				SetText(_T(""));
			if (m_pProperty->Put())
				bDontShow = FALSE;
			SendMessage(EM_SETREADONLY, FALSE); 
			break;
		}
		if (!bShowChild && !bDontShow)
		{
			if (m_hHook)
			{
				assert(FALSE);
				UnhookWindowsHookEx(m_hHook);
			}
			m_hHook = SetWindowsHookEx(WH_KEYBOARD, (HOOKPROC)KeyboardProc, 0, GetCurrentThreadId());
			m_pCurrentEdit = this;
			SetWindowPos(NULL, rcPosition.left, rcPosition.top, rcPosition.Width(), rcPosition.Height(), SWP_NOZORDER|SWP_SHOWWINDOW|SWP_NOACTIVATE);
			SetFocus();
			if (-1 == nVirtKey)
				SendMessage(EM_SETSEL, 0, -1);
			else
				SendMessage(EM_SETSEL, 1, -1);
			PostMessage(EM_SCROLLCARET);

			MSG msg;
			m_bProcessing = TRUE;
			while (m_bProcessing)
			{
				GetMessage(&msg, NULL, 0, 0);

				switch (msg.message)
				{
				case WM_LBUTTONDOWN:
					{
						CRect rcEdit;
						GetWindowRect(rcEdit);
						
						POINT pt;
						GetCursorPos(&pt);
						
						if (!PtInRect(&rcEdit, pt))
						{
							m_bProcessing = FALSE;
							Close(TRUE);
						}
					}
					break;
				}
				TranslateMessage(&msg);
				DispatchMessage(&msg);
			}
		}
		else
			ShowWindow(SW_HIDE);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// Close
//

void CBrowserEdit::Close(BOOL bHideParent)
{
	try
	{
		if (bHideParent)
			ShowWindow(SW_HIDE);

		if (NULL == m_pProperty)
			return;

		switch (m_pProperty->GetEditType())
		{
		case CProperty::eComboList:
			m_bChildActive = FALSE;
			ComboList()->CPropertyPopup::Hide();
			break;

		case CProperty::eShortCut:
			m_bChildActive = FALSE;
			ShortCut()->Hide();
			break;

		case CProperty::eBool:
		case CProperty::eEnum:
			m_bChildActive = FALSE;
			BrowserCombo()->CPropertyPopup::Hide();
			break;

		case CProperty::eColor:
			m_bChildActive = FALSE;
			ColorWnd()->CPropertyPopup::Hide();
			break;

		case CProperty::eEdit:
			m_bChildActive = FALSE;
			Edit()->CPropertyPopup::Hide();
			break;

		case CProperty::eFlags:
			m_bChildActive = FALSE;
			Flags()->CPropertyPopup::Hide();
			break;
		}
		m_pProperty = NULL;
		if (m_hHook)
		{
			UnhookWindowsHookEx(m_hHook);
			m_hHook = NULL;
			m_pCurrentEdit = NULL;
			m_bProcessing = FALSE;
		}
		m_pBrowser->SetFocus();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// Hide
//

void CBrowserEdit::Hide(BOOL bHideParent)
{
	try
	{
		if (bHideParent)
		{
			DWORD dwStyle = GetWindowLong(m_hWnd, GWL_STYLE);
			if (WS_VISIBLE & dwStyle)
				ShowWindow(SW_HIDE);
		}

		if (NULL == m_pProperty)
			return;

		switch (m_pProperty->GetEditType())
		{
		case CProperty::eComboList:
			m_bChildActive = FALSE;
			ComboList()->Hide();
			break;

		case CProperty::eShortCut:
			m_bChildActive = FALSE;
			ShortCut()->Close();
			break;

		case CProperty::eBool:
		case CProperty::eEnum:
			m_bChildActive = FALSE;
			BrowserCombo()->Hide();
			break;

		case CProperty::eColor:
			m_bChildActive = FALSE;
			ColorWnd()->Hide();
			break;

		case CProperty::eEdit:
			m_bChildActive = FALSE;
			Edit()->Hide();
			break;

		case CProperty::eFlags:
			m_bChildActive = FALSE;
			Flags()->Hide();
			break;
		}
		m_pProperty = NULL;
		if (m_hHook)
		{
			UnhookWindowsHookEx(m_hHook);
			m_hHook = NULL;
			m_pCurrentEdit = NULL;
			m_bProcessing = FALSE;
		}
		m_pBrowser->SetFocus();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

long ahextol(TCHAR* szHex)
{
	int nLen = lstrlen(szHex);
	long nTemp = 0;
	long nValue = 0;
	for (int nIndex = nLen-1, nIndex2 = 0; nIndex >= 0; nIndex--, nIndex2++)
	{
		if (szHex[nIndex] >= '0' && szHex[nIndex] <= '9')
			nValue = szHex[nIndex] - '0';
		else if (szHex[nIndex] >= 'A' && szHex[nIndex] <= 'F')
			nValue = (szHex[nIndex] - 'A') + 10;
		else
			nValue = 0;
		nTemp = nTemp + nValue * (long)pow(16, nIndex2);
	}
	return nTemp;
}

static BOOL IsNumber(LPCTSTR szString)
{
	int nLen = lstrlen(szString);
	if (nLen <= 0)
		return FALSE;
	for (int nIndex = 0; nIndex < nLen; nIndex++)
	{
		if (szString[nIndex] >= '0' && szString[nIndex] <= '9')
			continue;
		return FALSE;
	}
	return TRUE;
}

//
// UpdateText
//

BOOL CBrowserEdit::UpdateText()
{
	BOOL bResult = TRUE;
	try
	{
		if (NULL == m_pProperty)
			return bResult;

		switch (m_pProperty->GetEditType())
		{
		case CProperty::eVariant:
			{
				BOOL bResult = TRUE;
				LPTSTR szTemp = GetText();
				if (0 == lstrlen(szTemp))
				{
					VariantClear(&m_pProperty->GetData());
					m_pProperty->SetValueNoCheck();
					m_pProperty->VarType(VT_VARIANT);
					return TRUE;
				}
				else if (IsNumber(szTemp))
				{
					MAKE_WIDEPTR_FROMTCHAR(wEdit, szTemp);
					VARIANT vt;					
					vt.vt = VT_BSTR;
					vt.bstrVal = SysAllocString(wEdit);
					assert(vt.bstrVal);
					if (vt.bstrVal)
					{
						if (VT_I4 != m_pProperty->VarType() && VT_I2 != m_pProperty->VarType())
							m_pProperty->VarType(VT_I4);

						long nTemp = m_pProperty->GetData().lVal;
						HRESULT hResult = VariantChangeType(&m_pProperty->GetData(), 
															&vt, 
															0, 
															m_pProperty->VarType());
						if (SUCCEEDED(hResult))
						{
							hResult = m_pProperty->SetValueNoCheck();
							if (SUCCEEDED(hResult))
								PostMessage(GetGlobals().WM_BUTTONINACTIVE, CBrowser::eUpdate, 1);
							else
							{
								m_pProperty->GetData().lVal = nTemp;
								PostMessage(GetGlobals().WM_BUTTONINACTIVE);
								bResult = FALSE;
							}
						}
						SysFreeString(vt.bstrVal);
					}
				}
				else
				{
					MAKE_WIDEPTR_FROMTCHAR(wEdit, szTemp);
					m_pProperty->GetData().bstrVal = SysAllocString(wEdit);
					m_pProperty->VarType(VT_BSTR);
					m_pProperty->Clear();
					assert(m_pProperty->GetData().bstrVal);
					if (m_pProperty->GetData().bstrVal)
					{
						HRESULT hResult = m_pProperty->SetValueNoCheck();
						if (SUCCEEDED(hResult))
						{
							if (m_szPriorText)
							{
								delete [] m_szPriorText;
								m_szPriorText = NULL;
							}
							if (szTemp)
							{
								m_szPriorText = new TCHAR[lstrlen(szTemp)+1];
								_tcscpy(m_szPriorText, szTemp);
							}
							PostMessage(GetGlobals().WM_BUTTONINACTIVE, CBrowser::eUpdate, 1);
						}
						else
						{
							SysFreeString(m_pProperty->GetData().bstrVal);
							MAKE_WIDEPTR_FROMTCHAR(wEdit, m_szPriorText);
							m_pProperty->GetData().bstrVal = SysAllocString(wEdit);
							PostMessage(GetGlobals().WM_BUTTONINACTIVE);
							bResult = FALSE;
						}
					}
				}
				return bResult;
			}
			break;

		case CProperty::eNumber:
			{
				LPTSTR szTemp = GetText();
				if (szTemp && 0 == _tcscmp(m_szPriorText, szTemp))
					break;

				if (m_szPriorText)
				{
					delete [] m_szPriorText;
					m_szPriorText = NULL;
				}
				if (szTemp)
				{
					m_szPriorText = new TCHAR[lstrlen(szTemp)+1];
					if (m_szPriorText)
						lstrcpy(m_szPriorText, szTemp);
				}
				MAKE_WIDEPTR_FROMTCHAR(wEdit, m_szPriorText);
				VARIANT vt;					
				vt.vt = VT_BSTR;
				vt.bstrVal = SysAllocString(wEdit);
				assert(vt.bstrVal);
				if (vt.bstrVal)
				{
					if (VT_I4 != m_pProperty->VarType() && VT_I2 != m_pProperty->VarType())
						m_pProperty->VarType(VT_I4);

					long nTemp = m_pProperty->GetData().lVal;
					HRESULT hResult = VariantChangeType(&m_pProperty->GetData(), 
														&vt, 
														0, 
														m_pProperty->VarType());
					if (SUCCEEDED(hResult))
					{
						hResult = m_pProperty->SetValue();
						if (SUCCEEDED(hResult))
							PostMessage(GetGlobals().WM_BUTTONINACTIVE, CBrowser::eUpdate, 1);
						else
						{
							m_pProperty->GetData().lVal = nTemp;
							PostMessage(GetGlobals().WM_BUTTONINACTIVE);
							bResult = FALSE;
						}
					}
					SysFreeString(vt.bstrVal);
				}
			}
			break;

		case CProperty::eString:
			{
				LPTSTR szTemp = GetText();
				if (szTemp && 0 == _tcscmp(m_szPriorText, szTemp))
					break;

				MAKE_WIDEPTR_FROMTCHAR(wEdit, szTemp);
				m_pProperty->GetData().bstrVal = SysAllocString(wEdit);
				assert(m_pProperty->GetData().bstrVal);
				if (m_pProperty->GetData().bstrVal)
				{
					HRESULT hResult = m_pProperty->SetValue();
					if (SUCCEEDED(hResult))
					{
						if (m_szPriorText)
						{
							delete [] m_szPriorText;
							m_szPriorText = NULL;
						}
						if (szTemp)
						{
							m_szPriorText = new TCHAR[lstrlen(szTemp)+1];
							_tcscpy(m_szPriorText, szTemp);
						}
						PostMessage(GetGlobals().WM_BUTTONINACTIVE, CBrowser::eUpdate, 1);
					}
					else
					{
						SysFreeString(m_pProperty->GetData().bstrVal);
						MAKE_WIDEPTR_FROMTCHAR(wEdit, m_szPriorText);
						m_pProperty->GetData().bstrVal = SysAllocString(wEdit);
						PostMessage(GetGlobals().WM_BUTTONINACTIVE);
						bResult = FALSE;
					}
				}
			}
			break;

		case CProperty::ePicture:
			{
				LPCTSTR szEdit = GetText();
				if (szEdit && 0 == _tcscmp(szEdit, _T("")) && m_pProperty->GetData().pdispVal)
				{
					LPDISPATCH pDispatchOld = m_pProperty->GetData().pdispVal;
					m_pProperty->GetData().pdispVal = NULL;
					HRESULT hResult = m_pProperty->SetValue();
					if (SUCCEEDED(hResult))
					{
						DDString strPicture;
						strPicture.LoadString(IDS_PICNONE);
						SetText(strPicture);
						long nResult = pDispatchOld->Release();
						PostMessage(GetGlobals().WM_BUTTONINACTIVE, CBrowser::eUpdate, 1);
					}
					else
					{
						m_pProperty->GetData().pdispVal = pDispatchOld;
						PostMessage(GetGlobals().WM_BUTTONINACTIVE);
						bResult = FALSE;
					}
					Close(TRUE);
				}
			}
			break;

		case CProperty::eColor:
			{
				LPTSTR szEdit = GetText();
				if (szEdit)
				{
					enum STATES
					{
						eStart,
						eBeginAmpersand,
						eH,
						eNumber,
						eEndAmpersand,
						eError,
						eEnd
					} eState = eStart;

					int nLen = lstrlen(szEdit);
					LPTSTR szNumber = new TCHAR[9];
					if (NULL == szNumber)
						break;
					int nPos = 0;
					while (*szEdit)
					{
						if (islower(*szEdit))
							*szEdit = _toupper(*szEdit);
						switch (eState)
						{
						case eStart:
							if ('H' == *szEdit)
								eState = eH;
							else if ('&' == *szEdit)
								eState = eBeginAmpersand;
							else if (*szEdit >= '0' && *szEdit <= 'F')
							{
								eState = eNumber;
								szNumber[nPos] = *szEdit; 
								nPos++;
							}
							else
								eState = eError;
							break;

						case eBeginAmpersand:
							if ('H' == *szEdit)
								eState = eH;
							else if ('&' == *szEdit)
								eState = eError;
							else if (*szEdit >= '0' && *szEdit <= 'F')
							{
								eState = eNumber;
								szNumber[nPos] = *szEdit; 
								nPos++;
							}
							else
								eState = eError;
							break;

						case eH:
							if ('&' == *szEdit)
								eState = eError;
							else if (*szEdit >= '0' && *szEdit <= 'F')
							{
								eState = eNumber;
								szNumber[nPos] = *szEdit; 
								nPos++;
							}
							else
								eState = eError;
							break;

						case eNumber:
							if (*szEdit >= '0' && *szEdit <= 'F' && nPos < 9)
							{
								eState = eNumber;
								szNumber[nPos] = *szEdit; 
								nPos++;
							}
							else if ('&' == *szEdit)
								eState = eEnd;
							else
								eState = eError;
							break;

						case eEndAmpersand:
							szNumber[nPos] = NULL;
							break;

						case eError:
							break;

						case eEnd:
							szNumber[nPos] = NULL;
							break;
						}
						szEdit++;
					}
					if (nPos > 0 && eState == eEnd)
					{
						szNumber[nPos] = NULL;
						long nTemp = m_pProperty->GetData().lVal;
						m_pProperty->GetData().lVal = ahextol(szNumber);
						HRESULT hResult = m_pProperty->SetValue();
						if (SUCCEEDED(hResult))
							PostMessage(GetGlobals().WM_BUTTONINACTIVE, CBrowser::eUpdate, 1);
						else
						{
							m_pProperty->GetData().lVal = nTemp;
							PostMessage(GetGlobals().WM_BUTTONINACTIVE);
							bResult = FALSE;
						}
					}
					delete [] szNumber;
				}
			}
			break;
		}
	}
	CATCH
	{
		bResult = FALSE;
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return bResult;
}

//
// OnReturn
//

BOOL CBrowserEdit::OnReturn(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		if (NULL == m_pProperty)
			return FALSE;

		switch (m_pProperty->GetEditType())
		{
		case CProperty::eVariant:
			if (UpdateText())
				Close(TRUE);
			return TRUE;

		case CProperty::eNumber:
		case CProperty::eString:
			if (UpdateText())
				Close(TRUE);
			return TRUE;

		case CProperty::eColor:
			if (UpdateText())
			{
				ColorWnd()->SendMessage(nMsg, wParam, lParam);
				m_bChildActive = FALSE;
				Close(TRUE);
			}
			return TRUE;

		case CProperty::ePicture:
			if (UpdateText())
			{
				m_bChildActive = FALSE;
				Close(TRUE);
			}
			return TRUE;
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
// OnEscape
//

BOOL CBrowserEdit::OnEscape(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		if (NULL == m_pProperty)
			return FALSE;

		switch (m_pProperty->GetEditType())
		{
		case CProperty::eNumber:
		case CProperty::eString:
			Close(TRUE);
			return TRUE;

		case CProperty::eBool:
		case CProperty::eEnum:
			BrowserCombo()->SendMessage(nMsg, wParam, lParam);
			m_bChildActive = FALSE;
			Close(TRUE);
			return TRUE;

		case CProperty::eColor:
			ColorWnd()->SendMessage(nMsg, wParam, lParam);
			m_bChildActive = FALSE;
			Close(TRUE);
			return TRUE;
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
// WindowProc
//

LRESULT CBrowserEdit::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		if (GetGlobals().WM_BUTTONINACTIVE == nMsg)
		{
			try
			{
				Close();
				m_bChildActive = FALSE;
				return ::SendMessage(GetParent(m_hWnd), nMsg, wParam, lParam);
			}
			catch (...)
			{
				assert(FALSE);
			}
		}

		if (GetGlobals().WM_HIDE == nMsg)
		{
			try
			{
				DWORD dwStyle = GetWindowLong(m_hWnd, GWL_STYLE);
				if (WS_VISIBLE & dwStyle)
					ShowWindow(SW_HIDE);
				::SendMessage(GetParent(m_hWnd), GetGlobals().WM_BUTTONINACTIVE, wParam, 0);
			}
			catch (...)
			{
				assert(FALSE);
			}
			return 0;
		}

		switch (nMsg)
		{
		case WM_KILLFOCUS:
			try
			{
				if (m_pProperty && CProperty::eColor != m_pProperty->GetEditType())
					UpdateText();
				Close(TRUE);
			}
			catch (...)
			{
				assert(FALSE);
			}
			break;

		case WM_KEYDOWN:
			{
				TRACE(5, "WM_KEYDOWN\n");
				if (NULL == m_pProperty)
					break;

				switch (wParam)
				{
				case VK_TAB:
					switch (m_pProperty->GetEditType())
					{
					case CProperty::ePicture:
					case CProperty::eNumber:
					case CProperty::eString:
					case CProperty::eVariant:
						if (UpdateText())
						{
							DWORD dwStyle = GetWindowLong(m_hWnd, GWL_STYLE);
							if (WS_VISIBLE & dwStyle)
								ShowWindow(SW_HIDE);
						}
					}
					break;

				case VK_UP:
					switch (m_pProperty->GetEditType())
					{
					case CProperty::ePicture:
					case CProperty::eNumber:
					case CProperty::eString:
					case CProperty::eVariant:
						{
							if (UpdateText())
							{
								DWORD dwStyle = GetWindowLong(m_hWnd, GWL_STYLE);
								if (WS_VISIBLE & dwStyle)
									ShowWindow(SW_HIDE);
								m_pBrowser->PostMessage(nMsg, wParam, lParam);
							}
						}
						break;

					case CProperty::eBool:
					case CProperty::eEnum:
						if (m_bChildActive)
						{
							BrowserCombo()->SendMessage(nMsg, wParam, lParam);
							return 0;
						}
						break;

					case CProperty::eColor:
						if (m_bChildActive)
						{
							ColorWnd()->SendMessage(nMsg, wParam, lParam);
							return 0;
						}
						break;
					}
					break;

				case VK_DOWN:
					switch (m_pProperty->GetEditType())
					{
					case CProperty::ePicture:
					case CProperty::eNumber:
					case CProperty::eString:
					case CProperty::eVariant:
						{
							if (UpdateText())
							{
								DWORD dwStyle = GetWindowLong(m_hWnd, GWL_STYLE);
								if (WS_VISIBLE & dwStyle)
									ShowWindow(SW_HIDE);
								m_pBrowser->PostMessage(nMsg, wParam, lParam);
							}
						}
						break;

					case CProperty::eBool:
					case CProperty::eEnum:
						if (m_bChildActive)
						{
							BrowserCombo()->SendMessage(nMsg, wParam, lParam);
							return 0;
						}
						break;

					case CProperty::eColor:
						if (m_bChildActive)
						{
							ColorWnd()->SendMessage(nMsg, wParam, lParam);
							return 0;
						}
						break;
					}
					break;
				}
			}
			break;

		case WM_KEYUP:
			{
				switch (wParam)
				{
				case VK_ESCAPE:
					return 0;
				case VK_RETURN:
					return 0;
				}
			}
			break;

		case WM_PASTE:
		case WM_CUT:
			{
				LRESULT lResult = FWnd::WindowProc(nMsg, wParam, lParam);
				UpdateText();
				return lResult;
			}
			break;

		case WM_LBUTTONDBLCLK:
			{
				POINT pt = {LOWORD(lParam), HIWORD(lParam)};
				ClientToScreen(pt);
				m_pBrowser->ScreenToClient(pt);
				m_pBrowser->OnLButtonDblClk(wParam, pt);
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
// Button
//

Button::Button()
{
	m_szClassName = _T("BUTTON");
	m_bnType = eScroll;
}

Button::~Button()
{
	UnsubClass();
}

BOOL Button::Create(CBrowser* pBrowser, CRect rc)
{
	FWnd::Create(_T(""), 
			     WS_CHILD|BS_PUSHBUTTON|BS_OWNERDRAW,
			     rc.left,
			     rc.right,
			     rc.Width(),
			     rc.Height(),
			     pBrowser->hWnd(),
				 (HMENU)CBrowser::ePushButton);
	HFONT hFont = pBrowser->GetFont();
	if (hFont)
		SetFont(hFont);
	SubClassAttach();
	return m_hWnd ? TRUE : FALSE;
}

void Button::Type (BUTTONTYPE bnType)
{
	m_bnType = bnType;
}

LRESULT Button::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	switch (nMsg)
	{
	case WM_DRAWITEM:
		{
			LPDRAWITEMSTRUCT pDis = (LPDRAWITEMSTRUCT)lParam;
			switch (m_bnType)
			{
			case eEllipse:
				DrawFrameControl(pDis->hDC, &pDis->rcItem, DFC_BUTTON, DFCS_BUTTONPUSH);
				SetTextColor(pDis->hDC, GetSysColor(COLOR_BTNTEXT));
				DrawText(pDis->hDC, 
						 _T("..."), 
						 lstrlen(_T("...")), 
						 &pDis->rcItem, 
						 DT_CENTER|DT_VCENTER|DT_SINGLELINE);
				break;

			case eScroll:
				DrawFrameControl(pDis->hDC, &pDis->rcItem, DFC_SCROLL, DFCS_SCROLLDOWN);
				break;

			case eCancel:
				DrawFrameControl(pDis->hDC, &pDis->rcItem, DFC_SCROLL, DFCS_SCROLLDOWN);
				DrawIcon(pDis->hDC, pDis->rcItem.left, pDis->rcItem.top, GetGlobals().GetCancelIcon());
				break;
			}
		}
		break;
	}
	return FWnd::WindowProc(nMsg, wParam, lParam);
}

//
// CBrowser
//

CBrowser::CBrowser(CDesignerPage* pDesigner)
	: m_pTypeInfo(NULL),
	  m_pPerPropertyBrowsing(NULL),
	  m_pCurrentProperty(NULL),
	  m_pPrevProperty(NULL),
	  m_pDesigner(pDesigner),
	  m_pObject(NULL)
{
	m_szClassName = _T("LISTBOX");
	
	m_pFF = new CFlickerFree;
	assert(m_pFF);
	
	m_hPen = CreatePen(PS_SOLID, 1, RGB(128, 128, 128)); 
	assert(m_hPen);
	
	m_hBrushBackground = CreateSolidBrush(COLOR_WINDOW+1);
	assert(m_hBrushBackground);

	m_strTrue.LoadString(IDS_TRUE);
	m_strFalse.LoadString(IDS_FALSE);
	m_strName.LoadString(IDS_NAME);
	m_strCaption.LoadString(IDS_CAPTION);
	
	m_bButtonPressed = FALSE;
	m_bInClick = FALSE;
	m_nCurrentIndex = -1;
	m_nPrevIndex = -1;
}

CBrowser::~CBrowser()
{
	try
	{
		UnsubClass();
		if (m_pTypeInfo)
			m_pTypeInfo->Release();
		if (m_pObject)
			m_pObject->Release();
		if (m_pPerPropertyBrowsing)
			m_pPerPropertyBrowsing->Release();
		delete m_pFF;
		BOOL bResult = DeletePen(m_hPen);
		assert(bResult);
		CleanUpEnums();
		if (m_hBrushBackground)
		{
			bResult = DeleteBrush(m_hBrushBackground);
			assert(bResult);
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// CleanUpEnums
//

void CBrowser::CleanUpEnums()
{
	try
	{
		int nCount = m_aEnums.GetSize();
		for (int nIndex = 0; nIndex < nCount; nIndex++)
			delete m_aEnums.GetAt(nIndex);
		m_aEnums.RemoveAll();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// FindEnum
//

CEnum* CBrowser::FindEnum(BSTR bstrName)
{
	try
	{
		CEnum* pEnum;
		int nCount = m_aEnums.GetSize();
		for (int nIndex = 0; nIndex < nCount; nIndex++)
		{
			pEnum = m_aEnums.GetAt(nIndex);
			if (0 == _wcsicmp(bstrName, pEnum->EnumName()))
				return pEnum;
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
// Clear
//

void CBrowser::Clear()
{
	try
	{
		::SendMessage(GetDlgItem(::GetParent(m_hWnd), IDC_ST_DESC), WM_SETTEXT, 0, (LPARAM)_T(""));
		try
		{
			if (m_pPerPropertyBrowsing)
			{
				m_pPerPropertyBrowsing->Release();
				m_pPerPropertyBrowsing = NULL;
			}
		}
		CATCH
		{
			m_pPerPropertyBrowsing = NULL;
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
		}
		try
		{
			if (m_pObject)
			{
				m_pObject->Release();
				m_pObject = NULL;
			}
		}
		CATCH
		{
			m_pObject = NULL;
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
		}
		try
		{
			if (m_pTypeInfo)
			{
				m_pTypeInfo->Release();
				m_pTypeInfo = NULL;
			}
		}
		CATCH
		{
			m_pTypeInfo = NULL;
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
		}
		m_pPrevProperty = NULL;
		m_nCurrentIndex = -1;
		m_pCurrentProperty = NULL;
		m_theEdit.Hide();
		m_thePushButton.ShowWindow(SW_HIDE);
		SendMessage(LB_RESETCONTENT);
		CleanUpEnums();
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// ShowButton
//

void CBrowser::ShowButton(CRect rcButton, Button::BUTTONTYPE bnType)
{
	try
	{
		rcButton.left = rcButton.right - rcButton.Height();
		m_thePushButton.Type(bnType);
		m_thePushButton.SetWindowPos(NULL, 
									 rcButton.left - 1, 
									 rcButton.top - 2, 
									 0, 
									 0, 
									 SWP_NOZORDER|SWP_SHOWWINDOW|SWP_NOSIZE|SWP_NOACTIVATE);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// SubClassAttach
//

void CBrowser::SubClassAttach(HWND hWnd)
{
	try
	{
		FWnd::SubClassAttach(hWnd);
		m_theEdit.Create(this);
		m_nButtonHeight = SendMessage(LB_GETITEMHEIGHT) + 1;
		CRect rc(0, m_nButtonHeight, 0, m_nButtonHeight);
		m_thePushButton.Create(this, rc);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// SetEditFocus
//

void CBrowser::SetEditFocus(int nVirtKey)
{
	try
	{
		if (-1 == m_nCurrentIndex)
			return;

		if (NULL == m_pCurrentProperty)
			return;

		LRESULT lResult = SendMessage(LB_GETITEMRECT, m_nCurrentIndex, (LPARAM)&m_rcCurrentItem);
		if (LB_ERR == lResult)
			return;

		m_rcCurrentItem.left += (m_rcCurrentItem.Width() + 2) / 2;
		m_rcCurrentItem.Inflate(-1, -1);

		switch (m_pCurrentProperty->GetEditType())
		{
		case CProperty::eString:
		case CProperty::eNumber:
			m_theEdit.Show(m_rcCurrentItem, m_pCurrentProperty, FALSE, nVirtKey);
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
// WindowProc
//

LRESULT CBrowser::WindowProc(UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	try
	{
		if (GetGlobals().WM_BUTTONINACTIVE == nMsg)
		{
			switch (wParam)
			{
			case eUpdate:
				m_pObjectChanged->ObjectChanged();
				InvalidateItem(m_nCurrentIndex, TRUE);
				break;

			case eError:
				MessageBox(IDS_FAILEDTOSETPROPERTY);
				break;
			}
			if (m_bButtonPressed)
				m_bButtonPressed = FALSE;
			return 0;
		}

		if (GetGlobals().WM_SETEDITFOCUS == nMsg)
		{
			SetEditFocus(wParam);
			return 0;
		}

		switch (nMsg)
		{
		case WM_SIZE:
			m_theEdit.Hide();
			break;

		case WM_SHOWEDIT:
			m_theEdit.Show(m_rcCurrentItem, m_pCurrentProperty, wParam);
			break;

		case WM_SETFOCUS:
			TRACE(5, "WM_SETFOCUS\n");
			break;

		case WM_HELP:
			{
				int nIndex = SendMessage(LB_GETCURSEL);
				if (LB_ERR != nIndex)
				{
					CProperty* pProperty = (CProperty*)SendMessage(LB_GETITEMDATA, nIndex);
					if (pProperty)
					{
						WinHelp(GetParent(m_hWnd), HELP_CONTEXT, pProperty->ContextId()); 
						return TRUE;
					}
				}
			}
			break;

		case WM_CTLCOLORSTATIC:
			SetBkColor((HDC)wParam, GetSysColor(COLOR_WINDOW));
			return (LRESULT)m_hBrushBackground;

		case WM_VSCROLL:
			UpdatePrevSelection(-1);
			break;

		case WM_LBUTTONDBLCLK:
			{
				POINT pt;
				pt.x = LOWORD(lParam);
				pt.y = HIWORD(lParam);
				OnLButtonDblClk(wParam, pt);
			}
			break;

		case WM_CHARTOITEM:
		case WM_VKEYTOITEM:
			return OnKeyToItem(LOWORD(wParam), HIWORD(wParam));

		case WM_DRAWITEM:
			switch (wParam)
			{
			case IDC_LST_PROPERTIES:
				OnDrawItem((LPDRAWITEMSTRUCT)lParam);
				break;

			case ePushButton:
				return m_thePushButton.SendMessage(nMsg, wParam, lParam);
			}
			break;

		case WM_KILLFOCUS:
			if (!IsWindow(m_hWnd) || (IsWindow((HWND)wParam) && !IsChild(m_hWnd, (HWND)wParam)))
			{
				m_nPrevIndex = -1;
				m_pPrevProperty = NULL;
			}
			break;

		case WM_WINDOWPOSCHANGED:
			UpdatePrevSelection(-1);
			break;

		case WM_DELETEITEM:
			OnDeleteItem((LPDELETEITEMSTRUCT)lParam);
			break;

		case WM_KEYDOWN:
			switch (wParam)
			{
			case VK_F4:
				if (!m_bButtonPressed)
					OnButtonHit(m_nCurrentIndex, m_pCurrentProperty, m_rcCurrentItem);
				else if (m_theEdit.ChildActive())
					m_theEdit.Close();
				m_bButtonPressed = !m_bButtonPressed;
				return 0;
			}
			break;

		case WM_CHAR:
			PostMessage(GetGlobals().WM_SETEDITFOCUS, wParam);
			break;

		case WM_COMPAREITEM:
			return OnCompareItem((LPCOMPAREITEMSTRUCT)lParam);

		case WM_COMMAND:
			switch (HIWORD(wParam))
			{
			case LBN_SELCHANGE:
				OnSelChange();
				break;

			case BN_CLICKED:
				switch (LOWORD(wParam))
				{
				case ePushButton:
					if (!m_bButtonPressed)
						OnButtonHit(m_nCurrentIndex, m_pCurrentProperty, m_rcCurrentItem);
					else if (m_theEdit.ChildActive())
						m_theEdit.Close();
					m_bButtonPressed = !m_bButtonPressed;
					break;
				}
				break;
			}
			break;

		case LB_FINDSTRING:
			{
				int nCount = SendMessage(LB_GETCOUNT);
				CProperty* pProperty;
				for (int nIndex = 0; nIndex < nCount; nIndex++)
				{
					pProperty = reinterpret_cast<CProperty*>(SendMessage(LB_GETITEMDATA, nIndex));
					if (pProperty && 0 == _tcsicmp(pProperty->Name(), (LPCTSTR)lParam))
						return nIndex;
				}
				return LB_ERR;
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
// OnCompareItem
//

LRESULT CBrowser::OnCompareItem(LPCOMPAREITEMSTRUCT pCmpItem)
{
	LRESULT lResult = 0;
	try
	{
		if (0xFFFFFFFF == pCmpItem->itemID1)
			return -1;

		CProperty* pProperty1 = reinterpret_cast<CProperty*>(pCmpItem->itemData1);
		if (NULL == pProperty1 || NULL == pProperty1->Name())
			return -1;

		CProperty* pProperty2 = reinterpret_cast<CProperty*>(pCmpItem->itemData2);
		if (NULL == pProperty2 || NULL == pProperty2->Name())
			return 1;

		lResult = _tcsicmp(pProperty1->Name(), pProperty2->Name());
		if (0 == lResult)
			return lResult;
		if (0 == _tcsicmp(pProperty1->Name(), m_strName))
			return -1;
		if (0 == _tcsicmp(pProperty2->Name(), m_strName))
			return 1;
		if (0 == _tcsicmp(pProperty1->Name(), m_strCaption))
			return -1;
		if (0 == _tcsicmp(pProperty2->Name(), m_strCaption))
			return 1;
		if (0 == _tcsicmp(pProperty1->Name(), "ID"))
			return -1;
		if (0 == _tcsicmp(pProperty2->Name(), "ID"))
			return 1;
		if (0 == _tcsicmp(pProperty1->Name(), "Alignment"))
			return -1;
		if (0 == _tcsicmp(pProperty2->Name(), "Alignment"))
			return 1;
		if (0 == _tcsicmp(pProperty1->Name(), "CaptionPosition"))
			return -1;
		if (0 == _tcsicmp(pProperty2->Name(), "CaptionPosition"))
			return 1;
		if (0 == _tcsicmp(pProperty1->Name(), "ControlType"))
			return -1;
		if (0 == _tcsicmp(pProperty2->Name(), "ControlType"))
			return 1;
		if (0 == _tcsicmp(pProperty1->Name(), "DisplayHandles"))
			return -1;
		if (0 == _tcsicmp(pProperty2->Name(), "DisplayHandles"))
			return 1;
		if (0 == _tcsicmp(pProperty1->Name(), "DockingArea"))
			return -1;
		if (0 == _tcsicmp(pProperty2->Name(), "DockingArea"))
			return 1;
		if (0 == _tcsicmp(pProperty1->Name(), "Enabled"))
			return -1;
		if (0 == _tcsicmp(pProperty2->Name(), "Enabled"))
			return 1;
		if (0 == _tcsicmp(pProperty1->Name(), "Flags"))
			return -1;
		if (0 == _tcsicmp(pProperty2->Name(), "Flags"))
			return 1;
		if (0 == _tcsicmp(pProperty1->Name(), "GrabHandleStyle"))
			return -1;
		if (0 == _tcsicmp(pProperty2->Name(), "GrabHandleStyle"))
			return 1;
		if (0 == _tcsicmp(pProperty1->Name(), "PageStyle"))
			return -1;
		if (0 == _tcsicmp(pProperty2->Name(), "PageStyle"))
			return 1;
		if (0 == _tcsicmp(pProperty1->Name(), "Type"))
			return -1;
		if (0 == _tcsicmp(pProperty2->Name(), "Type"))
			return 1;
		if (0 == _tcsicmp(pProperty1->Name(), "Style"))
			return -1;
		if (0 == _tcsicmp(pProperty2->Name(), "Style"))
			return 1;
		if (0 == _tcsicmp(pProperty1->Name(), "SubBand"))
			return -1;
		if (0 == _tcsicmp(pProperty2->Name(), "SubBand"))
			return 1;
		if (0 == _tcsicmp(pProperty1->Name(), "Visible"))
			return -1;
		if (0 == _tcsicmp(pProperty2->Name(), "Visible"))
			return 1;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return lResult;
}

//
// UpdatePrevSelection
//

void CBrowser::UpdatePrevSelection(int nIndex)
{
	try
	{
		m_nPrevIndex = m_nCurrentIndex;
		m_pPrevProperty = m_pCurrentProperty;
		m_nCurrentIndex = nIndex;

		LRESULT lResult = SendMessage(LB_GETITEMDATA, m_nCurrentIndex);
		if (LB_ERR == lResult)
		{
			m_nCurrentIndex = -1;
			m_pCurrentProperty = NULL;
			m_thePushButton.ShowWindow(SW_HIDE);
		}
		else
			m_pCurrentProperty = reinterpret_cast<CProperty*>(lResult);

		if (m_pPrevProperty)
		{
			//
			// Clean up Prev Selection
			//

			switch (m_pPrevProperty->GetEditType())
			{
			case CProperty::eFont:
			case CProperty::eBitmap:
			case CProperty::ePicture:
			case CProperty::eBool:
			case CProperty::eEnum:
			case CProperty::eColor:
			case CProperty::eFlags:
			case CProperty::eEdit:
			case CProperty::eComboList:
			case CProperty::eShortCut:
				m_theEdit.Close(TRUE);
				m_thePushButton.ShowWindow(SW_HIDE);
				break;
			}
			InvalidateItem(m_nPrevIndex);
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnSelChange
//

void CBrowser::OnSelChange()
{
	try
	{
		LRESULT lResult;
		m_bButtonPressed = FALSE;
		int nIndex = SendMessage(LB_GETCURSEL);
		if (nIndex != m_nCurrentIndex)
			UpdatePrevSelection(nIndex);

		//
		// Handle the current selection
		//

		lResult = SendMessage(LB_GETITEMRECT, m_nCurrentIndex, (LPARAM)&m_rcCurrentItem);
		if (LB_ERR == lResult)
		{
			::SendMessage(GetDlgItem(GetParent(m_hWnd), IDC_ST_DESC), WM_SETTEXT, 0, (LPARAM)_T(""));
			return;
		}

		m_rcCurrentItem.left += (m_rcCurrentItem.Width() + 2) / 2;
		m_rcCurrentItem.Inflate(-1, -1);

		if (NULL == m_pCurrentProperty)
		{
			:: SendMessage(GetDlgItem(GetParent(m_hWnd), IDC_ST_DESC), WM_SETTEXT, 0, (LPARAM)_T(""));
			return;
		}

		switch (m_pCurrentProperty->GetEditType())
		{
		case CProperty::eVariant:
		case CProperty::eString:
		case CProperty::eNumber:
			PostMessage(WM_SHOWEDIT);
			break;

		case CProperty::ePicture:
			ShowButton(m_rcCurrentItem, Button::eEllipse);
			PostMessage(WM_SHOWEDIT);
			break;

		case CProperty::eBitmap:
			ShowButton(m_rcCurrentItem, Button::eEllipse);
			break;

		case CProperty::eFont:
			ShowButton(m_rcCurrentItem, Button::eEllipse);
			break;

		case CProperty::eColor:
		case CProperty::eEdit:
		case CProperty::eBool:
		case CProperty::eEnum:
		case CProperty::eFlags:
		case CProperty::eShortCut:
		case CProperty::eComboList:
			{
				ShowButton(m_rcCurrentItem, Button::eScroll);
				PostMessage(WM_SHOWEDIT);
			}
			break;

		default:
			m_theEdit.Hide();
			break;
		}

		if (m_pCurrentProperty)
		{
			MAKE_TCHARPTR_FROMWIDE(szDocString, m_pCurrentProperty->DocString());
			::SendMessage(GetDlgItem(GetParent(m_hWnd), IDC_ST_DESC), WM_SETTEXT, 0, (LPARAM)szDocString);
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnDrawItem
//

void CBrowser::OnDrawItem(LPDRAWITEMSTRUCT pDrawItem)
{
	try
	{
		CProperty* pProperty = (CProperty*)pDrawItem->itemData;
		if (NULL == pProperty)
			return;

		DDString strTemp;
		CRect rc;
		CRect rcBound = pDrawItem->rcItem;
		HFONT hFont = NULL;
		HFONT hFontOld = NULL;

		//
		// Setup flicker free drawing
		//

		HDC hDC;
		if (m_pFF)
		{
			hDC = m_pFF->RequestDC(pDrawItem->hDC, rcBound.Width(), rcBound.Height());
			if (NULL == hDC)
				hDC = pDrawItem->hDC;
			else
				rcBound.Offset(-rcBound.left, -rcBound.top);
		}
		else
			hDC = pDrawItem->hDC;

		hFont = (HFONT)SendMessage(WM_GETFONT);
		hFontOld = SelectFont(hDC, hFont);

		CRect rcPropertyName = rcBound;
		CRect rcPropertyValue = rcBound;
		rcPropertyName.right = rcPropertyName.left + rcPropertyName.Width() / 2;
		rcPropertyValue.left = rcPropertyName.right + 2;

		FillRect(hDC, &rcBound, (HBRUSH)(1+COLOR_WINDOW));

		if (pDrawItem->itemID >= 0 && 
			(pDrawItem->itemAction & (ODA_DRAWENTIRE | ODA_SELECT)))
		{
			SetBkMode(hDC,TRANSPARENT);

			BOOL bDisabled = !IsWindowEnabled(m_hWnd);
			if (!bDisabled && ODS_SELECTED & pDrawItem->itemState)
			{
				FillRect(hDC, &rcPropertyName, (HBRUSH)(1+COLOR_HIGHLIGHT));
				SetTextColor(hDC, GetSysColor(COLOR_HIGHLIGHTTEXT));
			}
			else
				SetTextColor(hDC, GetSysColor(COLOR_BTNTEXT));

			rcPropertyName.left += 2;

			//
			// Draw the grid lines
			//

			HPEN hPenOld = SelectPen(hDC, m_hPen);
			MoveToEx(hDC, rcBound.left, rcBound.bottom - 1, 0); 
			LineTo(hDC, rcBound.right, rcBound.bottom - 1);
			MoveToEx(hDC, rcPropertyName.right, rcPropertyName.top, 0); 
			LineTo(hDC, rcPropertyName.right, rcPropertyName.bottom - 1);
			SelectPen(hDC, hPenOld);
			
			//
			// Draw the property name
			//
			
			DrawText(hDC, 
					 pProperty->Name(), 
					 lstrlen(pProperty->Name()), 
					 &rcPropertyName, 
					 DT_LEFT|DT_VCENTER|DT_SINGLELINE);

			//
			// Draw the value
			//

			SetTextColor(hDC, GetSysColor(COLOR_BTNTEXT));
			UINT nFormat = DT_LEFT|DT_VCENTER|DT_SINGLELINE;
			switch (pProperty->GetEditType())
			{
			case CProperty::eEdit:
			case CProperty::eShortCut:
			case CProperty::eComboList:
			case CProperty::eBitmap:
				try
				{
					switch (pProperty->GetEditType())
					{
					case CProperty::eEdit:
						strTemp.LoadString(IDS_SHORTCUT);
						break;
					case CProperty::eShortCut:
						strTemp.LoadString(IDS_LIST);
						break;
					case CProperty::eComboList:
						strTemp.LoadString(IDS_COMBOLIST);
						break;
					case CProperty::eBitmap:
						strTemp.LoadString(IDS_BITMAPS);
						break;
					}
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
				break;

			case CProperty::eVariant:
				try
				{
					switch (pProperty->VarType())
					{
					case VT_BSTR:
						if (NULL != pProperty->GetData().bstrVal && NULL != *pProperty->GetData().bstrVal)
						{
							MAKE_TCHARPTR_FROMWIDE(szValue, pProperty->GetData().bstrVal);
							strTemp.Format(_T("%s"), szValue);
						}
						break;

					case VT_I2:
						strTemp.Format(_T("%li"), pProperty->GetData().iVal);
						break;

					case VT_INT:
					case VT_UINT:
					case VT_I4:
						strTemp.Format(_T("%li"), pProperty->GetData().lVal);
						break;
					}
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
				break;

			case CProperty::eString:
				try
				{
					if (NULL != pProperty->GetData().bstrVal && NULL != *pProperty->GetData().bstrVal)
					{
						MAKE_TCHARPTR_FROMWIDE(szValue, pProperty->GetData().bstrVal);
						strTemp.Format(_T("%s"), szValue);
					}
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
				break;

			case CProperty::eNumber:
				try
				{
					long nNumber = 0;
					switch (pProperty->GetData().vt)
					{
					case VT_I2:
						nNumber = pProperty->GetData().iVal;
						break;

					case VT_INT:
					case VT_UINT:
					case VT_I4:
						nNumber = pProperty->GetData().lVal;
						break;
					}
					strTemp.Format(_T("%li"), nNumber);
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
				break;

			case CProperty::eFlags:
				try
				{
					strTemp.Format(_T("%li"), pProperty->GetData().lVal);
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
				break;

			case CProperty::eColor:
				try
				{
					//
					// Draw a box around the Color
					//

					BOOL bResult;
					CRect rcColor = rcPropertyValue;
					rcColor.right = rcColor.left + rcColor.Height() - 2;
					rcColor.bottom -= 2;
					rcColor.Inflate(-1, -1);
					DrawBox(hDC, rcColor, GetStockPen(BLACK_PEN));

					//
					// Fill the box with the color
					//

					rcColor.left += 1;
					rcColor.top += 1;
					COLORREF crColor;
					OleTranslateColor(pProperty->GetData().lVal,
									  NULL,
									  &crColor);
					HBRUSH hColor = CreateSolidBrush(crColor);
					if (hColor)
					{
						FillRect(hDC, &rcColor, hColor);
						bResult = DeleteBrush(hColor);
						assert(bResult);
					}

					//
					// Format the color text
					//

					rcPropertyValue.left = rcPropertyValue.left + rcColor.Height() + 6;
					strTemp.Format(_T("&H%08X&"), pProperty->GetData().lVal);
					nFormat |= DT_NOPREFIX;
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
				break;

			case CProperty::eFont:
				try
				{
					LOGFONT lf;
					GetObject(pProperty->GetFont(), sizeof(LOGFONT), &lf);
					strTemp.Format(_T("%s"), lf.lfFaceName);
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
				break;

			case CProperty::ePicture:
				try
				{
					UINT nId = IDS_PICNONE;
					switch (pProperty->GetPictureType())
					{
					case PICTYPE_BITMAP:
						nId = IDS_PICBITMAP;
						break;

					case PICTYPE_ICON:
						nId = IDS_PICICON;
						break;

					case PICTYPE_METAFILE:
						nId = IDS_PICMETAFILE;
						break;

					case PICTYPE_ENHMETAFILE:
						nId = IDS_PICENHMETAFILE;
						break;
					}
					strTemp.LoadString(nId);
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
				break;

			case CProperty::eBool:
				try
				{
					if (VARIANT_TRUE == pProperty->GetData().boolVal)
						strTemp.Format(_T("%s"), m_strTrue);
					else
						strTemp.Format(_T("%s"), m_strFalse);
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
				break;

			case CProperty::eEnum:
				try
				{
					CEnum* pEnum = pProperty->GetEnum();
					assert(pEnum);
					if (pEnum)
					{
						int nIndex;
						switch (pProperty->VarType())
						{
						case VT_I4:
							{
								if (pProperty->GetData().lVal < 0)
									pProperty->GetData().lVal = 0;
			
								EnumItem* pEnumItem = pEnum->ByValue(pProperty->GetData().lVal, nIndex);
								assert(pEnumItem);
								if (pEnumItem)
									strTemp.Format(_T("%s"), pEnumItem->m_strDisplay);
							}
							break;

						case VT_BSTR:
							{
								if (pProperty->GetData().bstrVal && *pProperty->GetData().bstrVal)
								{
									MAKE_TCHARPTR_FROMWIDE(szValue, pProperty->GetData().bstrVal);
									strTemp.Format(_T("%s"), szValue);
								}
								else
									strTemp.LoadString(IDS_NONE);
							}
							break;
						}
					}
				}
				CATCH
				{
					assert(FALSE);
					REPORTEXCEPTION(__FILE__, __LINE__)
				}
				break;
			}

			DrawText(hDC, 
					 strTemp, 
					 strTemp.GetLength(), 
					 &rcPropertyValue, 
					 nFormat);

			if (hDC != pDrawItem->hDC)
			{
				m_pFF->Paint(pDrawItem->hDC,
							 pDrawItem->rcItem.left,
							 pDrawItem->rcItem.top);
				SelectFont(hDC, hFontOld);
			}
			if ((pDrawItem->itemAction & ODA_FOCUS))
				DrawFocusRect(pDrawItem->hDC, &(pDrawItem->rcItem));
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnDeleteItem
//

void CBrowser::OnDeleteItem(LPDELETEITEMSTRUCT pDeleteItem)
{
	try
	{
		CProperty* pProperty = (CProperty*)pDeleteItem->itemData;
		delete pProperty;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// OnButtonHit
//

BOOL CBrowser::OnButtonHit(int nIndex, CProperty* pProperty, CRect& rcItem)
{
	try
	{
		CRect rcText = rcItem;
		switch (pProperty->GetEditType())
		{
		case CProperty::ePicture:
			{
				m_bButtonPressed = TRUE;

				DDString strPictureTypes;
				strPictureTypes.LoadString(IDS_LOADPICTURETYPES);

				CFileDialog dlgFile(IDS_LOADPICT,
								TRUE,
								_T("*.bmp;*.ico;*.cur,*.gif;*.jpg"),
								_T("*.bmp;*.ico;*.cur,*.gif;*.jpg"),
								OFN_PATHMUSTEXIST|OFN_FILEMUSTEXIST|OFN_HIDEREADONLY,
								strPictureTypes,
								m_hWnd);
		
				if (IDOK == dlgFile.DoModal())
				{
					LPDISPATCH pPicOld = pProperty->GetData().pdispVal;

					pProperty->GetData().pdispVal = (LPDISPATCH)OleLoadPictureHelper(dlgFile.GetFileName());
					if (NULL == pProperty->GetData().pdispVal)
					{
						MessageBox (IDS_ERR_LOADINGIMAGEFAILED, MB_ICONSTOP);
						return FALSE;
					}
					HRESULT hResult = pProperty->SetValue();
					if (SUCCEEDED(hResult))
					{
						if (pPicOld)
							pPicOld->Release();
						m_pObjectChanged->ObjectChanged();
						InvalidateItem(m_nCurrentIndex, TRUE);
					}
					else
					{
						pProperty->GetData().pdispVal = pPicOld;
						m_theEdit.Close(TRUE);
						MessageBox(IDS_FAILEDTOSETPROPERTY);
					}
				}
				SetFocus();
			}
			break;

		case CProperty::eBitmap:
			{
				//
				// TODO
				//
				// This is a cheese way of getting to the icon for tools in ActiveBar
				// Should be fixed someday
				//

				EnableWindow(GetParent(m_hWnd), FALSE);
				m_bButtonPressed = TRUE;
				
				CIconEditor dlgIconEditor((ITool*)pProperty->Driver());
				if (IDOK == dlgIconEditor.DoModal(m_hWnd))
				{
					m_pObjectChanged->ObjectChanged();
					InvalidateItem(m_nCurrentIndex, TRUE);
				}
				EnableWindow(GetParent(m_hWnd), TRUE);
				SetFocus();
			}
			break;

		case CProperty::eFont:
			{
				EnableWindow(GetParent(m_hWnd), FALSE);
				m_bButtonPressed = TRUE;

				if (NULL == pProperty->GetData().pdispVal)
					break;

				CFontDialog theFont(pProperty->GetFont());
				if (theFont.DoModal())
				{
					pProperty->SetFont(theFont.GetFont());
					HRESULT hResult = pProperty->SetValue();
					if (SUCCEEDED(hResult))
					{
						m_pObjectChanged->ObjectChanged();
						InvalidateItem(m_nCurrentIndex, TRUE);
					}
				}
				EnableWindow(GetParent(m_hWnd), TRUE);
				SetFocus();
			}
			break;

		case CProperty::eColor:
			rcText.left = rcText.left + rcText.Height();
		case CProperty::eEdit:
		case CProperty::eEnum:
		case CProperty::eBool:
		case CProperty::eComboList:
		case CProperty::eShortCut:
		case CProperty::eFlags:
			m_theEdit.Show(rcText, m_pCurrentProperty, TRUE);
			break;
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
// OnKeyToItem
//

LRESULT CBrowser::OnKeyToItem(int nVirtKey, int nCaretPos)
{
	switch (nVirtKey)
	{
	case VK_UP:
		{
			if (nCaretPos <= 0)
				nCaretPos = SendMessage(LB_GETCOUNT) - 1;
			else
				nCaretPos--;
			SendMessage(LB_SETCURSEL, nCaretPos);
			OnSelChange();
		}
		return -2;

	case VK_DOWN:
		{
			if (nCaretPos >= SendMessage(LB_GETCOUNT) - 1)
				nCaretPos = 0;
			else
				nCaretPos++;
			SendMessage(LB_SETCURSEL, nCaretPos);
			OnSelChange();
		}
		return -2;
	}
	return -2;
}

//
// OnLButtonDblClk
//

void CBrowser::OnLButtonDblClk(UINT nFlags, POINT pt)
{
	try
	{
		LRESULT lResult = SendMessage(LB_ITEMFROMPOINT, 0, MAKELPARAM(pt.x, pt.y));
		if (LB_ERR == lResult)
			return;

		CProperty* pProperty = (CProperty*)SendMessage(LB_GETITEMDATA, lResult);
		if (NULL == pProperty || LB_ERR == (int)pProperty)
			return;

		if (pProperty != m_pCurrentProperty)
			return;

		CRect rcItem;
		lResult = SendMessage(LB_GETITEMRECT, m_nCurrentIndex, (LPARAM)&rcItem);
		rcItem.left = rcItem.left + rcItem.Width() / 2;

		switch (pProperty->GetEditType())
		{
		case CProperty::eBool:
			{
				VARIANT_BOOL vbTemp = m_pCurrentProperty->GetData().boolVal;
				if (VARIANT_TRUE == m_pCurrentProperty->GetData().boolVal)
					m_pCurrentProperty->GetData().boolVal = VARIANT_FALSE;
				else
					m_pCurrentProperty->GetData().boolVal = VARIANT_TRUE;

				if (SUCCEEDED(m_pCurrentProperty->SetValue()))
				{
					if (m_pDesigner)
						m_pDesigner->Update();
					m_pObjectChanged->ObjectChanged();
					InvalidateRect(&rcItem, FALSE);
				}
				else
				{
					m_pCurrentProperty->GetData().boolVal = vbTemp;
					MessageBox(IDS_FAILEDTOSETPROPERTY);
				}
			}
			break;

		case CProperty::eEnum:
			{
				CEnum* pEnum = m_pCurrentProperty->GetEnum();
				if (pEnum)
				{
					int nCurrentIndex;
					EnumItem* pEnumItem = pEnum->ByValue(m_pCurrentProperty->GetData().lVal, nCurrentIndex);
					if (pEnumItem)
					{
						int nLastIndex = pEnum->NumOfItems() - 1;
						int nNewIndex = nCurrentIndex + 1;
						if (nNewIndex > nLastIndex)
							nNewIndex = 0;

						EnumItem* pEnumItem = pEnum->ByIndex(nNewIndex);
						if (pEnumItem)
						{
							long nTemp = m_pCurrentProperty->GetData().lVal;
							m_pCurrentProperty->GetData().lVal = pEnumItem->m_nValue;
							if (SUCCEEDED(m_pCurrentProperty->SetValue()))
							{
								if (m_pDesigner)
									m_pDesigner->Update();
								m_pObjectChanged->ObjectChanged();
								InvalidateRect(&rcItem, FALSE);
							}
							else
							{
								m_pCurrentProperty->GetData().lVal = nTemp;
								MessageBox(IDS_FAILEDTOSETPROPERTY);
							}
						}
					}
				}
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
// SetTypeInfo
//

void CBrowser::SetTypeInfo(IDispatch* pObject, IObjectChanged* pObjectChanged)
{
	try
	{
		if (NULL == pObject)
			return;

		UpdatePrevSelection(-1);
		m_nCurrentIndex = -1;
		m_pCurrentProperty = NULL;

		if (m_pPerPropertyBrowsing)
		{
			m_pPerPropertyBrowsing->Release();
			m_pPerPropertyBrowsing = NULL;
		}
		if (m_pObject)
		{
			m_pObject->Release();
			m_pObject = NULL;
		}
		m_pObject = pObject;
		m_pObject->AddRef();

		m_pObjectChanged = pObjectChanged;

		if (m_pTypeInfo)
			m_pTypeInfo->Release();

		HRESULT hResult = m_pObject->GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT, &m_pTypeInfo);
		if (FAILED(hResult))
			return;

		hResult = m_pObject->QueryInterface(IID_IDDPerPropertyBrowsing, (LPVOID*)&m_pPerPropertyBrowsing);
		if (FAILED(hResult))
			m_pPerPropertyBrowsing = NULL;

		CProperty::EditType etProperty;
		TYPEATTR* pTypeAttr;
		FUNCDESC* pFuncDesc;
		TYPEDESC* pTypeDesc;
		ELEMDESC* pElem;
		USHORT    nVariantType;
		CEnum*    pEnum = NULL;
		UINT      nNumOfNames;
		BSTR      bstrFunctionName;
		BOOL      bValid;
		BOOL      bPtr;

		hResult = m_pTypeInfo->GetTypeAttr(&pTypeAttr);
		if (FAILED(hResult))
		{
			if (m_pPerPropertyBrowsing)
			{
				m_pPerPropertyBrowsing->Release();
				m_pPerPropertyBrowsing = NULL;
			}
			return;
		}

		CProperty* pProperty;
		int nResult;
		try 
		{
			for (int nFunction = 0; nFunction < pTypeAttr->cFuncs; nFunction++)
			{
				pProperty = NULL;
				bValid = FALSE;
				pEnum = NULL;
				bPtr = FALSE;

				hResult = m_pTypeInfo->GetFuncDesc(nFunction, &pFuncDesc);
				if (FAILED(hResult))
					continue;

				if (INVOKE_FUNC == pFuncDesc->invkind)
				{
					// If it is a method just skip it, we want just properties
					m_pTypeInfo->ReleaseFuncDesc(pFuncDesc);
					continue;
				}

				hResult = m_pTypeInfo->GetNames(pFuncDesc->memid, 
												&bstrFunctionName, 
												1, 
												&nNumOfNames);
				if (FAILED(hResult))
				{
					m_pTypeInfo->ReleaseFuncDesc(pFuncDesc);
					continue;
				}

				MAKE_TCHARPTR_FROMWIDE(szFunctionName, bstrFunctionName);
				if ('_' == szFunctionName[0])
				{
					//
					// Skip Hidden Properties
					//
					
					SysFreeString(bstrFunctionName);
					m_pTypeInfo->ReleaseFuncDesc(pFuncDesc);
					continue;
				}

				switch (pFuncDesc->invkind)
				{
				case INVOKE_PROPERTYGET:
					if (1 == pFuncDesc->cParams)
						pElem = pFuncDesc->lprgelemdescParam;
					else
						pElem = &pFuncDesc->elemdescFunc;
					break;

				case INVOKE_PROPERTYPUTREF:
					pElem = &pFuncDesc->elemdescFunc;
					break;
				
				default:
					pElem = pFuncDesc->lprgelemdescParam;
					break;
				}

				nVariantType = pElem->tdesc.vt;
				pTypeDesc = &(pElem->tdesc);
				if (NULL == m_pPerPropertyBrowsing || 
					!ResolvePerPropertyBrowsing(pFuncDesc->memid, 
											    bstrFunctionName, 
											    pEnum, 
											    nVariantType, 
											    etProperty, 
										        bValid))
				{
					switch (pElem->tdesc.vt)
					{
					case VT_PTR:
						{
							nVariantType = pElem->tdesc.lptdesc->vt;
							pTypeDesc = pElem->tdesc.lptdesc;
							bValid = ResolvePointerType(pTypeAttr->guid, 
														m_pTypeInfo, 
														pTypeDesc->hreftype, 
														nVariantType, 
														etProperty);
						}
						break;
					
					case VT_USERDEFINED:
						{
							bValid = ResolveUserDefinedType(pTypeAttr->guid, 
															m_pTypeInfo, 
															pTypeDesc->hreftype, 
															nVariantType, 
															etProperty, 
															pEnum);
						}
						break;

					case VT_BOOL:
						etProperty = CProperty::eBool;
						break;

					case VT_I2:
					case VT_I2|VT_BYREF:
					case VT_I4:
					case VT_I4|VT_BYREF:
					case VT_INT:
					case VT_INT|VT_BYREF:
					case VT_UINT:
					case VT_UINT|VT_BYREF:
						etProperty = CProperty::eNumber;
						break;

					case VT_BSTR:
						etProperty = CProperty::eString;
						break;

					case VT_DISPATCH|VT_ARRAY:
						if (0 == _tcsicmp(szFunctionName, _T("ShortCuts")))
						{
							nVariantType = VT_DISPATCH|VT_ARRAY;
							etProperty = CProperty::eShortCut;
						}
						break;

					case VT_VARIANT:
						if (0 == _tcsicmp(szFunctionName, _T("ShortCuts")))
						{
							nVariantType = VT_DISPATCH|VT_ARRAY;
							etProperty = CProperty::eShortCut;
						}
						else
							nVariantType = VT_VARIANT;
						break;
					}
				}
				if (!bValid && (VT_PTR == pElem->tdesc.vt || VT_USERDEFINED == pElem->tdesc.vt))
				{
					m_pTypeInfo->ReleaseFuncDesc(pFuncDesc);
					SysFreeString(bstrFunctionName);
					continue;
				}

				nResult = SendMessage(LB_FINDSTRING, -1, (LPARAM)szFunctionName);
				if (LB_ERR != nResult)
					pProperty = (CProperty*)SendMessage(LB_GETITEMDATA, nResult);
				else
				{
					BSTR bstrName = NULL;
					BSTR bstrDocString = NULL;
					DWORD nHelpContext = NULL;
					BSTR bstrHelpFileName;
					hResult = m_pTypeInfo->GetDocumentation(pFuncDesc->memid, 
															&bstrName, 
															&bstrDocString, 
															&nHelpContext, 
															&bstrHelpFileName);
					SysFreeString(bstrName);
					SysFreeString(bstrHelpFileName);

					pProperty = new CProperty(m_pObject, 
											  pFuncDesc->memid, 
											  nVariantType,
											  etProperty,
											  pFuncDesc->cParams,
											  bstrDocString,
											  nHelpContext,
											  pEnum);
					if (NULL == pProperty)
					{
						m_pTypeInfo->ReleaseFuncDesc(pFuncDesc);
						continue;
					}

					if (0 == wcscmp(bstrFunctionName, L"Bitmap"))
					{
						//
						// TODO
						//
						// This is a cheese way of getting to the icon for tools in ActiveBar
						// Should be fixed someday
						//

						pProperty->SetEditType(CProperty::eBitmap);
						pProperty->Name(szFunctionName);
						nResult = SendMessage(LB_ADDSTRING, 0, (LPARAM)pProperty);
					}
					else if (SUCCEEDED(pProperty->GetValue()))
					{
						if (VT_VARIANT == nVariantType)
						{
							switch (pProperty->GetData().vt)
							{
							case VT_BOOL:
								pProperty->SetEditType(CProperty::eBool);
								pProperty->VarType(pProperty->VarType());
								break;

							case VT_I2:
							case VT_I2|VT_BYREF:
							case VT_I4:
							case VT_I4|VT_BYREF:
							case VT_INT:
							case VT_INT|VT_BYREF:
							case VT_UINT:
							case VT_UINT|VT_BYREF:
							case VT_BSTR:
								pProperty->SetEditType(CProperty::eVariant);
								pProperty->VarType(pProperty->GetData().vt);
								nVariantType = pProperty->VarType();
								break;

							case VT_EMPTY:
								pProperty->SetEditType(CProperty::eVariant);
								pProperty->VarType(pProperty->VarType());
								nVariantType = pProperty->VarType();
								break;

							default:
								delete pProperty;
								pProperty = NULL;
								m_pTypeInfo->ReleaseFuncDesc(pFuncDesc);
								continue;
							}
						}
						pProperty->Name(szFunctionName);
						nResult = SendMessage(LB_ADDSTRING, 0, (LPARAM)pProperty);
					}
					else
					{
						delete pProperty;
						pProperty = NULL;
						m_pTypeInfo->ReleaseFuncDesc(pFuncDesc);
						continue;
					}
				}
				if (pProperty)
				{
					switch (pFuncDesc->invkind)
					{
					case INVOKE_PROPERTYGET:
						pProperty->Get(TRUE);
						pProperty->VarType(nVariantType);
						break;

					case INVOKE_PROPERTYPUT:
						pProperty->Put(TRUE);
						break;

					case INVOKE_PROPERTYPUTREF:
						pProperty->PutRef(TRUE);
						break;
					}
				}
				SysFreeString(bstrFunctionName);
				m_pTypeInfo->ReleaseFuncDesc(pFuncDesc);
			}
		}
		CATCH
		{
			assert(FALSE);
			REPORTEXCEPTION(__FILE__, __LINE__)
		}
		m_pTypeInfo->ReleaseTypeAttr(pTypeAttr);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}

	//
	// Set the current property
	//

	int nIndex = SendMessage(LB_GETCOUNT);
	if (LB_ERR == nIndex)
		return;

	nIndex = SendMessage(LB_SETCURSEL, 0);
	if (LB_ERR == nIndex)
		return;

	m_nCurrentIndex = nIndex;

	nIndex = SendMessage(LB_GETITEMDATA, nIndex);
	if (LB_ERR == nIndex)
		return;

	m_pCurrentProperty = reinterpret_cast<CProperty*>(nIndex);

	if (m_pCurrentProperty)
	{
		MAKE_TCHARPTR_FROMWIDE(szDocString, m_pCurrentProperty->DocString());
		::SendMessage(GetDlgItem(GetParent(m_hWnd), IDC_ST_DESC), WM_SETTEXT, 0, (LPARAM)szDocString);
	}
	
	m_nPrevIndex = -1;
	m_pPrevProperty = NULL;
}

//
// ResolvePerPropertyBrowsing 
//

BOOL CBrowser::ResolvePerPropertyBrowsing(UINT				   nDispId, 
										  BSTR				   bstrFunctionName, 
										  CEnum*&			   pEnum, 
										  USHORT&			   nVariantType, 
										  CProperty::EditType& etProperty, 
										  BOOL&				   bValid)
{
	BOOL bResult = TRUE;
	try
	{
		CALPOLESTR caStringsOut;
		caStringsOut.pElems = NULL;
		CADWORD caCookiesOut;
		caCookiesOut.pElems = NULL;
		HRESULT hResult = m_pPerPropertyBrowsing->GetPredefinedStrings(nDispId,
																	   &caStringsOut, 
																	   &caCookiesOut);
		if (FAILED(hResult))
			return FALSE;

		long nType;
		m_pPerPropertyBrowsing->GetType(nDispId, &nType);
		switch (nType)
		{
		case 0:
			nVariantType = VT_I4;
			etProperty = CProperty::eEnum;
			break;

		case 1:
			nVariantType = VT_BSTR;
			etProperty = CProperty::eEnum;
			break;

		case 2:
			nVariantType = VT_I4;
			etProperty = CProperty::eFlags;
			break;
		}

		pEnum = FindEnum(bstrFunctionName);
		if (pEnum)
		{
			bValid = TRUE;
			goto Cleanup;
		}

		pEnum = new CEnum(bstrFunctionName);
		if (NULL == pEnum)
		{
			bResult = TRUE;
			goto Cleanup;
		}

		m_aEnums.Add(pEnum);
		bValid = TRUE;

		EnumItem* pEnumItem;
		if (caStringsOut.cElems > 0)
		{
			BOOL bFirstString = TRUE;
			for (UINT nItem = 0; nItem < caStringsOut.cElems; nItem++)
			{
				pEnumItem = new EnumItem;
				if (NULL == pEnumItem)
					continue;

				// Build Display String
				MAKE_TCHARPTR_FROMWIDE(szName, caStringsOut.pElems[nItem]);
				switch (nType)
				{
				case 0:
					pEnumItem->m_strDisplay.Format(_T("%s"), szName);
					pEnumItem->m_nValue = caCookiesOut.pElems[nItem];
					break;

				case 1:
					if (bFirstString)
					{
						bFirstString = FALSE;
						pEnumItem->m_strDisplay.LoadString(IDS_NONE);
						pEnum->Add(pEnumItem);
						pEnumItem = new EnumItem;
						if (NULL == pEnumItem)
							continue;
					}
					pEnumItem->m_strDisplay.Format(_T("%s"), szName);
					break;

				case 2:
					pEnumItem->m_strDisplay.Format(_T("%s"), szName);
					pEnumItem->m_nValue = caCookiesOut.pElems[nItem];
					break;
				}
				pEnum->Add(pEnumItem);
			}
		}
		else
		{
			switch (nType)
			{
			case 1:
				pEnumItem = new EnumItem;
				if (pEnumItem)
				{
					pEnumItem->m_strDisplay.LoadString(IDS_NONE);
					pEnum->Add(pEnumItem);
				}
				break;
			}
		}
Cleanup:
		if (caStringsOut.pElems)
		{
			for (UINT nItem = 0; nItem < caStringsOut.cElems; nItem++)
				SysFreeString(caStringsOut.pElems[nItem]);
			CoTaskMemFree(caStringsOut.pElems);
		}
		if (caCookiesOut.pElems)
			CoTaskMemFree(caCookiesOut.pElems);
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
// ResolvePointerType
//

BOOL CBrowser::ResolvePointerType(REFCLSID			   rcClassId,
								  ITypeInfo*		   pTypeInfo,
								  HREFTYPE			   hRefType,
								  VARTYPE&			   vType,
								  CProperty::EditType& etProperty)
{
	ITypeInfo* pRefTypeInfo = NULL;
	HRESULT    hResult;
	DWORD	   nHelpContext;
	BSTR	   bstrHelpFileName = NULL;
	BSTR	   bstrDocString = NULL;
	BSTR	   bstrName = NULL;
	BOOL	   bValid = FALSE;

	try	
	{
		hResult = pTypeInfo->GetRefTypeInfo(hRefType, &pRefTypeInfo);
		if (S_OK != hResult && NULL == pRefTypeInfo)
			return FALSE;

		hResult = pRefTypeInfo->GetDocumentation(MEMBERID_NIL, 
												 &bstrName, 
												 &bstrDocString, 
												 &nHelpContext, 
												 &bstrHelpFileName);
		if (S_OK != hResult)
			goto Cleanup;

		if (NULL == bstrName)
			goto Cleanup;

		if (0 == _wcsicmp(bstrName, L"Picture") || 0 == _wcsicmp(bstrName, L"IPictureDisp"))
		{
			etProperty = CProperty::ePicture;
			vType = VT_DISPATCH;
			bValid = TRUE;
		}
		else if (0 == _wcsicmp(bstrName, L"Font") || 0 == _wcsicmp(bstrName, L"IFontDisp"))
		{
			etProperty = CProperty::eFont;
			vType = VT_DISPATCH;
			bValid = TRUE;
		}
		else if (0 == _wcsicmp(bstrName, L"ComboList"))
		{
			etProperty = CProperty::eComboList;
			vType = VT_DISPATCH;
			bValid = TRUE;
		}
Cleanup:
		SysFreeString(bstrName);
		SysFreeString(bstrDocString);
		SysFreeString(bstrHelpFileName);
		pRefTypeInfo->Release(); 
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return bValid;
}

//
// ResolveUserDefinedType
//

BOOL CBrowser::ResolveUserDefinedType(REFCLSID			   rcClassId,
									  ITypeInfo*		   pTypeInfo,
									  HREFTYPE			   hRefType,
									  VARTYPE&			   vType,
									  CProperty::EditType& etProperty,
									  CEnum*&			   pEnum)
{
	BOOL bValid = FALSE;
	try	
	{
		ITypeInfo* pRefTypeInfo;
		HRESULT hResult = pTypeInfo->GetRefTypeInfo(hRefType, &pRefTypeInfo);
		if (FAILED(hResult) && NULL == pRefTypeInfo)
			return FALSE;

		TYPEATTR* pRefTypeAttr = NULL;
		DWORD     nHelpContext;
		BSTR      bstrHelpFileName = NULL;
		BSTR      bstrDocString = NULL;
		BSTR      bstrName = NULL;
		hResult = pRefTypeInfo->GetDocumentation(MEMBERID_NIL, 
											     &bstrName, 
												 &bstrDocString, 
												 &nHelpContext, 
												 &bstrHelpFileName);
		if (FAILED(hResult))
			goto Cleanup;

		if (NULL == bstrName)
			goto Cleanup;

		//
		// Need to check for special OLE Types
		//

		if (0 == _wcsicmp(bstrName, L"OLE_YSIZE_PIXELS"))
		{
			etProperty = CProperty::eNumber;
			vType = VT_I4;
			bValid = TRUE;
			goto Cleanup;
		}
		if (0 == _wcsicmp(bstrName, L"OLE_XSIZE_PIXELS"))
		{
			etProperty = CProperty::eNumber;
			vType = VT_I4;
			bValid = TRUE;
			goto Cleanup;
		}
		if (0 == _wcsicmp(bstrName, L"OLE_COLOR"))
		{
			etProperty = CProperty::eColor;
			vType = VT_I4;
			bValid = TRUE;
			goto Cleanup;
		}

		hResult = pRefTypeInfo->GetTypeAttr(&pRefTypeAttr);
		if (FAILED(hResult) || TKIND_ENUM != pRefTypeAttr->typekind)
			goto Cleanup;

		//
		// It must be an Enum
		//

		pEnum = FindEnum(bstrName);
		if (NULL == pEnum)
		{
			//
			// If not found build the enum list
			//

			hResult = BuildEnumList(rcClassId, bstrName, pRefTypeInfo, pRefTypeAttr, pEnum);
			if (FAILED(hResult))
				goto Cleanup;
		}

		vType = VT_I4;
		etProperty = CProperty::eEnum;
		bValid = TRUE;

Cleanup:
		if (pRefTypeAttr)
			pRefTypeInfo->ReleaseTypeAttr(pRefTypeAttr); 

		pRefTypeInfo->Release(); 
		SysFreeString(bstrName);
		SysFreeString(bstrDocString);
		SysFreeString(bstrHelpFileName);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
	return bValid;
}

//
// Init
//

void CBrowser::Init()
{
/*	LV_COLUMN lvColumn;
	lvColumn.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;
	lvColumn.fmt = LVCFMT_LEFT;
	lvColumn.cx = 75;
	lvColumn.pszText = _T("Property");
	ListView_InsertColumn(m_hWnd, 0, &lvColumn);
	lvColumn.pszText = _T("Value");
	ListView_InsertColumn(m_hWnd, 1, &lvColumn);
	*/
}

//
// BuildEnumList
//
									  
HRESULT CBrowser::BuildEnumList(CLSID	   rcClassId, 
								BSTR	   bstrEnumName, 
								ITypeInfo* pRefTypeInfo, 
								TYPEATTR*  pRefTypeAttr, 
								CEnum*&    pEnum)
{
	HRESULT hResult;
	try
	{
		pEnum = new CEnum(bstrEnumName);
		if (NULL == pEnum)
			return E_OUTOFMEMORY;

		m_aEnums.Add(pEnum);

		EnumItem* pEnumItem;
		VARDESC*  pVarDesc;
		VARIANT   vValue;
		ULONG	  nHelpContext;
		BSTR	  bstrValueName;
		BSTR	  bstrDocName;
		BSTR	  bstrHelpFile;

		for (int nIndex = 0; nIndex < pRefTypeAttr->cVars; nIndex++)
		{
			hResult = pRefTypeInfo->GetVarDesc(nIndex, &pVarDesc); 
			if (FAILED(hResult))
				continue;

			hResult = pRefTypeInfo->GetDocumentation(pVarDesc->memid, 
													 &bstrValueName, 
													 &bstrDocName,
													 &nHelpContext,
													 &bstrHelpFile);
			if (FAILED(hResult))
			{
				pRefTypeInfo->ReleaseVarDesc(pVarDesc);
				continue;
			}

			try
			{
				pEnumItem = new EnumItem;
				if (NULL == pEnumItem)
					throw E_OUTOFMEMORY;

				vValue.vt = VT_EMPTY;
				hResult = VariantChangeType(&vValue, pVarDesc->lpvarValue, 0, VT_I4);
				if (SUCCEEDED(hResult))
					pEnumItem->m_nValue = vValue.lVal;
				else 
					pEnumItem->m_nValue = 0;
			}
			CATCH
			{
				assert(FALSE);
				REPORTEXCEPTION(__FILE__, __LINE__)
				SysFreeString(bstrValueName);
				SysFreeString(bstrDocName);
				SysFreeString(bstrHelpFile);
				pRefTypeInfo->ReleaseVarDesc(pVarDesc);
				delete pEnumItem;
				continue;
			}

			// Build Display String
			MAKE_TCHARPTR_FROMWIDE(szName, bstrValueName);
			pEnumItem->m_strDisplay.Format(_T("%d - %s"), vValue.lVal, szName);
			pEnum->Add(pEnumItem);

			SysFreeString(bstrValueName);
			SysFreeString(bstrDocName);
			SysFreeString(bstrHelpFile);
			pRefTypeInfo->ReleaseVarDesc(pVarDesc);
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
	return hResult;
}

