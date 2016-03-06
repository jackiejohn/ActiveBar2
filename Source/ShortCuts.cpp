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
#include "support.h"
#include "resource.h"
#include "debug.h"
#include "Errors.h"
#include "Bar.h"
#include "Band.h"
#include "Tool.h"
#include "Custom.h"
#include "PopupWin.h"
#include "Localizer.h"
#include "ShortCuts.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

//{OLE VERBS}
//{OLE VERBS}
//{OBJECT CREATEFN}
IUnknown *CreateFN_ShortCutStore(IUnknown *pUnkOuter)
{
	return (IUnknown *)(IShortCut *)new ShortCutStore();
}

DEFINE_AUTOMATIONOBJECT(&CLSID_ShortCut,ShortCut,_T("ShortCut Class"),CreateFN_ShortCutStore,1,0,&IID_IShortCut,_T(""));
void *ShortCutStore::objectDef=&ShortCutObject;
ShortCutStore *ShortCutStore::CreateInstance(IUnknown *pUnkOuter)
{
	return new ShortCutStore();
}
//{OBJECT CREATEFN}
ShortCutStore::ShortCutStore()
	: m_refCount(1)
	  ,m_pBar(NULL)
{
// {BEGIN INIT}
// {END INIT}
}

ShortCutStore::~ShortCutStore()
{
// {BEGIN CLEANUP}
// {END CLEANUP}
}

ShortCutStore::ShortCutPropV1::ShortCutPropV1()
{
	memset(m_nKeyCodes, 0, sizeof(m_nKeyCodes));
	memset(m_bShift, 0, sizeof(m_bShift));
	memset(m_nKeyboardCodes, 0, sizeof(m_nKeyboardCodes));
	m_nControlKeys = 0;
}

//{DEF IUNKNOWN MEMBERS}
STDMETHODIMP ShortCutStore::QueryInterface(REFIID riid, void **ppvObjOut)
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
	if (DO_GUIDS_MATCH(riid,IID_IShortCut))
	{
		*ppvObjOut=(void *)(IShortCut *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	if (DO_GUIDS_MATCH(riid,IID_ISupportErrorInfo))
	{
		*ppvObjOut=(void *)(ISupportErrorInfo *)this;
		((IUnknown *)(*ppvObjOut))->AddRef();
		return NOERROR;
	}
	//{CUSTOM QI}
	//{ENDCUSTOM QI}
	*ppvObjOut = NULL;
	return E_NOINTERFACE;
}
STDMETHODIMP_(ULONG) ShortCutStore::AddRef()
{
	return ++m_refCount;
}
STDMETHODIMP_(ULONG) ShortCutStore::Release()
{
	if (0!=--m_refCount)
		return m_refCount;
	delete this;
	return 0;
}
//{ENDDEF IUNKNOWN MEMBERS}
STDMETHODIMP ShortCutStore::InterfaceSupportsErrorInfo( REFIID riid)
{
    if (riid == (REFIID)INTERFACEOFOBJECT(11))
        return S_OK;

    return S_FALSE;
}
STDMETHODIMP ShortCutStore::IsEqual( ShortCut*aShortCut)
{
	if (*this == (*((ShortCutStore*)(aShortCut))))
		return NOERROR;
	return E_FAIL;
}
STDMETHODIMP ShortCutStore::GetTypeInfoCount( UINT *pctinfo)
{
	*pctinfo = 1; 
	return NOERROR; 
}
STDMETHODIMP ShortCutStore::GetTypeInfo( UINT itinfo, LCID lcid, ITypeInfo **pptinfo)
{
	*pptinfo=GetObjectTypeInfo(lcid,objectDef);
	if ((*pptinfo)==NULL)
		return E_FAIL;
	return NOERROR;
}
STDMETHODIMP ShortCutStore::GetIDsOfNames( REFIID riid, OLECHAR **rgszNames, UINT cNames, ULONG lcid, long *rgdispid)
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
STDMETHODIMP ShortCutStore::Invoke( long dispidMember, REFIID riid, ULONG lcid, USHORT wFlags, DISPPARAMS *pdispparams, VARIANT *pvarResult, EXCEPINFO *pexcepinfo, UINT *puArgErr)
{
#ifdef _DEBUG
	if (dispidMember<0)
		TRACE1(2, "IDispatch::Invoke -%X\n",-dispidMember)
	else
		TRACE1(2, "IDispatch::Invoke %X\n",dispidMember)
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
STDMETHODIMP ShortCutStore::GetKeyCode( short nIndex,  long *retval)
{
	*retval = scpV1.m_nKeyCodes[nIndex];
	return NOERROR;
}
STDMETHODIMP ShortCutStore::Clear()
{
	memset(scpV1.m_nKeyCodes, 0, sizeof(scpV1.m_nKeyCodes));
	memset(scpV1.m_bShift, 0, sizeof(scpV1.m_bShift));
	memset(scpV1.m_nKeyboardCodes, 0, sizeof(scpV1.m_nKeyboardCodes));
	scpV1.m_nControlKeys = 0;
	return NOERROR;
}
STDMETHODIMP ShortCutStore::get_Value(BSTR *retval)
{
	TCHAR szShortCut[120];
	GetShortCutDesc (szShortCut);
	MAKE_WIDEPTR_FROMTCHAR(wShortCut, szShortCut);
	*retval = SysAllocString(wShortCut);
	if (NULL == *retval)
		return E_FAIL;
	return NOERROR;
}
STDMETHODIMP ShortCutStore::put_Value(BSTR val)
{
	try
	{
		MAKE_TCHARPTR_FROMWIDE(szShortCut, val);
		TypedArray<Token*> aTokens;
		if (Parse(szShortCut, aTokens))
		{
			if (!Syntax(aTokens, scpV1))
			{
				int nCount = aTokens.GetSize();
				for (int nIndex = 0; nIndex < nCount; nIndex++)
					delete aTokens.GetAt(nIndex);
				SendError(0, IDS_ERR_INVALIDSHORTCUT);
				return CUSTOM_CTL_SCODE(IDERR_INVALIDSHORTCUT);
			}
		}
		else
		{
			int nCount = aTokens.GetSize();
			for (int nIndex = 0; nIndex < nCount; nIndex++)
				delete aTokens.GetAt(nIndex);
			SendError(0, IDS_ERR_INVALIDSHORTCUT);
			return CUSTOM_CTL_SCODE(IDERR_INVALIDSHORTCUT);
		}
		int nCount = aTokens.GetSize();
		for (int nIndex = 0; nIndex < nCount; nIndex++)
			delete aTokens.GetAt(nIndex);
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		SendError(0, IDS_ERR_INVALIDSHORTCUT);
		return E_FAIL;
	}
	return NOERROR;
}
STDMETHODIMP ShortCutStore::Set( short nIndex,  long nKeyCode,  long nKeyboardCode,  VARIANT_BOOL bShift)
{
	scpV1.m_nKeyCodes[nIndex] = (TCHAR)nKeyCode;
	scpV1.m_nKeyboardCodes[nIndex] = nKeyboardCode;
	scpV1.m_bShift[nIndex] = (bShift == VARIANT_TRUE ? TRUE : FALSE);
	return NOERROR;
}
STDMETHODIMP ShortCutStore::SetControlCode( long nControlCode)
{
	scpV1.m_nControlKeys = nControlCode;
	return NOERROR;
}
STDMETHODIMP ShortCutStore::Clone( IShortCut **pShortCutStore)
{
	ShortCutStore* pShortCutStore2 = ShortCutStore::CreateInstance(NULL);
	if (NULL == pShortCutStore2)
		return E_OUTOFMEMORY;
	memcpy(pShortCutStore2, this, sizeof(ShortCutStore));
	*pShortCutStore = (IShortCut*)pShortCutStore2;
	return NOERROR;
}
//{EVENT WRAPPERS}
//{EVENT WRAPPERS}
//
// SendError
//

void ShortCutStore::SendError(DWORD dwHelpContextID, DWORD dwDescriptionId)
{
	//
    // First get the Create Error Info Object.
    //

	ICreateErrorInfo* pCreateErrorInfo;
	HRESULT hResult = CreateErrorInfo(&pCreateErrorInfo);
    if (FAILED(hResult)) 
		return;

	//
	// Set up some default information on it.
    //
	
	pCreateErrorInfo->SetGUID((REFIID)INTERFACEOFOBJECT(11));
    
	pCreateErrorInfo->SetHelpContext(dwHelpContextID);

	MAKE_WIDEPTR_FROMTCHAR(wDesc, LoadStringRes(dwDescriptionId));
	BSTR bstrDescription = SysAllocString(wDesc);
    pCreateErrorInfo->SetDescription(bstrDescription);
	SysFreeString(bstrDescription);

#ifdef _UNICODE
	pCreateErrorInfo->SetHelpFile((LPWSTR)HELPFILEOFOBJECT(11));
	pCreateErrorInfo->SetSource((LPWSTR)NAMEOFOBJECT(11));
#else
	pCreateErrorInfo->SetHelpFile(L"Activebar20.hlp");
	WCHAR wszTmp[256];
    MultiByteToWideChar(CP_ACP, 0, NAMEOFOBJECT(11), -1, wszTmp, 256);
    pCreateErrorInfo->SetSource(wszTmp);
#endif

    //
	// Now set the error info, now it is up to the system
    //
	
	IErrorInfo* pErrorInfo;
    hResult = pCreateErrorInfo->QueryInterface(IID_IErrorInfo, (void**)&pErrorInfo);
    CLEANUP_ON_FAILURE(hResult);

    SetErrorInfo(0, pErrorInfo);
    pErrorInfo->Release();

CleanUp:
    pCreateErrorInfo->Release();
}

//
// GetShortCutDesc
//

BOOL ShortCutStore::GetShortCutDesc (LPTSTR szShortCut)
{
	if (NULL == szShortCut)
		return FALSE;

	TCHAR szBuffer[120];
	szBuffer[0] = NULL;
	szShortCut[0] = NULL;
	if (FALT & scpV1.m_nControlKeys)
	{
		if (m_pBar)
		{
			LPCTSTR szLocaleString = m_pBar->Localizer()->GetString(ddLTAlt);
			if (szLocaleString)
				lstrcat(szShortCut, szLocaleString);
			else
				lstrcat(szShortCut, Globals::m_szAlt);
		}
		else
			lstrcat(szShortCut, Globals::m_szAlt);
	}

	if (FCONTROL & scpV1.m_nControlKeys)
	{
		if (m_pBar)
		{
			LPCTSTR szLocaleString = m_pBar->Localizer()->GetString(ddLTControl);
			if (szLocaleString)
				lstrcat(szShortCut, szLocaleString);
			else
				lstrcat(szShortCut, Globals::m_szCtrl);
		}
		else
			lstrcat(szShortCut, Globals::m_szCtrl);
	}
	if (scpV1.m_nKeyCodes[0])
	{
		if (scpV1.m_bShift[0])
		{
			if (m_pBar)
			{
				LPCTSTR szLocaleString = m_pBar->Localizer()->GetString(ddLTShift);
				if (szLocaleString)
					lstrcat(szShortCut, szLocaleString);
				else
					lstrcat(szShortCut, Globals::m_szShift);
			}
			else
				lstrcat(szShortCut, Globals::m_szShift);
		}
	
		long nScanCode = MapVirtualKey(scpV1.m_nKeyCodes[0], 0);
		nScanCode <<= 16;
		if (scpV1.m_nKeyboardCodes[0] & 0x01000000)
			nScanCode |= 0x01000000;
		int nResult = GetKeyNameText(nScanCode, szBuffer, 50); 
		if (nResult > 0)
		{
			if (scpV1.m_nKeyCodes[0] >= VK_NUMPAD0 && scpV1.m_nKeyCodes[0] <= VK_NUMPAD9)
			{
				TCHAR* szSubString = _tcsstr(szBuffer, _T("Num "));
				if (szSubString)
					lstrcat(szShortCut, szSubString + lstrlen(_T("Num ")));
				else
					lstrcat(szShortCut, szBuffer);
			}
			else
				lstrcat(szShortCut, szBuffer);
		}
		if (scpV1.m_nKeyCodes[1])
		{
			lstrcat(szShortCut, _T(", "));
			if (scpV1.m_bShift[1])
				lstrcat(szShortCut, Globals::m_szShift);
			nScanCode = MapVirtualKey(scpV1.m_nKeyCodes[1], 0);
			nScanCode <<= 16;
			if (scpV1.m_nKeyboardCodes[1] & 0x01000000)
				nScanCode |= 0x01000000;
			nResult = GetKeyNameText(nScanCode, szBuffer, 50); 
			if (nResult > 0)
			{
				if (scpV1.m_nKeyCodes[1] >= VK_NUMPAD0 && scpV1.m_nKeyCodes[1] <= VK_NUMPAD9)
				{
					TCHAR* szSubString = _tcsstr(szBuffer, _T("Num "));
					if (szSubString)
						lstrcat(szShortCut, szSubString + lstrlen(_T("Num ")));
					else
						lstrcat(szShortCut, szBuffer);
				}
				else
					lstrcat(szShortCut, szBuffer);
			}
		}
	}
	return TRUE;
}

struct SpecialKeys
{
	TCHAR* szSpecialKey;
	int    cKey;
};

static SpecialKeys theControlKeys[] = 
{
	{{"Alt+"},{Token::eAlt}},
	{{"Control+"},{Token::eControl}},
	{{"Shift+"},{Token::eShift}}
};

static SpecialKeys theSpecialKeys[] = 
{	
	{{"Insert"},{VK_INSERT}},
	{{"Delete"},{VK_DELETE}},
	{{"Backspace"},{VK_BACK}},
	{{"Home"},{VK_HOME}},
	{{"End"},{VK_END}},
	{{"Page Up"},{VK_PRIOR}},
	{{"Page Down"},{VK_NEXT}},
	{{"Up"},{VK_UP}},
	{{"Down"},{VK_DOWN}},
	{{"Left"},{VK_LEFT}},
	{{"Right"},{VK_RIGHT}},
	{{"Num 0"},{VK_NUMPAD0}},
	{{"Num 1"},{VK_NUMPAD1}},
	{{"Num 2"},{VK_NUMPAD2}},
	{{"Num 3"},{VK_NUMPAD3}},
	{{"Num 4"},{VK_NUMPAD4}},
	{{"Num 5"},{VK_NUMPAD5}},
	{{"Num 6"},{VK_NUMPAD6}},
	{{"Num 7"},{VK_NUMPAD7}},
	{{"Num 8"},{VK_NUMPAD8}},
	{{"Num 9"},{VK_NUMPAD9}},
	{{"Num *"},{VK_MULTIPLY}},
	{{"Num /"},{VK_DIVIDE}},
	{{"Num -"},{VK_SUBTRACT}},
	{{"Num +"},{VK_ADD}},
	{{"F24"},{VK_F24}},
	{{"F23"},{VK_F23}},
	{{"F22"},{VK_F22}},
	{{"F21"},{VK_F21}},
	{{"F20"},{VK_F20}},
	{{"F19"},{VK_F19}},
	{{"F18"},{VK_F18}},
	{{"F17"},{VK_F17}},
	{{"F16"},{VK_F16}},
	{{"F15"},{VK_F15}},
	{{"F14"},{VK_F14}},
	{{"F13"},{VK_F13}},
	{{"F12"},{VK_F12}},
	{{"F11"},{VK_F11}},
	{{"F10"},{VK_F10}},
	{{"F1"},{VK_F1}},
	{{"F2"},{VK_F2}},
	{{"F3"},{VK_F3}},
	{{"F4"},{VK_F4}},
	{{"F5"},{VK_F5}},
	{{"F6"},{VK_F6}},
	{{"F7"},{VK_F7}},
	{{"F8"},{VK_F8}},
	{{"F9"},{VK_F9}}
};

static TCHAR* szComma = {", "};

//
// Parse
//

BOOL ShortCutStore::Parse(TCHAR* szShortCut, TypedArray<Token*>& aTokens)
{
	TCHAR* szCurrentPos = szShortCut;
	Token* pToken;
	int nCount = 3;
	int nIndex;
	while (*szCurrentPos)
	{
		nCount = sizeof(theControlKeys)/sizeof(SpecialKeys);
		for (nIndex = 0; nIndex < nCount && *szCurrentPos; nIndex++)
		{
			if (0 == _tcsncmp(szCurrentPos, theControlKeys[nIndex].szSpecialKey, lstrlen(theControlKeys[nIndex].szSpecialKey)))
			{
				pToken = new Token((Token::TokenTypes)theControlKeys[nIndex].cKey);
				if (pToken)
					aTokens.Add(pToken);
				szCurrentPos += lstrlen(theControlKeys[nIndex].szSpecialKey);
			}
		}
		if (NULL == *szCurrentPos)
			continue;
		nCount = sizeof(theSpecialKeys)/sizeof(SpecialKeys);
		for (nIndex = 0; nIndex < nCount && *szCurrentPos; nIndex++)
		{
			if (0 == _tcsncmp(szCurrentPos, theSpecialKeys[nIndex].szSpecialKey, lstrlen(theSpecialKeys[nIndex].szSpecialKey)))
			{
				pToken = new Token(Token::eChar, theSpecialKeys[nIndex].cKey);
				if (pToken)
					aTokens.Add(pToken);
				szCurrentPos += lstrlen(theSpecialKeys[nIndex].szSpecialKey);
			}
		}
		if (NULL == *szCurrentPos)
			continue;
		if (0 == _tcsncmp(szCurrentPos, szComma, 2))
		{
			pToken = new Token(Token::eComma);
			if (pToken)
				aTokens.Add(pToken);
			szCurrentPos += lstrlen(szComma);
			continue;
		}
		if (NULL == *szCurrentPos)
			continue;

		pToken = new Token(Token::eChar, *szCurrentPos);
		if (pToken)
			aTokens.Add(pToken);
		szCurrentPos++;
	}
	return TRUE;
}

//
// Syntax
//

BOOL ShortCutStore::Syntax(const TypedArray<Token*>& aTokens, ShortCutPropV1& scpShortCut)
{
	BOOL bResult = TRUE;
	memset(scpShortCut.m_nKeyCodes, 0, sizeof(scpShortCut.m_nKeyCodes));
	memset(scpShortCut.m_bShift, 0, sizeof(scpShortCut.m_bShift));
	memset(scpShortCut.m_nKeyboardCodes, 0, sizeof(scpShortCut.m_nKeyboardCodes));
	scpShortCut.m_nControlKeys = 0;
	ParseStates eState = eStart;
	Token* pToken;
	int nCount = aTokens.GetSize();
	int nIndex = 0;
	while (eEnd != eState)
	{
		if (nIndex < nCount)
			pToken = aTokens.GetAt(nIndex);
		switch (eState)
		{
		case eStart:
			switch (pToken->m_eType)
			{
			case Token::eAlt:
				eState = eAlt;
				break;

			case Token::eControl:
				eState = eControl;
				break;

			case Token::eShift:
				eState = eShift;
				break;

			case Token::eChar:
				scpShortCut.m_nKeyCodes[0] = pToken->m_cCharacter;
				eState = eChar;
				switch (scpShortCut.m_nKeyCodes[0])
				{
				case VK_INSERT:
				case VK_DELETE:
//				case VK_BACK:
				case VK_HOME:
				case VK_END:
				case VK_PRIOR:
				case VK_NEXT:
				case VK_UP:
				case VK_DOWN:
				case VK_LEFT:
				case VK_RIGHT:
					scpShortCut.m_nKeyboardCodes[0] = 0x01000000;
					break;
				}
				break;

			case Token::eComma:
				eState = eError;
				break;
			}
			break;

		case eAlt:
			scpShortCut.m_nControlKeys |= FALT;
			switch (pToken->m_eType)
			{
			case Token::eControl:
				eState = eControl;
				break;

			case Token::eShift:
				eState = eShift;
				break;

			case Token::eChar:
				scpShortCut.m_nKeyCodes[0] = pToken->m_cCharacter;
				eState = eChar;
				switch (scpShortCut.m_nKeyCodes[0])
				{
				case VK_INSERT:
				case VK_DELETE:
//				case VK_BACK:
				case VK_HOME:
				case VK_END:
				case VK_PRIOR:
				case VK_NEXT:
				case VK_UP:
				case VK_DOWN:
				case VK_LEFT:
				case VK_RIGHT:
					scpShortCut.m_nKeyboardCodes[0] = 0x01000000;
					break;
				}
				break;

			case Token::eAlt:
			case Token::eComma:
				eState = eError;
				break;
			}
			break;

		case eControl:
			scpShortCut.m_nControlKeys |= FCONTROL;
			switch (pToken->m_eType)
			{
			case Token::eShift:
				eState = eShift;
				break;

			case Token::eChar:
				scpShortCut.m_nKeyCodes[0] = pToken->m_cCharacter;
				eState = eChar;
				switch (scpShortCut.m_nKeyCodes[0])
				{
				case VK_INSERT:
				case VK_DELETE:
//				case VK_BACK:
				case VK_HOME:
				case VK_END:
				case VK_PRIOR:
				case VK_NEXT:
				case VK_UP:
				case VK_DOWN:
				case VK_LEFT:
				case VK_RIGHT:
					scpShortCut.m_nKeyboardCodes[0] = 0x01000000;
					break;
				}
				break;

			case Token::eAlt:
			case Token::eControl:
			case Token::eComma:
				eState = eError;
				break;
			}
			break;

		case eShift:
			scpShortCut.m_bShift[0] = TRUE;
			switch (pToken->m_eType)
			{
			case Token::eChar:
				scpShortCut.m_nKeyCodes[0] = pToken->m_cCharacter;
				eState = eChar;
				switch (scpShortCut.m_nKeyCodes[0])
				{
				case VK_INSERT:
				case VK_DELETE:
//				case VK_BACK:
				case VK_HOME:
				case VK_END:
				case VK_PRIOR:
				case VK_NEXT:
				case VK_UP:
				case VK_DOWN:
				case VK_LEFT:
				case VK_RIGHT:
					scpShortCut.m_nKeyboardCodes[0] = 0x01000000;
					break;
				}
				break;

			case Token::eAlt:
			case Token::eControl:
			case Token::eShift:
			case Token::eComma:
				eState = eError;
				break;
			}
			break;

		case eChar:
			if (nIndex < nCount)
			{
				switch (pToken->m_eType)
				{
				case Token::eComma:
					eState = eComma;
					break;
				case Token::eAlt:
				case Token::eControl:
				case Token::eShift:
				case Token::eChar:
					eState = eError;
					break;
				}
			}
			else
				eState = eEnd;
			break;

		case eComma:
			switch (pToken->m_eType)
			{
			case Token::eShift:
				eState = eShift2;
				break;

			case Token::eChar:
				scpShortCut.m_nKeyCodes[1] = pToken->m_cCharacter;
				eState = eChar2;
				switch (scpShortCut.m_nKeyCodes[1])
				{
				case VK_INSERT:
				case VK_DELETE:
//				case VK_BACK:
				case VK_HOME:
				case VK_END:
				case VK_PRIOR:
				case VK_NEXT:
				case VK_UP:
				case VK_DOWN:
				case VK_LEFT:
				case VK_RIGHT:
					scpShortCut.m_nKeyboardCodes[1] = 0x01000000;
					break;
				}
				break;

			case Token::eAlt:
			case Token::eControl:
			case Token::eComma:
				eState = eError;
				break;
			}
			break;

		case eShift2:
			scpShortCut.m_bShift[1] = TRUE;
			switch (pToken->m_eType)
			{
			case Token::eChar:
				scpShortCut.m_nKeyCodes[1] = pToken->m_cCharacter;
				eState = eChar2;
				break;

			case Token::eAlt:
			case Token::eControl:
			case Token::eShift:
			case Token::eComma:
				eState = eError;
				break;
			}
			break;

		case eChar2:
			if (nIndex < nCount)
			{
				switch (pToken->m_eType)
				{
				case Token::eChar:
				case Token::eAlt:
				case Token::eControl:
				case Token::eShift:
				case Token::eComma:
					eState = eError;
					break;
				}
			}
			else
				eState = eEnd;
			break;

		case eError:
			memset(scpShortCut.m_nKeyCodes, 0, sizeof(scpShortCut.m_nKeyCodes));
			memset(scpShortCut.m_bShift, 0, sizeof(scpShortCut.m_bShift));
			memset(scpShortCut.m_nKeyboardCodes, 0, sizeof(scpShortCut.m_nKeyboardCodes));
			scpShortCut.m_nControlKeys = 0;
			bResult = FALSE;
			eState = eEnd;
			break;

		case eEnd:
			break;
		}
		nIndex++;
	}
	return bResult;
}

//
// Exchange
//

HRESULT ShortCutStore::Exchange(IStream* pStream, VARIANT_BOOL vbSave)
{
	try
	{
		short nSize;
		short nSize2;
		long nStreamSize;
		HRESULT hResult;
		if (VARIANT_TRUE == vbSave)
		{
			nStreamSize = sizeof(scpV1) + sizeof(nSize);
			hResult = pStream->Write(&nStreamSize, sizeof(nStreamSize), NULL);
			if (FAILED(hResult))
				return hResult;

			nSize = sizeof(scpV1);
			hResult = pStream->Write(&nSize, sizeof(nSize), NULL);
			if (FAILED(hResult))
				return hResult;

			hResult = pStream->Write(&scpV1, sizeof(scpV1), NULL);
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

			hResult = pStream->Read(&nSize, sizeof(nSize), NULL);
			if (FAILED(hResult))
				return hResult;

			nStreamSize -= sizeof(nSize);
			if (nStreamSize <= 0)
				goto FinishedReading;

			nSize2 = sizeof(scpV1);
			hResult = pStream->Read(&scpV1, nSize < nSize2 ? nSize : nSize2, NULL);
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

//
// CShortCut
//

CShortCut::CShortCut()
	: m_pShortCutStore(NULL)
{
	m_nLastToolId = -1;
}

CShortCut::~CShortCut()
{
	if (m_pShortCutStore)
		m_pShortCutStore->Release();
}

//
// AddTool
//

HRESULT CShortCut::AddTool(CTool* pTool)
{
	BOOL bFound = FALSE;
	int nCount = m_aTools.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		if (m_aTools.GetAt(nIndex) == pTool)
		{
			bFound = TRUE;
			break;
		}
	}
	if (!bFound)
	{
		HRESULT hResult = m_aTools.Add(pTool);
		if (FAILED(hResult))
			return hResult;
	}
	return NOERROR;
}

//
// GetTool
//

HRESULT CShortCut::GetTool(CTool*& pTool)
{
	int nCount = m_aTools.GetSize();
	if (-1 == m_nLastToolId && nCount > 0)
		m_nLastToolId = 0;
	else if (m_nLastToolId > nCount - 1)
		m_nLastToolId = nCount - 1;

	if (-1 == m_nLastToolId)
	{
		pTool = NULL;
		return E_FAIL;
	}
	else
		pTool = m_aTools.GetAt(m_nLastToolId);
	return NOERROR;
}

//
// RemoveTool
//

HRESULT CShortCut::RemoveTool(CTool* pTool)
{
	int nCount = m_aTools.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
	{
		if (m_aTools.GetAt(nIndex) == pTool)
		{
			m_aTools.RemoveAt(nIndex);
			return NOERROR;
		}
	}
	return E_FAIL;
}

//
// CShortCuts
//

CShortCuts::CShortCuts()
	: m_pShortCutStore(NULL)
{
	m_pShortCutStore = ShortCutStore::CreateInstance(NULL);
	assert(m_pShortCutStore);
}

CShortCuts::~CShortCuts()
{
	CleanUp();
	if (m_pShortCutStore)
		m_pShortCutStore->Release();
}

//
// CleanUp
//

void CShortCuts::CleanUp()
{
	int nCount = m_aShortCuts.GetSize();
	for (int nIndex = 0; nIndex < nCount; nIndex++)
		delete m_aShortCuts.GetAt(nIndex);
	m_aShortCuts.RemoveAll();
}

//
// SetOwner
//

void CShortCuts::SetOwner(CBar* pBar)
{
	m_pBar = pBar;
}

//
// Count
//

HRESULT CShortCuts::Count(short* pnCount)
{
	*pnCount = m_aShortCuts.GetSize();
	return NOERROR;
}

//
// Add
//

HRESULT CShortCuts::Add(ShortCutStore* pShortCutStore, CTool* pTool)
{
	CShortCut* pShortCut;
	HRESULT hResult = Find(pShortCutStore, pShortCut);
	if (SUCCEEDED(hResult))
	{
		pShortCut->AddTool(pTool);
		return NOERROR;
	}
	else
	{
		pShortCut = new CShortCut;
		assert(pShortCut);
		if (pShortCut)
		{
			pShortCut->SetShortCut(pShortCutStore);
			pShortCut->AddTool(pTool);
			hResult = m_aShortCuts.Add(pShortCut);
			if (FAILED(hResult))
				return hResult;
			return NOERROR;
		}
	}
	return E_FAIL;
}

//
// Remove
//
//
// This function removes a tool from the shortcut collection.  If there are no more tools left
// in it's collection delete the short out of the collection.
//

HRESULT CShortCuts::Remove(const ShortCutStore* pShortCutStore, CTool* pTool)
{
	try
	{
		CShortCut* pShortCut;
		HRESULT hResult = Find(pShortCutStore, pShortCut);
		if (SUCCEEDED(hResult) && pShortCut)
		{
			 hResult = pShortCut->RemoveTool(pTool);
			 if (FAILED(hResult) || pShortCut->ToolCount() < 1)
			 {
				int nCount = m_aShortCuts.GetSize();
				for (int nIndex = 0; nIndex < nCount; nIndex++)
				{
					if (pShortCut == m_aShortCuts.GetAt(nIndex))
					{
						m_aShortCuts.RemoveAt(nIndex);
						delete pShortCut;
						break;
					}
				}
			 }
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
// Find
//

HRESULT CShortCuts::Find(const ShortCutStore* pShortCutStore,  CShortCut*& pShortCut)
{
	try
	{
		int nCount = m_aShortCuts.GetSize();
		for (int nIndex = 0; nIndex < nCount; nIndex++)
		{
			try
			{
				pShortCut = m_aShortCuts.GetAt(nIndex);
				if (NULL == pShortCut)
					continue;

				if ((*pShortCutStore) != (*pShortCut->GetShortCutStore()))
					continue;

				return NOERROR;
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
	pShortCut = NULL;
	return E_FAIL;
}

//
// FireDelayedToolClick
//

void CShortCuts::FireDelayedToolClick(CTool* pTool)
{
	try
	{
		if (NULL == pTool)
			return;

		if (m_pBar->m_bMenuLoop)
			m_pBar->m_bMenuLoop = FALSE;

		if (m_pBar->m_pPopupRoot)
			m_pBar->m_pPopupRoot->SetPopupIndex(CPopupWin::eRemoveSelection);

		pTool->AddRef();

		::PostMessage(m_pBar->m_hWnd, GetGlobals().WM_ACTIVEBARCLICK, 0, (LPARAM)pTool);
		if (m_pBar->m_bIsVisualFoxPro)
		{
			pTool->AddRef();
			::PostMessage(m_pBar->m_hWnd, GetGlobals().WM_ACTIVEBARCLICK, 0, (LPARAM)pTool);
		}
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
	}
}

//
// Process
//

BOOL CShortCuts::Process(long nControlState, WPARAM wKey, LPARAM lParam)
{
	CTool* pTool;
	CShortCut* pShortCut;
	BOOL bShift = FALSE;
	if (GetKeyState(VK_SHIFT) < 0)
		bShift = TRUE;
	
	if (m_pShortCutStore->scpV1.m_nKeyCodes[0])
	{
		//
		// Valid Second Keys
		//
		if ((wKey >= ' ' && wKey <= '@') ||
			(wKey >= '[' && wKey <= '`') || 
			(wKey >= '{' && wKey <= '~')
		   )
		{
			m_pShortCutStore->scpV1.m_nKeyCodes[1] = wKey;
			if (bShift)
				m_pShortCutStore->scpV1.m_bShift[1] = TRUE;
			HRESULT hResult = Find(m_pShortCutStore, pShortCut);
			if (SUCCEEDED(hResult) && pShortCut)
			{
				hResult = pShortCut->GetTool(pTool);
				if (SUCCEEDED(hResult) && VARIANT_TRUE == pTool->tpV1.m_vbEnabled && VARIANT_TRUE == pTool->tpV1.m_vbVisible)
					FireDelayedToolClick(pTool);
				m_pShortCutStore->Clear();
				return TRUE;
			}
			m_pShortCutStore->Clear();
			return ValidFirstKey(bShift, nControlState, wKey, lParam);
		}
		else if ((wKey >= VK_NUMPAD0 && wKey <= VK_F24) ||
		         (wKey >= VK_INSERT && wKey <= VK_DELETE) ||
		         (wKey >= VK_PRIOR && wKey <= VK_DOWN) ||
				 (wKey >= 'A' && wKey <= 'Z') ||
				 (196 == wKey || 197 == wKey || 214 == wKey) ||
			     (wKey == VK_BACK) ||
			     (wKey == VK_RETURN) ||
			     (wKey == VK_SCROLL)
	            )
		{
			//
			// Add shift to these keys if it is pressed
			//

			if (bShift)
				m_pShortCutStore->scpV1.m_bShift[1] = TRUE;
			m_pShortCutStore->scpV1.m_nKeyCodes[1] = wKey;
			HRESULT hResult = Find(m_pShortCutStore, pShortCut);
			if (SUCCEEDED(hResult) && pShortCut)
			{
				hResult = pShortCut->GetTool(pTool);
				if (SUCCEEDED(hResult) && VARIANT_TRUE == pTool->tpV1.m_vbEnabled && VARIANT_TRUE == pTool->tpV1.m_vbVisible)
					FireDelayedToolClick(pTool);
				m_pShortCutStore->Clear();
				return TRUE;
			}
			m_pShortCutStore->Clear();
			return ValidFirstKey(bShift, nControlState, wKey, lParam);
		}
		else 
			m_pShortCutStore->Clear();
	}
	else if (0 == nControlState)
	{
		if ((wKey >= VK_F1 && wKey <= VK_F24) ||
		    (wKey >= VK_INSERT && wKey <= VK_DELETE) ||
		    (wKey >= VK_PRIOR && wKey <= VK_DOWN)
	       )
		{
			m_pShortCutStore->scpV1.m_nKeyCodes[0] = wKey;
			if (bShift)
				m_pShortCutStore->scpV1.m_bShift[0] = TRUE;
			HRESULT hResult = Find(m_pShortCutStore, pShortCut);
			if (SUCCEEDED(hResult) && pShortCut)
			{
				hResult = pShortCut->GetTool(pTool);
				if (SUCCEEDED(hResult) && VARIANT_TRUE == pTool->tpV1.m_vbEnabled && VARIANT_TRUE == pTool->tpV1.m_vbVisible)
					FireDelayedToolClick(pTool);
				m_pShortCutStore->Clear();
				return TRUE;
			}
		}
	}
	else
	{
		//
		// Valid First Keys with control keys being pressed
		//
		
		return ValidFirstKey(bShift, nControlState, wKey, lParam);

	}
	return FALSE;
}

//
// ValidFirstKey
//

BOOL CShortCuts::ValidFirstKey(BOOL bShift, long nControlState, WPARAM wKey, LPARAM lParam)
{
	CTool* pTool;
	CShortCut* pShortCut;

	if ((wKey >= ' ' && wKey <= '@') ||
		(wKey >= '[' && wKey <= '`') || 
		(wKey >= '{' && wKey <= '~')
	   )
	{
		m_pShortCutStore->scpV1.m_nKeyCodes[0] = wKey;
		m_pShortCutStore->scpV1.m_nControlKeys = nControlState;
		if (bShift)
			m_pShortCutStore->scpV1.m_bShift[0] = TRUE;

		HRESULT hResult = Find(m_pShortCutStore, pShortCut);
		if (SUCCEEDED(hResult) && pShortCut)
		{
			hResult = pShortCut->GetTool(pTool);
			if (SUCCEEDED(hResult) && VARIANT_TRUE == pTool->tpV1.m_vbEnabled && VARIANT_TRUE == pTool->tpV1.m_vbVisible)
				FireDelayedToolClick(pTool);
			
			m_pShortCutStore->Clear();
			return TRUE;
		}
	}
	else if ((wKey >= VK_NUMPAD0 && wKey <= VK_F24) ||
			 (wKey >= VK_INSERT && wKey <= VK_DELETE) ||
			 (wKey >= VK_PRIOR && wKey <= VK_DOWN) ||
			 (wKey >= 'A' && wKey <= 'Z') ||
			 (196 == wKey || 197 == wKey || 214 == wKey) ||
			 (wKey == VK_BACK) ||
			 (wKey == VK_RETURN) ||
			 (wKey == VK_SCROLL)
	        )
	{
		//
		// Add shift to these keys if it is pressed
		//
		if (bShift)
			m_pShortCutStore->scpV1.m_bShift[0] = TRUE;
		m_pShortCutStore->scpV1.m_nKeyCodes[0] = wKey;
		m_pShortCutStore->scpV1.m_nControlKeys = nControlState;
		
		HRESULT hResult = Find(m_pShortCutStore, pShortCut);
		if (SUCCEEDED(hResult) && pShortCut)
		{
			hResult = pShortCut->GetTool(pTool);
			m_pShortCutStore->Clear();
			if (SUCCEEDED(hResult) && VARIANT_TRUE == pTool->tpV1.m_vbEnabled && VARIANT_TRUE == pTool->tpV1.m_vbVisible)
			{
				FireDelayedToolClick(pTool);
				return TRUE;
			}
		}
	}
	return FALSE;
}

//
// TranslateAccelerator
//

BOOL CShortCuts::TranslateAccelerator(MSG* pMsg)
{
	try
	{
		long nKeyState = 0;
		switch (pMsg->message)
		{
		case WM_SYSKEYDOWN:
		    if (GetKeyState(VK_CONTROL) < 0)
				nKeyState |= FCONTROL;
			
			if (GetKeyState(VK_MENU) < 0)
				nKeyState |= FALT;
			
			if (Process(nKeyState, pMsg->wParam, pMsg->lParam))
				return TRUE;
			break;

		case WM_CHAR:
		case WM_KEYDOWN:
			if (GetKeyState(VK_CONTROL) < 0)
				nKeyState |= FCONTROL;

			if (GetKeyState(VK_MENU) < 0)
				nKeyState |= FALT;

			if (Process(nKeyState, pMsg->wParam, pMsg->lParam))
				return TRUE;
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
