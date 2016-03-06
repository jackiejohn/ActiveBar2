//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "debug.h"
#include "support.h"
#include "xevents.h"


#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

const DWORD VTS_BYREF=0x40;

HRESULT DispatchHelper(LPDISPATCH lpDispatch,DISPID dwDispID, WORD wFlags,
	VARTYPE vtRet, void* pvRet, const BYTE* pbParamInfo, va_list argList)
{
	if (lpDispatch == NULL)
	{
		TRACE(1, "Warning: attempt to call Invoke with NULL lpDispatch!\n");
		return E_FAIL;
	}

	DISPPARAMS dispparams;
	memset(&dispparams, 0, sizeof dispparams);

	// determine number of arguments
	if (pbParamInfo != NULL)
	{
		dispparams.cArgs = lstrlenA((LPCSTR)pbParamInfo);
	}

	DISPID dispidNamed = DISPID_PROPERTYPUT;
	if (wFlags & (DISPATCH_PROPERTYPUT|DISPATCH_PROPERTYPUTREF))
	{
		ASSERT(dispparams.cArgs > 0,"Invalid parameter count in Dispatch Helper");
		dispparams.cNamedArgs = 1;
		dispparams.rgdispidNamedArgs = &dispidNamed;
	}

	if (dispparams.cArgs != 0)
	{
		// allocate memory for all VARIANT parameters
		VARIANT* pArg = new VARIANT[dispparams.cArgs];
		if (pArg == NULL)
			return E_OUTOFMEMORY;
		dispparams.rgvarg = pArg;
		memset(pArg, 0, sizeof(VARIANT) * dispparams.cArgs);

		// get ready to walk vararg list
		const BYTE* pb = pbParamInfo;
		pArg += dispparams.cArgs - 1;   // params go in opposite order

		while (*pb != 0)
		{
			ASSERT(pArg >= dispparams.rgvarg,"DispatchHelper error 001");
			pArg->vt = *pb; // set the variant type
			switch (pArg->vt)
			{
			case VT_I2:
				pArg->iVal = va_arg(argList, short);
				break;
			case VT_I4:
				pArg->lVal = va_arg(argList, long);
				break;
			case VT_R4:
				{
					// Note: All float arguments to vararg functions are passed
					//  as doubles instead.  Thats why they are passed as VT_R8
					//  instead of VT_R4.
					pArg->vt = VT_R8;
					*(DOUBLE_ARG*)&pArg->dblVal = va_arg(argList, DOUBLE_ARG);
					VariantChangeType(pArg, pArg, VARIANT_NOVALUEPROP, VT_R4);
				}
				break;
			case VT_R8:
				*(DOUBLE_ARG*)&pArg->dblVal = va_arg(argList, DOUBLE_ARG);
				break;
			case VT_DATE:
				*(DOUBLE_ARG*)&pArg->date = va_arg(argList, DOUBLE_ARG);
				break;
			case VT_CY:
				pArg->cyVal = *va_arg(argList, CY*);
				break;
			case VT_BSTR:
				{
					LPCOLESTR lpsz = va_arg(argList, LPOLESTR);
					pArg->bstrVal = ::SysAllocString(lpsz);
					if (lpsz != NULL && pArg->bstrVal == NULL)
						return E_OUTOFMEMORY;
				}
				break;
#if !defined(_UNICODE) && !defined(OLE2ANSI)
			case VT_BSTRA:
				{
					LPCSTR lpsz = va_arg(argList, LPSTR);
					MAKE_WIDEPTR_FROMANSI(pwsz,lpsz);
					pArg->bstrVal = ::SysAllocString(pwsz);
					if (lpsz != NULL && pArg->bstrVal == NULL)
						return E_OUTOFMEMORY;
					pArg->vt = VT_BSTR;
				}
				break;
#endif
			case VT_DISPATCH:
				pArg->pdispVal = va_arg(argList, LPDISPATCH);
				break;
			case VT_ERROR:
				pArg->scode = va_arg(argList, SCODE);
				break;
			case VT_BOOL:
				V_BOOL(pArg) = (VARIANT_BOOL)(va_arg(argList, BOOL) ? -1 : 0);
				break;
			case VT_VARIANT:
				*pArg = *va_arg(argList, VARIANT*);
				break;
			case VT_UNKNOWN:
				pArg->punkVal = va_arg(argList, LPUNKNOWN);
				break;

			case VT_I2|VTS_BYREF:
				pArg->vt=VT_I2|VT_BYREF;
				pArg->piVal = va_arg(argList, short*);
				break;
			case VT_I4|VTS_BYREF:
				pArg->vt=VT_I4|VT_BYREF;
				pArg->plVal = va_arg(argList, long*);
				break;
			case VT_R4|VTS_BYREF:
				pArg->vt=VT_R4|VT_BYREF;
				pArg->pfltVal = va_arg(argList, float*);
				break;
			case VT_R8|VTS_BYREF:
				pArg->vt=VT_R8|VT_BYREF;
				pArg->pdblVal = va_arg(argList, double*);
				break;
			case VT_DATE|VTS_BYREF:
				pArg->vt=VT_DATE|VT_BYREF;
				pArg->pdate = va_arg(argList, DATE*);
				break;
			case VT_CY|VTS_BYREF:
				pArg->vt=VT_CY|VT_BYREF;
				pArg->pcyVal = va_arg(argList, CY*);
				break;
			case VT_BSTR|VTS_BYREF:
				pArg->vt=VT_BSTR|VT_BYREF;
				pArg->pbstrVal = va_arg(argList, BSTR*);
				break;
			case VT_DISPATCH|VTS_BYREF:
				pArg->vt=VT_DISPATCH|VT_BYREF;
				pArg->ppdispVal = va_arg(argList, LPDISPATCH*);
				break;
			case VT_ERROR|VTS_BYREF:
				pArg->vt=VT_ERROR|VT_BYREF;
				pArg->pscode = va_arg(argList, SCODE*);
				break;
			case VT_BOOL|VTS_BYREF:
				{
					// coerce BOOL into VARIANT_BOOL
					BOOL* pboolVal = va_arg(argList, BOOL*);
#ifndef _MAC
					*pboolVal = *pboolVal ? MAKELONG(-1, 0) : 0;
#else
					*pboolVal = *pboolVal ? MAKELONG(0, -1) : 0;
#endif
					pArg->vt=VT_BOOL|VT_BYREF;
					pArg->pboolVal = (VARIANT_BOOL*)pboolVal;
				}
				break;
			case VT_VARIANT|VTS_BYREF:
				pArg->vt=VT_VARIANT|VT_BYREF;
				pArg->pvarVal = va_arg(argList, VARIANT*);
				break;
			case VT_UNKNOWN|VTS_BYREF:
				pArg->vt=VT_UNKNOWN|VT_BYREF;
				pArg->ppunkVal = va_arg(argList, LPUNKNOWN*);
				break;

			default:
				ASSERT(FALSE,"Unknown Type in DispatchHelper");
				break;
			}

			--pArg; // get ready to fill next argument
			++pb;
		}
	}

	// initialize return value
	VARIANT* pvarResult = NULL;
	VARIANT vaResult;
	memset(&vaResult,0,sizeof(VARIANT));
	if (vtRet != VT_EMPTY)
		pvarResult = &vaResult;

	// initialize EXCEPINFO struct
	EXCEPINFO excepInfo;
	memset(&excepInfo, 0, sizeof excepInfo);

	UINT nArgErr = (UINT)-1;  // initialize to invalid arg

	// make the call
	SCODE sc = lpDispatch->Invoke(dwDispID, IID_NULL, 0, wFlags,
		&dispparams, pvarResult, &excepInfo, &nArgErr);

	// cleanup any arguments that need cleanup
	if (dispparams.cArgs != 0)
	{
		VARIANT* pArg = dispparams.rgvarg + dispparams.cArgs - 1;
		const BYTE* pb = pbParamInfo;
		while (*pb != 0)
		{
			switch ((VARTYPE)*pb)
			{
#if !defined(_UNICODE) && !defined(OLE2ANSI)
			case VT_BSTRA:
#endif
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
		return ResultFromScode(sc);
	}

	if (vtRet != VT_EMPTY)
	{
		// convert return value
		if (vtRet != VT_VARIANT)
		{
			SCODE sc = VariantChangeType(&vaResult, &vaResult, 0, vtRet);
			if (FAILED(sc))
			{
				TRACE(1, "Warning: automation return value coercion failed.\n");
				VariantClear(&vaResult);
				return ResultFromScode(sc);
			}
			ASSERT(vtRet == vaResult.vt,"Invalid type coercion");
		}

		// copy return value into return spot!
		switch (vtRet)
		{
		case VT_I2:
			*(short*)pvRet = vaResult.iVal;
			break;
		case VT_I4:
			*(long*)pvRet = vaResult.lVal;
			break;
		case VT_R4:
			*(XFLOAT*)pvRet = *(XFLOAT*)&vaResult.fltVal;
			break;
		case VT_R8:
			*(XDOUBLE*)pvRet = *(XDOUBLE*)&vaResult.dblVal;
			break;
		case VT_DATE:
			*(XDOUBLE*)pvRet = *(XDOUBLE*)&vaResult.date;
			break;
		case VT_CY:
			*(CY*)pvRet = vaResult.cyVal;
			break;
		case VT_BSTR:
			*(BSTR *)pvRet=vaResult.bstrVal;
			break;
		case VT_DISPATCH:
			*(LPDISPATCH*)pvRet = vaResult.pdispVal;
			break;
		case VT_ERROR:
			*(SCODE*)pvRet = vaResult.scode;
			break;
		case VT_BOOL:
			*(BOOL*)pvRet = (V_BOOL(&vaResult) != 0);
			break;
		case VT_VARIANT:
			*(VARIANT*)pvRet = vaResult;
			break;
		case VT_UNKNOWN:
			*(LPUNKNOWN*)pvRet = vaResult.punkVal;
			break;

		default:
			ASSERT(FALSE,"invalid return type specified");
		}
	}
	return NOERROR;
}

/*
HRESULT DispatchHelper(LPDISPATCH lpDispatch,DISPID dwDispID, WORD wFlags,VARTYPE vtRet, void* pvRet, const BYTE* pbParamInfo, ...)
{
	HRESULT hr;
	va_list argList;
	va_start(argList, pbParamInfo);
	hr=DispatchHelper2(lpDispatch,dwDispID, wFlags, vtRet, pvRet, pbParamInfo, argList);
	va_end(argList);
	return hr;
}

*/