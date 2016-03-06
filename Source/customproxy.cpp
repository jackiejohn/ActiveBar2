//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

// Machine generated IDispatch wrapper class(es) created with ClassWizard

#include "precomp.h"
#include "customproxy.h"
#include "xevents.h"
#include "guids.h"


#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


CCustomTool::CCustomTool(LPDISPATCH pDispatch)
 : COleDispatchDriver(pDispatch) 
{
	m_pTool2=NULL;
	pDispatch->QueryInterface(IID_ICustomTool2,(LPVOID *)&m_pTool2);
}

CCustomTool::~CCustomTool()
{
	if (m_pTool2)
		m_pTool2->Release();
}

/////////////////////////////////////////////////////////////////////////////
// CCustomTool properties

/////////////////////////////////////////////////////////////////////////////
// CCustomTool operations
extern HRESULT DispatchHelper(LPDISPATCH lpDispatch,DISPID dwDispID, WORD wFlags,
	VARTYPE vtRet, void* pvRet, const BYTE* pbParamInfo, va_list argList);

void COleDispatchDriver::InvokeHelper( DISPID dwDispID, WORD wFlags, VARTYPE vtRet, 
		void* pvRet, const BYTE FAR* pbParamInfo, ... )
{
	va_list argList;
	va_start(argList, pbParamInfo);
	
	if (m_pDisp)
		hr=DispatchHelper(m_pDisp,dwDispID, wFlags,vtRet,pvRet,pbParamInfo,argList);
	va_end(argList);
}

void CCustomTool::SetHost(LPDISPATCH host)
{
	if (m_pTool2)
		hr=m_pTool2->SetHost(host);
	else
	{
		static BYTE parms[] =VTS_DISPATCH;
		InvokeHelper(0x90, DISPATCH_METHOD, VT_EMPTY, NULL, parms,host);
	}
}

void CCustomTool::GetFlags(long* pFlags)
{
	if (m_pTool2)
		hr=m_pTool2->GetFlags(pFlags);
	else
	{
		static BYTE parms[] =VTS_PI4;
		InvokeHelper(0x9d, DISPATCH_METHOD, VT_EMPTY, NULL, parms,pFlags);
	}
}

void CCustomTool::OnDraw(int hdc, int x, int y, int width, int height)
{
	if (m_pTool2)
		hr=m_pTool2->OnDraw(hdc,x,y,width,height);
	else
	{
		static BYTE parms[] =VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_I4;
		InvokeHelper(0x91, DISPATCH_METHOD, VT_EMPTY, NULL, parms,hdc, x, y, width, height);
	}
}

void CCustomTool::GetSize(int* width, int * height)
{
	if (m_pTool2)
		hr=m_pTool2->GetSize(width,height);
	else
	{
		static BYTE parms[] =VTS_PI4 VTS_PI4;
		InvokeHelper(0x2f, DISPATCH_METHOD, VT_EMPTY, NULL, parms,width, height);
	}
}

void CCustomTool::OnMouseEnter()
{
	if (m_pTool2)
		hr=m_pTool2->OnMouseEnter();
	else
		InvokeHelper(0x92, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CCustomTool::OnMouseExit()
{
	if (m_pTool2)
		hr=m_pTool2->OnMouseExit();
	else
		InvokeHelper(0x93, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void CCustomTool::OnMouseDown(short Button, int x, int y)
{
	if (m_pTool2)
		hr=m_pTool2->OnMouseDown(Button,x,y);
	else
	{
		static BYTE parms[] =VTS_I2 VTS_I4 VTS_I4;
		InvokeHelper(0x94, DISPATCH_METHOD, VT_EMPTY, NULL, parms,Button, x, y);
	}
}

void CCustomTool::OnMouseUp(short Button, int x, int y)
{
	if (m_pTool2)
		hr=m_pTool2->OnMouseUp(Button,x,y);
	else
	{
		static BYTE parms[] =VTS_I2 VTS_I4 VTS_I4;
		InvokeHelper(0x95, DISPATCH_METHOD, VT_EMPTY, NULL, parms,Button, x, y);
	}
}

void CCustomTool::OnMouseDblClk(short Button)
{
	if (m_pTool2)
		hr=m_pTool2->OnMouseDblClk(Button);
	else
	{
		static BYTE parms[] =VTS_I2;
		InvokeHelper(0x96, DISPATCH_METHOD, VT_EMPTY, NULL, parms,Button);
	}
}
