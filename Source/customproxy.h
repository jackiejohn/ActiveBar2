// Machine generated IDispatch wrapper class(es) created with ClassWizard
/////////////////////////////////////////////////////////////////////////////
// CCustomTool wrapper class

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "Interfaces.h"
#include "PrivateInterfaces.h"

class COleDispatchDriver
{
public:
	COleDispatchDriver() {m_pDisp=NULL;};
	COleDispatchDriver(LPDISPATCH pDisp) {m_pDisp=pDisp; if (pDisp) pDisp->AddRef();};
	virtual ~COleDispatchDriver() {if (m_pDisp) m_pDisp->Release();};

	HRESULT hr;

protected:
	IDispatch *m_pDisp;
	
	void InvokeHelper( DISPID dwDispID, WORD wFlags, VARTYPE vtRet, 
		void* pvRet, const BYTE FAR* pbParamInfo, ... );
};


class CCustomTool : public COleDispatchDriver
{
public:
	CCustomTool() {m_pTool2=NULL;}		// Calls COleDispatchDriver default constructor
	CCustomTool(LPDISPATCH pDispatch); 
	virtual ~CCustomTool();
	ICustomTool2 *m_pTool2;		
// Attributes
public:

// Operations
public:
	void SetHost(LPDISPATCH host);
	void GetFlags(long* pFlags);
	void OnDraw(int hdc, int x, int y, int width, int height);
	void GetSize(int* width, int* height);
	void OnMouseEnter();
	void OnMouseExit();
	void OnMouseDown(short Button, int x, int y);
	void OnMouseUp(short Button, int x, int y);
	void OnMouseDblClk(short Button);
};
