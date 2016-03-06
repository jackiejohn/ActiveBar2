#ifndef __CBands_H__
#define __CBands_H__

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "Interfaces.h"
class CBand;

class CBands : public IBands,public ISupportErrorInfo,public ICategorizeProperties
{
public:
	CBands();
	ULONG m_refCount;
	int m_objectIndex;
	~CBands();

	//{BEGIN INTERFACEDEFS}
	static void *objectDef;
	static CBands *CreateInstance(IUnknown *pUnkOuter);
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
	
	

	// IBands members

	STDMETHOD(Item)( VARIANT *Index, Band **retval);
	STDMETHOD(Count)( short *retval);
	STDMETHOD(Add)( BSTR name, Band **retval);
	STDMETHOD(Remove)( VARIANT *Index);
	STDMETHOD(RemoveAll)();
	STDMETHOD(InsertBand)( int index, IBand *pBand);
	STDMETHOD(_NewEnum)( IUnknown **retval);
	// DEFS for IBands
	
	

	// ISupportErrorInfo members

	STDMETHOD(InterfaceSupportsErrorInfo)( REFIID riid);
	// DEFS for ISupportErrorInfo
	
	

	// ICategorizeProperties members

	STDMETHOD(MapPropertyToCategory)( DISPID dispid,  PROPCAT*ppropcat);
	STDMETHOD(GetCategoryName)( PROPCAT propcat,  LCID lcid,  BSTR*pbstrName);
	// DEFS for ICategorizeProperties
	
	
	// Events 
	//{END INTERFACEDEFS}
#ifdef _DEBUG
	void Dump(DumpContext& dc);
#endif
	HRESULT Exchange(IStream* pStream, VARIANT_BOOL vbSave, BSTR bstrBandName, short nBandCount);
	HRESULT ExchangeConfig(IStream* pStream, VARIANT_BOOL vbSave, short nBandCount);
	HRESULT ExchangeMenuUsageData(IStream* pStream, VARIANT_BOOL vbSave);

	HRESULT ClearMenuUsageData();

	HRESULT RemoveEx(CBand* pBand);
	CBand* GetBand(int nIndex);
	BOOL CheckForStatusBand(BandTypes btType);
	void SetOwner(CBar* pBar);
	int GetBandCount();
	int GetPosOfItem(VARIANT* pvIndex);
	CBand* Find(BSTR bstrName);
	void RefreshFloating();

private:
	TypedArray<CBand*> m_aBands;
	CBar* m_pBar;
};

extern ITypeInfo *GetObjectTypeInfo(LCID lcid,void *objectDef);

inline void CBands::SetOwner(CBar* pBar)
{
	m_pBar = pBar;
}

inline CBand* CBands::GetBand(int nIndex)
{
	return m_aBands.GetAt(nIndex);
}

inline int CBands::GetBandCount()
{
	return m_aBands.GetSize();
}

#endif
