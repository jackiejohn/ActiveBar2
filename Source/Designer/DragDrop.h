#ifndef DRAGDROP_INCLUDED
#define DRAGDROP_INCLUDED

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

//
// CToolDataObject
//

class CToolDataObject : public IDataObject
{
public:
	enum Source
	{
		eActiveBarDragDropId = 1010,
		eDesignerDragDropId = 1111,
		eDesignerMainToolDragDropId = 2222,
		eLibraryDragDropId = 3333
	};

	enum TYPE
	{
		eToolId,
		eBandToolId,
		eBandChildBandToolId,
		eBand,
		eCategory,
		eTool
	};

	CToolDataObject(Source eSource, IActiveBar2* pBar, TypedArray<ITool*>& aTools, BSTR bstrBand = NULL, BSTR bstrChildBand = NULL);
	CToolDataObject(Source eSource, IActiveBar2* pBar, BSTR bstrCategory, TypedArray<ITool*>& aTools);
	CToolDataObject(Source eSource, IActiveBar2* pBar, IBand* pBand, TypedArray<ITool*>& aTools);
	~CToolDataObject();

    virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject);
        
    virtual ULONG __stdcall AddRef();
        
    virtual ULONG __stdcall Release();

    virtual HRESULT __stdcall GetData(FORMATETC* pFormatEtcIn,
									  STGMEDIUM* pMedium);
        
    virtual HRESULT __stdcall GetDataHere(FORMATETC* pFormatEtc,
										  STGMEDIUM* pMedium);
        
    virtual HRESULT __stdcall QueryGetData(FORMATETC* pFormatEtc);
        
    virtual HRESULT __stdcall GetCanonicalFormatEtc(FORMATETC* pFormatEctIn,
													FORMATETC* pFormatEtcOut);
        
    virtual HRESULT __stdcall SetData(FORMATETC* pFormatEtc,
									  STGMEDIUM* pMedium,
									  BOOL       bRelease);
        
    virtual HRESULT __stdcall EnumFormatEtc(DWORD			 dwDirection,
											IEnumFORMATETC** ppEnumFormatEtc);
        
    virtual HRESULT __stdcall DAdvise(FORMATETC*   pFormatEtc, 
									  DWORD		   advf, 
									  IAdviseSink* pAdvSink,
									  DWORD*       pdwConnection);
        
    virtual HRESULT __stdcall DUnadvise(DWORD dwConnection);
        
    virtual HRESULT __stdcall EnumDAdvise(IEnumSTATDATA** ppEnumAdvise);

private:
	TypedArray<ITool*>& m_aTools;
	IActiveBar2* m_pBar;
	IBand*       m_pBand;
	BSTR   m_bstrChildBand;
	BSTR   m_bstrCategory;
	BSTR   m_bstrBand;
	long   m_cRef;
	TYPE   m_eType;
	Source m_eSource;
};

//
// CToolEnumFORMATETC
//

class CToolEnumFORMATETC : public IEnumFORMATETC
{
public:
	CToolEnumFORMATETC(CToolDataObject* pToolDataObject);

    virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject);
        
    virtual ULONG __stdcall AddRef();
        
    virtual ULONG __stdcall Release();

	virtual HRESULT __stdcall Next(ULONG	  celt, 
								   FORMATETC* pFormatetc, 
								   ULONG*	  plFetched);

	virtual HRESULT __stdcall Skip(ULONG celt);
	
	virtual HRESULT __stdcall Reset();
	
	virtual HRESULT __stdcall Clone(IEnumFORMATETC** ppEnum);

private:
	CToolDataObject* m_pToolDataObject;
	FORMATETC		 m_FmEtc[6];
	long			 m_cRef;
	long			 m_nIndex;
};

#endif