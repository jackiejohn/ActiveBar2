#ifndef DRAGDROP_INCLUDED
#define DRAGDROP_INCLUDED

#include "Map.h"
interface ITool;
struct BandDragDrop;
class CTool;
class CDock;
class CBand;
class FWnd;

#define	DROPEFFECT_BAR	( 10 )

//
// Drag Drop
//

class CDragDropMgr
{
public:
	HRESULT RegisterDragDrop(HWND hWnd, IDropTarget* pDropTarget);
	HRESULT RevokeDragDrop(HWND hWnd);

	HRESULT DoDragDrop(IDataObject* pDataObject, 
					   IDropSource* pDropSource, 
					   DWORD	    dwOKEffect, 
					   DWORD*		pdwEffect);
private:
	TypedMap<HWND, IDropTarget*> m_mapTargets;
};

//
// CToolDataObject
//

class CToolDataObject : public IDataObject
{
public:
	CToolDataObject(ITool* pTool)
		: m_pTool(pTool)
	{
		m_cRef = 1;
	}

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
	ITool* m_pTool;
	long   m_cRef;
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
	FORMATETC		 m_FmEtc[3];
	long			 m_cRef;
	long			 m_nIndex;
};

//
// CToolDropSource
//

class CToolDropSource : public IDropSource
{
public:
	CToolDropSource()
	{
		m_cRef = 1;
	}

    virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject);
        
    virtual ULONG __stdcall AddRef();
        
    virtual ULONG __stdcall Release();

	virtual HRESULT __stdcall QueryContinueDrag(BOOL  bEscapePressed,
												DWORD grfKeyState);
        
    virtual HRESULT __stdcall GiveFeedback(DWORD dwEffect);
private:
	long m_cRef;
};

//
// BandDragDrop
//

struct BandDragDrop
{
	BandDragDrop()
		: m_pBar(NULL),
		  m_pWnd(NULL),
		  m_pTool(NULL)
	{
	}

	~BandDragDrop();

//	virtual BOOL DragEnter(const POINT& pt);
//	virtual BOOL DragOver(const POINT& pt);
//	virtual BOOL DragLeave();
//	virtual BOOL Drop(const POINT& pt, IDataObject* pDataObject);

	virtual CBand* GetBand(POINT pt) = 0;
	virtual void OffsetPoint(CBand* pBand, POINT& pt) {};

	CTool* m_pTool;
	CBar* m_pBar;
	FWnd* m_pWnd;
};

//
// CBandDropTarget
//

class CBandDropTarget : public IDropTarget
{
public:

	CBandDropTarget()
		: m_pBandDragDrop(NULL),
		  m_pDataObject(NULL)
	{
		m_cRef = 1;
	}

	~CBandDropTarget();

	void SetBandDragDrop (BandDragDrop* pBandDragDrop);

	virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject);
        
    virtual ULONG __stdcall AddRef();
        
    virtual ULONG __stdcall Release();

    virtual HRESULT __stdcall DragEnter(IDataObject* pDataObj, 
										DWORD        grfKeyState,
										POINTL       pt,
										DWORD*       pdwEffect);
        
    virtual HRESULT __stdcall DragOver(DWORD  grfKeyState,
									   POINTL pt,
									   DWORD* pdwEffect);
        
    virtual HRESULT __stdcall DragLeave();
        
    virtual HRESULT __stdcall Drop(IDataObject* pDataObj,
								   DWORD		grfKeyState,
								   POINTL		pt,
								   DWORD*		pdwEffect);
private:
	long		  m_cRef;
	BandDragDrop* m_pBandDragDrop;
	IDataObject*  m_pDataObject;
	CTool*		  m_pTool;
};

inline void CBandDropTarget::SetBandDragDrop (BandDragDrop* pBandDragDrop)
{
	m_pBandDragDrop = pBandDragDrop;
}

void SetDefFormatEtc(FORMATETC& fe, CLIPFORMAT& cf, DWORD dwMed);

#endif