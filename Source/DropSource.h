#ifndef DROPSOURCE_INCLUDED
#define DROPSOURCE_INCLUDED

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

interface ITool;
struct BandDragDrop;
class CTool;
class CDock;
class CBand;
class FWnd;

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

	virtual HRESULT __stdcall QueryContinueDrag(BOOL  bEscapePressed, DWORD grfKeyState);
        
    virtual HRESULT __stdcall GiveFeedback(DWORD dwEffect);
private:
	long m_cRef;
};

//
// BandDragDrop
//

struct BandDragDrop
{
	BandDragDrop();
	~BandDragDrop();

	virtual CBand* GetBand(POINT pt) = 0;
	virtual void OffsetPoint(CBand* pBand, POINT& pt);

	FWnd* m_pWnd;
};

inline void BandDragDrop::OffsetPoint(CBand* pBand, POINT& pt)
{
	assert(m_pWnd);
	if (m_pWnd)
		m_pWnd->ScreenToClient(pt);
};

//
// CBandDropTarget
//

class CBandDropTarget : public IDropTarget
{
public:
	enum MISC
	{
		eNumOfTypes = 6
	};

	CBandDropTarget(long nDestId);
	~CBandDropTarget();

	void Init (BandDragDrop* pBandDragDrop, IActiveBar2* pDestBar, long nDestId);

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
	IActiveBar2*  m_pDestBar;
	IActiveBar2*  m_pSrcBar;
	IDataObject*  m_pDataObject;
	BandDragDrop* m_pBandDragDrop;
	FORMATETC     m_fe[eNumOfTypes];
	ITool*		  m_pTool;
	long		  m_nFormatType;
	long		  m_nDestId;
	long		  m_nSrcId;
	long		  m_cRef;
	int			  m_nIndex;
};

inline void CBandDropTarget::Init (BandDragDrop* pBandDragDrop, IActiveBar2* pDestBar, long nDestId)
{
	m_pBandDragDrop = pBandDragDrop;
	m_pDestBar = pDestBar;
	m_nDestId = nDestId;
}

#endif