#ifndef __CTools_H__
#define __CTools_H__
#include "interfaces.h"

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

class CTool;
class CBand;

//
// CTools
//

class CTools : public ITools,public ISupportErrorInfo,public ICategorizeProperties
{
public:
	CTools();
	virtual ~CTools();
	
	ULONG m_refCount;
	int m_objectIndex;
	
	//{BEGIN INTERFACEDEFS}
	static void *objectDef;
	static CTools *CreateInstance(IUnknown *pUnkOuter);
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
	
	

	// ITools members

	STDMETHOD(Item)( VARIANT *Index, Tool **retval);
	STDMETHOD(Count)( short *retval);
	STDMETHOD(Add)( long toolid, BSTR name, Tool **retval);
	STDMETHOD(Remove)( VARIANT *Index);
	STDMETHOD(RemoveAll)();
	STDMETHOD(InsertTool)( int index,  ITool *pTool, VARIANT_BOOL fClone);
	STDMETHOD(CreateTool)( ITool **pTool);
	STDMETHOD(DeleteTool)( ITool *pTool);
	STDMETHOD(CopyTo)( ITools *pTools);
	STDMETHOD(Insert)( int index, Tool *tool);
	STDMETHOD(_NewEnum)( IUnknown **retval);
	STDMETHOD(ItemById)( LONG id, Tool **retval);
	// DEFS for ITools
	
	

	// ISupportErrorInfo members

	STDMETHOD(InterfaceSupportsErrorInfo)( REFIID riid);
	// DEFS for ISupportErrorInfo
	
	

	// ICategorizeProperties members

	STDMETHOD(MapPropertyToCategory)( DISPID dispid,  PROPCAT *ppropcat);
	STDMETHOD(GetCategoryName)( PROPCAT propcat,  LCID lcid, BSTR *pbstrName);
	// DEFS for ICategorizeProperties
	
	
	// Events 
	//{END INTERFACEDEFS}
#ifdef _DEBUG
	void Dump(DumpContext& dc);
#endif

	void SetBar(CBar* pBar);

	void SetBand(CBand* pBand);

	int	GetPosOfItem(VARIANT *pvIndex);

	HRESULT Exchange (IStream* pStream, VARIANT_BOOL vbSave);
	HRESULT ExchangeConfig(IStream* pStream, VARIANT_BOOL vbSave);
	HRESULT DragDropExchange(IStream* pStream, VARIANT_BOOL vbSave);
	HRESULT ItemByIdentity(DWORD Identity, CTool **ppTool);

	int GetToolCount();
	CTool* GetTool(int nToolIndex);

	int GetVisibleToolCount();
	CTool* GetVisibleTool(int nToolIndex);
	int GetVisibleToolIndex(const CTool* pTool);
	int FindVisibleTools(BOOL bCommitted = TRUE);

	CTool* CreateMoreTools();
	CTool* FindToolId(ULONG nId);
	void ParentWindowedTools(HWND hWndParent);
	void HideWindowedTools();
	void TabbedWindowedTools(BOOL bHide);
	void ShowTabbedWindowedTools();
	int GetToolIndex(const CTool* pTool);
	int GetVisibleTools(TypedArray<CTool*>& aTools);
	BOOL CommitVisibleTools(TypedArray<CTool*>& aTools);
	void CleanupVisible();

	CTool* m_pMDITools[2];
	CTool* m_pMoreTools;
private:
	void DeleteVisibleTool(ITool *pTool);

	TypedArray<CTool*> m_aTools;
	TypedArray<CTool*> m_aVisibleTools;
	CBand*			   m_pBand;
	CBar*			   m_pBar;
};

inline int CTools::GetToolCount()
{
	return m_aTools.GetSize();
}

inline CTool* CTools::GetTool(int nIndex)
{
	return m_aTools.GetAt(nIndex);
}

inline int CTools::GetVisibleToolCount()
{
	return m_aVisibleTools.GetSize();
}

inline CTool* CTools::GetVisibleTool(int nToolIndex)
{
	if (nToolIndex < 0 && nToolIndex > m_aVisibleTools.GetSize() - 1)
		return NULL;
	return m_aVisibleTools.GetAt(nToolIndex);
}

extern ITypeInfo *GetObjectTypeInfo(LCID lcid,void *objectDef);

#endif

