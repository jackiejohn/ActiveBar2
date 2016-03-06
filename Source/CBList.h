#ifndef __CBList_H__
#define __CBList_H__
#include "interfaces.h"

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

struct ListItem;
class CTool;

//
// CBList
//

class CBList : public IComboList
{
public:
	CBList();
	ULONG m_refCount;
	int m_objectIndex;
	~CBList();
	
	//{BEGIN INTERFACEDEFS}
	static void *objectDef;
	static CBList *CreateInstance(IUnknown *pUnkOuter);
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
	
	

	// IComboList members

	STDMETHOD(AddItem)( BSTR name);
	STDMETHOD(Remove)( VARIANT *Index);
	STDMETHOD(Item)( VARIANT *Index, BSTR *retval);
	STDMETHOD(Clear)();
	STDMETHOD(Count)( short *retval);
	STDMETHOD(InsertItem)( int index, BSTR name);
	STDMETHOD(put_ItemData)( VARIANT Index,  long Data);
	STDMETHOD(get_ItemData)( VARIANT Index,  long*Data);
	STDMETHOD(CopyTo)( IComboList **pComboList);
	STDMETHOD(get_NewIndex)(long *retval);
	// DEFS for IComboList
	
	
	// Events 
	//{END INTERFACEDEFS}
	int GetPosOfItem(BSTR strItem);
	int ConvertVariantToIndex(VARIANT *Index);
	int Count();
	void Cleanup();
	BSTR GetName(int nIndex);
	DWORD GetData(int nIndex);
	HRESULT Exchange(IStream* pStream, VARIANT_BOOL vbSave);
	void SetTool(CTool* pTool);
#pragma pack(push)
#pragma pack (1)
	struct CBListPropV1
	{
		CBListPropV1();

		ComboStyles m_nStyle;
		short m_nListIndex;
		short m_nLines;
		short m_nWidth;
	} lpV1;
	BOOL m_bNeedsSort:1;
	ULONG m_nNewIndex;
#pragma pack(pop)

	HWND m_hWndActive;
private:
	TypedArray<ListItem*> m_aItems;
	CTool* m_pTool;
};

inline void CBList::SetTool(CTool* pTool)
{
	m_pTool = pTool;
}

extern ITypeInfo *GetObjectTypeInfo(LCID lcid,void *objectDef);

#endif
