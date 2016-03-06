#ifndef __SHORTCUTS_H__
#define __SHORTCUTS_H__

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

struct Token
{
	enum TokenTypes
	{
		eAlt,
		eControl,
		eShift,
		eChar,
		eComma
	};

	Token(TokenTypes eType, TCHAR cCharacter = NULL)
		: m_eType(eType),
		  m_cCharacter(cCharacter)
	{
	}

	TokenTypes m_eType;
	TCHAR m_cCharacter;
};

class ShortCutStore : public IShortCut,public ISupportErrorInfo
{
public:
	ShortCutStore();
	~ShortCutStore();

	ULONG m_refCount;
	//{BEGIN INTERFACEDEFS}
	static void *objectDef;
	static ShortCutStore *CreateInstance(IUnknown *pUnkOuter);
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
	
	

	// IShortCut members

	STDMETHOD(get_Value)(BSTR *retval);
	STDMETHOD(put_Value)(BSTR val);
	STDMETHOD(Set)( short nIndex,  long nKeyCode,  long nKeyboardCode,  VARIANT_BOOL bShift);
	STDMETHOD(SetControlCode)( long nControlCode);
	STDMETHOD(Clone)( IShortCut **pShortCut);
	STDMETHOD(Clear)();
	STDMETHOD(GetKeyCode)( short nIndex,  long *retval);
	STDMETHOD(IsEqual)( ShortCut*aShortCut);
	// DEFS for IShortCut
	
	

	// ISupportErrorInfo members

	STDMETHOD(InterfaceSupportsErrorInfo)( REFIID riid);
	// DEFS for ISupportErrorInfo
	
	
	// Events 
	//{END INTERFACEDEFS}
	
	int operator==(ShortCutStore& rhs) const;
	int operator!=(ShortCutStore& rhs) const;
	BOOL GetShortCutDesc (LPTSTR szShortCut);
	HRESULT Exchange(IStream* pStream, VARIANT_BOOL vbSave);
	void SendError(DWORD dwHelpContextId, DWORD dwDescriptionId);

#pragma pack(push)
#pragma pack (1)
	struct ShortCutPropV1
	{
		ShortCutPropV1();

		TCHAR m_nKeyCodes[2];
		BOOL  m_bShift[2];
		long  m_nControlKeys;
		long  m_nKeyboardCodes[2];
	} scpV1;
#pragma pack(pop)

	enum ParseStates
	{ 
		eStart,
		eAlt,
		eControl,
		eShift,
		eChar,
		eComma,
		eShift2,
		eChar2,
		eError,
		eEnd
	};

	static BOOL Parse(TCHAR* szShortCut, TypedArray<Token*>& aTokens);
	static BOOL Syntax(const TypedArray<Token*>& aTokens, ShortCutPropV1& scpShortCut);

	CBar* m_pBar;
};


//
// operator==
//

inline int ShortCutStore::operator==(ShortCutStore& rhs) const 
{
	if (0 != memcmp(&scpV1.m_nKeyCodes, &rhs.scpV1.m_nKeyCodes, sizeof(scpV1.m_nKeyCodes)))
		return FALSE;
	if (0 != memcmp(&scpV1.m_bShift, &rhs.scpV1.m_bShift, sizeof(scpV1.m_bShift)))
		return FALSE;
	if (0 != memcmp(&scpV1.m_nControlKeys, &rhs.scpV1.m_nControlKeys, sizeof(scpV1.m_nControlKeys)))
		return FALSE;
	return TRUE;
}

//
// operator==
//

inline int ShortCutStore::operator!=(ShortCutStore& rhs) const 
{
	if (0 != memcmp(scpV1.m_nKeyCodes, &rhs.scpV1.m_nKeyCodes, sizeof(scpV1.m_nKeyCodes)))
		return TRUE;
	if (0 != memcmp(scpV1.m_bShift, &rhs.scpV1.m_bShift, sizeof(scpV1.m_bShift)))
		return TRUE;
	if (0 != memcmp(&scpV1.m_nControlKeys, &rhs.scpV1.m_nControlKeys, sizeof(scpV1.m_nControlKeys)))
		return TRUE;
	return FALSE;
}

//
// CShortCut
//

class CShortCut
{
public:
	CShortCut();
	~CShortCut();

	HRESULT AddTool (CTool* pTool);
	HRESULT GetTool(CTool*& pTool);
	HRESULT RemoveTool(CTool* pTool);

	int ToolCount();
	void SetShortCut(ShortCutStore* pShortCutStore);
	ShortCutStore* GetShortCutStore();

private:
	TypedArray<CTool*> m_aTools;
	ShortCutStore* m_pShortCutStore;
	long  m_nLastToolId;
};

inline int CShortCut::ToolCount()
{
	return m_aTools.GetSize();
}

inline ShortCutStore* CShortCut::GetShortCutStore()
{
	return m_pShortCutStore;
}

inline void CShortCut::SetShortCut(ShortCutStore* pShortCutStore)
{
	if (m_pShortCutStore)
		m_pShortCutStore->Release();
	pShortCutStore->Clone((IShortCut**)&m_pShortCutStore);
}

//
// CShortCuts
//

class CShortCuts
{
public:
	CShortCuts();
	~CShortCuts();

	HRESULT Count (short* retval);
	HRESULT Add (ShortCutStore* pShortCutStore, CTool* pTool);
	HRESULT Remove (const ShortCutStore* pShortCutStore, CTool* pTool);
	HRESULT Find (const ShortCutStore* pShortCutStore, CShortCut*& pShortCut);

	void SetOwner(CBar* pBar);
	void CleanUp();
	BOOL Process(long nControlState, WPARAM wKey, LPARAM lParam);
	BOOL TranslateAccelerator(MSG* pMsg);
	ShortCutStore* GetShortCutStore();

private:
	void FireDelayedToolClick(CTool* pTool);
	BOOL ValidFirstKey(BOOL bShift, long nControlState, WPARAM wKey, LPARAM lParam);
	TypedArray<CShortCut*> m_aShortCuts;
	ShortCutStore* m_pShortCutStore;
	CBar* m_pBar;
};

inline ShortCutStore* CShortCuts::GetShortCutStore()
{
	return m_pShortCutStore;
}

#endif
