#ifndef __LOCALIZER_H__
#define __LOCALIZER_H__

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

class CLocalizer
{
public:
	CLocalizer();

	ULONG AddRef();
	ULONG Release();

	void Add(int nCount, const long* pnResourceId);
	void Localize(int nIndex, BSTR bstrNew);
	LPCTSTR GetString(int nIndex);
	int GetCount();

protected:
	virtual ~CLocalizer();
	void Cleanup();

	ULONG m_nRefCount;
	BSTR* m_pStrings;
	const long* m_pnResourceIds;
	int   m_nStringCount;
};

inline int CLocalizer::GetCount()
{
	return m_nStringCount;
};

#endif