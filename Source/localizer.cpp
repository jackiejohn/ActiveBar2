//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "localizer.h"
#include "Support.h"
#include "Utility.h"
#include <assert.h>

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

CLocalizer::CLocalizer()
{
	m_nStringCount = 0;
	m_nRefCount = 1;
}

CLocalizer::~CLocalizer()
{
	Cleanup();
}

ULONG CLocalizer::AddRef()
{
	return ++m_nRefCount;
}

ULONG CLocalizer::Release()
{
	if (0 != --m_nRefCount)
		return m_nRefCount;
	delete this;
	return 0;
}

void CLocalizer::Cleanup()
{
	if (m_nStringCount > 0)
	{
		for (int nIndex = 0; nIndex < m_nStringCount; nIndex++)
			SysFreeString(m_pStrings[nIndex]);
		delete [] m_pStrings;
	}
	m_nStringCount = 0;
}

void CLocalizer::Add(int nStringCount, const long* pnResourceId)
{
	Cleanup();
	if (0 != nStringCount)
	{
		m_pStrings = new BSTR[nStringCount];
		if (m_pStrings)
			memset(m_pStrings, 0, sizeof(m_pStrings) * nStringCount);
		m_nStringCount = nStringCount;
		m_pnResourceIds = pnResourceId;
	}
}

LPCTSTR CLocalizer::GetString(int nIndex)
{
	static TCHAR szTemp[MAX_PATH];
	assert(nIndex >= 0 && nIndex < m_nStringCount);
	if (nIndex >= 0 && nIndex < m_nStringCount)
	{
		if (NULL == m_pStrings[nIndex])
		{
			if (m_pnResourceIds[nIndex] < 0)
				return NULL;
			DDString str;
			str.LoadString(m_pnResourceIds[nIndex]);
			m_pStrings[nIndex] = str.AllocSysString();
		}
		MAKE_TCHARPTR_FROMWIDE(szString, m_pStrings[nIndex]);
		lstrcpy(szTemp, szString);
		return szTemp;
	}
	else
		return NULL;
}

void CLocalizer::Localize(int nIndex, BSTR bstrNew)
{
	assert(nIndex >= 0 && nIndex < m_nStringCount);
	if (nIndex >= 0 && nIndex < m_nStringCount)
	{
		SysFreeString(m_pStrings[nIndex]);
		m_pStrings[nIndex] = SysAllocString(bstrNew);
	}
}