//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "flicker.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

CFlickerFree::~CFlickerFree()
{
	if (NULL != m_hBitmap)
	{
		SelectBitmap(m_hMemDC, m_hBitmapOld);
		DeleteBitmap(m_hBitmap);
	}
	if (NULL != m_hMemDC)
		DeleteDC(m_hMemDC);
}

HDC CFlickerFree::RequestDC(HDC hDCComp, int nWidth, int nHeight)
{
	if (NULL == m_hMemDC)
		m_hMemDC = CreateCompatibleDC(hDCComp);

	if (NULL == m_hMemDC)
		return NULL;

	if (nWidth > m_nCurWidth || nHeight > m_nCurHeight)
	{
		// old bitmap is not large enough create new
		if (NULL != m_hBitmap)
		{
			SelectBitmap(m_hMemDC, m_hBitmapOld);
			DeleteBitmap(m_hBitmap);
		}
		m_hBitmap = CreateCompatibleBitmap(hDCComp,nWidth,nHeight);
		if (NULL == m_hBitmap)
			return NULL;

		m_nCurWidth = nWidth;
		m_nCurHeight = nHeight;
		m_hBitmapOld = SelectBitmap(m_hMemDC, m_hBitmap);
		PatBlt(m_hMemDC, 0, 0, nWidth, nHeight, BLACKNESS);
	}
	m_nRequestWidth = nWidth;
	m_nRequestHeight = nHeight;
	return m_hMemDC;
}

