#ifndef __FLICKERFREE_H__
#define __FLICKERFREE_H__

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

//
// CFlickerFree
//

class CFlickerFree
{
public:
	CFlickerFree();
	~CFlickerFree();

	int& Height();
	int& Width();

	HDC RequestDC(HDC hDCComp, int nWidth, int nHeight);
	void Paint(HDC hDCDest, int x, int y);
	void Paint(HDC hDCDest, int x, int y, int nWidth, int nHeight);
	void Paint(HDC hDCDest, int x, int y, int nWidth, int nHeight, int xs, int ys);

private:
	HBITMAP m_hBitmapOld;
	HBITMAP m_hBitmap;
	HDC		m_hMemDC;
	int		m_nRequestHeight;
	int		m_nRequestWidth;
	int		m_nCurHeight;
	int		m_nCurWidth;
};

inline CFlickerFree::CFlickerFree()
 : m_hMemDC(NULL),
   m_hBitmap(NULL),
   m_hBitmapOld(NULL)
{
	m_nRequestHeight = m_nRequestWidth = m_nCurWidth = m_nCurHeight = 0;
}

inline int& CFlickerFree::Height()
{
	return m_nCurHeight;
}

inline int& CFlickerFree::Width()
{
	return m_nCurWidth;
}

inline void CFlickerFree::Paint(HDC hDCDest,int x, int y, int nWidth, int nHeight, int xs, int ys)
{
	BitBlt(hDCDest,x,y,nWidth,nHeight,m_hMemDC,xs,ys,SRCCOPY);
}

inline void CFlickerFree::Paint(HDC hDCDest,int x,int y)
{
	BitBlt(hDCDest,x,y,m_nRequestWidth,m_nRequestHeight,m_hMemDC,0,0,SRCCOPY);
}

inline void CFlickerFree::Paint(HDC hDCDest,int x,int y,int w,int h)
{
	StretchBlt(hDCDest,x,y,w,h,m_hMemDC,0,0,m_nRequestWidth,m_nRequestHeight,SRCCOPY);
}

#endif