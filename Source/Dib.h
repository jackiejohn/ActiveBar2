#ifndef DIB_INCLUDED
#define DIB_INCLUDED

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

//
// CDib
//

class CDib
{
	enum Alloc 
	{
		eNoAlloc, 
		eCreateAlloc, 
		eHeapAlloc
	};

public:
	CDib();
	// Builds BITMAPINFOHEADER
	CDib(const SIZE& size, const int& nBitCount);	

	~CDib();

	void Empty();

	BOOL IsEmpty();

	int GetSizeImage();

	int GetSizeHeader();

	int GetBitmapSize();

	int NumOfColors();

	LPRGBQUAD GetColors();

	SIZE GetDimensions();

	BOOL AttachMemory(LPVOID pvMem, BOOL bMustDelete = FALSE, HGLOBAL hGlobal = NULL);

	HGLOBAL SaveToMemory();

	BOOL Draw(HDC hDC, const POINT& ptOrigin, const SIZE& size);

	HBITMAP CreateSection(HDC hDC = NULL);

	UINT UsePalette(HDC hDC, BOOL bBackground = FALSE);

	BOOL MakePalette();
	
	HBITMAP MakeGray();

	BOOL SetSystemPalette(HDC hDC);
	
	BOOL Compress(HDC hDC, BOOL bCompress = TRUE); // FALSE means decompress
	
	HBITMAP CreateBitmap(HDC hDC);
	HBITMAP CreateMonochromeBitmap(HDC hDC);
	BOOL FromBitmap (HBITMAP hBitmap, HDC hDC, int nBitCount = 0);
	
	BOOL Read(HANDLE hFile);

	BOOL Write(HANDLE hFile);
	
	HRESULT Read(IStream* pStream);

	HRESULT Write(IStream* pStream);

	BOOL ReadSection(HANDLE hFile, HDC hDC = NULL);
	
	LPBITMAPINFOHEADER BitmapInfoHeader();

	LPVOID ImageBits();
private:
	void ComputePaletteSize(const int& nBitCount);
	void ComputeMetrics();

	HGLOBAL m_hGlobal; // For external windows we need to free;
	                   //  could be allocated by this class or allocated externally
	Alloc m_nBmihAlloc;
	Alloc m_nImageAlloc;
	DWORD m_dwSizeImage; // of bits -- not BITMAPINFOHEADER or BITMAPFILEHEADER
	int   m_nColorTableEntries;
	
	LPBITMAPINFOHEADER m_pBMIH; //  buffer containing the BITMAPINFOHEADER
	HPALETTE m_hPalette;
	HBITMAP  m_hBitmap;
	LPVOID   m_pvColorTable;
	LPVOID   m_pvFile;
	LPBYTE   m_pImage;
};

inline int CDib::NumOfColors()
{
	return m_nColorTableEntries;
}

inline LPRGBQUAD CDib::GetColors()
{
	return (LPRGBQUAD)m_pvColorTable;
}

inline int CDib::GetSizeHeader()
{
	return sizeof(BITMAPINFOHEADER) + sizeof(RGBQUAD) * m_nColorTableEntries;
}

inline int CDib::GetSizeImage() 
{
	return m_dwSizeImage;
}

inline BOOL CDib::IsEmpty()
{
	return NULL == m_pImage ? TRUE : FALSE;
}

inline LPBITMAPINFOHEADER CDib::BitmapInfoHeader()
{
	return m_pBMIH;
}

inline LPVOID CDib::ImageBits()
{
	return m_pImage;
}

//
// TreeNode
//

struct TreeNode 
{
    BOOL bIsLeaf;           // TRUE if node has no children
    UINT nPixelCount;       // Number of pixels represented by this leaf
    UINT nRedSum;           // Sum of red components
    UINT nGreenSum;         // Sum of green components
    UINT nBlueSum;          // Sum of blue components
    TreeNode* pChild[8];    // Pointers to child nodes
    TreeNode* pNext;        // Pointer to next reducible node
};

//
// CColorQuantizer
//

class CColorQuantizer
{
public:
	CColorQuantizer(UINT nMaxColors, UINT nColorBits);
	~CColorQuantizer();

	BOOL ProcessImage (CDib& theImage);
	HPALETTE CreatePalette();

	void AddColor (BYTE bRed, BYTE bGreen, BYTE bBlue);

private:
	void AddColor (TreeNode*& pNode, BYTE r, BYTE g, BYTE b, UINT nColorBits, UINT nLevel, UINT& nLeafCount, TreeNode** pReducibleNodes);
	TreeNode* CreateNode (UINT nLevel, UINT nColorBits, UINT& NLeafCount, TreeNode** pReducibleNodes);
	void ReduceTree (UINT nColorBits, UINT& nLeafCount, TreeNode** pReducibleNodes);
	void DeleteTree (TreeNode*& ppNode);
	void GetPaletteColors (TreeNode* pTree, PALETTEENTRY* pPalEntries, UINT& nIndex);
	int GetRightShiftCount (DWORD dwVal);
	int GetLeftShiftCount (DWORD dwVal);

	UINT m_nMaxColors;
	UINT m_nColorBits;
	UINT m_nLeafCount;
    TreeNode* m_pReducibleNodes[9];
    TreeNode* m_pTree;
};

#endif