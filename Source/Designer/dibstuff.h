#ifndef __DIBSTUFF_H__
#define __DIBSTUFF_H__

#ifndef HDIB
typedef HANDLE HDIB;
#endif

extern WORD WINAPI PaletteSize(LPBYTE lpbi);
extern BOOL WINAPI StSaveDIB(IStream *pStream,HDIB hDib);
extern HRESULT StReadDIB(IStream *pStream,DWORD dataSize,HDIB *pdib);
extern HPALETTE FAR GetSystemPalette(void);
extern HDIB FAR BitmapToDIB(HBITMAP hBitmap, HPALETTE hPal,int pixperpixel);
extern HBITMAP FAR DIBToBitmap(HDIB hDIB, HPALETTE hPal,BOOL monochrome);
extern HBITMAP FAR DIBToGrayBitmap(HDIB hDIB);


#endif