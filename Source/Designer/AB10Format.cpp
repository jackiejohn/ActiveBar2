//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "precomp.h"
#include "Support.h"
#include "..\EventLog.h"
#include "ImageMgr.h"
#include "Utility.h"
#include "AB10Format.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;

static void DebugPrintBitmap(HBITMAP hBitmap, int nOffset = 0)
{
	HDC hDC = GetDC(NULL);
	if (NULL == hDC)
		return;

	HDC hDCSrc = CreateCompatibleDC(hDC);
	HBITMAP hBitmapOld = SelectBitmap(hDCSrc, hBitmap);
	BITMAP bmInfo;
	GetObject(hBitmap, sizeof(bmInfo), &bmInfo);
	BitBlt(hDC, 
		   0, 
		   nOffset, 
		   bmInfo.bmWidth, 
		   bmInfo.bmHeight, 
		   hDCSrc, 
		   0, 
		   0, 
		   SRCCOPY);
	SelectBitmap(hDCSrc, hBitmapOld);
	ReleaseDC(NULL, hDC);
}

#endif

//
// AB10Tool
//

AB10Tool::AB10Tool(AB10* pBar)
	:	m_bstrDescription(NULL),
		m_bstrToolTipText(NULL),
		m_bstrCategory(NULL),
		m_bstrSubBand(NULL),
		m_bstrCaption(NULL),
		m_bstrText(NULL),
		m_bstrName(NULL),
		m_pBar(pBar)
{
}

AB10Tool::~AB10Tool()
{
	SysFreeString(m_bstrDescription);
	SysFreeString(m_bstrToolTipText);
	SysFreeString(m_bstrCategory);
	SysFreeString(m_bstrSubBand);
	SysFreeString(m_bstrCaption);
	SysFreeString(m_bstrText);
	SysFreeString(m_bstrName);
}

#define SHORTCUTSTART 2000

static TCHAR* s_szShortCuts[] = {"",
								 "Control+A",
								 "Control+B",
								 "Control+C",
								 "Control+D",
								 "Control+E",
								 "Control+F",
								 "Control+G",
								 "Control+H",
								 "Control+I",
								 "Control+J",
								 "Control+K",
								 "Control+L",
								 "Control+M",
								 "Control+N",
								 "Control+O",
								 "Control+P",
								 "Control+Q",
								 "Control+R",
								 "Control+S",
							 	 "Control+T",
								 "Control+U",
								 "Control+V",
								 "Control+W",
								 "Control+X",
								 "Control+Y",
								 "Control+Z",
								 "F1",
 								 "F2",
 								 "F3",
								 "F4",
								 "F5",
								 "F6",
								 "F7",
								 "F8",
								 "F9",
								 "F11",
								 "F12",
								 "Control+F1",
								 "Control+F2",
								 "Control+F3",
								 "Control+F4",
								 "Control+F5",
								 "Control+F6",
								 "Control+F7",
								 "Control+F8",
								 "Control+F9",
								 "Control+F11",
								 "Control+F12",
								 "Shift+F1",
								 "Shift+F2",
								 "Shift+F3",
								 "Shift+F4",
								 "Shift+F5",
								 "Shift+F6",
								 "Shift+F7",
								 "Shift+F8",
								 "Shift+F9",
								 "Shift+F11",
								 "Shift+F12",
								 "Shift+Control+F1",
								 "Shift+Control+F2",
								 "Shift+Control+F3",
								 "Shift+Control+F4",
								 "Shift+Control+F5",
								 "Shift+Control+F6",
								 "Shift+Control+F7",
								 "Shift+Control+F8",
								 "Shift+Control+F9",
								 "Shift+Control+F11",
								 "Shift+Control+F12",
								 "Control+Insert",
								 "Shift+Insert",
								 "Delete",
								 "Shift+Delete",
								 "Alt+Backspace"};

BOOL AB10Tool::Read(IStream* pStream)
{
	try
	{
		HRESULT hResult;
		BOOL bTemp;
		VARIANT_BOOL vbImages = VARIANT_FALSE;
		short nVersion;
		short nMarker;
		
		hResult = pStream->Read(&nVersion,sizeof(short),NULL);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = pStream->Read(&m_nToolId,sizeof(long),NULL);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = StReadBSTR(pStream, m_bstrName);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = pStream->Read(&m_nHelpContextId,sizeof(long),NULL);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = StReadBSTR(pStream,m_bstrToolTipText);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = StReadBSTR(pStream,m_bstrCategory);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = pStream->Read(&m_bEnabled,sizeof(BOOL),NULL);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = pStream->Read(&m_bChecked,sizeof(BOOL),NULL);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = StReadBSTR(pStream,m_bstrCaption);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = pStream->Read(&m_tsTool,sizeof(ToolStyles),NULL);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = pStream->Read(&m_cpTool,sizeof(CaptionPositionTypes),NULL);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = StReadBSTR(pStream,m_bstrDescription);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = pStream->Read(&m_ttTool,sizeof(ToolTypes),NULL);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = pStream->Read(&bTemp,sizeof(BOOL),NULL);
		if (FAILED(hResult))
			return FALSE;
		
		m_vbBeginGroup = -bTemp;
		
		hResult = pStream->Read(&m_nTag,sizeof(long),NULL);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = pStream->Read(&m_nWidth,sizeof(int),NULL);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = pStream->Read(&m_nHeight,sizeof(int),NULL);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = pStream->Read(&m_taTool,sizeof(ToolAlignmentTypes),NULL); 
		if (FAILED(hResult))
			return FALSE;
		
		hResult = pStream->Read(&m_dwShortCutKey,sizeof(DWORD),NULL);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = pStream->Read(&m_dwImageCookie,sizeof(DWORD),NULL);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = StReadBSTR(pStream, m_bstrSubBand);
		if (FAILED(hResult))
			return FALSE;

		hResult = pStream->Read(&nMarker,sizeof(short),NULL);
		if (FAILED(hResult))
			return FALSE;

		if (nMarker==1)
		{
			hResult = pStream->Read(&m_vbVisible,sizeof(VARIANT_BOOL),NULL);
			if (FAILED(hResult))
				return FALSE;
			
			hResult = pStream->Read(&nMarker,sizeof(short),NULL);
			if (FAILED(hResult))
				return FALSE;
		}

		if (nMarker==2)
		{
			hResult = pStream->Read(&(m_nImageIds[0]),sizeof(DWORD)*4,0);
			if (FAILED(hResult))
				return FALSE;

			hResult = pStream->Read(&nMarker,sizeof(short),NULL);
			if (FAILED(hResult))
				return FALSE;
		}

		if (nMarker==3)
		{
			if (VARIANT_TRUE == vbImages)
			{
				//
				// Loading Images
				//

				HPALETTE hPal = NULL;
				HRESULT hResult;
				HBITMAP hBitmap;
				BITMAP  bmInfo;
				HDC hDC = GetDC(NULL);
				if (hDC)
				{
					for (int nImage = 0; nImage < 4; nImage++)
					{
						if (-1 == m_nImageIds[nImage])
						{
							m_nImageIds[nImage] = -1;
							StReadHBITMAP(pStream, &hBitmap, 0, FALSE, hPal);
							if (NULL == hBitmap)
								return E_FAIL;

							GetObject(hBitmap, sizeof(bmInfo), &bmInfo);
							hResult = m_pBar->m_pImageMgr->CreateImage2(&m_nImageIds[nImage], bmInfo.bmWidth, bmInfo.bmHeight); 
							if (SUCCEEDED(hResult))
							{
								m_pBar->m_pImageMgr->put_ImageBitmap(m_nImageIds[nImage], (OLE_HANDLE)hBitmap);
								DeleteBitmap(hBitmap);
								StReadHBITMAP(pStream, &hBitmap, 0, FALSE, hPal);
								m_pBar->m_pImageMgr->put_MaskBitmap(m_nImageIds[nImage], (OLE_HANDLE)hBitmap);
								DeleteBitmap(hBitmap);
							}
						}
					}
					ReleaseDC(NULL, hDC);
				}
				hResult = pStream->Read(&nMarker,sizeof(short),NULL);
			}
		}
		return TRUE;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
}

//
// Convert
//

BOOL AB10Tool::Convert(ITool* pTool)
{
	try
	{
		BOOL bResult = TRUE;
		HRESULT hResult;

		hResult = pTool->put_Description(m_bstrDescription);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = pTool->put_TooltipText(m_bstrToolTipText);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = pTool->put_Category(m_bstrCategory);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = pTool->put_SubBand(m_bstrSubBand);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = pTool->put_Caption(m_bstrCaption);
		if (FAILED(hResult))
			return FALSE;
		
		if (m_bstrText && m_bstrText)
		{
			hResult = pTool->put_Text(m_bstrText);
			if (FAILED(hResult))
				return FALSE;
		}
		
		hResult = pTool->put_CaptionPosition(m_cpTool);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = pTool->put_Alignment(m_taTool);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = pTool->put_ControlType(m_ttTool);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = pTool->put_Style(m_tsTool);
		if (FAILED(hResult))
			return FALSE;
		
		m_bEnabled = -m_bEnabled;
		hResult = pTool->put_Enabled(m_bEnabled);
		if (FAILED(hResult))
			return FALSE;
		
		m_bChecked = -m_bChecked;
		hResult = pTool->put_Checked(m_bChecked);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = pTool->put_HelpContextID(m_nHelpContextId);
		if (FAILED(hResult))
			return FALSE;
		
		SIZE size;
		if (m_nHeight > 0)
		{
			size.cx = m_nHeight;
			PixelToTwips(&size, &size);
			m_nHeight = size.cx;
		}
		hResult = pTool->put_Height(m_nHeight);
		if (FAILED(hResult))
			return FALSE;
		
		if (m_nWidth > 0)
		{
			size.cx = m_nWidth;
			PixelToTwips(&size, &size);
			m_nWidth = size.cx;
		}
		hResult = pTool->put_Width(m_nWidth);
		if (FAILED(hResult))
			return FALSE;
		
		VARIANT vTag;
		vTag.vt = VT_I4;
		vTag.lVal = m_nTag;
		hResult = pTool->put_TagVariant(vTag);
		if (FAILED(hResult))
			return FALSE;
		
		hResult = pTool->put_Visible(m_vbVisible);
		if (FAILED(hResult))
			return FALSE;

		//
		// ShortCuts
		//
		
		long nShortCutIndex = (long)m_dwShortCutKey;
		if (nShortCutIndex > 0 && nShortCutIndex < sizeof(s_szShortCuts) / sizeof(TCHAR*))
		{
			TCHAR* szShortCut = s_szShortCuts[nShortCutIndex];
			if (szShortCut)
			{
				VARIANT vShortCut;
				vShortCut.vt = VT_ARRAY|VT_DISPATCH;
				vShortCut.parray = NULL;

				SAFEARRAYBOUND rgsabound[1];    
				rgsabound[0].lLbound = 0;
				rgsabound[0].cElements = 1;
				vShortCut.parray = SafeArrayCreate(VT_DISPATCH, 1, rgsabound);
				if (vShortCut.parray)
				{
					IShortCut** ppShortCut = NULL;

					// Get a pointer to the elements of the array.
					HRESULT hResult = SafeArrayAccessData(vShortCut.parray, (void**)&ppShortCut);
					if (SUCCEEDED(hResult))
					{
						HRESULT hResult = CoCreateInstance(CLSID_ShortCut, 
														   NULL, 
														   CLSCTX_INPROC_SERVER, 
														   IID_IShortCut, 
														   (LPVOID*)&ppShortCut[0]);
						if (SUCCEEDED(hResult))
						{
							MAKE_WIDEPTR_FROMTCHAR(wShortCut, szShortCut);
							BSTR bstr = SysAllocString(wShortCut);
							if (bstr)
							{
								hResult = ppShortCut[0]->put_Value(bstr);
								SysFreeString(bstr);
							}
						}
						SafeArrayUnaccessData(vShortCut.parray);

						hResult = pTool->put_ShortCuts(vShortCut);

						ppShortCut[0]->Release();
					}
				}
			}
		}

		//
		// Images
		//

		OLE_HANDLE hBitmap = NULL;
		OLE_HANDLE hBitmapMask = NULL;
		for (int nImage = 0; nImage < eNumOfImageTypes; nImage++)
		{
			if (-1 == m_nImageIds[nImage])
				continue;

			hResult = m_pBar->m_pImageMgr->get_ImageBitmap(m_nImageIds[nImage], &hBitmap);
			if (FAILED(hResult))
				continue;

			hResult = pTool->put_Bitmap((ImageTypes)nImage, (OLE_HANDLE)hBitmap);
			if (hBitmap)
			{
				bResult = DeleteBitmap(hBitmap);
				assert(bResult);
				hBitmap = NULL;
			}

			hResult = m_pBar->m_pImageMgr->get_MaskBitmap(m_nImageIds[nImage], &hBitmapMask);
			if (FAILED(hResult))
			{
				bResult = DeleteBitmap(hBitmap);
				assert(bResult);
				bResult = DeleteBitmap(hBitmapMask);
				assert(bResult);
				continue;
			}

			hResult = pTool->put_MaskBitmap((ImageTypes)nImage, (OLE_HANDLE)hBitmapMask);
			if (hBitmapMask)
			{
				bResult = DeleteBitmap(hBitmapMask);
				assert(bResult);
				hBitmapMask = NULL;
			}
		}
		return bResult;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
}

//
// AB10Tools
//

AB10Tools::~AB10Tools()
{
	int nCount = m_faTools.GetSize();
	for (int nTool = 0; nTool < nCount; nTool++)
		delete m_faTools.GetAt(nTool);
}

//
// Read
//

BOOL AB10Tools::Read(IStream* pStream)
{
	try
	{
		BitmapMgr theBitmapMgr;
		m_pBar->m_pImageMgr->GetConvertInfo(pStream, FALSE, theBitmapMgr);

		DWORD nCount;
		HRESULT hResult = pStream->Read(&nCount, sizeof(DWORD), NULL);
		if (FAILED(hResult))
			return FALSE;

		for (int nTool = 0; (DWORD)nTool < nCount; nTool++)
		{
			AB10Tool* pTool = new AB10Tool(m_pBar);
			if (NULL == pTool)
				return E_FAIL;

			pTool->Read(pStream);
			if (pTool->m_nToolId > m_pBar->m_nHighestToolId)
				m_pBar->m_nHighestToolId = pTool->m_nToolId;

			m_faTools.Add(pTool);
		}
		m_pBar->m_pImageMgr->Convert(theBitmapMgr, m_faTools);
		return TRUE;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
}

//
// Convert
//

BOOL AB10Tools::Convert(ITools* pTools)
{
	try
	{
		BOOL bResult = TRUE;
		HRESULT hResult;
		ITool* pTool;
		AB10Tool* pTool10;
		int nCount = m_faTools.GetSize();
		for (int nTool = 0; nTool < nCount; nTool++)
		{
			pTool10 = m_faTools.GetAt(nTool);
			if (NULL == pTool10)
				continue;

			if (VARIANT_TRUE == pTool10->m_vbBeginGroup)
			{
				hResult = pTools->Add(++m_pBar->m_nHighestToolId, L"miSeparator", (Tool**)&pTool);
				if (SUCCEEDED(hResult))
				{
					pTool->put_ControlType(ddTTSeparator);
					pTool->Release();
				}
			}
			WCHAR* wName = NULL;
			if (NULL == pTool10->m_bstrName || NULL == *pTool10->m_bstrName)
			{
				WCHAR wCaption[256];
				CleanCaption(pTool10->m_bstrCaption, wCaption);
				wName = wCaption;
			}
			else
				wName = pTool10->m_bstrName;

			hResult = pTools->Add(pTool10->m_nToolId, wName, (Tool**)&pTool);
			if (FAILED(hResult))
			{
				hResult = pTools->Add(++m_pBar->m_nHighestToolId, wName, (Tool**)&pTool);
				if (FAILED(hResult))
					continue;
			}

			bResult = pTool10->Convert(pTool);

			pTool->Release();
			pTool = NULL;
			if (!bResult)
				return FALSE;
		}
		return bResult;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return E_FAIL;
	}
}

//
// AB10Page
//

AB10ChildBand::AB10ChildBand(AB10* pBar)
	: m_pBar(pBar),
	  m_bstrCaption(NULL)
{
	m_pTools = new AB10Tools(m_pBar);
	assert(m_pTools);
}

AB10ChildBand::~AB10ChildBand()
{
	delete m_pTools;
	SysFreeString(m_bstrCaption);
}

BOOL AB10ChildBand::Read(IStream* pStream)
{
	try
	{
		HRESULT hResult = StReadBSTR(pStream, m_bstrCaption);
		if (FAILED(hResult))
			return FALSE;

		hResult = m_pTools->Read(pStream);
		if (FAILED(hResult))
			return FALSE;
		return TRUE;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
}

//
// Convert
//

BOOL AB10ChildBand::Convert(IBand* pChildBand)
{
	try
	{
		BOOL bResult = TRUE;
		HRESULT hResult;
		ITools* pTools = NULL;

		hResult = pChildBand->get_Tools((Tools**)&pTools);
		if (FAILED(hResult))
		{
			bResult = FALSE;
			goto Cleanup;
		}
		m_pTools->Convert(pTools);

Cleanup:
		if (pTools)
			pTools->Release();
		return bResult;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
}

//
// AB10ChildBands
//

AB10ChildBands::~AB10ChildBands()
{
	int nCount = m_faChildBands.GetSize();
	for (int nBand = 0; nBand < nCount; nBand++)
		delete m_faChildBands.GetAt(nBand);
}

//
// Read
//

BOOL AB10ChildBands::Read(IStream* pStream)
{
	try
	{
		HRESULT hResult;
		short nVersion = 0;
		hResult = pStream->Read(&nVersion,sizeof(short),NULL);
		if (FAILED(hResult))
			return FALSE;

		short nPageCount;
		hResult = pStream->Read(&nPageCount,sizeof(short),NULL);
		if (FAILED(hResult))
			return FALSE;

		if (nPageCount > 0)
		{
			for (int nPage = 0; nPage < nPageCount; nPage++)
			{
				AB10ChildBand* pNewChildBand = new AB10ChildBand(m_pBar);
				if (!pNewChildBand)
					return E_OUTOFMEMORY;
				pNewChildBand->Read(pStream);
				m_faChildBands.Add(pNewChildBand);
			}
		}

		// End marker
		DWORD nId=0;
		hResult = pStream->Read(&nId, sizeof(DWORD), NULL);
		if (FAILED(hResult))
			return FALSE;
		return TRUE;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
}

//
// Convert
//

BOOL AB10ChildBands::Convert(IChildBands* pChildBands)
{
	try
	{
		BOOL bResult = TRUE;
		HRESULT hResult;
		AB10ChildBand* pChildBand10 = NULL;
		IBand* pChildBand = NULL;

		hResult = pChildBands->RemoveAll();
		if (FAILED(hResult))
			return FALSE;

		int nCount = m_faChildBands.GetSize();
		for (int nBand = 0; nBand < nCount; nBand++)
		{
			pChildBand10 = m_faChildBands.GetAt(nBand);
			if (NULL == pChildBand10)
				continue;

			hResult = pChildBands->Add(pChildBand10->m_bstrCaption, (Band**)&pChildBand);
			if (FAILED(hResult))
				continue;

			bResult = pChildBand10->Convert(pChildBand);
			pChildBand->Release();
			pChildBand = NULL;
			if (!bResult)
				return bResult;
		}
		return bResult;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
}

//
// AB10Band
//

AB10Band::AB10Band(AB10* pBar)
	: m_bstrCaption(NULL),
	  m_bstrName(NULL),
	  m_pBar(pBar)
{
	m_pTools = new AB10Tools(pBar);
	assert(m_pTools);
	m_pChildBands = new AB10ChildBands(pBar);
	assert(m_pChildBands);
}

AB10Band::~AB10Band()
{
	SysFreeString(m_bstrCaption);
	SysFreeString(m_bstrName);
	delete m_pTools;
	delete m_pChildBands;
}

//
// Read
//

BOOL AB10Band::Read(IStream* pStream)
{
	try
	{
		HRESULT hResult = pStream->Read(&m_bIsDetached,sizeof(BOOL),NULL);
		if (FAILED(hResult))
			return FALSE;

		hResult = StReadBSTR(pStream, m_bstrCaption);
		if (FAILED(hResult))
			return FALSE;

		hResult = pStream->Read(&m_dwFlags,sizeof(DWORD),NULL);
		if (FAILED(hResult))
			return FALSE;

		hResult = pStream->Read(&m_daDockingArea,sizeof(DockingAreaTypes),NULL);
		if (FAILED(hResult))
			return FALSE;

		hResult = pStream->Read(&m_nDockOffset,sizeof(int),NULL); // non portable
		if (FAILED(hResult))
			return FALSE;

		hResult = pStream->Read(&m_nDockLine,sizeof(short),NULL);
		if (FAILED(hResult))
			return FALSE;

		hResult = pStream->Read(&m_nToolsHPadding,sizeof(int),NULL);
		if (FAILED(hResult))
			return FALSE;

		hResult = pStream->Read(&m_nToolsVPadding,sizeof(int),NULL);
		if (FAILED(hResult))
			return FALSE;

		hResult = pStream->Read(&m_nToolsHSpacing,sizeof(int),NULL);
		if (FAILED(hResult))
			return FALSE;

		hResult = pStream->Read(&m_nToolsVSpacing,sizeof(int),NULL);
		if (FAILED(hResult))
			return FALSE;

		hResult = pStream->Read(&m_tsMouseTracking,sizeof(TrackingStyles),NULL);
		if (FAILED(hResult))
			return FALSE;

		hResult = pStream->Read(&m_vbDisplayHandles,sizeof(VARIANT_BOOL),NULL);
		if (FAILED(hResult))
			return FALSE;

		hResult = PersistFont(pStream,&m_font,NULL,FALSE);
		if (FAILED(hResult))
			return FALSE;

		hResult = pStream->Read(&m_btType,sizeof(BandTypes),NULL);	
		if (FAILED(hResult))
			return FALSE;

		hResult = pStream->Read(&m_vbVisible,sizeof(VARIANT_BOOL),NULL);
		if (FAILED(hResult))
			return FALSE;

		hResult = pStream->Read(&m_vbWrappable,sizeof(VARIANT_BOOL),NULL);
		if (FAILED(hResult))
			return FALSE;

		hResult = pStream->Read(&m_rcFloat,sizeof(RECT),NULL);
		if (FAILED(hResult))
			return FALSE;

		hResult = pStream->Read(&m_rcDimension,sizeof(RECT),NULL);
		if (FAILED(hResult))
			return FALSE;

		// Now Exchange tools
		hResult = m_pTools->Read(pStream);
		DWORD dwId = 0;
		BOOL bNameLoaded = FALSE;
		while (TRUE)
		{
			hResult = pStream->Read(&dwId, sizeof(DWORD), NULL);
			if (FAILED(hResult))
				return FALSE;

			switch (dwId)
			{
			case 0xB0B0:
				hResult = m_pChildBands->Read(pStream);
				if (FAILED(hResult))
					return FALSE;
				break;

			case 0xB0B1:
				hResult = pStream->Read(&m_nGrabHandleStyle,sizeof(short),NULL);
				if (FAILED(hResult))
					return FALSE;
				break;

			case 0xB0B2:
				hResult = pStream->Read(&m_psStyle,sizeof(m_psStyle),NULL);
				if (FAILED(hResult))
					return FALSE;
				break;

			case 0xB0B3:
				hResult = pStream->Read(&m_nCreatedBy,sizeof(m_nCreatedBy),NULL);
				if (FAILED(hResult))
					return FALSE;
				break;

			case 0xB0B4:
				hResult = StReadBSTR(pStream, m_bstrName);
				if (FAILED(hResult))
					return FALSE;
				bNameLoaded = TRUE;
				break;

			default:
				dwId = 0;
				break;
			}
			if (0 == dwId)
				break;
		}
		if (!bNameLoaded)
			m_bstrName = SysAllocString(m_bstrCaption);
		return TRUE;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
}

int AB10Band::ConvertDockingArea()
{
	switch (m_daDockingArea)
	{
    case 0x1:
		return 0;
	case 0x2:
		return 1;
	case 0x4:
		return 2;
	case 0x8:
		return 3;
	case 0x10:
		return 4;
	case 0x20:
		return 5;
	}
	return 0;
}

//
// Convert
//

BOOL AB10Band::Convert(IBand* pBand)
{	
	try
	{
		BOOL bResult = TRUE;
		HRESULT hResult;
		IChildBands* pChildBands = NULL;
		ITools* pTools = NULL;

		hResult = pBand->put_Caption(m_bstrCaption);
		if (FAILED(hResult))
			return FALSE;

		SIZE size;
		if (m_nDockOffset > 0)
		{
			size.cx = m_nDockOffset;
			PixelToTwips(&size, &size);
			m_nDockOffset = size.cx;
		}
		hResult = pBand->put_DockingOffset(m_nDockOffset);
		if (FAILED(hResult))
			return FALSE;

		hResult = pBand->put_DockingArea((DockingAreaTypes)ConvertDockingArea());
		if (FAILED(hResult))
			return FALSE;

		hResult = pBand->put_DockLine(m_nDockLine);
		if (FAILED(hResult))
			return FALSE;

		hResult = pBand->put_Flags((BandFlags)m_dwFlags);
		if (FAILED(hResult))
			return FALSE;

		if (VARIANT_FALSE == m_vbDisplayHandles)
			hResult = pBand->put_GrabHandleStyle(ddGSNone);
		else
			hResult = pBand->put_GrabHandleStyle((GrabHandleStyles)(m_nGrabHandleStyle+1));
		if (FAILED(hResult))
			return FALSE;

		hResult = pBand->put_IsDetached(m_bIsDetached);
		if (FAILED(hResult))
			return FALSE;

		hResult = pBand->put_CreatedBy((CreatedByTypes)m_nCreatedBy);
		if (FAILED(hResult))
			return FALSE;

		hResult = pBand->put_MouseTracking(m_tsMouseTracking);
		if (FAILED(hResult))
			return FALSE;

		hResult = pBand->put_Visible(m_vbVisible);
		if (FAILED(hResult))
			return FALSE;

		hResult = pBand->put_Type(m_btType);
		if (FAILED(hResult))
			return FALSE;

		hResult = pBand->put_ChildBandStyle(m_psStyle);
		if (FAILED(hResult))
			return FALSE;

		hResult = pBand->put_WrapTools(m_vbWrappable);
		if (FAILED(hResult))
			return FALSE;

		if (m_nToolsHPadding > 0)
		{
			size.cx = m_nToolsHPadding;
			PixelToTwips(&size, &size);
			m_nToolsHPadding = size.cx;
		}
		hResult = pBand->put_ToolsHPadding(m_nToolsHPadding);
		if (FAILED(hResult))
			return FALSE;

		if (m_nToolsVPadding > 0)
		{
			size.cx = m_nToolsVPadding;
			PixelToTwips(&size, &size);
			m_nToolsVPadding = size.cx;
		}
		hResult = pBand->put_ToolsVPadding(m_nToolsVPadding);
		if (FAILED(hResult))
			return FALSE;

		if (m_nToolsHSpacing > 0)
		{
			size.cx = m_nToolsHSpacing;
			PixelToTwips(&size, &size);
			m_nToolsHSpacing = size.cx;
		}
		hResult = pBand->put_ToolsHSpacing(m_nToolsHSpacing);
		if (FAILED(hResult))
			return FALSE;

		if (m_nToolsVSpacing > 0)
		{
			size.cx = m_nToolsVSpacing;
			PixelToTwips(&size, &size);
			m_nToolsVSpacing = size.cx;
		}
		hResult = pBand->put_ToolsVSpacing(m_nToolsVSpacing);
		if (FAILED(hResult))
			return FALSE;

		size.cx = m_rcDimension.Width();
		PixelToTwips(&size, &size);
		hResult = pBand->put_Width(size.cx);
		if (FAILED(hResult))
			return FALSE;

		size.cx = m_rcDimension.Height();
		PixelToTwips(&size, &size);
		hResult = pBand->put_Height(size.cx);
		if (FAILED(hResult))
			return FALSE;

		hResult = pBand->get_Tools((Tools**)&pTools);
		if (FAILED(hResult))
			return FALSE;

		bResult = m_pTools->Convert(pTools);
		if (!bResult)
		{
			bResult = FALSE;
			goto Cleanup;
		}

		hResult = pBand->get_ChildBands((ChildBands**)&pChildBands);
		if (FAILED(hResult))
		{
			//
			// No child bands
			//

			goto Cleanup;
		}

		bResult = m_pChildBands->Convert(pChildBands);
		if (!bResult)
		{
			bResult = FALSE;
			goto Cleanup;
		}

Cleanup:
		if (pTools)
			pTools->Release();

		if (pChildBands)
			pChildBands->Release();

		return bResult;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
}

//
// AB10Bands
//

AB10Bands::~AB10Bands()
{
	int nCount = m_faBands.GetSize();
	for (int nBand = 0; nBand < nCount; nBand++)
		delete m_faBands.GetAt(nBand);
}

//
// Read
//

BOOL AB10Bands::Read(IStream* pStream, int nCount, BOOL bSpecificBand)
{
	try
	{
		BOOL bResult = TRUE;
		BSTR bstrBandName = NULL;
		AB10Band* pBand;
		for (int nBand = 0; nBand < nCount; nBand++)
		{
			pBand = new AB10Band(m_pBar);
			if (pBand)
			{
				bResult = pBand->Read(pStream);
				if (!bResult)
					return bResult;
			}

			if (bSpecificBand)
			{
				if (pBand->m_bstrName && 0 == wcscmp(pBand->m_bstrName, bstrBandName))
					m_faBands.Add(pBand);
				else
					delete pBand;
			}
			else
				m_faBands.Add(pBand);
		}
		return bResult;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
}

//
// Convert
//

BOOL AB10Bands::Convert(IBands* pBands)
{
	try
	{
		BOOL bResult = TRUE;
		HRESULT hResult;
		IBand* pBand;
		AB10Band* pBand10;
		int nCount = m_faBands.GetSize();
		for (int nBand = 0; nBand < nCount; nBand++)
		{
			pBand10 = m_faBands.GetAt(nBand);
			if (NULL == pBand10)
				continue;

			hResult = pBands->Add(pBand10->m_bstrName, (Band**)&pBand);
			if (FAILED(hResult))
				goto Cleanup;

			pBand10->Convert(pBand);

			pBand->Release();
			pBand = NULL;
		}

Cleanup:
		if (pBand)
			pBand->Release();
		return bResult;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
}

//
// AB10
//

AB10::AB10()
	: m_bstrLocaleString(NULL)
{
	m_pBands = new AB10Bands(this);
	assert(m_pBands);
	m_pTools = new AB10Tools(this);
	assert(m_pTools);
	m_pImageMgr = new CImageMgr;
	assert(m_pImageMgr);
	m_nHighestToolId = 0;
	m_ocBackColor = 0x80000000 + COLOR_BTNFACE;
	m_ocForeColor = 0x80000000 + COLOR_BTNTEXT;
	m_ocHighlightColor = 0x80000000 + COLOR_BTNHIGHLIGHT;
	m_ocShadowColor = 0x80000000 + COLOR_BTNSHADOW;
	m_msMenuFontStyle = ddMSSystem;
	m_vbDisplayKeysInToolTip = VARIANT_FALSE;;
	m_vbDisplayToolTips = VARIANT_TRUE;
}

AB10::~AB10()
{
	delete m_pTools;
	delete m_pBands;
	delete m_pImageMgr;
	SysFreeString(m_bstrLocaleString);
}

//
// ReadState
//

BOOL AB10::ReadState(IStream* pStream)
{
	try
	{
		HRESULT hResult;
		DWORD dwFilever = 0xA0B0;
		ULONG nBytesRead;
		DWORD nElem;
		DWORD nItemId;
		BOOL bSpecificBand = FALSE;
		BOOL bResult;

		hResult = pStream->Read(&dwFilever, sizeof(DWORD), &nBytesRead);
		if (dwFilever != 0xA0B0)
			return S_FALSE;

		DWORD dwFlags;
		hResult = pStream->Read(&dwFlags, sizeof(DWORD), &nBytesRead);
		if (FAILED(hResult))
			return FALSE;

		if (pStream->Read(&nElem,sizeof(DWORD),NULL) != S_OK)
			return FALSE; // failed

		 // Extended Info Marker
		if (0xFFFFFFFF == nElem)
		{
			while(1)
			{
				nItemId=0;
				if (pStream->Read(&nItemId,sizeof(DWORD),NULL)!=S_OK)
					break;

				switch(nItemId)
				{
				case 0:
					break;

				case 0xC0B0:
					StReadPalette(pStream, m_hPal);
					hResult = m_pImageMgr->put_Palette((OLE_HANDLE)m_hPal);
					break;

				default:
					nItemId=0; // unrecog.
					break;
				}
				if (0 == nItemId)
					break;
			}
			if (pStream->Read(&nElem,sizeof(DWORD),NULL)!=S_OK)
				return FALSE;
		}
		if (bSpecificBand)
		{
			m_bSaveImages = TRUE;
			m_pImageMgrBandSpecific = new CImageMgr;
			if (NULL == m_pImageMgrBandSpecific)
				return FALSE;

			if (0xB0B7 == nElem)
			{
				hResult = m_pImageMgrBandSpecific->Exchange(pStream, FALSE, nElem);
				if (FAILED(hResult))
					return FALSE;

				if (pStream->Read(&nElem, sizeof(DWORD), 0) != S_OK)
					return FALSE; // failed
			}

			if (0xB0B8 == nElem)
			{
				hResult = m_pImageMgrBandSpecific->Exchange(pStream, FALSE, nElem);
				if (FAILED(hResult))
					return FALSE;

				if (pStream->Read(&nElem, sizeof(DWORD), 0) != S_OK)
					return FALSE; // failed
			}
			if (0xB0B9 == nElem)
			{
				if (pStream->Read(&nElem, sizeof(DWORD), 0) != S_OK)
					return FALSE; // failed
			}
			if (0xB0C1 == nElem)
			{
				hResult = m_pImageMgrBandSpecific->Exchange(pStream, FALSE, nElem);
				if (FAILED(hResult))
					return FALSE;

				if (pStream->Read(&nElem, sizeof(DWORD), 0) != S_OK)
					return FALSE; // failed
			}
			m_bSaveImages = FALSE;
		}
		else
		{
			if (0xB0B7 == nElem)
			{
				hResult = m_pImageMgr->Exchange(pStream, FALSE, nElem);
				if (FAILED(hResult))
					return FALSE;

				if (pStream->Read(&nElem, sizeof(DWORD), 0) != S_OK)
					return FALSE; // failed
			}
			if (0xB0B8 == nElem)
			{
				hResult = m_pImageMgr->Exchange(pStream, FALSE, nElem);
				if (FAILED(hResult))
					return FALSE;

				if (pStream->Read(&nElem, sizeof(DWORD), 0) != S_OK)
					return FALSE; // failed
			}   
			if (0xB0B9 == nElem)
			{
				if (pStream->Read(&nElem, sizeof(DWORD), 0) != S_OK)
					return FALSE; // failed
			}
			if (0xB0C1 == nElem)
			{
				hResult = m_pImageMgr->Exchange(pStream, FALSE, nElem);
				if (FAILED(hResult))
					return FALSE;

				if (S_OK != pStream->Read(&nElem, sizeof(DWORD), 0))
					return FALSE; // failed
			}
		}

		if (nElem > 10000)
			return FALSE;

		bResult = m_pBands->Read(pStream, nElem, bSpecificBand);
		if (!bResult)
			return bResult;

		while(1)
		{
			nItemId=0;
			if (S_OK != pStream->Read(&nItemId,sizeof(DWORD),NULL))
				break;

			switch(nItemId)
			{
			case 0:
				break;

			case 0xB0B0:
				bResult = m_pTools->Read(pStream);
				if (!bResult)
					return bResult;
				break;

			case 0xB0B1:
				hResult = pStream->Read(&m_vbDisplayToolTips,sizeof(VARIANT_BOOL),NULL);
				if (FAILED(hResult))
					return FALSE;
				
				hResult = pStream->Read(&m_vbDisplayKeysInToolTip,sizeof(VARIANT_BOOL),NULL);
				if (FAILED(hResult))
					return FALSE;
				break;

			case 0xB0B2:
				hResult = pStream->Read(&m_nColorDepth,sizeof(m_nColorDepth),NULL);
				if (FAILED(hResult))
					return FALSE;
				
				hResult = m_pImageMgr->put_BitDepth(m_nColorDepth);
				if (FAILED(hResult))
					return FALSE;
				break;

			case 0xB0B3:
				hResult = pStream->Read(&m_msMenuFontStyle,sizeof(m_msMenuFontStyle),NULL);
				if (FAILED(hResult))
					return FALSE;
				break;

			case 0xB0B4:
				hResult = StReadBSTR(pStream,m_bstrLocaleString);
				if (FAILED(hResult))
					return FALSE;
				break;

			case 0xB0B5:
				hResult = pStream->Read(&m_ocBackColor,sizeof(m_ocBackColor),NULL);
				if (FAILED(hResult))
					return FALSE;

				hResult = pStream->Read(&m_ocForeColor,sizeof(m_ocForeColor),NULL);
				if (FAILED(hResult))
					return FALSE;

				hResult = pStream->Read(&m_ocHighlightColor,sizeof(m_ocHighlightColor),NULL);
				if (FAILED(hResult))
					return FALSE;

				hResult = pStream->Read(&m_ocShadowColor,sizeof(m_ocShadowColor),NULL);
				if (FAILED(hResult))
					return FALSE;

				PersistPict(pStream,&m_pPict,FALSE);

				hResult = PersistFont(pStream,&m_controlFont,NULL,FALSE);
				if (FAILED(hResult))
					return FALSE;
				break;

			case 0xB0C0:
				{
					VARIANT_BOOL vbUseHooks;
					hResult = pStream->Read(&vbUseHooks,sizeof(vbUseHooks),NULL);
					if (FAILED(hResult))
						return FALSE;
				}
				break;

			default:
				nItemId=0; // unrecog.
				break;
			}
			if (nItemId==0)
				break;
		}
		return TRUE;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
}

//
// Convert
//

BOOL AB10::Convert(IActiveBar2* pBar)
{
	try
	{
		BOOL bResult = TRUE;
		IBands* pBands = NULL;
		ITools* pTools = NULL;

		HRESULT hResult = pBar->get_Bands((Bands**)&pBands);
		if (FAILED(hResult))
		{
			bResult = FALSE;
			goto Cleanup;
		}

		hResult = pBands->RemoveAll();
		if (FAILED(hResult))
		{
			bResult = FALSE;
			goto Cleanup;
		}

		hResult = pBar->get_Tools((Tools**)&pTools);
		if (FAILED(hResult))
		{
			bResult = FALSE;
			goto Cleanup;
		}

		hResult = pTools->RemoveAll();
		if (FAILED(hResult))
		{
			bResult = FALSE;
			goto Cleanup;
		}

		hResult = m_pBands->Convert(pBands);
		if (FAILED(hResult))
		{
			bResult = FALSE;
			goto Cleanup;
		}

		hResult = m_pTools->Convert(pTools);
		if (FAILED(hResult))
		{
			bResult = FALSE;
			goto Cleanup;
		}

		if (m_controlFont.GetFontDispatch())
			hResult = pBar->put_ControlFont(m_controlFont.GetFontDispatch());
		if (m_pPict.GetPictureDispatch())
			hResult = pBar->put_Picture(m_pPict.GetPictureDispatch());
		hResult = pBar->put_BackColor(m_ocBackColor);
		hResult = pBar->put_ForeColor(m_ocForeColor);
		hResult = pBar->put_HighlightColor(m_ocHighlightColor);
		hResult = pBar->put_ShadowColor(m_ocShadowColor);
		hResult = pBar->put_MenuFontStyle(m_msMenuFontStyle);
		hResult = pBar->put_DisplayKeysInToolTip(m_vbDisplayKeysInToolTip);
		hResult = pBar->put_DisplayToolTips(m_vbDisplayToolTips);

Cleanup:
		if (pBands)
			pBands->Release();

		if (pTools)
			pTools->Release();

		return bResult;
	}
	CATCH
	{
		assert(FALSE);
		REPORTEXCEPTION(__FILE__, __LINE__)
		return FALSE;
	}
}

