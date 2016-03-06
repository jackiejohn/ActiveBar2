#ifndef ABFORMAT_INCLUDED
#define ABFORMAT_INCLUDED

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#include "..\Interfaces.h"

class CImageMgr;
class AB10;

//
// Conversion Structures
//

//
// AB10Tool
//

class AB10Tool
{
public:
	AB10Tool(AB10* pBar);
	~AB10Tool();

	enum
	{
		eNumOfImageTypes = 4
	};

	BOOL Read(IStream* pStream);
	BOOL Convert(ITool* pTool);

	long m_nHelpContextId;
	long m_nToolId;
	long m_nTag;
	int  m_nHeight;
	int  m_nWidth;

	BSTR m_bstrDescription;
	BSTR m_bstrToolTipText;
	BSTR m_bstrCategory;
	BSTR m_bstrSubBand;
	BSTR m_bstrCaption;
	BSTR m_bstrText;
	BSTR m_bstrName;

	BOOL m_bChecked;
	BOOL m_bEnabled;

	CaptionPositionTypes m_cpTool;
	ToolAlignmentTypes m_taTool;
	ToolTypes m_ttTool;
	ToolStyles m_tsTool;
	
	DWORD m_dwShortCutKey;
	DWORD m_dwImageCookie;
	DWORD m_dwCustomFlag;
	long m_nImageIds[4];
	
	VARIANT_BOOL m_vbBeginGroup;
	VARIANT_BOOL m_vbVisible;

	AB10* m_pBar;
};

//
// AB10Tools
//

class AB10Tools
{
public:
	AB10Tools(AB10* pBar)
		: m_pBar(pBar)
	{
	}

	~AB10Tools();

	BOOL Read(IStream* pStream);
	BOOL Convert(ITools* pTools);

	TypedArray<AB10Tool*> m_faTools;
	AB10* m_pBar;
};

//
// AB10ChildBand
//

class AB10ChildBand
{
public:
	AB10ChildBand(AB10* pBar);
	~AB10ChildBand();

	BOOL Read(IStream* pStream);
	BOOL Convert(IBand* pChildBand);

	AB10Tools* m_pTools;
	BSTR m_bstrCaption;
	AB10* m_pBar;
};

//
// AB10ChildBands
//

class AB10ChildBands
{
public:
	AB10ChildBands(AB10* pBar)
		: m_pBar(pBar)
	{
	}

	~AB10ChildBands();

	BOOL Read(IStream* pStream);
	BOOL Convert(IChildBands* pChildBands);

	TypedArray<AB10ChildBand*> m_faChildBands;
	AB10* m_pBar;
};

//
// AB10Band
//

class AB10Band
{
public:
	AB10Band(AB10* pBar);
	~AB10Band();

	BOOL Read(IStream* pStream);
	BOOL Convert(IBand* pBand);

	int ConvertDockingArea();

	TrackingStyles m_tsMouseTracking;
	VARIANT_BOOL m_vbDisplayHandles;
	VARIANT_BOOL m_vbVisible;
	VARIANT_BOOL m_vbWrappable;
	CFontHolder  m_font;
	DockingAreaTypes  m_daDockingArea;
	ChildBandStyles   m_psStyle;
	BandTypes    m_btType;
	DWORD m_dwFlags;
	short m_nGrabHandleStyle;
	short m_nCurrentPage;
	short m_nCreatedBy;
	short m_nDockLine;
	CRect m_rcDimension;
	CRect m_rcFloat;
	BSTR  m_bstrCaption;
	BSTR  m_bstrName;
	BOOL  m_bIsDetached;
	int   m_nToolsHPadding;
	int   m_nToolsVPadding;
	int   m_nToolsHSpacing;
	int   m_nToolsVSpacing;
	int   m_nDockOffset;
	int   m_nCurrentTool;

	AB10Tools* m_pTools;
	AB10ChildBands* m_pChildBands;
	AB10* m_pBar;
};

class AB10Bands
{
public:
	AB10Bands(AB10* pBar)
		: m_pBar(pBar)
	{
	}

	~AB10Bands();

	BOOL Read(IStream* pStream, int nCount, BOOL bSpecificBand);
	BOOL Convert(IBands* pBands);

	TypedArray<AB10Band*> m_faBands;
	AB10* m_pBar;
};

//
// AB10
//

class AB10
{
public:
	AB10();
	~AB10();

	BOOL ReadState(IStream* pStream);
	BOOL Convert(IActiveBar2* pBar);

	MenuFontStyles m_msMenuFontStyle;
	CPictureHolder m_pPict;
	VARIANT_BOOL   m_vbDisplayKeysInToolTip;
	VARIANT_BOOL   m_vbDisplayToolTips;
	OLE_COLOR	   m_ocBackColor;
	OLE_COLOR	   m_ocForeColor;
	OLE_COLOR	   m_ocHighlightColor;
	OLE_COLOR	   m_ocShadowColor;
	BSTR		   m_bstrLocaleString;
	int            m_nColorDepth;
	CFontHolder m_controlFont;

	AB10Bands*	   m_pBands;
	AB10Tools*	   m_pTools;
	CImageMgr*     m_pImageMgr;
	HPALETTE       m_hPal;
	BOOL		   m_bSaveImages;
	CImageMgr*     m_pImageMgrBandSpecific;
	long m_nHighestToolId;
};

#endif
