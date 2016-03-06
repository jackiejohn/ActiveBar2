#ifndef __XEVENTS_H__
#define __XEVENTS_H__

//
//  Copyright © 1995-1999, Data Dynamics. All rights reserved.
//
//  Unpublished -- rights reserved under the copyright laws of the United
//  States.  USE OF A COPYRIGHT NOTICE IS PRECAUTIONARY ONLY AND DOES NOT
//  IMPLY PUBLICATION OR DISCLOSURE.
//
//

#define VTS_I2              "\x02"      // a 'short'
#define VTS_I4              "\x03"      // a 'long'
#define VTS_R4              "\x04"      // a 'float'
#define VTS_R8              "\x05"      // a 'double'
#define VTS_CY              "\x06"      // a 'CY' or 'CY*'
#define VTS_DATE            "\x07"      // a 'DATE'
#define VTS_WBSTR           "\x08"      // an 'LPCOLESTR'
#define VTS_DISPATCH        "\x09"      // an 'IDispatch*'
#define VTS_SCODE           "\x0A"      // an 'SCODE'
#define VTS_BOOL            "\x0B"      // a 'BOOL'
#define VTS_VARIANT         "\x0C"      // a 'const VARIANT&' or 'VARIANT*'
#define VTS_UNKNOWN         "\x0D"      // an 'IUnknown*'
#if defined(_UNICODE) || defined(OLE2ANSI)
	#define VTS_BSTR            VTS_WBSTR// an 'LPCOLESTR'
	#define VT_BSTRT            VT_BSTR
#else
	#define VTS_BSTR            "\x0E"  // an 'LPCSTR'
	#define VT_BSTRA            14
	#define VT_BSTRT            VT_BSTRA
#endif
// parameter types: by reference VTs
#define VTS_PI2             "\x42"      // a 'short*'
#define VTS_PI4             "\x43"      // a 'long*'
#define VTS_PR4             "\x44"      // a 'float*'
#define VTS_PR8             "\x45"      // a 'double*'
#define VTS_PCY             "\x46"      // a 'CY*'
#define VTS_PDATE           "\x47"      // a 'DATE*'
#define VTS_PBSTR           "\x48"      // a 'BSTR*'
#define VTS_PDISPATCH       "\x49"      // an 'IDispatch**'
#define VTS_PSCODE          "\x4A"      // an 'SCODE*'
#define VTS_PBOOL           "\x4B"      // a 'VARIANT_BOOL*'
#define VTS_PVARIANT        "\x4C"      // a 'VARIANT*'
#define VTS_PUNKNOWN        "\x4D"      // an 'IUnknown**'
// special VT_ and VTS_ values
#define VTS_NONE            "\x00"        // used for members with 0 params
#define VTS_COLOR           VTS_I4      // OLE_COLOR
#define VTS_XPOS_PIXELS     VTS_I4      // OLE_XPOS_PIXELS
#define VTS_YPOS_PIXELS     VTS_I4      // OLE_YPOS_PIXELS
#define VTS_XSIZE_PIXELS    VTS_I4      // OLE_XSIZE_PIXELS
#define VTS_YSIZE_PIXELS    VTS_I4      // OLE_YSIZE_PIXELS
#define VTS_XPOS_HIMETRIC   VTS_I4      // OLE_XPOS_HIMETRIC
#define VTS_YPOS_HIMETRIC   VTS_I4      // OLE_YPOS_HIMETRIC
#define VTS_XSIZE_HIMETRIC  VTS_I4      // OLE_XSIZE_HIMETRIC
#define VTS_YSIZE_HIMETRIC  VTS_I4      // OLE_YSIZE_HIMETRIC
#define VTS_TRISTATE        VTS_I2      // OLE_TRISTATE
#define VTS_OPTEXCLUSIVE    VTS_BOOL    // OLE_OPTEXCLUSIVE
#define VTS_PCOLOR          VTS_PI4     // OLE_COLOR*
#define VTS_PXPOS_PIXELS    VTS_PI4     // OLE_XPOS_PIXELS*
#define VTS_PYPOS_PIXELS    VTS_PI4     // OLE_YPOS_PIXELS*
#define VTS_PXSIZE_PIXELS   VTS_PI4     // OLE_XSIZE_PIXELS*
#define VTS_PYSIZE_PIXELS   VTS_PI4     // OLE_YSIZE_PIXELS*
#define VTS_PXPOS_HIMETRIC  VTS_PI4     // OLE_XPOS_HIMETRIC*
#define VTS_PYPOS_HIMETRIC  VTS_PI4     // OLE_YPOS_HIMETRIC*
#define VTS_PXSIZE_HIMETRIC VTS_PI4     // OLE_XSIZE_HIMETRIC*
#define VTS_PYSIZE_HIMETRIC VTS_PI4     // OLE_YSIZE_HIMETRIC*
#define VTS_PTRISTATE       VTS_PI2     // OLE_TRISTATE*
#define VTS_POPTEXCLUSIVE   VTS_PBOOL   // OLE_OPTEXCLUSIVE*
#define VTS_FONT            VTS_DISPATCH    // IFontDispatch*
#define VTS_PICTURE         VTS_DISPATCH    // IPictureDispatch*
#define VTS_HANDLE          VTS_I4      // OLE_HANDLE
#define VTS_PHANDLE         VTS_PI4     // OLE_HANDLE*

#define EVENT_PARAM(vtsParams) (BYTE*)(vtsParams)


struct XFLOAT   { BYTE floatBits[sizeof(float)]; };
struct XDOUBLE  { BYTE doubleBits[sizeof(double)]; };

#if defined(_68K_) || defined(_X86_)
	#define DOUBLE_ARG  XDOUBLE
#else
	#define DOUBLE_ARG  double
#endif

extern HRESULT DispatchHelper(LPDISPATCH pDispatch,DISPID dwDispID, WORD wFlags,
						   VARTYPE vtRet, void* pvRet, const BYTE* pbParamInfo, va_list argList);


#endif