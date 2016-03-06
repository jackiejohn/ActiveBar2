

/* this ALWAYS GENERATED file contains the IIDs and CLSIDs */

/* link this file in with the server and any clients */


 /* File created by MIDL compiler version 7.00.0500 */
/* at Fri Jan 29 07:25:35 2016
 */
/* Compiler settings for .\Private.odl:
    Oicf, W1, Zp8, env=Win32 (32b run)
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
//@@MIDL_FILE_HEADING(  )

#pragma warning( disable: 4049 )  /* more than 64k source lines */


#ifdef __cplusplus
extern "C"{
#endif 


#include <rpc.h>
#include <rpcndr.h>

#ifdef _MIDL_USE_GUIDDEF_

#ifndef INITGUID
#define INITGUID
#include <guiddef.h>
#undef INITGUID
#else
#include <guiddef.h>
#endif

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        DEFINE_GUID(name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8)

#else // !_MIDL_USE_GUIDDEF_

#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        const type name = {l,w1,w2,{b1,b2,b3,b4,b5,b6,b7,b8}}

#endif !_MIDL_USE_GUIDDEF_

MIDL_DEFINE_GUID(IID, LIBID_ActiveBar2PrivateLibrary,0x140C8648,0xA2B8,0x47B6,0x9A,0x6F,0xD5,0xEF,0x27,0x6D,0xAF,0xEA);


MIDL_DEFINE_GUID(IID, DIID_ICustomTool,0xA4672F71,0x1E31,0x11D1,0xA9,0x6E,0x00,0x60,0x08,0x1C,0x43,0xD9);


MIDL_DEFINE_GUID(IID, IID_ICustomTool2,0xD56E5130,0x44D4,0x11D1,0xA9,0xF8,0x00,0x60,0x08,0x1C,0x43,0xD9);


MIDL_DEFINE_GUID(IID, IID_IDesignerNotifications,0xDDC9DB53,0x5F79,0x11D2,0xA1,0xB6,0x00,0x60,0x08,0x1C,0x43,0xD9);


MIDL_DEFINE_GUID(IID, IID_IDesignerNotify,0x1B9CB537,0xC02F,0x11D2,0xA2,0x77,0x00,0x60,0x08,0x1C,0x43,0xD9);


MIDL_DEFINE_GUID(IID, IID_IDragDropManager,0x55F1245B,0xB9DC,0x11D2,0xA2,0x65,0x00,0x60,0x08,0x1C,0x43,0xD9);


MIDL_DEFINE_GUID(IID, IID_ICategorizeProperties,0x4D07FC10,0xF931,0x11CE,0xB0,0x01,0x00,0xAA,0x00,0x68,0x84,0xE5);


MIDL_DEFINE_GUID(IID, IID_IBarPrivate,0x89541522,0x2D31,0x11D2,0xA1,0x66,0x00,0x60,0x08,0x1C,0x43,0xD9);


MIDL_DEFINE_GUID(IID, IID_ICustomHost,0x89541524,0x2D31,0x11D2,0xA1,0x66,0x00,0x60,0x08,0x1C,0x43,0xD9);


MIDL_DEFINE_GUID(IID, IID_IPerPropertyBrowsing,0x51973C54,0xCB0C,0x11D0,0xB5,0xC9,0x00,0xA0,0x24,0x4A,0x0E,0x7A);


MIDL_DEFINE_GUID(IID, IID_IDDPerPropertyBrowsing,0x6AFD2D9C,0xF2B0,0x4E83,0x86,0x06,0x12,0x42,0x79,0x61,0x72,0x2A);

#undef MIDL_DEFINE_GUID

#ifdef __cplusplus
}
#endif



