/* this file contains the actual definitions of */
/* the IIDs and CLSIDs */

/* link this file in with the server and any clients */


/* File created by MIDL compiler version 5.01.0164 */
/* at Fri Mar 01 10:38:31 2002
 */
/* Compiler settings for C:\Dev\ActBar20\Source\Designer\prj.odl:
    Os (OptLev=s), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )
#ifdef __cplusplus
extern "C"{
#endif 


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

const IID LIBID_Designer = {0x4EB91002,0x2661,0x11D2,{0xBC,0x36,0x8D,0xFE,0xBE,0x3A,0x8B,0x36}};


const IID IID_IDesigner = {0x0564AE52,0xFEFE,0x11D2,{0xA2,0xF4,0x00,0x60,0x08,0x1C,0x43,0xD9}};


const CLSID CLSID_DesignerPage = {0x5A168040,0x2657,0x11D2,{0xBC,0x36,0x8D,0xFE,0xBE,0x3A,0x8B,0x36}};


const CLSID CLSID_Designer = {0x0564AE54,0xFEFE,0x11D2,{0xA2,0xF4,0x00,0x60,0x08,0x1C,0x43,0xD9}};


#ifdef __cplusplus
}
#endif

