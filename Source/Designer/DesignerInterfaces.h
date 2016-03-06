/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Fri Mar 01 10:38:31 2002
 */
/* Compiler settings for C:\Dev\ActBar20\Source\Designer\prj.odl:
    Os (OptLev=s), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __DesignerInterfaces_h__
#define __DesignerInterfaces_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __IDesigner_FWD_DEFINED__
#define __IDesigner_FWD_DEFINED__
typedef interface IDesigner IDesigner;
#endif 	/* __IDesigner_FWD_DEFINED__ */


#ifndef __DesignerPage_FWD_DEFINED__
#define __DesignerPage_FWD_DEFINED__

#ifdef __cplusplus
typedef class DesignerPage DesignerPage;
#else
typedef struct DesignerPage DesignerPage;
#endif /* __cplusplus */

#endif 	/* __DesignerPage_FWD_DEFINED__ */


#ifndef __Designer_FWD_DEFINED__
#define __Designer_FWD_DEFINED__

#ifdef __cplusplus
typedef class Designer Designer;
#else
typedef struct Designer Designer;
#endif /* __cplusplus */

#endif 	/* __Designer_FWD_DEFINED__ */


void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 


#ifndef __Designer_LIBRARY_DEFINED__
#define __Designer_LIBRARY_DEFINED__

/* library Designer */
/* [version][lcid][helpstring][uuid] */ 

typedef /* [uuid] */ 
enum __MIDL___MIDL_itf_prj_0094_0001
    {	Hierarchy	= 0,
	Category	= 1
    }	DesignerTreeTypes;


DEFINE_GUID(LIBID_Designer,0x4EB91002,0x2661,0x11D2,0xBC,0x36,0x8D,0xFE,0xBE,0x3A,0x8B,0x36);

#ifndef __IDesigner_INTERFACE_DEFINED__
#define __IDesigner_INTERFACE_DEFINED__

/* interface IDesigner */
/* [object][dual][uuid] */ 


DEFINE_GUID(IID_IDesigner,0x0564AE52,0xFEFE,0x11D2,0xA2,0xF4,0x00,0x60,0x08,0x1C,0x43,0xD9);

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("0564AE52-FEFE-11D2-A2F4-0060081C43D9")
    IDesigner : public IDispatch
    {
    public:
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE OpenDesigner( 
            /* [in] */ /* external definition not present */ OLE_HANDLE hWndParent,
            /* [in] */ LPDISPATCH pActiveBar,
            /* [in] */ VARIANT_BOOL vbStandALone) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE CloseDesigner( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE SetFocus( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE UIDeactivateCloseDesigner( void) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IDesignerVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IDesigner __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IDesigner __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IDesigner __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IDesigner __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IDesigner __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IDesigner __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IDesigner __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *OpenDesigner )( 
            IDesigner __RPC_FAR * This,
            /* [in] */ /* external definition not present */ OLE_HANDLE hWndParent,
            /* [in] */ LPDISPATCH pActiveBar,
            /* [in] */ VARIANT_BOOL vbStandALone);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *CloseDesigner )( 
            IDesigner __RPC_FAR * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetFocus )( 
            IDesigner __RPC_FAR * This);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *UIDeactivateCloseDesigner )( 
            IDesigner __RPC_FAR * This);
        
        END_INTERFACE
    } IDesignerVtbl;

    interface IDesigner
    {
        CONST_VTBL struct IDesignerVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IDesigner_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IDesigner_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IDesigner_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IDesigner_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IDesigner_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IDesigner_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IDesigner_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IDesigner_OpenDesigner(This,hWndParent,pActiveBar,vbStandALone)	\
    (This)->lpVtbl -> OpenDesigner(This,hWndParent,pActiveBar,vbStandALone)

#define IDesigner_CloseDesigner(This)	\
    (This)->lpVtbl -> CloseDesigner(This)

#define IDesigner_SetFocus(This)	\
    (This)->lpVtbl -> SetFocus(This)

#define IDesigner_UIDeactivateCloseDesigner(This)	\
    (This)->lpVtbl -> UIDeactivateCloseDesigner(This)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [id] */ HRESULT STDMETHODCALLTYPE IDesigner_OpenDesigner_Proxy( 
    IDesigner __RPC_FAR * This,
    /* [in] */ /* external definition not present */ OLE_HANDLE hWndParent,
    /* [in] */ LPDISPATCH pActiveBar,
    /* [in] */ VARIANT_BOOL vbStandALone);


void __RPC_STUB IDesigner_OpenDesigner_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE IDesigner_CloseDesigner_Proxy( 
    IDesigner __RPC_FAR * This);


void __RPC_STUB IDesigner_CloseDesigner_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE IDesigner_SetFocus_Proxy( 
    IDesigner __RPC_FAR * This);


void __RPC_STUB IDesigner_SetFocus_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [id] */ HRESULT STDMETHODCALLTYPE IDesigner_UIDeactivateCloseDesigner_Proxy( 
    IDesigner __RPC_FAR * This);


void __RPC_STUB IDesigner_UIDeactivateCloseDesigner_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IDesigner_INTERFACE_DEFINED__ */


DEFINE_GUID(CLSID_DesignerPage,0x5A168040,0x2657,0x11D2,0xBC,0x36,0x8D,0xFE,0xBE,0x3A,0x8B,0x36);

#ifdef __cplusplus

class DECLSPEC_UUID("5A168040-2657-11D2-BC36-8DFEBE3A8B36")
DesignerPage;
#endif

DEFINE_GUID(CLSID_Designer,0x0564AE54,0xFEFE,0x11D2,0xA2,0xF4,0x00,0x60,0x08,0x1C,0x43,0xD9);

#ifdef __cplusplus

class DECLSPEC_UUID("0564AE54-FEFE-11D2-A2F4-0060081C43D9")
Designer;
#endif
#endif /* __Designer_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
