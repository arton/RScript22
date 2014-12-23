

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 7.00.0555 */
/* at Tue Dec 23 18:11:59 2014
 */
/* Compiler settings for RScript22.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 7.00.0555 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__


#ifndef __RScript22_i_h__
#define __RScript22_i_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __RubyScript_FWD_DEFINED__
#define __RubyScript_FWD_DEFINED__

#ifdef __cplusplus
typedef class RubyScript RubyScript;
#else
typedef struct RubyScript RubyScript;
#endif /* __cplusplus */

#endif 	/* __RubyScript_FWD_DEFINED__ */


#ifndef __IRubyize_FWD_DEFINED__
#define __IRubyize_FWD_DEFINED__
typedef interface IRubyize IRubyize;
#endif 	/* __IRubyize_FWD_DEFINED__ */


#ifndef __Rubyize_FWD_DEFINED__
#define __Rubyize_FWD_DEFINED__

#ifdef __cplusplus
typedef class Rubyize Rubyize;
#else
typedef struct Rubyize Rubyize;
#endif /* __cplusplus */

#endif 	/* __Rubyize_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"
#include "activscp.h"

#ifdef __cplusplus
extern "C"{
#endif 



#ifndef __RScript22Lib_LIBRARY_DEFINED__
#define __RScript22Lib_LIBRARY_DEFINED__

/* library RScript22Lib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_RScript22Lib;

EXTERN_C const CLSID CLSID_RubyScript;

#ifdef __cplusplus

class DECLSPEC_UUID("456A3763-90A4-4F2A-BFF1-4B773C1056EC")
RubyScript;
#endif

#ifndef __IRubyize_INTERFACE_DEFINED__
#define __IRubyize_INTERFACE_DEFINED__

/* interface IRubyize */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IRubyize;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("0A4CBEBD-C46B-4A7C-A1E2-AD474C330C7A")
    IRubyize : public IDispatch
    {
    public:
        virtual /* [propget][id] */ HRESULT STDMETHODCALLTYPE get_Version( 
            /* [retval][out] */ BSTR *pVersion) = 0;
        
        virtual /* [propget][id] */ HRESULT STDMETHODCALLTYPE get_RubyVersion( 
            /* [retval][out] */ BSTR *pVersion) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE rubyize( 
            /* [in] */ VARIANT val,
            /* [retval][out] */ VARIANT *pObj) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE erubyize( 
            /* [in] */ BSTR script,
            /* [retval][out] */ VARIANT *pObj) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IRubyizeVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IRubyize * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IRubyize * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IRubyize * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IRubyize * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IRubyize * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IRubyize * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IRubyize * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_Version )( 
            IRubyize * This,
            /* [retval][out] */ BSTR *pVersion);
        
        /* [propget][id] */ HRESULT ( STDMETHODCALLTYPE *get_RubyVersion )( 
            IRubyize * This,
            /* [retval][out] */ BSTR *pVersion);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *rubyize )( 
            IRubyize * This,
            /* [in] */ VARIANT val,
            /* [retval][out] */ VARIANT *pObj);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *erubyize )( 
            IRubyize * This,
            /* [in] */ BSTR script,
            /* [retval][out] */ VARIANT *pObj);
        
        END_INTERFACE
    } IRubyizeVtbl;

    interface IRubyize
    {
        CONST_VTBL struct IRubyizeVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IRubyize_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IRubyize_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IRubyize_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IRubyize_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IRubyize_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IRubyize_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IRubyize_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IRubyize_get_Version(This,pVersion)	\
    ( (This)->lpVtbl -> get_Version(This,pVersion) ) 

#define IRubyize_get_RubyVersion(This,pVersion)	\
    ( (This)->lpVtbl -> get_RubyVersion(This,pVersion) ) 

#define IRubyize_rubyize(This,val,pObj)	\
    ( (This)->lpVtbl -> rubyize(This,val,pObj) ) 

#define IRubyize_erubyize(This,script,pObj)	\
    ( (This)->lpVtbl -> erubyize(This,script,pObj) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IRubyize_INTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_Rubyize;

#ifdef __cplusplus

class DECLSPEC_UUID("0BCFF05A-C2BF-4CB2-A778-3428A8E85A21")
Rubyize;
#endif
#endif /* __RScript22Lib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


