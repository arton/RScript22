/*
 * Copyright(c) 2014 arton
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 */

#ifndef __RUBYIZE_H_
#define __RUBYIZE_H_

#include "resource.h"
#include "RScript22_i.h"
#include "giplip.h"

#if defined(RUBY_2_2)
#define MAJOR_VERSION 2
#define MINOR_VERSION 2
#else
#error "no ruby version defined"
#endif

/////////////////////////////////////////////////////////////////////////////
// CRubyize
class ATL_NO_VTABLE CRubyize : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CRubyize, &CLSID_Rubyize>,
    public ISupportErrorInfo,
    public IDispatchImpl<IRubyize, &IID_IRubyize, &LIBID_RScript22Lib, MAJOR_VERSION, MINOR_VERSION>,
    public CObjectStore
{
public:
    CRubyize()
        : m_pAsr(NULL), m_asr(Qnil)
    {
    }

    static _ATL_REGMAP_ENTRY RegEntries[];
    static HRESULT WINAPI UpdateRegistry(BOOL bRegister);

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CRubyize)
    COM_INTERFACE_ENTRY(IRubyize)
    COM_INTERFACE_ENTRY(IDispatch)
    COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

    HRESULT FinalConstruct();

    void FinalRelease();

    inline void SetPassedObject(VARIANT& v)
    {
        _ASSERT(m_pPassedObject);
        VariantCopy(m_pPassedObject, &v);
    }

// ISupportsErrorInfo
    STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IRubyize
public:
    HRESULT STDMETHODCALLTYPE get_Version( 
            /* [retval][out] */ BSTR __RPC_FAR *pVersion);

    HRESULT STDMETHODCALLTYPE get_RubyVersion( 
            /* [retval][out] */ BSTR __RPC_FAR *pVersion);

    HRESULT STDMETHODCALLTYPE rubyize( 
            /* [in] */ VARIANT val,
            /* [retval][out] */ VARIANT __RPC_FAR *pObj);
        
    HRESULT STDMETHODCALLTYPE erubyize( 
            /* [in] */ BSTR script,
            /* [retval][out] */ VARIANT __RPC_FAR *pObj);

private:
    CComObject<CRubyScript>* m_pRubyScript;
    VALUE m_asr;
    IDispatch* m_pAsr;
    VARIANT* m_pPassedObject;
    DISPID m_dispidRubyize;
};

#endif //__RUBYIZE_H_
