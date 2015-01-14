/*
 * Copyright(c) 2015 arton
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

#include "stdafx.h"
#include "RubyScript.h"
#include "EnumVariant.h"
#include "ScriptError.h"

HRESULT  STDMETHODCALLTYPE CEnumVariant::QueryInterface(
    const IID & riid,  
    void **ppvObj)
{
    if (!ppvObj) return E_POINTER;

    if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IEnumVARIANT))
    {
        *ppvObj = this;
    }
    else
    {
        return E_NOINTERFACE;
    }
    InterlockedIncrement(&m_lCount);
    return S_OK;
}

ULONG  STDMETHODCALLTYPE CEnumVariant::AddRef()
{
    return InterlockedIncrement(&m_lCount);
}

ULONG  STDMETHODCALLTYPE CEnumVariant::Release()
{
    if (InterlockedDecrement(&m_lCount) == 0)
    {
        delete this;
        return 0;
    }
    return m_lCount;
}

HRESULT STDMETHODCALLTYPE CEnumVariant::Next( 
    /* [in] */ ULONG celt,
    /* [length_is][size_is][out] */ VARIANT *rgVar,
    /* [out] */ ULONG *pCeltFetched)
{
    if (m_endOfEnum) return S_FALSE;
    VALUE params[] = {
        m_enum,
        rb_intern("next"),
        0
    };
    if (pCeltFetched) *pCeltFetched = 0;
    int nstate(0);
    for (ULONG i = 0; i < celt; i++)
    {
        VariantInit(rgVar + i);
        volatile VALUE v = rb_protect(CRubyScript::safe_funcall, (VALUE)params, &nstate);
        if (nstate)
        {
            volatile VALUE v = rb_errinfo();
            if (!NIL_P(v))
            {
                volatile VALUE msg = rb_funcall(v, rb_intern("message"), 0);
                if (!NIL_P(msg)) ATLTRACE(_T("msg:%hs\n"), StringValueCStr(msg));
            }
            return S_FALSE;
        }
        CRubyScript::CreateVariant(v, rgVar + i, true);
        if (pCeltFetched) *pCeltFetched++;
        m_nIterate++;
    }
    return S_OK;
}
        
HRESULT STDMETHODCALLTYPE CEnumVariant::Skip( 
    /* [in] */ ULONG celt)
{
    if (m_endOfEnum) return S_FALSE;
    VALUE params[] = {
        m_enum,
        rb_intern("next"),
        0
    };
    int nstate(0);
    for (ULONG i = 0; i < celt; i++)
    {
        rb_protect(CRubyScript::safe_funcall, (VALUE)params, &nstate);
        if (nstate)
        {
            return S_FALSE;
        }
        m_nIterate++;
    }
    return S_OK;
}
        
HRESULT STDMETHODCALLTYPE CEnumVariant::Reset( void)
{
    m_nIterate = 0;
    rb_funcall(m_object, rb_intern("rewind"), 0);
    return S_OK;
}
        
HRESULT STDMETHODCALLTYPE CEnumVariant::Clone( 
    /* [out] */ __RPC__deref_out_opt IEnumVARIANT **ppEnum)
{
    *ppEnum = new CEnumVariant(*this);
    return S_OK;
}

