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

#include "stdafx.h"
#include "RubyScript.h"
#include "ScriptObject.h"

HRESULT  STDMETHODCALLTYPE CScriptObject::QueryInterface(
    const IID & riid,  
    void **ppvObj)
{
    if (!ppvObj) return E_POINTER;

    if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch) || IsEqualIID(riid, IID_IDispatchEx)
        || IsEqualIID(riid, IID_IRubyScriptObject))
    {
        *ppvObj = (IDispatchEx*)this;
    }
    else if (IsEqualIID(riid, IID_ISupportErrorInfo))
    {
        *ppvObj = (ISupportErrorInfo*)this;
    }
    else
    {
        return E_NOINTERFACE;
    }
    InterlockedIncrement(&m_lCount);
    return S_OK;
}

ULONG  STDMETHODCALLTYPE CScriptObject::AddRef()
{
    return InterlockedIncrement(&m_lCount);
}

ULONG  STDMETHODCALLTYPE CScriptObject::Release()
{
    if (InterlockedDecrement(&m_lCount) == 0)
    {
        delete this;
        return 0;
    }
    return m_lCount;
}


HRESULT STDMETHODCALLTYPE CScriptObject::GetIDsOfNames( 
    /* [in] */ REFIID riid,
    /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
    /* [in] */ UINT cNames,
    /* [in] */ LCID lcid,
    /* [size_is][out] */ DISPID __RPC_FAR *rgDispId)
{
    ATLTRACE(_T("CScriptObject::GetIDsOfNames(%ls)\n"), *rgszNames);
    if (m_pDispatch)
    {
        HRESULT hr = m_pDispatch->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
        if (hr != DISP_E_UNKNOWNNAME) 
        {
            if (hr == S_OK)
            {
                for (int i = 0; i < m_cDispIds; i++)
                {
                    if (*(m_DispIds + i) == *rgDispId) return hr;
                }
                if (m_cDispIds == m_sizeDispIds)
                {
                    DISPID* p = new DISPID[m_sizeDispIds * 2];
                    memcpy(p, m_DispIds, sizeof(DISPID) * m_sizeDispIds);
                    delete[] m_DispIds;
                    m_sizeDispIds *= 2;
                    m_DispIds = p;
                }
                *(m_DispIds + m_cDispIds++) = *rgDispId;
            }
            return hr;
        }
    }
    USES_CONVERSION;
    volatile VALUE v = rb_str_new_cstr(W2A(*rgszNames));
    v = rb_funcall(m_object, rb_intern("get_method_id"), 1, v);
    *rgDispId = NUM2INT(v);
    for (UINT i = 2; i < cNames; i++)
    {
        *(rgDispId + i) = DISPID_UNKNOWN;
    }
    return (*rgDispId == DISPID_UNKNOWN) ? DISP_E_UNKNOWNNAME : S_OK;
}

static ruby_value_type WIN32OLE_CONPATIBLE[] = {
    T_ARRAY,
    T_STRING,
    T_FIXNUM,
    T_BIGNUM,
    T_FLOAT,
    T_TRUE,
    T_FALSE,
    T_NIL,
};

HRESULT STDMETHODCALLTYPE CScriptObject::Invoke( 
    /* [in] */ DISPID dispIdMember,
    /* [in] */ REFIID riid,
    /* [in] */ LCID lcid,
    /* [in] */ WORD wFlags,
    /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
    /* [out] */ VARIANT __RPC_FAR *pVarResult,
    /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
    /* [out] */ UINT __RPC_FAR *puArgErr)
{
    if (m_pDispatch)
    {
        for (int i = 0; i < m_cDispIds; i++)
        {
            if (*(m_DispIds + i) == dispIdMember)
            {
                return m_pDispatch->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
            }
        }
    }
    VALUE* argvalues = reinterpret_cast<VALUE*>(_alloca(sizeof(VALUE) * (pDispParams->cArgs + 4)));
    *argvalues = m_object;
    *(argvalues + 1) = rb_intern("call");
    *(argvalues + 2) = pDispParams->cArgs + 1;
    *(argvalues + 3) = INT2FIX(dispIdMember);
    for (UINT i = 1; i <= pDispParams->cArgs; i++)
    {
        volatile VALUE variant = CRubyScript::CreateVariant(*(pDispParams->rgvarg + pDispParams->cArgs - i));
        *(argvalues + 3 + i) = rb_funcall(variant, rb_intern("value"), 0);
    }
    int state(0);
    VALUE v = rb_protect(CRubyScript::safe_funcall, (VALUE)argvalues, &state);
    if (state)
    {
        volatile VALUE v = rb_errinfo();
        if (!NIL_P(v))
        {
            size_t len(0);
            volatile VALUE msg = rb_funcall(v, rb_intern("message"), 0);
            if (!NIL_P(msg)) len = strlen(StringValueCStr(msg));
            volatile VALUE excep = rb_funcall(v, rb_intern("class"), 0);
            excep = rb_funcall(excep, rb_intern("to_s"), 0);
            len += strlen(StringValueCStr(excep));
            LPOLESTR p = reinterpret_cast<LPOLESTR>(_alloca(sizeof(OLECHAR) * (len + 8)));
            swprintf(p, L"%hs: %hs", StringValueCStr(excep), (NIL_P(msg)) ? "" : StringValueCStr(msg));
            ICreateErrorInfo* pceinfo;
            IErrorInfo* perrinfo;
            if (CreateErrorInfo(&pceinfo) == S_OK)
            {
                pceinfo->SetGUID(IID_IRubyScriptObject);
                pceinfo->SetDescription(p);
                pceinfo->SetSource(L"RubyScript");
                if (pceinfo->QueryInterface(IID_IErrorInfo, (void**)&perrinfo) == S_OK)
                {
                    SetErrorInfo(0, perrinfo);
                    perrinfo->Release();
                }
                pceinfo->Release();
            }
            if (pExcepInfo)
            {
                pExcepInfo->wCode = 1001;
                pExcepInfo->wReserved = 0;
                pExcepInfo->bstrSource = SysAllocString(L"RubyScript");
                pExcepInfo->bstrDescription = SysAllocString(p);
                pExcepInfo->bstrHelpFile = NULL;
                pExcepInfo->dwHelpContext = 0;
                pExcepInfo->pvReserved = NULL;
                pExcepInfo->pfnDeferredFillIn = NULL;
                pExcepInfo->scode = DISP_E_EXCEPTION;
            }
        }
        return DISP_E_EXCEPTION;
    }
    if (!pVarResult) return S_OK;
    VariantInit(pVarResult);
    UINT vtype = TYPE(v);
    for (UINT i = 0; i < COUNT_OF(WIN32OLE_CONPATIBLE); i++)
    {
        if (vtype == WIN32OLE_CONPATIBLE[i])
        {
            CRubyScript::CreateVariant(v, pVarResult);
            if (pVarResult->vt == VT_ERROR && pVarResult->scode == DISP_E_PARAMNOTFOUND)
            {
                pVarResult->vt = VT_DISPATCH;
                pVarResult->pdispVal = NULL;
            }
            return S_OK;
        }
    }
    pVarResult->vt = VT_DISPATCH;
    pVarResult->pdispVal = m_pObjectStore->CreateDispatch(v);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE CScriptObject::InterfaceSupportsErrorInfo(REFIID riid)
{
    if (InlineIsEqualGUID(IID_IRubyScriptObject, riid))
    {
	return S_OK;
    }
    return S_FALSE;
}
