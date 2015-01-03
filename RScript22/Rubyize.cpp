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
#include "Rubyize.h"
#include "ScriptObject.h"

#if defined(RUBY_2_2)
#define RUBYIZE_VERSION L"2.2.0"
#else
#error "no ruby version defined"
#endif

#define RUBYIZE NULL
/////////////////////////////////////////////////////////////////////////////
// CRubyize
HRESULT CRubyize::FinalConstruct()
{
    HRESULT hr = CComObject<CRubyScript>::CreateInstance(&m_pRubyScript); // initialize ruby interpreter
    if (hr == S_OK)
    {
        m_pRubyScript->AddRef();
        VALUE bridge = CRubyScript::CreateWin32OLE(new CBridgeDispatch(this));
        m_asr = rb_class_new_instance(1, &bridge, CRubyScript::s_asrClass);
        rb_gc_register_address(&m_asr);
        VARIANT v;
        VariantInit(&v);
        m_pPassedObject = &v;
        rb_funcall(m_asr, rb_intern("self_to_variant"), 0);
        m_pAsr = (v.vt == (VT_DISPATCH | VT_BYREF)) ? *v.ppdispVal : v.pdispVal;
    }
    return hr;
}

void CRubyize::FinalRelease()
{
    if (m_pAsr)
    {
        m_pAsr->Release();
    }
    rb_gc_unregister_address(&m_asr);
    if (m_pRubyScript)
    {
        m_pRubyScript->Close();
        m_pRubyScript->Release();
    }
}

HRESULT STDMETHODCALLTYPE CRubyize::get_Version( 
            /* [retval][out] */ BSTR __RPC_FAR *pVersion)
{
    if (!pVersion) return E_POINTER;
    *pVersion = SysAllocString(RUBYIZE_VERSION);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE CRubyize::get_RubyVersion( 
            /* [retval][out] */ BSTR __RPC_FAR *pVersion)
{
    if (!pVersion) return E_POINTER;

    VALUE v = rb_funcall(m_asr, rb_intern("ruby_version"), 0);
    USES_CONVERSION;
    *pVersion = SysAllocString(A2W(StringValueCStr(v)));
    return S_OK;
}

HRESULT STDMETHODCALLTYPE CRubyize::rubyize( 
            /* [in] */ VARIANT val,
            /* [retval][out] */ VARIANT __RPC_FAR *pObj)
{
    if (!pObj) return E_POINTER;
    VariantInit(pObj);

    volatile VALUE variant = CRubyScript::CreateVariant(val);
    VARIANT v;
    VariantInit(&v);
    m_pPassedObject = &v;
    VALUE obj = rb_funcall(variant, rb_intern("value"), 0);
    pObj->vt = VT_DISPATCH;
    pObj->pdispVal = CreateDispatch(obj);
    return S_OK;
}
        
HRESULT STDMETHODCALLTYPE CRubyize::erubyize( 
            /* [in] */ BSTR script,
            /* [retval][out] */ VARIANT __RPC_FAR *pObj)
{
    if (!pObj) return E_POINTER;
    VariantInit(pObj);

    USES_CONVERSION;
    int len = SysStringLen(script);
    LPSTR psz = new char[len * 2 + 1];
    size_t m = WideCharToMultiByte(GetACP(), 0, script ? script : L"", len, psz, (int)len * 2 + 1, NULL, NULL);
    volatile VALUE vscript = rb_str_new(psz, m);
    delete[] psz;
    volatile VALUE fn = rb_str_new_cstr("(rubyize)");
    volatile VALUE vret = rb_funcall(m_asr, rb_intern("instance_eval"), 3, vscript, fn, LONG2FIX(0));
    pObj->vt = VT_DISPATCH;
    pObj->pdispVal = CreateDispatch(vret);
    return S_OK;
}

IDispatch* CRubyize::CreateDispatch(VALUE val)
{
    VARIANT v;
    VariantInit(&v);
    m_pPassedObject = &v;
    VALUE proxy = rb_funcall(m_asr, rb_intern("to_proxy"), 1, val);
    return new CScriptObject(this, proxy, (v.vt == (VT_DISPATCH | VT_BYREF)) ? *v.ppdispVal : v.pdispVal);
}

STDMETHODIMP CRubyize::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IRubyize
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}
