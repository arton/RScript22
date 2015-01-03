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

#if !defined(SCRIPT_OBJECT_H)
#define SCRIPT_OBJECT_H

class CScriptObject : public IDispatchEx
{
public:
    // pobjdispatch should be AddRefed
    CScriptObject(VALUE v, IDispatch* pobjdispatch, IDispatch* pdisp = NULL) :
        m_object(v),
        m_objectDispatch(pobjdispatch),
        m_pDispatch(pdisp),
        m_lCount(1),
        m_cDispIds(0),
        m_sizeDispIds(16)
    {
        if (!IMMEDIATE_P(v))
        {
            rb_gc_register_address(&v);
        }
        m_DispIds = new DISPID[m_sizeDispIds];
    }

    ~CScriptObject()
    {
        m_objectDispatch->Release();
        if (m_pDispatch)
        {
            m_pDispatch->Release();
        }
        if (!IMMEDIATE_P(m_object))
        {
            rb_gc_unregister_address(&m_object);
        }
        delete[] m_DispIds;
    }

    HRESULT  STDMETHODCALLTYPE QueryInterface(
		const IID & riid,  
		void **ppvObj)
    {
	if (!ppvObj) return E_POINTER;

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch) || IsEqualIID(riid, IID_IDispatchEx))
	{
            m_lCount++;
            *ppvObj = this;
            return S_OK;
	}
	return E_NOINTERFACE;
    }

    ULONG  STDMETHODCALLTYPE AddRef()
    {
        return ++m_lCount;
    }

    ULONG  STDMETHODCALLTYPE Release()
    {
        m_lCount--;
        if (m_lCount != 0)
        {
            return m_lCount;
        }
        delete this;
        return 0;
    }

    // IDispatch
    virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount( 
        /* [out] */ UINT __RPC_FAR *pctinfo)
    {
	ATLTRACENOTIMPL(_T("CRubyObject::GetTypeInfoCount"));
    }
        
    virtual HRESULT STDMETHODCALLTYPE GetTypeInfo( 
        /* [in] */ UINT iTInfo,
        /* [in] */ LCID lcid,
        /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo)
    {
        ATLTRACENOTIMPL(_T("CRubyObject::GetTypeInfo"));
    }
        
    virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames( 
        /* [in] */ REFIID riid,
        /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
        /* [in] */ UINT cNames,
        /* [in] */ LCID lcid,
        /* [size_is][out] */ DISPID __RPC_FAR *rgDispId)
    {
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
        
    virtual /* [local] */ HRESULT STDMETHODCALLTYPE Invoke( 
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
        DISPPARAMS params = { NULL, pDispParams->rgdispidNamedArgs, 1 + pDispParams->cArgs, pDispParams->cNamedArgs };
        params.rgvarg = (VARIANT*)_alloca(sizeof(VARIANT) * params.cArgs);
        memcpy(params.rgvarg, pDispParams->rgvarg, sizeof(VARIANTARG) * pDispParams->cArgs);
        VariantInit(&params.rgvarg[pDispParams->cArgs]);
        params.rgvarg[pDispParams->cArgs].vt = VT_I4;
        params.rgvarg[pDispParams->cArgs].lVal = dispIdMember;
        return m_objectDispatch->Invoke(DISPID_VALUE, riid, lcid, DISPATCH_METHOD, &params, pVarResult, pExcepInfo, puArgErr);
    }

    HRESULT STDMETHODCALLTYPE GetDispID( 
        /* [in] */ BSTR bstrName,
        /* [in] */ DWORD grfdex,
        /* [out] */ DISPID __RPC_FAR *pid)
    {
	    HRESULT hr = GetIDsOfNames(IID_NULL, &bstrName, 1, LOCALE_SYSTEM_DEFAULT, pid);
	    return hr;
    }
        
    HRESULT STDMETHODCALLTYPE InvokeEx( 
        /* [in] */ DISPID id,
        /* [in] */ LCID lcid,
        /* [in] */ WORD wFlags,
        /* [in] */ DISPPARAMS __RPC_FAR *pdp,
        /* [out] */ VARIANT __RPC_FAR *pvarRes,
        /* [out] */ EXCEPINFO __RPC_FAR *pei,
        /* [unique][in] */ IServiceProvider __RPC_FAR *pspCaller)
    {
	HRESULT hr = Invoke(id, IID_NULL, lcid, wFlags, pdp, pvarRes, pei, NULL);
	return hr;
    }
        
    HRESULT STDMETHODCALLTYPE DeleteMemberByName( 
        /* [in] */ BSTR bstr,
        /* [in] */ DWORD grfdex)
    {
        ATLTRACENOTIMPL(_T("DeleteMemberByName"));
    }

    HRESULT STDMETHODCALLTYPE DeleteMemberByDispID( 
        /* [in] */ DISPID id)
    {
	ATLTRACENOTIMPL(_T("DeleteMemberByDispID"));
    }
        
    HRESULT STDMETHODCALLTYPE GetMemberProperties( 
        /* [in] */ DISPID id,
        /* [in] */ DWORD grfdexFetch,
        /* [out] */ DWORD __RPC_FAR *pgrfdex)
    {
	ATLTRACENOTIMPL(_T("GetMemberProperties"));
    }
        
    HRESULT STDMETHODCALLTYPE GetMemberName( 
        /* [in] */ DISPID id,
        /* [out] */ BSTR __RPC_FAR *pbstrName)
    {
	ATLTRACENOTIMPL(_T("GetMemberName"));
    }
        
    HRESULT STDMETHODCALLTYPE GetNextDispID( 
        /* [in] */ DWORD grfdex,
        /* [in] */ DISPID id,
        /* [out] */ DISPID __RPC_FAR *pid)
    {
	ATLTRACENOTIMPL(_T("GetNextDispID"));
    }
        
    HRESULT STDMETHODCALLTYPE GetNameSpaceParent( 
        /* [out] */ IUnknown __RPC_FAR *__RPC_FAR *ppunk)
    {
	ATLTRACENOTIMPL(_T("GetNameSpaceParent"));
    }

private:
    VALUE m_object;
    IDispatch* m_objectDispatch;
    IDispatch* m_pDispatch;
    ULONG m_lCount;
    DISPID* m_DispIds;
    int m_cDispIds;
    int m_sizeDispIds;
};

#endif
