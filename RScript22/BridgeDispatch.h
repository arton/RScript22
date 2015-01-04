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
#if !defined(BRIDGE_DISPATCH_H)
#define BRIDGE_DISPATCH

class CObjectStore
{
public:
    // set variant for receiving an object that ASR passed
    virtual void SetPassedObject(VARIANT&) = 0;
    // create proxy dispatch
    virtual IDispatch* CreateDispatch(VALUE) = 0;
};

class CBridgeDispatch : public IDispatch 
{
public:
    CBridgeDispatch(CObjectStore* p)
        : m_lCount(1), m_pObjectStore(p)
    {
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(
	const IID & riid,  
	void **ppvObj)
    {
	if (!ppvObj) return E_POINTER;
	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch))
	{
            m_lCount++;
            *ppvObj = this;
            return S_OK;
	}
	return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef()
    {
	return ++m_lCount;
    }

    ULONG STDMETHODCALLTYPE Release()
    {
	m_lCount--;
	if (m_lCount <= 0)
	{
	    delete this;
	    return 0;
	}
	return m_lCount;
    }

    // IDispatch
    HRESULT STDMETHODCALLTYPE GetTypeInfoCount( 
	/* [out] */ UINT __RPC_FAR *pctinfo)
    {
	ATLTRACENOTIMPL(_T("CBridgeDispatch::GetTypeInfoCount"));
    }
        
    HRESULT STDMETHODCALLTYPE GetTypeInfo( 
	/* [in] */ UINT iTInfo,
	/* [in] */ LCID lcid,
	/* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo)
    {
	ATLTRACENOTIMPL(_T("CBridgeDispatch::GetTypeInfo"));
    }

    static const DISPID DISPID_PASSOBJECT = 32;

    HRESULT STDMETHODCALLTYPE GetIDsOfNames( 
	/* [in] */ REFIID riid,
	/* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
	/* [in] */ UINT cNames,
	/* [in] */ LCID lcid,
	/* [size_is][out] */ DISPID __RPC_FAR *rgDispId)
    {
        static struct Method {
            LPCOLESTR name;
            DISPID id;
        } dispids[] = {
            { L"PassObject", DISPID_PASSOBJECT },
        };
        HRESULT hr = S_OK;
        for (UINT i = 0; i < cNames; i++) 
        {
            *(rgDispId + i) = DISPID_UNKNOWN;
            for (int n = 0; n < COUNT_OF(dispids); n++)
            {
                if (_wcsicmp(*(rgszNames + i), dispids[n].name) == 0)
                {
                    *(rgDispId + i) = dispids[n].id;
                }
            }
            if (*(rgDispId + i) == DISPID_UNKNOWN) hr = DISP_E_UNKNOWNNAME;
        }
        return hr;
    }

    HRESULT STDMETHODCALLTYPE Invoke( 
	/* [in] */ DISPID dispIdMember,
	/* [in] */ REFIID riid,
	/* [in] */ LCID lcid,
	/* [in] */ WORD wFlags,
	/* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
	/* [out] */ VARIANT __RPC_FAR *pVarResult,
	/* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
	/* [out] */ UINT __RPC_FAR *puArgErr)
    {
	if (!pDispParams) return E_POINTER;
	if (pVarResult)
	{
            VariantInit(pVarResult);
	}

        VARIANT* pv;
        switch (dispIdMember)
        {
        case DISPID_PASSOBJECT:
            if (pDispParams->cArgs != 1)
                return DISP_E_BADPARAMCOUNT;
            pv = (pDispParams->rgvarg[0].vt == (VT_VARIANT | VT_BYREF)) ? pDispParams->rgvarg[0].pvarVal : &pDispParams->rgvarg[0];
            m_pObjectStore->SetPassedObject(*pv);
            break;
        default:
            return DISP_E_PARAMNOTFOUND;
        }
        return S_OK;
    }

private:
    ULONG m_lCount;
    CObjectStore* m_pObjectStore;
};

#endif
