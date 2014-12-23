/*
 *  Copyright(c) 2009, 2014 arton
 *
 *  You may distribute under the terms of either the GNU General Public
 *  License
 *
 *  $Date$
 */

#include "stdafx.h"
#include "RubyScript.h"
#include "Rubyize.h"

#if defined(RUBY_2_2)
#define RUBYIZE_VERSION L"2.2.0"
#else
#define RUBYIZE_VERSION L"2.1.0"
#endif

#define RUBYIZE NULL
/////////////////////////////////////////////////////////////////////////////
// CRubyize
HRESULT CRubyize::FinalConstruct()
{
    HRESULT hr = CComObject<CRubyScript>::CreateInstance(&m_pRubyScript);
    if (hr == S_OK)
    {
        m_pRubyScript->AddRef();
        m_pRubyScript->SetScriptSite(this);
        CComPtr<IActiveScriptParse> pParse;
        HRESULT hr = m_pRubyScript->QueryInterface(IID_IActiveScriptParse, (void**)&pParse);
        if (hr == S_OK)
        {
            m_pRubyScript->SetScriptState(SCRIPTSTATE_CONNECTED);
        }
    }
    return hr;
}

void CRubyize::FinalRelease()
{
    if (m_pRubyScript)
    {
        m_pRubyScript->Close();
        m_pRubyScript->Release();
        m_pRubyScript = NULL;
    }
}

HRESULT STDMETHODCALLTYPE CRubyize::GetLCID( 
            /* [out] */ LCID __RPC_FAR *plcid)
{
    return E_NOTIMPL;
}
        
HRESULT STDMETHODCALLTYPE CRubyize::GetItemInfo( 
            /* [in] */ LPCOLESTR pstrName,
            /* [in] */ DWORD dwReturnMask,
            /* [out] */ IUnknown __RPC_FAR *__RPC_FAR *ppiunkItem,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppti)
{
    return TYPE_E_ELEMENTNOTFOUND;
}
        
HRESULT STDMETHODCALLTYPE CRubyize::GetDocVersionString( 
            /* [out] */ BSTR __RPC_FAR *pbstrVersion)
{
      *pbstrVersion = SysAllocString(RUBYIZE_VERSION);
      return S_OK;
}
        
HRESULT STDMETHODCALLTYPE CRubyize::OnScriptTerminate( 
            /* [in] */ const VARIANT __RPC_FAR *pvarResult,
            /* [in] */ const EXCEPINFO __RPC_FAR *pexcepinfo)
{
    return S_OK;
}
        
HRESULT STDMETHODCALLTYPE CRubyize::OnStateChange( 
            /* [in] */ SCRIPTSTATE ssScriptState)
{
    return S_OK;
}
        
HRESULT STDMETHODCALLTYPE CRubyize::OnScriptError( 
            /* [in] */ IActiveScriptError __RPC_FAR *pscripterror)
{
    return S_OK;
}
        
HRESULT STDMETHODCALLTYPE CRubyize::OnEnterScript( void)
{
    return S_OK;
}
        
HRESULT STDMETHODCALLTYPE CRubyize::OnLeaveScript( void)
{
    return S_OK;
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
    VARIANT v;
    VariantInit(&v);
    HRESULT hr = Call(L"ruby_version", 0, NULL, &v);
    if (hr == S_OK)
    {
        if (v.vt == VT_BSTR)
        {
            *pVersion = SysAllocString(v.bstrVal);
        }
        else if (v.vt == (VT_BSTR | VT_BYREF))
        {
            *pVersion = SysAllocString(*v.pbstrVal);
        }
        else
        {
            *pVersion = NULL;
        }
        VariantClear(&v);
    }
    return hr;
}

HRESULT CRubyize::Call(LPCOLESTR method, int cargs, VARIANT* args, VARIANT* pResult)
{
    CComPtr<IDispatch> pdisp;
    HRESULT hr = m_pRubyScript->GetScriptDispatch(RUBYIZE, &pdisp);
    if (hr == S_OK)
    {
        DISPID dispid[1];
        hr = pdisp->GetIDsOfNames(IID_NULL, const_cast<LPOLESTR*>(&method), 1, LOCALE_SYSTEM_DEFAULT, dispid);
        if (hr == S_OK)
        {
            DISPPARAMS params = { args, NULL, cargs, 0 };
            unsigned int aerr;
            hr = pdisp->Invoke(dispid[0], IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD,
                &params, pResult, NULL, &aerr);
        }
    }
    return hr;    
}

HRESULT STDMETHODCALLTYPE CRubyize::rubyize( 
            /* [in] */ VARIANT val,
            /* [retval][out] */ VARIANT __RPC_FAR *pObj)
{
    if (!pObj) return E_POINTER;
    return Call(L"rubyize", 1, &val, pObj);
}
        
HRESULT STDMETHODCALLTYPE CRubyize::erubyize( 
            /* [in] */ BSTR script,
            /* [retval][out] */ VARIANT __RPC_FAR *pObj)
{
    if (!pObj) return E_POINTER;
    VARIANT v;
    v.vt = VT_BSTR;
    v.bstrVal = script;
    return Call(L"erubyize", 1, &v, pObj);
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
