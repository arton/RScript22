/*
 *  Copyright(c) 2000,2003 arton
 *
 *  You may distribute under the terms of either the GNU General Public
 *  License
 *
 */

#include "stdafx.h"
#include "RubyScript.h"
#include "eventsink.h"

CEventCode::CEventCode(int n, LPCOLESTR s) :
    m_nStartLine(n)
{
    USES_CONVERSION;
    m_strCode = W2A(s);
}

CEventCode::CEventCode(const CEventCode& c) :
    m_nStartLine(c.m_nStartLine), m_strCode(c.m_strCode)
{
}

CEventSink::CEventSink(CRubyScript* pEngine)
    : m_fDone(false), m_pEngine(pEngine), m_pDisp(NULL), m_lRefCount(0)
{
}

CEventSink::~CEventSink()
{
    if (m_pDisp)
    {
        Unadvise();
        m_pDisp->Release();
        m_pDisp = NULL;
    }
}

HRESULT CEventSink::Advise(IDispatch* pDisp)
{
    if (m_fDone) return S_OK;

    m_pDisp = pDisp;
    pDisp->AddRef();
    IProvideClassInfo2* pPCI;
    HRESULT hr = pDisp->QueryInterface(IID_IProvideClassInfo2, (void**)&pPCI);
    if (hr == S_OK)
    {
	if (SUCCEEDED(pPCI->GetGUID(GUIDKIND_DEFAULT_SOURCE_DISP_IID,
		&m_iidEvent)))
	{
	    ITypeInfo* pTypeInfo;
	    if (pPCI->GetClassInfo(&pTypeInfo) == S_OK)
	    {
		GetEventNames(m_iidEvent, pTypeInfo);
		pTypeInfo->Release();
	    }

	    IConnectionPointContainer* pConnPt;
	    hr = pDisp->QueryInterface(IID_IConnectionPointContainer, (void**)&pConnPt);
	    if (hr == S_OK)
	    {
		IConnectionPoint* pConn;
		hr = pConnPt->FindConnectionPoint(m_iidEvent, &pConn);
		if (hr == S_OK)
		{
		    hr = pConn->Advise(this, &m_dwCookie);
		    pConn->Release();
		}
		pConnPt->Release();
	    }
	}
	pPCI->Release();
    }
    if (hr == S_OK)
    {
	m_fDone = true;
    }
    return hr;
}

HRESULT CEventSink::GetEventNames(IID& iid, ITypeInfo* pInfo)
{
    ITypeLib* pTypeLib;
    UINT uIndex;
    HRESULT hr = pInfo->GetContainingTypeLib(&pTypeLib, &uIndex);
    if (hr == S_OK)
    {
	ITypeInfo* pTypeInfo;
	hr = pTypeLib->GetTypeInfoOfGuid(iid, &pTypeInfo);
	if (hr == S_OK)
	{
	    TYPEATTR* pTypeAttr;
	    if (pTypeInfo->GetTypeAttr(&pTypeAttr) == S_OK)
	    {
	        int cFunc = pTypeAttr->cFuncs;
	        for (int i = 0; i < cFunc; i++)
	        {
		    FUNCDESC* pFuncDesc;
		    if (pTypeInfo->GetFuncDesc(i, &pFuncDesc) == S_OK)
		    {
			BSTR bstr;
			bstr = NULL;
			if (pTypeInfo->GetDocumentation(pFuncDesc->memid, &bstr, NULL, NULL, NULL) == S_OK)
			{
			    ATLTRACE(_T(" Event:%ls, id=%d\n"), bstr, pFuncDesc->memid);
			    m_mapDisp.insert(DispMap::value_type(bstr, pFuncDesc->memid));
			    SysFreeString(bstr);
			}
			pTypeInfo->ReleaseFuncDesc(pFuncDesc);
		    }
	        }
	        pTypeInfo->ReleaseTypeAttr(pTypeAttr);
	    }
	    pTypeInfo->Release();
	}
	pTypeLib->Release();
    }
    return S_OK;
}

HRESULT CEventSink::Advise(IDispatch* pDisp, LPOLESTR pstrName)
{
    if (m_fDone) return S_OK;

    HRESULT hr = E_UNEXPECTED;
    IDispatch* pSubDisp = NULL;
    DISPID dispid = -1;
    DISPPARAMS dispparm = { NULL, NULL, 0, 0 };
    VARIANT vResult;
    VariantInit(&vResult);

    IDispatchEx* pEx;
    if (pDisp->QueryInterface(IID_IDispatchEx, (void**)&pEx) == S_OK)
    {
	BSTR bstr = SysAllocString(pstrName);
	hr = pEx->GetDispID(bstr, fdexNameCaseSensitive, &dispid);
	SysFreeString(bstr);
	if (hr == S_OK)
	{
	    hr = pEx->InvokeEx(dispid, 0, DISPATCH_PROPERTYGET, &dispparm, &vResult, NULL, m_pEngine);
	}
	pEx->Release();
    }
    else if (pDisp->GetIDsOfNames(IID_NULL, &pstrName, 1, 0, &dispid) == S_OK)
    {
	hr = pDisp->Invoke(dispid, IID_NULL, 0, DISPATCH_PROPERTYGET, &dispparm, &vResult, NULL, NULL);
    }

    if (hr == S_OK)
    {
	hr = Advise(vResult.pdispVal);
	VariantClear(&vResult);
    }
    return hr;
}

HRESULT CEventSink::Unadvise()
{
    if (!m_pDisp) return S_OK;
    IConnectionPointContainer* pConnPt;
    HRESULT hr = m_pDisp->QueryInterface(IID_IConnectionPointContainer, (void**)&pConnPt);
    if (hr == S_OK)
    {
	IConnectionPoint* pConn;
	hr = pConnPt->FindConnectionPoint(m_iidEvent, &pConn);
	if (hr == S_OK)
	{
	    hr = pConn->Unadvise(m_dwCookie);
	    pConn->Release();
	}
	pConnPt->Release();
    }
    return hr;
}

HRESULT CEventSink::ResolveEvent(LPCOLESTR pstrEventName, int nLineStart, LPCOLESTR pstrCode)
{
    DispMapIter it = m_mapDisp.find(pstrEventName);
    if (it == m_mapDisp.end()) return E_UNEXPECTED;

    m_mapIvk[(*it).second] = CEventCode(nLineStart, pstrCode);
    return S_OK;
}

STDMETHODIMP CEventSink::GetIDsOfNames(
    REFIID riid,
    LPOLESTR* rgszNames,
    unsigned int cNames,
    LCID lcid,
    DISPID FAR* rgDispId)
{
    if (!rgDispId) return E_POINTER;
    DispMapIter it = m_mapDisp.find(*rgszNames);
    if (it == m_mapDisp.end())
    {
        *rgDispId = DISPID_UNKNOWN;
        return DISP_E_UNKNOWNNAME; 
    }
    *rgDispId = (*it).second;
    for (unsigned int i = 1; i < cNames; i++)
    {
        *(rgDispId + i) = DISPID_UNKNOWN;
    }
    return S_OK;
}

STDMETHODIMP CEventSink::Invoke(
    DISPID dispIdMember,
    REFIID riid,
    LCID lcid,
    WORD wFlags,
    DISPPARAMS FAR* pDispParams,
    VARIANT FAR* pVarResult,
    EXCEPINFO FAR* pExcepInfo,
    unsigned int FAR* puArgErr)
{
    ATLTRACE(_T("EventSink::Invoke dispid=%d\n"), dispIdMember);
    VARIANTARG* pArg = pDispParams->rgvarg;
    IvkMapIter it = m_mapIvk.find(dispIdMember);
    if (it == m_mapIvk.end()) return DISP_E_MEMBERNOTFOUND;

    std::string& str = (*it).second.GetCode();
    CRubyScript* pEngine = m_pEngine;
    m_pEngine->EnterScript();
    HRESULT hr = m_pEngine->EvalString((*it).second.GetStartLine(), str.length(), str.c_str(), pVarResult, pExcepInfo);
    if (pEngine == m_pEngine)
    {
	// If Engine was destroyed for some reason, this pointer makes access violation.
	m_pEngine->LeaveScript();
    }
    return hr;
}

HRESULT STDMETHODCALLTYPE CEventSink::GetDispID( 
            /* [in] */ BSTR bstrName,
            /* [in] */ DWORD grfdex,
            /* [out] */ DISPID __RPC_FAR *pid)
{
    return GetIDsOfNames(IID_NULL, &bstrName, 1, LOCALE_SYSTEM_DEFAULT, pid);
}
        
HRESULT STDMETHODCALLTYPE CEventSink::InvokeEx( 
    /* [in] */ DISPID id,
    /* [in] */ LCID lcid,
    /* [in] */ WORD wFlags,
    /* [in] */ DISPPARAMS __RPC_FAR *pdp,
    /* [out] */ VARIANT __RPC_FAR *pvarRes,
    /* [out] */ EXCEPINFO __RPC_FAR *pei,
    /* [unique][in] */ IServiceProvider __RPC_FAR *pspCaller)
{
//    if (pspCaller)
//	m_pEngine->PushServiceProvider(pspCaller);
    HRESULT hr = Invoke(id, IID_NULL, lcid, wFlags, pdp, pvarRes, pei, NULL);
//    if (pspCaller)
//	m_pEngine->PopServiceProvider();
    return hr;
}

