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

#ifndef EVENTSINK_HEADER
#define EVENTSINK_HEADER

#include "ScriptObject.h"

typedef std::map<std::wstring, DISPID> DispMap;
typedef std::map<std::wstring, DISPID>::iterator DispMapIter;
typedef std::map<DISPID, DISPID> IvkMap;
typedef std::map<DISPID, DISPID>::iterator IvkMapIter;

class CEventSink : public IDispatchEx
{
public:
	CEventSink(CRubyScript*);
	~CEventSink();
	HRESULT Advise(IDispatch*);
	HRESULT Advise(IDispatch*, LPOLESTR);
	HRESULT Unadvise();
	HRESULT ResolveEvent(LPCOLESTR pstrEventName, VALUE handler, VALUE methodid);

	STDMETHOD(QueryInterface)(REFIID iid, void ** ppvObject)
	{
		if (IsEqualIID(iid, IID_IUnknown) 
			|| IsEqualIID(iid, IID_IDispatchEx) 
			|| IsEqualIID(iid, IID_IDispatch) 
			|| IsEqualIID(iid, m_iidEvent))
		{
			*ppvObject = this;
			AddRef();
			return S_OK;
		}
		return E_NOINTERFACE;
	}

	STDMETHOD_(ULONG,AddRef)()
	{
		return InterlockedIncrement(&m_lRefCount);
	}
	STDMETHOD_(ULONG,Release)()
	{
		long l = InterlockedDecrement(&m_lRefCount);
		if (!l)
		{
			delete this;
		}
		return l;
	}
	STDMETHOD(GetTypeInfoCount)(unsigned int FAR *iTInfo)
	{
		return E_NOTIMPL;
	}
	STDMETHOD(GetTypeInfo)(unsigned int iTInfo, LCID lcid, ITypeInfo FAR* FAR* ppTInfo)
	{
		return E_NOTIMPL;
	}
	STDMETHOD(GetIDsOfNames)(
		REFIID riid,
		LPOLESTR* rgszNames,
		unsigned int cNames,
		LCID lcid,
		DISPID FAR* rgDispId);
	STDMETHOD(Invoke)(
		DISPID dispIdMember,
		REFIID riid,
		LCID lcid,
		WORD wFlags,
		DISPPARAMS FAR* pDispParams,
		VARIANT FAR* pVarResult,
		EXCEPINFO FAR* pExcepInfo,
		unsigned int FAR* puArgErr);

	HRESULT STDMETHODCALLTYPE GetDispID( 
            /* [in] */ BSTR bstrName,
            /* [in] */ DWORD grfdex,
            /* [out] */ DISPID __RPC_FAR *pid);
        
	HRESULT STDMETHODCALLTYPE InvokeEx( 
            /* [in] */ DISPID id,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [in] */ DISPPARAMS __RPC_FAR *pdp,
            /* [out] */ VARIANT __RPC_FAR *pvarRes,
            /* [out] */ EXCEPINFO __RPC_FAR *pei,
            /* [unique][in] */ IServiceProvider __RPC_FAR *pspCaller);
        
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
	HRESULT GetEventNames(IID&, ITypeInfo*);
	IID m_iidEvent;
	DispMap m_mapDisp;
	IvkMap m_mapIvk;
	CRubyScript* m_pEngine;
        CScriptObject* m_pHandler;
	IDispatch* m_pDisp;
	DWORD m_dwCookie;
	bool m_fDone;
	long m_lRefCount;
};

#endif
