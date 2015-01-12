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

#pragma once
#include "resource.h"

#include "RScript22_i.h"
#include "giplip.h"

using namespace ATL;

class CItemDisp;
typedef std::map<std::wstring, CItemDisp*> ItemMap;
typedef std::map<std::wstring, CItemDisp*>::iterator ItemMapIter;
class CRubyScript;
class CScriptlet
{
public:
    CScriptlet(LPCOLESTR code, LPCOLESTR item, LPCOLESTR subitem, LPCOLESTR event, ULONG startline);
    HRESULT Add(CRubyScript*);
private:
    std::wstring m_code;
    std::wstring m_item;
    std::wstring m_subitem;
    std::wstring m_event;
    ULONG m_startline;
};
typedef std::list<CScriptlet> ScriptletList;
typedef std::list<CScriptlet>::iterator ScriptletListIter;
class CEventSink;
typedef std::map<std::wstring, CEventSink*> EventMap;
typedef std::map<std::wstring, CEventSink*>::iterator EventMapIter;

// CRubyScript
#include "BridgeDispatch.h"

class ATL_NO_VTABLE CRubyScript :
    public CComObjectRootEx<CComMultiThreadModel>,
    public CComCoClass<CRubyScript, &CLSID_RubyScript>,
    public IActiveScript,
    public IActiveScriptParse,
    public IActiveScriptGarbageCollector,
    public IActiveScriptParseProcedure,
    public IServiceProvider,
    public CObjectStore
{
    friend class CRubyize;
    friend class CScriptObject;
public:
    CRubyScript();

    static _ATL_REGMAP_ENTRY RegEntries[];
    static HRESULT WINAPI UpdateRegistry(BOOL bRegister);

BEGIN_COM_MAP(CRubyScript)
    COM_INTERFACE_ENTRY(IActiveScriptParse)
    COM_INTERFACE_ENTRY(IActiveScript)
    COM_INTERFACE_ENTRY(IActiveScriptGarbageCollector)
    COM_INTERFACE_ENTRY(IActiveScriptParseProcedure)
    COM_INTERFACE_ENTRY(IServiceProvider)
    COM_INTERFACE_ENTRY_AGGREGATE(IID_IMarshal, m_pUnkMarshaler.p)
END_COM_MAP()

    DECLARE_PROTECT_FINAL_CONSTRUCT()
    DECLARE_GET_CONTROLLING_UNKNOWN()

    HRESULT FinalConstruct()
    {
	return CoCreateFreeThreadedMarshaler(
	    GetControllingUnknown(), &m_pUnkMarshaler.p);
    }

    void FinalRelease()
    {
        Close();
	m_pUnkMarshaler.Release();
    }

    CComPtr<IUnknown> m_pUnkMarshaler;

    void EnterScript();
    void LeaveScript();

    inline void EnterScript(IActiveScriptSite* pSite)
    {
        m_threadState = SCRIPTTHREADSTATE_RUNNING;
        pSite->OnEnterScript();
    }
    inline void LeaveScript(IActiveScriptSite* pSite)
    {
        m_threadState = SCRIPTTHREADSTATE_NOTINSCRIPT;
        pSite->OnLeaveScript();
    }
    inline void CopyPersistent(int n, std::string& s) { m_nStartLinePersistent = n; m_strScriptPersistent = s; }
    HRESULT EvalString(int line, int len, LPCSTR script, VARIANT* result = NULL, EXCEPINFO FAR* pExcepInfo = NULL, DWORD dwFlags = 0);

    inline void SetPassedObject(VARIANT& v)
    {
        _ASSERT(m_pPassedObject);
        VariantCopy(m_pPassedObject, &v);
    }

public:
    // IServiceProvider
    HRESULT STDMETHODCALLTYPE QueryService(
	    REFGUID guidService,
	    REFIID riid,
	    void **ppv);


    HRESULT STDMETHODCALLTYPE SetScriptSite( 
        /* [in] */ __RPC__in_opt IActiveScriptSite *pass);
        
    HRESULT STDMETHODCALLTYPE GetScriptSite( 
        /* [in] */ __RPC__in REFIID riid,
        /* [iid_is][out] */ __RPC__deref_out_opt void **ppvObject);
        
    HRESULT STDMETHODCALLTYPE SetScriptState( 
        /* [in] */ SCRIPTSTATE ss);
        
    HRESULT STDMETHODCALLTYPE GetScriptState( 
        /* [out] */ __RPC__out SCRIPTSTATE *pssState);
        
    HRESULT STDMETHODCALLTYPE Close( void);
        
    HRESULT STDMETHODCALLTYPE AddNamedItem( 
        /* [in] */ __RPC__in LPCOLESTR pstrName,
        /* [in] */ DWORD dwFlags);
        
    HRESULT STDMETHODCALLTYPE AddTypeLib( 
        /* [in] */ __RPC__in REFGUID rguidTypeLib,
        /* [in] */ DWORD dwMajor,
        /* [in] */ DWORD dwMinor,
        /* [in] */ DWORD dwFlags);
        
    HRESULT STDMETHODCALLTYPE GetScriptDispatch( 
        /* [in] */ __RPC__in LPCOLESTR pstrItemName,
        /* [out] */ __RPC__deref_out_opt IDispatch **ppdisp);
        
    HRESULT STDMETHODCALLTYPE GetCurrentScriptThreadID( 
        /* [out] */ __RPC__out SCRIPTTHREADID *pstidThread);
        
    HRESULT STDMETHODCALLTYPE GetScriptThreadID( 
        /* [in] */ DWORD dwWin32ThreadId,
        /* [out] */ __RPC__out SCRIPTTHREADID *pstidThread);
        
    HRESULT STDMETHODCALLTYPE GetScriptThreadState( 
        /* [in] */ SCRIPTTHREADID stidThread,
        /* [out] */ __RPC__out SCRIPTTHREADSTATE *pstsState);
        
    HRESULT STDMETHODCALLTYPE InterruptScriptThread( 
        /* [in] */ SCRIPTTHREADID stidThread,
        /* [in] */ __RPC__in const EXCEPINFO *pexcepinfo,
        /* [in] */ DWORD dwFlags);
        
    HRESULT STDMETHODCALLTYPE Clone( 
        /* [out] */ __RPC__deref_out_opt IActiveScript **ppscript);
#if defined(_WIN64)
	    // IActiveScriptParse64
    HRESULT STDMETHODCALLTYPE InitNew( void);
        
    HRESULT STDMETHODCALLTYPE AddScriptlet( 
        /* [in] */ __RPC__in LPCOLESTR pstrDefaultName,
        /* [in] */ __RPC__in LPCOLESTR pstrCode,
        /* [in] */ __RPC__in LPCOLESTR pstrItemName,
        /* [in] */ __RPC__in LPCOLESTR pstrSubItemName,
        /* [in] */ __RPC__in LPCOLESTR pstrEventName,
        /* [in] */ __RPC__in LPCOLESTR pstrDelimiter,
        /* [in] */ DWORDLONG dwSourceContextCookie,
        /* [in] */ ULONG ulStartingLineNumber,
        /* [in] */ DWORD dwFlags,
        /* [out] */ __RPC__deref_out_opt BSTR *pbstrName,
        /* [out] */ __RPC__out EXCEPINFO *pexcepinfo);
        
    HRESULT STDMETHODCALLTYPE ParseScriptText( 
        /* [in] */ __RPC__in LPCOLESTR pstrCode,
        /* [in] */ __RPC__in LPCOLESTR pstrItemName,
        /* [in] */ __RPC__in_opt IUnknown *punkContext,
        /* [in] */ __RPC__in LPCOLESTR pstrDelimiter,
        /* [in] */ DWORDLONG dwSourceContextCookie,
        /* [in] */ ULONG ulStartingLineNumber,
        /* [in] */ DWORD dwFlags,
        /* [out] */ __RPC__out VARIANT *pvarResult,
        /* [out] */ __RPC__out EXCEPINFO *pexcepinfo);

    // IActiveScriptParseProcedure64
    HRESULT STDMETHODCALLTYPE ParseProcedureText( 
        /* [in] */ __RPC__in LPCOLESTR pstrCode,
        /* [in] */ __RPC__in LPCOLESTR pstrFormalParams,
        /* [in] */ __RPC__in LPCOLESTR pstrProcedureName,
        /* [in] */ __RPC__in LPCOLESTR pstrItemName,
        /* [in] */ __RPC__in_opt IUnknown *punkContext,
        /* [in] */ __RPC__in LPCOLESTR pstrDelimiter,
        /* [in] */ DWORDLONG dwSourceContextCookie,
        /* [in] */ ULONG ulStartingLineNumber,
        /* [in] */ DWORD dwFlags,
        /* [out] */ __RPC__deref_out_opt IDispatch **ppdisp);
#else
	    // IActiveScriptParse32
    HRESULT STDMETHODCALLTYPE InitNew( void);

    HRESULT STDMETHODCALLTYPE AddScriptlet( 
        /* [in] */ __RPC__in LPCOLESTR pstrDefaultName,
        /* [in] */ __RPC__in LPCOLESTR pstrCode,
        /* [in] */ __RPC__in LPCOLESTR pstrItemName,
        /* [in] */ __RPC__in LPCOLESTR pstrSubItemName,
        /* [in] */ __RPC__in LPCOLESTR pstrEventName,
        /* [in] */ __RPC__in LPCOLESTR pstrDelimiter,
        /* [in] */ DWORD dwSourceContextCookie,
        /* [in] */ ULONG ulStartingLineNumber,
        /* [in] */ DWORD dwFlags,
        /* [out] */ __RPC__deref_out_opt BSTR *pbstrName,
        /* [out] */ __RPC__out EXCEPINFO *pexcepinfo);
        
    HRESULT STDMETHODCALLTYPE ParseScriptText( 
        /* [in] */ __RPC__in LPCOLESTR pstrCode,
        /* [in] */ __RPC__in LPCOLESTR pstrItemName,
        /* [in] */ __RPC__in_opt IUnknown *punkContext,
        /* [in] */ __RPC__in LPCOLESTR pstrDelimiter,
        /* [in] */ DWORD dwSourceContextCookie,
        /* [in] */ ULONG ulStartingLineNumber,
        /* [in] */ DWORD dwFlags,
        /* [out] */ __RPC__out VARIANT *pvarResult,
        /* [out] */ __RPC__out EXCEPINFO *pexcepinfo);

    // IActiveScriptParseProcedure32
    HRESULT STDMETHODCALLTYPE ParseProcedureText( 
        /* [in] */ __RPC__in LPCOLESTR pstrCode,
        /* [in] */ __RPC__in LPCOLESTR pstrFormalParams,
        /* [in] */ __RPC__in LPCOLESTR pstrProcedureName,
        /* [in] */ __RPC__in LPCOLESTR pstrItemName,
        /* [in] */ __RPC__in_opt IUnknown *punkContext,
        /* [in] */ __RPC__in LPCOLESTR pstrDelimiter,
        /* [in] */ DWORD dwSourceContextCookie,
        /* [in] */ ULONG ulStartingLineNumber,
        /* [in] */ DWORD dwFlags,
        /* [out] */ __RPC__deref_out_opt IDispatch **ppdisp);
#endif // _WIN64

    // IActiveScriptGarbageCollector
    HRESULT STDMETHODCALLTYPE CollectGarbage(
        SCRIPTGCTYPE scriptgctype
    );
private:
    DWORD m_dwThreadID;
    HANDLE m_hThread;
    SCRIPTSTATE m_state;
    SCRIPTTHREADSTATE m_threadState;
    int m_nStartLinePersistent;
    std::string m_strScriptPersistent;
    std::wstring m_strGlobalObjectName;
    static VALUE s_asrModule;
    static VALUE s_asrClass;
    static VALUE s_asrProxy;
    VALUE m_asr;
    GIP(IActiveScriptSite) m_pSite;
    ItemMap m_mapItem;
    ScriptletList m_listScriptlets;
    EventMap m_mapEvent;
    HRESULT Connect();
    void ConnectToEvents();
    void Disconnect(bool fSinkOnly = false);
    void CopyNamedItem(ItemMap&);
    void AddNamedItemToScript(LPCOLESTR, DWORD);
    void BindNamedItem();
    void UnbindNamedItem();
    void AddConst(LPCOLESTR, VARIANT&);
    VARIANT* m_pPassedObject;
    HRESULT LoadTypeLib(REFGUID rguidTypeLib, DWORD dwMajor, DWORD dwMinor, ITypeLib** ppResult);
    void CreateEventSource(IActiveScriptSite*, ItemMap::value_type&, bool);
    void CreateActiveScriptRuby();
    static VALUE CreateWin32OLE(IDispatch* pdisp);
    static VALUE CreateVariant(VARIANT&);
    static void CreateVariant(VALUE, VARIANT*);
    HRESULT OnScriptError(LPCSTR script);
public:
    IDispatch* CreateGlobalDispatch();
    // create proxied dispatch
    IDispatch* CreateDispatch(VALUE obj);
    // create event only dispatch
    IDispatch* CreateEventDispatch(VALUE);
    // create asr dispatch, take global dispatch
    IDispatch* CreateAsrDispatch(IDispatch* pglobal);
    // safe rb_funcall
    static VALUE safe_funcall(VALUE args);
};

//OBJECT_ENTRY_AUTO(__uuidof(RubyScript), CRubyScript)
