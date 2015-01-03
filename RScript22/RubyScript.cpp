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
#include "dllmain.h"
#include "ItemDisp.h"
#include "eventsink.h"
#include "ScriptObject.h"

// Win32OLE
struct OLE_DATA {
    IDispatch *pDispatch;
};

struct OLE_VARIANT_DATA {
    VARIANT realvar;
    VARIANT var;
};

static VALUE s_win32ole(Qnil);
static VALUE s_win32ole_variant(Qnil);

VALUE CRubyScript::CreateWin32OLE(IDispatch* pdisp)
{
    volatile VALUE obj = rb_obj_alloc(s_win32ole);
    OLE_DATA* pole = reinterpret_cast<OLE_DATA*>(DATA_PTR(obj));
    pole->pDispatch = pdisp;
    return obj;
}

VALUE CRubyScript::CreateVariant(VARIANT& var)
{
    volatile VALUE val = rb_obj_alloc(s_win32ole_variant);
    OLE_VARIANT_DATA* pvar = reinterpret_cast<OLE_VARIANT_DATA*>(DATA_PTR(val));
    VariantCopy(&pvar->realvar, &var);
    VariantCopy(&pvar->var, &var);
    return val;
}

void CRubyScript::CreateVariant(VALUE v, VARIANT* pvar)
{
    VALUE var = rb_class_new_instance(1, &v, s_win32ole_variant);
    OLE_VARIANT_DATA* pvariant = reinterpret_cast<OLE_VARIANT_DATA*>(DATA_PTR(var));
    VariantCopy(pvar, &pvariant->var);
}

IDispatch* CRubyScript::CreateDispatch(VALUE obj)
{
    VARIANT v;
    VariantInit(&v);
    m_pPassedObject = &v;
    rb_funcall(m_asr, rb_intern("to_variant"), 1, obj);
    return (v.vt == (VT_DISPATCH | VT_BYREF)) ? *v.ppdispVal : v.pdispVal;
}

IDispatch* CRubyScript::CreateGlobalDispatch()
{
    IDispatch* pglobal = NULL;
    ItemMapIter it = (m_strGlobalObjectName.size() > 0) ? m_mapItem.find(m_strGlobalObjectName.c_str()) : m_mapItem.end();
    if (it == m_mapItem.end())
    {
        for (it = m_mapItem.begin(); it != m_mapItem.end(); it++)
        {
            if ((*it).second->GetFlag() & SCRIPTITEM_GLOBALMEMBERS)
            {
                break;
            }
        }
    }
    if (it != m_mapItem.end())
    {
        pglobal = (*it).second->GetDispatch();
    }
    return new CScriptObject(this, m_asr, CreateDispatch(m_asr), pglobal);
}

void CRubyScript::CreateActiveScriptRuby()
{
    if (m_asr == Qnil)
    {
        VALUE v = CreateWin32OLE(new CBridgeDispatch(this));
        m_asr = rb_class_new_instance(1, &v, s_asrClass);
        rb_gc_register_address(&m_asr);
    }
}

// CRubyScript

template<class T, const IID* piid> T* GetInterface(GlobalInterfacePointer<T, piid>& gip)
{
    T* p = NULL;
    HRESULT hr = gip.Localize(&p);
    return p;
}

#define GET_POINTER(T, gip) T* p##T = GetInterface<T, &IID_##T>(gip);
#define RELEASE_POINTER(T) p##T->Release();

static bool bRubyInitialized(false);
static char* asr_argv[] = {"ActiveScriptRuby", "-e", ";", NULL};
VALUE CRubyScript::s_asrModule(Qnil);
VALUE CRubyScript::s_asrClass(Qnil);
VALUE CRubyScript::s_asrProxy(Qnil);

__declspec( thread ) bool s_threadInitialized = false;

CRubyScript::CRubyScript()
    : m_state(SCRIPTSTATE_UNINITIALIZED),
      m_threadState(SCRIPTTHREADSTATE_NOTINSCRIPT),
      m_pUnkMarshaler(NULL),
      m_asr(Qnil),
      m_dwThreadID(GetCurrentThreadId()),
      m_nStartLinePersistent(0),
      m_pPassedObject(NULL)
{
    ATLTRACE(_T("CRubyScript on thread:%d\n"), GetCurrentThreadId());
    if (!bRubyInitialized)
    {
        bRubyInitialized = true;
        s_threadInitialized = true;
        int dummyargc(1);
        char* dummyargv[] = {"dummy", NULL };
        char** pargv;
        ruby_sysinit(&dummyargc, &pargv);
        ruby_init();
        ruby_options(3, asr_argv);
        rb_require("activescriptruby");
        s_asrModule = rb_const_get(rb_cObject, rb_intern("ActiveScriptRuby"));
        s_asrClass = rb_const_get(s_asrModule, rb_intern("GlobalProxy"));
        s_asrProxy = rb_const_get(s_asrModule, rb_intern("AsrProxy"));
        s_win32ole = rb_const_get(rb_cObject, rb_intern("WIN32OLE"));
        s_win32ole_variant = rb_const_get(rb_cObject, rb_intern("WIN32OLE_VARIANT"));
    }
    else if (!s_threadInitialized)
    {
        s_threadInitialized = true;
        rb_thread_create(NULL, NULL);
    }
    _ASSERT(!NIL_P(s_win32ole));
    _ASSERT(!NIL_P(s_asrClass));
    ATLTRACE(_T("asrClass:%08X\n"), s_asrClass);
}

void CRubyScript::EnterScript()
{
    m_hThread = GetCurrentThread();
    GET_POINTER(IActiveScriptSite, m_pSite)
    EnterScript(pIActiveScriptSite);
    RELEASE_POINTER(IActiveScriptSite)
}

void CRubyScript::LeaveScript()
{
    GET_POINTER(IActiveScriptSite, m_pSite)
    LeaveScript(pIActiveScriptSite);
    RELEASE_POINTER(IActiveScriptSite)
}

HRESULT STDMETHODCALLTYPE CRubyScript::QueryService(
    REFGUID guidService,
    REFIID riid,
    void **ppv)
{
    if (!ppv) return E_POINTER;
    HRESULT hr = E_NOINTERFACE;
    *ppv = NULL;
    if (InlineIsEqualGUID(guidService, SID_GetCaller)
	    || InlineIsEqualGUID(guidService, IID_IActiveScriptSite))
    {
	GET_POINTER(IActiveScriptSite, m_pSite);
	if (pIActiveScriptSite)
	{
            hr = pIActiveScriptSite->QueryInterface(riid, ppv);
        }
        RELEASE_POINTER(IActiveScriptSite)
    }
    return hr;
}

HRESULT STDMETHODCALLTYPE CRubyScript::SetScriptSite( 
    /* [in] */ __RPC__in_opt IActiveScriptSite *pass)
{
    if (!pass) return E_POINTER;
    if (m_pSite.IsOK()) return E_UNEXPECTED;

    pass->AddRef();
    m_pSite.Globalize(pass);
    m_dwThreadID = GetCurrentThreadId();
    return S_OK;
}
        
HRESULT STDMETHODCALLTYPE CRubyScript::GetScriptSite( 
    /* [in] */ __RPC__in REFIID riid,
    /* [iid_is][out] */ __RPC__deref_out_opt void **ppvObject)
{
    if (!ppvObject) return E_POINTER;
    *ppvObject = NULL;
    if (!m_pSite.IsOK()) return S_FALSE;
    GET_POINTER(IActiveScriptSite, m_pSite)
    HRESULT hr = pIActiveScriptSite->QueryInterface(riid, ppvObject);
    RELEASE_POINTER(IActiveScriptSite)
    return hr;
}
        
HRESULT STDMETHODCALLTYPE CRubyScript::SetScriptState( 
    /* [in] */ SCRIPTSTATE ss)
{
    if (!m_pSite.IsOK()) return E_UNEXPECTED;
    if (m_state != ss)
    {
	switch (ss)
	{
	case SCRIPTSTATE_STARTED:
            CreateActiveScriptRuby();
	    BindNamedItem();
	    break;
	case SCRIPTSTATE_CONNECTED:
            CreateActiveScriptRuby();
	    Connect();
	    BindNamedItem();
	    ConnectToEvents();
	    break;
	case SCRIPTSTATE_INITIALIZED:
            CreateActiveScriptRuby();
	    Disconnect(true);	// Disconnect all Sink
	    UnbindNamedItem();
	    break;
	case SCRIPTSTATE_DISCONNECTED:
	    Disconnect(true);	// Disconnect all Sink
	    break;
	case SCRIPTSTATE_UNINITIALIZED:
	    return E_FAIL;
	    break;
	default:
	    break;
	}

	m_state = ss;
	if (ss != SCRIPTSTATE_UNINITIALIZED)
	{
            if (m_pSite.IsOK())
            {
	        GET_POINTER(IActiveScriptSite, m_pSite)
                if (pIActiveScriptSite)
                {
	            pIActiveScriptSite->OnStateChange(ss);
                    RELEASE_POINTER(IActiveScriptSite)
                }
                else
                {
                    m_pSite.Unglobalize();
                    ATLTRACE(_T("no scriptsite in SetScriptState %d\n"), ss);
                }
            }
	}
	return S_OK;
    }
    return S_FALSE;
}
        
HRESULT STDMETHODCALLTYPE CRubyScript::GetScriptState( 
    /* [out] */ __RPC__out SCRIPTSTATE *pssState)
{
    if (!pssState) return E_POINTER;
    *pssState = m_state;
    return S_OK;
}
        
HRESULT STDMETHODCALLTYPE CRubyScript::Close( void)
{
    if (!m_pSite.IsOK()) return S_FALSE;

    m_state = SCRIPTSTATE_CLOSED;
    Disconnect();
    m_pSite.Unglobalize();
    return S_OK;
}
        
HRESULT STDMETHODCALLTYPE CRubyScript::AddNamedItem( 
    /* [in] */ __RPC__in LPCOLESTR pstrName,
    /* [in] */ DWORD dwFlags)
{
    if (!pstrName) return E_POINTER;

    ATLTRACE(_T("AddNameItem %ls flag=%08X\n"), pstrName, dwFlags);

    if (!(dwFlags & SCRIPTITEM_ISVISIBLE)) return S_OK;
    if (dwFlags & SCRIPTITEM_GLOBALMEMBERS)
        m_strGlobalObjectName = pstrName;

    ItemMapIter it = m_mapItem.find(pstrName);
    if (it != m_mapItem.end())
    {
        EnterScript();
        USES_CONVERSION;
        volatile VALUE itemName = rb_str_new_cstr(W2A((*it).first.c_str()));
        rb_funcall(m_asr, rb_intern("remove_item"), 1, itemName);
        if ((*it).second)
        {
            (*it).second->Empty();
            delete (*it).second;
        }
        LeaveScript();
    }
    m_mapItem.insert(ItemMap::value_type(pstrName, new CItemDisp(dwFlags)));
    if (m_state == SCRIPTSTATE_STARTED || m_state == SCRIPTSTATE_CONNECTED)
    {
        AddNamedItemToScript(pstrName, dwFlags);
    }
    return S_OK;
}

void CRubyScript::AddConst(LPCOLESTR pname, VARIANT& var)
{
    volatile VALUE val = CreateVariant(var);
    USES_CONVERSION;
    volatile VALUE vname = rb_str_new_cstr(W2A(pname));
    rb_funcall(m_asr, rb_intern("add_constant"), 2, vname, val);
}
        
HRESULT STDMETHODCALLTYPE CRubyScript::AddTypeLib( 
    /* [in] */ __RPC__in REFGUID rguidTypeLib,
    /* [in] */ DWORD dwMajor,
    /* [in] */ DWORD dwMinor,
    /* [in] */ DWORD dwFlags)
{
    ATLTRACE(_T("AddTypeLib\n"));
    ITypeLib* pLib;
    HRESULT hr = LoadTypeLib(rguidTypeLib, dwMajor, dwMinor, &pLib);
    if (hr != S_OK)
    {
        return hr;
    }

    int count = pLib->GetTypeInfoCount();
    for (int i = 0; i < count; i++)
    {
        ITypeInfo* pInfo;
        if (pLib->GetTypeInfo(i, &pInfo) != S_OK)
        {
            continue;
        }
        TYPEATTR* pattr;
        if (pInfo->GetTypeAttr(&pattr) != S_OK)
        {
            pInfo->Release();
            continue;
        }
        VARDESC* pvardesc;
        BSTR strName;
        UINT pc;
        for (int iv = 0; iv < pattr->cVars; iv++)
        {
            if (pInfo->GetVarDesc(iv, &pvardesc) != S_OK)
            {
                continue;
            }
            strName = NULL;
            if(pvardesc->varkind == VAR_CONST &&
                !(pvardesc->wVarFlags & (VARFLAG_FHIDDEN | VARFLAG_FRESTRICTED | VARFLAG_FNONBROWSABLE))) 
            {
                if (pInfo->GetNames(pvardesc->memid, &strName, 1, &pc) != S_OK || pc == 0 || !strName)
                {
	            continue;
                }
                AddConst(strName, *pvardesc->lpvarValue);
                SysFreeString(strName);
            }
            pInfo->ReleaseVarDesc(pvardesc);
        }
        pInfo->ReleaseTypeAttr(pattr);
        pInfo->Release();
    }
    return S_OK;
}
        
HRESULT STDMETHODCALLTYPE CRubyScript::GetScriptDispatch( 
    /* [in] */ __RPC__in LPCOLESTR pstrItemName,
    /* [out] */ __RPC__deref_out_opt IDispatch **ppdisp)
{
    if (!ppdisp) return E_POINTER;
    HRESULT hr = S_FALSE;
    *ppdisp = NULL;
    ATLTRACE(_T("GetScriptDispatch for %ls\n"), (pstrItemName) ? pstrItemName : L"GLOBALNAMESPACE");
    if (!pstrItemName || m_strGlobalObjectName == pstrItemName)
    {
        *ppdisp = CreateGlobalDispatch();
        return S_OK;
    }

    GET_POINTER(IActiveScriptSite, m_pSite)
    ItemMapIter it = m_mapItem.find(pstrItemName);
    if (it != m_mapItem.end())
    {
        IDispatch* pdisp = (*it).second->GetDispatch(pIActiveScriptSite, pstrItemName, m_dwThreadID == GetCurrentThreadId());
        hr = pdisp->QueryInterface(IID_IDispatch, (void**)ppdisp);
    }
    RELEASE_POINTER(IActiveScriptSite)
    return hr;
}
        
HRESULT STDMETHODCALLTYPE CRubyScript::GetCurrentScriptThreadID( 
    /* [out] */ __RPC__out SCRIPTTHREADID *pstidThread)
{
    if (!pstidThread) return E_POINTER;
    *pstidThread = m_dwThreadID;
    return S_OK;
}
        
HRESULT STDMETHODCALLTYPE CRubyScript::GetScriptThreadID( 
    /* [in] */ DWORD dwWin32ThreadId,
    /* [out] */ __RPC__out SCRIPTTHREADID *pstidThread)
{
    if (!pstidThread) return E_POINTER;
    *pstidThread = m_dwThreadID;
    return S_OK;
}
        
HRESULT STDMETHODCALLTYPE CRubyScript::GetScriptThreadState( 
    /* [in] */ SCRIPTTHREADID stidThread,
    /* [out] */ __RPC__out SCRIPTTHREADSTATE *pstsState)
{
    if (!pstsState) return E_POINTER;
    *pstsState = m_threadState;
    return S_OK;
}

static void interruptThread()
{
	ATLTRACE(_T("interruptThread()\n"));
	rb_raise(rb_eInterrupt, "interrput");
}
        
HRESULT STDMETHODCALLTYPE CRubyScript::InterruptScriptThread( 
    /* [in] */ SCRIPTTHREADID stidThread,
    /* [in] */ __RPC__in const EXCEPINFO *pexcepinfo,
    /* [in] */ DWORD dwFlags)
{
    if (m_dwThreadID == GetCurrentThreadId())
    {
        return E_UNEXPECTED;
    }

    SuspendThread(m_hThread);
    CONTEXT con;
    ZeroMemory(&con, sizeof(CONTEXT));
    con.ContextFlags = CONTEXT_CONTROL;
    GetThreadContext(m_hThread, &con);
    FARPROC proc = (FARPROC)&interruptThread;
    con.ContextFlags = CONTEXT_CONTROL;
#if !defined(_M_X64)
    con.Eip = (long)proc;
#else
    con.Rip = (DWORD64)proc;
#endif
    SetThreadContext(m_hThread, &con);
    ResumeThread(m_hThread);
    return S_OK;
}
        
HRESULT STDMETHODCALLTYPE CRubyScript::Clone( 
    /* [out] */ __RPC__deref_out_opt IActiveScript **ppscript)
{
    ATLTRACE(_T("Clone\n"));
    if (!ppscript) return E_POINTER;

    HRESULT hr = S_OK;

    CComObject<CRubyScript>* p;
    hr = CComObject<CRubyScript>::CreateInstance(&p);
    if (hr == S_OK)
    {
        hr = p->QueryInterface(IID_IActiveScript, (void**)ppscript);
        if (hr != S_OK)
        {
            delete p;
        }
        else
        {
            p->CopyNamedItem(m_mapItem);
            p->CopyPersistent(m_nStartLinePersistent, m_strScriptPersistent);
        }
    }
    return (hr == S_OK) ? S_OK : E_UNEXPECTED;
}

HRESULT STDMETHODCALLTYPE CRubyScript::InitNew( void)
{
    Disconnect();
    if (m_pSite.IsOK())
    {
        GET_POINTER(IActiveScriptSite, m_pSite)
        pIActiveScriptSite->OnStateChange(m_state = SCRIPTSTATE_INITIALIZED);
        RELEASE_POINTER(IActiveScriptSite)
    }
    m_threadState = SCRIPTTHREADSTATE_NOTINSCRIPT;
    return S_OK;
}

CScriptlet::CScriptlet(LPCOLESTR code, LPCOLESTR item, LPCOLESTR subitem, LPCOLESTR event, ULONG startline) :
    m_startline(startline)
{
    if (code) m_code = code;
    if (item) m_item = item;
    if (subitem) m_subitem = subitem;
    if (event) m_event = event;
}

HRESULT CScriptlet::Add(CRubyScript* pscr)
{
    BSTR bstr = NULL;
    EXCEPINFO excep;
    return pscr->AddScriptlet(L"", m_code.c_str(), m_item.c_str(), m_subitem.c_str(), m_event.c_str(),
		L";", SCRIPTTEXT_ISVISIBLE , m_startline, 0, &bstr, &excep);
}

HRESULT STDMETHODCALLTYPE CRubyScript::AddScriptlet( 
#if defined(_WIN64)
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
    /* [out] */ __RPC__out EXCEPINFO *pexcepinfo)
#else
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
    /* [out] */ __RPC__out EXCEPINFO *pexcepinfo)
#endif            
{
    if (!(dwFlags | SCRIPTTEXT_ISVISIBLE)) return S_OK;
    if (m_state != SCRIPTSTATE_CONNECTED)
    {
        m_listScriptlets.push_back(CScriptlet(pstrCode, pstrItemName, pstrSubItemName, pstrEventName, ulStartingLineNumber));
        return S_OK;
    }
    ATLTRACE(_T("AddScriptLet default=%ls, item=%ls, subitem=%ls, event=%ls\n"), (pstrDefaultName) ? pstrDefaultName : L"(null)", pstrItemName, pstrSubItemName, pstrEventName);
    if (pstrItemName && pstrEventName && pstrCode)
    {
        USES_CONVERSION;
	HRESULT hr = E_UNEXPECTED;
        volatile VALUE vname = rb_str_new_cstr(W2A((!pstrSubItemName || wcslen(pstrSubItemName) == 0) ? pstrItemName : pstrSubItemName));
        volatile VALUE eventname = rb_str_new_cstr(W2A(pstrEventName));
        size_t len = wcslen(pstrCode);
        LPSTR pScript = new char[len * 2 + 1];
        size_t m = WideCharToMultiByte(GetACP(), 0, pstrCode, (int)len, pScript, (int)len * 2 + 1, NULL, NULL);
        volatile VALUE vscript = rb_str_new(pScript, m);
        volatile VALUE handler = rb_funcall(m_asr, rb_intern("create_event_handler"), 4, vname, eventname, vscript, LONG2FIX(ulStartingLineNumber));
        VALUE methodid = rb_funcall(handler, rb_intern("get_method_id"), 1, eventname);
        delete[] pScript;

        ItemMapIter it = m_mapItem.end();

        if (pstrSubItemName)
        {
    	    it = m_mapItem.find(pstrSubItemName);
            if (it == m_mapItem.end())
            {
                AddNamedItem(pstrSubItemName, SCRIPTTEXT_ISVISIBLE | SCRIPTITEM_ISSOURCE);
                it = m_mapItem.find(pstrSubItemName);
                GET_POINTER(IActiveScriptSite, m_pSite)
                CreateEventSource(pIActiveScriptSite, *it, (m_dwThreadID == GetCurrentThreadId()));
                RELEASE_POINTER(IActiveScriptSite)
            }
        }
        else
        {
	    it = m_mapItem.find(pstrItemName);
        }
	if (it != m_mapItem.end())
	{
	    EventMapIter itev = m_mapEvent.end();
	    if (pstrSubItemName)
	    {
		itev = m_mapEvent.find(pstrSubItemName);
	    }
	    else
	    {
		itev = m_mapEvent.find(pstrItemName);
	    }

	    if (itev != m_mapEvent.end())
	    {
		hr = (*itev).second->ResolveEvent(pstrEventName, handler, methodid);
	    }
	    if (hr == S_OK)
	    {
		*pbstrName = SysAllocString(pstrDefaultName);
	    }
	    else
	    {
		USES_CONVERSION;

		CComBSTR bstr(pstrSubItemName);
		bstr += L".";
		bstr += pstrEventName;
		*pbstrName = bstr.Copy();
                hr = S_OK;
	    }
	}
	return hr;
    }
    return E_INVALIDARG;
}

HRESULT STDMETHODCALLTYPE CRubyScript::ParseScriptText( 
#if defined(_WIN64)            
    /* [in] */ __RPC__in LPCOLESTR pstrCode,
    /* [in] */ __RPC__in LPCOLESTR pstrItemName,
    /* [in] */ __RPC__in_opt IUnknown *punkContext,
    /* [in] */ __RPC__in LPCOLESTR pstrDelimiter,
    /* [in] */ DWORDLONG dwSourceContextCookie,
    /* [in] */ ULONG ulStartingLineNumber,
    /* [in] */ DWORD dwFlags,
    /* [out] */ __RPC__out VARIANT *pvarResult,
    /* [out] */ __RPC__out EXCEPINFO *pexcepinfo)
#else
    /* [in] */ __RPC__in LPCOLESTR pstrCode,
    /* [in] */ __RPC__in LPCOLESTR pstrItemName,
    /* [in] */ __RPC__in_opt IUnknown *punkContext,
    /* [in] */ __RPC__in LPCOLESTR pstrDelimiter,
    /* [in] */ DWORD dwSourceContextCookie,
    /* [in] */ ULONG ulStartingLineNumber,
    /* [in] */ DWORD dwFlags,
    /* [out] */ __RPC__out VARIANT *pvarResult,
    /* [out] */ __RPC__out EXCEPINFO *pexcepinfo)
#endif
{
    size_t len = wcslen(pstrCode);
    LPSTR psz = new char[len * 2 + 1];
    size_t cb = WideCharToMultiByte(GetACP(), 0, pstrCode, (int)len, psz, (int)len * 2 + 1, NULL, NULL);
    *(psz + cb) = '\0';

    if ((dwFlags & (SCRIPTTEXT_ISPERSISTENT | SCRIPTTEXT_ISVISIBLE)) == (SCRIPTTEXT_ISPERSISTENT | SCRIPTTEXT_ISVISIBLE)
	    && m_state == SCRIPTSTATE_UNINITIALIZED)
    {
	m_nStartLinePersistent = ulStartingLineNumber;
	m_strScriptPersistent = psz;
	delete[] psz;
	return S_OK;
    }

    HRESULT hr = E_UNEXPECTED;
    try 
    {
	if (pvarResult)
	{
	    VariantInit(pvarResult);
	}
	Connect();
        EnterScript();
	hr = EvalString(ulStartingLineNumber, cb, psz, pvarResult, NULL, dwFlags);
	LeaveScript();
    }
    catch (...)
    {
    }
    delete[] psz;

    return hr;
}

// IActiveScriptParseProcedure64
HRESULT STDMETHODCALLTYPE CRubyScript::ParseProcedureText( 
#if defined(_WIN64)
    /* [in] */ __RPC__in LPCOLESTR pstrCode,
    /* [in] */ __RPC__in LPCOLESTR pstrFormalParams,
    /* [in] */ __RPC__in LPCOLESTR pstrProcedureName,
    /* [in] */ __RPC__in LPCOLESTR pstrItemName,
    /* [in] */ __RPC__in_opt IUnknown *punkContext,
    /* [in] */ __RPC__in LPCOLESTR pstrDelimiter,
    /* [in] */ DWORDLONG dwSourceContextCookie,
    /* [in] */ ULONG ulStartingLineNumber,
    /* [in] */ DWORD dwFlags,
    /* [out] */ __RPC__deref_out_opt IDispatch **ppdisp)
#else
    /* [in] */ __RPC__in LPCOLESTR pstrCode,
    /* [in] */ __RPC__in LPCOLESTR pstrFormalParams,
    /* [in] */ __RPC__in LPCOLESTR pstrProcedureName,
    /* [in] */ __RPC__in LPCOLESTR pstrItemName,
    /* [in] */ __RPC__in_opt IUnknown *punkContext,
    /* [in] */ __RPC__in LPCOLESTR pstrDelimiter,
    /* [in] */ DWORD dwSourceContextCookie,
    /* [in] */ ULONG ulStartingLineNumber,
    /* [in] */ DWORD dwFlags,
    /* [out] */ __RPC__deref_out_opt IDispatch **ppdisp)
#endif
{
    static int seqcnt(0);
    if (!(dwFlags | SCRIPTTEXT_ISVISIBLE)) return S_OK;
    if (!pstrCode || !ppdisp) return E_POINTER;
    *ppdisp = NULL;
    VARIANT v;
    VariantInit(&v);
    EXCEPINFO excep;
    memset(&excep, 0, sizeof(EXCEPINFO));

    USES_CONVERSION;
    LPCSTR iname = (pstrItemName) ? W2A(pstrItemName) : W2A(m_strGlobalObjectName.c_str());
    size_t len = wcslen(pstrCode);
    LPSTR psz = new char[len * 2 + 1];
    size_t m = WideCharToMultiByte(GetACP(), 0, pstrCode, (int)len, psz, (int)len * 2 + 1, NULL, NULL);
    volatile VALUE vscript = rb_str_new(psz, m);
    volatile VALUE vname = rb_str_new_cstr(iname);
    volatile VALUE handler = rb_funcall(m_asr, rb_intern("create_event_proc"), 3, vname, vscript, LONG2FIX(ulStartingLineNumber));
    delete[] psz;
    *ppdisp = CreateDispatch(handler);
    return S_OK;
}

HRESULT CRubyScript::CollectGarbage(SCRIPTGCTYPE scriptgctype)
{
    if (scriptgctype == SCRIPTGCTYPE_EXHAUSTIVE)
    {
        rb_gc_start();
    }
    else
    {
        rb_gc_start();
    }
    return S_OK;
}

HRESULT CRubyScript::Connect()
{
    if (m_state == SCRIPTSTATE_CONNECTED) return S_OK;

    if (m_mapItem.size() <= 0) return S_OK;

    bool fSameApt = (m_dwThreadID == GetCurrentThreadId());
    GET_POINTER(IActiveScriptSite, m_pSite)
    for (ItemMapIter it = m_mapItem.begin(); it != m_mapItem.end(); it++)
    {
	if ((*it).second->IsOK() == false)
            AddNamedItem((*it).first.c_str(), (*it).second->GetFlag());

	if ((*it).second->IsSource())
	{
            CreateEventSource(pIActiveScriptSite, *it, fSameApt);
	}
    }
    HRESULT hr = pIActiveScriptSite->OnStateChange(SCRIPTSTATE_CONNECTED);
    RELEASE_POINTER(IActiveScriptSite)
    return hr;
}

void CRubyScript::CreateEventSource(IActiveScriptSite* pSite, ItemMap::value_type& val, bool fSameApt)
{
    IDispatch* pDisp = val.second->GetDispatch(pSite, const_cast<LPOLESTR>(val.first.c_str()), fSameApt);
    if (pDisp)
    {
	EventMapIter itev = m_mapEvent.find(val.first.c_str());
	if (itev == m_mapEvent.end())
	{
	    itev = (m_mapEvent.insert(EventMap::value_type(val.first, new CEventSink(this)))).first;
	}
	(*itev).second->AddRef();
	(*itev).second->Advise(pDisp);
	pDisp->Release();
    }
}

void CRubyScript::ConnectToEvents()
{
    for (ScriptletListIter it = m_listScriptlets.begin(); it != m_listScriptlets.end(); it++)
    {
        it->Add(this);
    }
    m_listScriptlets.clear();
}

void CRubyScript::Disconnect(bool fSinkOnly)
{
    for (EventMapIter itev = m_mapEvent.begin(); itev != m_mapEvent.end(); itev++)
    {
        if ((*itev).second)
        {
            (*itev).second->Release();
        }
    }
    m_mapEvent.clear();
    if (fSinkOnly) return;	// return If triggered by SetScriptState -> DISCONNECTED

    if (m_asr != Qnil)
    {
        rb_funcall(m_asr, rb_intern("remove_items"), 0);
        rb_gc_unregister_address(&m_asr);
        m_asr = Qnil;
    }
    for (ItemMapIter it = m_mapItem.begin(); it != m_mapItem.end(); it++)
    {
        if ((*it).second)
        {
            (*it).second->Empty();
            delete (*it).second;
        }
    }
    m_mapItem.clear();
    CreateActiveScriptRuby();
}

void CRubyScript::CopyNamedItem(ItemMap& map)
{
    for (ItemMapIter it = map.begin(); it != map.end(); it++)
    {
        m_mapItem.insert(ItemMap::value_type((*it).first, new CItemDisp((*it).second->GetFlag())));
    }
}

void CRubyScript::BindNamedItem()
{
    for (ItemMapIter it = m_mapItem.begin(); it != m_mapItem.end(); it++)
    {
        if ((*it).second->IsOK() == false)
            AddNamedItemToScript((*it).first.c_str(), (*it).second->GetFlag());
    }
}

void CRubyScript::UnbindNamedItem()
{
    rb_funcall(m_asr, rb_intern("remove_items"), 0);
    for (ItemMapIter it = m_mapItem.begin(); it != m_mapItem.end(); it++)
    {
        if ((*it).second)
            (*it).second->Empty();
    }
}

void CRubyScript::AddNamedItemToScript(LPCOLESTR pstrName, DWORD dwFlags)
{
    ItemMapIter it = m_mapItem.find(pstrName);
    if (it == m_mapItem.end()) return;

    // Check NamedItem
    ATLTRACE(_T("AddNamedItem %ls\n"), pstrName);
    GET_POINTER(IActiveScriptSite, m_pSite)
    IDispatch* pDisp = (*it).second->GetDispatch(pIActiveScriptSite, pstrName, GetCurrentThreadId() == m_dwThreadID);
    if (!pDisp) 
    {
        RELEASE_POINTER(IActiveScriptSite)
        return;
    }
    volatile VALUE vdisp = CreateWin32OLE(pDisp);
    USES_CONVERSION;
    volatile VALUE itemName = rb_str_new_cstr(W2A(pstrName));
    EnterScript(pIActiveScriptSite);
    rb_funcall(m_asr, rb_intern("add_named_item"), 3, itemName, vdisp, (dwFlags & SCRIPTITEM_GLOBALMEMBERS) ? Qtrue : Qfalse);
    LeaveScript(pIActiveScriptSite);
    RELEASE_POINTER(IActiveScriptSite)
}

HRESULT CRubyScript::LoadTypeLib(
            /* [in] */ REFGUID rguidTypeLib,
            /* [in] */ DWORD dwMajor,
            /* [in] */ DWORD dwMinor,
            /* [out]*/ ITypeLib** ppResult)
{
    OLECHAR strGuid[64];
    StringFromGUID2(rguidTypeLib, strGuid, 64);
    WCHAR key[MAX_PATH];
    wsprintf(key, L"TypeLib\\%s\\%d.%d", strGuid, dwMajor, dwMinor);
    HKEY hTypeLib;
    if (RegOpenKeyExW(HKEY_CLASSES_ROOT, key, 0, KEY_READ, &hTypeLib) != ERROR_SUCCESS)
    {
        return E_INVALIDARG;
    }
    bool found(false);
    FILETIME ft;
    WCHAR buff[MAX_PATH];
    for (int i = 0; !found; i++)
    {
	DWORD szbuff = sizeof(buff);
	if (RegEnumKeyExW(hTypeLib, i, buff, &szbuff, NULL, NULL, NULL, &ft) != ERROR_SUCCESS)
	{
            break;
	}
	wcscat(buff, L"\\win32");
	HKEY hWin32;
	if (RegOpenKeyExW(hTypeLib, buff, 0, KEY_READ, &hWin32) != ERROR_SUCCESS)
	{
            break;
	}
	szbuff = COUNT_OF(buff);
	if (RegQueryValueExA(hWin32, NULL, NULL, NULL, reinterpret_cast<LPBYTE>(&buff), &szbuff) == ERROR_SUCCESS)
	{
            found = true;
	}
	RegCloseKey(hWin32);
    }
    RegCloseKey(hTypeLib);
    if (!found)
    {
        return E_INVALIDARG;
    }
    USES_CONVERSION;
    if (::LoadTypeLib(buff, ppResult) != S_OK)
    {
        return TYPE_E_CANTLOADLIBRARY;
    }
    return S_OK;
}

HRESULT CRubyScript::EvalString(int line, int len, LPCSTR script, VARIANT* result, EXCEPINFO FAR* pExcepInfo, DWORD dwFlags)
{
    volatile VALUE vscript = rb_str_new(script, len);
    volatile VALUE fn = rb_str_new_cstr("(asr)");
    VALUE vret = rb_funcall(m_asr, rb_intern("instance_eval"), 3, vscript, fn, LONG2FIX(line));
    if (!result) return S_OK;
    m_pPassedObject = result;
    rb_funcall(m_asr, rb_intern("to_variant"), 1, vret);
    return S_OK;
}
