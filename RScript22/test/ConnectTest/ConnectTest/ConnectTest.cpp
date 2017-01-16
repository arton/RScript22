// ConnectTest.cpp : コンソール アプリケーションのエントリ ポイントを定義します。
//

#include "stdafx.h"

HANDLE hConsole;

class CMyObj : public IDispatch {
public:
	enum {
		DISPID_DO = 1,
	};
	CMyObj() : m_lCount(1)
	{
		CoCreateFreeThreadedMarshaler(this, &m_pUnkMarshaler);
	}
	~CMyObj()
	{
		m_pUnkMarshaler->Release();
	}
	HRESULT  STDMETHODCALLTYPE QueryInterface(
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
		HRESULT hr = m_pUnkMarshaler->QueryInterface(riid, ppvObj);
		return hr;
	}
	ULONG STDMETHODCALLTYPE AddRef()
	{
		return ++m_lCount;
	}
	ULONG STDMETHODCALLTYPE Release()
	{
		--m_lCount;
		if (m_lCount <= 0)
		{
			delete this;
			return 0;
		}
		return m_lCount;
	}
	HRESULT STDMETHODCALLTYPE GetTypeInfoCount( 
		/* [out] */ UINT __RPC_FAR *pctinfo)
	{
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE GetTypeInfo( 
		/* [in] */ UINT iTInfo,
		/* [in] */ LCID lcid,
		/* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo)
	{
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE GetIDsOfNames( 
		/* [in] */ REFIID riid,
		/* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
		/* [in] */ UINT cNames,
		/* [in] */ LCID lcid,
		/* [size_is][out] */ DISPID __RPC_FAR *rgDispId)
	{
		HRESULT hr = S_OK;
		for (UINT i = 0; i < cNames; i++)
		{
			if (_wcsicmp(*(rgszNames + i), L"do") == 0
				|| _wcsicmp(*(rgszNames + i), L"domethod") == 0)
			{
				*(rgDispId + i) = DISPID_DO;
			}
			else
			{
				*(rgDispId + i) = DISPID_UNKNOWN;
				hr = DISP_E_MEMBERNOTFOUND;
			}
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
		HRESULT hr = S_OK;
		if (dispIdMember == DISPID_DO)
		{
			char sz[16];
			VARIANTARG* pv = pDispParams->rgvarg;
			if (pv->vt == (VT_VARIANT | VT_BYREF))
				pv = pv->pvarVal;
			DWORD n = sprintf(sz, "%d\n", pv->iVal);
			WriteConsole(hConsole, sz, n, &n, NULL);
		}
		else
		{
			hr = DISP_E_MEMBERNOTFOUND;
		}
		return hr;
	}

private:
	long m_lCount;
	IUnknown* m_pUnkMarshaler;
};

class MyScriptSite : public IActiveScriptSite
{
public:
	MyScriptSite() : m_lCount(1)
	{
	}
	HRESULT  STDMETHODCALLTYPE QueryInterface(
		const IID & riid,  
		void **ppvObj)
	{
		if (!ppvObj) return E_POINTER;
		if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IActiveScriptSite))
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
		--m_lCount;
		if (m_lCount <= 0)
		{
			delete this;
			return 0;
		}
		return m_lCount;
	}
	HRESULT STDMETHODCALLTYPE GetLCID(LCID*)
	{
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE GetItemInfo(
		LPCOLESTR pstrName,     // address of item name
		DWORD dwReturnMask,     // bit mask for information retrieval
		IUnknown **ppunkItem,   // address of pointer to item's IUnknown
		ITypeInfo **ppTypeInfo  // address of pointer to item's ITypeInfo
		)
	{
		if ((dwReturnMask & SCRIPTINFO_IUNKNOWN)&& _wcsicmp(pstrName, L"MyCustomObj") == 0)
		{
			*ppunkItem = new CMyObj;
			return S_OK;
		}
		return TYPE_E_ELEMENTNOTFOUND;
	}
	HRESULT STDMETHODCALLTYPE GetDocVersionString(
		BSTR *pbstrVersionString  // address of document version string
		)
	{
		*pbstrVersionString = SysAllocString(L"1.0");
		return S_OK;
	}
	HRESULT STDMETHODCALLTYPE OnScriptTerminate(
		const VARIANT *pvarResult,   // address of script results
		const EXCEPINFO *pexcepinfo  // address of structure with exception information
		)
	{
		return S_OK;
	}
	HRESULT STDMETHODCALLTYPE OnStateChange(
		SCRIPTSTATE ssScriptState  // new state of engine
		)
	{
		return S_OK;
	}
	HRESULT STDMETHODCALLTYPE OnScriptError(
		IActiveScriptError *pase  // address of error interface
		)
	{
		BSTR b;
		pase->GetSourceLineText(&b);
		DWORD n;
		char sz[128];
		n = sprintf(sz, "err:%ls\n", b);
		SysFreeString(b);
		WriteConsole(hConsole, sz, n, &n, NULL);
		return S_OK;
	}
	HRESULT STDMETHODCALLTYPE OnEnterScript(void)
	{
		return S_OK;
	}
	HRESULT STDMETHODCALLTYPE OnLeaveScript(void)
	{
		return S_OK;
	}
private:
	long m_lCount;
};

int _tmain(int argc, _TCHAR* argv[])
{
	CoInitialize(NULL);
	hConsole = GetStdHandle(STD_OUTPUT_HANDLE);

	CLSID clsid;
	HRESULT hr = ::CLSIDFromProgID(L"RubyScript.2.4", &clsid );
	IActiveScript* pAS;
	hr = ::CoCreateInstance( clsid, 0, CLSCTX_ALL, IID_IActiveScript, reinterpret_cast<void **>( &pAS ));

	MyScriptSite* pASS = new MyScriptSite;
	hr = pAS->SetScriptSite( pASS );

	IActiveScriptParse* pASP;
	hr = pAS->QueryInterface(IID_IActiveScriptParse, ( void ** )&pASP);
	pASP->InitNew();
	hr = pAS->AddNamedItem(L"MyCustomObj", SCRIPTITEM_GLOBALMEMBERS | SCRIPTITEM_ISVISIBLE);
	//  スクリプトテキストの解析と実行
	hr = pASP->ParseScriptText(L"puts 'hello'", 0, 0, 0, 0, 0, 0, 0, 0);
	hr = pAS->SetScriptState(SCRIPTSTATE_CONNECTED);  

	pASP->Release();
	hr = pAS->Close();
	pAS->Release();

	CloseHandle(hConsole);
	CoUninitialize();

	return 0;
}

