/*
 * Copyright(c) 2015 arton
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
#if !defined(SCRIPT_ERROR_H)
#define SCRIPT_ERROR_H

class CScriptError : public IActiveScriptError
{
public:
    CScriptError(LPCSTR);

    HRESULT STDMETHODCALLTYPE QueryInterface(
	const IID & riid,  
	void **ppvObj)
    {
	if (!ppvObj) return E_POINTER;
	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IActiveScriptError))
	{
            InterlockedIncrement(&m_lCount);
            *ppvObj = this;
            return S_OK;
	}
	return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef()
    {
	return InterlockedIncrement(&m_lCount);
    }

    ULONG STDMETHODCALLTYPE Release()
    {
	if (InterlockedDecrement(&m_lCount) == 0)
	{
	    delete this;
	    return 0;
	}
	return m_lCount;
    }
    HRESULT STDMETHODCALLTYPE GetExceptionInfo(
	    EXCEPINFO *pexcepinfo  // structure for exception information
    );
    HRESULT STDMETHODCALLTYPE GetSourcePosition(
	DWORD *pdwSourceContext,  // context cookie
	    ULONG *pulLineNumber,     // line number of error
	    LONG *pichCharPosition    // character position of error
    );
    HRESULT STDMETHODCALLTYPE GetSourceLineText(
	    BSTR *pbstrSourceLine  // address of buffer for source line
    );

private:
    ULONG m_lCount;
    int m_nLine;
    int GetLineNumber(LPCSTR);
    void SetupSourceLine(LPCSTR, int);
    std::string m_strMessage;
    std::string m_strBacktrace;
    std::string m_strSource;
};

#endif
