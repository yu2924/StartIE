//
//  StartIE.cpp
//  StartIE
//
//  created by yu2924 on 2022-08-07
//  (c) 2022 yu2924
//

#include "targetver.h"
#include <atlbase.h>
#include <atlstr.h>
#include <strsafe.h>
#include "resource.h"

static CString GetErrorText(HRESULT r)
{
    WCHAR s[256] = {};
    if(!FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, NULL, r, 0, s, _countof(s), NULL))
        if(!FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, NULL, HRESULT_CODE(r), 0, s, _countof(s), NULL))
            StringCchPrintfW(s, _countof(s), L"0x%08X", r);
    StrTrimW(s, L" \r\n");
    return s;
}

struct ScopedCoinitialize
{
    ScopedCoinitialize(DWORD coinit) { CoInitializeEx(NULL, coinit); }
    ~ScopedCoinitialize() { CoUninitialize(); }
};

int APIENTRY wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine, _In_ int nCmdShow)
{
    ScopedCoinitialize coinit(COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    try
    {
        HRESULT r = S_OK;
        CComPtr<IDispatch> disp;
        if(FAILED(r = disp.CoCreateInstance(L"InternetExplorer.Application", NULL, CLSCTX_LOCAL_SERVER))) throw r;
        CComVariant visible(true);
        if(FAILED(r = disp.PutPropertyByName(L"Visible", &visible))) throw r;
        r = S_OK;
    }
    catch(HRESULT r)
    {
        MessageBoxW(NULL, CString(L"Error\n") + GetErrorText(r), CString(MAKEINTRESOURCE(IDR_STARTIE)), MB_ICONERROR | MB_OK);
    }
    return 0;
}
