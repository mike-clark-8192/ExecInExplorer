// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
// Copyright (c) Microsoft Corporation. All rights reserved

#include <atlbase.h>
#include <comdef.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <sstream>
#include <string>
#include <tuple>
#include <vector>
#include <windows.h>

#pragma comment(lib, "shlwapi.lib")

// use the shell view for the desktop using the shell windows automation to find the
// desktop web browser and then grabs its view
//
// returns:
//      IShellView, IFolderView and related interfaces

HRESULT GetShellViewForDesktop(REFIID riid, void** ppv)
{
    *ppv = NULL;

    IShellWindows* psw;
    HRESULT hr = CoCreateInstance(CLSID_ShellWindows, NULL, CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&psw));
    if (SUCCEEDED(hr)) {
        HWND hwnd;
        IDispatch* pdisp;
        VARIANT vEmpty = {}; // VT_EMPTY
        if (S_OK == psw->FindWindowSW(&vEmpty, &vEmpty, SWC_DESKTOP, (long*)&hwnd, SWFO_NEEDDISPATCH, &pdisp)) {
            IShellBrowser* psb;
            hr = IUnknown_QueryService(pdisp, SID_STopLevelBrowser, IID_PPV_ARGS(&psb));
            if (SUCCEEDED(hr)) {
                IShellView* psv;
                hr = psb->QueryActiveShellView(&psv);
                if (SUCCEEDED(hr)) {
                    hr = psv->QueryInterface(riid, ppv);
                    psv->Release();
                }
                psb->Release();
            }
            pdisp->Release();
        } else {
            hr = E_FAIL;
        }
        psw->Release();
    }
    return hr;
}

// From a shell view object gets its automation interface and from that gets the shell
// application object that implements IShellDispatch2 and related interfaces.

HRESULT GetShellDispatchFromView(IShellView* psv, REFIID riid, void** ppv)
{
    *ppv = NULL;

    IDispatch* pdispBackground;
    HRESULT hr = psv->GetItemObject(SVGIO_BACKGROUND, IID_PPV_ARGS(&pdispBackground));
    if (SUCCEEDED(hr)) {
        IShellFolderViewDual* psfvd;
        hr = pdispBackground->QueryInterface(IID_PPV_ARGS(&psfvd));
        if (SUCCEEDED(hr)) {
            IDispatch* pdisp;
            hr = psfvd->get_Application(&pdisp);
            if (SUCCEEDED(hr)) {
                hr = pdisp->QueryInterface(riid, ppv);
                pdisp->Release();
            }
            psfvd->Release();
        }
        pdispBackground->Release();
    }
    return hr;
}

inline CComVariant WStringToVTBSTR(std::wstring wstr)
{
    CComVariant vt = wstr.c_str();
    return vt;
}

HRESULT ShellExecInExplorerProcess(const std::wstring file, const std::wstring args)
{
    IShellView* psv;
    HRESULT hr = GetShellViewForDesktop(IID_PPV_ARGS(&psv));
    if (SUCCEEDED(hr)) {
        IShellDispatch2* psd;
        hr = GetShellDispatchFromView(psv, IID_PPV_ARGS(&psd));
        if (SUCCEEDED(hr)) {
            CComBSTR bstrFile = file.c_str();
            CComVariant vtArgs = args.c_str();
            VARIANT vtEmpty = {};
            hr = psd->ShellExecuteW(bstrFile, vtArgs, vtEmpty, vtEmpty, vtEmpty);
            psd->Release();
        }
        psv->Release();
    }
    return hr;
}

std::vector<wchar_t> StringToVector(const std::wstring& str)
{
    std::vector<wchar_t> vec(str.length() + 1);
    std::copy(str.begin(), str.end(), vec.begin());
    vec.push_back(L'\0');
    return vec;
}

std::wstring VectorToString(const std::vector<wchar_t>& vec)
{
    std::wstring str(vec.begin(), vec.end());
    return str;
}

std::tuple<std::wstring, std::wstring> SplitCmdline(std::wstring cmdline)
{
    auto cmdlineVec = StringToVector(cmdline);
    auto cmdlineBuf = cmdlineVec.data();
    std::wstring args = PathGetArgsW(cmdlineBuf);
    PathRemoveArgsW(cmdlineBuf);
    std::wstring exe = cmdlineBuf;
    return std::make_tuple(exe, args);
}

std::wstring DescribeLastError(DWORD gle)
{
    wchar_t* message = NULL;
    int fmtOk = FormatMessage(
        FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
        NULL, gle, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPWSTR)&message, 0, NULL);
    if (fmtOk != 0 || message == NULL) {
        return (std::wostringstream() << "Could not format message for: " << gle).str();
    }
    std::wstring result(message);
    LocalFree(message);
    return result;
}

std::wstring DescribeHRESULT(HRESULT hr)
{
    if (SUCCEEDED(hr)) {
        return L"No error.";
    }
    _com_error comError(hr);
    _bstr_t bstrDescription = comError.Description();
    if (bstrDescription.length() > 0) {
        // Error description was retrieved successfully
        return static_cast<const wchar_t*>(bstrDescription);
    } else {
        // Fallback if no description is available
        return L"Error description not available.";
    }
}

void ShowErrorMessageBox(std::wstring message, HRESULT code)
{
    std::wostringstream msg;
    msg << message << L"\n\n"
        << DescribeHRESULT(code);
    MessageBox(NULL, msg.str().c_str(), L"Error", MB_OK | MB_ICONERROR);
}

void ShowUsageMessageBox()
{
    std::wstring path(UNICODE_STRING_MAX_CHARS, L'\0');
    std::wstring thisModulePath = L"ExecInExplorer.asdf";
    if (GetModuleFileNameW(NULL, path.data(), UNICODE_STRING_MAX_CHARS)) {
        thisModulePath = path;
    }
    std::wstring thisModuleName = PathFindFileNameW(thisModulePath.c_str());
    const auto msg = std::wstring(L"Usage: ") + thisModuleName + L" <cmd> <args>";
    MessageBoxW(NULL, msg.c_str(), L"Usage", MB_OK | MB_ICONINFORMATION);
}

std::tuple<std::wstring, std::wstring> GetTargetExeAndArgs()
{
    wchar_t* commandLine = GetCommandLine();
    auto thisArgs = PathGetArgs(commandLine);
    auto splitCmdline = SplitCmdline(thisArgs);
    return splitCmdline;
}

int wWinMain(
    _In_ HINSTANCE /*hInstance*/,
    _In_opt_ HINSTANCE /*hPrevInstance*/,
    _In_ LPWSTR /*lpCmdLine*/,
    _In_ int /*nShowCmd*/)
{
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    if (SUCCEEDED(hr)) {
        auto targetExeAndArgs = GetTargetExeAndArgs();
        auto targetExe = std::get<0>(targetExeAndArgs);
        auto targetArgs = std::get<1>(targetExeAndArgs);
        if (targetExe.empty()) {
            ShowUsageMessageBox();
            return 1;
        }
        hr = ShellExecInExplorerProcess(targetExe, targetArgs);
        if (!SUCCEEDED(hr)) {
            std::wostringstream errMsg;
            errMsg << L"Failed to execute cmd[" << targetExe << L"] args[" << targetArgs << L"]";
            ShowErrorMessageBox(errMsg.str(), hr);
        }
        CoUninitialize();
    } else {
        ShowErrorMessageBox(L"Failed to initialize COM.", hr);
    }
    return SUCCEEDED(hr) ? 0 : abs(hr);
}
