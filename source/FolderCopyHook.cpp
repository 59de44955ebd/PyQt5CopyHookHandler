#include "FolderCopyHook.h"
#include <strsafe.h>
#include <assert.h>
#include <Shlwapi.h>
#pragma comment(lib, "shlwapi.lib")

#include "resource.h"

extern HINSTANCE g_hInst;
extern long g_cDllRef;

#define MAX_WIN_CLASS_LEN 256

char *utf16ToUtf8(const wchar_t *utf16)
{
	// Get size of destination UTF-8 buffer, in chars
	int output_len = WideCharToMultiByte(CP_UTF8, 0, utf16, -1, NULL, 0, NULL, NULL);
	if (!output_len) return NULL;

	// Allocate destination buffer to store UTF-16 string
	char * output = (char *)malloc(output_len + 1);
	if (!output) return NULL;

	// Do the conversion from UTF-8 to UTF-16
	int result = WideCharToMultiByte(
		CP_UTF8,	// convert to UTF-8
		0,

		utf16,		// source UTF-16 string
		-1,

		output,		// destination buffer
		output_len,

		NULL,
		NULL
	);
	if (!result)
	{
		free(output);
		return NULL;
	}
	return output;
}

FolderCopyHook::FolderCopyHook(void) : m_cRef(1)
{
    InterlockedIncrement(&g_cDllRef);
}

FolderCopyHook::~FolderCopyHook(void)
{
    InterlockedDecrement(&g_cDllRef);
}


#pragma region IUnknown

// Query to the interface the component supported.
IFACEMETHODIMP FolderCopyHook::QueryInterface(REFIID riid, void **ppv)
{
    static const QITAB qit[] =
    {
        QITABENT(FolderCopyHook, ICopyHookW),
        { 0 },
    };
    return QISearch(this, qit, riid, ppv);
}

// Increase the reference count for an interface on an object.
IFACEMETHODIMP_(ULONG) FolderCopyHook::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

// Decrease the reference count for an interface on an object.
IFACEMETHODIMP_(ULONG) FolderCopyHook::Release()
{
    ULONG cRef = InterlockedDecrement(&m_cRef);
    if (0 == cRef)
    {
        delete this;
    }

    return cRef;
}

#pragma endregion


#pragma region ICopyHook

IFACEMETHODIMP_(UINT) FolderCopyHook::CopyCallback(HWND hwnd, UINT wFunc,
    UINT wFlags, LPCWSTR pszSrcFile, DWORD dwSrcAttribs, LPCWSTR pszDestFile,
    DWORD dwDestAttribs)
{
	wchar_t pszWinClass[MAX_WIN_CLASS_LEN];
	LoadStringW(g_hInst, IDS_WIN_CLASS, pszWinClass, sizeof(wchar_t) * MAX_WIN_CLASS_LEN);
    if (wcsstr(pszSrcFile, L"__qt5drop__") != NULL && wFunc == FO_COPY)
    {
		char * pDestFile = utf16ToUtf8(pszDestFile);
		if (pDestFile)
		{
			COPYDATASTRUCT ds;
			ds.dwData =	NULL;
			ds.cbData =	(DWORD)strlen(pDestFile);
			ds.lpData =	(void *)pDestFile;

			// send to all top-level windows with class "Qt5152QWindowIcon"
			HWND hWnd = (HWND)NULL;
			while (1)
			{
				hWnd = FindWindowEx((HWND)NULL, hWnd, pszWinClass, NULL);
				if (!hWnd)
					break;
				SendMessage(hWnd, WM_COPYDATA, (WPARAM)(HWND)0, (LPARAM)&ds);
			}

			free(pDestFile);
		}

		return IDNO;
    }

    return IDYES;
}

#pragma endregion