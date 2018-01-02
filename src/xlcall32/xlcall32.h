#pragma once

#ifdef __cplusplus
#include <vector>
#include <string>
#endif

#define WIN32_LEAN_AND_MEAN
#include <windows.h>

#ifdef XLCALL32_EXPORTS
#define XLCALL32_CAPI __declspec(dllexport)
#else
#define XLCALL32_CAPI __declspec(dllimport)
#endif

#ifdef __cplusplus
extern "C" {
#endif

	/*
	This API is not reentrant (returned strings are kept in internal static variables).
	It is not intended to be multi-threaded which is consistent with the way most XLL work.
	The intent is however to distribute the XLL calculation on a grid (e.g. multi-procesing rather than multi-threading).
	*/

	// Excel 97+ exported function (exported by xlcall32.dll)
	struct xloper;
	typedef struct xloper* LPXLOPER;
	int pascal XLCallVer(void);
	long pascal LPenHelper(int wCode, VOID *lpv);
	int pascal Excel4v(int xlfn, LPXLOPER operRes, int count, LPXLOPER opers[]);
	int _cdecl Excel4(int xlfn, LPXLOPER operRes, int count, ...);
	
	// Excel 2007+ exported functions (exported by excel.exe)
	// struct xloper12;
	// typedef struct xloper12* LPXLOPER12;
	// int pascal MdCallBack_impl(int xlfn, int n, LPXLOPER args[], LPXLOPER ret);
	// int pascal MdCallBack12_impl(int xlfn, int n, LPXLOPER12 args[], LPXLOPER12 ret);

	// Excel and XLL SDK version reported to XLLs (used for XLCallVer)
	XLCALL32_CAPI int  xlcall32_excel_version_get();
	XLCALL32_CAPI void xlcall32_excel_version_set(int);
	

#ifdef __cplusplus
}
#endif
