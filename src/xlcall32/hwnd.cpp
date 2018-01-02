#include "hwnd.h"
#include <windows.h>
#include <stdio.h>

HWND get_hwnd() {
	static HWND sHwnd = 0;
	if (sHwnd)
		return sHwnd;
	
	DWORD dwStyle = WS_VISIBLE | WS_CAPTION;
	dwStyle = WS_VISIBLE | WS_POPUP;
	dwStyle = 0;

	sHwnd = ::CreateWindowA(
		"STATIC", // _In_opt_  LPCTSTR lpClassName,
		"ExcelHwd", // _In_opt_  LPCTSTR lpWindowName,
		dwStyle, // _In_      DWORD dwStyle,
		0, //CW_USEDEFAULT, // _In_      int x,
		0, //CW_USEDEFAULT, // _In_      int y,
		8, //CW_USEDEFAULT, // _In_      int nWidth,
		8, //CW_USEDEFAULT, // _In_      int nHeight,
		NULL, // _In_opt_  HWND hWndParent,
		NULL, // _In_opt_  HMENU hMenu,
		NULL, // _In_opt_  HINSTANCE hInstance,
		NULL // _In_opt_  LPVOID lpParam
		);

	if (sHwnd)
		return sHwnd;
	
	fprintf(stderr, "XLCall32/hwnd/CreateWindowA: error");

	return sHwnd;
}