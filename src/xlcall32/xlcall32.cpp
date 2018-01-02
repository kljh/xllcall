/*

Excel 97 (v8.0) introduced XLL addins. Callbacks from XLL addin to Excel are exported in XLCALL32.LIB/DLL.
Exported functions are: XLCallVer, Excel4, Excel4v, LPenHelper
This library is a replacement for Microsoft's one to allow when calling XLL from C programs.

Excel 2007 (v12) introduced new XLOPER to handle bigger array and bigger strings.
Callback from XLL addins to Excel have also been implemented in a different way than for previous versions.
XLCall32.lib is no longer used, instead EXCEL.EXE (and XLLSERVE.EXE) exports symbols (the same way a DLL would).

Those symbols are exported using the /EXPORT:<exportname>=<fctname> option when calling LINK:EXE.
  /EXPORT:LPenHelper=LPenHelper_impl
  /EXPORT:MdCallBack=MdCallBack_impl 
  /EXPORT:MdCallBack12=MdCallBack12_impl 

Sometimes we do not have the control on the exe (Python.exe, node.exe, etc) so we may hook the GetProcAddress and other Win32 API calls. 
That's can be done with CAPIHook.

*/

#define XLCALL32_EXPORTS
#include "xlcall32.h"
#include "../xlsdk97/xlcall.h"

// functions exposed by the XLL to XL
typedef int (*xlAutoOpen_t)(void);
typedef int (*xlAutoClose_t)(void);
typedef void (*xlAutoFree_t)(xloper*);
typedef void (*xlAutoFree12_t)(xloper12*);

#include "hwnd.h"
#include <sstream>

BOOL APIENTRY DllMain( 
	HMODULE hModule,
	DWORD  ul_reason_for_call,
	LPVOID lpReserved
	)
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
		break;
	case DLL_PROCESS_DETACH:
		break;
	case DLL_THREAD_ATTACH:
		break;
	case DLL_THREAD_DETACH:
		break;
	}
	return TRUE;
}

// ----------------------------------------------------------------------------
// Excel 97+ exports function (exported by xlcall32.dll)

// pragma below is equivalent to the information from xlcall32.def 
// #pragma comment(linker, "/EXPORT:XLCallVer,@1")
int pascal XLCallVer(void)
{
	return xlcall32_excel_version_get() * 256;
}

// pragma below is equivalent to the information from xlcall32.def 
// #pragma comment(linker, "/EXPORT:LPenHelper,@4")
long pascal LPenHelper(int wCode, VOID *lpv)
{
	return 0;
}

// pragma below is equivalent to the information from xlcall32.def 
// #pragma comment(linker, "/EXPORT:Excel4,@2")
int _cdecl Excel4(int xlfn, LPXLOPER operRes, int count, ... )
{
	va_list va;
	va_start(va, count);

	std::vector<LPXLOPER> opers;
	for (int i=0; i<count; i++)
		opers.push_back(va_arg(va, LPXLOPER));

	va_end(va);

	return Excel4v(xlfn, operRes, count, opers.empty()?0:&opers[0]);
}

// pragma below is equivalent to the information from xlcall32.def 
// #pragma comment(linker, "/EXPORT:Excel4v,@3")
int pascal Excel4v(int xlfn, LPXLOPER ret, int n, LPXLOPER args[])
{
	// Thread Local Storage (TLS) to handle multi-threading

	if (!ret) {
		// no return expected;
		return xlretSuccess;
	}

	const char* xlfn_name = 0;
	XLOPER& res = *ret;
	switch (xlfn) {
		case xlFree:
			xlfn_name = "xlFree MISSING => LEAKING !!";
			break;
		case xlSheetId:
			res.val.mref.idSheet = 0;
			return xlretSuccess;
		case xlfGetWorkspace:
			// xlfGetWorkspace with first argument set to 2 queries for the version of Excel formated as string
			xlfn_name = "xlfGetWorkspace";
			break;
		case xlSheetNm:
			xlfn_name = "xlSheetNm";
			break;
		case xlAbort:
			xlfn_name = "xlAbort";
			break;
		case xlGetHwnd:
		{
			HWND hwnd = get_hwnd();
			res.val.w = (short)hwnd;
			return xlretSuccess;
		}
		case xlGetName:
			// 	xlGetName: not defined outside xlAutoCall
			//  returns currently loaded XLL module path
			xlfn_name = "xlGetName";
			break;
		case xlfRegister:
			xlfn_name = "xlfRegister MISSING";
			break;
		
		case xlfCaller:
			xlfn_name = "xlfCaller";
			break;
		case xlcAlert:
			xlfn_name = "xlcAlert";
			break;
		case xlcMessage:
			xlfn_name = "xlcMessage";
			break;

	}

	if (xlfn_name)
		fprintf(stderr, "Excel4v xlfn_name=%s not implemented.\n", xlfn_name);
	else
		fprintf(stderr, "Excel4v xlfn=%i not implemented.\n", xlfn);
	
	return xlretSuccess;
}

// ----------------------------------------------------------------------------

int* xlcall32_excel_version() {
	static int v = 8;
	return &v;
}
int xlcall32_excel_version_get() {
	return *xlcall32_excel_version();
}
void xlcall32_excel_version_set(int v) {
	*xlcall32_excel_version() = v;
}

// ----------------------------------------------------------------------------
/*

http://msdn.microsoft.com/en-us/library/office/bb687923%28v=office.15%29.aspx

Auxiliary function numbers (only from the C API) : 
xlFree 0 : No return value. Arguments : one or more XLOPER/XLOPER12s to be freed. 
xlStack : Returns the number of bytes (xltypeInt) remaining on the stack.
xlCoerce : Converts one type of XLOPER/XLOPER12 to another, or looks up cell values on a sheet.
xlSheetId : 
xlSheetNm : Returns the name of the sheet (xltypeStr) in the form [Book1]Sheet1.
xlAbort : Returns TRUE (xltypeBool) if the user has pressed ESC.
xlGetHwnd : Contains the window handle (xltypeInt) in the val.w field. 
xlGetName : Returns the path and file name (xltypeStr). No arguments.

Excel command numbers 0x8000 :
xlcAlert 118 :
xlcMessage 122 

Excel function numbers :
xlfRegister 149 : (Form 1) On success, returns the register ID of the function (xltypeNum). Otherwise, returns a #VALUE! error. Arguments : pxModuleText, pxProcedure, ...
xlfUnregister 201 : (Form 1) Unregisters an individual command or function. (Form 2) Unloads and deactivates an XLL.
xlfCaller 89 : 

in ExcelDNA

xlfGetName 107 :
xlfSetName 88 : 
	
xlfGetWorkspace 186 :
xlfEvaluate
xlfCall
xlfRtd

xlfGetBar 182 :
xlfGetToolbar 258
xlfAddMenu
xlfDeleteMenu
xlfAddCommand
xlfDeleteCommand

*/

