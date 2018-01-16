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

//#pragma pack(1)

#include "../xloper.h"
#define XLCALL32_EXPORTS
#include "xlcall32.h"
#include "../xlsdk97/xlcall.h"

#include "hwnd.h"
#include <sstream>
#include <map>

static excel4v_t s_excel4v_ptr = NULL;

extern "C"
__declspec(dllexport)
void set_excel4v(excel4v_t ptr) {
	printf("%s: sizeof int %i.\n", __FUNCTION__, (int)sizeof(int));
	printf("%s: sizeof XLOPER %i.\n", __FUNCTION__, (int)sizeof(XLOPER));
	s_excel4v_ptr = ptr;
}
set_excel4v_t sxl4v = set_excel4v;




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

struct xll_function_t {
	std::string xll_path, proc_name, proto, fct_name, arg_names;
	size_t id;
};
typedef std::map<std::string, xll_function_t> xll_function_registry_t;
static xll_function_registry_t xll_function_registry;

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
			if (s_excel4v_ptr)
				return s_excel4v_ptr(xlfn, ret, n, args);
			else {
				xlfn_name = "xlfGetWorkspace";
				res.val.str = "x4.0";
				res.val.str[0] = 1;
				res.xltype = xltypeStr;
				break;
			}
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
		{
			xlfn_name = "xlfRegister";
			xll_function_t xll_fct;
			xll_fct.xll_path = (XLOper&)*args[0];
			xll_fct.proc_name = (XLOper&)*args[1]; // "\tSyncMacro\x2>B*SyncMacro_eb7302a7120a5b6c8ec6d9fbf379dd53\x5value"
			xll_fct.proto = (XLOper&)*args[2]; //  "\x2>B*SyncMacro_eb7302a7120a5b6c8ec6d9fbf379dd53\x5value"
			xll_fct.fct_name = (XLOper&)*args[3]; //  "*SyncMacro_eb7302a7120a5b6c8ec6d9fbf379dd53\x5value"
			xll_fct.arg_names = (XLOper&)*args[4]; //  "\x5value"
			//double method_kind = (XLOper&)*args[5]; //  num=2.0
			xll_fct.id = xll_function_registry.size();

			xll_function_registry[xll_fct.fct_name] = xll_fct;

			res.xltype = xltypeNum;
			res.val.num = xll_fct.id;
			break;
		}
		case xlfCaller:
			xlfn_name = "xlfCaller";
			break;
		case xlcAlert:
			xlfn_name = "xlcAlert";
			break;
		case xlcMessage:
			xlfn_name = "xlcMessage";
			if (n>1) {
				std::string msg = (XLOper&)*args[1]; // 2nd argument (1st argument is a boolean)
				fprintf(stderr, ">>> %s\n", msg.c_str());
			} else {
				fprintf(stderr, ">>> ---\n");
			}
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

