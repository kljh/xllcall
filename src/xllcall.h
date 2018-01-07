#include <node.h>

void xllcall_debug_v8(const v8::FunctionCallbackInfo<v8::Value>& args);

void xllcall_ffi_v8(const v8::FunctionCallbackInfo<v8::Value>& args);

int xll_load(const char* xll_path);

int xll_call(const char* fct_name, double a);


void xldyncall_ctypes_impl(
	const char* fct_name,
	void* fct_ptr,
	void* free_ptr, 
	const char* fct_abi, 
	const char* fct_type,
	const std::vector<std::string>& arg_types,	// [nbArgs]
	/*const*/ std::vector<v8::Value*>& arg_val,	// [nbArgs]
	v8::ReturnValue<v8::Value>& ret_val
	);