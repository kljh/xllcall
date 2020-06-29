#include "xllcall.h"
#include "xloper.h"
#include "v8conv.h"
#include "xlcall32/xlcall32.h"
#include <dyncall/dyncall.h>
#include <dyncall/dyncall_signature.h>

#include <excpt.h>
#include <stdio.h>
#include <map>

typedef void (*xlAutoFree_t)(xloper*);
typedef void (*xlAutoFree12_t)(xloper12*);

bool& xllcall_debug() {
    static bool s_b = false;
    return s_b;
}

namespace {
void throw_v8_exception(v8::Isolate* isolate, const char * msg) {
    isolate->ThrowException(v8::Exception::TypeError(
        v8::String::NewFromUtf8(isolate, msg, v8::NewStringType::kInternalized).ToLocalChecked() ));    // kNormal or kInternalized ?
}}

void xllcall_debug_v8(const v8::FunctionCallbackInfo<v8::Value>& args) {
    v8::Isolate* isolate = args.GetIsolate();
    
    if (args.Length() > 0) {
        if (!args[0]->IsBoolean()) {
            throw_v8_exception(isolate, "xllcall_debug: expects one optional boolean argument"); 
            return;
        } else {
            bool new_value = args[0]->BooleanValue(isolate);
            xllcall_debug() = new_value; 
        }
    }

    bool curr_value = xllcall_debug();
    v8::Local<v8::Value> res = v8::Boolean::New(isolate, curr_value);
    args.GetReturnValue().Set(res);
}


// http://blog.aaronballman.com/2012/02/describing-the-msvc-abi-for-structure-return-types/

std::map<std::string, HMODULE> module_handler_cache;
std::map<std::string, void*> function_pointer_cache;

void* get_fct_ptr(const std::string& dll_path, const std::string& fct_name) {
    std::string dll_name, dll_folder;
    std::string::size_type sep = dll_path.find_last_of('\\');
    if (sep!=std::string::npos) {
        dll_name = std::string(dll_path.begin()+sep+1, dll_path.end());
        dll_folder = std::string(dll_path.begin(), dll_path.begin()+sep);
        BOOL bOk = SetCurrentDirectory(dll_folder.c_str());
        if (!bOk) fprintf(stderr, "%s: SetCurrentDirectory(%s) failed.\n",  __FUNCTION__, dll_folder.c_str());
    } else {
        dll_name = dll_path;
    }

    HMODULE dll_handle = NULL;
    if (!dll_name.empty()) {
        if (module_handler_cache.find(dll_name)!=module_handler_cache.end()) {
            dll_handle = module_handler_cache[dll_name];
        } else {
            dll_handle = LoadLibraryA(dll_name.c_str());
        }
        if (!dll_handle) {
            DWORD err = GetLastError();
            fprintf(stderr, "%s: LoadLibrary(%s) failed. Error=0x%x.\n", __FUNCTION__, dll_name.c_str(), err);
            return 0;
        } else {
            module_handler_cache[dll_name] = dll_handle;
        }
    }

    void* fct_ptr = NULL;
    if (dll_handle && !fct_name.empty()) {
        if (function_pointer_cache.find(fct_name)!=function_pointer_cache.end()) {
            fct_ptr = function_pointer_cache[fct_name];
        } else {
            fct_ptr = GetProcAddress(
                dll_handle, fct_name.c_str()); 
        }
        if (!fct_ptr) {
            fprintf(stderr, "%s: GetProcAddress(%s) failed.\n", __FUNCTION__, fct_name.c_str());
            return 0;
        } else {
            function_pointer_cache[fct_name] = fct_ptr;
        }
    }

    return fct_ptr;
}

int seh_filter(unsigned int code, struct _EXCEPTION_POINTERS *ep) {
    fprintf(stderr, "seh_filter (%i).\n", code);
    return EXCEPTION_EXECUTE_HANDLER;  // 1
    //return EXCEPTION_CONTINUE_SEARCH;  // 0 
    //return EXCEPTION_CONTINUE_EXECUTION;  // -1
};

void xllload_xlAutoOpen_v8(const v8::FunctionCallbackInfo<v8::Value>& args);
void xllload_xlAutoOpen_v8_seh(const v8::FunctionCallbackInfo<v8::Value>& args) {
    __try {
        xllload_xlAutoOpen_v8(args);
    } __except(seh_filter(GetExceptionCode(), GetExceptionInformation())) {
        fprintf(stderr, " --- XLLCALL32 STRUCTURED EXCEPTION HANDLER --- \n");
    }
}
void xllload_xlAutoOpen_v8(const v8::FunctionCallbackInfo<v8::Value>& args) {
    printf("%s: sizeof int %i.\n", __FUNCTION__, (int)sizeof(int));
	printf("%s: sizeof XLOPER %i.\n", __FUNCTION__, (int)sizeof(XLOPER));
    set_excel4v_t sxl4v = (set_excel4v_t)get_fct_ptr("xlcall32.dll", "set_excel4v");
	sxl4v((excel4v_t)nullptr);

    v8::Isolate* isolate = args.GetIsolate();
    
    if (args.Length()==0 || !args[0]->IsString()) {
        return throw_v8_exception(isolate, "xllcall_debug: expects one string argument");
    }

    std::string xll_path;
    if (!v8value_2_native(isolate, args[0], xll_path)) return;

    typedef short   (__stdcall *PFN_SHORT_VOID)();
    printf("%s: looking for xlAutoOpen in %s.\n", __FUNCTION__, xll_path.c_str());
    PFN_SHORT_VOID xlAutoOpen_ptr = (PFN_SHORT_VOID)get_fct_ptr(xll_path, "xlAutoOpen");
    printf("%s: calling for xlAutoOpen in %s.\n", __FUNCTION__, xll_path.c_str());
    xlAutoOpen_ptr();

    printf("%s: done.\n", __FUNCTION__);
    args.GetReturnValue().Set(args[0]);
}

struct ctypes_memory_holder_item_t {
    std::string m_string;
    std::wstring m_wstring;
    XLOper m_xop;
    XLOper12 m_xop12;
};
typedef std::vector<ctypes_memory_holder_item_t> ctypes_memory_holder_t;

void xllcall_ffi_v8(const v8::FunctionCallbackInfo<v8::Value>& args);
void xllcall_ffi_v8_seh(const v8::FunctionCallbackInfo<v8::Value>& args) {
    __try {
        xllcall_ffi_v8(args);
    } __except(seh_filter(GetExceptionCode(), GetExceptionInformation())) {
        fprintf(stderr, " --- XLLCALL32 STRUCTURED EXCEPTION HANDLER --- \n");
    }
}
void xllcall_ffi_v8(const v8::FunctionCallbackInfo<v8::Value>& args) {
    bool bVerbose = xllcall_debug();

    if (bVerbose) printf("%s BEGIN.\n", __FUNCTION__);
    v8::Isolate* isolate = args.GetIsolate();
    v8::Local<v8::Context> context = isolate->GetCurrentContext();

    // Check the number of arguments passed.
    if (args.Length() != 6) {
        // Throw an Error that is passed back to JavaScript
        fprintf(stderr, "%s: #args %i", __FUNCTION__, (int)args.Length());
        throw_v8_exception(isolate, "xllcall_ffi: wrong number of arguments, 5 expected: (xll_path, fct_name, return_type, arg_types, arg_names, arg_vals)");
        return;
    }

    // Check the argument types
    if (!args[0]->IsString() || !args[1]->IsString() || !args[2]->IsString() || !args[3]->IsArray() || (!args[4]->IsUndefined()&&!args[4]->IsArray()) || !args[5]->IsArray()) {
        throw_v8_exception(isolate, "xllcall_ffi: wrong arguments types, expected (string, string, string, array, ?array, array)");
        return;
    }

    v8::String::Utf8Value s(isolate, args[0]);
    std::string xll_path(*s);
    //std::string xll_path = v8::String::Utf8Value(isolate, args[0]->ToString());
    std::string fct_name(*v8::String::Utf8Value(isolate, args[1]));
    std::string return_type(*v8::String::Utf8Value(isolate, args[2]));
    
    void* fct_ptr = get_fct_ptr(xll_path, fct_name);
    void* xlAutoFree_ptr = get_fct_ptr(xll_path, "xlAutoFree");
    void* xlAutoFree12_ptr = get_fct_ptr(xll_path, "xlAutoFree12");
    if (!fct_ptr && !module_handler_cache[xll_path]) {
        throw_v8_exception(isolate, "could not load XLL (check x86 vs x64 arch)");
        return;
    }
    if (!fct_ptr) {
        throw_v8_exception(isolate, "could not load XLL function");
        return;
    }
    
    const char* fct_abi = 0; 
    const char* fct_type = return_type.c_str();
    std::vector<std::string> arg_types;	// [nbArgs]
    std::vector<std::string> arg_names;	// [nbArgs]
    if (!v8value_2_native(isolate, args[3], arg_types)) return;
    if (!v8value_2_native(isolate, args[4], arg_types)) return;
    size_t nb_arg_types = arg_types.size();
    size_t nb_arg_names = arg_names.size();
    if (nb_arg_names!=0 && nb_arg_names!=nb_arg_types) {
        fprintf(stderr, "nb_arg_types: %i", (int)nb_arg_types);
        fprintf(stderr, "nb_arg_names: %i", (int)nb_arg_names);
        throw_v8_exception(isolate, "arg_types and arg_names are different size.");
        return;
    }

    v8::Local<v8::Array> arg_vals = v8::Local<v8::Array>::Cast(args[5]);
    size_t nb_args_vals = arg_vals->Length();

    if (bVerbose) {
        printf("%s: fct_name %s\n", __FUNCTION__, fct_name.c_str());
        printf("%s: #arg_types %i\n", __FUNCTION__, (int)nb_arg_types);
        printf("%s: #args_vals %i\n", __FUNCTION__, (int)nb_args_vals);
    }

    DCCallVM* vm = 0;
    //try 
    {
        vm = dcNewCallVM(4096);
        dcReset(vm);

#ifdef _WIN64
        // only one callling convention in 64 bit
        dcMode(vm, DC_CALL_C_X64_WIN64);
#else
        if (fct_abi==0 || strcmp(fct_abi, "") == 0 || strcmp(fct_abi, "stdcall") == 0) 
        {
            // stdcall is default convention
            // stdcall also known as WINAPI
            dcMode(vm, DC_CALL_C_X86_WIN32_STD); 
        } 
        else if (strcmp(fct_abi,"cdecl")==0) 
        {
            // cdecl
            dcMode(vm, DC_CALL_C_X86_CDECL);
        } 
        else 
        {
            dcFree(vm); vm = 0;
            fprintf(stderr, "%s: unknown calling convention %s.\n", __FUNCTION__, fct_abi);
            throw_v8_exception(isolate, "unknown calling convention");
            return;
        }
#endif

        /* 
        Excel XLL arguments :

        Data type			ByValue	ByRef 	Comments
        Boolean					A	L	short int
        double					B	E	
        char *						C,F	Null-terminated ASCII byte string
        unsigned char *				D,G	Counted ASCII byte string
        unsigned short [int]	H		16-bit WORD
        [signed] short [int]	I	M	16-bit signed integer
        [signed long] int		J	N	32-bit signed integer
        FP							K	Floating-point array structure
        Array						O	unsigned short int* / unsigned short int* / double[]
        XLOPER						P	Variable-type worksheet values and arrays
        XLOPER Ref					R	Values, arrays, and range references

        Introduced  in Excel 2007 to support large grids and long Unicode strings:

        unsigned short *		C%, F%	Null-terminated Unicode wide-character string
        unsigned short *		D%, G%	Counted Unicode wide-character string
        FP12						K%	Larger grid floating-point array structure
        Array						O%	signed int * / signed int * / double[]
        XLOPER12					Q	Variable-type worksheet values and arrays
        XLOPER12 Ref				U	Values, arrays, and range references

        Notes:
        - The string types F, F%, G, and G% are used for arguments that are modified-in-place.
        - The C-language declarations assume that your compiler uses 8-byte doubles, 2-byte short integers, and 4-byte long integers by default.
        - All functions in DLLs and code resources are called using the __stdcall calling convention.
        - Any function that returns a data type by reference, that is, that returns a pointer to something, can safely return a null pointer. Excel interprets a null pointer as a #NUM! error.

        */
        
        // push the arguments
        dcReset(vm);
        {
        ctypes_memory_holder_t mem_hold(nb_arg_types); // we must reserve full size otherwise items get destructed on push_back

        static XLOPER missing;
        static XLOPER12 missing12;
        missing.xltype = xltypeMissing;
        missing12.xltype = xltypeMissing;

        for (size_t a=0; a<nb_arg_types; a++) {
            bool missing_arg = !(a<nb_args_vals);
            
            const char* arg_type = arg_types[a].c_str();
            v8::MaybeLocal<v8::Value> maybe_arg_val;
            v8::Local<v8::Value> arg_val;
            if (!missing_arg) maybe_arg_val = arg_vals->Get(context, a);
            if (maybe_arg_val.IsEmpty()) 
                missing_arg = true;
            else
                arg_val = maybe_arg_val.ToLocalChecked();
            
            if (bVerbose) printf("%s: arg %i of type %s.\n", __FUNCTION__, (int)a, arg_type);
            if (_stricmp(arg_type, "XLOPER12*") == 0)
            {
                if (missing_arg) {
                    dcArgPointer(vm, (void*)&missing12);
                } else {
                    bool bOk = v8value_2_xloper(isolate, a, arg_val, mem_hold[a].m_xop12);
                    if (!bOk) fprintf(stderr, "%s: problem converting argument %i/%i into xloper.\n", __FUNCTION__, (int)(a+1), (int)nb_arg_types);
                    dcArgPointer(vm, (void*)&mem_hold[a].m_xop12); // leaking memory ? !!
                }
            }
            else if (_stricmp(arg_type, "XLOPER*") == 0)
            {
                if (missing_arg) {
                    dcArgPointer(vm, (void*)&missing);
                } else {
                    bool bOk = v8value_2_xloper(isolate, a, arg_val, mem_hold[a].m_xop);
                    if (!bOk) fprintf(stderr, "%s: problem converting argument %i/%i into xloper.\n", __FUNCTION__, (int)(a+1), (int)nb_arg_types);
                    dcArgPointer(vm, (void*)&mem_hold[a].m_xop); // leaking memory ? !!
                }
            }
            else if (_stricmp(arg_type, "wchar_t*") == 0)
            {
                if (missing_arg) {
                    dcArgPointer(vm, (void*)L"");
                } else {
                    if (!v8value_2_native(isolate, arg_val, mem_hold[a].m_wstring)) return;
                    dcArgPointer(vm, (void*)(mem_hold[a].m_wstring.c_str()));
                }
            }
            else if (_stricmp(arg_type, "char*") == 0)
            {
                if (missing_arg) {
                    dcArgPointer(vm, "");
                } else {
                    if (!v8value_2_native(isolate, arg_val, mem_hold[a].m_string)) return;
                    dcArgPointer(vm, (void*)(mem_hold[a].m_string.c_str()));
                }
            }
            else if (_stricmp(arg_type, "double") == 0)
            {
                if (missing_arg) {
                    dcArgDouble(vm, 0.0);
                } else {
                    double d;
                    if (!v8value_2_native(isolate, arg_val, d)) return;
                    dcArgDouble(vm, d);
                }
            } 
            //else if (_stricmp(arg_type, "double*") == 0)
            //{
            //	experimental support of type 'E' but input is not a floating number
            //	const double* p = ;
            //	dcArgPointer(vm, (void*)p);
            //}
            else if (_stricmp(arg_type, "bool") == 0)
            {
                if (missing_arg) {
                    dcArgBool(vm, false);
                } else {
                    bool b;
                    if (!v8value_2_native(isolate, arg_val, b)) return;
                    short int sib = (short int)b;
                    dcArgBool(vm, sib);
                }
            } 
            else if (_stricmp(arg_type, "int32") == 0)
            {
                if (missing_arg) {
                    dcArgLong(vm, 0);
                } else {
                    long l;
                    if (!v8value_2_native(isolate, arg_val, l)) return;
                    dcArgLong(vm, l);
                }
            } 
            else 
            {
                dcFree(vm); vm = 0;
                fprintf(stderr, "%s: unknown argument %i type %s.\n", __FUNCTION__, (int)a, arg_type);
                throw_v8_exception(isolate, "unknown argument type (see stderr for details)");
                return;
            }
        }
        
        v8::Local<v8::Value> ret_val; // = v8::Number::New(isolate, 7865);
        
        if (bVerbose) printf("%s: res type %s.\n", __FUNCTION__, fct_type);
        if (_stricmp(fct_type, "XLOPER12*") == 0)
        {
            // call the function 
            if (bVerbose) printf("%s: dcCallPointer...\n", __FUNCTION__);
            void *res = dcCallPointer(vm, fct_ptr);
            
            // convert the result
            if (bVerbose) printf("%s: xloper_2_v8value...\n", __FUNCTION__);
            XLOPER12* xop = (XLOPER12*)res;
            ret_val = xloper_2_v8value(isolate, *xop);

            // free 
            if (bVerbose) printf("%s: dcFree...\n", __FUNCTION__);
            dcFree(vm); vm = 0;
            
            if (xop->xltype & xlbitDLLFree) {
                if (xlAutoFree12_ptr) {
                    ((xlAutoFree12_t)xlAutoFree12_ptr)(xop);
                } else {
                    fprintf(stderr, "%s: no xlAutoFree12 function.\n", __FUNCTION__);
                    throw_v8_exception(isolate, "no xlAutoFree12 function");
                    return;
                }
            }
        }
        else if (_stricmp(fct_type, "XLOPER*") == 0)
        {
            // call the function 
            void *res  = dcCallPointer(vm, fct_ptr);

            // convert the result
            XLOPER* xop = (XLOPER*)res;
            ret_val = xloper_2_v8value(isolate, *xop);

            // free 
            dcFree(vm); vm = 0;

            if (xop->xltype & xlbitDLLFree) {
                if (xlAutoFree_ptr) {
                    ((xlAutoFree_t)xlAutoFree_ptr)(xop);
                } else {
                    fprintf(stderr, "%s: no xlAutoFree function.\n", __FUNCTION__);
                    throw_v8_exception(isolate, "no xlAutoFree function");
                    return;
                }
            }
        }
        else if (_stricmp(fct_type, "double") == 0) 
        {
            double d = dcCallDouble(vm, fct_ptr);
            dcFree(vm); vm = 0;

            ret_val = v8::Number::New(isolate, d);
        }
        else if (_stricmp(fct_type, "void") == 0)
        {
            dcCallVoid(vm, fct_ptr);
            dcFree(vm); vm = 0;

            // ret_val = undefined;
        }
        else 
        {
            dcFree(vm); vm = 0;
            fprintf(stderr, "%s: unknown return type %s.\n", __FUNCTION__, fct_type);
            throw_v8_exception(isolate, "unknown return type (see stderr)");
            return;
        }
        
        if (bVerbose) printf("%s: returning result (of type %s).\n", __FUNCTION__, fct_type);
        args.GetReturnValue().Set(ret_val);
    
        //if (bVerbose) printf("%s: memory holder for arguments - before destruction.\n", __FUNCTION__);
        }
        //if (bVerbose) printf("%s: memory holder for arguments - after destruction.\n", __FUNCTION__);

    /* } catch (std::exception& e) {
        if (vm) {
            dcFree(vm); vm = 0;
        }
        fprintf(stderr, "%s: C++ exception caught: %s.\n", __FUNCTION__, e.what());
        isolate->ThrowException(v8::Exception::TypeError(
            v8::String::NewFromUtf8(isolate, "unknown return type (see stderr)")));
        return;
    */
    }
    if (bVerbose) printf("%s END.\n", __FUNCTION__);
}
