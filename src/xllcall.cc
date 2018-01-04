#include "xllcall.h"
#include "xloper.h"
#include "v8conv.h"
#include <dyncall/dyncall.h>
#include <dyncall/dyncall_signature.h>

#include <stdio.h>
#include <map>

typedef void (*xlAutoFree_t)(xloper*);
typedef void (*xlAutoFree12_t)(xloper12*);

// http://blog.aaronballman.com/2012/02/describing-the-msvc-abi-for-structure-return-types/

std::map<std::string, HMODULE> module_handler_cache;
std::map<std::string, void*> function_pointer_cache;

void* get_fct_ptr(const std::string& dll_path, const std::string& fct_name) {
	std::string dll_name, dll_folder;
	std::string::size_type sep = dll_path.find_last_of("\\");
	if (sep!=std::string::npos) {
		dll_name = std::string(dll_path.begin()+sep+1, dll_path.end());
		dll_folder = std::string(dll_path.begin(), dll_path.end()+
		sep);
		BOOL b = SetCurrentDirectory(dll_folder.c_str());
		printf("SetCurrentDirectory(%s) returned %i\n", dll_folder.c_str(), (int)b);
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
      fprintf(stderr, "%s: LoadLibrary(%s) failed. Error=0x%x\n", __FUNCTION__, dll_name.c_str(), err);
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

struct ctypes_memory_holder_item_t {
	std::string m_string;
	std::wstring m_wstring;
	XLOper m_xop;
	XLOper12 m_xop12;
};
typedef std::vector<ctypes_memory_holder_item_t> ctypes_memory_holder_t;

void xllcall_ffi_v8(const v8::FunctionCallbackInfo<v8::Value>& args) {
	printf("%s BEGIN.\n", __FUNCTION__);
  v8::Isolate* isolate = args.GetIsolate();

  // Check the number of arguments passed.
  if (args.Length() != 5) {
    // Throw an Error that is passed back to JavaScript
    isolate->ThrowException(v8::Exception::TypeError(
      v8::String::NewFromUtf8(isolate, "xllcall_ffi: wrong number of arguments, 5 expected: (xll_path, fct_name, return_type, arg_types, arg_vals)")));
    return;
  }

  // Check the argument types
  if (!args[0]->IsString() || !args[1]->IsString() || !args[2]->IsString() || !args[3]->IsArray() || !args[4]->IsArray()) {
    isolate->ThrowException(v8::Exception::TypeError(
      v8::String::NewFromUtf8(isolate, "xllcall_ffi: wrong arguments types, expected (string, string, string, array, array)")));
    return;
  }

  v8::String::Utf8Value s(args[0]);
  std::string xll_path(*s);
  //std::string xll_path = v8::String::Utf8Value(args[0]->ToString());
  std::string fct_name(*v8::String::Utf8Value(args[1]));
  std::string return_type(*v8::String::Utf8Value(args[2]));
  
	void* fct_ptr = get_fct_ptr(xll_path, fct_name);
	void* free_ptr = get_fct_ptr(xll_path, "xlAutoFree");
	void* free12_ptr = get_fct_ptr(xll_path, "xlAutoFree12");
	if (!fct_ptr && !module_handler_cache[xll_path]) {
		isolate->ThrowException(v8::Exception::TypeError(
			v8::String::NewFromUtf8(isolate, "could not load XLL (check x86 vs x64 arch)")));
		return;
	}
	if (!fct_ptr) {
		isolate->ThrowException(v8::Exception::TypeError(
			v8::String::NewFromUtf8(isolate, "could not load XLL function")));
		return;
  }
  
  const char* fct_abi = 0; 
	const char* fct_type = return_type.c_str();
  std::vector<std::string> arg_types;	// [nbArgs]
  if (!v8value_2_native(isolate, args[3], arg_types)) return;
  size_t nb_arg_types = arg_types.size();

  if (!args[4]->IsArray()) {
    isolate->ThrowException(v8::Exception::TypeError(
      v8::String::NewFromUtf8(isolate, "arg_vals not an array")));
    return;
  }
  v8::Local<v8::Array> arg_vals = v8::Handle<v8::Array>::Cast(args[4]);
  size_t nb_args_vals = arg_vals->Length();

  //printf("#arg_types %i\n", (int)nb_arg_types);
  //printf("#args_vals %i\n", (int)nb_args_vals);
  

	DCCallVM* vm = 0;
	try {
		vm = dcNewCallVM(4096);
		dcReset(vm);

#ifdef WIN64
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
      isolate->ThrowException(v8::Exception::TypeError(
        v8::String::NewFromUtf8(isolate, "unknown calling convention")));
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
		ctypes_memory_holder_t mem_hold(nb_arg_types); // we must reserve full size otherwise items get destructed on push_back

		static XLOPER missing;
		static XLOPER12 missing12;
		missing.xltype = xltypeMissing;
		missing12.xltype = xltypeMissing;

		for (size_t a=0; a<nb_arg_types; a++) {
      bool missing_arg = !(a<nb_args_vals);
      
			const char* arg_type = arg_types[a].c_str();
      v8::Local<v8::Value> arg_val;
      if (!missing_arg) arg_val = arg_vals->Get(a);
			
			printf("arg %i of type %s.\n", (int)a, arg_type);
			if (_stricmp(arg_type, "XLOPER12*") == 0)
			{
				if (missing_arg) {
          dcArgPointer(vm, (void*)&missing12);
        } else {
					bool bOk = v8value_2_xloper(isolate, a, arg_val, mem_hold[a].m_xop12);
					if (!bOk) fprintf(stderr, "problem converting argument %i/%i into xloper.\n", (int)(a+1), (int)nb_arg_types);
          dcArgPointer(vm, (void*)&mem_hold[a].m_xop12); // leaking memory ? !!
        }
			}
			else if (_stricmp(arg_type, "XLOPER*") == 0)
			{
        if (missing_arg) {
          dcArgPointer(vm, (void*)&missing);
        } else {
          bool bOk = v8value_2_xloper(isolate, a, arg_val, mem_hold[a].m_xop);
					if (!bOk) fprintf(stderr, "problem converting argument %i/%i into xloper.\n", (int)(a+1), (int)nb_arg_types);
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
			//else if (_stricmp(arg_type, "int32") == 0)
			//{
			//	if (missing_arg) {
      //    dcArgLong(vm, 0);
      //  } else {
			//		long l;
			//		if (!v8value_2_native(isolate, arg_val, l)) return;
      //    dcArgLong(vm, l);
			//	}
			//} 
			else 
			{
        dcFree(vm); vm = 0;
				fprintf(stderr, "%s: unknown argument %i type %s.\n", __FUNCTION__, (int)a, arg_type);
        isolate->ThrowException(v8::Exception::TypeError(
          v8::String::NewFromUtf8(isolate, "unknown argument type (see stderr for details)")));
        return;
			}
		}
    
    v8::Local<v8::Value> ret_val; // = v8::Number::New(isolate, 7865);
    
		printf("res type %s.\n", fct_type);
		if (_stricmp(fct_type, "XLOPER12*") == 0)
		{
			// call the function 
			printf("dcCallPointer...\n");
			void *res = dcCallPointer(vm, fct_ptr);
			
			// convert the result
			printf("xloper_2_v8value...\n");
			XLOPER12* xop = (XLOPER12*)res;
			ret_val = xloper_2_v8value(isolate, *xop);

			// free 
			printf("dcFree...\n");
			dcFree(vm); vm = 0;
			
			if (xop->xltype & xlbitDLLFree) {
				if (free12_ptr)
					((xlAutoFree12_t)free12_ptr)(xop);
				else {
					fprintf(stderr, "no xlAutoFree12 function");
					isolate->ThrowException(v8::Exception::TypeError(
						v8::String::NewFromUtf8(isolate, "no xlAutoFree12 function")));
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
				if (free_ptr)
					((xlAutoFree_t)free_ptr)(xop);
				else {
					fprintf(stderr, "no xlAutoFree function");
					isolate->ThrowException(v8::Exception::TypeError(
						v8::String::NewFromUtf8(isolate, "no xlAutoFree function")));
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
      printf("%s: unknown return type %s.\n", __FUNCTION__, fct_type);
      isolate->ThrowException(v8::Exception::TypeError(
        v8::String::NewFromUtf8(isolate, "unknown return type (see stderr)")));
      return;
    }
		
		printf("%s: returning result (of type %s).\n", __FUNCTION__, fct_type);
    args.GetReturnValue().Set(ret_val);
  
	} catch (std::exception& e) {
		if (vm) {
			dcFree(vm); vm = 0;
		}
		fprintf(stderr, "%s: C++ exception caught: %s.\n", __FUNCTION__, e.what());
		isolate->ThrowException(v8::Exception::TypeError(
			v8::String::NewFromUtf8(isolate, "unknown return type (see stderr)")));
		return;
	}  
	printf("%s END.\n", __FUNCTION__);
}
