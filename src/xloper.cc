#include "xloper.h"
#include "stdio.h"

XLOper::XLOper() {
    xltype = 0;
}

XLOper::~XLOper() {
    int is_xl_free = (this->xltype & xlbitXLFree) ? 1 : 0;
    int is_dll_free = (this->xltype & xlbitDLLFree) ? 1 : 0;
    int is_string = (this->xltype & xltypeStr) ? 1 : 0;
    int is_multi = (this->xltype & xltypeMulti) ? 1 : 0;
    //printf("------------ XL is_xl_free %i   is_dll_free %i   is_string %i   is_multi %i \n", is_xl_free, is_dll_free, is_string, is_multi);
    if (is_xl_free) {
        if (is_string) 
            delete[] this->val.str;
        if (is_multi) 
            delete[] (XLOper*)this->val.array.lparray;
    }
    if (is_dll_free) {
        fprintf(stderr, "!!!!!!!!!!!! XL - XLCALL NOT SUPPOSED TO FREE DLL ALLOCATED MEMORY - is_xl_free %i   is_dll_free %i   is_string %i   is_multi %i \n", is_xl_free, is_dll_free, is_string, is_multi);
    }
}

XLOper12::XLOper12() {
    xltype = 0;
}

XLOper12::~XLOper12() {
    int is_xl_free = (this->xltype & xlbitXLFree) ? 1 : 0;
    int is_dll_free = (this->xltype & xlbitDLLFree) ? 1 : 0;
    int is_string = (this->xltype & xltypeStr) ? 1 : 0;
    int is_multi = (this->xltype & xltypeMulti) ? 1 : 0;
    //printf("------------ XL12 is_xl_free %i   is_dll_free %i   is_string %i   is_multi %i \n", is_xl_free, is_dll_free, is_string, is_multi);
    if (is_xl_free) {
        if (is_string) 
            delete[] this->val.str;
        if (is_multi) 
            delete[] (XLOper12*)this->val.array.lparray;
    }
    if (is_dll_free) {
        fprintf(stderr, "!!!!!!!!!!!! XL12 - XLCALL NOT SUPPOSED TO FREE DLL ALLOCATED MEMORY - is_xl_free %i   is_dll_free %i   is_string %i   is_multi %i \n", is_xl_free, is_dll_free, is_string, is_multi);
    }
}
