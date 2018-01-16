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

// (XLL result) conversion from xloper to v8::Value

namespace {
std::string xloper_char_ptr_2_string(const char * pc) {
    std::string s(pc+1, pc+1 + (unsigned char)pc[0]); // kljh: unsigned
    return s;
}}

namespace {
std::string xloper_char_ptr_2_string(const XCHAR *pc) {
    std::string s(pc+1, pc+1 + pc[0]); // kljh: XCHAR is unsigned short (so already unsigned)
    return s;
}}

template <class X>
bool xloper_2_native_scalar(const X& xop, std::string& t) {
    if (xop.xltype & xltypeStr) {
        t = xloper_char_ptr_2_string(xop.val.str);
    } else {
        return false;
    }
    return true;
}

template <class X>
bool xloper_2_native_scalar(const X& xop, double& t) {
    if (xop.xltype & xltypeNum) {
        t = xop.val.num;
    } else if (xop.xltype & xltypeBool) {
        t = double(xop.val.xbool);
    } else if (xop.xltype & xltypeInt) {
        t = double(xop.val.w);
    } else {
        return false;
    }
    return true;
}

template <typename T, class X>
bool xloper_2_native(const X& xop, T& t) {
    if (xop.xltype & xltypeMulti) {
        size_t n = (size_t)xop.val.array.rows;
        size_t m = (size_t)xop.val.array.columns;
        
        if (n==1 && m==1) { 
            size_t i=0, j=0;
            X* xop_ij = (X*) &(xop.val.array.lparray[i*m+j]);
            return xloper_2_native_scalar(*xop_ij, t);
        } else {
            return false;
        }
    } else {
        return xloper_2_native_scalar(xop, t);
    }
}

XLOper::operator std::string() const {
    std::string x;
    bool bOk = xloper_2_native(*this, x);
    return x;
}

XLOper12::operator std::string() const {
    std::string x;
    bool bOk = xloper_2_native(*this, x);
    return x;
}
