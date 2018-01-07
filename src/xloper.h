#pragma once

#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include "xlsdk97/xlcall.h"

struct XLOper : public xloper {
    // scalar constructors
    ~XLOper();
    XLOper();

    /*
    XLOper(const XLOper&);
    XLOper(const XLOper&&);
    
    XLOper(const std::string&);
    XLOper(const std::wstring&);
    XLOper(double);
    XLOper(int);
    XLOper(bool);
    // array constructor and accessor
    XLOper(size_t w, size_t h);
    XLOper& item(size_t i, size_t j);
    const XLOper& item(size_t i, size_t j) const;

    operator std::string();
    operator std::wstring();
    operator double();
    operator int();
    operator bool();
    */
};

struct XLOper12 : public xloper12 {
    ~XLOper12();
    XLOper12();
};
