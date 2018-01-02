#include "v8conv.h"
#include <algorithm>

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
v8::Local<v8::Value> xloper_2_v8value_impl(v8::Isolate* isolate, X& xop) {  
  v8::Local<v8::Value> res;
  
  //printf("xloper_2_v8value_impl xltype=0x%x \n", (int)xop.xltype);

	if (!xop.xltype) {
		//res = v8::Undefined();
	} else if (xop.xltype & xltypeNum) {
		res = v8::Number::New(isolate, xop.val.num);
  } else if (xop.xltype & xltypeBool) {
		res = v8::Boolean::New(isolate, xop.val.xbool);
	} else if (xop.xltype & xltypeInt) {
		res = v8::Number::New(isolate, xop.val.w);
	} else if (xop.xltype & xltypeStr) {
    std::string s = xloper_char_ptr_2_string(xop.val.str);
    res = v8::String::NewFromUtf8(isolate, s.c_str(), v8::NewStringType::kNormal, (int)s.length())  // v8::NewStringType::kInternalized
      .ToLocalChecked();
	} else if (xop.xltype & xltypeErr) {
    printf("xloper_2_v8value_impl: error %i.\n", (int)xop.val.err);
    isolate->ThrowException(v8::Exception::TypeError(
      v8::String::NewFromUtf8(isolate, "xloper_2_v8value_impl: error")));
    //res = v8::Undefined();
	} else if (xop.xltype & xltypeNil) {
		//res = v8::Undefined();
	} else if (xop.xltype & xltypeMulti) {
    size_t n = (size_t)xop.val.array.rows;
    size_t m = (size_t)xop.val.array.columns;
  
    printf("xloper_2_v8value_impl multi %i x %i.\n", (int)n, (int)m);
    
    v8::Local<v8::Array> rows = v8::Array::New(isolate, n);
    for (size_t i=0; i<n; i++) {
      v8::Local<v8::Array> row = v8::Array::New(isolate, m);
      for (size_t j=0; j<m; j++) {
        X* xop_ij = &(xop.val.array.lparray[i*m+j]);
        row->Set(j, xloper_2_v8value_impl(isolate, *xop_ij));
      }
      rows->Set(i, row);
    }

    res = rows;
  } else {
    fprintf(stderr, "%s: to complete. xop.xltype=%i.\n", __FUNCTION__, (int)xop.xltype);
		
		throw std::logic_error("xloper_2_v8value to complete");
  }

  return res;
}


v8::Local<v8::Value> xloper_2_v8value(v8::Isolate* isolate, const xloper& x) {
  v8::Local<v8::Value> res = xloper_2_v8value_impl(isolate, x);
  return res;
}
v8::Local<v8::Value> xloper_2_v8value(v8::Isolate* isolate, const xloper12& x) {
  v8::Local<v8::Value> res = xloper_2_v8value_impl(isolate, x);
  return res;
}

// (XLL args) conversion from v8::Value to xloper

bool v8value_2_native(v8::Isolate* isolate, const v8::Local<v8::Value>& v, std::string& x)
{
  if (!v->IsString()) {
    isolate->ThrowException(v8::Exception::TypeError(
      v8::String::NewFromUtf8(isolate, "v8 value not a string")));
    return false;
  }

  x = std::string(*v8::String::Utf8Value(v));
  return true;
}

bool v8value_2_native(v8::Isolate* isolate, const v8::Local<v8::Value>& v, std::wstring& x)
{
  if (!v->IsString()) {
    isolate->ThrowException(v8::Exception::TypeError(
      v8::String::NewFromUtf8(isolate, "v8 value not a wstring")));
    return false;
  }

  std::string s(*v8::String::Utf8Value(v));
  x = std::wstring(s.begin(), s.end());
  return true;
}

bool v8value_2_native(v8::Isolate* isolate, const v8::Local<v8::Value>& v, double& x)
{
  if (!v->IsNumber()) {
    isolate->ThrowException(v8::Exception::TypeError(
      v8::String::NewFromUtf8(isolate, "v8 value not a double")));
    return false;
  }

  // NumberValue --OR-- IntegerValue
  x = v->NumberValue();
  return true;
}

bool v8value_2_native(v8::Isolate* isolate, const v8::Local<v8::Value>& v, bool& x)
{
  if (!v->IsBoolean()) {
    isolate->ThrowException(v8::Exception::TypeError(
      v8::String::NewFromUtf8(isolate, "v8 value not a boolean")));
    return false;
  }

  x = v->BooleanValue();
  return true;
}

bool v8value_2_native(v8::Isolate* isolate, const v8::Local<v8::Value>& v8_val, std::vector<std::string>& vx)
{
  if (!v8_val->IsArray()) {
    isolate->ThrowException(v8::Exception::TypeError(
      v8::String::NewFromUtf8(isolate, "v8 value not an array")));
    return false;
  }

  v8::HandleScope scope(isolate);
  v8::Local<v8::Array> v8_array = v8::Handle<v8::Array>::Cast(v8_val);

  size_t n = v8_array->Length();
  vx.resize(n);
  for (size_t i=0; i<n; i++) {
    v8::Local<v8::Value> v8_element = v8_array->Get(i);
    
    if (!v8value_2_native(isolate, v8_element, vx[i])) return false;
  }

  return true;
}

v8::Local<v8::Array> v8Value_to_v8array(v8::Isolate* isolate, size_t arg_pos, const v8::Local<v8::Value>& v) {
  if (!v->IsArray()) {
    isolate->ThrowException(v8::Exception::TypeError(
      v8::String::NewFromUtf8(isolate, "v8Value_to_v8array: v8 value not an array")));
    return v8::Local<v8::Array>();
  }
  return v8::Local<v8::Array>::Cast(v);
}

bool v8value_2_xloper_scalar(v8::Isolate* isolate, size_t arg_pos, const v8::Local<v8::Value>& v8_val, XLOper& x)
{
  if (v8_val->IsString()) {
    std::string s;
    if (!v8value_2_native(isolate, v8_val, s)) return false;
    size_t n = std::min(s.size(), (size_t)254);

    x.xltype = xltypeStr | xlbitDLLFree;
    x.val.str = new char[n+2];
    x.val.str[0] = n;
    x.val.str[n+1] = '\0';
    memcpy(x.val.str+1, s.c_str(), n);
    
  } else if (v8_val->IsNumber()) {
    double d;
    if (!v8value_2_native(isolate, v8_val, d)) return false;

    x.xltype = xltypeNum;
    x.val.num = d;

  } else if (v8_val->IsBoolean()) {
    bool b;
    if (!v8value_2_native(isolate, v8_val, b)) return false;

    x.xltype = xltypeBool;
    x.val.xbool = b;

  } else {
    isolate->ThrowException(v8::Exception::TypeError(
      v8::String::NewFromUtf8(isolate, "v8value_2_xloper_scalar: v8 value not a scalar")));
    return false;
  } 
  return true;
}

bool v8value_2_xloper_array(v8::Isolate* isolate, size_t arg_pos, const v8::Local<v8::Value>& v8_val, XLOper& x)
{
  v8::Local<v8::Array> v8_rows = v8Value_to_v8array(isolate, arg_pos, v8_val);
  if (v8_rows.IsEmpty()) return false;

  size_t nb_rows = v8_rows->Length();
  size_t nb_cols = 0;
  for (size_t i=0; i<nb_rows; i++) {
    v8::Local<v8::Array> v8_row = v8Value_to_v8array(isolate, arg_pos, v8_rows->Get(0));
    if (v8_row.IsEmpty()) { printf("sub-element %i is not an array.\n", (int)i); return false; }

    nb_cols = std::max((size_t)v8_rows->Length(), nb_cols);      
  }
  //printf("v8value_2_xloper_array: %i x %i.\n", (int)nb_rows, (int)nb_cols);

  x.xltype = xltypeMulti | xlbitDLLFree;
  x.val.array.rows = nb_rows;
  x.val.array.columns = nb_cols;
  x.val.array.lparray = new XLOper[nb_rows*nb_cols];
  
  // initialise array with sensible default values
  for (size_t i=0, k=0; i<nb_rows; i++)
    for (size_t j=0; j<nb_cols; j++, k++)
      x.val.array.lparray[k].xltype = xltypeNil;

  for (size_t i=0; i<nb_rows; i++) {
    v8::Local<v8::Array> v8_row = v8Value_to_v8array(isolate, arg_pos, v8_rows->Get(i));
    for (size_t j=0; j<nb_cols; j++) {
      v8::Local<v8::Value> v8_cell = v8_row->Get(j);
      if (!v8value_2_xloper_scalar(isolate, arg_pos, v8_cell, (XLOper&)x.val.array.lparray[i*nb_cols+j])) return false;
    }
  }
  return true;
}

bool v8value_2_xloper(v8::Isolate* isolate, size_t arg_pos, const v8::Local<v8::Value>& v8_val, XLOper& x)
{
  if (v8_val->IsArray()) 
    return v8value_2_xloper_array(isolate, arg_pos, v8_val, x);
  else
    return v8value_2_xloper_scalar(isolate, arg_pos, v8_val, x);
}

bool v8value_2_xloper(v8::Isolate* isolate, size_t arg_pos, const v8::Local<v8::Value>& v8_val, XLOper12& x)
{
  printf("v8value_2_xloper12: %i", (int) x.xltype);
  return true;
}



// conversion from xloper

/*
template <class X>
void xloper_2_any(X* xop, AnyVal& a) {
	if (!xop.xltype) {
		a =  AnyVal(); // ret val
	} else if (xop->xltype & xltypeNum) {
		a = xop->val.num;
	} else if (xop->xltype & xltypeBool) {
		a =  xop->val.xbool;
	} else if (xop->xltype & xltypeInt) {
		a =  xop->val.w;
	} else if (xop->xltype & xltypeStr) {
		a =  xloper_char_ptr_2_string(xop->val.str);
	} else if (xop->xltype & xltypeErr) {
		a =  AnyVal();
	} else if (xop->xltype & xltypeNil) {
		a = AnyVal();
	} else {
		throw std::logic_error("XLOper parse to complete");
	}

}
*/