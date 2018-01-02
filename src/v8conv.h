#pragma once
#include <node.h>
#include "xloper.h"


// inherit plain-old-data structs to add destructors

/*

__filename and __dirname
require('path').basename(__dirname);
process.cwd

        Handle<Object> Result = Object::New();
        Result->Set(String::New("javascript_tagged"), Number::New(317566));

The relevant functions to create new string objects are:
String::NewFromUtf8 (UTF-8 encoded, obviously)
String::NewFromOneByte (Latin-1 encoded)
String::NewFromTwoByte (UTF-16 encoded)

Alternatively, you can avoid copying the string data and construct a V8 string object that refers to existing data (whose lifecycle you control):
String::NewExternalOneByte (Latin-1 encoded)
String::NewExternalTwoByte (UTF-16 encoded)

v8::String::New((uint16_t*)path, wcslen(path))

*/

// (XLL result) conversion from xloper to v8::Value

v8::Local<v8::Value> xloper_2_v8value(v8::Isolate* isolate, const xloper& x);
v8::Local<v8::Value> xloper_2_v8value(v8::Isolate* isolate, const xloper12& x);

// (XLL args) conversion from v8::Value to xloper

bool v8value_2_native(v8::Isolate* isolate, const v8::Local<v8::Value>& v, std::string& x);
bool v8value_2_native(v8::Isolate* isolate, const v8::Local<v8::Value>& v, std::wstring& x);
bool v8value_2_native(v8::Isolate* isolate, const v8::Local<v8::Value>& v, double& x);
bool v8value_2_native(v8::Isolate* isolate, const v8::Local<v8::Value>& v, bool& x);
bool v8value_2_native(v8::Isolate* isolate, const v8::Local<v8::Value>& v, int& x);
bool v8value_2_native(v8::Isolate* isolate, const v8::Local<v8::Value>& v, std::vector<std::string>& x);

bool v8value_2_xloper(v8::Isolate* isolate, size_t arg_pos, const v8::Local<v8::Value>& v, XLOper& x);
bool v8value_2_xloper(v8::Isolate* isolate, size_t arg_pos, const v8::Local<v8::Value>& v, XLOper12& x);

bool v8value_2_xloper_vector(const v8::Value& v, std::vector<XLOper>& x);
