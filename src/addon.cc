#include <node.h>
#include "xllcall.h"

#include <stdio.h>

namespace xllcall {

void init(v8::Local<v8::Object> exports) {
    NODE_SET_METHOD(exports, "xllcall_debug_v8", xllcall_debug_v8);
    NODE_SET_METHOD(exports, "xllload_xlAutoOpen_v8", xllload_xlAutoOpen_v8_seh);
    NODE_SET_METHOD(exports, "xllcall_ffi_v8", xllcall_ffi_v8_seh);
}

NODE_MODULE(NODE_GYP_MODULE_NAME, init)

}