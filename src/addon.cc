#include <node.h>
#include "xllcall.h"

#include <stdio.h>

namespace xllcall {

void init(v8::Local<v8::Object> exports) {
  NODE_SET_METHOD(exports, "xllcall_ffi_v8", xllcall_ffi_v8);
}

NODE_MODULE(NODE_GYP_MODULE_NAME, init)

}