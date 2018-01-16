#ifdef _WIN64
#include "txt/dyncall_callvm_x64.c"
#else
#include "dyncall_callvm_x86.c"
#endif
