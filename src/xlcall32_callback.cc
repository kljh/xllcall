#include "xllcall.h"
#include "xloper.h"
#include "v8conv.h"
#include <dyncall/dyncall.h>
#include <dyncall/dyncall_signature.h>

#include <excpt.h>
#include <stdio.h>
#include <map>

template <class X>
int excel_callback(int xlfn, X* ret, int n, X* args[]) {	
	return xlretSuccess;
}

int excel4v_callback(int xlfn, LPXLOPER ret, int n, LPXLOPER args[]) { 
	return excel_callback(int xlfn, LPXLOPER ret, int n, LPXLOPER args[]);
}
