{
  "targets": [
    {
      "target_name": "addon",
      "include_dirs": [ "src/dyncall" ],
      "sources": [ 
        "src/addon.cc", 
        "src/xllcall.cc", 
        "src/xloper.cc",
        "src/v8conv.cc",
        "src/dyncall/dyncall/dyncall_api.c",
        "src/dyncall/dyncall/dyncall_struct.c",
        "src/dyncall/dyncall/dyncall_vector.c",
        "src/dyncall/dyncall/dyncall_callvm_base.c",
        "src/dyncall/dyncall/dyncall_callvm_x64.c",
      ],
      'conditions': [
        [
          'OS=="win_x86"', 
          {
            "sources": [
              "src/dyncall/dyncall/dyncall_callvm_x86.c",
            ],
            "libraries": [
              "C:\\Users\\kljh\\Documents\\Code\\GitHub\\xllcall\\src/dyncall/dyncall/dyncall_call_x86_generic_masm.asm.obj",
            ]
          },
          {
            "sources": [
              "src/dyncall/dyncall/dyncall_callvm_x64.c",
            ],
            "libraries": [
              "C:\\Users\\kljh\\Documents\\Code\\GitHub\\xllcall\\src/dyncall/dyncall/dyncall_call_x64_generic_masm.asm.obj",
            ]
          },
        ], 
      ],
      
    }, 
    {
      "target_name": "xlcall32",
      "type": "shared_library",
      "include_dirs": [ "src/dyncall" ],
      "sources": [ 
        "src/xlcall32/xlcall32.def", 
        "src/xlcall32/xlcall32.cpp", 
        "src/xlcall32/hwnd.cpp", 
      ],
 
    }
  ]
}