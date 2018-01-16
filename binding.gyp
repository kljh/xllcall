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
        "src/dyncall/dyncall/dyncall_callvm__auto_arch.c",
        #"src/dyncall/dyncall/dyncall_call_generic_masm__auto_arch.asm",
      ],
      'conditions': [
        [
          # How to test for ia32 vs x64 arch ? !!
          '"x86"=="x86"', 
          {
            "sources": [
              #"src/dyncall/dyncall/dyncall_callvm_x86.c",
              "src/dyncall/dyncall/dyncall_call_x86_generic_masm.asm",
            ],
            # Include ASM file above rather than pre-compiled OBJ file below
            # "libraries": [
            #   "<!(cd)/src/dyncall/dyncall/dyncall_call_x86_generic_masm.asm.obj",
            # ],
            'msvs_settings': {
              # By default, msvs passes /SAFESEH for Link, but not for MASM.
              # In order for test_safeseh_default to link successfully, we need to explicitly specify /SAFESEH for MASM.
              # ia32 only ?
              'MASM': {
                'UseSafeExceptionHandlers': 'true',
              },
            },
          },
          {
            "sources": [
              #"src/dyncall/dyncall/dyncall_callvm_x64.c",
              "src/dyncall/dyncall/dyncall_call_x64_generic_masm.asm",
            ],
            # Include ASM file above rather than pre-compiled OBJ file below
            # "libraries": [
            #   "!(cd)/src/dyncall/dyncall/dyncall_call_x64_generic_masm.asm.obj",
            # ],
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
        "src/xloper.cc",
        "src/xlcall32/hwnd.cpp", 
      ],
    }
  ]
}