IFDEF RAX

  ECHO "MASM x64"
  INCLUDE <dyncall_call_x64_generic_masm.asm>

ELSE

  ECHO "MASM x86"
  INCLUDE <dyncall_call_x86_generic_masm.asm>

ENDIF
