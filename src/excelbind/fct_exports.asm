; The code here contains the actual functions exported from the dll.
; They use the thunks_objects and thunks_methods tables to redirect the calls from Excel 
; to the correct method on the PythonFunctionAdapter object wrapping the python function
; The win32 functions are quite simple, as the (internal Visual Studio/Microsoft calling 
; convention for calling a non-virtual method on an object, is the same as stdcall + ptr 
; to the object in ECX. I.e. all function arguments are on the stack.
; The x64 functions are a bit more involved, as the Microsoft x64 calling conventions 
; already use RCX for the first function argument. So we need to shift all function arguments
; and then put the ptr to the object in RCX.
; Note that to simplify the code, we always shift all 10 arguments even if the function actually has less.
; See (https://docs.microsoft.com/en-us/cpp/build/x64-calling-convention?view=vs-2019)
; for details on x64 calling conventions.

IFDEF RAX; small hack - RAX is only defined in x64  - but it doesn't work with IFNDEF!

ELSE
    .model flat, c
ENDIF

.data
EXTERN thunks_objects:PTR PTRTYPE
EXTERN thunks_methods:PTR PTRTYPE

IFDEF RAX; small hack - RAX is only defined in x64
    fct macro i
        PUBLIC f&i
        f&i PROC EXPORT 
            sub    RSP, 058h
    
            mov RAX, QWORD PTR[RSP + 0A8h]
            mov [RSP + 050h], RAX

            mov RAX, QWORD PTR[RSP + 0A0h]
            mov [RSP + 048h], RAX

            mov RAX, QWORD PTR[RSP + 098h]
            mov [RSP + 040h], RAX

            mov RAX, QWORD PTR[RSP + 090h]
            mov [RSP + 038h], RAX

            mov RAX, QWORD PTR[RSP + 088h]
            mov [RSP + 030h], RAX

            mov RAX, QWORD PTR[RSP + 080h]
            mov [RSP + 028h], RAX
            mov [RSP + 020h], R9
            mov R9, R8
            mov R8, RDX
            mov RDX, RCX
            mov RCX, QWORD PTR [thunks_objects + i * 8]
	        mov RAX, thunks_methods + i * 8
            call RAX

            add      RSP, 058h
            ret
        f&i ENDP
    endm
ELSE
    fct macro i
        PUBLIC f&i
        f&i PROC EXPORT 
            mov ECX, DWORD PTR [thunks_objects + i * 4]
	        jmp thunks_methods + i * 4
        f&i ENDP
    endm
ENDIF
.code


; Roll out 10000 functions - note the number of functions here should match the space allocated to 
; the thunks tables in 'function_registration.cpp'

counter = 0
REPEAT 10000
    fct %counter
    counter = counter + 1
ENDM



end