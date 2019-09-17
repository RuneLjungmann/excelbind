
IFDEF RAX; small hack - RAX is only defined in x64
    PTRTYPE TEXTEQU <QWORD>
    OBJECT_REG TEXTEQU <rcx>
ELSE
    .model flat, c

    PTRTYPE TEXTEQU <DWORD>
    OBJECT_REG TEXTEQU <ecx>
ENDIF


.data
EXTERN thunks_objects:PTR PTRTYPE
EXTERN thunks_methods:PTR PTRTYPE

fct macro i
PUBLIC f&i
    f&i PROC EXPORT 
        mov OBJECT_REG, PTRTYPE ptr [thunks_objects + (SIZEOF PTRTYPE) * i]
	    jmp thunks_methods + (SIZEOF PTRTYPE) * i
    f&i ENDP
endm

.code

fct 0
fct 1
fct 2
fct 3
fct 4
fct 5
fct 6
fct 7
fct 8
fct 9

end