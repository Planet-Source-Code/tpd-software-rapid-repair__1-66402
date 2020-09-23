;
; Rapid Repair assembler code
;
SEGMENT code USE32

;Dll entry point
GLOBAL _DllMain
_DllMain:
  	mov	eax, 0ffffffffh
        retn    12


;Sub XOR_Buffer (3 LONG Arguments)
GLOBAL fix_xorbuffer
fix_xorbuffer:
	enter	0, 0
        push    edi
        push    esi
        push    eax
        push    ebx
        push    ecx 
	mov	eax,[ebp+8]	;pointer to orgBuffer(0)
	mov	esi,[ebp+12]	;pointer to xorBuffer(0)
	mov	ecx,[ebp+16]	;Size
        xor     ebx,ebx
_xor_loop:
        mov     bl,[eax]        ;bl=orgBuffer(eax)
        xor	[esi],bl 	;xorBuffer(esi)=xorBuffer(esi) xor bl
	inc     eax
        inc     esi
        loopnz  _xor_loop 
	pop     ecx
        pop     ebx
        pop     eax
        pop     esi
        pop     edi 
        leave
        retn    12              ;return, with 12 bytes of arguments (3 DWords)



;Sub MUL_Buffer (6 LONG Arguments)
GLOBAL fix_mulbuffer
fix_mulbuffer:
        enter   0, 0
        push    edi
        push    esi
        push    eax
        push    ebx
        push    ecx
        push    edx
        mov     eax,[ebp+8]       ; ptr to inBuffer(0)
        mov     edx,[ebp+12]      ; ptr to outBuffer(0)
        mov     esi,[ebp+16]      ; ptr to B_TO_J(0)
        mov     edi,[ebp+20]      ; ptr to J_TO_B(0)
        mov     ecx,[ebp+24]      ; Size
	xor     ebx,ebx
_mul_next:
        mov     bl,[eax]          ; bl = inBuffer(eax)
        cmp     ebx,0             ; if bl=0
        je      _mul_skip_0       ;    skip back 
        mov     ebx,[esi+ebx*4]   ; ebx = B_TO_J(bl*4)
        add     ebx,[ebp+28]      ; ebx = ebx + flog
        mov     ebx,[edi+ebx*4]   ; ebx = J_TO_B(ebx*4)  
        mov     [edx],bl          ; outBuffer(edx) = bl
        xor     ebx,ebx
_mul_skip_0:
        inc     eax               
        inc     edx
        loopnz  _mul_next         ;    next byte 
_mul_done:
        pop     edx
        pop     ecx
        pop     ebx
        pop     eax
        pop     esi
        pop     edi
        leave
        retn    24


;Sub fix_emptybuffer (2 LONG Arguments)
GLOBAL fix_emptybuffer
fix_emptybuffer:
        enter 0, 0
        push    eax
        push    ebx
        push    ecx
        mov     eax,[ebp+8]       ; ptr to emptyBuffer(0)
        mov     ecx,[ebp+12]      ; Size
        xor     ebx, ebx
_empty_next:
        mov     [eax],ebx
        inc     eax
        loopnz  _empty_next
        pop     ecx
        pop     ebx
        pop     eax
        leave
        retn 8


;Sub fix_crc32buffer (4 LONG Arguments)
GLOBAL fix_crc32buffer
fix_crc32buffer:
        enter  0, 0 
        push   edi
        push   esi
        push   eax
        push   ebx
        push   ecx
        mov    eax,[ebp+8]	;ptr to CRC value
        mov    eax,[eax]	;CRC value
        mov    esi,[ebp+12]	;ptr to Buffer(0)
        mov    edi,[ebp+16]	;ptr to m_Table(0)
        mov    ecx,[ebp+20]	;Size
        xor    ebx, ebx
_crc_next:
	mov    bl,[esi]		;bl = Buffer(esi)
	xor    bl,al		;bl = (bl xor CRC) and 255
	shr    eax,8		;CRC = CRC / 256
	xor    eax,[edi+ebx*4]  ;CRC = CRC xor m_Table(bl)
	inc    esi
	loopnz _crc_next
        mov    ecx,[ebp+8]
        mov    [ecx],eax        ;return CRC 
        pop    ecx
        pop    ebx
        pop    eax
        pop    esi
        pop    edi
        leave  
        retn   16

	ENDS