; ******************************************************
; BASIC .ASM template file for AVR
; ******************************************************

.include "C:\VMLAB\include\m8def.inc"

; Define here the variables
;

.equ first =0xa5e7
.equ second = 0xc3d8
.def x1 = r16
.def x2 = r17
.def y1 = r18
.def y2 = r19
.def z1 = r20
.def z2 = r21
.def z3 = r22
.def z4 = r23
.def i1 = r25
.def i2 = r26
.def i3 = r27
.def f1 = r29
.def f2 = r30
.def f3 = r31

; Define here Reset and interrupt vectors, if any
;
reset:
   rjmp start
   reti      ; Addr $01
   reti      ; Addr $02
   reti      ; Addr $03
   reti      ; Addr $04
   reti      ; Addr $05
   reti      ; Addr $06        Use 'rjmp myVector'
   reti      ; Addr $07        to define a interrupt vector
   reti      ; Addr $08
   reti      ; Addr $09
   reti      ; Addr $0A
   reti      ; Addr $0B        This is just an example
   reti      ; Addr $0C        Not all MCUs have the same
   reti      ; Addr $0D        number of interrupt vectors
   reti      ; Addr $0E
   reti      ; Addr $0F
   reti      ; Addr $10

; Program starts here after Reset
;
start:
    ldi x1, low(RAMEND)     ;Íàñòðîéêà ñòåêà
    out SPL, x1
    ldi x1, high(RAMEND)
    out SPH, x1
    
    
    ldi x1, high(first)    ;
   ldi x2, low(first) ; cleanup RAM, etc.
    ldi y1, high(second)     ; Initialize here ports, stack pointer,
   ldi y2, low(second) ;


    ldi z1, 0
    ldi z2, 0
    ldi z3, 0
    ldi z4, 0
    ldi i1, 0
    ldi i2, 0
    ldi i3, 0

mul_for_numbers:
   mul x2, y2
   add z4, r0
   add z3, r1

   mul x1, y2
   add z3, r0
   adc z2, r1

   mul x2, y1
   add z3, r0
   adc z2, r1
   adc z1, z1

   mul x1, y1
   add z2, r0
   adc z1, r1
   nop       ; Infinite loop.
   nop       ; Define your main system
   nop       ; behaviour here
summ:
    add i3, x2
    add i3, y2
    adc i2, x1
    add i2, y1
    adc i1, i1

sub numbers:
mov f2, xl ; Копируем старший байт первого числа в i2
mov f3, x2 ; Копируем младший байт первого числа в із
sub f3, v2 ; Вычитаем младший байт второго числа
sbc f2, vl ; Вычитаем старший байт с учетом заема
1di f1,    ; Обнуляем i1 (для расширения знака)
sbc fl, f1 ; Распространяем заем на i1 (0 если нет заема, ОхFF если был заем)


