ORG 0000H
SJMP MAIN
ORG 0030H
MAIN:
MOV R0, #40H  ;被加数首地址，从个位开始加
MOV R1, #50H  ;加数首址
MOV R2, 30H   ;循环变量，逐次递减用于判断是否加完每一位
INC R2
CLR C
PP：            ;循环，计算每一位相加后的结果，C为进位标志
MOV A, @R1   ;取加数(50H)
ADDC A,@ R0    
DA A          ;进制转换    
MOV 20H, A   ;存储
CLR C
SUBB A,#10H
JNC NEXT
CLR C
MOV @ R0,20H  ;A＜10，取原数
LCALL FINAL
NEXT:   ;A＞10
MOV @R0,A
SETB C
FINAL:
INC R0
INC R1
DJNZ R2,PP
AJMP $
END
