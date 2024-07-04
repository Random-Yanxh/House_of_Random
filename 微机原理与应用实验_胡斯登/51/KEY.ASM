ORG     0000H
SJMP     MAIN
ORG     0080H

MAIN:;数码管显示
MOV  DPTR,#KEY_TAB
LCALL KEY
MOVC  A,@A+DPTR
MOV   P0,A
CLR   P2.1
LJMP  MAIN

KEY:;按键去抖
LCALL  KS
JNZ    K1;有键按下继续
LCALL  DELAY2 ;无键按下调用延时去抖程序
AJMP   KEY ;返回继续监测有无按键按下
K1:;  键值判断（行扫描）
LCALL   DELAY2 ;有按键按下继续延时去抖动
LCALL   DELAY2 ;
LCALL   KS ; 按键检测
JNZ     K2  ; 确认有按下进行下一步
AJMP    KEY  ;无键按下返回继续监测
K2:
MOV     42H,#0EFH
MOV     43H,#00H
K3:
MOV     P1,42H
L6:
JB      P1.0,L1 ;P1.0=1则跳转L1
MOV     A,#00H
AJMP    LK
L1:
JB      P1.1,L2
MOV     A,#04H
AJMP    LK
L2:
JB      P1.2,L3
MOV     A,#08H
AJMP    LK
L3:
JB      P1.3,NEXT
MOV     A,#0CH
LK:
ADD     A,43H
PUSH    ACC
K4:
LCALL   DELAY2
LCALL   KS
JNZ     K4
POP     ACC
RET

KS:; 按下扫描
MOV  P1,#0FH
MOV  A,P1
XRL  A,#0FH
RET

NEXT:;列扫描
INC  43H
MOV  A,42H
JNB  ACC.7,KEY
RL   A
MOV  42H,A
AJMP K3

DELAY2:
MOV  R0,#08H
L7:  MOV  R7,#0FAH
L8:  DJNZ R7,L8
     DJNZ R0,L7
     RET

KEY_TAB:;键值与显示
DB   28H,34H,28H,34H,0A9H,60H,20H,7AH,20H,21H,61H,74H,30H,62H,0A2H,7EH
;0H0HC9878654A321轮流显示键盘因为无法表达*#就用H表示，B用8表示?
;FE0C 8A09 8765 4321
;DB  FH,EH,28H,CH, 20H,AH,28H,60H, 20H,7AH,21H,61H, 74H,62H,0A2H,7EH 



