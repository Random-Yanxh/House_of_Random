ORG 0000H
SJMP BEGIN
ORG 0020H
BEGIN:
; 初始化寄存器
MOV 31H,#63H
MOV 32H,#82H
MOV 33H,#0C5H
MOV 34H,#36H

MOV R0, 31H
MOV A,R0
XRL 32H,A  ;  A XOR B,结果放在32H中

; 计算 NOT C
MOV A,33H
CPL A
ANL 32H,A;
MOV A,32H
CPL A
MOV 32H,A;前面一部分算完放在32H备用

MOV A,34H
ORL A,31H
CPL A ;算完或非
ANL A,32H ;最后与，直接留在A里
MOV 40H,A  ;以16进制呈现在40H单元
;逐位显示部分 ，实现了将40H单元字节逐位依次存放在41H~48H单元中
MOV R0,#41H;起始地址放入R0
MOV R1,#8 ;一个字节，8位 ，控制循环次数
LOOP:
RLC A ;带位左移，这样高位先呈现，符合从左到右的顺序 ,并且最高位放入CY中
JC PUT1 ;如果A的最高位为1，那么这里会直接跳转到PUT1
MOV @R0,#0 ;相当于else，最高位是0就在R0存放的地址中放0
LP:INC R0 ;##这里借鉴EX2的想法，是一个给PUT1的传送点，指令本身实现“逐位”
DJNZ R1,LOOP  ;控制循环次数
HERE:SJMP HERE  ;结束指令，和“函数”定义区隔开

PUT1: ;用来放1
MOV @R0,#1
AJMP LP ;回城，不然会乱跑

