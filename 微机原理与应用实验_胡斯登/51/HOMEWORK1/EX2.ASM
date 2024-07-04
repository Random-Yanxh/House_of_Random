ORG     0000h
;在RAM 31H单元存放一组8位带符号数，字节个数放在30H中，
;请编写程序统计出其中正数，负数以及0的数目，结果存放在41H，42H以及43H中。
;##正负判断：利用JB ACC.7,rel这条指令进行负数的判断。
;0的判断：利用CJNE A,#0判断;也可使用 JZ rel判断。两种相反。
LJMP    BEGIN
ORG     0030h
BEGIN:
MOV 41H,#0 ;初始化正数计数器
MOV 42H,#0 ;初始化负数计数器
MOV 43H,#0 ;初始化0计数器

MOV  R1,30H;数据长度，用于控制循环次数
MOV  R0,#31H

LOOP:  ;循环主体
MOV  A,@R0
JB   ACC.7,CNT_NEGA;##本题核心指令，ACC.7是A中数的最高位，存放带符号数
;最高位就是符号位，若为1则为负数，跳转到负数计数器
;否则顺序进行，起到if-elif-else的作用，下一条类似
JZ CNT_ZERO;#判断0，为0跳转，否则就剩正数
;#此处也可采用CJNE A,#0,CNT_POSI,正数跳转，对应着下一条指令要改成INC 43H
INC  41H;正数计数器+1
JMP LP;这条可不要，保留是为了与另两个分支判断格式一致，好看
LP:  INC R0;此处是与C的不同，下面类似调用函数，C是调用，汇编是跳转
;跳转了还得回来，因此配对下面的计数“函数”，相当于回城点。
DJNZ R1,LOOP ;控制循环次数

LJMP STOP;如果这里直接用END，会报错；如果在最后用END
;由于顺序执行，会一直跑计数器加1并且去LP，R1会减到0FF再减，出问题
;因此这里设一个传送点，与下面“函数”定义区隔开
;下面类比为函数定义区
CNT_POSI:INC 41H
AJMP LP;“调用”完了跳转回去，回城
CNT_NEGA:INC 42H
AJMP LP
CNT_ZERO:INC 43H
AJMP LP

STOP:SJMP STOP;正式的结束位置


