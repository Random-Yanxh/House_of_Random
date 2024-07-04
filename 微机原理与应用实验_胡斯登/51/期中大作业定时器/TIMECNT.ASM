;*******功能*
;1. 初始归零，归零时按开始正计时
;2. 10min(02H)   1min(03H)   10s(04H)
;    归零（0FH） 开始/暂停（0EH）
;3. 仅暂停时候可以改时间
;4. 10秒无操作自动息屏
;5. 计时结束蜂鸣器响，按归零停止响并恢复初始状态
;
;
;************
        B_WORK  BIT     00H	;0-没有进行计时；1-在进行计时工作
        B_2ms     BIT	 01H	;2ms定时时间标记
        ;B_TIME_N    BIT     02H     ;倒计时
        B_TIME_P   BIT     03H      ;=1正计时,=0倒计时
        B_SLEEP   BIT     05H   ;0-正常；1-无操作满10s进入息屏
        B_TMOVER    BIT   06H   ;0-计时没有结束；1-计时结束，随后蜂鸣器响
        B_KEYDO     BIT   07H   ;0-没有有效按键输入；1-有有效按键

;;内部RAM空间
        LED_BUF  EQU     28H          ;;LED显示缓存28H~2FH8个地址，存放显示的内容

        DELH     EQU     30H          ; 用于小延迟
        DELL     EQU     31H          ; 用于小延迟
;;按键相关寄存器
        KEYCODE  EQU     32H          ; 实际使用键码
        LASTCODE EQU     33H          ; 上次采集的键码，用于防抖
        N_SAME   EQU     34H          ; 重复按键次数，用于防抖
        N_NOKEY  EQU     37H
;;显示相关
        TEMP     EQU     35H
        N_DISP   EQU     36H          ; 配合使用，第n个灯


;;时钟相关寄存器
        T_2MS    EQU    40H             ;2ms定时计数
        T_100MS  EQU    41H             ;100ms定时计数;正计时的时候用
        T_2MS_N  EQU    45H
        T_100MS_N EQU   46H             ;倒计时的时候用
        T_2MS_SLEEP EQU 47H
        T_100MS_SLEEP EQU 48H           ;无操作时间，息屏判断时候用

        SECOND  EQU     42H     ;计数，也是存放对应的时间值
        MINUTE  EQU     43H
        HOUR    EQU     44H


        ORG     0000H
        LJMP    BEGIN

        ORG     000BH
        LJMP    TIMER0_INT
;R0     =       LED地址寄存器

        ORG     0100H
BEGIN:
        MOV     SP,#0E0H
;工作状态初始化
        CLR     B_WORK
        CLR     B_SLEEP
        SETB    B_TIME_P;初始是正计时
        ;CLR     B_TIME_N
        CLR     B_KEYDO
        CLR     B_TMOVER
;;按键初始化
        MOV     KEYCODE,#0FFH
        MOV     LASTCODE,#0FFH
        MOV     N_SAME,#0

;;LED显示初始化
        MOV     R1,#8;8个灯动态显示扫描
        MOV     R0,#LED_BUF;8个地址，到哪算哪

CLRLOOP:
        MOV     @R0,#0FH;
        INC     R0
        DJNZ    R1,CLRLOOP;8次一循环
;;显示时间值初始化
        LCALL   TIMER_INI

LOOP:;主
        LCALL   TIMER_P          ;正计时
        LCALL   TIMER_N          ;倒计时
        LCALL   DISP_TIMER       ;时钟放入显示缓冲
        LCALL   KEY_SCAN         ;键盘扫描
        LCALL   TIMER_SLEEP     ;休眠时间检测
        LCALL   DO_KEY           ;根据键码值运行相应操作
        LCALL   TIMER_N_TF      ;倒计时结束判断
        LCALL   DISP_LED         ;LED显示
        LCALL   TMOVER           ;计时结束及蜂鸣器

        JNB      B_2ms,$
        CLR      B_2ms
        LJMP    LOOP
;************定时器初始化************
TIMER_INI:
        MOV     T_2MS,#0
        MOV     T_100MS,#0
        MOV     T_2MS_N,#0
        MOV     T_100MS_N,#0
        MOV     T_2MS_SLEEP,#0
        MOV     T_100MS_SLEEP,#0
        MOV     SECOND,#0
        MOV     MINUTE,#00
        MOV     HOUR,#00
        MOV     TMOD,#01H               ;TIMER0 MODE1
        MOV     TCON,#0
        MOV     TH0,#0F8H
        MOV     TL0,#30H               ;65536-2000=63536=F830H ,2ms单元
        SETB    TR0
        SETB    ET0
        SETB    EA               ;计时器的基本设置
        RET
;定时器中断
TIMER0_INT:
        MOV     TH0,#0F8H
        MOV     TL0,#30H               ;2mS
        SETB       B_2ms               ;2ms计时标志
        RETI

;;内部时钟运行
;;INPUT:NONE
;;OUTPUT:TIP,SECOND,MINUTE,HOUR
;;INFLUENCED:ACC
;****************正计时和倒计时处理
TIMER_P:;Positive 正计时
        JNB      B_WORK,TIMER_P_RET;B_WORK=1 -->
        JNB      B_TIME_P,TIMER_P_RET;正计时
        INC     T_2MS ;;2ms计数加1
        MOV     A,T_2MS;放入A中方便实用SUBB 或者 CJNE
        CLR     C       ;用SUBB之前一定要先把C清零，否则影响判断
        SUBB    A,#50   ;也可以用CJNE
        ;IF T_2MS<50
        JC      TIMER_P_RET;T_2MS<50时，小减大C=1，跳转RET,计时继续滚动
        ;IF T_2MS=50
        MOV     T_2MS,A;T_2MS=50时，减完=0，把0重新放回T_2MS，计数器重置，也可以直接放入#0
;100ms计数加1
        INC     T_100MS
        MOV     A,T_100MS
        CLR     C
        SUBB    A,#10
        JC      TIMER_P_RET
        MOV     T_100MS,A;与前面同理
;秒计数加1
        INC     SECOND
        MOV     A,SECOND
        CLR     C
        SUBB    A,#60;60S
        JC      TIMER_P_RET
        MOV     SECOND,A
;;分钟计数加1
        INC     MINUTE
        MOV     A,MINUTE
        CLR     C
        SUBB    A,#99;
        JC      TIMER_P_RET
        MOV     MINUTE,A
;;小时计数加1
        ;MOV     A,HOUR

TIMER_P_RET:
        RET

;倒计时

TIMER_N:
        JNB     B_WORK,TIMER_N_RET
        JB     B_TIME_P,TIMER_N_RET
        INC     T_2MS_N ;;2ms计数加1
        MOV     A,T_2MS_N;放入A中方便实用SUBB 或者 CJNE
        CLR     C       ;用SUBB之前一定要先把C清零，否则影响判断
        SUBB    A,#50   ;也可以用CJNE
        ;IF T_2MS<50
        JC      TIMER_N_RET;T_2MS<50时，小减大C=1，跳转RET,计时继续滚动
        ;IF T_2MS=50
        MOV     T_2MS_N,A;T_2MS=50时，减完=0，把0重新放回T_2MS，计数器重置，也可以直接放入#0
;;100ms计数加1
        INC     T_100MS_N
        MOV     A,T_100MS_N
        CLR     C
        SUBB    A,#10
        JC      TIMER_N_RET
        MOV     T_100MS_N,A;与前面同理
;;秒计数-1
TIMER_N_SEC:

        MOV     A,SECOND
        ;CJNE    A,#0,TIMER_
        CLR     C
        SUBB    A,#1
        JC      TIMER_N_MIN;减到0跳转减分
        MOV     SECOND,A
        SJMP    TIMER_N_RET
        ;CJNE    A,#0,TIMER_N_RET
        ;SJMP    TIMER_N_MIN

TIMER_N_MIN:
        DEC     MINUTE
        MOV     A,MINUTE
        MOV     SECOND,#59;没结束，秒59
        CJNE    A,#0,TIMER_N_MIN1
        JB      B_TMOVER,TMOVER
        LJMP    TIMER_N_RET
        ;CLR     C
        ;SUBB    A,#1

        ;MOV     MINUTE,A

        ;JC      TIMER_N_MIN1;分到0，
        ;LJMP    TIMER_N_RET
TIMER_N_MIN1:
        ;JB      B_TMOVER,TMOVER
        LJMP    TIMER_N_RET

;TIMER_N_ZERO:
        ;MOV     SECOND,#0
        ;MOV     MINUTE,#0
        ;RET

TIMER_N_RET:
        RET
;*******无操作时间
TIMER_SLEEP:;
        JB      B_SLEEP,TIMER_S_RET0
        JB      B_KEYDO,TIMER_S_RET0
        JB      B_WORK,TIMER_S_RET0

        ;B_WORK=1 -->
        ;JNB      B_TIME_P,TIMER_S_RET
        ;正计时
        INC     T_2MS_SLEEP ;;2ms计数加1
        MOV     A,T_2MS_SLEEP;放入A中方便实用SUBB 或者 CJNE
        CLR     C       ;用SUBB之前一定要先把C清零，否则影响判断
        SUBB    A,#50   ;也可以用CJNE
        ;IF T_2MS<50
        JC      TIMER_S_RET;T_2MS<50时，小减大C=1，跳转RET,计时继续滚动
        ;IF T_2MS=50
        MOV     T_2MS_SLEEP,A;T_2MS=50时，减完=0，把0重新放回T_2MS，计数器重置，也可以直接放入#0

        INC     T_100MS_SLEEP
        MOV     A,T_100MS_SLEEP
        CLR     C
        SUBB    A,#100
        JC      TIMER_S_RET
        MOV     T_100MS_SLEEP,A;与前面同理
        SETB    B_SLEEP
        MOV     T_100MS_SLEEP,#0
        MOV     T_2MS_SLEEP,#0
TIMER_S_RET:
        RET
TIMER_S_RET0:
        MOV     T_100MS_SLEEP,#0
        MOV     T_2MS_SLEEP,#0
        SJMP    TIMER_S_RET
;*****是否结束的判断
TIMER_N_TF:
        JNB     B_WORK,TIMER_N_TF_RET
        JB      B_TIME_P,TIMER_N_TF_RET
        MOV     A,MINUTE
        CJNE    A,#0,NOT_OVER
        MOV     A,SECOND
        CJNE    A,#0,NOT_OVER
        SETB    B_TMOVER
        CLR     B_WORK
TIMER_N_TF_RET:
        RET
NOT_OVER:
        SJMP    TIMER_N_TF_RET

TMOVER:
        JNB     B_TMOVER,TMOVER_RET
        CLR     B_WORK
        SJMP    BUZZ
TMOVER_RET:
    RET
BUZZ:;蜂鸣器;\
        JNB B_TMOVER,BUZ_RET
        CLR     B_SLEEP
	CPL	P3.3
	lcall	DEL_BUZ

	;ajmp	BUZZ
DEL_BUZ:
	mov	r7,#00h
LLA:
    DJNZ	R7,LLA
    RET
BUZ_RET:
	AJMP    TMOVER_RET
;***********时间处理***********
;;显示时间，将显示数据放入显示缓冲区（LED_BUF作为起始地址，往后8位28H-2FH即缓冲区
;;INPUT:SECOND,MINUTE,HOUR
;;OUTPUT:LED_BUF
;;INFLUENCED:ACC，B，R0
DISP_TIMER:

        MOV     R0,#LED_BUF     ;把缓存地址读入R0
        MOV     A,HOUR
        MOV     B,#10
        DIV     AB
        MOV     @R0,A           ;存放十位（高位）
        INC     R0
        MOV     @R0,B           ;存放个位（低位）
        ;上面几行是用来分离HOUR中的十位和个位
        INC     R0              ;DISPLAY HOUR
        ;-1_DISP
        MOV     @R0,#0EH        ;0EH还是0F7,显示横杠 "-"
        INC     R0
        ;MIN_DISP
        MOV     A,MINUTE
        MOV     B,#10
        DIV     AB
        MOV     @R0,A
        INC     R0
        MOV     @R0,B
        INC     R0               ;DISPLAY MINUTE
        ;-2_DISP
        MOV     @R0,#0EH        ;-
        INC     R0
        ;SEC_DISP
        MOV     A,SECOND
        MOV     B,#10
        DIV     AB
        MOV     @R0,A
        INC     R0
        MOV     @R0,B            ;;DISPLAY SECOND
DISPTIMER_RET:
        RET
;****************DELAY*********
;;DELAY 2ms



;DELAY2MS:
        ;MOV     DELH,#2
;DELAY2:
        ;MOV     DELL,#250
;DELAY1:
       ; NOP
        ;NOP
        ;DJNZ    DELL,DELAY1
       ; DJNZ    DELH,DELAY2;250*2 *4 =2000US=2MS
        ;RET


;****************LED显示**************
;;LED动态扫描
;;INPUT:LED_BUF
;;OUTPUT:KEYCODE
;;INFLUENCED:Acc，DPTR
DISP_LED:
        JB      B_SLEEP,DISP_LED_NO
        MOV     A,#LED_BUF
        ADD     A,N_DISP;定位
        MOV     R0,A
        MOV     A,@R0;取出缓存中的时间数
        MOV     DPTR,#TAB_LED
        ANL     A,#0FH  ;0000 1111
        MOVC    A,@A+DPTR
        MOV     P0,A;查表，数码管中数的内容

        MOV     A,#0FFH;选择点亮的管
        CLR     C
        INC     N_DISP
        MOV     R0,N_DISP
LOOP_D1:
        RLC     A
        DJNZ    R0,LOOP_D1
        MOV     P2,A
        ANL     N_DISP,#07H;0000 0111,与完之后找位置
        SJMP    DISP_LED_RET
DISP_LED_NO:
        MOV     A,#0FFH
        MOV     P2,A
        MOV     P0,A
DISP_LED_RET:
        RET

;*************按键*************
;;按键处理

DO_KEY:
        MOV     A,KEYCODE
        CJNE    A,#0FFH,KEY_1;A!=FF --> KEY_1，这里键盘是低电平有效的，全为FF就是没有有效按键按下的状态
        CLR     B_KEYDO
        LJMP    KEY_RET
KEY_1:;开始
        CLR     B_SLEEP
        SETB    B_KEYDO
        CLR     B_TMOVER
        MOV     A,KEYCODE
        CJNE    A,#0FH,KEY_ZERO;A!=0F --> KEY_1
        SETB    B_KEYDO
        CLR     B_SLEEP
        CLR     B_TMOVER
        JNB     B_WORK,RESTART;B_WORK=0,重启计时；=1，暂停
        SJMP    PAUSE ;=1,暂停


PAUSE:  CLR     B_WORK
        LJMP    KEY_RET
RESTART:SETB    B_WORK
        LJMP    KEY_RET

KEY_ZERO:
        SETB    B_KEYDO
        MOV     A,KEYCODE
        CJNE    A,#0EH,KEY_10S
        CLR     B_TMOVER
        CLR     B_WORK;叫停
        SETB    B_TIME_P
        MOV     SECOND,#0
        MOV     MINUTE,#0
        LJMP    KEY_RET


KEY_10S:
        SETB    B_KEYDO
        CLR     B_TIME_P
        MOV     A,KEYCODE
        MOV     B,#60
        CJNE    A,#04H,KEY_1M
        MOV     A,SECOND
        ADD     A,#10
        DIV     AB;这里考虑进位的，如果秒数超过60，/60，取出商进位和余数放秒
        MOV     SECOND,B
        ADD     A,MINUTE
        MOV     MINUTE,A
        CLR     C
        SUBB    A,#100
        JC      KEY_RET
        MOV     MINUTE,#0

        LJMP    KEY_RET

KEY_1M:
        SETB    B_KEYDO
        MOV     A,KEYCODE
        CJNE    A,#03H,KEY_10M
        MOV     A,MINUTE
        ADD     A,#1
        MOV     MINUTE,A
        CLR     C
        SUBB    A,#100
        JC      KEY_RET
        MOV     MINUTE,#0
        LJMP    KEY_RET

KEY_10M:
        SETB    B_KEYDO
        MOV     A,KEYCODE
        CJNE    A,#02H,KEY_RET
        MOV     A,MINUTE
        ADD     A,#10
        MOV     MINUTE,A
        CLR     C
        SUBB    A,#100
        JC      KEY_RET
        MOV     MINUTE,#0

KEY_RET:
        MOV     KEYCODE,#0FFH
        RET



;*******************************
;;键盘扫描
;;INPUT:P1，P2
;;OUTPUT:KEYCODE
;;INFLUENCED:LASTCODE-上次采样的键码
;;           N_SAME -采样到相同键码的次数
;;DPTR
KEY_SCAN:
        MOV     P1,#0F0H
        NOP
        MOV     A,P1
        ANL     A,#0F0H;1111 0000 P1.4-P1,7，行扫描
        MOV     B,A;暂存保护
        MOV     P1,#0FH
        NOP
        MOV     A,P1
        ANL     A,#0FH;0000 1111 P1.0-1.3，列扫描
        ORL     A,B
        MOV     DPTR,#TAB_KEY
        MOVC    A,@A+DPTR
;这是防抖动的程序，只有连续三次扫描得到的是同一个按键按下时才把它写进键码
        CJNE    A,LASTCODE,GET_NEW_KEY
        ;采到的键码若等于上一个键码，把相同的次数与3比较
        MOV     A,N_SAME
        CJNE    A,#5,GET_NEXT;<5,继续获取
;=5,确定是新键码
        MOV     KEYCODE,LASTCODE
KEY_EXIT:
        INC     N_SAME
        RET
GET_NEXT:
        JC      KEY_EXIT
        RET
GET_NEW_KEY:
        MOV     LASTCODE,A
        MOV     N_SAME,#0
        RET

TAB_KEY:
 DB      0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH ;0-10H
        DB      0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH ;10-1FH
        DB      0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH ;20-2FH
        DB      0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH ;30-3FH
        DB      0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH ;40-4FH
        DB      0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH ;50-5FH
        DB      0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH ;60-6FH
        DB      0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0CH,0FFH,0FFH,0FFH,0DH,0FFH,0EH,0FH,0FFH ;70-7FH
        DB      0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH  ;80-8FH
        DB      0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH  ;90-9FH
        DB      0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH  ;A0-AFH
        DB      0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0AH,0FFH,0FFH,0FFH,09H,0FFH,00H,0BH,0FFH  ;B0-BFH
        DB      0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH  ;C0-CFH
        DB      0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,05H,0FFH,0FFH,0FFH,06H,0FFH,07H,08H,0FFH  ;D0-DFH
        DB      0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,01H,0FFH,0FFH,0FFH,02H,0FFH,03H,04H,0FFH  ;E0-EFH
        DB      0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH,0FFH  ;F0-FFH


        ;0EH 开始/暂停  0FH 归零  01H-10MIN  02H-1MIN  03H-10S


TAB_LED:
        DB      28H,7EH,0A2H,62H,74H
        DB      61H,21H,7AH,20H,60H
        DB      0FFH,0FFH,0FFH,0FFH,0F7H,0EFH







