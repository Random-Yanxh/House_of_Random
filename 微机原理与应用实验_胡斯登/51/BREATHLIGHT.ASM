ORG 0000H
SJMP START
ORG 0080H
START:
MOV R5,#0FEH ; 1111 1110初始化亮度级别
LOOP: ;变暗的循环
     MOV A,#00H ;初始化，点灯
     MOV P1,A
     CALL DELAY     ;间隔
     MOV A, #0FFH   ;灭灯，初步实现闪烁
     MOV P1,A
     MOV A,R5       ;亮度级别，把FE放在A中
     CALL DELAY2     ;变暗速率调节，Delay2初始是把R5的值扔进A
     DEC R5           ;亮度下降
     CJNE R5,#20H,LOOP   ;R5！=20H则跳转回LOOP，这个20H就是亮度最低值（似乎不是这样）

MOV R5,#20H  ;似乎没用

LOOP2:  ;回亮的循环
      MOV A,#00H  ;亮
      MOV P1,A       ;
      CALL DELAY
      MOV A, #0FFH
      MOV P1,A
      MOV A,R5
      CALL DELAY2  ;此处为止都和上面一样
      MOV A,#00H
      MOV P1,A
      CALL DELAY   
      MOV A, #0FFH
      MOV P1,A
      MOV A,R5
      CALL DELAY2
      INC R5
      CJNE R5,#0FEH,LOOP2 ;回最高亮度，与初值对应。应该说FE效果最好，和FF最接近，最平滑

JMP START  ;往复，一直循环呼吸

DELAY:MOV R3,#20H ;跑了一个20H*20H=400H=1024D，经过实验尝试，这一段就是用来调整体亮度，也就是实现PWM调光的部分
D1:MOV R4,#40  ;时间越短，整体越暗，显示出来的占空比越大；反之，观感越亮，显示出的占空比越小。
DJNZ R4,$
DJNZ R3,D1
RET

DELAY2: MOV R3,A
D2: MOV R4,#20H ;似乎是频闪速率，20H比较合适
DJNZ R4,$
DJNZ R3,D2
RET

END
