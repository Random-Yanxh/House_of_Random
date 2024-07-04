ORG     0000h
;30H与31H中存放4位BCD码数字，编写程序将数字倒序排列
;##两步：1. 各数倒序
;2. 两数交换
LJMP    BEGIN   
ORG     0030h

BEGIN:
MOV R0,#30H
MOV R1,#31H
MOV A,@R0
SWAP A  ;交换A高低位，12变21
MOV @R0,A
MOV A,@R1
SWAP A  ;交换A高低位，12变21
XCH A,@R0;注意XCH对A操作，因此这里直接将本来放回@R1再与@R0交换的数
;放到@R0。
MOV @R1,A ;将原来@R0中倒序后的数放入@R1
HERE:SJMP HERE

