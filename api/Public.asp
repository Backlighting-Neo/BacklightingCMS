<%


	'-----------------------------------------------------------------------------------
	'公共函数接口
	'函数接口帮助:http://bbs.actcms.com   QQ 5382862
	'-----------------------------------------------------------------------------------
 	Function demotest()'改函数会有demo.inc来调用.然后返回值给demo,这里如果只是独立函数调用的话.不需要赋值给actcool了
		
		demotest="我是公共函数"

	End Function 

	'=============汉字转换为UTF-8================== 
	function chinese2unicode(Str) 
	for i=1 to len(Str) 
	Str_one=Mid(Str,i,1) 
	Str_unicode=Str_unicode&chr(38) 
	Str_unicode=Str_unicode&chr(35) 
	Str_unicode=Str_unicode&chr(120) 
	Str_unicode=Str_unicode& Hex(ascw(Str_one)) 
	Str_unicode=Str_unicode&chr(59) 
	next 
	chinese2unicode = Str_unicode 
	end function
	'=============UTF-8转换为汉字================== 
	function UTF2GB(UTFStr) 
	for Dig=1 to len(UTFStr) 
	if mid(UTFStr,Dig,1)="%" then 
	if len(UTFStr) >= Dig+8 then 
	GBStr=GBStr & ConvChinese(mid(UTFStr,Dig,9)) 
	Dig=Dig+8 
	else 
	GBStr=GBStr & mid(UTFStr,Dig,1) 
	end if 
	else 
	GBStr=GBStr & mid(UTFStr,Dig,1) 
	end if 
	next 
	UTF2GB=GBStr 
	end function
	function ConvChinese(x) 
	A=split(mid(x,2),"%") 
	i=0 
	j=0 
	for i=0 to ubound(A) 
	A(i)=c16to2(A(i)) 
	next 
	for i=0 to ubound(A)-1 
	DigS=instr(A(i),"0") 
	Unicode="" 
	for j=1 to DigS-1 
	if j=1 then 
	A(i)=right(A(i),len(A(i))-DigS) 
	Unicode=Unicode & A(i) 
	else 
	i=i+1 
	A(i)=right(A(i),len(A(i))-2) 
	Unicode=Unicode & A(i) 
	end if 
	next 
	if len(c2to16(Unicode))=4 then 
	ConvChinese=ConvChinese & chrw(int("&H" & c2to16(Unicode))) 
	else 
	ConvChinese=ConvChinese & chr(int("&H" & c2to16(Unicode))) 
	end if 
	next 
	end function 
	function c2to16(x) 
	i=1 
	for i=1 to len(x) step 4 
	c2to16=c2to16 & hex(c2to10(mid(x,i,4))) 
	next 
	end function 
	function c2to10(x) 
	c2to10=0 
	if x="0" then exit function 
	i=0 
	for i= 0 to len(x) -1 
	if mid(x,len(x)-i,1)="1" then c2to10=c2to10+2^(i) 
	next 
	end function 
	function c16to2(x) 
	i=0 
	for i=1 to len(trim(x)) 
	tempstr= c10to2(cint(int("&h" & mid(x,i,1)))) 
	do while len(tempstr)<4 
	tempstr="0" & tempstr 
	loop 
	c16to2=c16to2 & tempstr 
	next 
	end function 
	function c10to2(x) 
	mysign=sgn(x) 
	x=abs(x) 
	DigS=1 
	do 
	if x<2^DigS then 
	exit do 
	else 
	DigS=DigS+1 
	end if 
	loop 
	tempnum=x 
	i=0 
	for i=DigS to 1 step-1 
	if tempnum>=2^(i-1) then 
	tempnum=tempnum-2^(i-1) 
	c10to2=c10to2 & "1" 
	else 
	c10to2=c10to2 & "0" 
	end if 
	next 
	if mysign=-1 then c10to2="-" & c10to2 
	end function 

	'个人代码风格注释（变量名中第一个小写字母表表示变量类型） 
	'i:为Integer型; 
	's:为String; 
	Function U2UTF8(Byval a_iNum) 
	Dim sResult,sUTF8 
	Dim iTemp,iHexNum,i 
	iHexNum = Trim(a_iNum) 
	If iHexNum = "" Then 
	Exit Function 
	End If 
	sResult = "" 
	If (iHexNum < 128) Then 
	sResult = sResult & iHexNum 
	ElseIf (iHexNum < 2048) Then 
	sResult = ChrB(&H80 + (iHexNum And &H3F)) 
	iHexNum = iHexNum \ &H40 
	sResult = ChrB(&HC0 + (iHexNum And &H1F)) & sResult 
	ElseIf (iHexNum < 65536) Then 
	sResult = ChrB(&H80 + (iHexNum And &H3F)) 
	iHexNum = iHexNum \ &H40 
	sResult = ChrB(&H80 + (iHexNum And &H3F)) & sResult 
	iHexNum = iHexNum \ &H40 
	sResult = ChrB(&HE0 + (iHexNum And &HF)) & sResult 
	End If 
	U2UTF8 = sResult 
	End Function 
	Function GB2UTF(Byval a_sStr) 
	Dim sGB,sResult,sTemp 
	Dim iLen,iUnicode,iTemp,i 
	sGB = Trim(a_sStr) 
	iLen = Len(sGB) 
	For i = 1 To iLen 
	sTemp = Mid(sGB,i,1) 
	iTemp = Asc(sTemp) 
	If (iTemp>127 OR iTemp<0) Then 
	iUnicode = AscW(sTemp) 
	If iUnicode<0 Then 
	iUnicode = iUnicode + 65536 
	End If 
	Else 
	iUnicode = iTemp 
	End If 
	sResult = sResult & U2UTF8(iUnicode) 
	Next 
	GB2UTF = sResult 
	End Function 

%>