'======================================
'函数名：DisplayPageNav
'作用：用于在二级页面和内容页面显示当前栏目下的二级子栏目，可以显示深度最大为2的子栏目
'传入参数：ClassID，可以用{$ClassID}标签
'程序员：逆光
'日期：2013年10月28日 21:14:17

'苦逼的程序员大晚上写程序！！！谁能懂！！！
'没加ClassID的容错处理，所以调用时你敢胡传参数我就敢不执行，你信不信！！
'======================================
Function DisplayPageNav(ClassID)
	Dim rs1,rs2,resultstring
	set rs1=actcms.actexe("SELECT Class_Act.ClassName, Class_Act.ID, Class_Act.ClassID FROM Class_Act WHERE (((Class_Act.ParentID)='" & ClassID & "')) ORDER BY Class_Act.OrderID, Class_Act.ID;")
	If Not rs1.eof Then
		resultstring="<table border='0' cellpadding='0' cellspacing='0' wIDth='100%' align='center'>"
		Do While Not rs1.eof
			resultstring=resultstring+"<tr><td height='33'><img src='/images/NavigaterImage.jpg' border='0'/>&nbsp;<a  href='/List.asp?L-" & rs1("ClassID") & ".html' target='_self'>" & rs1("ClassName") & "</a></td></tr>"
			set rs2=actcms.actexe("SELECT Class_Act.ClassName, Class_Act.ID, Class_Act.ClassID FROM Class_Act WHERE (((Class_Act.ParentID)='" & rs1("ClassID") & "')) ORDER BY Class_Act.OrderID, Class_Act.ID;")
			If Not rs2.eof Then
				Do While Not rs2.eof
					resultstring=resultstring+"<tr><td height='33'><table><tr><td>&nbsp;&nbsp;&nbsp;&nbsp;<img src='/images/NavigaterImage.jpg' border='0'/>&nbsp;</td><td><a  href='/List.asp?L-" & rs2("ClassID") & ".html' target='_self'>" & rs2("ClassName") & "</a></td></tr></table></td></tr>"					  
					rs2.movenext
				Loop
			End If
			
		    rs2.close
			set rs2=nothing
			rs1.movenext
		Loop
  End If
  rs1.close
  set rs=nothing
  
  resultstring=resultstring+"</table>"
  
  actcool=resultstring
End Function
