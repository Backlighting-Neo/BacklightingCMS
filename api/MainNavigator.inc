'======================================
'函数名：MainNavigator
'作用：替换原有的导航栏菜单，以实现鼠标划过时显示二级菜单
'传入参数：无
'程序员：逆光
'日期：2013年11月21日 19:44:05
'======================================
Function MainNavigator()
	Dim rs1,rs2,resultstring
	set rs1=actcms.actexe("SELECT Class_Act.ClassName, Class_Act.ID, Class_Act.ClassID FROM Class_Act WHERE (((Class_Act.ParentID)='0')) AND dh=1 ORDER BY Class_Act.OrderID, Class_Act.ID;")
	If Not rs1.eof Then
		Do While Not rs1.eof
		    resultstring=resultstring+"<li ><a href='/List.asp?L-" & rs1("ClassID") & ".html'><strong>" & rs1("ClassName") & "</strong></a>"+vbCrLf
			set rs2=actcms.actexe("SELECT Class_Act.ClassName, Class_Act.ID, Class_Act.ClassID FROM Class_Act WHERE (((Class_Act.ParentID)='" & rs1("ClassID") & "')) AND dh=1 ORDER BY Class_Act.OrderID, Class_Act.ID;")
			If Not rs2.eof Then
			    ZiChang=0
			    Do While Not rs2.eof
					if len(rs2("ClassName"))>ZiChang then ZiChang=len(rs2("ClassName"))
					rs2.movenext
				Loop
				rs2.movefirst
			    resultstring=resultstring+"<ul style='width:" & cstr(15*ZiChang+20) & "'>"+vbCrLf
				Do While Not rs2.eof
					resultstring=resultstring+"<li style='height:30px; width:" & cstr(15*ZiChang+20) & "px'><a href='/List.asp?L-" & rs2("ClassID") & ".html'>" & rs2("ClassName") & "</a></li>"+vbCrLf
					rs2.movenext
				Loop
				resultstring=resultstring+"</ul>"+vbCrLf
			End If
			
		    rs2.close
			set rs2=nothing
			rs1.movenext
		Loop
	resultstring=resultstring+"</li>"+vbCrLf
	End If
	rs1.close
	set rs=nothing
	actcool=resultstring
End Function