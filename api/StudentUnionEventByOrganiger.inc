'======================================
'函数名：StudentUnionEventByOrganiger
'作用：查找出指定组织者的组织的活动
'传入参数：Organiger-组织者，以字符串形式传入，例如"自动化工程学院学生会文艺部"
'         DisplayNumber-要显示的专题条数，0表示无限，其他数字表示条数
'程序员：逆光
'日期：2013年11月24日 18:13:10
'======================================


Function StudentUnionEventByOrganiger(Organiger,DisplayNumber)
	Dim rs1,resultstring,SQLString
	If DisplayNumber="0" Then
	  SQLString="SELECT ID,Title FROM Article_ACT WHERE HDZuZhiZhe_ACT='" & Organiger & "' ORDER BY ID DESC;"
    Else
      SQLString="SELECT Top " & Cstr(DisplayNumber) & " ID,Title FROM Article_ACT WHERE HDZuZhiZhe_ACT='" & Organiger & "' ORDER BY ID DESC;"
	End If
	set rs1=actcms.actexe(SQLString)
	If Not rs1.eof Then
		resultstring=resultstring+"<div class='eventlist'><ul>"+vbCrLf
		Do While Not rs1.eof
			resultstring=resultstring+"<li><a class='style2' href='List.asp?C-1-" & Rs1("ID") & ".html'>" & Rs1("Title") & "</a></li>"+vbCrlf
			rs1.movenext
		Loop
		resultstring=resultstring+"</ul></div>"+vbCrLf
	End If
	rs1.close
	set rs=nothing
	actcool=resultstring
End Function