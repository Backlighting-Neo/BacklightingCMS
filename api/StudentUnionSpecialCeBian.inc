'======================================
'函数名：StudentUnionSpecialCeBian
'作用：在学生会模板中提专题部分内容在侧边栏显示
'传入参数：DisplayNumber-要显示的专题条数，0表示无限，其他数字表示条数
'程序员：逆光
'日期：2013年11月24日 11:02:14
'======================================


Function StudentUnionSpecialCeBian()
	Dim rs1,rs2,resultstring,SQLString,ZhuanTiString
	  SQLString="SELECT Top 3 Special_ACT.ID, Special_ACT.PicIndex FROM Special_ACT ORDER BY Special_ACT.ID DESC;"
	set rs1=actcms.actexe(SQLString)
	If Not rs1.eof Then
		Do While Not rs1.eof
			resultstring=resultstring+"<div class='project_thum'>"+vbCrlf
			resultstring=resultstring+"<ul>"+vbCrlf
			resultstring=resultstring+"<li><a href=plus/Special.asp?ID=" & rs1("ID") & " target=_blank ><img src=" & rs1("PicIndex") & " width='280px' height='80px' style='border-radius:3px;'/></a></li>"+vbCrlf
			resultstring=resultstring+"</ul>"+vbCrlf
			resultstring=resultstring+"</div>"+vbCrlf
			rs1.movenext
		Loop
	End If
	rs1.close
	set rs=nothing
	actcool=resultstring
End Function


