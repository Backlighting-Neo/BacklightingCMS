'======================================
'��������StudentUnionSpecial
'���ã���ѧ����ģ������ר�ⲿ������
'���������DisplayNumber-Ҫ��ʾ��ר��������0��ʾ���ޣ��������ֱ�ʾ����
'����Ա�����
'���ڣ�2013��11��24�� 11:02:14
'======================================


Function StudentUnionSpecial(DisplayNumber)
	Dim rs1,rs2,resultstring,SQLString,ZhuanTiString,Counter
	Counter=0
	If DisplayNumber="0" Then
	  SQLString="SELECT Special_ACT.ID, Special_ACT.title, Special_ACT.PicIndex FROM Special_ACT ORDER BY Special_ACT.ID DESC;"
	Else
	  SQLString="SELECT Top " & Cstr(DisplayNumber) & " Special_ACT.ID, Special_ACT.title, Special_ACT.PicIndex FROM Special_ACT ORDER BY Special_ACT.ID DESC;"
	End If
	set rs1=actcms.actexe(SQLString)
	If Not rs1.eof Then
		Do While Not rs1.eof
			Counter=Counter+1
			resultstring=resultstring+"<div class=this_project>"+vbCrlf
			resultstring=resultstring+"<div class=project_picbg>"+vbCrlf
			resultstring=resultstring+"<div class=project_pic><a href=plus/Special.asp?ID=" & rs1("ID") & " target=_blank><img src=" & rs1("PicIndex") & "></a></div>"+vbCrlf
			resultstring=resultstring+"<div class=clr></div>"+vbCrlf
			resultstring=resultstring+"<div class=project_text>"+vbCrlf
			resultstring=resultstring+"<ul>"+vbCrlf
			resultstring=resultstring+"<li>&#x4E13;&#x9898;&#xFF1A;<a href=plus/Special.asp?ID=" & rs1("ID") & " target=_blank class=more2>" & rs1("title") & "</a></li>"+vbCrlf
			resultstring=resultstring+"</ul>"+vbCrlf
			resultstring=resultstring+"</div>"+vbCrlf
			resultstring=resultstring+"<div class=clr></div>"+vbCrlf
			resultstring=resultstring+"</div>"+vbCrlf
			resultstring=resultstring+"</div>"+vbCrlf
			rs1.movenext
			If Counter Mod 3=0 Then
				resultstring=resultstring+"<br><br>"
			End If
		Loop
	End If
	rs1.close
	set rs=nothing
	actcool=resultstring
End Function

