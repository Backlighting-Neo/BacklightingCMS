'======================================
'��������StudentUnionSpecialCeBian
'���ã���ѧ����ģ������ר�ⲿ�������ڲ������ʾ
'���������DisplayNumber-Ҫ��ʾ��ר��������0��ʾ���ޣ��������ֱ�ʾ����
'����Ա�����
'���ڣ�2013��11��24�� 11:02:14
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

