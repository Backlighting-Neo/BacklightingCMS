'======================================
'��������StudentUnionEventByOrganiger
'���ã����ҳ�ָ����֯�ߵ���֯�Ļ
'���������Organiger-��֯�ߣ����ַ�����ʽ���룬����"�Զ�������ѧԺѧ�������ղ�"
'         DisplayNumber-Ҫ��ʾ��ר��������0��ʾ���ޣ��������ֱ�ʾ����
'����Ա�����
'���ڣ�2013��11��24�� 18:13:10
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