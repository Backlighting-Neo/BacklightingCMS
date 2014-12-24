  <% 
Dim ACT_L
Set ACT_L = New ACT_Code
Class Cls_Special
	Dim Reg
    Dim   ID,rs,sql,i,ModeID,DateForm,TitleLen,ContentLen,ListNumber,keywords,ClassID,Parameter,sqlstr,DiyContents
	Public TemplateContent
 	Private Sub Class_Initialize()
		Set Reg = New RegExp
		Reg.Ignorecase = True
		Reg.Global = True
	End Sub
	Function loads(ID)	 
  	 if ID=0 or ID="" Then ID=1
  	 Application(AcTCMSN&"ModeID")=1
	 Set rs=actcms.actexe("select id,title,writer,pubdate,tempurl from Special_ACT where id="&id)
	 If rs.eof Then response.write "参数错误":response.end
 	 TemplateContent = ACT_L.LoadTemplate(rs("tempurl"))
 	 If InStr(TemplateContent, "{$Special_writer}") > 0 Then
		 TemplateContent = Replace(TemplateContent,"{$Special_writer}",Trim(rs("writer")))
	 Else
		 TemplateContent = Replace(TemplateContent,"{$Special_writer}","")
	 End If 
 	 If InStr(TemplateContent, "{$Special_pubdate}") > 0 Then
 		 TemplateContent = Replace(TemplateContent,"{$Special_pubdate}",rs("pubdate"))
	 Else
		 TemplateContent = Replace(TemplateContent,"{$Special_pubdate}","")
	 End If 

	 If InStr(TemplateContent, "{$ID}") > 0 Then
 		 TemplateContent = Replace(TemplateContent,"{$ID}",rs("ID"))
 	 End If 
  
 	 If InStr(TemplateContent, "{$Special_Content}") > 0 Then
 		 TemplateContent = Replace(TemplateContent,"{$Special_Content}",rs("Content"))
	 Else
		 TemplateContent = Replace(TemplateContent,"{$Special_Content}","")
	 End If 
	 
	 
	 Set rs=actcms.actexe("select id,notename,arcid,ModeID,DiyContent,DateForm,TitleLen,ContentLen,ListNumber,isauto,keywords,ClassID from Node_ACT where sid=0 or  sid="&id)
	 If  Not rs.eof Then 
 	  sql=rs.getrows(-1)
	  for i=0 to ubound(sql,2)
 		 If InStr(TemplateContent, "{$node_" & sql(1,i) & "}") > 0 Then
			ModeID=sql(3,i)
			If ModeID="0" Then ModeID=Cint(Application(AcTCMSN & "ModeID"))
 			DateForm=sql(5,i)
			TitleLen=sql(6,i)
			ContentLen=sql(7,i)
			ListNumber=sql(8,i)
			keywords=sql(10,i)
			ClassID=sql(11,i)
			If ClassID<>"0" Then 
 				If InStr(ClassID, ",") > 0 Then
					 Parameter="ClassID In (" & ClassID & ") And"
				Else
					Parameter="ClassID='" & Replace(ClassID,"'","") & "' And"
				End If 
			 End If 
 			If ListNumber=0 Then ListNumber=10
			If Trim(sql(9,i))=1 Then 
 				If Trim(sql(2,i))<>"" Then Parameter=" id in("&Trim(sql(2,i))&")  and "
				Sqlstr="Select TOP  " & ListNumber & " ID,ClassID,Title,UpdateTime,ActLink,FileName,infopurview,readpoint,PicUrl,Intro,Content,CopyFrom,Author,KeyWords"&ACTCMS.DIYField(ModeID)&" From  "&ACTCMS.ACT_C(ModeID,2)&"  Where  "&Parameter&" isAccept=0 AND delif=0   ORDER BY IsTop Desc,ID Desc"  
 				DiyContents=ACT_L.ACTCMS_A_Code(SqlStr,TitleLen,"",DateForm,Trim(sql(4,i)),ModeID,ContentLen) 
			Else 
 				Sqlstr="Select TOP  " & ListNumber & " ID,ClassID,Title,UpdateTime,ActLink,FileName,infopurview,readpoint,PicUrl,Intro,Content,CopyFrom,Author,KeyWords"&ACTCMS.DIYField(ModeID)&" From  "&ACTCMS.ACT_C(ModeID,2)&"  Where title like '%" & keywords & "%' and "&Parameter&" isAccept=0 AND delif=0   ORDER BY IsTop Desc,ID Desc"  
 			 DiyContents=ACT_L.ACTCMS_A_Code(SqlStr,TitleLen,"",DateForm,Trim(sql(4,i)),ModeID,ContentLen) 
 			End If 
 			 Parameter=""
			 TemplateContent = Replace(TemplateContent,"{$node_" & sql(1,i) & "}",DiyContents)
 		 Else
		  TemplateContent = Replace(TemplateContent,"{$node_" & sql(1,i) & "}","")
		 End If
 		next	 
	  End If 

       Set rs=actcms.actexe("select id,title,picurl from SpecialPicUrl_ACT where sid="&id&"")
	   If Not  rs.eof Then 
	   sql=rs.getrows(-1)
		  for i=0 to ubound(sql,2)
				If InStr(TemplateContent, "{$SpecialPic_"&sql(0,i)&"_"&sql(1,i)&"}") > 0 Then
					 TemplateContent = Replace(TemplateContent,"{$SpecialPic_"&sql(0,i)&"_"&sql(1,i)&"}",actcms.PathDoMain&Trim(sql(2,i)))
				Else
					 TemplateContent = Replace(TemplateContent,"{$SpecialPic_"&sql(0,i)&"_"&sql(1,i)&"}","")
				End If 
	 
		  Next 
		End If 
		TemplateContent = ACT_L.LabelReplaceAll(TemplateContent)
		TemplateContent=ACT_L.actcmsexe(TemplateContent)
		labelreg
		loads=TemplateContent 
    End Function 
	Function labelreg()
	Dim regEx,Matches,Match,CTemp
	Dim TagLabs, Tagsstr, Loopstr
	Set regEx = New RegExp
	regEx.Pattern = "<!--(.+?)\:(.+?)-->([\s\S]*?)<!--\1-->"
	regEx.IgnoreCase = True
	regEx.Global = True
	Set Matches = regEx.Execute(TemplateContent)
 	Dim Lab_Table,Lab_top,Lab_order,Lab_Field,Lab_where,BackValue,actsql
	For Each Match In Matches
			TagLabs = Match.SubMatches(0)	  ' 标签
			Tagsstr = Match.SubMatches(1)	  ' 属性
			Loopstr = Match.SubMatches(2)	  ' innerText

			Lab_Table = GetAttr(Tagsstr, "table")  ' 组合查询,表
			Lab_top = GetAttr(Tagsstr, "top")  '数量
			Lab_order = GetAttr(Tagsstr, "order")  '排序
			Lab_Field = GetAttr(Tagsstr, "Field")  '字段
			Lab_where = GetAttr(Tagsstr, "where")  '条件
			If len(Lab_top)=0 Or Not IsNumeric(Lab_top) Then Lab_top = 10
			If Len(Lab_Field) = 0 Then Lab_Field = "*"
			If Len(Lab_Where) = 0 Then Lab_Where = " 1=1 "
			If Len(Lab_Order) = 0 Then Lab_Order = " [ID] Desc"
			BackValue = ""
			Err.Clear
			actsql="select Top "&Lab_top&" "&Lab_Field&"   from  "&Lab_Table&" Where "&Lab_Where&" Order By "&Lab_Order&" "
  				Set rs=conn.execute(actsql) 
 			    If Err   Then Response.Write "<font color=red>标签执行错误[" & actsql & " =>   " & Err.Description & "]</font>": Response.End

 				For i = 1 To Lab_top
					If Rs.Eof Then Exit For ' 不存在记录就退出
					If Len(TagLabs) = 0 Then TagLabs = "field"
						BackValue = BackValue & Parser_Tags("\[(.+?)/]", Loopstr, Rs) ' 替换
 					Rs.MoveNext
				Next
				Rs.Close
 			    TemplateContent = Replace(TemplateContent, Match.Value,BackValue)
 		Next

	   If RegExists("<!--(.+?):\{(.+?)\}-->([\s\S]*?)<!--\1-->", TemplateContent) Then Call labelreg() 
    End Function 
 
	Public Function Parser_Tags(ByVal Pattern, ByVal Temp, ByVal Dat)
   		On Error Resume Next
  		Dim Matches, Match
		Dim Tagsstr, Tagsval, Tagsvalt, TagTitle: TagTitle = False
		Dim Tag_Len, Tag_Lenext, Tag_Format, Tag_Replace, Tag_Function,Tag_width,Tag_Height
		Dim Re, Re1, Re2
		Dim i, c, l, t,reg,tagset,isset,temptitle,date_Len
		Set Reg = New RegExp
		Reg.Pattern = Pattern
		Set Matches = Reg.Execute(Temp)
		For Each Match In Matches
 			Tagsstr = Split(Match.SubMatches(0), " ")
        	tagset = Tagsstr(0)
			If   UBound(Tagsstr)=0 Then 
				Tagsval=""
			Else
 		 		Tagsval=Tagsstr(1)
			End If 
 			isset=""
   		 	isset = GetAttr(Tagsval, "isset")
 			If Len(isset)=0 Then 
  				temptitle=dat(tagset)
			Else 
 				'参数省略
			End If 
  		 	Tag_Len = GetAttr(Temp, "len")
 		 	date_Len = GetAttr(Temp, "date")
			If len(date_Len)<>0 Then temptitle=formmdate(temptitle,LCase(date_Len))
			If len(Tag_Len)<>0 Then temptitle=Left(temptitle,Tag_Len)
			
 

 			If Trim(temptitle)<>"" Then 
				Temp = Replace(Temp, Match.Value,temptitle)
			Else 
				Temp = Replace(Temp, Match.Value,"")
			End If 
  		Next
		Parser_Tags = Temp
		temptitle=""
 	End Function

	Function formmdate(FieldValue,DateMB)
		  IF Not IsDate(FieldValue) Then formmdate=FieldValue:Exit Function
		  formmdate=Replace(DateMB,"yyyy",Year(FieldValue))
		  formmdate=Replace(formmdate,"yy",Right("0" & Year(FieldValue), 2))
		  formmdate=Replace(formmdate,"mm",Right("0" & Month(FieldValue), 2))
		  formmdate=Replace(formmdate,"dd",Right("0" & Day(FieldValue), 2))
		  formmdate=Replace(formmdate,"hh",Right("0" & hour(FieldValue), 2))
		  formmdate=Replace(formmdate,"mm",Right("0" & minute(FieldValue), 2))
		  formmdate=Replace(formmdate,"ss",Right("0" & second(FieldValue), 2))
		  thisdate=Year(FieldValue) & "-" & Right("0" & Month(FieldValue), 2) & "-" & Right("0" & Day(FieldValue), 2)	
		  thisdaydate=Year(date) & "-" & Right("0" & Month(date), 2) & "-" & Right("0" & Day(date), 2)
 				If thisdate=thisdaydate Then 
					Dim m
					m=Datediff("n",FieldValue,now())
 					if m<60  Then
						If FieldValue<=now() And abs(Datediff("s",FieldValue,now()))<60 Then 
							formmdate="<font color=red>"&abs(Datediff("s",FieldValue,now()))&"秒前</font>"
						Else 
							formmdate="<font color=red>"&m&"分种前</font>"
						End If 
					Else
						If CLng(abs(Datediff("h",FieldValue,now())))>5 Then 
							formmdate="<font color=red>今天</font>"
						Else 
							formmdate="<font color=red>"&abs(Datediff("h",FieldValue,now()))&"小时前</font>"
						End If 
					End  if
 				Else 
					formmdate=year(thisdate)&"-"&month(thisdate)&"-"&day(thisdate)
				End If 
   	End Function 
  	
    Function GetAttr(ByVal Tagsstr, ByVal AttrName)
 		Dim Matches, Match,regEx
		Set regEx = New RegExp
		regEx.Pattern =  AttrName & "=""(.+?)"""
		Set Matches = regEx.Execute(Tagsstr)
 		For Each Match In Matches
 			 GetAttr = Match.SubMatches(0)
 		Next
 	End Function
	' 是否存在此类标签
	Public Function RegExists(ByVal Pattern, ByVal TestContent)
		Reg.Pattern = Pattern
		RegExists = Reg.Test(TestContent)
	End Function
  End Class
%>  