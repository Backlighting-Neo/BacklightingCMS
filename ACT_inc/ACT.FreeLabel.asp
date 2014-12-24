<%
Class ACTFreeLabel
		Private DataSourceType,DataSourceStr,tconn
		Private Sub Class_Initialize()
		End Sub
        Private Sub Class_Terminate()
		   If isobject(tconn) Then
		   TConn.Close:Set TConn=Nothing
		   End If
		End Sub
		'替换自定义函数标签
		Function ReplaceReeLabel(Content)
			Dim regEx, Matches, SqlLabel,Match
			Dim Matchn,n
			Set regEx = New RegExp
			regEx.Pattern = "{ACTSQL_[^{]*\)}"
			regEx.IgnoreCase = True
			regEx.Global = True
			Set Matches = regEx.Execute(Content)
			ReplaceReeLabel=Content
			For Each Match In Matches
			  SqlLabel=Match.value
			  ReplaceReeLabel=Replace(ReplaceReeLabel,SqlLabel,ReplaceDIYFunctionLabel(SqlLabel,"label"))
			Next		
		End Function
		'返回循环次数
		Function GetLoopNum(Content)
			 Dim regEx, Matches, Match
			 Set regEx = New RegExp
			 regEx.Pattern="\[loop=\d*]"
			 regEx.IgnoreCase = True
			 regEx.Global = True
			 Set Matches = regEx.Execute(Content)
			 If Matches.count > 0 Then
			  GetLoopNum=Replace(Replace(Matches.item(0),"[loop=",""),"]","")
			 Else
			  GetLoopNum=0
			 end if
		End Function
		Function ReplaceDIYFunctionLabel(SqlLabel,GetFrom)
		  Dim I,ARs,LabelName,UserParamArr,FunctionLabelParamArr,CirLabelContent,FunctionSQL,LabelContent
		  Dim FunctionLabelType,ItemName,PageStyle,PerPageNumber,TotalPut,PageNum,J,TempStr,Ajax
		  LabelName    = Replace(Replace(Split(SqlLabel,"(")(0),"""",""),"'","")
		  '用户函数参数
		  UserParamArr = Split(Replace(Replace(Replace(Replace(SqlLabel,LabelName&"(",""),")}",""),"""",""),"'",""),",")   
		  Set ARs=actcms.actexe("Select  top 1 LabelContent,Description From Label_Act Where LabelName='" & LabelName & "}'")
		  IF ARs.Eof And ARs.Bof Then
		     ARs.Close:Set ARs=Nothing:ReplaceDIYFunctionLabel="":Exit Function
		  Else
		    FunctionLabelParamArr = Split(ARs(0),"§")
			LabelContent          = Replace(ARs(1),vbcrlf,"$ACT:Page$")
		  End If
		   ARs.Close
		  DataSourceType=FunctionLabelParamArr(3)
		  DataSourceStr=FunctionLabelParamArr(4)
		  FunctionSQL=FunctionLabelParamArr(5)
 		  if DataSourceType=1 Or DataSourceType=5 Or DataSourceType=6 then	DataSourceStr=ACTCMS.GetAbsolutePath(DataSourceStr)
 		   FunctionSQL=Replace(FunctionSQL,"{$CurrClassID}",""&ACTCMS.TempClassID(Application(AcTCMSN & "ClassID"))&"")'当前栏目ID
		   FunctionSQL=Replace(FunctionSQL,"{$ThisClassID}","'"&Application(AcTCMSN & "ClassID")&"'")'当前栏目ID
		   FunctionSQL=Replace(FunctionSQL,"{$CurrInfoID}",Application(AcTCMSN & "ID"))'当前ID
 		   For I=0 To Ubound(UserParamArr)
		    FunctionSQL  = Replace(FunctionSQL,"{$Param("&I&")}",UserParamArr(I))
			LabelContent = Replace(LabelContent,"{$Param("&I&")}",UserParamArr(I))
		   Next
		   FunctionLabelType=FunctionLabelParamArr(2)
		   If Not Isnumeric(FunctionLabelType) Then FunctionLabelType=0
			FunctionLabelType=FunctionLabelParamArr(0)
			PageStyle=FunctionLabelParamArr(2)
			ItemName=FunctionLabelParamArr(1)
  		   If OpenExtConn=false Then ReplaceDIYFunctionLabel="外部数据库连接出错!":Exit Function
           If DataSourceType=0 Then
			  ARs.Open FunctionSQL,Conn,1,1
		   Else
			  ARs.Open FunctionSQL,TConn,1,1
		   End If 
  		   If Not ARs.Eof Then
			    Dim regEx, Matches, Match,LoopTimes
				Set regEx = New RegExp
				regEx.Pattern = "\[loop=\d*].+?\[/loop]"
				regEx.IgnoreCase = True
				regEx.Global = True
				Set Matches = regEx.Execute(LabelContent)
				If FunctionLabelType=1  Then  '分页标签
				         PerPageNumber=0
				         For Each Match In Matches
							PerPageNumber=PerPageNumber+GetLoopNum(Match.Value) '每页记录数
						 Next
                         If PerPageNumber=0 Then ReplaceDIYFunctionLabel="自由标签的循环次数必须大于0":Exit Function
						 
				  		TotalPut = ARs.recordcount
						if (TotalPut mod PerPageNumber)=0 then
								PageNum = TotalPut \ PerPageNumber
						else
								PageNum = TotalPut \ PerPageNumber + 1
						end if
						Application(AcTCMSN & "PageStyle") = PageStyle
						
						Dim Url,urlarr,ClassID
						On Error Resume Next
						Url=Request.ServerVariables("QUERY_STRING")
						urlarr=Split(Split(url,".")(0),"-")
					    ClassID = RSQL(urlarr(1))

						If  ClassID<>"" Then
							 If UBound(urlarr)=2 Then CurrPage=ChkNumeric(urlarr(2))
						     Dim CurrPage:CurrPage=ChkNumeric(CurrPage)
							 If CurrPage<=0 Then CurrPage=1
						     Application(AcTCMSN & "PageNum")=PageNum
						     Application(AcTCMSN & "TotalPut")=TotalPut
						     Application(AcTCMSN & "CurrPage")=CurrPage
 							 TempCirContent    = LabelContent
							 ARs.Move (CurrPage - 1) * PerPageNumber
						     For Each Match In Matches
								  LoopTimes=GetLoopNum(Match.Value)   '循环次数
								  CirLabelContent = Replace(Replace(Match.value,"[loop=" & LoopTimes&"]",""),"[/loop]","")
								   TempCirContent    = Replace(TempCirContent,"[loop="&LoopTimes&"]"&CirLabelContent&"[/loop]",GetCirLabelContent(CirLabelContent,ARs,LoopTimes),1,1)
								  If ARs.Eof Then Exit For
							 Next
 							  TempStr = TempCirContent & ACTCMS.GetPageList(PageStyle,ItemName,PageNum,CurrPage,TotalPut,PerPageNumber)
 							  TempStr=TempStr &"{$PageList}" '加上分页符
							  ReplaceDIYFunctionLabel=Replace(CleanLabel(TempStr),"$ACT:Page$",vbcrlf)
						Else
						    dim TempCirContent
							For I = 1 To Cint(PageNum)
							     TempCirContent    = LabelContent
								 For Each Match In Matches
								  LoopTimes=GetLoopNum(Match.Value)   '循环次数
								  CirLabelContent = Replace(Replace(Match.value,"[loop=" & LoopTimes&"]",""),"[/loop]","")
								   TempCirContent=Replace(TempCirContent,"[loop="&LoopTimes&"]"&CirLabelContent&"[/loop]",GetCirLabelContent(CirLabelContent,ARs,LoopTimes),1,1)
								  If ARs.Eof Then Exit For
								 Next
						      Application(AcTCMSN & "pagecount")=TotalPut 
							  TempStr = TempStr & TempCirContent & actcms.GetPageList(PageStyle,ItemName,PageNum,I,TotalPut,PerPageNumber)
							  TempStr=TempStr & "{$PageList}" '加上分页符
							Next
							Application(Cstr(AcTCMSN & "PageList")) = Replace(CleanLabel(TempStr),"$ACT:Page$",vbcrlf)
							ReplaceDIYFunctionLabel="{PageListStr}"
					 End If

				Else
					Do While Not ARs.Eof
						For Each Match In Matches
						  LoopTimes=GetLoopNum(Match.Value)   '循环次数
						  CirLabelContent = Replace(Replace(Match.value,"[loop=" & LoopTimes&"]",""),"[/loop]","")
						  LabelContent    = Replace(LabelContent,"[loop="&LoopTimes&"]"&CirLabelContent&"[/loop]",GetCirLabelContent(CirLabelContent,ARs,LoopTimes),1,1)
						  If ARs.Eof Then Exit For
						Next
						If ARs.Eof Then Exit Do
					Loop
					'消除多余的循环体
					ReplaceDIYFunctionLabel=Replace(CleanLabel(LabelContent),"$ACT:Page$",vbcrlf)
				End If		 
		   Else
		     ReplaceDIYFunctionLabel="":Exit Function
		   End if
		   ARs.Close:Set ARs=Nothing
		   
		End Function
		'消除多余的循环体
		Function CleanLabel(Content)
				Dim regEx, Matches, Match,LoopTimes
				Set regEx = New RegExp
					regEx.Pattern = "\[loop=\d*][^\[\]]*\[/loop]"
					regEx.IgnoreCase = True
					regEx.Global = True
					Set Matches = regEx.Execute(Content)
					For Each Match In Matches
					  Content=Replace(Content,Match.value,"")
					Next
					CleanLabel=Content
		End Function
		'替换循环部分内容
		Function GetCirLabelContent(CirLabelContent,ByRef ARs,LoopTimes)
		Dim regEx, Matches, Match, TempStr
		Dim FieldParam,FieldParamArr,FieldName,FieldType,ReturnFieldValue
		Dim DB_FieldValue,FieldParamLength,I,FieldPosition,N
			If Not IsNumeric(LoopTimes) Then LoopTimes=10
			For N=1 To LoopTimes
			  If Not ARs.Eof Then
					Set regEx = New RegExp
					regEx.Pattern = "{\$Field\([^{\$}]*}"
					regEx.IgnoreCase = True
					regEx.Global = True
					Set Matches = regEx.Execute(CirLabelContent)
					TempStr=Replace(CirLabelContent,"{$AutoID}",N)
					For Each Match In Matches
					  FieldParam    = Replace(Replace(Match.Value,"{$Field(",""),")}","")
					  FieldParamArr = Split(FieldParam,",")
					  FieldParamLength=Ubound(FieldParamArr) '参数数组长度
 					  FieldName     = FieldParamArr(0)       '根据参数得到字段名称
					  FieldType     = FieldParamArr(1)       '根据参数得到字段类型
					  FieldPosition=0
					  For I=0 To ARs.Fields.count-1
					    IF lcase(FieldName)=lcase(ARs.Fields(I).name) Then FieldPosition=I:Exit For
					  Next
						  DB_FieldValue=ARs(FieldPosition)      '得到字段的值
					 
						  Select Case Lcase(FieldType)
						 
						   Case "text"
							 ReturnFieldValue=Get_Text_Field(DB_FieldValue,FieldParamArr(2),FieldParamArr(3),FieldParamArr(4),FieldParamArr(5))
						   Case "num"
							 ReturnFieldValue=Get_Num_Field(DB_FieldValue,FieldParamArr(2),FieldParamArr(3))
						   Case "date"
							 ReturnFieldValue=Get_Date_Field(DB_FieldValue,FieldParamArr(2))
						   Case "getinfourl"
							 ReturnFieldValue=Get_InfoUrl_Field(FieldName,DB_FieldValue,FieldParamArr(2),FieldParamArr(3))
						   Case "getclassurl"
							 ReturnFieldValue=Get_ClassUrl_Field(FieldName,DB_FieldValue,FieldParamArr(2),FieldParamArr(3))
						  End Select
					 
					  on error resume next
				      TempStr=Replace(TempStr,"{$Field(" &FieldParam &")}",ReturnFieldValue)
					Next

					 GetCirLabelContent=GetCirLabelContent &TempStr
				Else
				  Exit For
				End If
				 ARs.MoveNext
			Next
		
		End Function
		
		'取文本字段的值
		Function Get_Text_Field(FieldValue,CutNum,EndTag,HtmlTag,DefaultChar)
		 Dim TempStr:TempStr=FieldValue
		 If FieldValue="" Or IsNull(FieldValue) Then TempStr=DefaultChar
		 If Not IsNumeric(HtmlTag) Or Not IsNumeric(CutNum) Then Exit Function
		 If HtmlTag=1 Then
		  TempStr=Server.HtmlEncode(TempStr)
		 ElseIF HtmlTag=2 Then
		  TempStr=ACTCMS.CloseHtml(TempStr)
		 End If
          If EndTag="0" Then EndTag=""
		  if actcms.strLength(TempStr)>cint(CutNum) and CutNum<>0 then TempStr = actcms.GetStrValue(TempStr, CutNum) & EndTag
		 Get_Text_Field=TempStr
		End Function
		
		'取数字字段的值
		Function Get_Num_Field(FieldValue,OutType,XSWS)
		 If Not IsNumeric(FieldValue) Then Get_Num_Field=FieldValue:Exit Function
		 If Not IsNumeric(OutType) Then OutType=0
		 If Not IsNumeric(XSWS) Then XSWS=0
         If OutType=1 Then
		   Get_Num_Field=FormatNumber(FieldValue,XSWS)
		 ElseIf OutType=2 Then
		   Get_Num_Field=FormatPercent(FieldValue)
		 Else
		   Get_Num_Field=FieldValue
		 End if  
		End Function
		
		'取日期字段的值
		Function Get_Date_Field(FieldValue,DateMB)
		  IF Not IsDate(FieldValue) Then Get_Date_Field=FieldValue:Exit Function
		  Get_Date_Field=Replace(DateMB,"YYYY",Year(FieldValue))
		  Get_Date_Field=Replace(Get_Date_Field,"YY",Right("0" & Year(FieldValue), 2))
		  Get_Date_Field=Replace(Get_Date_Field,"MM",Right("0" & Month(FieldValue), 2))
		  Get_Date_Field=Replace(Get_Date_Field,"DD",Right("0" & Day(FieldValue), 2))
		  Get_Date_Field=Replace(Get_Date_Field,"hh",Right("0" & hour(FieldValue), 2))
		  Get_Date_Field=Replace(Get_Date_Field,"mm",Right("0" & minute(FieldValue), 2))
		  Get_Date_Field=Replace(Get_Date_Field,"ss",Right("0" & second(FieldValue), 2))
		End Function
		
		'取对象的链接URL
		Function Get_InfoUrl_Field(byval FieldName,byval FieldValue,ModeID,OutType)
		 If OutType=2  Then Get_InfoUrl_Field=FieldValue:Exit Function
		 Dim SqlStr
		 If Not Isnumeric(ModeID) Then Exit Function
		  SqlStr="Select ID,Classid,Title,UpdateTime,actlink,FileName,InfoPurview,ReadPoint From " & ACTCMS.ACT_C(ModeID,2) & " Where " & FieldName &"=" &FieldValue
		   Dim ARs:Set ARs=Server.CreateObject("ADODB.RECORDSET")
		   ARs.Open SqlStr,Conn,1,1
		  IF ARs.Eof Then
			   ARs.Close:Set ARs=Nothing:Exit Function
			  Else
					If OutType=0 Then
					 Get_InfoUrl_Field="<a href="""&AcTCMS.GetInfoUrl(ModeID,ARs(1),ARs(0),ARs(4),ARs(5),ARs(6),ARs(7))&""" target=""_blank"">" & FieldValue &"</a>"
					ElseIF OutType=1 Then
					 Get_InfoUrl_Field=AcTCMS.GetInfoUrl(ModeID,ARs(1),ARs(0),ARs(4),ARs(5),ARs(6),ARs(7))
					End If		
			  End if
			  ARs.Close:Set ARs=Nothing
		End Function
		'得到栏目的链接URL
		Function Get_ClassUrl_Field(FieldName,FieldValue,ModeID,OutType)
		  If OutType=2 Then Get_ClassUrl_Field=FieldValue:Exit Function
		  Dim ClassID:ClassID=FieldValue
			 If FieldName="id" Then
				 Dim SqlStr:SqlStr="Select Classid From Class_ACT Where Classid='" & Conn.Execute("Select Classid From " & ACTCMS.ACT_C(ModeID,2) & " Where " & FieldName &"=" &FieldValue)(0)&"'"
				Dim ARs:Set ARs=Server.CreateObject("ADODB.RECORDSET")
				ARs.Open SqlStr,Conn,1,1
				IF ARs.Eof Then
				   ARs.Close:Set ARs=Nothing:Exit Function
				Else
				   ClassID  = ARs(0)
				End if
				ARs.Close:Set ARs=Nothing
			 End IF
		     If OutType=0 Then
				 Get_ClassUrl_Field="<a href="""&actcms.DiyClassName(ClassID)&""" target=""_blank"">" & ACTCMS.ACT_L(classID,2) &"</a>"
			ElseIF OutType=1 Then
				 Get_ClassUrl_Field=actcms.DiyClassName(ClassID)
			 End If
		  
		End Function
		Function OpenExtConn()
		 If DataSourceType=0 Then
		   OpenExtConn=True
		 Else
			on error resume next
		    Set tconn = Server.CreateObject("ADODB.Connection")
			tconn.open datasourcestr
			If DataSourceType=7 Then
			 tconn.execute("set names 'utf-8'")
			End If
			If Err Then 
			  Err.Clear
			  Set tconn = Nothing
			   OpenExtConn=False
			Else 
			   OpenExtConn=true
			End If
		 End If
    	End Function
	End Class
%> 
