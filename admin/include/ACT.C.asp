<!--#include file="../ACT.Function.asp"-->
<!--#include file="../../ACT_inc/ACT.Code.asp"-->
<% 	
		Dim ACTCode,ModeID,Table
		Set ACTCode =New ACT_Code
		Dim StartRefreshTime,RefreshFlag
		Dim FsoHtmlList,ItemName,f,PauseNum,PauseTime,Types
		RefreshFlag = Request("RefreshFlag")
		StartRefreshTime = Request("StartRefreshTime")
 		f=Request("f")
		PauseNum=100'多少页刷新
		PauseTime=5'停止多长时间
		If StartRefreshTime = "" Then StartRefreshTime = Timer()
	ModeID =ChkNumeric(request("ModeID"))
	IF ModeID = 0 Then ModeID = 1
 	Server.ScriptTimeOut=9999999
	Table=actcms.ACT_C(ModeID,2)
	Application(AcTCMSN&"ModeID")=ModeID
	If ACTCMS.ACT_C(ModeID,3) = 0 Or ACTCMS.ACT_C(ModeID,3) = 2 Then Call ACTCMS.Alert("ACTCMS系统提醒您：\n\n1、此模型没有启用生成静态HTML功能\n\n2、请进入该模型->系统参数设置 >启用生成静态Html功能","../Mode/ACT.MX.asp?A=E&ModeID="&ModeID&"")
Call MakeContent
 	 Set ACTCode=Nothing:Set ACTCMS=Nothing
		Sub Main()
		Dim ReturnInfo
 		  Echo ("<html>")
		  Echo ("<head>")
		  Echo ("<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">")
		  Echo ("<title>系统信息</title>")
		  Echo ("</head>")
		  Echo ("<body>")
		  If RefreshFlag="ID" Then
              Echo "<div style=""display:none"">"
				Echo "<br><br><br><table style=""display:none"" id=""BarShowArea"" width=""400"" border=""0"" align=""center"" cellspacing=""1"" cellpadding=""1"">"
		 Else
				Echo "<br><br><br><table id=""BarShowArea"" width=""400"" border=""0"" align=""center"" cellspacing=""1"" cellpadding=""1"">"
		 End iF
				Echo "<tr> "
				Echo "<td bgcolor=000000>"
				Echo " <table width=""400"" border=""0"" cellspacing=""0"" cellpadding=""1"">"
				Echo "<tr> "
				Echo "<td bgcolor=ffffff height=20><img src=""../images/bar9.gif"" width=0 height=19 id=img2 name=img2 align=absmiddle></td></tr></table>"
				Echo "</td></tr></table>"
				Echo "<table width=""550"" border=""0"" align=""center"" cellspacing=""1"" cellpadding=""1""><tr> "
				Echo "<td align=center> <span id=txt2 name=txt2 style=""font-size:9pt"">0</span><span id=txt4 style=""font-size:9pt"">%</span></td></tr> "
				Echo "<tr><td align=center><span id=txt3 name=txt3 style=""font-size:9pt"">0</span></td></tr>"
				Echo "</table>"
		 Echo ("   <div id=""fsohtml"">")
		 Echo (FsoHtmlList)
		  Echo ("   </div>")
		 Echo ("</body>")
		 Echo ("</html>")
 		End Sub
		Sub MakeContent()
 		Dim AlreadyRefreshByID, NowNum, R_Sql, R_RS, TotalNum,ID 
		Dim StartDate, EndDate, ClassID, RefreshTotalNum,StartID,EndID
		AlreadyRefreshByID = Request.QueryString("AlreadyRefreshByID")
		RefreshTotalNum = Request.QueryString("RefreshTotalNum")
		NowNum = Request.QueryString("NowNum") '正在刷新第几篇文章
 		if request("ismake")="1" then
			R_Sql=" Where ismake=0 and isAccept=0 and delif=0 and actlink<>1"
		Else 
			R_Sql=" Where isAccept=0 and delif=0 and actlink<>1"
		End  If 
			If NowNum = "" Then NowNum = 0
 	Select Case RefreshFlag
		   Case "ID"
				ID =ACTCMS.S("ID")
 				R_Sql="Select Top 2 ID,ClassID,Title,UpdateTime,ActLink,FileName,infopurview,readpoint,PicUrl,Intro,Content,CopyFrom,Author,TemplateUrl,UserID,IntactTitle,KeyWords,ArticleInput"&ACTCMS.DIYFieldList(ModeID)&" From " & Table & R_SQL&" and ID IN(Select top 2 id from " & Table & R_Sql & " And ID<=" & id & " Order By ID Desc) Order By ID"
				RefreshTotalNum=conn.execute("select count(id) from " & Table  &" where isAccept=0 and ID<=" & ID)(0)
				If RefreshTotalNum>2 Then RefreshTotalNum=2
			Case "New"
			  TotalNum = ChkNumeric(Request("TotalNum"))
			  If TotalNum >conn.execute("select count(id) from "& Table )(0) Then TotalNum = conn.execute("select count(id) from "& Table )(0)
			  RefreshTotalNum = TotalNum
			  If TotalNum=0 Then TotalNum=1
			  R_Sql="Select Top " & TotalNum & " ID,ClassID,Title,UpdateTime,ActLink,FileName,infopurview,readpoint,PicUrl,Intro,Content,CopyFrom,Author,TemplateUrl,UserID,IntactTitle,KeyWords,ArticleInput"&ACTCMS.DIYFieldList(ModeID)&" from " & Table  & " Order By ID Desc"
		 Case "InfoID"
				 StartID = ChkNumeric(Request("StartID"))
				 EndID = ChkNumeric(Request("EndID"))
				 RefreshTotalNum=conn.execute("select count(id) from " & Table  & R_Sql & " and ID>= " & StartID & " And  ID <=" & EndID)(0)
				 R_Sql = "Select ID,ClassID,Title,UpdateTime,ActLink,FileName,infopurview,readpoint,PicUrl,Intro,Content,CopyFrom,Author,TemplateUrl,UserID,IntactTitle,KeyWords,ArticleInput"&ACTCMS.DIYFieldList(ModeID)&" from " & Table  & R_Sql & " and ID>= " & StartID & " And  ID <=" & EndID & " order by ID desc"
		   Case "Date"
			  StartDate = Request("StartDate"):EndDate = DateAdd("d", 1, Request("EndDate"))
				 RefreshTotalNum=conn.execute("select count(id) from " & Table  & R_Sql & " and UpdateTime>= #" & StartDate & "# And  UpdateTime <=#" & EndDate & "#")(0)
				 R_Sql = "Select ID,ClassID,Title,UpdateTime,ActLink,FileName,infopurview,readpoint,PicUrl,Intro,Content,CopyFrom,Author,TemplateUrl,UserID,IntactTitle,KeyWords,ArticleInput"&ACTCMS.DIYFieldList(ModeID)&" from " & Table  & R_Sql & " and UpdateTime>= #" & StartDate & "# And  UpdateTime <=#" & EndDate & "# order by ID desc"
		  Case "Folder"
			 ClassID = Trim(Replace(Request("CID")," ",""))
			
			 RefreshTotalNum=conn.execute("select count(id) from " & Table  & R_Sql& " And ClassID IN(" & ClassID & ")")(0)
			 R_Sql = "Select ID,ClassID,Title,UpdateTime,ActLink,FileName,infopurview,readpoint,PicUrl,Intro,Content,CopyFrom,Author,TemplateUrl,UserID,IntactTitle,KeyWords,ArticleInput"&ACTCMS.DIYFieldList(ModeID)&" from " & Table  & R_Sql & " And ClassID IN(" & ClassID & ") order by ID desc"
 		  Case "Make"
			ID =ACTCMS.S("ID")
 			 RefreshTotalNum=conn.execute("select count(id) from " & Table  & R_Sql& " And ID IN(" & ID & ")")(0)
			 R_Sql = "Select ID,ClassID,Title,UpdateTime,ActLink,FileName,infopurview,readpoint,PicUrl,Intro,Content,CopyFrom,Author,TemplateUrl,UserID,IntactTitle,KeyWords,ArticleInput"&ACTCMS.DIYFieldList(ModeID)&" from " & Table  & R_Sql & " And ID IN(" & ID & ") order by ID desc"
		  Case "All"
		      RefreshTotalNum=conn.execute("select count(id) from " & Table  & R_Sql)(0)
 			  R_Sql = "Select ID,ClassID,Title,UpdateTime,ActLink,FileName,infopurview,readpoint,PicUrl,Intro,Content,CopyFrom,Author,TemplateUrl,UserID,IntactTitle,KeyWords,ArticleInput"&ACTCMS.DIYFieldList(ModeID)&"  from " & Table  & R_Sql & " order by ID desc"
		  Case "Pause"
			 R_Sql=Request.QueryString("R_Sql")
			 RefreshTotalNum=ChkNumeric(ACTCMS.S("RefreshTotalNum"))
		 Case Else
			R_Sql = ""
			RefreshTotalNum = 0
		End Select
	 
	 	Call Main
 		
		If R_Sql <> "" Then
	    Set R_RS=Conn.Execute(R_Sql)
  			If R_RS.EOF And R_RS.BOF Then
				echo "<script>img2.width=""0"";" & vbCrLf
				echo "txt2.innerHTML=""没有可生成的内容页！<br><br><input name='button1' type='button' onclick=javascript:location='ACT.Make.asp?Action=ref&ModeID=" & ModeID  &"'; class='button' value=' 返 回 '>"";" & vbCrLf
				echo "txt3.innerHTML="""";" & vbCrLf
				echo "txt4.innerHTML="""";" & vbCrLf
				echo "document.all.BarShowArea.style.display='none';" & vbCrLf
				echo "</script>" & vbCrLf
				Response.Flush
				R_RS.Close:Set R_RS=Nothing
				Exit Sub
			Else
				'On Error Resume Next
				Dim CurrNowNum:CurrNowNum=ChkNumeric(ACTCMS.S("CurrNowNum"))
				If CurrNowNum=0 Then CurrNowNum=1
				R_RS.Move(CurrNowNum-1)
 				For NowNum=CurrNowNum To RefreshTotalNum
				   Dim Node
				   Dim DocXML:Set DocXML=actcms.arrayToXml(R_RS.GetRows(1),R_RS,"row","root")
				     Set Node=DocXml.DocumentElement.SelectSingleNode("row")
				     Set ACTCode.Nodes=DocXml.DocumentElement.SelectSingleNode("row")
					 FsoHtmlList=GetRefreshSucc(Node,ItemName)
				     conn.execute("update " & Table & " set ismake=1 where id=" & Node.SelectSingleNode("@id").text)
					 Call ACTCode.ArticleContent(ModeID)
 				If Err.Number <> 0 Then
				 FsoHtmlList = "操作失败!<br><font color=red>" & Err.Description & "</font>"
				 Call TypeJS(NowNum,RefreshTotalNum,ACTCMS.ACT_C(ModeID,5))
				End If
				If RefreshFlag="ID" and NowNum=2 Then
				 Call TypeJS(NowNum,RefreshTotalNum,ACTCMS.ACT_C(ModeID,5)):R_RS.Close:Set R_RS=Nothing:.Die ""
				Else
				 Call TypeJS(NowNum,RefreshTotalNum,ACTCMS.ACT_C(ModeID,5))
				End If
 				if Not Response.IsClientConnected then Exit FOR
				If PauseNum>0 Then
					If RefreshTotalNum>1 and NowNum Mod PauseNum=0 Then
					    R_RS.Close:Set R_RS=Nothing
						 echo "<script>"
						 echo "fsohtml.innerHTML='<div style=""text-align:cdenter""><div style=""margin:10px;height:80px;padding:8px;border:1px dashed #cccccc;text-align:left;""><img src=""../images/succeed.gif"" align=""left""><br>&nbsp;&nbsp;&nbsp;&nbsp;<b>温馨提示：</b><br>&nbsp;&nbsp;&nbsp;&nbsp;以免过度占用服务器资源，系统暂停" & PauseTime & "秒后继续<img src=""../images/wait.gif""><br>&nbsp;&nbsp;&nbsp;&nbsp;如果" & PauseTime & "秒后没有继续，请点此<a href=""ACT.C.asp?CurrNowNum=" & NowNum+1 & "&ModeID=" & ModeID & "&RefreshFlag=Pause&Types=" & Types & "&StartRefreshTime=" & StartRefreshTime & "&R_Sql=" & Server.UrlEncode(R_Sql) & "&RefreshTotalNum=" & RefreshTotalNum & """><font color=red>继续</font></a>或点此<a href=""ACT.C.asp?Action=ref&ModeID=" & ModeID & """><font color=red>停止</font></a>!</div></div>';" & vbCrLf
						 echo "</script>" &vbcrlf
						 die "<meta http-equiv=""refresh"" content=""" & PauseTime & ";url=ACT.C.asp?f=" & f & "&CurrNowNum=" & NowNum+1 & "&ModeID=" & ModeID & "&RefreshFlag=Pause&Types=" & Types & "&StartRefreshTime=" & StartRefreshTime & "&R_Sql=" & Server.UrlEncode(R_Sql) & "&RefreshTotalNum=" & RefreshTotalNum & """>"
					End If
			   End If
			Next	
 				echo "<script>"
				echo "fsohtml.innerHTML='';" & vbCrLf
				echo "img2.width=400;" & vbCrLf
				echo "txt2.innerHTML=""生成内容页结束！100"";" & vbCrLf
				echo "txt3.innerHTML=""总共生成了 <font color=red><b>" & RefreshTotalNum & "</b></font> 条,总费时:<font color=red>" & Left((Timer() - StartRefreshTime), 4) & "</font> 秒<br><br><input name='button1' type='button' onclick=javascript:location='ACT.Make.asp?ModeID=" & ModeID & "'; class='button' value=' 返 回 '>"";" & vbCrLf
				echo "img2.title=""(" & NowNum & ")"";</script>" & vbCrLf
 			End If
		Else
				echo "<script>img2.width=""0"";" & vbCrLf
				echo "txt2.innerHTML=""没有可生成的内容页！<br><br><input name='button1' type='button' onclick=javascript:location='ACT.Make.asp?ModeID=" & ModeID & "'; class='button' value=' 返 回 '>"";" & vbCrLf
				echo "txt3.innerHTML="""";" & vbCrLf
				echo "txt4.innerHTML="""";" & vbCrLf
				echo "document.all.BarShowArea.style.display='none';" & vbCrLf
				echo "</script>" & vbCrLf
 		End If
 		End Sub
 		Function GetRefreshSucc(Node,ItemName)
		 Dim str 
		 str=""
		 if RefreshFlag<>"ID" Then str="<img src=""../images/succeed.gif"" align=""left""><table border=""0"">"
		 GetRefreshSucc=str & "<table border=""0"">"_
									& "<tr><td><li><strong>ID 号为：</strong></li></td><td> <font color=red>"  & Node.SelectSingleNode("@id").text & "</font> 的" & ItemName & "已生成</td></tr>"_
									& "<tr><td><li><strong>" & ItemName & "标题：</strong></li></td><td><font color=red>" & Node.SelectSingleNode("@title").text & "</font></li></td><tr>" _
 									& "</table>"
  		End Function

	 


		Sub TypeJS(NowNum,TotalNum,itemname)
 				Echo "<script>"
				Echo "fsohtml.innerHTML='<div style=""margin:10px;height:80px;padding:8px;border:1px dashed #cccccc;text-align:left;"">" & FsoHtmlList & "</div>';" & vbCrLf
				Echo "img2.width=" & Fix((NowNum / TotalNum) * 400) & ";" & vbCrLf
				Echo "txt2.innerHTML=""生成进度:" & FormatNumber(NowNum / TotalNum * 100, 2, -1) & """;" & vbCrLf
				Echo "txt3.innerHTML=""总共需要生成 <font color=red><b>" & TotalNum & "</b></font> " & itemname & ",<font color=red><b>在此过程中请勿刷新此页面！！！</b></font> 系统正在生成第 <font color=red><b>" & NowNum & "</b></font> " & itemname & """;" & vbCrLf
				Echo "img2.title=""(" & NowNum & ")"";</script>" & vbCrLf
				response.Flush
 		End Sub
 %>
