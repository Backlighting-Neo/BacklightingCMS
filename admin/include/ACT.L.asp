<!--#include file="../ACT.Function.asp"-->
<!--#include file="../../ACT_inc/ACT.Code.asp"-->
<% 		With Response
		Dim ACTCode,ModeID
		Set ACTCode =New ACT_Code
		Dim StartRefreshTime,RefreshFlag
		Dim FsoHtmlList,ItemName
		RefreshFlag = Request("RefreshFlag")
		StartRefreshTime = Request("StartRefreshTime")
 		If StartRefreshTime = "" Then StartRefreshTime = Timer()
	Server.ScriptTimeOut=9999999
	Call MakeFolder
	End With
	 Set ACTCode=Nothing:Set ACTCMS=Nothing
		Sub Main()
		Dim ReturnInfo
		  With Response
		  .Write ("<html>")
		  .Write ("<head>")
		  .Write ("<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">")
		  .Write ("<title>系统信息</title>")
		  .Write ("</head>")
		  .Write ("<body>")
		  If RefreshFlag="ID" Then
              .Write "<div style=""display:none"">"
				.Write "<br><br><br><table style=""display:none"" id=""BarShowArea"" width=""400"" border=""0"" align=""center"" cellspacing=""1"" cellpadding=""1"">"
		 Else
				.Write "<br><br><br><table id=""BarShowArea"" width=""400"" border=""0"" align=""center"" cellspacing=""1"" cellpadding=""1"">"
		 End iF
				.Write "<tr> "
				.Write "<td bgcolor=000000>"
				.Write " <table width=""400"" border=""0"" cellspacing=""0"" cellpadding=""1"">"
				.Write "<tr> "
				.Write "<td bgcolor=ffffff height=20><img src=""../images/bar9.gif"" width=0 height=19 id=img2 name=img2 align=absmiddle></td></tr></table>"
				.Write "</td></tr></table>"
				.Write "<table width=""550"" border=""0"" align=""center"" cellspacing=""1"" cellpadding=""1""><tr> "
				.Write "<td align=center> <span id=txt2 name=txt2 style=""font-size:9pt"">0</span><span id=txt4 style=""font-size:9pt"">%</span></td></tr> "
				.Write "<tr><td align=center><span id=txt3 name=txt3 style=""font-size:9pt"">0</span></td></tr>"
				.Write "</table>"
		 .Write ("   <div id=""fsohtml"">")
		 .Write (FsoHtmlList)
		  .Write ("   </div>")
		 .Write ("</body>")
		 .Write ("</html>")
		 End With
		End Sub
		Sub MakeFolder()
		With Response
		 Dim ClassID, FolderSql, MaxNum, FolderRS, NewsTotalNum, NewsNo	  
		 If NewsNo = "" Then NewsNo = 0
		  Select Case RefreshFlag
		    Case "IDS"
			   ClassID = RSQL(Trim(Request("CID")))
			    ModeID=actcms.act_l(ClassID,10)
			 	Application(AcTCMSN&"ModeID")=ModeID
 			    FolderSql = "Select * from Class_ACT where  ClassID ='" & ClassID & "'  and  actlink <>2  "
 			Case "Folder"
				ClassID = Trim(Request("CID"))
 				FolderSql = "Select * from Class_ACT where  ClassID IN (" & ClassID & ")  and  actlink  <>2    Order By OrderID ASC"
		   Case "All"
				FolderSql = "Select * from Class_ACT where  actlink <>2  Order By OrderID ASC"
		   Case Else
			FolderSql = ""
		  End Select
   			Call Main
			If FolderSql <> "" Then
 			Set FolderRS = Server.CreateObject("ADODB.RecordSet")
			FolderRS.Open FolderSql, Conn, 1, 1

				If FolderRS.EOF Then
					.Write "<script>img2.width=""0"";" & vbCrLf
					.Write "txt2.innerHTML=""没有可生成的" & ItemName & "栏目！<br><br><input name='button1' type='button' onclick=javascript:location.href='ACT.Make.asp?ModeID=" & ModeID &"'; class='button' value=' 返 回 '>"";" & vbCrLf
					.Write "txt3.innerHTML="""";" & vbCrLf
					.Write "txt4.innerHTML="""";" & vbCrLf
					.Write "document.all.BarShowArea.style.display='none';" & vbCrLf
					.Write "</script>" & vbCrLf
					FolderRS.Close:Set FolderRS = Nothing
					.end 
				Else
				   NewsTotalNum = FolderRS.RecordCount
				   For NewsNo=1 to NewsTotalNum
				
					   ModeID=actcms.act_l(FolderRS("ClassID"),10)
				   	   Application(AcTCMSN&"ModeID")=ModeID
  						If FolderRS("GroupIDClass")<>""  Or  ACTCMS.ACT_C(ModeID,3) = 0 Or ACTCMS.ACT_C(ModeID,3) = 2  Then 
							FsoHtmlList="<div align=center><li>栏目名称为<font color=red>"  & FolderRS("ClassName") & "</font>的没有生成</div>"
						Else
							 FsoHtmlList="<div align=center><li>栏目名称为<font color=red>"  & FolderRS("ClassName")& "</font>已生成</div>"
					 	Call ACTCode.CreateArticleList(ModeID,FolderRS)
					
						End If 
				    If RefreshFlag="ID" Then Call TypeJS(NewsNo,NewsTotalNum,ACTCMS.ACT_C(ModeID,5) & "栏目"):.End
					Call TypeJS(NewsNo,NewsTotalNum,ACTCMS.ACT_C(ModeID,5) & "栏目")
					FolderRS.MoveNext
					if Not Response.IsClientConnected then Exit FOR
				  Next
				.Write "<script>"
				.Write "fsohtml.innerHTML='';" & vbCrLf
				.Write "img2.width=400;" & vbCrLf
				.Write "txt2.innerHTML=""生成" & ItemName & "栏目结束！100"";" & vbCrLf
				.Write "txt3.innerHTML=""总共生成了 <font color=red><b>" & NewsTotalNum & "</b></font> 个" & ItemName & "栏目,总费时:<font color=red>" & Left((Timer() - StartRefreshTime), 4) & "</font> 秒<br><br><input name='button1' type='button' onclick=javascript:location='ACT.Make.asp?ModeID=" & ModeID &"'; class='button' value=' 返 回 '>"";" & vbCrLf
				.Write "img2.title=""(" & NewsNo & ")"";</script>" & vbCrLf
				
				FolderRS.Close:Set FolderRS = Nothing
			End If
		Else
				.Write "<script>img2.width=""0"";" & vbCrLf
				.Write "txt2.innerHTML=""没有可生成的栏目！<br><br><input name='button1' type='button' onclick=javascript:location='ACT.Make.asp?ModeID=" & ModeID & "'; class='button' value=' 返 回 '>"";" & vbCrLf
				.Write "txt3.innerHTML="""";" & vbCrLf
				.Write "txt4.innerHTML="""";" & vbCrLf
				.Write "document.all.BarShowArea.style.display='none';" & vbCrLf
				.Write "</script>" & vbCrLf
		End If
		End With
		End Sub

		Sub TypeJS(NowNum,TotalNum,itemname)
		  With Response
				.Write "<script>"
				.Write "fsohtml.innerHTML='" & FsoHtmlList & "';" & vbCrLf
				.Write "img2.width=" & Fix((NowNum / TotalNum) * 400) & ";" & vbCrLf
				.Write "txt2.innerHTML=""生成进度:" & FormatNumber(NowNum / TotalNum * 100, 2, -1) & """;" & vbCrLf
				.Write "txt3.innerHTML=""总共需要生成 <font color=red><b>" & TotalNum & "</b></font> " & itemname & ",<font color=red><b>在此过程中请勿刷新此页面！！！</b></font> 系统正在生成第 <font color=red><b>" & NowNum & "</b></font> " & itemname & """;" & vbCrLf
				.Write "img2.title=""(" & NowNum & ")"";</script>" & vbCrLf
				.Flush
		  End With
		End Sub
 %>
