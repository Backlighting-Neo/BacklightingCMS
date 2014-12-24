<!--#include file="../act_Inc/ACT.FreeLabel.asp" -->
 <%
Class ACT_Space
	
	
 	Public UID,UIDSQL
	Private Domain,ASys
	Private Sub Class_Initialize()
		Domain = AcTCMS.ACTURL
		ASys = Actcms.ActSys
 		 UID = ChkNumeric(urlarr(1))
		 If U="0" Or UID="0" Then response.write "errors":response.End
		 UIDSQL="   And userid="&UID&"  "
 		End Sub
        Private Sub Class_Terminate()
		End Sub
		Public Function LabelReplaceAll(TemplateContent)
		  TemplateContent=Loadfile(TemplateContent)
 		  TemplateContent = LableFlag(AllLabel(TemplateContent))
		  TemplateContent =ReplaceAllLabel(TemplateContent)
		  TemplateContent = GeneralLabel(TemplateContent)   
		  LabelReplaceAll = TemplateContent
 	    End Function
 

	Function  LoadTemplate(TempString) 
 		    'on error resume next
			Dim  Str,A_W
			set A_W=server.CreateObject("adodb.Stream")
			A_W.Type=2:A_W.mode=3:A_W.charset="utf-8":A_W.open
			A_W.loadfromfile server.MapPath(actcms.ThisTheme&"/"&TempString)
			If Err.Number<>0 Then Err.Clear:LoadTemplate="模板没有找到 <br> by BackLighting Software":Exit Function
			Str=A_W.readtext
			A_W.Close
			Set  A_W=nothing
			LoadTemplate=Str
 	End  function

	Function actcmsexe(TemplateContent)
 		dim HtmlLabel,HtmlLabelArr,i,Param
 		 If InStr(TemplateContent, "{=ACTEXE") > 0 Then
			 HtmlLabel = SelectLabelParameter(TemplateContent, "{=ACTEXE")
 			 HtmlLabelArr=Split(HtmlLabel,"$$$")
			 For I=0 To Ubound(HtmlLabelArr)
				 Param = Split(FunctionLabelParam(HtmlLabelArr(I), "{=ACTEXE"),",")
  				  TemplateContent = Replace(TemplateContent, HtmlLabelArr(I), actcms.AEXE(Ubound(Param),Param))
 			 Next
		 End If
 		 actcmsexe=aspexecute(TemplateContent)
	End Function 
	
 	Public Function Loadfile(TemplateContent) 
  		 'on error resume next
          If InStr(TemplateContent, "{file:") > 0 Then
            Dim Match, Matches, FileBody,FilePath
            Reg.Pattern = "{file:(.+?)}"
            Set Matches = Reg.Execute(TemplateContent)
            For Each Match In Matches
				FilePath= Match.SubMatches(0)
 				FileBody = LoadTemplate(FilePath)
   			    TemplateContent = Replace(TemplateContent, Match.Value,FileBody)
            Next
        End If
		Loadfile=TemplateContent
 	End Function

	Public Function aspexecute(TemplateContent)
		'on error resume next
 		Dim Matches, Match 
 		Reg.Pattern = "{aspexe:([\s\S]*?)}"
		Set Matches = Reg.Execute(TemplateContent)
		For Each Match In Matches
 			Execute(replace(Match.SubMatches(0),"'",""""))
   			TemplateContent = Replace(TemplateContent, Match.Value, aspexe) 
			If Err Then Response.Write "<font color=red>语法错误[" & Match.SubMatches(0) & "]" & Err.Description & "</font>": Err.Clear: Response.End
		Next
		aspexecute=TemplateContent
 	End Function
 	Public Function GeneralLabel(FileContent)
		' 'on error resume next
		 FileContent = ReplaceChannel(FileContent)'栏目标签
 		 Dim HtmlLabel,HtmlLabelArr, Param,I,Act_S
 		 FileContent = Replace(FileContent, "{$SiteName}",AcTCMS.ActCMS_Sys(0))
		 FileContent = Replace(FileContent, "{$SiteTitle}", AcTCMS.ActCMS_Sys(1))
		 FileContent = Replace(FileContent, "{$Keywords}", AcTCMS.ActCMS_Other(1))
		 FileContent = Replace(FileContent, "{$Description}", AcTCMS.ActCMS_Other(2))
		 FileContent = Replace(FileContent, "{$CopyRight}", AcTCMS.ActCMS_Other(0))
		 FileContent = Replace(FileContent, "{$SiteUrl}", AcTCMS.ACTUrl)
		 FileContent = Replace(FileContent, "{$InstallDir}", AcTCMS.ActCMSDM)
		 FileContent = Replace(FileContent, "{$Path}", AcTCMS.acturl)
		 FileContent = Replace(FileContent, "{$Now}", now)
		 FileContent = Replace(FileContent, "{$Skin}", ACTCMS.NowTheme)
		 FileContent = Replace(FileContent, "{$ThemePath}", ACTCMS.ThisTheme&"/")
 		 FileContent = Replace(FileContent, "{$Logo}", AcTCMS.ActCMS_Sys(5))
		 FileContent = Replace(FileContent, "{$AdminName}", AcTCMS.ActCMS_Sys(6))
		 FileContent = Replace(FileContent, "{$AdminMail}", AcTCMS.ActCMS_Sys(7))
		 FileContent = Replace(FileContent, "{$AdminDir}", AcTCMS.ActCMS_Sys(8))
		 If Trim(AcTCMS.ActCMS_Sys(26))<>"" Then 
			 FileContent = Replace(FileContent, "{$Beian}", "<a href=""http://www.miibeian.gov.cn"" rel=""nofollow"" target=""_blank"">"&AcTCMS.ActCMS_Sys(26)&"</a>")
		 Else 
			 FileContent = Replace(FileContent, "{$Beian}", "")
		 End If 
		 FileContent = Replace(FileContent, "{$Statcode}", AcTCMS.ActCMS_Sys(27))
 		 FileContent = Replace(FileContent, "{$JSUserlogin}", " <Script Language=""Javascript"" Src="""&actcms.acturl&"User/Userlogin.asp?A=JS""></Script>")
 		 FileContent = Replace(FileContent, "{$actcms}", "Powered by <A href=""http://www.actcms.com"" target=""_blank"">ACTCMS</a> 3.0")
 		 If InStr(FileContent, "{=GetTags") <> 0 Then
			 HtmlLabel = SelectLabelParameter(FileContent, "{=GetTags")
			 HtmlLabelArr=Split(HtmlLabel,"$$$")
			 For I=0 To Ubound(HtmlLabelArr)
				 Param = Split(FunctionLabelParam(HtmlLabelArr(I), "{=GetTags"),",")
				 FileContent = Replace(FileContent, HtmlLabelArr(I), GetTags(Param(0),Param(1)))
			 Next
		 End If

		 If InStr(FileContent, "{=TodayRenewal") <> 0 Then
			 HtmlLabel = SelectLabelParameter(FileContent, "{=TodayRenewal")
			 HtmlLabelArr=Split(HtmlLabel,"$$$")
			 For I=0 To Ubound(HtmlLabelArr)
				 Param = Split(FunctionLabelParam(HtmlLabelArr(I), "{=TodayRenewal"),",")
				 FileContent = Replace(FileContent, HtmlLabelArr(I), AcTCMS.TodayRenewal(Param(0)))
			 Next
		 End If
 
	     		
 		 If InStr(FileContent, "{=CountClass") <> 0 Then
			 HtmlLabel = SelectLabelParameter(FileContent, "{=CountClass")
			 HtmlLabelArr=Split(HtmlLabel,"$$$")
			 For I=0 To Ubound(HtmlLabelArr)
				 Param = Split(FunctionLabelParam(HtmlLabelArr(I), "{=CountClass"),",")
				 FileContent = Replace(FileContent, HtmlLabelArr(I), AcTCMS.CountClass(Param(0)))
			 Next
		 End If

		 If InStr(FileContent, "{=SysCount") <> 0 Then
			 HtmlLabel = SelectLabelParameter(FileContent, "{=SysCount")
			 HtmlLabelArr=Split(HtmlLabel,"$$$")
			 For I=0 To Ubound(HtmlLabelArr)
				 Param = Split(FunctionLabelParam(HtmlLabelArr(I), "{=SysCount"),",")
				 FileContent = Replace(FileContent, HtmlLabelArr(I), AcTCMS.SysCount(Param(0)))
			 Next
		 End If

		If InStr(FileContent, "{=UserLogin") <> 0 Then
		 HtmlLabel = SelectLabelParameter(FileContent, "{=UserLogin")
		 HtmlLabelArr=Split(HtmlLabel,"@@@")
		 For I=0 To Ubound(HtmlLabelArr)
			 Param = Split(FunctionLabelParam(HtmlLabelArr(I), "{=UserLogin"),",")
			 FileContent = Replace(FileContent, HtmlLabelArr(I), "<iframe Width="&Param(0)&" height="&Param(1)&" ID=""loginframe"" name=""loginframe"" src=""" & Domain & "User/Userlogin.asp?A=Html"" frameBorder=""0"" scrolling=""no"" allowtransparency=""true""></iframe>")
	  	 Next
	    End If
		 GeneralLabel = FileContent
    End Function

	Function ReplaceChannel(FileContent)
		 'on error resume next
		 If Application(AcTCMSN & "ACTCMS_TCJ_Type")<>"Folder"   Then ReplaceChannel=FileContent:Exit Function
		 If InStr(FileContent, "{$SeoTitle}") > 0  Then
			If Trim(Actcms.ACT_L(Application(AcTCMSN & "classid"),25))<>"" Then 
				 FileContent = Replace(FileContent, "{$SeoTitle}",Actcms.ACT_L(Application(AcTCMSN & "classid"),25))
			Else 
				 FileContent = Replace(FileContent, "{$SeoTitle}",Actcms.ACT_L(Application(AcTCMSN & "classid"),2))
			End If 
 		 End If 
		 FileContent = Replace(FileContent, "{$ClassPicUrl}",Actcms.ACT_L(Application(AcTCMSN & "classid"),26))
		 FileContent = Replace(FileContent, "{$ClassID}",Actcms.ACT_L(Application(AcTCMSN & "classid"),0))
		 FileContent = Replace(FileContent, "{$ClassName}",Actcms.ACT_L(Application(AcTCMSN & "classid"),2))
		 FileContent = Replace(FileContent, "{$ClassKeywords}", Actcms.ACT_L(Application(AcTCMSN & "classid"),8))
		 FileContent = Replace(FileContent, "{$ClassDescription}", Actcms.ACT_L(Application(AcTCMSN & "classid"),9))
 		 ReplaceChannel = FileContent
	End Function 
 	Function GetTags(Num,TagType)
	  'on error resume next
	  if not isnumeric(num) then exit function
	  dim sqlstr,sql,i,n,str
	  select case cint(tagtype)
	   case 1:sqlstr="select top "&Num&" TagsChar,ModeID from Tags_ACT order by hits desc"
	   case 2:sqlstr="select top "&Num&" TagsChar,ModeID from Tags_ACT order by ClicksTime desc,ID desc"
	   case 3:sqlstr="select top "&Num&" TagsChar,ModeID from Tags_ACT order by AddTime desc,ID desc"
	   Case Else : sqlstr="select top "&Num&" TagsChar,ModeID from Tags_ACT order by hits desc"
	  end Select
	  dim rs:set rs=ACTCMS.ActExe(sqlstr)
	  if rs.eof then rs.close:set rs=nothing:exit function
	  sql=rs.getrows(-1)
	  rs.close:set rs=Nothing
	  for i=0 to ubound(sql,2)
	   if Actcms.FoundInArr(str,sql(0,i),",")=false Then
		n=n+1
		str=str & "," & sql(0,i)
		gettags=gettags & "<span><a href=""" & Domain & "plus/search/index.asp?searchtype=3&ModeID=" & sql(1,i) & "&keyword=" & sql(0,i)& """ target=""_blank"">" & sql(0,i) & "</a></span>"& vbCrLf
	   end if
	   if n>=cint(num) then exit for
	  next
	End Function
 	  Public Function ArrayToxml(DataArray,Recordset,row,xmlroot)
			Dim i,node,rs,j
			If xmlroot="" Then xmlroot="xml"
			Set ArrayToxml = Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			ArrayToxml.appendChild(ArrayToxml.createElement(xmlroot))
			If row="" Then row="row"
			For i=0 To UBound(DataArray,2)
				Set Node=ArrayToxml.createNode(1,row,"")
				j=0
				For Each rs in Recordset.Fields
						 node.attributes.setNamedItem(ArrayToxml.createNode(2,LCase(rs.name),"")).text= DataArray(j,i)& ""
						 j=j+1
				Next
				ArrayToxml.documentElement.appendChild(Node)
			Next
		End Function
		
		Function AllLabel(Content)
  			Dim  node 
			If not IsObject(Application(AcTCMSN&"_ReplaceAllLabel")) then
					Dim RS:Set RS = Server.CreateObject("ADODB.Recordset")
 					Set Rs=ACTCMS.ACTEXE("Select LabelType,LabelName,LabelContent from Label_ACT")
					if Not RS.eof then
						Set Application(AcTCMSN&"_ReplaceAllLabel")=ArrayToXml(RS.GetRows(-1),rs,"row","")
					end if
					RS.Close:Set RS = Nothing
 			End if
			For Each Node In Application(AcTCMSN&"_ReplaceAllLabel").documentElement.SelectNodes("row")
 				If Node.SelectSingleNode("@labeltype").text = "2" Then 
					Content = Replace(Content, Node.SelectSingleNode("@labelname").text, FreeLabel(Node.SelectSingleNode("@labelcontent").text))
				Else 
				    Content = Replace(Content, Node.SelectSingleNode("@labelname").text, Node.SelectSingleNode("@labelcontent").text)
				End If 
  			Next   
 			AllLabel = Content
		End Function

		Function ReplaceAllLabel(Content)
			Dim D:Set D=New ACTFreeLabel
			Content=D.ReplaceReeLabel(Content) '替换自定义函数标签 
			Set D=nothing
			ReplaceAllLabel =Content
		End Function
 
		'替换自由标签为内容
		Function FreeLabel(Content)
  			Dim  node 
			If not IsObject(Application(AcTCMSN&"_ReplaceAllLabel")) then
				Dim RS:Set RS = Server.CreateObject("ADODB.Recordset")
				Set Rs=ACTCMS.ACTEXE("Select  LabelName,LabelContent from Label_ACT")
				if Not RS.eof then
					Set Application(AcTCMSN&"_ReplaceAllLabel")=ArrayToXml(RS.GetRows(-1),rs,"row","")
				end if
				RS.Close:Set RS = Nothing
 			End if
			For Each Node In Application(AcTCMSN&"_ReplaceAllLabel").documentElement.SelectNodes("row")
				Content = Replace(Content, Node.SelectSingleNode("@labelname").text, Node.SelectSingleNode("@labelcontent").text)
  			Next   
			FreeLabel = Content
		End Function

		Function LableFlag(Content)
			Dim regEx, Matches, Match, TempStr
			Set regEx = New RegExp
			regEx.Pattern = "{\$[^{\$}]*}"
			regEx.IgnoreCase = True
			regEx.Global = True
			Set Matches = regEx.Execute(Content)
			LableFlag = Content
			For Each Match In Matches
				on error resume next
				TempStr = Match.Value
				TempStr = Replace(TempStr, Chr(13) & Chr(10), "")
				TempStr = Replace(TempStr, "{$", "")
				TempStr = Replace(TempStr, "}", "")
				TempStr = Left(TempStr, InStr(TempStr, "(") - 1) & "§" & MID(TempStr, InStr(TempStr, "(") + 1)
				TempStr = Left(TempStr, InStrRev(TempStr, ")") - 1)
				'TempStr = Replace(TempStr, """", "")
				If Err.Number = 0 Then
					LableFlag = Replace(LableFlag, Match.Value, MakeLablelFunction(TempStr))'转换标签
				End If
			Next
		End Function	

 		Function MakeLablelFunction(LabelContent)
		   Dim LabelArr:LabelArr = Split(LabelContent, "§")
			If LabelArr(0) = "" Then
				  MakeLablelFunction = ""
				  Exit Function
			End If
			Select Case UCase(LabelArr(0))
				 Case "GETARTICLELIST"
							MakeLablelFunction = ACT_A_List(LabelArr(1), LabelArr(2), LabelArr(3), LabelArr(4), LabelArr(5), LabelArr(6), LabelArr(7), LabelArr(8), LabelArr(9), LabelArr(10), LabelArr(11), LabelArr(12), LabelArr(13),LabelArr(14),LabelArr(15),LabelArr(16),LabelArr(17),LabelArr(18),LabelArr(19),LabelArr(20),LabelArr(21),LabelArr(22),LabelArr(23),LabelArr(24),LabelArr(25))'函数调用并执行SQL返回结果
				 Case "GETNAVIGATION"
							  MakeLablelFunction =GetNavigation(LabelArr(1), LabelArr(2), LabelArr(3), LabelArr(4), LabelArr(5))
				 Case "GETSPECIAL"'GetSpecial
							  MakeLablelFunction =GetSpecial(LabelArr(1), LabelArr(2), LabelArr(3), LabelArr(4))
				 Case "GETLINKLIST"
					   MakeLablelFunction = GetLinkList(LabelArr(1), LabelArr(2), LabelArr(3), LabelArr(4), LabelArr(5), LabelArr(6), LabelArr(7), LabelArr(8), LabelArr(9), LabelArr(10), LabelArr(11))
				Case "GETARTICLEPIC"'图文混排
							 MakeLablelFunction =ACT_P(LabelArr(1), LabelArr(2), LabelArr(3), LabelArr(4), LabelArr(5), LabelArr(6), LabelArr(7), LabelArr(8), LabelArr(9), LabelArr(10), LabelArr(11), LabelArr(12),LabelArr(13),LabelArr(14),LabelArr(15),LabelArr(16),LabelArr(17),LabelArr(18),LabelArr(19))
				 Case "GETSLIDE" '幻灯片
							 MakeLablelFunction = ACTCMS_GetSlIDe(LabelArr(1), LabelArr(2), LabelArr(3), LabelArr(4), LabelArr(5), LabelArr(6))
				 Case "GETLASTARTICLELIST"  '文章分页列表函数
						If 	 AcTCMS.ACT_C(Application(AcTCMSN & "modeid"),3) = "0" Or Application(AcTCMSN & "Make")="No" Then 
								MakeLablelFunction=LabelContent
								Application(AcTCMSN &"PageParam")=LabelContent
								Application(AcTCMSN & "Make")="Yes"
						Else 
								MakeLablelFunction = GetLastArticleList(LabelArr(1), LabelArr(2), LabelArr(3), LabelArr(4), LabelArr(5), LabelArr(6), LabelArr(7), LabelArr(8), LabelArr(9), LabelArr(10), LabelArr(11), LabelArr(12), LabelArr(13),LabelArr(14),LabelArr(15),LabelArr(16),LabelArr(17),LabelArr(18),LabelArr(19),LabelArr(20),LabelArr(21),LabelArr(22),LabelArr(23))
						End If 
				Case "GETCLASSNAVIGATION"'总导航和栏目导航
							 MakeLablelFunction = GetClassNavigation(LabelArr(1), LabelArr(2), LabelArr(3), LabelArr(4), LabelArr(5), LabelArr(6), LabelArr(7), LabelArr(8), LabelArr(9), LabelArr(10), LabelArr(11))
				Case "CORRELATIONARTICLELIST"
							 MakeLablelFunction = ACT_Correlation_Article(LabelArr(1), LabelArr(2), LabelArr(3), LabelArr(4), LabelArr(5), LabelArr(6), LabelArr(7), LabelArr(8), LabelArr(9), LabelArr(10), LabelArr(11), LabelArr(12), LabelArr(13),LabelArr(14),LabelArr(15),LabelArr(16),LabelArr(17),LabelArr(18))
				Case "GETCLASSFORARTICLELIST"
							 MakeLablelFunction = GetClassForArticleList(LabelArr(1), LabelArr(2), LabelArr(3), LabelArr(4), LabelArr(5), LabelArr(6), LabelArr(7), LabelArr(8), LabelArr(9), LabelArr(10), LabelArr(11), LabelArr(12), LabelArr(13),LabelArr(14),LabelArr(15),LabelArr(16),LabelArr(17),LabelArr(18),LabelArr(19),LabelArr(20),LabelArr(21),LabelArr(22),LabelArr(23),LabelArr(24),LabelArr(25),LabelArr(26),LabelArr(27),LabelArr(28),LabelArr(29),LabelArr(30),LabelArr(31))'函数调用并执行SQL返回结果
				Case Else
					   MakeLablelFunction = LabelArr(0)&"标签执行错误"
					   Exit Function
				 End Select
		End Function


		Function FSOSaveFile(Templetcontent,FileName)
			Templetcontent=actcmsexe(Templetcontent)
			'on error resume next 
			Dim FileFSO,FileType
			 Set FileFSO = Server.CreateObject("ADODB.Stream")
				With FileFSO
				.Type = 2
				.Mode = 3
				.Open
				.Charset = "utf-8"
				.Position = FileFSO.Size
				.WriteText  Templetcontent
				.SaveToFile Server.MapPath(FileName),2
				If Err.Number<>0 Then 
					Err.Clear 
					Call actcms.InsertLog(RSQL(Request.Cookies(AcTCMSN)("AdminName")),4,"生成错误",Request.ServerVariables("QUERY_STRING"))
					Exit Function 
				End If 
				.Close
				End With
			Set FileType = nothing
			Set FileFSO = nothing
		End Function

		 Function ReplaceArticleContent(ModeID,TempletContent,ArticleContents)
				Dim TempStr
			'   'on error resume next 
			   ArticleContents=ACTCMS.ReplaceSitelink(ArticleContents)
			   If InStr(TempletContent, "{$ArticleSize}") <> 0 Then
				   ArticleContents = "<span ID=""ContentArea"">" & ArticleContents & "</span>"
				   TempStr = "<script Language=Javascript>" & _
					  "function ContentSize(size)" & _
					  "{document.all.ContentArea.style.fontSize=size+""px"";}" & _
					  "</script>"
				  TempStr = TempStr & "【字体：<A href=""javascript:ContentSize(16)"">大</A> <A href=""javascript:ContentSize(14)"">中</A> <A href=""javascript:ContentSize(12)"">小</A>】"
				  TempletContent = Replace(TempletContent, "{$ArticleSize}", TempStr)
			  End If
			TempletContent=ReplaceMX(ModeID,TempletContent)
 			'TempletContent = Replace(TempletContent,"{$ArticleContent}",ArticleContents)
			TempletContent = Replace(TempletContent,"{$ArticleTitle}",GetNodeText("title"))
 			If InStr(TempletContent, "{$KeyTags}") > 0  Then
				TempletContent = Replace(TempletContent, "{$KeyTags}",ReplaceKeyTags(1,GetNodeText("keywords")))
			End if
 			If InStr(TempletContent, "{$CTitle}") > 0  Then
				TempletContent = Replace(TempletContent, "{$CTitle}",ACTCMS.CloseHtml(GetNodeText("title")))
			End if


 			If InStr(TempletContent, "{$CommentYes}") > 0  Then'审核通过的评论
 			 TempletContent = Replace(TempletContent, "{$CommentCount}", "<Script Language=""Javascript"" Src=""" & Domain & "plus/Comment/Comment.List.asp?Action=CommentYes&ModeID="&ModeID&"&ClassID=" & GetNodeText("classid") & "&ID=" & GetNodeText("id") & """></Script>")
			
			End if

 			If InStr(TempletContent, "{$CommentCount}") > 0  Then'共多少评论
 			
			 TempletContent = Replace(TempletContent, "{$CommentCount}", "<Script Language=""Javascript"" Src=""" & Domain & "plus/Comment/Comment.List.asp?Action=CommentCount&ModeID="&ModeID&"&ClassID=" & GetNodeText("classid") & "&ID=" & GetNodeText("id") & """></Script>")
 			End if
 
  			If InStr(TempletContent, "{$ArticleAuthor}") > 0  Then
 				If GetNodeText("userid")=0 Then 
						If GetNodeText("author")<>"" Then 
							TempletContent = Replace(TempletContent, "{$ArticleAuthor}",ACTCMS.Author(GetNodeText("author")))
						Else 
 							TempletContent = Replace(TempletContent, "{$ArticleAuthor}",GetNodeText("articleinput"))
						End If 
 				Else 
					If ACTCMS.UserM(GetNodeText("userid"))=False Then 
					
						If GetNodeText("author")<>"" Then 
							TempletContent = Replace(TempletContent, "{$ArticleAuthor}",ACTCMS.Author(GetNodeText("author")))
						Else 
							TempletContent = Replace(TempletContent, "{$ArticleAuthor}",GetNodeText("articleinput"))
						End If 
					Else
						TempletContent = Replace(TempletContent, "{$ArticleAuthor}",ACTCMS.UserM(GetNodeText("userid")))
					End If 
				End If 
 			End If

			If Not IsNull(GetNodeText("copyfrom")) And Trim(GetNodeText("copyfrom")) <> "" Then
			   TempletContent = Replace(TempletContent, "{$ArticleCopyFrom}", ACTCMS.CopyFrom(GetNodeText("copyfrom")))
			Else
			   TempletContent = Replace(TempletContent, "{$ArticleCopyFrom}", "本站原创")
			End If
		
 

			If InStr(TempletContent, "{$UserID}") > 0  Then
				TempletContent = Replace(TempletContent, "{$UserID}", GetNodeText("userid"))
			End If 

			If InStr(TempletContent, "{$PicUrl}") > 0   And Trim(GetNodeText("picurl")) <> "" Then
				TempletContent = Replace(TempletContent, "{$PicUrl}",ACTCMS.PathDoMain&GetNodeText("picurl"))
			Else
				TempletContent = Replace(TempletContent, "{$PicUrl}","")
			End If 

 			If InStr(TempletContent, "{$ArticleUrl}") > 0  Then
				TempletContent = Replace(TempletContent, "{$ArticleUrl}", ACTCMS.GetInfoUrlall(ModeID,GetNodeText("classid"),GetNodeText("id"),GetNodeText("actlink"),GetNodeText("filename"),GetNodeText("infopurview"),GetNodeText("readpoint")))
			End If 

			If InStr(TempletContent, "{$ClassName}") > 0  Then
				TempletContent = Replace(TempletContent, "{$ClassName}", Actcms.ACT_L(GetNodeText("classid"),2))
			End If 


			If InStr(TempletContent, "{$IntactTitle}") <> 0 And Trim(GetNodeText("intacttitle")) <> ""  Then
				TempletContent = Replace(TempletContent, "{$IntactTitle}", GetNodeText("intacttitle"))
			Else
				TempletContent = Replace(TempletContent, "{$IntactTitle}", GetNodeText("title"))
			End If 

			If InStr(TempletContent, "{$ArticleKeyWord}") > 0  Then
				TempletContent = Replace(TempletContent, "{$ArticleKeyWord}", GetNodeText("keywords"))
			End If 
			If InStr(TempletContent, "{$ID}") > 0  Then
				TempletContent = Replace(TempletContent, "{$ID}", GetNodeText("id"))
			End If 

			If InStr(TempletContent, "{$ClassID}") > 0  Then
				TempletContent = Replace(TempletContent, "{$ClassID}", Application(AcTCMSN & "classid"))
			End If 
			If InStr(TempletContent, "{$ModeID}") > 0  Then
				TempletContent = Replace(TempletContent, "{$ModeID}", ModeID)
			End If 

			If InStr(TempletContent, "{$ArticleHits}") <> 0 Then
			 TempletContent = Replace(TempletContent, "{$ArticleHits}", "<Script Language=""Javascript"" Src=""" & Domain & "Plus/ACT.Hits.asp?ModeID="&ModeID&"&ID=" & GetNodeText("id") & """></Script>")
			End If
			If InStr(TempletContent, "{$ArticleDate}") <> 0 Then
			   TempletContent = Replace(TempletContent, "{$ArticleDate}", Year(GetNodeText("updatetime")) & "年" & Right("0" & Month(GetNodeText("updatetime")), 2) & "月" & Right("0" & Day(GetNodeText("updatetime")), 2)&"日")
			End If
			If InStr(TempletContent, "{$ArticleIntro}") >0  Then
				If Trim(GetNodeText("intro")) <> "" Then 
					TempletContent = Replace(TempletContent, "{$ArticleIntro}", GetNodeText("intro"))
				Else
					TempletContent = Replace(TempletContent, "{$ArticleIntro}", ACTCMS.GetStrValue(Replace(Replace(Replace(AcTCMS.CloseHtml(GetNodeText("content")), vbCrLf, ""), "[NextPage]", ""), "&nbsp;", ""),200))
				End If 
			End If 

		   If InStr(TempletContent, "{$TypeComment}")  Then
			 TempletContent = Replace(TempletContent, "{$TypeComment}", "<Script Language=""Javascript"" Src=""" & Domain & "plus/Comment/Comment.js.asp?ModeID="&ModeID&"&ClassID=" & GetNodeText("classid") & "&ID=" & GetNodeText("id") & """></Script>")
		   Else
			TempletContent = Replace(TempletContent, "{$TypeComment}", "")
		   End If
		   If InStr(TempletContent, "{$WriteComment}") > 0 Then
			 TempletContent = Replace(TempletContent, "{$WriteComment}", "<Script Language=""Javascript"" Src=""" & Domain & "plus/Comment/ACT.Comment.asp?Action=Write&ModeID="&ModeID&"&ClassID=" & GetNodeText("classid") & "&ID=" & GetNodeText("id") & """></Script>")
		   Else
			 TempletContent = Replace(TempletContent, "{$WriteComment}", "")
		   End If
		 	 TempletContent = Replace(TempletContent, "{$PrevArticle}", NextArticle(GetNodeText("id"), GetNodeText("classid"), "Prev",ModeID))
		 	 TempletContent = Replace(TempletContent, "{$NextArticle}", NextArticle(GetNodeText("id"), GetNodeText("classid"), "Next",ModeID))
			 ReplaceArticleContent = TempletContent
 		 End Function

		Function ReplaceMX(ModeID,TempletContent)
			Dim MX_Arr,K
			MX_Arr=ACTCMS.Act_MX_Arr(ModeID,1)
			If IsArray(MX_Arr) Then
			  For K=0 To Ubound(MX_Arr,2)
				 If Not IsNull(GetNodeText("" &LCase(MX_Arr(0,K)) & "")) Then
				  TempletContent = Replace(TempletContent,"{$" & MX_Arr(0,K) & "}",GetNodeText("" &LCase(MX_Arr(0,K)) & ""))
				 Else
				  TempletContent = Replace(TempletContent,"{$" & MX_Arr(0,K) & "}","")
				 End If
			  Next
			End If
			ReplaceMX=TempletContent
		End Function

		Function ReplaceKeyTags(ModeID,KeyStr)
		  'on error resume next 
		  If Trim(KeyStr)="" Then Exit Function
		  Dim I,ActArr:ActArr=Split(KeyStr,",")
		  For I=0 To Ubound(ActArr)
		    ReplaceKeyTags=ReplaceKeyTags & "<a href=""" & Domain & "plus/search/index.asp?searchtype=3&ModeID=" & ModeID & "&keyword=" &server.URLEncode(ActArr(i))  & """ target=""_blank"">" & ActArr(i) & "</a> "
		  Next
		End Function 
		'上一篇、下一篇
		Function NextArticle(NowID, classID, TypeStr,ModeID)
			Dim SqlStr
			If Trim(TypeStr) = "Prev" Then
				   SqlStr = " SELECT Top 1 ClassID,ID,ActLink,FileName,infopurview,readpoint,title From  "&ACTCMS.ACT_C(ModeID,2)&"  Where classID='" & Trim(classID) & "' And ID<" & NowID & "  And isAccept=0 AND delif=0  Order By ID Desc"
			ElseIf Trim(TypeStr) = "Next" Then
				   SqlStr = " SELECT Top 1 ClassID,ID,ActLink,FileName,infopurview,readpoint,title From  "&ACTCMS.ACT_C(ModeID,2)&"  Where classID='" & Trim(classID) & "' And ID>" & NowID & "  And isAccept=0 AND delif=0  Order By ID"
			Else
				NextArticle = "":Exit Function
			End If
			 Dim RS:Set RS=Server.CreateObject("ADODB.Recordset")
			 RS.Open SqlStr, Conn, 1, 1
			 If RS.EOF And RS.BOF Then
				NextArticle = "没有了"
			 Else
				NextArticle = "<a href=""" &ACTCMS.GetInfoUrl(ModeID,Rs(0),Rs(1),Rs(2),Rs(3),Rs(4),Rs(5)) & """ title=""" & ACTCMS.CloseHtml(RS("title")) & """>" & RS("title") & "</a>"
			 End If
			 RS.Close:Set RS = Nothing
		End Function

		Function CreateArticleList(ModeID,FolderRs)
			Dim TemplateContent,FilePath,IndexHtml,FolderDir
			Application(AcTCMSN & "ACTCMS_TCJ_Type")="Folder"
			Application(AcTCMSN & "modeid")=FolderRs("modeid")
			Application(AcTCMSN & "classid")=FolderRs("classid")
			If Trim(FolderRs("ParentID")) = "0" Then Application(AcTCMSN & "ModeHome")= True	Else Application(AcTCMSN & "ModeHome") = False
			TemplateContent = LoadTemplate(FolderRs("FolderTemplate"))'模版
 			If TemplateContent = "" Then TemplateContent ="模板文件丢失"
		    TemplateContent=Loadfile(TemplateContent)
			TemplateContent = AllLabel(TemplateContent)'标签转换
			TemplateContent = LableFlag(GeneralLabel(TemplateContent))'通用标签转换
			TemplateContent =ReplaceAllLabel(TemplateContent)
			IndexHtml = FolderRs("Extension")
 		 	If Trim(FolderRs("actlink"))="3" Then 
				  TemplateContent=Replace(TemplateContent,"{$GetClassIntro}",FolderRs("content"))
   			   	  Call FSOSaveFile(TemplateContent,actcms.GetPath(FolderRs("classid"),FolderRs("makehtmlname")))
			Else 
				
				If  FolderRs("ParentID")="0" Then 
				    If actcms.ACT_L(actcms.GetParent(FolderRs("classid")),13)="1" Then 
 						FilePath = asys &actcms.ACT_L(FolderRs("classid"),14)
					Else 
						FilePath = asys &actcms.ACT_L(FolderRs("classid"),14)& Actcms.ACT_C(ModeID,6)& FolderRs("ClassEName")
					End If 
				Else 
				    If actcms.ACT_L(actcms.GetParent(FolderRs("classid")),13)="1" Then 
						FilePath = asys &actcms.ACT_L(actcms.GetParent(FolderRs("classid")),14)&"/"&  FolderRs("ClassEName")
					Else 
						FilePath = asys & Actcms.ACT_C(ModeID,6)& FolderRs("ClassEName")
					End If 
 					
				End If 
 				Call Actcms.CreateFolder(FilePath)

 
				If (Application(Cstr(AcTCMSN & "PageList")) <> "")   Then
				  Call GetPageStr(Application(Cstr(AcTCMSN & "PageList")), IndexHtml, TemplateContent,FilePath,  True)
				  Application.Contents.Remove (AcTCMSN & "PageList")
				Else
				  TemplateContent = Replace(TemplateContent,"{$pagelist}"," 首页 上一页 <strong>[1]</strong>下一页 尾页 &nbsp;<span>1/1页</span>")
				  TemplateContent=Replace(TemplateContent,"{$pagecount}",0)
				  TemplateContent=Replace(TemplateContent,"{$pagethis}",0)
				  TemplateContent=Replace(TemplateContent,"{$pagenum}",0)
				  TemplateContent = Replace(TemplateContent, "{PageListStr}", "")
			     Call FSOSaveFile(TemplateContent,FilePath & IndexHtml)
				End If 
 			 End If
		End Function
		Sub GetPageStr(PageContent, Index, FileContent,FilePath,  TypeSelect)
			Dim CurrPage, PageStr, TempFileContent, I, PageContentArr, J, SelectStr
			Dim TotalPage
			Dim HomeLink     
			Dim LinkUrlFileName 
			Dim FileName       
			Dim FExt         
			  HomeLink = Index
			  FExt = MID(Trim(Index), InStrRev(Trim(Index), ".")) 
			  FileName = Replace(Trim(Index), FExt, "") 
			  LinkUrlFileName = FileName
			  PageContentArr = Split(PageContent, "{$PageList}")
 			  TotalPage = UBound(PageContentArr)
			  For I = LBound(PageContentArr) To TotalPage - 1
			   CurrPage = I + 1
 			  If Application(AcTCMSN & "PageStyle")=4 Then 
  			    If CurrPage=1 Then
			     PageStr="首页 上一页"
				ElseIf CurrPage=2 Then
			     PageStr="<a href=""" & HomeLink & """ title=""首页"">首页</a> <a href=""" & HomeLink & """ title=""上一页"">上一页</a>"& vbcrlf
				Else
				 PageStr="<a href=""" & HomeLink & """ title=""首页"">首页</a> <a href=""" & LinkUrlFileName & "_" & CurrPage - 1 & FExt & """ title=""上一页"">上一页</a> "& vbcrlf
				End If
				 For J=CurrPage To CurrPage+9
				    If J>TotalPage Then Exit For
				    If J= CurrPage Then
				     PageStr=PageStr & " <strong>[" & J &"]</strong>"& vbcrlf
				    Else
				     PageStr=PageStr & " <a href=""" & LinkUrlFileName & "_" & J & FExt & """>[" & J &"]</a>"& vbcrlf
					End If
				 Next
				 If CurrPage=TotalPage Then
				  PageStr=PageStr & " 下一页 尾页"
				 Else
				  PageStr=PageStr & " <a href=""" & LinkUrlFileName & "_" & CurrPage + 1 & FExt & """ title=""下一页"">下一页</a> <a href=""" & LinkUrlFileName & "_" & TotalPage & FExt & """>尾页</a> "& vbcrlf
				 End If
 			  Else 
 			   Select Case Application(AcTCMSN & "PageStyle")
			    Case 1   
				   If CurrPage = 1 And CurrPage <> TotalPage Then
					PageStr = "首页  上一页 <a href=""" & LinkUrlFileName & "_" & CurrPage + 1 & FExt & """>下一页</a>  <a href= """ & LinkUrlFileName & "_" & TotalPage & FExt & """>尾页</a>"
				   ElseIf CurrPage = 1 And CurrPage = TotalPage Then
					PageStr = "首页  上一页 下一页 尾页"
				   ElseIf CurrPage = TotalPage And CurrPage <> 2 Then
					 PageStr = "<a href=""" & HomeLink & """>首页</a>  <a href=""" & LinkUrlFileName & "_" & CurrPage - 1 & FExt & """>上一页</a> 下一页  尾页"
				   ElseIf CurrPage = TotalPage And CurrPage = 2 Then
					 PageStr = "<a href=""" & HomeLink & """>首页</a>  <a href=""" & HomeLink & """>上一页</a> 下一页  尾页"
				   ElseIf CurrPage = 2 Then
					PageStr = "<a href=""" & HomeLink & """>首页</a>  <a href=""" & HomeLink & """>上一页</a> <a href=""" & LinkUrlFileName & "_" & CurrPage + 1 & FExt & """>下一页</a>  <a href= """ & LinkUrlFileName & "_" & (TotalPage & FExt) & """>尾页</a>"
				   Else
					PageStr = "<a href=""" & HomeLink & """>首页</a>  <a href=""" & LinkUrlFileName & "_" & CurrPage - 1 & FExt & """>上一页</a> <a href=""" & LinkUrlFileName & "_" & CurrPage + 1 & FExt & """>下一页</a>  <a href= """ & LinkUrlFileName & "_" & (TotalPage & FExt) & """>尾页</a>"
				   End If
			  Case 2
			    If CurrPage=1 Then
			     PageStr="<font face=webdings>9</font> <font face=webdings>7</font>"
				ElseIf CurrPage=2 Then
			     PageStr="<a href=""" & HomeLink & """ title=""首页""><font face=webdings>9</font></a> <a href=""" & HomeLink & """ title=""上一页""><font face=webdings>7</font></a>"
				Else
				 PageStr="<a href=""" & HomeLink & """ title=""首页""><font face=webdings>9</font></a> <a href=""" & LinkUrlFileName & "_" & CurrPage - 1 & FExt & """ title=""上一页""><font face=webdings>7</font></a> "
				End If
				 For J=CurrPage To CurrPage+9
				    If J>TotalPage Then Exit For
				    If J= CurrPage Then
				     PageStr=PageStr & " <font color=red>[" & J &"]</font>"
				    Else
				     PageStr=PageStr & " <a href=""" & LinkUrlFileName & "_" & J & FExt & """>[" & J &"]</a>"
					End If
				 Next
				 If CurrPage=TotalPage Then
				  PageStr=PageStr & " <font face=webdings>8</font> <font face=webdings>:</font>"
				 Else
				  PageStr=PageStr & " <a href=""" & LinkUrlFileName & "_" & CurrPage + 1 & FExt & """ title=""上一页""><font face=webdings>8</font></a> <a href=""" & LinkUrlFileName & "_" & TotalPage & FExt & """><font face=webdings>:</font></a> "
				 End If
			  Case 3
			    If CurrPage=1 Then
			     PageStr="<font face=webdings>9</font> <font face=webdings>7</font>"
				ElseIf CurrPage=2 Then
			     PageStr="<a href=""" & HomeLink & """ title=""首页""><font face=webdings>9</font></a> <a href=""" & HomeLink & """ title=""上一页""><font face=webdings>7</font></a>"
				Else
				 PageStr="<a href=""" & HomeLink & """ title=""首页""><font face=webdings>9</font></a> <a href=""" & LinkUrlFileName & "_" & CurrPage - 1 & FExt & """ title=""上一页""><font face=webdings>7</font></a> "
				End If			
				 If CurrPage=TotalPage Then
				  PageStr=PageStr & " <font face=webdings>8</font> <font face=webdings>:</font>"
				 Else
				  PageStr=PageStr & " <a href=""" & LinkUrlFileName & "_" & CurrPage + 1 & FExt & """ title=""上一页""><font face=webdings>8</font></a> <a href=""" & LinkUrlFileName & "_" & TotalPage & FExt & """><font face=webdings>:</font></a> "
				 End If
			  End Select	   
			 End If 

 
  			If Application(AcTCMSN & "PageStyle")=4 Then 
			   TempFileContent = Replace(FileContent, "{PageListStr}", PageContentArr(I))
 			   TempFileContent = Replace(TempFileContent,"{$pagelist}",   PageStr)
 			  TempFileContent=Replace(TempFileContent,"{$pagecount}",Application(AcTCMSN & "pagecount"))
			 ' Application(AcTCMSN & "pagecount")=""
			  TempFileContent=Replace(TempFileContent,"{$pagethis}",CurrPage)
   			  TempFileContent=Replace(TempFileContent,"{$pagenum}",TotalPage)
  			Else
			   TempFileContent = Replace(FileContent, "{PageListStr}", PageContentArr(I) & PageStr & "</div></div>")
			End If 
			   Dim TempFilePath
			   If CurrPage = 1 Then
				  TempFilePath =FilePath&Index
			   Else
				 TempFilePath = FilePath&FileName & "_" & CurrPage & FExt
			   End If
			  Call FSOSaveFile( TempFileContent, TempFilePath)
			  Next
		End Sub

		Function ACT_A_List(ClassID,ActF,ATT,ArticleSort,OpenTypeStr,ListNumber,RowHeight,TitleLen,ColNumber,TypeClassName,TypeNew,ACTIF,NavType,Nav,MoreLinkType,MoreLink,Division,DateForm,DateAlign,TitleCss,DateCss,DiyContent,ModeID,SubClass,ContentLen) 
			Dim SqlStr, Parameter,OpenType,MoreLinkStr,ACT_IF,ACTCMS_ATT
			Select Case ClassID 
			    Case "","0":Parameter=""
				Case "1"
					If Application(AcTCMSN & "classid")<>"0"  Then 
						If  CBool(SubClass)=True Then 
							 Parameter="ClassID In (" & ACTCMS.TempClassID(Application(AcTCMSN & "classid")) & ") And"
						Else 
							 Parameter="ClassID='" & Application(AcTCMSN & "classid") & "' And" 
							 ClassID=Application(AcTCMSN & "classid")
						End If 
					End If 
				Case Else
					If InStr(ClassID, ",") > 0 Then
						 Parameter="ClassID In (" & ClassID & ") And"
					Else
						If CBool(SubClass)=True Then 
						 Parameter="ClassID In (" & ACTCMS.TempClassID(ClassID) & ") And"
						Else 
						 Parameter="ClassID='" & Replace(ClassID,"'","") & "' And"
						End If 
					End If 
			End Select
 			If ModeID="0" Then ModeID=Cint(Application(AcTCMSN & "modeid"))
			If ACTIF<>"" Then ACT_IF = "  "&ACTIF
			If ATT="0" Then  ACTCMS_ATT="" Else ACTCMS_ATT = " And ATT="&ATT
			If MoreLink <> "" And InStr(ClassID, ",") = 0 And ClassID <> "0"  And ClassID <> "1" Then
				If ActF=1 Then 
					MoreLinkStr=MLink(ColNumber,RowHeight,MoreLinkType, MoreLink, AcTCMS.DiyClassName(ClassID),OpenTypeStr)
				Else 
					MoreLinkStr=CodeMLink(MoreLinkType, MoreLink, AcTCMS.DiyClassName(ClassID),OpenTypeStr)
				End If 
			End If 
			If Ucase(Left(Trim(ArticleSort),2))<>"id" Then  ArticleSort=ArticleSort & ",ID Desc"
			Sqlstr="Select TOP " & ListNumber & " ID,ClassID,Title,UpdateTime,ActLink,FileName,infopurview,readpoint,PicUrl,Intro,Content,CopyFrom,Author,KeyWords"&ACTCMS.DIYField(ModeID)&" From "&ACTCMS.ACT_C(ModeID,2)&" Where " & Parameter & " isAccept=0 AND delif=0 " & ACTCMS_ATT &ACT_IF&UIDSQL& " ORDER BY IsTop Desc," & ArticleSort
 			If ActF=2 Then 
				ACT_A_List = ACTCMS_A_Code(SqlStr,TitleLen,MoreLinkStr,DateForm,DiyContent,ModeID,ContentLen) 
 			Else 
				ACT_A_List = ACTCMS_A_SQL(SqlStr,OpenTypeStr,RowHeight,TitleLen,ColNumber,TypeClassname,TypeNew,NavType,Nav,MoreLinkStr,Division,DateForm,DateAlign,TitleCss,DateCss,ModeID) 
			End If 
 		End Function

 		Function ACTCMS_A_Code(SqlStr,TitleLen,MoreLinkStr,DateForm,DiyContent,ModeID,ContentLen) 
			 Dim RS,K,N,TempTitle,ACTSQL,DiyContents,ModID,J
			 Set RS=ACTCMS.ActExe(SqlStr)
 			 If RS.EOF Then	 ACTCMS_A_Code="":RS.Close:Set RS=Nothing:Exit Function
			 ACTSQL=RS.GetRows(-1):Set RS = Nothing
 			 Dim ActNum:ActNum=Ubound(ACTSQL,2)
			 Dim DIYFieldText
			 DIYFieldText=ACTCMS.DIYField(ModeID)
 			 J=1
 				For K=0 To ActNum'17
					    DiyContents=DiyContent
 						If InStr(DiyContents, "#ID") > 0  Then
						   DiyContents = Replace(DiyContents, "#ID",ACTSQL(0,N))
						End If	
						If InStr(DiyContents, "#Link") > 0  Then
						   DiyContents = Replace(DiyContents, "#Link", AcTCMS.GetInfoUrl(ModeID,ACTSQL(1,N),ACTSQL(0,N),ACTSQL(4,N),ACTSQL(5,N),ACTSQL(6,N),ACTSQL(7,N)))
						End if
						If InStr(DiyContents, "#Title") > 0  Then
						   DiyContents = Replace(DiyContents, "#Title",ACTCMS.GetStrValue(ACTSQL(2,N),TitleLen) )
						End if
 						If InStr(DiyContents, "#CTitle") > 0  Then
						   DiyContents = Replace(DiyContents, "#CTitle",AcTCMS.CloseHtml(ACTSQL(2,N)))
						End if
 					 	If InStr(DiyContents, "#KeyWord") > 0  Then
							If ACTSQL(13,N)<>"" Then 
						    DiyContents = Replace(DiyContents, "#KeyWord","<a href=""" & Domain & "plus/search/index.asp?searchtype=3&ModeID=" & ModeID & "&keyword=" & ACTSQL(13,N)& """ target=""_blank"">" & ACTSQL(13,N) & "</a>")
							Else
						    DiyContents = Replace(DiyContents, "#KeyWord","")
							End If 
					 	End If
      					If InStr(DiyContents, "#Thumb") > 0  Then
						   If  Trim(AcTSQL(8,N))<>"" Then 
							   DiyContents = Replace(DiyContents, "#Thumb",ACTCMS.PathDoMain&ACTSQL(8,N))
						   Else
							   DiyContents = Replace(DiyContents, "#Thumb",ACTCMS.ActCMSDM&"images/nopic.gif")
						   End If 
						End if
   						If InStr(DiyContents, "#PicUrl") > 0  Then
						   If  Trim(AcTSQL(8,N))<>"" Then 
						   DiyContents = Replace(DiyContents, "#PicUrl",ACTCMS.PathDoMain&Replace(ACTSQL(8,N),"thumb_",""))
						   Else
							   DiyContents = Replace(DiyContents, "#PicUrl",ACTCMS.ActCMSDM&"images/nopic.gif")
						   End If 
 						End If
 						If InStr(DiyContents, "#Intro") > 0  Then'暂用
							If Trim(ACTSQL(9,N))<>"" Then 
							   DiyContents = Replace(DiyContents, "#Intro",ACTCMS.GetStrValue(Replace(Replace(Replace(AcTCMS.CloseHtml(ACTSQL(9,N)), vbCrLf, ""), "[NextPage]", ""), "&nbsp;", ""), ContentLen))
							Else
							   DiyContents = Replace(DiyContents, "#Intro",ACTCMS.GetStrValue(Replace(Replace(Replace(AcTCMS.CloseHtml(ACTSQL(10,N)), vbCrLf, ""), "[NextPage]", ""), "&nbsp;", ""), ContentLen))
							End If 
						End if
  						If InStr(DiyContents, "#ClassName") > 0  Then
						   DiyContents = Replace(DiyContents, "#ClassName",ACTCMS.ACT_L(ACTSQL(1,N),2))
						End if
  						If InStr(DiyContents, "#ClassLink") > 0  Then
						   DiyContents = Replace(DiyContents, "#ClassLink",AcTCMS.DiyClassName(ACTSQL(1,N)))
						End If
  						If InStr(DiyContents, "#Time") > 0  Then
						   DiyContents = Replace(DiyContents, "#Time",CodeDateStr(ACTSQL(3,N),DateForm))
						End if
 						If InStr(DiyContents, "#Hits") > 0  Then
						   DiyContents = Replace(DiyContents, "#Hits","<Script Language=""Javascript"" Src=""" & Domain & "Plus/ACT.Hits.asp?A=List&ModeID="&ModeID&"&ID=" & ACTSQL(0,N)  & """></Script>")
						End If
  						If InStr(DiyContents, "#CopyFrom") > 0   Then
						   If   Trim(ACTSQL(11,N)) <> "" Then 
							 DiyContents = Replace(DiyContents, "#CopyFrom",ACTSQL(11,N))
						   Else
							 DiyContents = Replace(DiyContents, "#CopyFrom","")
						   End If 
 						End If
 						If InStr(DiyContents, "#Author") > 0   Then
						   If   Trim(ACTSQL(12,N)) <> "" Then 
							 DiyContents = Replace(DiyContents, "#Author",ACTSQL(12,N))
						   Else
							 DiyContents = Replace(DiyContents, "#Author","佚名")
						   End If 
 						End If
 						If InStr(DiyContents, "#AutoID") > 0  Then
						   DiyContents = Replace(DiyContents, "#AutoID",J)
						End if
 						If InStr(DiyContents, "#ModID") > 0  Then
 							If  N Mod 2 =0 Then ModID=0 Else ModID=1
						    DiyContents = Replace(DiyContents, "#ModID",ModID)
						End If
 						If InStr(DiyContents, "#Path") > 0  Then
						   DiyContents = Replace(DiyContents, "#Path",ACTCMS.ActCMSDM)
						End if
 						If InStr(DiyContents, "#New") > 0  Then
						   If  (Year(ACTSQL(3,N))&Month(ACTSQL(3,N))&Day(ACTSQL(3,N)) =Year(Now)&Month(Now)&Day(Now)) Then
							   DiyContents = Replace(DiyContents, "#New","<img src=""" & ACTCMS.ActCMSDM&"ACT_inc/share/new.gif"" border=""0""/>")
						   Else 
							   DiyContents = Replace(DiyContents, "#New","")
						   End If 
 						End If
  						If InStr(DiyContents, "#ClassSeo") > 0  Then
						   DiyContents = Replace(DiyContents, "#ClassSeo",ACTCMS.ACT_L(ACTSQL(1,N),25))
						End if
  						If InStr(DiyContents, "#ClassPicUrl") > 0  Then
						   DiyContents = Replace(DiyContents, "#ClassPicUrl",ACTCMS.ACT_L(ACTSQL(1,N),26))
						End if
  						If InStr(DiyContents, "#ClassPicFile") > 0  Then
							If ACTCMS.ACT_L(ACTSQL(1,N),26)<>"" Then 
							  DiyContents = Replace(DiyContents, "#ClassPicFile","<img src="""&ACTCMS.ACT_L(ACTSQL(1,N),26)&""" alt="""&ACTCMS.ACT_L(ACTSQL(1,N),2)&"""  /> ")
							Else 
							  DiyContents = Replace(DiyContents, "#ClassPicFile","")
							End If 
						End if

   						 
 					    ACTCMS_A_Code = ACTCMS_A_Code & DiyContents& vbCrLf
						N=N+1:j=j+1
 			    Next
				        ACTCMS_A_Code = ACTCMS_A_Code&MoreLinkStr& vbCrLf
 		End Function 

 
		Function ACTCMS_A_SQL(SqlStr,OpenType,RowHeight,TitleLen,ColNumber,TypeClassname,TypeNew,NavType,Nav,MoreLinkStr,Division,DateForm,DateAlign,TitleCss,DateCss,ModeID) 
			 'on error resume next
			 Dim RS,I,k,N,ColSpanNum,TypeNews,TempTitle,NaviStr,ACTSQL,ClassnameLink
			 Set RS=ACTCMS.ActExe(SqlStr)
 			 If RS.EOF Then	 ACTCMS_A_SQL="":RS.Close:Set RS=Nothing:Exit Function
			 ACTSQL=RS.GetRows(-1):Set RS = Nothing
			 Dim ActNum:ActNum=Ubound(ACTSQL,2)
  				 Dim TitleCssName,DateCssStr,DateStr
				 TitleCssName = GCss(TitleCss):DateCssStr = GCss(DateCss):RowHeight = GRowHeight(RowHeight):NaviStr = GNavi(NavType,Nav)
				 ACTCMS_A_SQL = "<table border=""0"" cellpadding=""0"" cellspacing=""0"" wIDth=""100%"">" & vbCrLf
				 For K=0 To ActNum
					 ACTCMS_A_SQL = ACTCMS_A_SQL & "<tr>" & vbCrLf
					 For I = 1 To ColNumber
					  If CBool(TypeClassname) = True Then ClassnameLink = "[" & AcTCMS.GainClassName(ACTSQL(1,N),OpenType,TitleCssName) & "]"			
					  If Cbool(TypeNew)=True And (Year(ACTSQL(3,N))&Month(ACTSQL(3,N))&Day(ACTSQL(3,N)) =Year(Now)&Month(Now)&Day(Now)) Then TypeNews="<img src=""" & Domain&"ACT_inc/share/new.gif"" border=""0""/>" Else TypeNews=""
					  DateStr=GDateStr(ACTSQL(3,N),DateForm,DateAlign,DateCssStr,ColNumber,ColSpanNum)
					  TempTitle = "<a " & TitleCssName &  " href=""" &AcTCMS.GetInfoUrl(ModeID,ACTSQL(1,N),ACTSQL(0,N),ACTSQL(4,N),ACTSQL(5,N),ACTSQL(6,N),ACTSQL(7,N)) &  """"  & Gopen(OpenType) & " title=""" & AcTCMS.CloseHtml(ACTSQL(2,N)) & """>" &ACTCMS.GetStrValue(ACTSQL(2,N),TitleLen) & "</a>" 
						  If ColNumber=1 Then
							  ACTCMS_A_SQL = ACTCMS_A_SQL & ("  <td height=""" & RowHeight & """>"  &NaviStr&ClassnameLink&TempTitle&TypeNews&DateStr& "</td>" & vbCrLf)
						  Else
							  ACTCMS_A_SQL = ACTCMS_A_SQL & ("  <td  wIDth=""" & CInt(100 / CInt(ColNumber)) & "%"" height=""" &RowHeight&  """>" & vbCrLf)
							  ACTCMS_A_SQL = ACTCMS_A_SQL & ("    <table wIDth=""90%"" height=""100%"" cellpadding=""0"" cellspacing=""0"" border=""0"">" & vbCrLf)
							  ACTCMS_A_SQL = ACTCMS_A_SQL & ("     <tr><td> " &NaviStr&ClassnameLink&TempTitle&TypeNews &DateStr )
							  ACTCMS_A_SQL = ACTCMS_A_SQL & ("      </td></tr>" & vbcrlf &"   </table>" & vbCrLf & "  </td>" & vbCrLf)
						  End if
						  N=N+1
					      If N>=ActNum+1 Then Exit For
					 Next
					 ACTCMS_A_SQL = ACTCMS_A_SQL & "</tr>" & vbCrLf
					 ACTCMS_A_SQL = ACTCMS_A_SQL & (GbgPic(Division,ColSpanNum) & vbCrLf)
					 If N>=ActNum+1 Then Exit For
				Next
					 ACTCMS_A_SQL = ACTCMS_A_SQL & MoreLinkStr& ("</table>" & vbCrLf)
 		End Function

 		Function ACT_P(ClassID,ActF,ATT,ArticleSort,OpenType,ListNumber,ColNumber,TitleLen,Titlecss,PiCcss,PicWIDth,PicHeight,ContentLen,PicStyle,TypeTitle,ACTIF,DiyContent,ModeID,SubClass)     
			Dim SqlStr, ACT_IF,Parameter,ACTCMS_ATT
			Select Case ClassID 
			    Case "","0":Parameter=""
				Case "1"
					If Application(AcTCMSN & "classid")<>"0"  Then 
						If  CBool(SubClass)=True Then 
							 Parameter="ClassID In (" & ACTCMS.TempClassID(Application(AcTCMSN & "classid")) & ") And"
						Else 
							Parameter="ClassID='" & Application(AcTCMSN & "classid") & "' And" 
							ClassID=Application(AcTCMSN & "classid")
						End If 
					End If 
				Case Else
					If InStr(ClassID, ",") > 0 Then
						 Parameter="ClassID In (" & ClassID & ") And"
					Else
						If CBool(SubClass)=True Then 
						 Parameter="ClassID In (" & ACTCMS.TempClassID(ClassID) & ") And"
						Else 
						 Parameter="ClassID='" & Replace(ClassID,"'","") & "' And"
						End If 
					End If 
			End Select
			If ModeID="0" Then ModeID=Cint(Application(AcTCMSN & "modeid"))
			If ACTIF<>"" Then ACT_IF = "  "&ACTIF
			If ATT="0" Then  ACTCMS_ATT="" Else ACTCMS_ATT = " And ATT="&ATT
			OpenType = Gopen(OpenType)
			If Ucase(Left(Trim(ArticleSort),2))<>"id" Then  ArticleSort=ArticleSort & ",ID Desc"
			Sqlstr="Select TOP " & ListNumber & " ID,ClassID,Title,UpdateTime,ActLink,FileName,infopurview,readpoint,PicUrl,Intro,Content,CopyFrom,Author,KeyWords"&ACTCMS.DIYField(ModeID)&" From  "&ACTCMS.ACT_C(ModeID,2)&"  Where " & Parameter & " isAccept=0 AND delif=0 AND PicUrl<>'' " & ACTCMS_ATT &ACT_IF&UIDSQL& " ORDER BY IsTop Desc," & ArticleSort
			If ActF=2 Then 
				ACT_P =ACT_P_Code(SqlStr,ContentLen,TitleLen,DiyContent,ModeID)
			Else 
				ACT_P = ACT_P_SQL(SqlStr,OpenType,ColNumber,TitleLen,Titlecss,PiCcss,PicWIDth,PicHeight,ContentLen,PicStyle,TypeTitle,ModeID)
			End If 
		End Function

		Function ACT_P_Code(SqlStr,ContentLen,TitleLen,DiyContent,ModeID)
			 Dim K,N,AcTSQL,RS,ArticleC,DiyContents,j,ModID
			 Set RS=ACTCMS.ActExe(SqlStr)
 			 If RS.EOF Then	 ACT_P_Code="":RS.Close:Set RS=Nothing:Exit Function
			 AcTSQL=RS.GetRows(-1):RS.Close:Set RS = Nothing
			 Dim ActNum:ActNum=Ubound(AcTSQL,2)
			 Dim DIYFieldText
			 DIYFieldText=ACTCMS.DIYField(ModeID)
					J=1	
  					For K=0 To ActNum
  						 DiyContents=DiyContent
 						If InStr(DiyContents, "#ID") > 0  Then
						   DiyContents = Replace(DiyContents, "#ID",ACTSQL(0,N))
						End If	
						If InStr(DiyContents, "#Link") > 0  Then
						   DiyContents = Replace(DiyContents, "#Link", AcTCMS.GetInfoUrl(ModeID,ACTSQL(1,N),ACTSQL(0,N),ACTSQL(4,N),ACTSQL(5,N),ACTSQL(6,N),ACTSQL(7,N)))
						End if
						If InStr(DiyContents, "#Title") > 0  Then
						   DiyContents = Replace(DiyContents, "#Title",ACTCMS.GetStrValue(ACTSQL(2,N),TitleLen) )
						End if
 						If InStr(DiyContents, "#CTitle") > 0  Then
						   DiyContents = Replace(DiyContents, "#CTitle",AcTCMS.CloseHtml(ACTSQL(2,N)))
						End if
 					 	If InStr(DiyContents, "#KeyWord") > 0  Then
							If ACTSQL(13,N)<>"" Then 
						    DiyContents = Replace(DiyContents, "#KeyWord","<a href=""" & Domain & "plus/search/index.asp?searchtype=3&ModeID=" & ModeID & "&keyword=" & ACTSQL(13,N)& """ target=""_blank"">" & ACTSQL(13,N) & "</a>")
							Else
						    DiyContents = Replace(DiyContents, "#KeyWord","")
							End If 
					 	End If
    					If InStr(DiyContents, "#Thumb") > 0  Then
						   If  Trim(AcTSQL(8,N))<>"" Then 
							   DiyContents = Replace(DiyContents, "#Thumb",ACTCMS.PathDoMain&ACTSQL(8,N))
						   Else
							   DiyContents = Replace(DiyContents, "#Thumb",ACTCMS.ActCMSDM&"images/nopic.gif")
						   End If 
						End if
   						If InStr(DiyContents, "#PicUrl") > 0  Then
						   If  Trim(AcTSQL(8,N))<>"" Then 
							  DiyContents = Replace(DiyContents, "#PicUrl",ACTCMS.PathDoMain&Replace(ACTSQL(8,N),"thumb_",""))
						   Else
							  DiyContents = Replace(DiyContents, "#PicUrl",ACTCMS.ActCMSDM&"images/nopic.gif")
						   End If 
 						End If
 						If InStr(DiyContents, "#Intro") > 0  Then'暂用
							If Trim(ACTSQL(9,N))<>"" Then 
							   DiyContents = Replace(DiyContents, "#Intro",ACTCMS.GetStrValue(Replace(Replace(Replace(AcTCMS.CloseHtml(ACTSQL(9,N)), vbCrLf, ""), "[NextPage]", ""), "&nbsp;", ""), ContentLen))
							Else
							   DiyContents = Replace(DiyContents, "#Intro",ACTCMS.GetStrValue(Replace(Replace(Replace(AcTCMS.CloseHtml(ACTSQL(10,N)), vbCrLf, ""), "[NextPage]", ""), "&nbsp;", ""), ContentLen))
							End If 
						End if
  						If InStr(DiyContents, "#ClassName") > 0  Then
						   DiyContents = Replace(DiyContents, "#ClassName",ACTCMS.ACT_L(ACTSQL(1,N),2))
						End if
  						If InStr(DiyContents, "#ClassLink") > 0  Then
						   DiyContents = Replace(DiyContents, "#ClassLink",AcTCMS.DiyClassName(ACTSQL(1,N)))
						End If
  						If InStr(DiyContents, "#Time") > 0  Then
						   DiyContents = Replace(DiyContents, "#Time",CodeDateStr(ACTSQL(3,N),DateForm))
						End if
 						If InStr(DiyContents, "#Hits") > 0  Then
						   DiyContents = Replace(DiyContents, "#Hits","<Script Language=""Javascript"" Src=""" & Domain & "Plus/ACT.Hits.asp?A=List&ModeID="&ModeID&"&ID=" & ACTSQL(0,N)  & """></Script>")
						End If
  						If InStr(DiyContents, "#CopyFrom") > 0   Then
						   If  Trim(ACTSQL(11,N)) <> "" Then 
							 DiyContents = Replace(DiyContents, "#CopyFrom",ACTSQL(11,N))
						   Else
							 DiyContents = Replace(DiyContents, "#CopyFrom","")
						   End If 
 						End If
 						If InStr(DiyContents, "#Author") > 0   Then
						   If   Trim(ACTSQL(12,N)) <> "" Then 
							 DiyContents = Replace(DiyContents, "#Author",ACTSQL(12,N))
						   Else
							 DiyContents = Replace(DiyContents, "#Author","佚名")
						   End If 
 						End If
 						If InStr(DiyContents, "#AutoID") > 0  Then
						   DiyContents = Replace(DiyContents, "#AutoID",J)
						End if
 						If InStr(DiyContents, "#ModID") > 0  Then
 							If  N Mod 2 =0 Then ModID=0 Else ModID=1
						    DiyContents = Replace(DiyContents, "#ModID",ModID)
						End If
 						If InStr(DiyContents, "#Path") > 0  Then
						   DiyContents = Replace(DiyContents, "#Path",ACTCMS.ActCMSDM)
						End if

 						If InStr(DiyContents, "#New") > 0  Then
						   If  (Year(ACTSQL(3,N))&Month(ACTSQL(3,N))&Day(ACTSQL(3,N)) =Year(Now)&Month(Now)&Day(Now)) Then
							   DiyContents = Replace(DiyContents, "#New","<img src=""" & ACTCMS.ActCMSDM&"ACT_inc/share/new.gif"" border=""0""/>")
						   Else 
							   DiyContents = Replace(DiyContents, "#New","")
						   End If 
 						End if
  						If InStr(DiyContents, "#ClassSeo") > 0  Then
						   DiyContents = Replace(DiyContents, "#ClassSeo",ACTCMS.ACT_L(ACTSQL(1,N),25))
						End if
  						If InStr(DiyContents, "#ClassPicUrl") > 0  Then
						   DiyContents = Replace(DiyContents, "#ClassPicUrl",ACTCMS.ACT_L(ACTSQL(1,N),26))
						End if
  						If InStr(DiyContents, "#ClassPicFile") > 0  Then
							If ACTCMS.ACT_L(ACTSQL(1,N),26)<>"" Then 
							  DiyContents = Replace(DiyContents, "#ClassPicFile","<img src="""&ACTCMS.ACT_L(ACTSQL(1,N),26)&""" alt="""&ACTCMS.ACT_L(ACTSQL(1,N),2)&"""  /> ")
							Else 
							  DiyContents = Replace(DiyContents, "#ClassPicFile","")
							End If 
						End if
 
						ACT_P_Code =ACT_P_Code& DiyContents & vbCrLf
						N=N+1:j=j+1
  					Next 
 		End Function 

 		Function ACT_P_SQL(SqlStr,OpenType,ColNumber,TitleLen,Titlecss,PiCcss,PicWIDth,PicHeight,ContentLen,PicStyle,TypeTitle,ModeID)
			 'on error resume next
			 Dim PicStr,I,TempPicStr,ActCMSURL,K,N,TempTitle,AcTSQL,RS,ArticleC
			 Set RS=ACTCMS.ActExe(SqlStr)
 			 If  RS.EOF Then	 ACT_P_SQL="":RS.Close:Set RS=Nothing:Exit Function
			 AcTSQL=RS.GetRows(-1):RS.Close:Set RS = Nothing
			 Dim ActNum:ActNum=Ubound(AcTSQL,2)
					TitleCss = GCss(TitleCss):PicCss = GCss(PicCss)
 					PicStr="<table border=""0"" cellpadding=""0"" cellspacing=""0"" wIDth=""100%"">"
					For K=0 To ActNum
						 PicStr = PicStr & "<tr>" & vbCrLf
						 For I = 1 To ColNumber
							ActCMSURL=AcTCMS.GetInfoUrl(ModeID,ACTSQL(1,N),ACTSQL(0,N),ACTSQL(4,N),ACTSQL(5,N),ACTSQL(6,N),ACTSQL(7,N))
							TempPicStr = "<a href=""" &ActCMSURL & """" & OpenType & " title=""" & AcTCMS.CloseHtml(AcTSQL(2,N)) & """><Img "& PicCss &" Src=""" & AcTSQL(8,N) & """ border=""0"" wIDth=""" & PicWIDth & """ height=""" & PicHeight & """ align=""absmIDdle""/></a>"
							TempTitle = "<a " & TitleCss & " href=""" &ActCMSURL  & """" & OpenType & " title=""" & AcTCMS.CloseHtml(AcTSQL(2,N)) & """>" & ACTCMS.GetStrValue(ACTSQL(2,N),TitleLen) & "</a>"
							PicStr = PicStr & ("<td wIDth=""" & CInt(100 / CInt(ColNumber)) & "%"">" & vbCrLf)
							If AcTSQL(9,N)="" Or IsNull(AcTSQL(9,N)) Then ArticleC=AcTSQL(10,N) Else ArticleC=AcTSQL(10,N)
							Select Case PicStyle
								Case "1"
									 PicStr = PicStr & ("<span align=center><p>" & TempPicStr & "</p></span>" & vbCrLf)
								Case "2"
									 PicStr = PicStr & ("<table border=""0"" cellspacing=""0"" cellpadding=""0"" wIDth=""100%""> ")
									 PicStr = PicStr & ("<tr><td align=center>" & TempPicStr & "</td></tr>" & vbCrLf)
									 If CBool(TypeTitle) = True Then
									  PicStr = PicStr & ("<tr><td align=center>" & TempTitle  & "</td></tr>" & vbCrLf)
									 End If
									 PicStr = PicStr & ("</table>")
								Case "3"
									 PicStr = PicStr & "<table cellSpacing=""0"" cellPadding=""0"" wIDth=""100%"" border=""0"">" & vbCrLf
									 PicStr = PicStr & " <TR>" & vbCrLf
									 PicStr = PicStr & " <TD align=center>" & vbCrLf
									 PicStr = PicStr & "  <TABLE align=center cellSpacing=0 cellPadding=0 border=0>" & vbCrLf
									 PicStr = PicStr & "  <TBODY><TR><TD wIDth=110 align=center>" & TempPicStr & "</TD></TR></TBODY>" & vbCrLf
									 PicStr = PicStr & " </TABLE></TD>" & vbCrLf
									 PicStr = PicStr & "<TD> <TABLE wIDth=""100%"" border=""0"">" & vbCrLf
									 PicStr = PicStr & "<TBODY>"
									 If CBool(TypeTitle) = True Then
										PicStr = PicStr & "<TR><TD>" & TempTitle & "</TD></TR>" & vbCrLf
									 End If
									 PicStr = PicStr & "<TR><TD>"&ACTCMS.GetStrValue(Replace(Replace(Replace(AcTCMS.CloseHtml(ArticleC), vbCrLf, ""), "[NextPage]", ""), "&nbsp;", ""), ContentLen) & "...[<a href=""" & ActCMSURL & """" & OpenType & ">全文</a>]</TD></TR>" & vbCrLf
									 PicStr = PicStr & "</TBODY>" & vbCrLf
									 PicStr = PicStr & "</TABLE></TD>" & vbCrLf
									 PicStr = PicStr & " </TR>" & vbCrLf
									 PicStr = PicStr & "</TABLE>" & vbCrLf
								Case "4"
									 PicStr = PicStr & "<TABLE cellSpacing=""0"" cellPadding=""0"" wIDth=""100%"" border=""0"">" & vbCrLf
									 PicStr = PicStr & " <TBODY>" & vbCrLf
									 PicStr = PicStr & " <TR>" & vbCrLf
									 PicStr = PicStr & "<TD> <TABLE width=""100%"" border=""0"">" & vbCrLf
									 PicStr = PicStr & "<TBODY>"
									 If CBool(TypeTitle) = True Then
										PicStr = PicStr & "<TR><TD>" & TempTitle & "</TD></TR>" & vbCrLf
									 End If
									 PicStr = PicStr & "<TR><TD>"&ACTCMS.GetStrValue(Replace(Replace(Replace(AcTCMS.CloseHtml(ArticleC), vbCrLf, ""), "[NextPage]", ""), "&nbsp;", ""), ContentLen) & "...[<a href=""" & ActCMSURL & """" & OpenType & ">全文</a>]</TD></TR>" & vbCrLf
									 PicStr = PicStr & "</TBODY>" & vbCrLf
									 PicStr = PicStr & "</TABLE></TD>" & vbCrLf
									 PicStr = PicStr & " <TD align=center>" & vbCrLf
									 PicStr = PicStr & "<TABLE align=center cellSpacing=0 cellPadding=0 border=0>" & vbCrLf
									 PicStr = PicStr & "<TBODY><TR><TD wIDth=110 align=center>" & TempPicStr & "</TD></TR></TBODY>" & vbCrLf
									 PicStr = PicStr & "</TABLE></TD>" & vbCrLf
									 PicStr = PicStr & " </TR>" & vbCrLf
									 PicStr = PicStr & "</TBODY></TABLE>" & vbCrLf
								Case Else
									 PicStr=PicStr&PicStyle
								End Select
									 PicStr = PicStr & ("</td>" & vbCrLf)
									 N=N+1
									 If N>=ActNum+1 Then Exit For
							 Next
						 PicStr = PicStr & ("</tr>" & vbCrLf)
						 PicStr = PicStr & ("<tr><td colspan=""" & ColNumber & """ height=""5""></td></tr>")
						 IF N>=ActNum+1 Then Exit For
					Next
						 ACT_P_SQL = PicStr & ("</table>" & vbCrLf)
			 
		End Function

		Function GetSpecial(ListNumber,ContentLen,DiyContent,ArticleSort)
			 Dim Sqlstr,Rs,ACTSQL,i,DiyContents,j,FileName
			 If Ucase(Left(Trim(ArticleSort),2))<>"id" Then  ArticleSort=ArticleSort & ",ID Desc"
  			 Sqlstr= "Select TOP " & ListNumber & " ID,Title,PicIndex,writer,pubdate,Hits,Content,filename From Special_ACT Order By  "&ArticleSort
			 Set RS=ACTCMS.ActExe(SqlStr)
 			 If RS.EOF Then	 GetSpecial="":RS.Close:Set RS=Nothing:Exit Function
			 ACTSQL=RS.GetRows(-1):Set RS = Nothing
			 Dim ActNum:ActNum=Ubound(ACTSQL,2)
			 For i=0 To ActNum
 	 			   DiyContents=DiyContent
  						If InStr(DiyContents, "#ID") > 0  Then
							  DiyContents = Replace(DiyContents, "#ID",ACTSQL(0,i))
						End If	
  						If InStr(DiyContents, "#Title") > 0  Then
						  If  Trim(ACTSQL(1,i)) <> "" Then 
							 DiyContents = Replace(DiyContents, "#Title",ACTSQL(1,i))
						  Else 
							 DiyContents = Replace(DiyContents, "#Title","")
						  End If 
 						End If	
  						If InStr(DiyContents, "#Link") > 0  Then
  								If InStr(ACTSQL(7,i),"/")>0 Then 
  									If Right(ACTSQL(7,i),1)="/"   Then
										FileName=actcms.ActSys&ACTSQL(7,i)&"index.html"
									Else
										FileName=actcms.ActSys&ACTSQL(7,i)
									End If 
 								Else 
									FileName=actcms.ActSys&ACTSQL(7,i)
								End If 
 						   DiyContents = Replace(DiyContents, "#Link",FileName)
						End If	
  						If InStr(DiyContents, "#Thumb") > 0  Then
						  If  Trim(ACTSQL(2,i)) <> "" Then 
 							 DiyContents = Replace(DiyContents, "#Thumb",ACTSQL(2,i))
						  Else 
 							 DiyContents = Replace(DiyContents, "#Thumb","")
						  End If 
						End If	
  						If InStr(DiyContents, "#Writer") > 0  Then
						  If  Trim(ACTSQL(3,i)) <> "" Then 
							   DiyContents = Replace(DiyContents, "#Writer",ACTSQL(3,i))
						  Else 
							   DiyContents = Replace(DiyContents, "#Writer","")
						  End If 
						End If	
  						If InStr(DiyContents, "#Hits") > 0  Then
						   DiyContents = Replace(DiyContents, "#Hits",ACTSQL(5,i))
						End If	
  						If InStr(DiyContents, "#Time") > 0  Then
						   DiyContents = Replace(DiyContents, "#Time",ACTSQL(4,i))
						End If	
						If InStr(DiyContents, "#Content") > 0  Then
						   DiyContents = Replace(DiyContents, "#Content",ACTCMS.GetStrValue(ACTSQL(6,i), ContentLen))
						End If	
    				  GetSpecial = GetSpecial & DiyContents& vbCrLf
			 Next 
 
		End Function 

		Function ACTCMS_GetSlIDe(ClassID,ListNumber,TitleLen,ModeID,DiyContent,ContentLen)
			If ModeID=0 Then ModeID=Application(AcTCMSN & "modeid")
 			Dim SqlStr, Parameter
			Select Case ClassID 
			    Case "","0":Parameter=""
				Case "1"
					If Application(AcTCMSN & "classid")<>"0"  Then 
 							 Parameter="ClassID In (" & ACTCMS.TempClassID(Application(AcTCMSN & "classid")) & ") And"
 					End If 
				Case Else
					If InStr(ClassID, ",") > 0 Then
						 Parameter="ClassID In (" & ClassID & ") And"
					Else
 						 Parameter="ClassID In (" & ACTCMS.TempClassID(ClassID) & ") And"
 					End If 
			End Select
 			 Sqlstr= "Select TOP " & ListNumber & " ID,ClassID,Title,UpdateTime,ActLink,FileName,infopurview,readpoint,PicUrl,Intro,Content,CopyFrom,Author,KeyWords"&ACTCMS.DIYField(ModeID)&" From "&ACTCMS.ACT_C(ModeID,2)&" Where " & Parameter & " isAccept=0 AND delif=0 AND picurl<>'' AND SlIDe=1  ORDER BY IsTop Desc, ID Desc "
		     ACTCMS_GetSlIDe = ACT_SlIDeSQL(SqlStr,TitleLen,ModeID,DiyContent,ContentLen)
		End Function
		Function ACT_SlIDeSQL(SqlStr,TitleLen,ModeID,DiyContent,ContentLen)
   			 Dim RS,K,N,TempTitle,ACTSQL,DiyContents,ModID,J
			 Set RS=ACTCMS.ActExe(SqlStr)
 			 If RS.EOF Then	 ACT_SlIDeSQL="":RS.Close:Set RS=Nothing:Exit Function
			 ACTSQL=RS.GetRows(-1):Set RS = Nothing
			 Dim ActNum:ActNum=Ubound(ACTSQL,2)
			 Dim DIYFieldText
			 DIYFieldText=ACTCMS.DIYField(ModeID)
 			 J=1
 				For K=0 To ActNum
					   DiyContents=DiyContent
 						If InStr(DiyContents, "#ID") > 0  Then
						   DiyContents = Replace(DiyContents, "#ID",ACTSQL(0,N))
						End If	
						If InStr(DiyContents, "#Link") > 0  Then
						   DiyContents = Replace(DiyContents, "#Link", AcTCMS.GetInfoUrl(ModeID,ACTSQL(1,N),ACTSQL(0,N),ACTSQL(4,N),ACTSQL(5,N),ACTSQL(6,N),ACTSQL(7,N)))
						End if
						If InStr(DiyContents, "#Title") > 0  Then
						   DiyContents = Replace(DiyContents, "#Title",ACTCMS.GetStrValue(ACTSQL(2,N),TitleLen) )
						End if
 						If InStr(DiyContents, "#CTitle") > 0  Then
						   DiyContents = Replace(DiyContents, "#CTitle",AcTCMS.CloseHtml(ACTSQL(2,N)))
						End if
 					 	If InStr(DiyContents, "#KeyWord") > 0  Then
							If ACTSQL(13,N)<>"" Then 
						    DiyContents = Replace(DiyContents, "#KeyWord","<a href=""" & Domain & "plus/search/index.asp?searchtype=3&ModeID=" & ModeID & "&keyword=" & ACTSQL(13,N)& """ target=""_blank"">" & ACTSQL(13,N) & "</a>")
							Else
						    DiyContents = Replace(DiyContents, "#KeyWord","")
							End If 
					 	End If
    					If InStr(DiyContents, "#Thumb") > 0  Then
						   If  Trim(AcTSQL(8,N))<>"" Then 
							   DiyContents = Replace(DiyContents, "#Thumb",ACTCMS.PathDoMain&ACTSQL(8,N))
						   Else
							   DiyContents = Replace(DiyContents, "#Thumb",ACTCMS.ActCMSDM&"images/nopic.gif")
						   End If 
						End if
   						If InStr(DiyContents, "#PicUrl") > 0  Then
						   If  Trim(AcTSQL(8,N))<>"" Then 
						   DiyContents = Replace(DiyContents, "#PicUrl",ACTCMS.PathDoMain&Replace(ACTSQL(8,N),"thumb_",""))
						   Else
							   DiyContents = Replace(DiyContents, "#PicUrl",ACTCMS.ActCMSDM&"images/nopic.gif")
						   End If 
 						End If
 						If InStr(DiyContents, "#Intro") > 0  Then'暂用
							If Trim(ACTSQL(9,N))<>"" Then 
							   DiyContents = Replace(DiyContents, "#Intro",ACTCMS.GetStrValue(Replace(Replace(Replace(AcTCMS.CloseHtml(ACTSQL(9,N)), vbCrLf, ""), "[NextPage]", ""), "&nbsp;", ""), ContentLen))
							Else
							   DiyContents = Replace(DiyContents, "#Intro",ACTCMS.GetStrValue(Replace(Replace(Replace(AcTCMS.CloseHtml(ACTSQL(10,N)), vbCrLf, ""), "[NextPage]", ""), "&nbsp;", ""), ContentLen))
							End If 
						End if
  						If InStr(DiyContents, "#ClassName") > 0  Then
						   DiyContents = Replace(DiyContents, "#ClassName",ACTCMS.ACT_L(ACTSQL(1,N),2))
						End if
  						If InStr(DiyContents, "#ClassLink") > 0  Then
						   DiyContents = Replace(DiyContents, "#ClassLink",AcTCMS.DiyClassName(ACTSQL(1,N)))
						End If
  						If InStr(DiyContents, "#Time") > 0  Then
						   DiyContents = Replace(DiyContents, "#Time",CodeDateStr(ACTSQL(3,N),DateForm))
						End if
 						If InStr(DiyContents, "#Hits") > 0  Then
						   DiyContents = Replace(DiyContents, "#Hits","<Script Language=""Javascript"" Src=""" & Domain & "Plus/ACT.Hits.asp?A=List&ModeID="&ModeID&"&ID=" & ACTSQL(0,N)  & """></Script>")
						End If
  						If InStr(DiyContents, "#CopyFrom") > 0   Then
						   If   Trim(ACTSQL(11,N)) <> "" Then 
							 DiyContents = Replace(DiyContents, "#CopyFrom",ACTSQL(11,N))
						   Else
							 DiyContents = Replace(DiyContents, "#CopyFrom","")
						   End If 
 						End If
 						If InStr(DiyContents, "#Author") > 0   Then
						   If   Trim(ACTSQL(12,N)) <> "" Then 
							 DiyContents = Replace(DiyContents, "#Author",ACTSQL(12,N))
						   Else
							 DiyContents = Replace(DiyContents, "#Author","佚名")
						   End If 
 						End If
 						If InStr(DiyContents, "#AutoID") > 0  Then
						   DiyContents = Replace(DiyContents, "#AutoID",J)
						End if
 						If InStr(DiyContents, "#ModID") > 0  Then
 							If  N Mod 2 =0 Then ModID=0 Else ModID=1
						    DiyContents = Replace(DiyContents, "#ModID",ModID)
						End If
 						If InStr(DiyContents, "#Path") > 0  Then
						   DiyContents = Replace(DiyContents, "#Path",ACTCMS.ActCMSDM)
						End if

 						If InStr(DiyContents, "#New") > 0  Then
						   If  (Year(ACTSQL(3,N))&Month(ACTSQL(3,N))&Day(ACTSQL(3,N)) =Year(Now)&Month(Now)&Day(Now)) Then
							   DiyContents = Replace(DiyContents, "#New","<img src=""" & ACTCMS.ActCMSDM&"ACT_inc/share/new.gif"" border=""0""/>")
						   Else 
							   DiyContents = Replace(DiyContents, "#New","")
						   End If 
 						End if
					
  						If InStr(DiyContents, "#ClassSeo") > 0  Then
						   DiyContents = Replace(DiyContents, "#ClassSeo",ACTCMS.ACT_L(ACTSQL(1,N),25))
						End if
  						If InStr(DiyContents, "#ClassPicUrl") > 0  Then
						   DiyContents = Replace(DiyContents, "#ClassPicUrl",ACTCMS.ACT_L(ACTSQL(1,N),26))
						End if
  						If InStr(DiyContents, "#ClassPicFile") > 0  Then
							If ACTCMS.ACT_L(ACTSQL(1,N),26)<>"" Then 
							  DiyContents = Replace(DiyContents, "#ClassPicFile","<img src="""&ACTCMS.ACT_L(ACTSQL(1,N),26)&""" alt="""&ACTCMS.ACT_L(ACTSQL(1,N),2)&"""  /> ")
							Else 
							  DiyContents = Replace(DiyContents, "#ClassPicFile","")
							End If 
						End if
 					 
 					    ACT_SlIDeSQL = ACT_SlIDeSQL & DiyContents& vbCrLf
						N=N+1:j=j+1
 			    Next
				        ACT_SlIDeSQL = replace(ACT_SlIDeSQL,"&","%26")& vbCrLf
  		End Function

		Function GetLastArticleList(ActF,ClassID,PageStyle,ArticleSort,OpenType,PageNumber,RowHeight,TitleLen,ColNumber,TypeClassName,TypeNew,NavType,Nav,Division,DateForm,DateAlign,TitleCss,DateCss,ACTIF,DiyContent,ModeID,SubClass,ContentLen) 
		 Dim Parameter,SqlStr,ACTSQL,ACT_IF
		 If Application(AcTCMSN & "ACTCMS_TCJ_Type") <> "Folder" Then GetLastArticleList="该标签位置出错":Exit Function 
			Select Case ClassID 
			    Case "","0":Parameter=""
				Case "1"
					If Application(AcTCMSN & "classid")<>"0"  Then 
						If  CBool(SubClass)=True Then 
							 Parameter="ClassID In (" & ACTCMS.TempClassID(Application(AcTCMSN & "classid")) & ") And"
						Else 
							Parameter="ClassID='" & Application(AcTCMSN & "classid") & "' And" 
							ClassID=Application(AcTCMSN & "classid")
						End If 
					End If 
				Case Else
					If InStr(ClassID, ",") > 0 Then
						 Parameter="ClassID In (" & ClassID & ") And"
					Else
						If CBool(SubClass)=True Then 
						 Parameter="ClassID In (" & ACTCMS.TempClassID(ClassID) & ") And"
						Else 
						 Parameter="ClassID='" & Replace(ClassID,"'","") & "' And"
						End If 
					End If 
			End Select
			If ACTIF<>"" Then ACT_IF = "  "&ACTIF
			If ModeID="0" Then ModeID=Cint(Application(AcTCMSN & "modeid"))
			If Ucase(Left(Trim(ArticleSort),2))<>"id" Then  ArticleSort=ArticleSort & ",ID Desc"
			   SqlStr = "SELECT ID FROM "&ACTCMS.ACT_C(ModeID,2)&" Where " & Parameter & " isAccept=0 AND delif=0  "&ACT_IF&UIDSQL&"  order by IsTop Desc," &ArticleSort 
 			   Dim RS,N
			   Set RS=ACTCMS.ACTEXE(SqlStr)
			   If RS.EOF Then	GetLastArticleList = "<p>此栏目下没有数据</p>":Application(Cstr(AcTCMSN & "PageList")) = "":RS.Close:Set RS = Nothing:Exit Function
			   ACTSQL=RS.GetRows(-1):Set RS = Nothing
			   TotalPut=Ubound(ACTSQL,2)+1
			   PageNum=cint(PageNumber)
			   Dim PageNum, I, J, k, TempStr, OpenTypeStr,FolderNameAndLinkStr, TempTitle, NaviStr, ColSpanNum,AddDate,totalput,TempIDArrStr
				OpenTypeStr = Gopen(OpenType)
						if (TotalPut mod PageNumber)=0 then
							PageNum = TotalPut \ PageNumber
						Else 
							PageNum = TotalPut \ PageNumber + 1
						End  If 
					  For I = 1 To PageNum
						 TempIDArrStr = ""
						 For J = 1 To PageNumber
						   TempIDArrStr = TempIDArrStr &ACTSQL(0,N) & ","
						   N=N+1
						   If N>=TotalPut Then Exit For
						 Next
						  TempIDArrStr = Left(TempIDArrStr, Len(TempIDArrStr) - 1)				
						  SqlStr = "SELECT ID,ClassID,Title,UpdateTime,ActLink,FileName,infopurview,readpoint,PicUrl,Intro,Content,CopyFrom,Author,KeyWords"&ACTCMS.DIYField(ModeID)&" FROM "&ACTCMS.ACT_C(ModeID,2)&"   Where ID in (" & TempIDArrStr & ") AND isAccept=0 AND delif=0  "&ACT_IF&UIDSQL&" order by IsTop Desc," &ArticleSort 
 						  TempStr = TempStr & ACTCMS_Page_SQL(SqlStr,OpenType,RowHeight,TitleLen,ColNumber,TypeClassname,TypeNew,NavType,Nav,Division,DateForm,DateAlign,TitleCss,DateCss,ACTF,DiyContent,ModeID,SubClass,ContentLen)
						  TempStr = TempStr & AcTCMS.GetPageList(PageStyle,ACTCMS.ACT_C(ModeID,5),PageNum,I,TotalPut,PageNumber)
						  TempStr = TempStr & "{$PageList}" '加上分页符
					
						  If N>=TotalPut Then Exit For
					 Next
 						 Application(AcTCMSN & "PageList") = TempStr
						 Application(AcTCMSN & "PageStyle")= PageStyle
						 Application(AcTCMSN & "pagecount")= TotalPut
 						 GetLastArticleList = "{PageListStr}"

		End Function





		Function ACTCMS_Page_SQL(SqlStr,OpenType,RowHeight,TitleLen,ColNumber,TypeClassname,TypeNew,NavType,Nav,Division,DateForm,DateAlign,TitleCss,DateCss,ACTF,DiyContent,ModeID,SubClass,ContentLen) '21
			 Dim RS,I,K,N,DateStr,TitleCssName,ColSpanNum,TypeNews,TempTitle,NaviStr,DateCssStr,ACTSQL,DiyContents,ArticleC
			 Set RS=ACTCMS.ACTEXE(SqlStr)
			 If RS.EOF Then	 ACTCMS_Page_SQL="暂无内容":RS.Close:Set RS=Nothing:Exit Function
			 ACTSQL=RS.GetRows(-1):Set RS = Nothing
			 Dim ActNum:ActNum=Ubound(ACTSQL,2)
			 Dim Title,ClassnameLink
			 TitleCssName = GCss(TitleCss):DateCssStr = GCss(DateCss):RowHeight = GRowHeight(RowHeight):NaviStr = GNavi(NavType,Nav)
			 If ActF=2 Then 
			 Dim DIYFieldText,ModID
			 DIYFieldText=ACTCMS.DIYField(ModeID)
 			 Dim J:J=1
				For K=0 To ActNum
 					    DiyContents=DiyContent
 
 						If InStr(DiyContents, "#ID") > 0  Then
						   DiyContents = Replace(DiyContents, "#ID",ACTSQL(0,N))
						End If	
						If InStr(DiyContents, "#Link") > 0  Then
						   DiyContents = Replace(DiyContents, "#Link", AcTCMS.GetInfoUrl(ModeID,ACTSQL(1,N),ACTSQL(0,N),ACTSQL(4,N),ACTSQL(5,N),ACTSQL(6,N),ACTSQL(7,N)))
						End if
						If InStr(DiyContents, "#Title") > 0  Then
						   DiyContents = Replace(DiyContents, "#Title",ACTCMS.GetStrValue(ACTSQL(2,N),TitleLen) )
						End if
 						If InStr(DiyContents, "#CTitle") > 0  Then
						   DiyContents = Replace(DiyContents, "#CTitle",AcTCMS.CloseHtml(ACTSQL(2,N)))
						End if
 					 	If InStr(DiyContents, "#KeyWord") > 0  Then
							If ACTSQL(13,N)<>"" Then 
						    DiyContents = Replace(DiyContents, "#KeyWord","<a href=""" & Domain & "plus/search/index.asp?searchtype=3&ModeID=" & ModeID & "&keyword=" & ACTSQL(13,N)& """ target=""_blank"">" & ACTSQL(13,N) & "</a>")
							Else
						    DiyContents = Replace(DiyContents, "#KeyWord","")
							End If 
					 	End If
    					If InStr(DiyContents, "#Thumb") > 0  Then
						   If  Trim(AcTSQL(8,N))<>"" Then 
							   DiyContents = Replace(DiyContents, "#Thumb",ACTCMS.PathDoMain&ACTSQL(8,N))
						   Else
							   DiyContents = Replace(DiyContents, "#Thumb",ACTCMS.ActCMSDM&"images/nopic.gif")
						   End If 
						End if
   						If InStr(DiyContents, "#PicUrl") > 0  Then
						   If  Trim(AcTSQL(8,N))<>"" Then 
						   DiyContents = Replace(DiyContents, "#PicUrl",ACTCMS.PathDoMain&Replace(ACTSQL(8,N),"thumb_",""))
						   Else
							   DiyContents = Replace(DiyContents, "#PicUrl",ACTCMS.ActCMSDM&"images/nopic.gif")
						   End If 
 						End If
 						If InStr(DiyContents, "#Intro") > 0  Then'暂用
							If Trim(ACTSQL(9,N))<>"" Then 
							   DiyContents = Replace(DiyContents, "#Intro",ACTCMS.GetStrValue(Replace(Replace(Replace(AcTCMS.CloseHtml(ACTSQL(9,N)), vbCrLf, ""), "[NextPage]", ""), "&nbsp;", ""), ContentLen))
							Else
							   DiyContents = Replace(DiyContents, "#Intro",ACTCMS.GetStrValue(Replace(Replace(Replace(AcTCMS.CloseHtml(ACTSQL(10,N)), vbCrLf, ""), "[NextPage]", ""), "&nbsp;", ""), ContentLen))
							End If 
						End if
  						If InStr(DiyContents, "#ClassName") > 0  Then
						   DiyContents = Replace(DiyContents, "#ClassName",ACTCMS.ACT_L(ACTSQL(1,N),2))
						End if
  						If InStr(DiyContents, "#ClassLink") > 0  Then
						   DiyContents = Replace(DiyContents, "#ClassLink",AcTCMS.DiyClassName(ACTSQL(1,N)))
						End If
  						If InStr(DiyContents, "#Time") > 0  Then
						   DiyContents = Replace(DiyContents, "#Time",CodeDateStr(ACTSQL(3,N),DateForm))
						End if
 						If InStr(DiyContents, "#Hits") > 0  Then
						   DiyContents = Replace(DiyContents, "#Hits","<Script Language=""Javascript"" Src=""" & Domain & "Plus/ACT.Hits.asp?A=List&ModeID="&ModeID&"&ID=" & ACTSQL(0,N)  & """></Script>")
						End If
  						If InStr(DiyContents, "#CopyFrom") > 0   Then
						   If   Trim(ACTSQL(11,N)) <> "" Then 
							 DiyContents = Replace(DiyContents, "#CopyFrom",ACTSQL(11,N))
						   Else
							 DiyContents = Replace(DiyContents, "#CopyFrom","")
						   End If 
 						End If
 						If InStr(DiyContents, "#Author") > 0   Then
						   If   Trim(ACTSQL(12,N)) <> "" Then 
							 DiyContents = Replace(DiyContents, "#Author",ACTSQL(12,N))
						   Else
							 DiyContents = Replace(DiyContents, "#Author","佚名")
						   End If 
 						End If
 						If InStr(DiyContents, "#AutoID") > 0  Then
						   DiyContents = Replace(DiyContents, "#AutoID",J)
						End if
 						If InStr(DiyContents, "#ModID") > 0  Then
 							If  N Mod 2 =0 Then ModID=0 Else ModID=1
						    DiyContents = Replace(DiyContents, "#ModID",ModID)
						End If
 						If InStr(DiyContents, "#Path") > 0  Then
						   DiyContents = Replace(DiyContents, "#Path",ACTCMS.ActCMSDM)
						End if
 						If InStr(DiyContents, "#New") > 0  Then
						   If  (Year(ACTSQL(3,N))&Month(ACTSQL(3,N))&Day(ACTSQL(3,N)) =Year(Now)&Month(Now)&Day(Now)) Then
							   DiyContents = Replace(DiyContents, "#New","<img src=""" & ACTCMS.ActCMSDM&"ACT_inc/share/new.gif"" border=""0""/>")
						   Else 
							   DiyContents = Replace(DiyContents, "#New","")
						   End If 
 						End if
  						If InStr(DiyContents, "#ClassSeo") > 0  Then
						   DiyContents = Replace(DiyContents, "#ClassSeo",ACTCMS.ACT_L(ACTSQL(1,N),25))
						End if
  						If InStr(DiyContents, "#ClassPicUrl") > 0  Then
						   DiyContents = Replace(DiyContents, "#ClassPicUrl",ACTCMS.ACT_L(ACTSQL(1,N),26))
						End if
  						If InStr(DiyContents, "#ClassPicFile") > 0  Then
							If ACTCMS.ACT_L(ACTSQL(1,N),26)<>"" Then 
							  DiyContents = Replace(DiyContents, "#ClassPicFile","<img src="""&ACTCMS.ACT_L(ACTSQL(1,N),26)&""" alt="""&ACTCMS.ACT_L(ACTSQL(1,N),2)&"""  /> ")
							Else 
							  DiyContents = Replace(DiyContents, "#ClassPicFile","")
							End If 
						End if

					 
 
 					   ACTCMS_Page_SQL= ACTCMS_Page_SQL &DiyContents& vbCrLf
 					   N=N+1:j=j+1
				Next
			  Else
					  ACTCMS_Page_SQL = "<table border=""0"" cellpadding=""0"" cellspacing=""0"" wIDth=""100%"">" & vbCrLf
				For K=0 To ActNum
					 ACTCMS_Page_SQL = ACTCMS_Page_SQL & "<tr>" & vbCrLf
					 For I = 1 To ColNumber
					  If CBool(TypeClassname) = True Then ClassnameLink = "[" & AcTCMS.GainClassName(ACTSQL(1,N),OpenType,TitleCssName) & "]&nbsp;"			
					  If Cbool(TypeNew)=True And (Year(ACTSQL(3,N))&Month(ACTSQL(3,N))&Day(ACTSQL(3,N)) =Year(Now)&Month(Now)&Day(Now)) Then TypeNews="<img src=""" & Domain&"ACT_inc/share/new.gif"" border=""0""/>" Else TypeNews=""
					  DateStr=GDateStr(ACTSQL(3,N),DateForm,DateAlign,DateCssStr,ColNumber,ColSpanNum)
					  TempTitle = "<a " & TitleCssName &  " href=""" &AcTCMS.GetInfoUrl(Application(AcTCMSN & "ModeID"),ACTSQL(1,N),ACTSQL(0,N),ACTSQL(4,N),ACTSQL(5,N),ACTSQL(6,N),ACTSQL(7,N)) &  """"  & Gopen(OpenType) & " title=""" & AcTCMS.CloseHtml(ACTSQL(2,N)) & """>" &ACTCMS.GetStrValue(ACTSQL(2,N),TitleLen) & "</a>" 
						  If ColNumber=1 Then
							  ACTCMS_Page_SQL = ACTCMS_Page_SQL & ("  <td height=""" & RowHeight & """>"  &NaviStr&ClassnameLink&TempTitle&TypeNews&DateStr& "</td>" & vbCrLf)
						  Else
							  ACTCMS_Page_SQL = ACTCMS_Page_SQL & ("  <td  wIDth=""" & CInt(100 / CInt(ColNumber)) & "%"" height=""" &RowHeight&  """>" & vbCrLf)
							  ACTCMS_Page_SQL = ACTCMS_Page_SQL & ("    <table wIDth=""90%"" height=""100%"" cellpadding=""0"" cellspacing=""0"" border=""0"">" & vbCrLf)
							  ACTCMS_Page_SQL = ACTCMS_Page_SQL & ("     <tr><td> " &NaviStr&ClassnameLink&TempTitle&TypeNews &DateStr )
							  ACTCMS_Page_SQL = ACTCMS_Page_SQL & ("      </td></tr>" & vbcrlf &"   </table>" & vbCrLf & "  </td>" & vbCrLf)
						  End if
						  N=N+1
					      If N>=ActNum+1 Then Exit For
					 Next
					 ACTCMS_Page_SQL = ACTCMS_Page_SQL & "</tr>" & vbCrLf
					 ACTCMS_Page_SQL = ACTCMS_Page_SQL & (GbgPic(Division,ColSpanNum) & vbCrLf)
					 If N>=ActNum+1 Then Exit For
				Next
					 ACTCMS_Page_SQL = ACTCMS_Page_SQL &  ("</table>" & vbCrLf)
			End If 
		End Function

	Function GetClassNavigation(ModeID, OpenType, ColNumber, NavHeight, TCss, Division,NavType,Nav,ACTF,ThisCss,DiyContent)
			Dim I,SqlStr,RS
			Dim TempTitle,NaviStr,ColSpanNum,ACTSQL,K,N,TitleCss,DiyContents
			If ModeID ="0" Then 
				  SqlStr = "Select ClassID,ClasseName,ClassName,ParentID From Class_Act Where  ParentID='0' AND dh=1  Order by Orderid asc,ID asc"
			ElseIf  IsNumeric(ModeID) And ModeID<"20" Then 
				  SqlStr = "Select ClassID,ClasseName,ClassName From Class_Act where ParentID='0' And ModeID=" & ModeID & " AND dh=1    Order by Orderid asc,ID asc"
			ElseIf ModeID ="888" Then 
				  SqlStr = "Select ClassID,ClasseName,ClassName From Class_Act  Where  ParentID='" & Application(AcTCMSN & "classid") & "' AND dh=1       Order by Orderid asc,ID asc"
			Else
				If InStr(ModeID,",") > 0 Then
				    SqlStr = "Select ClassID,ClasseName,ClassName From Class_Act  Where  ClassID IN (" & ModeID & ")  AND dh=1  Order by Orderid asc,ID asc"
				Else
				    SqlStr = "Select ClassID,ClasseName,ClassName From Class_Act  Where   ParentID = '" & ModeID & "'  AND dh=1   Order by Orderid asc,ID asc"
				End If 
			End If 
			 Set RS=ACTCMS.ActExe(SqlStr)
			 If RS.EOF Then	 
				If  ModeID ="888" And actcms.ACT_L(Application(AcTCMSN & "classid"),11)<>"0" Then 
				  SqlStr = "Select ClassID,ClasseName,ClassName From Class_Act  Where  ParentID='" & actcms.ACT_L(Application(AcTCMSN & "classid"),11) & "' AND dh=1     Order by Orderid asc,ID asc"
				  Set RS=ACTCMS.ActExe(SqlStr)
				  If Rs.eof Then GetClassNavigation="":RS.Close:Set RS=Nothing:Exit Function
				Else 
					GetClassNavigation="":RS.Close:Set RS=Nothing:Exit Function
				End If 
			End If 
  			 ACTSQL=RS.GetRows(-1):Set RS = Nothing
			 Dim ActNum:ActNum=Ubound(ACTSQL,2)
			 Dim ModID
			 Dim J:J=1
			 OpenType = Gopen(OpenType)
			 If ActF=2 Then 
				 For K=0 To ActNum
						DiyContents=DiyContent
						If InStr(DiyContents, "#ClassName") > 0  Then
						   DiyContents = Replace(DiyContents, "#ClassName",ACTCMS.ACT_L(ACTSQL(0,N),2))
						End if
					
						If InStr(DiyContents, "#Css") > 0  Then
							If  CStr(Application(AcTCMSN & "classid"))=ACTSQL(0,N) And Trim(ThisCss)<>"" Then 
								 DiyContents = Replace(DiyContents, "#Css",ThisCss)
							Else 
								 DiyContents = Replace(DiyContents, "#Css",TCss)
							End If 
						End if
					

 						If InStr(DiyContents, "#AutoID") > 0  Then
						   DiyContents = Replace(DiyContents, "#AutoID",J)
						End if
 						If InStr(DiyContents, "#ModID") > 0  Then
 							If  N Mod 2 =0 Then ModID=0 Else ModID=1
						    DiyContents = Replace(DiyContents, "#ModID",ModID)
						End If
 						If InStr(DiyContents, "#Path") > 0  Then
						   DiyContents = Replace(DiyContents, "#Path",ACTCMS.ActCMSDM)
						End if
 						 
  						If InStr(DiyContents, "#ClassSeo") > 0  Then
						   DiyContents = Replace(DiyContents, "#ClassSeo",ACTCMS.ACT_L(ACTSQL(0,N),25))
						End if
  						If InStr(DiyContents, "#ClassPicUrl") > 0  Then
						   DiyContents = Replace(DiyContents, "#ClassPicUrl",ACTCMS.ACT_L(ACTSQL(0,N),26))
						End if
  						If InStr(DiyContents, "#ClassPicFile") > 0  Then
							If ACTCMS.ACT_L(ACTSQL(0,N),26)<>"" Then 
							  DiyContents = Replace(DiyContents, "#ClassPicFile","<img src="""&ACTCMS.ACT_L(ACTSQL(0,N),26)&""" alt="""&ACTCMS.ACT_L(ACTSQL(0,N),2)&"""  /> ")
							Else 
							  DiyContents = Replace(DiyContents, "#ClassPicFile","")
							End If 
						End if
 
						If InStr(DiyContents, "#ClassLink") > 0  Then
						   DiyContents = Replace(DiyContents, "#ClassLink",AcTCMS.DiyClassName(ACTSQL(0,N)))
						End If
					 GetClassNavigation = GetClassNavigation &DiyContents
 				     N=N+1:j=j+1
 				 Next
			 Else
				GetClassNavigation = "<table border=""0"" cellpadding=""0"" cellspacing=""0"" wIDth=""100%"" align=""center"">" & vbCrLf
				  NavHeight = GRowHeight(NavHeight):NaviStr = GNavi(NavType, Nav)&"&nbsp;"
				   For K=0 To ActNum
						GetClassNavigation = GetClassNavigation & "<tr>" & vbCrLf
						For I = 1 To ColNumber
						If  CStr(Application(AcTCMSN & "classid"))=ACTSQL(0,N) And Trim(ThisCss)<>"" Then 
							TitleCss = GCss(ThisCss)
						Else 
							TitleCss = GCss(TCss)
						End If 
						If ColNumber>=2 Then ColSpanNum = ColNumber
						TempTitle =AcTCMS.GainClassName(ACTSQL(0,N),OpenType,TitleCss) 
						 Select Case  ColNumber
							 Case 1
							 GetClassNavigation=  GetClassNavigation & ("  <td height=""" & NavHeight & """>" &NaviStr&TempTitle & "</td>" & vbCrLf)
							 Case 989
							  GetClassNavigation = GetClassNavigation & ("  <td   height=""" &NavHeight&  """>" & vbCrLf)
							  GetClassNavigation = GetClassNavigation & ("    <table wIDth=""90%"" height=""100%"" cellpadding=""0"" align=""center"" cellspacing=""0"" border=""0"">" & vbCrLf)
							  GetClassNavigation = GetClassNavigation & ("     <tr><td> " &NaviStr&TempTitle)
							  GetClassNavigation = GetClassNavigation & ("      </td></tr>" & vbcrlf &"   </table>" & vbCrLf & "  </td>" & vbCrLf)
							Case Else 
							  GetClassNavigation = GetClassNavigation & ("  <td  wIDth=""" & CInt(100 / CInt(ColNumber)) & "%"" height=""" &NavHeight&  """>" & vbCrLf)
							  GetClassNavigation = GetClassNavigation & ("    <table wIDth=""90%"" height=""100%"" align=""center"" cellpadding=""0"" cellspacing=""0"" border=""0"">" & vbCrLf)
							  GetClassNavigation = GetClassNavigation & ("     <tr><td> " &NaviStr&TempTitle)
							  GetClassNavigation = GetClassNavigation & ("      </td></tr>" & vbcrlf &"   </table>" & vbCrLf & "  </td>" & vbCrLf)
							End Select 
						  N=N+1
					      If N>=ActNum+1 Then Exit For
						Next 
				  If N>=ActNum+1 Then Exit For 
				  GetClassNavigation = GetClassNavigation & "</tr>" & vbCrLf
				  GetClassNavigation = GetClassNavigation & (GbgPic(Division,ColSpanNum) & vbCrLf)
				 Next 
				 GetClassNavigation = GetClassNavigation & ("</table>" & vbCrLf)
			 End If 
		End Function

		Function GetClassForArticleList(ClassID,ActF,ATT,ArticleSort,OpenTypeStr,ListNumber,RowHeight,TitleLen,ColNumber,TypeClassName,TypeNew,ACTIF,NavType,Nav,MoreLinkType,MoreLink,Division,DateForm,DateAlign,TitleCss,DateCss,SubColNumber,outerfor,DiyContent,MainTitleCss,PicA,PicNum,PicContentNum,ForClassContent,SubClass,ModeID) 
			If  Application(AcTCMSN & "ACTCMS_TCJ_Type")  =  "ARTICLECONTENT" Then GetClassForArticleList="标签位置出错":Exit Function 
				 Dim  SqlStr,RS,ACTSQL,n,k,Sqlstrs,Pmode,PicContent,ForClassContents
				 Dim  Parameter,MoreLinkStr,ACT_IF,ACTCMS_ATT,RSs,ACT_SQL,ACT
 				  If ModeID<>"0" Then Pmode=" And ModeID="&ModeID&" "
					If  ClassID ="0" Then '查询所有
						SqlStr = "Select ClassID,ModeID,ParentID From Class_act Where  ParentID='0' and ACTlink=1 "&Pmode&"   Order by Orderid asc,ID asc "
					ElseIf ClassID ="1" Then 
						SqlStr = "Select ClassID,ModeID,ParentID From Class_act Where  ParentID='" & Application(AcTCMSN & "classid") & "'  and  ACTlink=1 "&Pmode&"   Order by Orderid asc,ID asc"
					Else
						If InStr(ClassID, ",") > 0 Then
							SqlStr = "Select ClassID,ModeID,ParentID From Class_act Where   ClassID IN (" & ClassID & ")  and  ACTlink=1 "&Pmode&"   Order by Orderid asc,ID asc"
						Else
							SqlStr = "Select ClassID,ModeID,ParentID From Class_act Where  ClassID = '" & ClassID & "'   and  ACTlink=1 "&Pmode&"   Order by Orderid asc,ID asc"
						End If 
					End If  
				 Set RS=ACTCMS.ActExe(SqlStr)
 				 Dim TempStrs,outerfors
 				 If RS.EOF Then	 GetClassForArticleList="":RS.Close:Set RS=Nothing:Exit Function
				 ACTSQL=RS.GetRows(-1):Set RS = Nothing
				 Dim ActNum:ActNum=Ubound(ACTSQL,2)
					Dim I,jj:N=0:jj=1
				If 	ActF=2 Then 
					
 						N=0
 						 For K=0 To ActNum
							 outerfors=outerfor
 							 If InStr(ClassID, ",") > 0 Then
								 Parameter="ClassID In ('" & ACTSQL(0,N) & "') And"
							 Else
								If CBool(SubClass)=True Then 
								 Parameter="ClassID In (" & ACTCMS.TempClassID(ACTSQL(0,N)) & ") And"
								Else 
								 Parameter="ClassID='" & Replace(ACTSQL(0,N),"'","") & "' And"
								End If 
							 End If 
 							 If ACTIF<>"" Then ACT_IF = "  "&ACTIF
							 If ATT="0" Then  ACTCMS_ATT="" Else ACTCMS_ATT = " And ATT="&ATT
 							 If Ucase(Left(Trim(ArticleSort),2))<>"id" Then  ArticleSort=ArticleSort & ",ID Desc"
 							 Sqlstrs="Select TOP " & ListNumber & " ID,ClassID,Title,UpdateTime,ActLink,FileName,infopurview,readpoint,PicUrl,Intro,Content,CopyFrom,Author,KeyWords"&ACTCMS.DIYField(ACTSQL(1,N))&" From "&ACTCMS.ACT_C(ACTSQL(1,N),2)&"  Where " & Parameter & " isAccept=0 AND delif=0 " & ACTCMS_ATT &ACT_IF&UIDSQL& " ORDER BY IsTop Desc," & ArticleSort
							 Set RSs=ACTCMS.ActExe(SqlStrs)
  							 TempStrs= ACTCMS_A_Code(SqlStrs,TitleLen,MoreLinkStr,DateForm,DiyContent,ACTSQL(1,N),PicContentNum) 
 							 If PicA=1 Then 
								 Dim PicSQL,rsss,Pic_SQL,pi
								 PicSQL="Select TOP " & PicNum & " ID,ClassID,Title,UpdateTime,ActLink,FileName,infopurview,readpoint,PicUrl,Intro,Content,CopyFrom,Author,KeyWords"&ACTCMS.DIYField(ACTSQL(1,N))&" From "&ACTCMS.ACT_C(ACTSQL(1,N),2)&"  Where " & Parameter & " isAccept=0 and picurl<>''  AND delif=0 " & ACTCMS_ATT &ACT_IF&UIDSQL& " ORDER BY IsTop Desc," & ArticleSort
								 Set RSss=ACTCMS.ActExe(PicSQL)
 								 PicContent=""
								 If Not RSss.eof Then 
									 Pic_SQL=RSss.GetRows(-1):Set RSss = Nothing
									 For pi=0 To Ubound(Pic_SQL,2)
 									  ForClassContents=ForClassContent
									 
									   ForClassContents = Replace(ForClassContents, "#Link", ACTCMS.GetInfoUrl(ACTSQL(1,N),Pic_SQL(1,pi),Pic_SQL(0,pi),Pic_SQL(4,pi),Pic_SQL(5,pi),Pic_SQL(6,pi),Pic_SQL(7,pi)))
									   ForClassContents = Replace(ForClassContents, "#PicUrl", Pic_SQL(8,pi))
									   ForClassContents = Replace(ForClassContents, "#Title", ACTCMS.GetStrValue(Pic_SQL(2,pi),TitleLen))
									   ForClassContents = Replace(ForClassContents, "#Intro", ACTCMS.GetStrValue(AcTCMS.CloseHtml(Pic_SQL(9,pi)),PicContentNum))
									   PicContent=PicContent&ForClassContents
									 Next 
 										ForClassContents=""
 									 Else 
										PicContent=""
										ForClassContents=""
								  End If 
							 End If 							 
 							If InStr(outerfors, "#ClassName") > 0  Then
								 outerfors= Replace(outerfors,"#ClassName",AcTCMS.ACT_L(ACTSQL(0,N),2))
							End if
 							If InStr(outerfors, "#ClassLink") > 0  Then
								 outerfors= Replace(outerfors,"#ClassLink",AcTCMS.DiyClassName(ACTSQL(0,N)))
							End If

 							If InStr(outerfors, "#ForClassKeywords") > 0  Then
								 outerfors= Replace(outerfors,"#ForClassKeywords",AcTCMS.ACT_L(ACTSQL(0,N),8))
							End If

 							If InStr(outerfors, "#ForClassDescription") > 0  Then
								 outerfors= Replace(outerfors,"#ForClassDescription",AcTCMS.ACT_L(ACTSQL(0,N),9))
							End If


							If InStr(outerfors, "#ForClassSeo") > 0  Then
							   outerfors = Replace(outerfors, "#ForClassSeo",ACTCMS.ACT_L(ACTSQL(0,N),25))
							End if
							If InStr(outerfors, "#ForClassPicUrl") > 0  Then
							   outerfors = Replace(outerfors, "#ForClassPicUrl",ACTCMS.ACT_L(ACTSQL(0,N),26))
							End if
							If InStr(outerfors, "#ForClassPicFile") > 0  Then
								If ACTCMS.ACT_L(ACTSQL(0,N),26)<>"" Then 
								  outerfors = Replace(outerfors, "#ForClassPicFile","<img src="""&ACTCMS.ACT_L(ACTSQL(0,N),26)&""" alt="""&ACTCMS.ACT_L(ACTSQL(0,N),2)&"""  /> ")
								Else 
								  outerfors = Replace(outerfors, "#ForClassPicFile","")
								End If 
							End if


						    outerfors= Replace(outerfors,"#outerfor",TempStrs)
						    outerfors= Replace(outerfors,"#subpic",PicContent)
							outerfors = Replace(outerfors, "#AutoID",Jj)
							GetClassForArticleList =GetClassForArticleList&vbCrLf&Replace(outerfors,"#outerfor",TempStrs)
							N=N+1:jj=jj+1
						 Next 
				Else
					  Dim TypeMenuBg,TempStr,MenuBg,Act_MenuBg,OpenType,ACT_ArticleList
					  TempStr = "<TABLE BORDER=""0"" Cellpadding=""0"" Cellspacing=""2"" WIDth=""100%"">" & vbCrLf
					  Act_MenuBg = Act_HS_MenuBg(TypeMenuBg, MenuBg, ColNumber):OpenType = Gopen(OpenTypeStr)
					  For K=0 To ActNum
							TempStr = TempStr & "<TR>" & vbCrLf
							For I = 1 To ColNumber							
 							 If InStr(ClassID, ",") > 0 Then
								 Parameter=ACTSQL(0,N)
							 Else
								If CBool(SubClass)=True Then 
								 Parameter=ACTCMS.TempClassID(ACTSQL(0,N))
								Else 
								 Parameter=Replace(ACTSQL(0,N),"'","")
								End If 
								 If InStr(Parameter, ",") = 0 Then
 										 Parameter=Replace(Parameter,"'","")		
								 End If 
							 End If 
 								TempStr = TempStr & "<TD WIDth=""" & CInt(100 / CInt(ColNumber)) & "%"" height=""100%"" Valign=""top"">" & vbCrLf
								TempStr = TempStr & "<table height=""100%"" wIDth=""100%"" border=""0"" align=""center"" cellPadding=""0"" cellSpacing=""0"">" & vbCrLf
								TempStr = TempStr & "<tr><td class=""main_title""" & Act_MenuBg & ">"
								TempStr = TempStr & AcTCMS.GainClassName(ACTSQL(0,N), OpenTypeStr, "class=""main_link""") & "</td></tr>" & vbCrLf
								TempStr = TempStr & "<tr><td class=""main_tdbg"" Valign=""top"">" & vbCrLf
								ACT_ArticleList = ACT_A_List(Parameter,1,ATT,ArticleSort,OpenTypeStr,ListNumber,RowHeight,TitleLen,SubColNumber,TypeClassName,TypeNew,ACTIF,NavType,Nav,MoreLinkType,MoreLink,Division,DateForm,DateAlign,TitleCss,DateCss,DiyContent,ACTSQL(1,N),SubClass,"")
								If Trim(ACT_ArticleList) = "" Then ACT_ArticleList = "<li>此栏目下还没有文章</li>"
								TempStr = TempStr & ACT_ArticleList
								TempStr = TempStr & "</div>" & vbCrLf
								TempStr = TempStr & "</td></tr></table></td>" & vbCrLf
								N=N+1
								If N>=ActNum+1 Then Exit For
							Next
							TempStr = TempStr & "</tr>" & vbCrLf
						If N>=ActNum+1 Then Exit For
					 Next
					   TempStr = TempStr & "</TABLE>" & vbCrLf
					   GetClassForArticleList = TempStr
 				End If 
  		End Function
		

		Function ACT_Correlation_Article(ACTF,ArticleSort,ListNumber,OpenType,RowHeight,TitleLen,ColNumber,TypeClassname,TypeNew,NavType,Nav,Division,DateForm,DateAlign,TitleCss,DateCss,ACTIF,DiyContent)
		Dim ACT_IF
			 If Application(AcTCMSN & "ACTCMS_TCJ_Type") = "ARTICLECONTENT"  Then
				 Dim RS,SqlStr:SqlStr = "Select KeyWords From "&ACTCMS.ACT_C(Application(AcTCMSN & "modeid"),2)&" Where ID=" & RSQL(Application(AcTCMSN & "id")) & ""
				 Set RS=ACTCMS.ActExe(SqlStr)
 				 If RS.EOF Then	 ACT_Correlation_Article="<li>暂无相关链接":RS.Close:Set RS=Nothing:Exit Function
					 If Trim(RS(0)) <> "" And IsNull(RS(0)) = False Then
						Dim KeyWordsArr, I, SqlKeyWordStr
						KeyWordsArr = Split(Trim(RS(0)), ",")
						 For I = 0 To UBound(KeyWordsArr)
							If SqlKeyWordStr = "" Then
								SqlKeyWordStr = "KeyWords like '%" & KeyWordsArr(I) & "%' "
							Else
								SqlKeyWordStr = SqlKeyWordStr & "or KeyWords like '%" & KeyWordsArr(I) & "%' "
							End If
						Next
 					  If Ucase(Left(Trim(ArticleSort),2))<>"id" Then  ArticleSort=ArticleSort & ",ID Desc"
					   	If ACTIF<>"" Then ACT_IF = "  "&ACTIF

					  SqlStr = "Select TOP " & ListNumber & " ID,ClassID,Title,UpdateTime,ActLink,FileName,infopurview,readpoint,PicUrl,Intro,Content,CopyFrom,Author,KeyWords"&ACTCMS.DIYField(Application(AcTCMSN & "modeid"))&"  FROM "&ACTCMS.ACT_C(Application(AcTCMSN & "modeid"),2)&"  Where  (" & SqlKeyWordStr & ")   "&ACT_IF&UIDSQL&" AND isAccept=0 AND delif=0 order by IsTop Desc," & ArticleSort
 						If ActF=2 Then 
							ACT_Correlation_Article = ACTCMS_A_Code(SqlStr,TitleLen,"",DateForm,DiyContent,Application(AcTCMSN & "modeid"),"") 
						Else 
							ACT_Correlation_Article = ACTCMS_A_SQL(SqlStr,OpenType,RowHeight,TitleLen,ColNumber,TypeClassname,TypeNew,NavType,Nav,"",Division,DateForm,DateAlign,TitleCss,DateCss,Application(AcTCMSN & "modeid")) 
						End if
  					Else
					 ACT_Correlation_Article = "<li>暂无相关链接</li>"
				   End If
				 Else
					ACT_Correlation_Article = "<li>暂无相关链接</li>"
				 End If
		End Function


		 Function GetNavigation(TitleCss,OpenType,NavType,Nav,TypeMode)
		 Dim StrNav:StrNav = GetNavLocation(NavType,Nav)
 			 OpenType = Gopen(OpenType):TitleCss = GCss(TitleCss)
			Select Case UCase(Application(AcTCMSN & "ACTCMS_TCJ_Type"))
				 Case "INDEX" 
					GetNavigation = ACTCMS.GetIndexNavigation(TitleCss,OpenType,StrNav)
				 Case "FOLDER"        '栏目的位置导航
					GetNavigation = ACTCMS.GetClassNavigation(TitleCss,OpenType,StrNav, Application(AcTCMSN & "classid"),TypeMode)
				 Case "ARTICLECONTENT"     '内容页的位置导航
					GetNavigation = ACTCMS.GetContentNavigation(TitleCss,OpenType,StrNav, Application(AcTCMSN & "classid"),TypeMode)
				 Case "ACTCMSMODE"
					GetNavigation = ACTCMS.TypeModeName(TitleCss,OpenType,Application(AcTCMSN & "modeid"),StrNav)&"首页"
				 Case "OTHER"'插件导航
					GetNavigation =   "<a "& TitleCss &" href=""" & Application(AcTCMSN & "link") & """" &OpenType& ">" & Application(AcTCMSN & "ACTCMSTCJ") & "</a>"
				 Case Else
					GetNavigation = ""
			End Select
		 End Function
		Function GetNavLocation(NavType, Nav)
			If CStr(NavType) = "0" Then
			  If Nav = "" Then
			   GetNavLocation = " >> "
			  Else
			   GetNavLocation = Nav
			  End If
			Else
			  If Nav = "" Then
				GetNavLocation = " >> "
			  Else
				If Left(Nav, 1) = "/" Or Left(Nav, 1) = "\" Then Nav = Right(Nav, Len(Nav) - 1)
				GetNavLocation = "<img src=""" & ASys & Nav & """ border=""0"" align=""absmIDdle"">"
			  End If
			End If
		End Function


		Function GetLinkList(ClassLinkID, LinkType, TypeStyle, LogoWIDth, LogoHeight, ListNumber, TitleLen, ColNumber,ActF,LinkSorts,LinkContent)
			' 'on error resume next
			 Dim SqlStr, Para, SiteName, TitleStr, WIDthStr, LinkRegStr
			 Dim LinkRs:Set LinkRs=Server.CreateObject("ADODB.RECORDSET")
			 Dim k, I, NoLinkRowNumber,ACTSQL
			 LinkRegStr = Domain & "plus/Link/LinkReg.asp" '注册链接
 			 If Ucase(Left(Trim(LinkSorts),2))<>"id" Then  LinkSorts=LinkSorts & ",ID Desc"
 			 WIDthStr = CInt(100 / CInt(ColNumber)) & "%"
			 ClassLinkID = CInt(ClassLinkID):LinkType = CInt(LinkType)
			 Para = " Where Locked=0 And sh=1"
			 If ClassLinkID <> 0 Then
			   Para = Para & " And ClassLinkID=" & ClassLinkID
			 End If
			 If LinkType = 2 Then
			   Para = Para & " Order BY LinkType Desc,Rec Desc,"& LinkSorts
			 Else
			   Para = Para & " And LinkType=" & LinkType & " Order BY Rec desc,"& LinkSorts
			 End If
			 If ListNumber = 0 Then                     '列出所有友情链接站点
			   SqlStr = "Select ID,LinkType,SiteName,Description,Logo,AddDate,ClassLinkID,Url From Link_ACT" & Para
			 Else
			   SqlStr = "Select TOP " & ListNumber & " ID,LinkType,SiteName,Description ,Logo,AddDate,ClassLinkID,Url From Link_ACT" & Para
			 End If
 			 Set LinkRs=ACTCMS.ActExe(SqlStr)
 			 If LinkRs.EOF Then	 GetLinkList="":LinkRs.Close:Set LinkRs=Nothing:Exit Function

			 If ActF=2 Then 
					ACTSQL=LinkRs.GetRows(-1):Set LinkRs = Nothing
					Dim ActNum,LinkContents:ActNum=Ubound(ACTSQL,2)
  					For K=0 To ActNum
						LinkContents=LinkContent
 						If InStr(LinkContents, "#Link") > 0  Then
						   LinkContents = Replace(LinkContents, "#Link",ACTSQL(7,K))
						End if
						If InStr(LinkContents, "#Title") > 0  Then
						   LinkContents = Replace(LinkContents, "#Title",ACTSQL(2,K))
						End if
					
						If InStr(LinkContents, "#Logo") > 0  Then
							If Trim(ACTSQL(4,K))<>"" Then 
 							   LinkContents = Replace(LinkContents, "#Logo",ACTSQL(4,K))
							Else
 							   LinkContents = Replace(LinkContents, "#Logo",actcms.actsys&"images/nologo.gif")
							End If 
						End If
						
						If InStr(LinkContents, "#Description") > 0  Then
							If Trim(ACTSQL(3,K))<>"" Then 
 							   LinkContents = Replace(LinkContents, "#Description",ACTSQL(3,K))
							Else
 							   LinkContents = Replace(LinkContents, "#Description","")
							End If 
  						End If
 					   GetLinkList = GetLinkList & LinkContents& vbCrLf
 					Next 
 			Else 
 			 Select Case (CInt(TypeStyle))
				Case 0                '向上滚动
			    GetLinkList = "<div ID=rolllinkArea style=""overflow:hIDden;height:100%;wIDth:100%"">"& vbCrLf
				 GetLinkList = GetLinkList & "<div ID=rolllinkArea1>" & vbCrLf
				 GetLinkList = GetLinkList & " <table wIDth=""100%"" cellSpacing=""2""> " & vbCrLf
				  If LinkRs.EOF And LinkRs.BOF Then
					 If ClassLinkID = 0 Then                  '当显示所有类别的友情链接时,显示点击申请
					   For I = 1 To ListNumber
						 GetLinkList = GetLinkList & "<tr align=""center"" height=""22"">" & vbCrLf
						 If LinkType = 0 Then
						   GetLinkList = GetLinkList & "<td><a href=""" & LinkRegStr & """ target=""_blank"" title=""点击申请"">点击申请</a></td>"
						 Else
						   GetLinkList = GetLinkList & "<td><a href=""" & LinkRegStr & """ target=""_blank"" title=""点击申请""><Img src=""" & Domain & "ACT_inc/share/nologo.gif"" border=""0""/></a></td>"
						 End If
						GetLinkList = GetLinkList & "</tr>" & vbCrLf
					  Next
					End If
				  Else
				   Do While Not LinkRs.EOF
					 GetLinkList = GetLinkList & "<tr align=""center"" height=""22"">" & vbCrLf
					 SiteName = LinkRs(2)
					 TitleStr = " title=""网站名称:" & SiteName & "&#13;&#10;添加日期:" & LinkRs(5) & "&#13;&#10;网站描述:" & LinkRs(3) & """"
						If LinkType = 2 Then
						  If LinkRs(1) = 0 Then
						   GetLinkList = GetLinkList & "<td><a href=""" & LinkRs("Url") & """ target=""_blank""" & TitleStr & ">" & ACTCMS.GetStrValue(SiteName, TitleLen) & "</a></td>"
						  Else
						   GetLinkList = GetLinkList & "<td><a href=""" & LinkRs("Url") & """ target=""_blank""><img src=""" & LinkRs(4) & """" & TitleStr & " alt=""" & SiteName & """  wIDth=""" & LogoWIDth & """ height=""" & LogoHeight & """ border=""0""/></a></td>"
						  End If
						ElseIf LinkType = 0 Then
						  GetLinkList = GetLinkList & "<td><a href=""" & LinkRs("Url") & """ target=""_blank""" & TitleStr & ">" & ACTCMS.GetStrValue(SiteName, TitleLen) & "</a></td>"
						ElseIf LinkType = 1 Then
						  GetLinkList = GetLinkList & "<td><a href=""" & LinkRs("Url") & """ target=""_blank""><img src=""" & LinkRs(4) & """" & TitleStr & " alt=""" & SiteName & """  wIDth=""" & LogoWIDth & """ height=""" & LogoHeight & """  border=""0""/></a></td>"
						End If
						GetLinkList = GetLinkList & "</tr>" & vbCrLf
						LinkRs.MoveNext
						I = I + 1
					Loop
					If ClassLinkID = 0 Then
					 Do While I < CLng(ListNumber)
					   GetLinkList = GetLinkList & "<tr align=""center"" height=""22"">" & vbCrLf
					  If LinkType = 0 Then
					   GetLinkList = GetLinkList & "<td><a href=""" & LinkRegStr & """ target=""_blank"" title=""点击申请"">点击申请</a></td>"
					  Else
					   GetLinkList = GetLinkList & "<td><a href=""" & LinkRegStr & """ target=""_blank""  title=""点击申请""><Img src=""" & Domain & "ACT_inc/share/nologo.gif"" alt=""点击申请"" border=""0""/></a></td>"
					  End If
					   GetLinkList = GetLinkList & "</tr>" & vbCrLf
					  I = I + 1
					Loop
				   End If
				  End If
				GetLinkList = GetLinkList & "</table>"
				LinkRs.Close
				Set LinkRs = Nothing
				 GetLinkList = GetLinkList & "</div>" & vbCrLf
				 GetLinkList = GetLinkList & "<div ID=rolllinkArea2></div>" & vbCrLf
				 GetLinkList = GetLinkList & "</div>" & vbCrLf
				 GetLinkList = GetLinkList & "<script>" & vbCrLf
				 GetLinkList = GetLinkList & "var rollspeed = 20" & vbCrLf
				 GetLinkList = GetLinkList & "rolllinkArea2.innerHTML = rolllinkArea1.innerHTML" & vbCrLf
				 GetLinkList = GetLinkList & "function Marquee(){" & vbCrLf
				 GetLinkList = GetLinkList & "if(rolllinkArea2.offsetTop-rolllinkArea.scrollTop<=0)" & vbCrLf
				 GetLinkList = GetLinkList & "rolllinkArea.scrollTop-=rolllinkArea1.offsetHeight" & vbCrLf
				 GetLinkList = GetLinkList & "else{" & vbCrLf
				 GetLinkList = GetLinkList & "rolllinkArea.scrollTop++" & vbCrLf
				 GetLinkList = GetLinkList & "}}" & vbCrLf
				 GetLinkList = GetLinkList & "var MyMar = setInterval(Marquee, rollspeed)" & vbCrLf
				 GetLinkList = GetLinkList & "rolllinkArea.onmouseover=function() {clearInterval(MyMar)}" & vbCrLf
				 GetLinkList = GetLinkList & "rolllinkArea.onmouseout=function() {MyMar=setInterval(Marquee,rollspeed)}" & vbCrLf
				 GetLinkList = GetLinkList & "</script>" & vbCrLf
				
				Case 1                '横向列表
				  GetLinkList = " <table wIDth=""100%"" cellSpacing=""2""> " & vbCrLf
				  If LinkRs.EOF And LinkRs.BOF Then
					If ClassLinkID = 0 Then
					   If ListNumber = 0 Then
						  NoLinkRowNumber = 1
					   Else
						  NoLinkRowNumber = ListNumber \ ColNumber
					   End If

					   For I = 1 To NoLinkRowNumber
						  GetLinkList = GetLinkList & "<tr align=""center"">" & vbCrLf
						  For k = 1 To ColNumber
							If LinkType = 1 Then
							  GetLinkList = GetLinkList & "<td wIDth=""" & WIDthStr & """><a href=""" & LinkRegStr & """ target=""_blank"" title=""点击申请""><Img src=""" & Domain & "ACT_inc/share/nologo.gif"" alt=""点击申请"" border=""0""/></a></td>"
							Else
							  GetLinkList = GetLinkList & "<td wIDth=""" & WIDthStr & """ nowrap><a href=""" & LinkRegStr & """ target=""_blank"" title=""点击申请"">点击申请</a></td>"
							End If
						  Next
						  GetLinkList = GetLinkList & "</tr>" & vbCrLf
					   Next
					End If
				  Else
				   Do While Not LinkRs.EOF
					 If ColNumber = 1 Then
					   GetLinkList = GetLinkList & "<tr align=""center"">" & vbCrLf
					 Else
					   GetLinkList = GetLinkList & "<tr>" & vbCrLf
					 End If
					 For k = 1 To ColNumber
					 	 SiteName = LinkRs(2)
					 TitleStr = " title=""网站名称:" & SiteName & "&#13;&#10;添加日期:" & LinkRs(5) & "&#13;&#10;网站描述:" & LinkRs(3) & """"
						
						If LinkType = 2 Then
							  If LinkRs(1) = 0 Then
							   GetLinkList = GetLinkList & "<td wIDth=""" & WIDthStr & """ nowrap><a href=""" & LinkRs("Url") & """ target=""_blank""" & TitleStr & ">" & ACTCMS.GetStrValue(SiteName, TitleLen) & "</a></td>"
							  Else
							   GetLinkList = GetLinkList & "<td wIDth=""" & WIDthStr & """><a href=""" & LinkRs("Url") & """ target=""_blank""><img src=""" & LinkRs(4) & """" & TitleStr & " alt=""" & SiteName & """ wIDth=""" & LogoWIDth & """ height=""" & LogoHeight & """ border=""0""/></a></td>"
							  End If
						ElseIf LinkType = 0 Then
							  GetLinkList = GetLinkList & "<td wIDth=""" & WIDthStr & """ nowrap><a href=""" & LinkRs("Url") & """ target=""_blank""" & TitleStr & ">" & ACTCMS.GetStrValue(SiteName, TitleLen) & "</a></td>"
						ElseIf LinkType = 1 Then
							  GetLinkList = GetLinkList & "<td wIDth=""" & WIDthStr & """><a href=""" & LinkRs("Url") & """ target=""_blank""><img src=""" & LinkRs(4) & """" & TitleStr & " alt=""" & SiteName & """ wIDth=""" & LogoWIDth & """ height=""" & LogoHeight & """  border=""0""/></a></td>"
						End If
						LinkRs.MoveNext
						If LinkRs.EOF Then Exit For
					  Next

						 for  k=k+1 to ColNumber
							If LinkType = 1 Then
								   GetLinkList = GetLinkList & "<td wIDth=""" & WIDthStr & """><a href=""" & LinkRegStr & """ target=""_blank""  title=""点击申请""><Img src=""" & Domain & "ACT_inc/share/nologo.gif"" alt=""点击申请"" border=""0""/></a></td>"
								  Else
								   GetLinkList = GetLinkList & "<td wIDth=""" & WIDthStr & """ nowrap><a href=""" & LinkRegStr & """ target=""_blank"" title=""点击申请"">点击申请</a></td>"
							End If
						 next
					 GetLinkList = GetLinkList & "</tr>" & vbCrLf
					Loop
					
				  End If
				GetLinkList = GetLinkList & "</table>"
				LinkRs.Close:Set LinkRs = Nothing
				Case 2                '下拉列表
				  GetLinkList = "<select name=""Link"" onchange=""if(this.options[this.selectedIndex].value!=''){window.open(this.options[this.selectedIndex].value,'_blank');}"">" & vbCrLf
				 If LinkRs.EOF And LinkRs.BOF Then
				 GetLinkList = GetLinkList &  "<option value=''>---没有任何链接---</option>"
				 Else
				 GetLinkList = GetLinkList &  "<option value=''>---" & ACTCMS.ActExe("Select SiteName From Link_Act Where ClassLinkID=" & LinkRs("ClassLinkID"))(0) & "---</option>"
				 End If
				 Do While Not LinkRs.EOF
				   GetLinkList = GetLinkList & "<option value='" & LinkRs("Url") & "'>" & ACTCMS.GetStrValue(LinkRs(2), TitleLen) & "</option>" & vbCrLf
				   LinkRs.MoveNext
				 Loop
				  GetLinkList = GetLinkList & "</select>" & vbCrLf
				  LinkRs.Close:Set LinkRs = Nothing
			 End Select
		  End If 
 		End Function

		Function Act_HS_MenuBg(TypeMenuBg, MenuBg, ColNumber)
		  If TypeMenuBg = 0 Then
			 If MenuBg = "" Then Act_HS_MenuBg = "" Else Act_HS_MenuBg = " bgcolor=""" & MenuBg & """"
		  Else
			 If MenuBg = "" Then
			   Act_HS_MenuBg = " background=""" & Domain & "Images/Share/MenuBg" & ColNumber & ".Gif"""
			 Else
			   If Left(MenuBg, 1) = "/" Or Left(MenuBg, 1) = "\" Then MenuBg = Right(MenuBg, Len(MenuBg) - 1)
			   If LCase(Left(MenuBg, 4)) = "http" Then MenuBg = MenuBg Else MenuBg = ASys & MenuBg
			   Act_HS_MenuBg = " background=""" & MenuBg & """"
			 End If
		  End If
		End Function
 
		Function Gopen(OpenType)
			  If OpenType = "" Or OpenType = False Then
				Gopen = ""
			  ElseIf OpenType = True Then
				Gopen = " target=""_blank"""
			  Else
				Gopen = " target=""" & OpenType & """"
			  End If
		End Function
		Function GRowHeight(RowHeight)
			If IsNumeric(RowHeight) Then GRowHeight = RowHeight Else GRowHeight = 20
		End Function
		
		Function MLink(ColNumber,RowHeight,MoreLinkType, LinkNameStr, LinkUrl,OpenTypeStr)
		   If LinkNameStr = "" Then GetMoreLink = "":Exit Function
		   LinkNameStr = Trim(LinkNameStr):LinkUrl = Trim(LinkUrl)
		   If CStr(MoreLinkType) = "0" Then
			  MLink = "<tr><td colspan= """ & ColNumber+1 & """ height=""" & RowHeight & """ align=""right""><a href=""" & LinkUrl & """" & OpenTypeStr & " > " & LinkNameStr & "</a></td></tr>"
		   ElseIf CStr(MoreLinkType) = "1" Then
				MLink = "<tr><td colspan= """ & ColNumber+1 & """ height=""" & RowHeight & """ align=""right""><a href=""" & LinkUrl & """" & OpenTypeStr & " > <img src=""" & LinkNameStr & """ border=""0"" align=""absmIDdle""/></a></td></tr>"
		   Else
			 MLink = ""
		   End If
		End Function	
		
		Function MLink_D(ColNumber,RowHeight,MoreLinkType, LinkNameStr, LinkUrl,OpenTypeStr)
		   If LinkNameStr = "" Then GetMoreLink = "":Exit Function
		   LinkNameStr = Trim(LinkNameStr):LinkUrl = Trim(LinkUrl)
		   If CStr(MoreLinkType) = "0" Then
			  MLink_D = "<h4><a href=""" & LinkUrl & """" & OpenTypeStr & " > " & LinkNameStr & "</a></h4>"
		   ElseIf CStr(MoreLinkType) = "1" Then
				MLink_D = "<h4><a href=""" & LinkUrl & """" & OpenTypeStr & " > <img src=""" & LinkNameStr & """ border=""0"" align=""absmIDdle""/></a></h4>"
		   Else
			 MLink_D = ""
		   End If
		End Function	
 
		Function GbgPic(Division, ColSpanNum)
			 Dim ColStr
			 If Division = "" Then
			   GbgPic = ""
			 Else
			   If ColSpanNum>=2 Then ColStr=" colspan=" & ColSpanNum 
			   GbgPic = "<tr><td Height=1"  & ColStr & " background=" & actcms.actsys&Division & " ></td></tr>"
			 End If
		End Function

		Function GCssID(ID)
		  If ID="" Then GCssID="" Else GCssID=" ID=""" & ID & """"
		End Function  

		Function GCss(CssName)
 			 If CssName = "" Then  GCss = "" Else GCss = " class=""" & CssName & """"
		End Function

		Function GNavi(NaviType, NaviStr)
		 If CStr(NaviType) = "0" Then
			 If NaviStr = "" Then GNavi = "" Else GNavi = NaviStr
		 ElseIf CStr(NaviType) = "1" Then
		     If NaviStr <> "" Then  GNavi = "<img src=""" & actcms.actsys&NaviStr & """ border=""0""/>&nbsp;"
		 Else
			 GNavi = ""
		 End If
		End Function

		Function GDateStr(UpdateTime,DateForm,DateAlign,DateCss,ByVal ColNumber,ByRef ColSpanNum)
			   If CStr(DateForm) <> "0" And CStr("DateForm") <> "" Then
					Dim NowDate
					NowDate=Now
 					If Lcase(DateAlign)="left" Then
						GDateStr="&nbsp;<span " &  DateCss &">" & DateFormat(UpdateTime, DateForm) & "</span>"
						ColSpanNum = 1
					Else
						GDateStr="</td><td wIDth=""*"" nowrap align=" & DateAlign & "><span " &  DateCss & " >" & DateFormat(UpdateTime, DateForm) & "</span>&nbsp;"
						ColSpanNum = 2
					End If
				Else
				GDateStr="":ColSpanNum = 1
				End If
				If ColNumber>=2 Then ColSpanNum = ColNumber
		End Function
 		Function CodeDateStr(UpdateTime,DateForm)
			   If CStr(DateForm) <> "0" And CStr("DateForm") <> "" Then
					Dim NowDate
						 CodeDateStr=DateFormat(UpdateTime, DateForm)
			  Else
						 CodeDateStr= DateFormat(UpdateTime,1)
			  End If
		End Function
 		Function DateFormat(DateStr, Types)
			Dim DateString
			If IsDate(DateStr) = False Then
				DateFormat = "":Exit Function
			End If
			Select Case CStr(Types)
			  Case "0"
				DateFormat = ""
				Exit Function
			  Case 1,21,41
			      DateString=Year(DateStr) & "-" & Right("0" & Month(DateStr), 2) & "-" & Right("0" & Day(DateStr), 2)
			      if Types=21 then
				   DateString = "(" & DateString &")"
				  elseIf Types=41 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 2,22,42
			      DateString=Year(DateStr) & "." & Right("0" & Month(DateStr), 2) & "." & Right("0" & Day(DateStr), 2)
			      if Types=22 then
				   DateString = "(" & DateString &")"
				  elseIf Types=42 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 3,23,43
			      DateString=Year(DateStr) & "/" & Right("0" & Month(DateStr), 2) & "/" & Right("0" & Day(DateStr), 2)
			      if Types=23 then
				   DateString = "(" & DateString &")"
				  elseIf Types=43 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 4,24,44
			      DateString=Right("0" & Month(DateStr), 2) & "/" & Right("0" & Day(DateStr), 2) & "/" & Year(DateStr)
			      if Types=24 then
				   DateString = "(" & DateString &")"
				  elseIf Types=44 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 5,25,45
				  DateString = Year(DateStr) & "年" & Right("0" & Month(DateStr), 2) & "月"
			      if Types=25 then
				   DateString = "(" & DateString &")"
				  elseIf Types=45 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 6,26,46
				  DateString = Year(DateStr) & "年" & Right("0" & Month(DateStr), 2) & "月" & Right("0" & Day(DateStr), 2) & "日"
			      if Types=26 then
				   DateString = "(" & DateString &")"
				  elseIf Types=46 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 7,27,47
				  DateString = Right("0" & Month(DateStr), 2) & "." & Right("0" & Day(DateStr), 2) & "." & Year(DateStr)
			      if Types=27 then
				   DateString = "(" & DateString &")"
				  elseIf Types=47 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 8,28,48
				  DateString = Right("0" & Month(DateStr), 2) & "-" & Right("0" & Day(DateStr), 2) & "-" & Year(DateStr)
				  if Types=28 then
				   DateString = "(" & DateString &")"
				  elseIf Types=48 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 9,29,49
				  DateString = Right("0" & Month(DateStr), 2) & "/" & Right("0" & Day(DateStr), 2)
				  if Types=29 then
				   DateString = "(" & DateString &")"
				  elseIf Types=49 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 10,30,50
				  DateString = Right("0" & Month(DateStr), 2) & "." & Right("0" & Day(DateStr), 2)
			      if Types=30 then
				   DateString = "(" & DateString &")"
				  elseIf Types=50 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 11,31,51
				  DateString = Right("0" & Month(DateStr), 2) & "月" & Right("0" & Day(DateStr), 2) & "日"
			      if Types=31 then
				   DateString = "(" & DateString &")"
				  elseIf Types=51 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 12,32,52
				  DateString = Right("0" & Day(DateStr), 2) & "日" & Right("0" & Hour(DateStr), 2) & "时"
				  if Types=32 then
				   DateString = "(" & DateString &")"
				  elseIf Types=52 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 13,33,53
				  DateString = Right("0" & Day(DateStr), 2) & "日" & Right("0" & Hour(DateStr), 2) & "点"
			      if Types=33 then
				   DateString = "(" & DateString &")"
				  elseIf Types=53 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 14,34,54
				  DateString = Right("0" & Hour(DateStr), 2) & "时" & Minute(DateStr) & "分"
				  if Types=34 then
				   DateString = "(" & DateString &")"
				  elseIf Types=54 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 15,35,55
				  DateString = Right("0" & Hour(DateStr), 2) & ":" & Right("0" & Minute(DateStr), 2)
			      if Types=35 then
				   DateString = "(" & DateString &")"
				  elseIf Types=55 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 16,36,56
				  DateString = Right("0" & Month(DateStr), 2) & "-" & Right("0" & Day(DateStr), 2)
				 if Types=36 then
				   DateString = "(" & DateString &")"
				  elseIf Types=56 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 17,37,57
				  DateString = Right("0" & Month(DateStr), 2) & "/" & Right("0" & Day(DateStr), 2) &" " &Right("0" & Hour(DateStr), 2)&":"&Right("0" & Minute(DateStr), 2)
				  if Types=37 then
				   DateString = "(" & DateString &")"
				  elseIf Types=57 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case Else
				  DateString = DateStr
			 End Select
			 DateFormat = DateString
	   End Function

		Function CodeMLink(MoreLinkType, LinkNameStr, LinkUrl,OpenTypeStr)
		   If LinkNameStr = "" Then GetMoreLink = "":Exit Function
		   LinkNameStr = Trim(LinkNameStr):LinkUrl = Trim(LinkUrl)
		   If CStr(MoreLinkType) = "0" Then
			  CodeMLink = "<h4><a href=""" & LinkUrl & """" & OpenTypeStr & " > " & LinkNameStr & "</a></h4>"
		   ElseIf CStr(MoreLinkType) = "1" Then
				CodeMLink = "<h4><a href=""" & LinkUrl & """" & OpenTypeStr & " > <img src=""" & LinkNameStr & """ border=""0"" align=""absmIDdle""/></a></h4>"
		   Else
			 CodeMLink = ""
		   End If
		End Function	

		Function SelectLabelParameter(Content, MatchStr)
			Dim regEx, Matches, Match,N
			Set regEx = New RegExp
			regEx.Pattern = MatchStr & "[^{\=}]([\s\S]+?)}"
			regEx.IgnoreCase = True
			regEx.Global = True
			Set Matches = regEx.Execute(Content)
			SelectLabelParameter = ""
			For Each Match In Matches
				 'on error resume next
 				N=N+1
				IF N=1 Then
				SelectLabelParameter = Match.Value
				Else
				 SelectLabelParameter=SelectLabelParameter & "$$$" & Match.Value
				End IF
			Next
		End Function
		Function FunctionLabelParam(Content, MatchStr)
				FunctionLabelParam = Replace(Content, MatchStr & "(", "")
				FunctionLabelParam = Replace(Replace(FunctionLabelParam, ")}", ""), """", "")
		End Function
End Class
%>