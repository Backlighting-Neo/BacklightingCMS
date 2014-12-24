<!--#include file="../../act_inc/ACT.User.asp"-->
<!--#include file="../../act_inc/cls_pageview.asp"-->
<%	ConnectionDatabase
	On Error Resume Next
     Dim ACT_L,TemplateContent,SearchContent,SearchContents,Content,KeyWord,ModeID,searchtype,CTemp
  	 KeyWord=RSQL(request("KeyWord"))
	 If Trim(KeyWord)="" Then response.write "请输入关键字":response.end
 	 Set ACT_L = New ACT_Code
     TemplateContent = ACT_L.LoadTemplate("plus/Search.html")
  	 ModeID = ChkNumeric(Request("ModeID"))
 	 searchtype = ChkNumeric(Request("searchtype"))
	 if ModeID=0 or ModeID="" Then ModeID=1
 	 If InStr(TemplateContent, "{$KeyWord}") > 0  Then
		TemplateContent = Replace(TemplateContent, "{$KeyWord}", KeyWord)
	 End If 
 	 If InStr(TemplateContent, "{$ModeID}") > 0  Then
		TemplateContent = Replace(TemplateContent, "{$ModeID}", ModeID)
	 End If 
	Function Search(SearchContent)
 	 Dim strLocalUrl
	 strLocalUrl = request.ServerVariables("SCRIPT_NAME")
	 Dim intPageNow
	 intPageNow = ChkNumeric(request.QueryString("page"))
	 Dim intPageSize, strPageInfo
	 intPageSize = 10
	 Dim Arr, i,sql,sqlCount,sqls
	 Select Case searchtype
			Case "1": Sqls = " where title Like '%" & KeyWord & "%' "'标题
			Case "2": Sqls = " where title Like '%" & KeyWord & "%' or Content Like '%" & KeyWord & "%'"'标题和内容
			Case "3": Sqls = " where KeyWords Like '%" & KeyWord & "%' "'tag
			actcms.actexe("Update Tags_ACT set hits=hits+1,ClicksTime=" & NowString & " where TagsChar='" & keyword & "'")
			'Case "4": Sqls = "  "'自定义搜索
			Case Else
			  Sqls = " where title Like '%" & RSQL(request("KeyWord")) & "%' "
 	 End Select 
	

	 sql = "SELECT  [ID],[ClassID],[Title],[UpdateTime],[ActLink],[FileName],[infopurview],[readpoint],[PicUrl],[Intro],[Content],[Hits]" & _
		" FROM ["&ACTCMS.ACT_C(ModeID,2)&"] " &Sqls& _
		 " and   isAccept=0 AND delif=0 ORDER BY [ID] deSC"
 	 sqlCount = "SELECT Count([ID])" & _
			" FROM ["&ACTCMS.ACT_C(ModeID,2)&"] "&Sqls&"  and   isAccept=0 AND delif=0 "
		Dim clsRecordInfo
		Set clsRecordInfo = New Cls_PageView
			clsRecordInfo.intRecordCount = 2816
			clsRecordInfo.strSqlCount = sqlCount
			clsRecordInfo.strSql = sql
			clsRecordInfo.intPageSize = intPageSize
			clsRecordInfo.intPageNow = intPageNow
			clsRecordInfo.strPageUrl = strLocalUrl
			clsRecordInfo.strPageVar = "KeyWord="&server.URLEncode(KeyWord)&"&searchtype="&searchtype&"&ModeID="&ModeID&"&page"
			clsRecordInfo.objConn = Conn		
			Arr = clsRecordInfo.arrRecordInfo
			strPageInfo = clsRecordInfo.strPageInfo
 			If IsArray(Arr) Then
				For i = 0 to UBound(Arr, 2)	
					SearchContents=SearchContent
 						If InStr(SearchContents, "{$ArticleTitle}") > 0  Then
						   SearchContents = Replace(SearchContents, "{$ArticleTitle}", Highlight(Arr(2,i),keyword))
						End if
					
 						If InStr(SearchContents, "{$ClassName}") > 0  Then
						   SearchContents = Replace(SearchContents, "{$ClassName}", actcms.ACT_L(Arr(1,i),2))
						End if
 					
						If InStr(SearchContents, "{$ClassUrl}") > 0  Then
						   SearchContents = Replace(SearchContents, "{$ClassUrl}", actcms.DiyClassName(Arr(1,i)))
						End if
 					
						If InStr(SearchContents, "{$ArticleDate}") > 0  Then
						   SearchContents = Replace(SearchContents, "{$ArticleDate}", Arr(3,i))
						End if

 						If InStr(SearchContents, "{$ArticleHits}") > 0  Then
						   SearchContents = Replace(SearchContents, "{$ArticleHits}", Arr(11,i))
						End if

						If InStr(SearchContents, "{$ArticleUrl}") > 0  Then
						   SearchContents = Replace(SearchContents, "{$ArticleUrl}", AcTCMS.GetInfoUrl(ModeID,Arr(1,I),Arr(0,I),Arr(4,I),Arr(5,I),Arr(6,I),Arr(7,I)))
						End if
					
						If InStr(SearchContents, "{$ArticleIntro}") > 0  Then
							If Arr(9,i)<>"" Then 
 							 SearchContents = Replace(SearchContents, "{$ArticleIntro}", ACTCMS.GetStrValue(Replace(Replace(Replace(AcTCMS.CloseHtml(Arr(9,I)), vbCrLf, ""), "[NextPage]", ""), "&nbsp;", ""), 50))
							Else
 							 SearchContents = Replace(SearchContents, "{$ArticleIntro}", ACTCMS.GetStrValue(Replace(Replace(Replace(AcTCMS.CloseHtml(Arr(10,I)), vbCrLf, ""), "[NextPage]", ""), "&nbsp;", ""), 50))
							End If 
						End if
 						Content=Content&SearchContents
				Next 
				Search=Content&"§"&strPageInfo
			 Else 
				Search=Content&"§"&strPageInfo
			 End IF	
    End Function 
	Dim regEx,Matches,Match
	Set regEx = New RegExp
	regEx.Pattern = "<!--ActPlus-->([\s\S]*?)<!--ActPlus-->"
	regEx.IgnoreCase = True
	regEx.Global = True
	Set Matches = regEx.Execute(TemplateContent)
	For Each Match In Matches
		 CTemp=Search(Match.SubMatches(0))
 		 TemplateContent =Replace(TemplateContent, Match.Value,Split(CTemp,"§")(0))
		 TemplateContent = Replace(TemplateContent,"{$page}",Split(CTemp,"§")(1))
	Next
	 TemplateContent = ACT_L.LabelReplaceAll(TemplateContent)
  	Call CloseConn()	
	Set regEx = Nothing
   response.write TemplateContent
 	Function Highlight(strContent,keyword)  '标记高亮关键字 
		Dim RegEx  
		Set RegEx=new RegExp  
		RegEx.IgnoreCase =True  '不区分大小写 
		RegEx.Global=True  
		Dim ArrayKeyword,i 
		ArrayKeyword = Split(keyword," ") '用空格隔开的多关键字 
		For i=0 To Ubound(ArrayKeyword) 
			RegEx.Pattern="("&ArrayKeyword(i)&")" 
			strContent=RegEx.Replace(strContent,"<font color=red>$1</font>" )       ' 在这如只是实现不分大小写替换,可将颜色去掉
		Next 
		Set RegEx=Nothing  
		Highlight=strContent  
	End Function 

%>