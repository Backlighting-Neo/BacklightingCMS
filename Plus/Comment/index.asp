<!--#include file="../../ACT_inc/ACT.User.asp"-->
<!--#include file="../../act_inc/cls_pageview.asp"-->
<%ConnectionDatabase
	 Dim ACT_L,TemplateContent,ModeID,CommUser,ClassID,ID,Rs,Comment,Comments,Content,UN,CTemp
	 ModeID=ChkNumeric(request("ModeID"))
	 IF ModeID=0  Then Response.End
	 ClassID=RSQL(request("ClassID"))
	 ID=ChkNumeric(request("ID"))
     Set ACT_L = New ACT_Code
     Set CommUser = New ACT_User
 	 Set Rs=actcms.actexe("Select  [ClassID],[ID],[actlink],[FileName],[InfoPurview],[ReadPoint],[title],[updatetime], * From "&ACTCMS.ACT_C(ModeID,2)&" where ID=" & ID)
	 If   rs.eof Then response.write "error":response.end
	Dim DocXML,Node:Set DocXML=actcms.arrayToXml(Rs.GetRows(1),Rs,"row","root")
	Set Node=DocXml.DocumentElement.SelectSingleNode("row")
	Set ACT_L.Nodes=DocXml.DocumentElement.SelectSingleNode("row")
	 Application(AcTCMSN & "classid") = ACT_L.GetNodeText("classid")
	 Application(AcTCMSN & "modeid")=ModeID
	 Application(AcTCMSN & "id")=ACT_L.GetNodeText("id")
 	 TemplateContent = ACT_L.LoadTemplate(Actcms.ACT_C(ModeID,10))
  	 If InStr(TemplateContent, "{$ArticleUrl}") > 0  Then
		TemplateContent = Replace(TemplateContent, "{$ArticleUrl}", ACTCMS.GetInfoUrl(ModeID,rs(0),rs(1),rs(2),rs(3),rs(4),rs(5)))
	 End If 
  	 
	 Function CommentContent(Comment)
	 Dim strLocalUrl
	 strLocalUrl = request.ServerVariables("SCRIPT_NAME")
	 Dim intPageNow
	 intPageNow = ChkNumeric(request.QueryString("page"))
	 Dim intPageSize, strPageInfo
	 intPageSize = 10
	 Dim Arr, i,sql,sqlCount,sqls
 
	 sql = "Select  ID,ModeID,ClassID,acticleID,Email,UserIP,Content,Locked,AddDate,userid,Y,N From Comment_Act Where Locked=1 And ModeID=" & ModeID & " And ClassID='" & ClassID & "' And acticleID=" & ID & "  Order By AddDate Desc"
 	sqlCount = "SELECT Count([ID]) from  Comment_Act  Where Locked=1 And ModeID=" & ModeID & " And ClassID='" & ClassID & "' And acticleID=" & ID & " "
 		Dim clsRecordInfo
		Set clsRecordInfo = New Cls_PageView
			clsRecordInfo.intRecordCount = 2816
			clsRecordInfo.strSqlCount = sqlCount
			clsRecordInfo.strSql = sql
			clsRecordInfo.intPageSize = intPageSize
			clsRecordInfo.intPageNow = intPageNow
			clsRecordInfo.strPageUrl = strLocalUrl
			clsRecordInfo.strPageVar = "ModeID="&ModeID&"&ClassID="& ClassID&"&ID="&ID&"&page"
			clsRecordInfo.objConn = Conn		
			Arr = clsRecordInfo.arrRecordInfo
			strPageInfo = clsRecordInfo.strPageInfo
			If IsArray(Arr) Then
				For i = 0 to UBound(Arr, 2)	
				On Error Resume Next
					Comments=Comment
 						If InStr(Comments, "{$ID}") > 0  Then
						   Comments = Replace(Comments, "{$ID}", Arr(0,i))
						End if
						If InStr(Comments, "{$ModeID}") > 0  Then
						   Comments = Replace(Comments, "{$ModeID}", Arr(1,i))
						End if
						If InStr(Comments, "{$ClassID}") > 0  Then
						   Comments = Replace(Comments, "{$ClassID}", Arr(2,i))
						End if
						If InStr(Comments, "{$acticleID}") > 0  Then
						   Comments = Replace(Comments, "{$acticleID}", Arr(3,i))
						End if
						If InStr(Comments, "{$Email}") > 0  Then
						   Comments = Replace(Comments, "{$Email}", Arr(4,i))
						End if
						If InStr(Comments, "{$UserIP}") > 0  Then
						   Comments = Replace(Comments, "{$UserIP}", Replace(Arr(5,i),split(Arr(5,i),".")(3),"*"))
						End if
						If InStr(Comments, "{$Content}") > 0  Then
						   Comments = Replace(Comments, "{$Content}", Arr(6,i))
						End if
						If InStr(Comments, "{$Locked}") > 0  Then
						   Comments = Replace(Comments, "{$Locked}", Arr(7,i))
						End if
					 
						If InStr(Comments, "{$AddDate}") > 0  Then
						   Comments = Replace(Comments, "{$AddDate}", Arr(8,i))
						End if
						
						If InStr(Comments, "{$userid}") > 0  Then
						   Comments = Replace(Comments, "{$userid}", Arr(9,i))
						End if
						
						If InStr(Comments, "{$Y}") > 0  Then
						   Comments = Replace(Comments, "{$Y}", Arr(10,i))
 						End if
						
						If InStr(Comments, "{$N}") > 0  Then
						   Comments = Replace(Comments, "{$N}", Arr(11,i))
						End if
						
						If InStr(Comments, "{$UserName}") > 0  Then
							Dim u
							 u=ActCMS.UserM(Arr(9,i))
 							If u=false Then 
							   Comments = Replace(Comments, "{$UserName}","匿名")
							Else
							   Comments = Replace(Comments, "{$UserName}",u)
							End If 
   						End if
 						
						 
 					 CommentContent=CommentContent&Comments
				Next 
				CommentContent=CommentContent&"§"&strPageInfo
			 Else 
				CommentContent="§"&strPageInfo
  			 End IF	
 	End Function 
 	IF Cbool(CommUser.UserLoginChecked)=false then
		UN=("用户名:<input type='text' name='username' size='16' class='ipt-txt' />E-Mail:<input type='text' name='Email' size='16' class='ipt-txt' />")&vbcrlf&vbcrlf
 	Else
		UN=("用户名:"&CommUser.username&"")&vbcrlf&vbcrlf
  	End If	

	TemplateContent= Replace(TemplateContent,"{$Login}",UN)
	UN=""
	If Actcms.ACT_C(ModeID,13)=0 Then
		UN = ("&nbsp;验证码：<input type='text' class='ipt-txt' size='4' name='Code' />")&vbcrlf&vbcrlf
		UN = UN &("<img style='cursor:hand;' src='"&ACTCMS.ActSys&"ACT_INC/Code.asp?s=+Math.random();' id='IMG1' onclick=this.src='"&ACTCMS.ActSys&"ACT_INC/Code.asp?s=+Math.random();' alt='看不清楚? 换一张！'>")&vbcrlf&vbcrlf
	End If
	
	
	TemplateContent= Replace(TemplateContent,"{$Code}",UN)
	UN=""
	If InStr(TemplateContent, "{$UserName}") > 0  Then
 		UN=ActCMS.UserM(CommUser.UserID)
		If UN=false Then 
		   TemplateContent = Replace(TemplateContent, "{$UserName}","匿名")
		Else
		   TemplateContent = Replace(TemplateContent, "{$UserName}",UN)
		End If 
	End If
	
	Dim regEx,Matches,Match
	Set regEx = New RegExp
	regEx.Pattern = "<!--ActPlus-->([\s\S]*?)<!--ActPlus-->"
	regEx.IgnoreCase = True
	regEx.Global = True
	Set Matches = regEx.Execute(TemplateContent)
	For Each Match In Matches
		 CTemp=CommentContent(Match.SubMatches(0))
 		 TemplateContent =Replace(TemplateContent, Match.Value,Split(CTemp,"§")(0))
		 TemplateContent = Replace(TemplateContent,"{$page}",Split(CTemp,"§")(1))
	Next
 	 TemplateContent = ACT_L.LabelReplaceAll(TemplateContent)
	 TemplateContent = ACT_L.ReplaceArticleContent(ModeID,TemplateContent,"")
  	Call CloseConn()	
	Set regEx = Nothing
 response.write TemplateContent

 
 %>