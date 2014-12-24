<!--#include file="../../ACT_inc/ACT.User.asp"-->
<!--#include file="../../ACT_inc/cls_pageview.asp"-->

<%	Dim ACTCode,TemplateContent,contenttemp,j
	Dim Rs1,ActCMS_Book
	Set Rs1=ACTCMS.ACTEXE("Select PlusConfig,IsUse from Plus_ACT where PlusID='lyxt_ACT'")
	If Rs1("IsUse")=1 Then Call actcms.alert("该系统已经被管理员关闭","")
	ActCMS_Book=Split(Rs1(0),"^@$@^")
	Application(AcTCMSN & "ACTCMS_TCJ_Type")= "OTHER"
	Application(AcTCMSN & "ACTCMSTCJ")="留言本"
	Application(AcTCMSN & "link")=actcms.actsys&"plus/book/index.asp"
	Set ACTCode =New ACT_Code
	TemplateContent = ACTCode.LoadTemplate(ActCMS_Book(7))
	contenttemp=ACTCode.LoadTemplate(ActCMS_Book(8))
	If TemplateContent = "" Then TemplateContent = "模板不存在 by ACTCMS"
  	If ActCMS_Book(3)="0" Then 
	 TemplateContent=Replace(TemplateContent,"{$Code}","<input type='text' size='10' name='Code'> <img style='cursor:hand;'  src='"&ACTCMS.ActSys&"ACT_INC/Code.asp?s=+Math.random();' id='IMG1' onclick=this.src='"&ACTCMS.ActSys&"ACT_INC/Code.asp?s=+Math.random();' alt='看不清楚? 换一张！'>")
	Else
	 TemplateContent=Replace(TemplateContent,"{$Code}","")
	End If 
	if ActCMS_Book(0)="1" then 
	 TemplateContent=Replace(TemplateContent,"{$title}","<div id=""notice""><p>留言系统已经关闭,当前只能查看留言</p></div>")
	Else
	 TemplateContent=Replace(TemplateContent,"{$title}","")
	End  If
	
   Function BookContent(contenttemp)
	Dim intDateStart,content,contents
	Dim strLocalUrl
	strLocalUrl = request.ServerVariables("SCRIPT_NAME")
	Dim intPageNow
	intPageNow = ChkNumeric(request.QueryString("page"))
	Dim intPageSize, strPageInfo
	intPageSize = ActCMS_Book(4)
	Dim arrRecordInfo, i,sql,sqlCount
	sql = "SELECT [show],[name],[qq],[mail],[url],[xq],[nr],[hf],[ip],[addtime],[id]" & _
		" FROM [Book_ACT] " & _
		 " where sh=0 ORDER BY [addtime] deSC"
	sqlCount = "SELECT Count([ID])" & _
			" FROM [Book_ACT]  where sh=0 "
		Dim clsRecordInfo
		Set clsRecordInfo = New Cls_PageView
			clsRecordInfo.intRecordCount = 2816
			clsRecordInfo.strSqlCount = sqlCount
			clsRecordInfo.strSql = sql
			clsRecordInfo.intPageSize = intPageSize
			clsRecordInfo.intPageNow = intPageNow
			clsRecordInfo.strPageUrl = strLocalUrl
			clsRecordInfo.strPageVar = "page"
			clsRecordInfo.objConn = Conn		
			arrRecordInfo = clsRecordInfo.arrRecordInfo
			strPageInfo = clsRecordInfo.strPageInfo
 			If IsArray(arrRecordInfo) Then
			j=1
				For i = 0 to UBound(arrRecordInfo, 2)	
				content=contenttemp
				If InStr(content, "{$name}") > 0  And Trim(ChkCharacter(arrRecordInfo(1,i)))<>""   Then
					content = Replace(contenttemp, "{$name}", ChkCharacter(arrRecordInfo(1,i)))
				End If
				
				If InStr(content, "{$content}") > 0   And  Trim(ChkCharacter(arrRecordInfo(6,i)))<>"" Then
					If 	arrRecordInfo(0,i)="0" Then 
						If ActCMS_Book(1)="0" then
							content = Replace(content, "{$content}", ChkCharacter(arrRecordInfo(6,i)))
						Else
							content = Replace(content, "{$content}",arrRecordInfo(6,i))
						End If 
					Else
							content = Replace(content, "{$content}","<div align=""center""><font color=""#ff8000""><br>给管理员的悄悄话...</font> </div>")
					End If 
				End If 

				If InStr(content, "{$i}") > 0   Then
					content = Replace(content, "{$i}", j)
					j=j+1
				End If 
			
				If InStr(content, "{$face}") > 0  And  Trim(arrRecordInfo(5,i))<>""  Then
					content = Replace(content, "{$face}", arrRecordInfo(5,i))
				End If 

				If InStr(content, "{$ip}") > 0   Then
					content = Replace(content, "{$ip}", left(arrRecordInfo(8,i),(len(arrRecordInfo(8,i))-2))+".*")
				End If 

				If InStr(content, "{$time}") > 0   Then
					content = Replace(content, "{$time}",arrRecordInfo(9,i))
				End If 

				If InStr(content, "{$admin}") > 0  And  Trim(arrRecordInfo(7,i))<>""  Then
					content = Replace(content, "{$admin}", "<IMG src=""face/dot.gif"" width=""21"" height=""10"" border=0><font color=""red"">管理员回复</font><FONT color=#ff8000 >："&arrRecordInfo(7,i)& "</FONT>")
				Else
					content = Replace(content, "{$admin}","")
				End If 

				If InStr(content,"{$qq}") > 0 And  Trim(arrRecordInfo(2,i))<>"" Then
					content = Replace(content, "{$qq}",arrRecordInfo(2,i))
				Else
					content = Replace(content, "{$qq}","<img src=""face/nooicq.gif"" title=""没有填写""  border=0>")
				End If 
					
				If InStr(content,"{$mail}") > 0 And  Trim(arrRecordInfo(3,i))<>"" Then
					content = Replace(content, "{$mail}","<a title=""点击这里给"&arrRecordInfo(1,i)&"发送邮件"" href=""mailto:"&arrRecordInfo(3,i)&"""><img src=""face/email.gif""  border=0></a>")
				Else
					content = Replace(content, "{$mail}","<img src=""face/email1.gif"" title=""没有填写""  border=0>")
				End If 
					
				If InStr(content,"{$url}") > 0 And  Trim(arrRecordInfo(4,i))<>"" Then
					content = Replace(content, "{$url}","<a title=""请浏览我的主页""  target=""_blank""  href="""&arrRecordInfo(4,i)&"""><img src=""face/home.gif""  border=0></a>")
				Else
					content = Replace(content, "{$url}","<img src=""face/nooicq.gif"" title=""没有填写""  border=0>")
				End If 
					BookContent=BookContent&content
				Next
				BookContent=BookContent&"§"&strPageInfo
			 Else
			 
				BookContent="§"&strPageInfo
			 End IF	
	End Function 
 	Dim regEx,Matches,Match,CTemp
	Set regEx = New RegExp
	regEx.Pattern = "<!--ActPlus-->([\s\S]*?)<!--ActPlus-->"
	regEx.IgnoreCase = True
	regEx.Global = True
	Set Matches = regEx.Execute(TemplateContent)
	For Each Match In Matches
		 CTemp=BookContent(Match.SubMatches(0))
 		 TemplateContent =Replace(TemplateContent, Match.Value,Split(CTemp,"§")(0))
		 TemplateContent = Replace(TemplateContent,"{$page}",Split(CTemp,"§")(1))
	Next
    TemplateContent = ACTCode.LabelReplaceAll(TemplateContent)
 	response.write TemplateContent
 	Public Function ChkCharacter(Str)
	  If IsNull(Str) Then Exit Function
	  Dim i,tempCharacter
	  tempCharacter = Split(ActCMS_Book(6),",")
	  For i = 0 To Ubound(tempCharacter)
	   Str = Replace(Str,tempCharacter(i),"<font color=red>***</font>")
	 Next
	  ChkCharacter = Str
	End Function
  	Call CloseConn()	
   %>