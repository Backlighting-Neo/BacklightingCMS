<!--#include file="../ACT_INC/ACT.User.asp"-->
<!--#include file="../ACT_INC/MD5.asp"-->
<!--#include file="../ACT_inc/ACT.U_M.ASP"-->
<!--#include file="../Field.asp"-->
<%			

 		Dim Rs,ModeID,i,TableName,CheckMode,ACTCode
		Set ACTCode =New ACT_Code
		if ACTCMS.ActCMS_Sys(12) = "1" or ACTCMS.ActCMS_Sys(13) = "1"  Then
			Call ACTCMS.Alert("系统关闭注册或会员系统已经关闭",ACTCMS.ActCMS_Sys(3))		
			Response.End
		End If
	ModeID=ChkNumeric(request("ModeID"))
	Select Case Request.QueryString("action")	
			   Case "save" 
			   		Call SaveUser()
					Response.End
			   Case "RegM"
			   		call regok()
			   Case else
					Call main()
					Response.End
		End Select
		Sub SaveUser()
		Dim UserName ,Codes,PassWord,RPassWord,Email,TempRs,AdminRS,AdminSql,Regrz,CheckNum,Question,Answer
		if actcms.actexe("select RegCode from ModeUser_Act where ModeID="&ModeID&"")(0) = 0 Then
		 Codes = Request.Form("Code")
			If CStr(Codes) <>CStr(Session("GetCode")) then
				Call ACTCMS.Alert("验证码有误，重新输入","")		
				Response.End
			End if
		End  If 
			 Dim rs1
	  	     Set Rs1=ACTCMS.actexe("select ModeTable from ModeUser_Act where ModeID="&ModeID&" ")
			 If Not Rs1.eof Then
				TableName=Rs1("ModeTable")
			 Else 
				Call actcms.alert("未定义操作","")
				response.End 
			 End if	
			UserName = ACTCMS.S("UserName")
			PassWord = Request.Form("PassWord")
			RPassWord = Request.Form("RPassWord")	
			Email = ACTCMS.S("Email")	
			Question = ACTCMS.S("Question")
			Answer = ACTCMS.S("Answer")
			IF Trim(PassWord) <> Trim(RPassWord) Then
				Call ACTCMS.Alert("2次输入的密码不一致!","")		
				Response.end
			End IF
			
			IF UserName <> "" Then
				IF Len(UserName) >= 100 Then
				    Call ACTCMS.Alert("用户名不能超过50个字符","")		
					Response.end
				End IF
			Else
				    Call ACTCMS.Alert("请输入用户名","")		
					Response.end
			End if	
		
			if UserName="" then
			 Call ACTCMS.Alert("请输入会员名","")	
			elseif InStr(UserName, "=") > 0 Or InStr(UserName, ".") > 0 Or InStr(UserName, "%") > 0 Or InStr(UserName, Chr(32)) > 0 Or InStr(UserName, "?") > 0 Or InStr(UserName, "&") > 0 Or InStr(UserName, ";") > 0 Or InStr(UserName, ",") > 0 Or InStr(UserName, "'") > 0 Or InStr(UserName, ",") > 0 Or InStr(UserName, Chr(34)) > 0 Or InStr(UserName, Chr(9)) > 0 Or InStr(UserName, "") > 0 Or InStr(UserName, "$") > 0 Or InStr(UserName, "*") Or InStr(UserName, "|") Or InStr(UserName, """") > 0 Then
			 Call ACTCMS.Alert("用户名中含有非法字符","")	
			elseif len(UserName)<4 or len(UserName)>16 then
			 Call ACTCMS.Alert("会员名长度应为4-16位!","")	
			elseif ACTCMS.FoundInArr(ACTCMS.ActCMS_Sys(16), UserName, "|") = True Then
			 Call ACTCMS.Alert("您输入的用户名为系统禁止注册的用户名","")	
			end if
 		 if ACTCMS.IsValidEmail(Email)=false then
			  Call ACTCMS.Alert("请输入正确的电子邮箱!","")	
		 End if
 			Set CheckMode = ACTCMS.ACTEXE("Select ModeTable from ModeUser_Act")
			If Not CheckMode.eof Then 
				Do While Not CheckMode.eof
					Set TempRs = Conn.Execute("Select UserName from User_ACT where UserName='" & UserName & "'")
					IF Not TempRs.Eof Then
							Call ACTCMS.Alert("数据库中已存在该用户名","")		
							Response.end
					End IF			
				 CheckMode.movenext
				 Loop
			End If 
				  Dim grs,GroupSettingSet,dianshu
				  Set grs = actcms.actexe("Select GroupSetting From Group_Act Where DefaultGroup=1 and ModeID=" & ModeID)
				  If  grs.eof Then response.End 
				  GroupSettingSet=Split(grs("GroupSetting"),"^@$@^")
				  Regrz=GroupSettingSet(12)
				  dianshu=GroupSettingSet(14)
				  Set AdminRS = Server.CreateObject("adodb.recordset")
				  AdminSql = "select * from User_ACT"
				  AdminRS.Open AdminSql, Conn, 1, 3
				  AdminRS.AddNew
				  AdminRS("RegDate") = Now
				  AdminRS("GroupID") = actcms.actexe("Select GroupID From Group_Act Where DefaultGroup=1 and ModeID=" & ModeID)(0)
				  AdminRS("Loginip") = GetIP()
				  AdminRS("LoginTime") = Now
				  AdminRS("UserName") = UserName
				  AdminRS("templetsid") = 1
				  AdminRS("UModeID") = ModeID
				  AdminRS("ChargeType") = 1
  				  If dianshu<>"0" Then AdminRS("Score") = dianshu
				  AdminRS("PassWord")=MD5(RPassWord)
				  AdminRS("Answer")=Answer
				  AdminRS("Question") = Question
				  Select Case Regrz
						Case "1"
							AdminRS("Locked") = 0
						Case "2"
							 if ACTCMS.IsValidEmail(Email)=false then
								  Call ACTCMS.Alert("请输入正确的电子邮箱!","")	
							 End if
							CheckNum=MD5(actcms.MakeRandom(10))
							AdminRS("Locked") = 1
							AdminRS("CheckNum") = CheckNum
						   Dim Errors,MailBodyStr
						   MailBodyStr="欢迎您注册成为本站会员！<BR><BR>验证码："&CheckNum&"<BR>请点击下面的地址，输入上面的验证码进行邮件验证。验证通过后，您就可以正式成为我们的会员，享受有关服务了！<BR><BR><A HREF="&ACTCMS.ActUrl &"User/UserRegCheck.asp?UserName=" & UserName &"&CheckNum=" & CheckNum&" TARGET=_BLANK>"&ACTCMS.ActUrl &"User/UserRegCheck.asp?UserName=" & UserName &"&CheckNum=" & CheckNum
						   Errors=ACTCMS.SendMail(ACTCMS.ActCMS_Other(3), ACTCMS.ActCMS_Other(4), ACTCMS.ActCMS_Other(5), AcTCMS.ActCMS_Sys(0) & "-会员注册确认信", Email,UserName, MailBodyStr,ACTCMS.ActCMS_Other(4))
						   IF Errors="OK" Then
								 Errors="注册成功，注册验证码已发送到您的信箱" &Email &"，只有激活后才可以正式成为本站会员!"
						   Else
								Call ACTCMS.Alert("信件发送失败!失败原因:" & Errors & "，请联系网站管理员!","")
						   End if
						Case "3"
							AdminRS("Locked") = 1
				  End Select 
				  AdminRS("Email") = Email
			

				  AdminRS.Update
				  '自定义字段开始入库
			Dim UserID,FieldRS,FieldSql
			Set rs=ACTCMS.actexe("Select top 1 UserID from User_ACT  order by UserID desc")
			If Not rs.eof Then UserID = rs("UserID")

 
				  Set FieldRS = Server.CreateObject("adodb.recordset")
				  FieldSql = "select * from "&ACTCMS.ACT_U(ModeID,2)&" where userid<>"&userid&""

 				  FieldRS.Open FieldSql, Conn, 1, 3
				  FieldRS.AddNew
				  FieldRS("UserID") = UserID
 

 				 Dim IF_NULL
				IF_NULL=ActUser_MX_Arr(ModeID)
				If IsArray(IF_NULL) Then
				For I=0 To Ubound(IF_NULL,2)
				 If IF_NULL(2,I)=0 And Trim(ACTCMS.S(IF_NULL(0,I)))="" Then  Call  ACTCMS.ALERT(IF_NULL(1,I)&"不能为空","")
				Next
				End If
				If IsArray(IF_NULL) Then
					For I=0 To Ubound(IF_NULL,2)
 						If IF_NULL(3,I)="NumberType" Then 
						   If actcms.regexField(ACTCMS.S(IF_NULL(0,I)),"^\d+$")=True Then 
							   FieldRS("" & IF_NULL(0,I) & "" )= ACTCMS.S(IF_NULL(0,I))
						   End If 
						ElseIf IF_NULL(3,I)="DateType" Then 
							If IsDate(ACTCMS.S(IF_NULL(0,I)))=False Then 
								FieldRS("" & IF_NULL(0,I) & "")= Now()
							Else 
								FieldRS("" & IF_NULL(0,I) & "")=ACTCMS.S(IF_NULL(0,I))
							End If
						ElseIf IF_NULL(4,I)="1" Then 
								 FieldRS("" & IF_NULL(0,I) & "")= actcms.AField(IF_NULL(5,I))
						ElseIf IF_NULL(4,I)="2" Then 
								If actcms.regexField(ACTCMS.S(IF_NULL(0,I)),IF_NULL(5,I))=True Then 
									FieldRS("" & IF_NULL(0,I) & "")=ACTCMS.S(IF_NULL(0,I))
								Else 
									Call Actcms.Alert(IF_NULL(6,I),"")
								End If 
						Else 
							FieldRS("" & IF_NULL(0,I) & "")=ACTCMS.S(IF_NULL(0,I))
						End If 
						actField=""
					Next
				End If
			'结束

				  FieldRS.Update
 				If Regrz="1" Then	
						  Response.Cookies(AcTCMSN)("UserName") = UserName
						  Response.Cookies(AcTCMSN)("GroupID") = AdminRS("GroupID")
						  Response.Cookies(AcTCMSN)("Password") = MD5(RPassWord)
						  
						Errors="注册成功!您的用户名:" & UserName & ",您已成为了本站的正式会员!"
				Else
					Errors="注册成功!您的用户名:" & UserName & ",您需要通过管理员的认证才能成为正式会员!"	
				End If 
				Call ACTCMS.Alert(Errors,"index.asp")
			Response.end
		End Sub

   Function ActUser_MX_Arr(ModeID)'返回模型数组
	  Dim Rs
	  Set Rs=actcms.ACTEXE("Select FieldName,Title,IsNotNull,FieldType,[check],regex,regError from Table_ACT  Where ModeID=" & ModeID & " and SearchIF=1 and actcms=2  order by OrderID desc,ID Desc")
	 If Not Rs.Eof Then
	  ActUser_MX_Arr=Rs.GetRows(-1)
	 Else
	  ActUser_MX_Arr=""
	 End If
	 Rs.Close:Set Rs=Nothing
   End Function

Sub main()
	Dim TemplateContent,content,RsU
	Application(AcTCMSN & "ACTCMS_TCJ_Type")= "OTHER"
	Application(AcTCMSN & "ACTCMSTCJ")="用户注册协议"
	Application(AcTCMSN & "link")=actcms.actsys&"user/reg.asp"

	
	TemplateContent=ACTCode.LoadTemplate(ACTCMS.ActCMS_Sys(17))
	If TemplateContent = "" Then TemplateContent = "模板不存在 by ACTCMS"
	TemplateContent = ACTCode.LabelReplaceAll(TemplateContent)
	Set RsU=ACTCMS.ACTEXE("Select ModeID, ModeName From ModeUser_Act order by ModeID asc")
	IF not Rsu.eof then 
		Do while Not Rsu.eof
		content=content& "<input  class=""button_style""  type=""button""  value=""   "&rsu("ModeName")&"   "" onclick=""document.getElementById('readpact').checked ?window.location.href='?action=RegM&ModeID="&rsu("ModeID")&"' : alert('您必须同意本站注册协议才能进行注册！');"" />&nbsp;&nbsp;"
		Rsu.movenext
		loop	
	Else 
		response.write "未定义操作"
		response.end
	End  if 
	TemplateContent=Replace(TemplateContent,"{$ModeName}",content)
	response.write TemplateContent
	
	%>
<%End Sub


sub regok()
	Dim rss,Regrz
	Set Rss = ACTCMS.ACTEXE("SELECT Groupid,GroupSetting FROM Group_Act Where DefaultGroup=1 and  ModeID=" & ModeID & " order by ModeID desc")
	 If rss.eof Then 
		response.write "<font color=red>还未设置默认用户组</font>"
		response.end
	 End if	

	 Regrz= Split(Rss("GroupSetting"),"^@$@^")(12)



	Dim TemplateContent
	Application(AcTCMSN & "ACTCMS_TCJ_Type")= "OTHER"
	Application(AcTCMSN & "ACTCMSTCJ")="用户注册"
	Application(AcTCMSN & "link")=actcms.actsys&"user/reg.asp?action=RegM&ModeID="&ModeID&""
	TemplateContent=ACTCode.LoadTemplate(ACTCMS.ActCMS_Sys(18))
	If TemplateContent = "" Then TemplateContent = "模板不存在 by ACTCMS"
	TemplateContent = ACTCode.LabelReplaceAll(TemplateContent)

	TemplateContent = Replace(TemplateContent, "{$ModeID}",ModeID)
	If InStr(TemplateContent, "{$regtitle}") > 0  And  regrz>1  Then
		TemplateContent = Replace(TemplateContent, "{$regtitle}", "<font color=red>该组注册需要经过管理员的认证才可以通过</font>")
	Else
		TemplateContent = Replace(TemplateContent, "{$regtitle}","")
	End If 

	If InStr(TemplateContent, "{$RegMode}") > 0   Then
		TemplateContent = Replace(TemplateContent, "{$RegMode}",  U_M.ACT_NoRormMXList(ModeID))
	Else
		TemplateContent = Replace(TemplateContent, "{$RegMode}","")
	End If 
	
	
	if  actcms.actexe("select RegCode from ModeUser_Act where ModeID="&ModeID&"")(0)= 0 Then
	 TemplateContent=Replace(TemplateContent,"{$regcode}","<tr class=""td_bg""><th align=""right"">验 证 码：</th> <td><input type='text' size='10' name='Code'> <img style='cursor:hand;'  src='"&ACTCMS.ActSys&"ACT_INC/Code.asp?s=+Math.random();' id='IMG1' onclick=this.src='"&ACTCMS.ActSys&"ACT_INC/Code.asp?s=+Math.random();' alt='看不清楚? 换一张！'></td></tr>")
	Else
	 TemplateContent=Replace(TemplateContent,"{$regcode}","")
	End If 

	if  actcms.actexe("select RegCode from ModeUser_Act where ModeID="&ModeID&"")(0)= 0 Then
	 TemplateContent=Replace(TemplateContent,"{$regjscode}","if (form.Code.value=='')"&vbCrLf&"{ alert(""请输入验证码!"");"&vbCrLf&"form.Code.focus(); "&vbCrLf&"return false;"&vbCrLf&"}")
	Else
	 TemplateContent=Replace(TemplateContent,"{$regjscode}","")
	End If 
	
	
	response.write TemplateContent

	end sub%>
