<!--#include file="../act_inc/ACT.User.asp"-->
<!--#include file="../act_inc/MD5.asp"-->
<% ConnectionDatabase
	Select Case Request("Action")
			Case "LoginCheck"
				Call LoginCheck()
			Case "LoginOut"
				Call LoginOut()
	End Select
	Sub LoginCheck()
		if ACTCMS.ActCMS_Sys(12) = 1   Then
			Call ACTCMS.Alert("会员系统已经关闭!","")		
			Response.End
		End IF
		Dim A_PWD,PassWord,UserName,Act_Code,SqlStr,RS,act,CheckMode,CookieDate
		CookieDate=ChkNumeric(Request.Form("CookieDate"))
		if  ACTCMS.ActCMS_Sys(15) = 0 Then
			 Act_Code = Request.Form("Code")
			If CStr(Act_Code) <>CStr(Session("GetCode")) then
				Response.Write("<script>alert('验证码有误，重新输入！');history.back();</script>")
				Response.End
			End If
		End If 
		A_PWD = MD5(Request.Form("PassWord"))
		UserName = RSQL(Request.Form("UserName"))
		act = Request("act")
			 Set RS = Server.CreateObject("ADODB.RecordSet")
			  
			SqlStr="select UserID,UserName,LoginTime,LoginIP,LoginNumber,PassWord,GroupID,Locked from User_ACT where username='"&UserName&"'"
 			 RS.Open SqlStr,Conn,1,3
		 	IF Rs.eof And Rs.bof Then
				Response.Write("<script>alert('登录失败:\n\n您输入了错误的用户名，请再次输入！');history.back();</script>")
				Response.End
			Else
				IF Rs("PassWord") = A_PWD Then
					IF Rs("Locked") = 1 Then
						Response.Write("<script>alert('登录失败:\n\n您的账号未通过审核,或者账号被管理员锁定，请与您的系统管理员联系！');history.back();</script>")
						Response.End
					Else
							Dim UserRs:Set UserRS=Server.CreateOBject("ADODB.RECORDSET")
							 Dim  UserHS
							 Set UserHS = New ACT_User
 							UserRS.Open "Select * From User_ACT  Where UserID="&Rs("UserID")&"",Conn,2,3
							If not  UserRS.Bof Then
								UserRS("LoginTime") = Now
								UserRS("LoginIP") = GetIP()
								UserRS("LoginNumber") = Rs("LoginNumber")+1
							   UserRS.update
							End If 
						    If RSQL(Request.Form("CookieDate"))<>"" Then Response.Cookies(AcTCMSN).Expires = Date + 365
							Response.Cookies(AcTCMSN)("UserName") = UserName
							Response.Cookies(AcTCMSN)("PassWord") = A_PWD
							Response.Cookies(AcTCMSN)("GroupID") = Rs("GroupID")


 					End IF
				Else
						Response.Write("<script>alert('登录失败:\n\n您输入了错误的口令，请再次输入！');history.back();</script>")
						Response.End
				End If
						Rs.Close:Set Rs = Nothing
						If act="cool" Then 
						Response.Write "<script>top.location.href ='Index.asp' ;</script>"
						Else
						Response.Redirect("../Index.asp")
						End If 
			End IF


	End Sub
	Sub LoginOut()
		Response.Cookies(AcTCMSN)("UserName") = ""
		Response.Cookies(AcTCMSN)("Password") = ""
		Response.Write "<script>top.location.href ='../' ;</script>"
	End Sub

 %>
