<!--#include file="../Conn.asp"-->
<!--#include file="CheckCode.asp"-->
<!--#include file="../ACT_inc/ACT.Main.asp"-->
<!--#include file="../ACT_inc/ACT.Code.asp"-->

<% 	
	Dim ACTCMS,ACTERR
	Set ACTCMS = New ACT_Main
	Public Function UserLoginChecked()
	on error resume next
	Dim AdminName,PassWord
	 	If CheckManageCode=True Then 
			If CStr(Request.Cookies(AcTCMSN)("CheckManageCode")) <>CStr(CheckManageCodeContent) then
			    UserLoginChecked=false
			    Exit Function
				Response.End
			End If
		End If 
		UserLoginChecked = false 
		AdminName = RSQL(Trim(Request.Cookies(AcTCMSN)("AdminName")))
		PassWord= RSQL(Trim(Request.Cookies(AcTCMSN)("AdminPassword")))
		IF AdminName="" Or PassWord = "" Then
		   UserLoginChecked=false
		   Exit Function
		Else
			Dim UserRs
			Set Userrs=Actcms.Actexe("Select Admin_Name,PassWord From Admin_ACT Where Admin_Name='" & AdminName & "' And PassWord='" & PassWord & "'")
			IF UserRS.Eof And UserRS.Bof Then
				UserLoginChecked=false
			Else
				UserLoginChecked = true
			End if
			UserRS.Close:Set UserRS=Nothing
	   End IF
	End Function 
	IF Cbool(UserLoginChecked)=false Then
	  Response.Write "<script>top.location.href='Login.asp';</script>"
	  Response.end
	End If
	
	Function echo(content)
 		response.write content
	End Function 
	

		  %>
