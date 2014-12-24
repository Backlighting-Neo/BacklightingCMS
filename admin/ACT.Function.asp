<!--#include file="../Conn.asp"-->
<!--#include file="CheckCode.asp"-->
<!--#include file="../ACT_inc/ACT.Main.asp"-->

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
	
	'按ID显示静态标签的函数，逆光与2013年10月22日添加
	Function ShowStaticContent(LID)
		Dim ID,ModeID,rs,TemplateContent,ACT_L
		ID = 0
		ModeID = 0
		If ModeID="0" Then ModeID=1
		
		Application(AcTCMSN & "ClassID") = RSQL(actcms.s("ClassID"))
		Application(AcTCMSN & "ModeID")=ModeID
		Application(AcTCMSN & "ID")=ID
		Set ACT_L = New ACT_Code
		Set rs=actcms.actexe("select id,LabelName from Label_ACT where id="&LID)
		If rs.eof Then response.write "参数错误":response.end
		
		TemplateContent=rs("LabelName")	
		TemplateContent = ACT_L.LabelReplaceAll(TemplateContent)
		TemplateContent=ACT_L.actcmsexe(TemplateContent)
		TemplateContent=Replace(TemplateContent,vbCrLf, "")
		 Response.Write TemplateContent
	End Function
		  %>
