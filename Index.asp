<!--#include file="Conn.asp"-->
<!--#include file="ACT_inc/ACT.Code.asp"-->
<!--#include file="ACT_inc/ACT.Main.asp"-->

 <%
Dim ACTCMS,TemplateContent,ACT_L,Fso
Set Fso = Server.CreateObject("scripting.FileSystemObject")
If Fso.FileExists(Server.MapPath("Install/Index.Asp")) And Not Fso.FileExists(Server.MapPath("ACT_inc/Lock/Install.lock")) Then
    Set Fso = Nothing: Response.Redirect "Install/Index.Asp": Response.End
Else
	Set ACTCMS = New ACT_Main
	If Split(ACTCMS.ActCMS_Sys(4),".")(1)<>"asp" Then 
		Response.Redirect ACTCMS.ActCMS_Sys(4):Response.End
	Else
		Set ACT_L = New ACT_Code
 			TemplateContent = ACT_L.LoadTemplate(ACTCMS.ActCMS_Sys(9))
  			Application(AcTCMSN & "ACTCMS_TCJ_Type") = "Index"
			Application(AcTCMSN & "ClassID")="0"
			Application(AcTCMSN & "ModeID")=1
			If TemplateContent = "" Then TemplateContent = "模板不存在 by ACTCMS"
			TemplateContent = ACT_L.LabelReplaceAll(TemplateContent)
  		 Response.write ACT_L.actcmsexe(TemplateContent)
	End If 
End If 
Call CheckWebSiteOnline()
Set ACT_L = Nothing: Set ACTCMS = Nothing
 %> 