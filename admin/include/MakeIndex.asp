<!--#include file="../ACT.Function.asp"-->
<!--#include file="../../act_inc/ACT.Code.asp"-->
<% 
	Dim MakeIndex,TemplateContent
	Application(AcTCMSN & "ACTCMS_TCJ_Type") = "Index"
	Application(AcTCMSN & "ClassID")="0"
	Application(AcTCMSN & "ModeID")=1
	Set MakeIndex =New ACT_Code
	If Split(ACTCMS.ActCMS_Sys(4),".")(1)="asp" Then Call actcms.Alert("ACTCMS系统提醒您：\n\n1、站点首页文件以.ASP结尾的不能生成静态HTML\n\n2、请到系统设置->基本信息设置->改成Index.html","../ACT.Sys.asp"):Response.end
	Application(AcTCMSN & "ACTCMS_TCJ_Type") = "Index"
	TemplateContent = MakeIndex.LoadTemplate(ACTCMS.ActCMS_Sys(9))
	TemplateContent = MakeIndex.LabelReplaceAll(TemplateContent)
	TemplateContent = MakeIndex.actcmsexe(TemplateContent)
	IF TemplateContent = "" Then Response.Write "error":Response.End
	Call MakeIndex.FSOSaveFile(TemplateContent,ACTCMS.ActSys&ACTCMS.ActCMS_Sys(4))	
	response.Write "<a  href=""" & ACTCMS.ActSys&ACTCMS.ActCMS_Sys(4) & """target=""_blank""" &  ">首页生成成功,点击浏览</a>"
 Set MakeIndex=Nothing:Set ACTCMS=Nothing
 %>