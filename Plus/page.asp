<!--#include file="../act_inc/ACT.User.asp"-->
 <%	ConnectionDatabase
	on error resume next
     Dim ACT_L,TemplateContent,ID,rs
  	 Set ACT_L = New ACT_Code
  	 ID = ChkNumeric(Request("ID"))
 	 if ID=0 or ID="" Then ID=1
	Application(AcTCMSN&"ModeID")=1
	 Set rs=actcms.actexe("select id,tempurl,pagename,content from DiyPage_ACT where id="&id)
	 If rs.eof Then response.write "参数错误":response.end
	 
	 TemplateContent = ACT_L.LoadTemplate(rs("tempurl"))
	
	If InStr(TemplateContent, "{$diycontent}") > 0   Then
		TemplateContent = Replace(TemplateContent, "{$diycontent}", rs("content"))
	End If 

	If InStr(TemplateContent, "{$pagename}") > 0  And  rs("pagename")<>""  Then
		TemplateContent = Replace(TemplateContent, "{$pagename}", rs("pagename"))
	Else
		TemplateContent = Replace(TemplateContent, "{$pagename}", "")
	End If 

	 
	 
	 TemplateContent = ACT_L.LabelReplaceAll(TemplateContent)
	 TemplateContent=ACT_L.actcmsexe(TemplateContent)
	 response.write TemplateContent
%>