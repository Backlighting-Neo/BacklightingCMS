<!--#include file="../act_inc/ACT.User.asp"-->
<%


Dim ID,LID,ModeID,rs,TemplateContent,ACT_L
LID = ChkNumeric(Request("LID"))
ID = ChkNumeric(Request("ID"))
ModeID = ChkNumeric(Request("ModeID"))
If ModeID="0" Then ModeID=1


Application(AcTCMSN & "ClassID") = RSQL(actcms.s("ClassID"))
Application(AcTCMSN & "ModeID")=ModeID
Application(AcTCMSN & "ID")=ID
Set ACT_L = New ACT_Code
Set rs=actcms.actexe("select id,LabelName from Label_ACT where id="&LID)
If rs.eof Then response.write "document.write('参数错误');":response.end


TemplateContent=rs("LabelName")

TemplateContent = ACT_L.LabelReplaceAll(TemplateContent)
TemplateContent=ACT_L.actcmsexe(TemplateContent)

TemplateContent=Replace(TemplateContent,"""","\""")
TemplateContent=Replace(TemplateContent,vbCrLf, "")
TemplateContent=Replace(TemplateContent,"'", "\'")
TemplateContent=Replace(TemplateContent,"/", "\/")

 Response.Write "document.write('"&TemplateContent&"');"
%> 

 