<!--#include file="../../act_inc/ACT.User.asp"-->
<%
	Dim ModeID,ClassID,ID,rs,CommentStr
	ModeID=ChkNumeric(request("ModeID"))
	ClassID=RSQL(request("ClassID"))
	ID=ChkNumeric(request("ID"))

	If request("action")="CommentYes" Then 
     CommentStr=actcms.actexe("SELECT Count([ID]) from  Comment_Act  Where Locked=1 And ModeID=" & ModeID & " And ClassID='" & ClassID& "' And acticleID=" &id)(0)
	Else 
     CommentStr=actcms.actexe("SELECT Count([ID]) from  Comment_Act  Where  ModeID=" & ModeID & " And ClassID='" & ClassID& "' And acticleID=" &id)(0)

	End If 

  	echo  "document.write("&CommentStr&");"
 %>