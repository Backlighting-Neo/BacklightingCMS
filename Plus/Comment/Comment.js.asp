<!--#include file="../../act_inc/ACT.User.asp"-->
<%
	Dim ModeID,ClassID,ID,rs,CommentStr,names
	ModeID=ChkNumeric(request("ModeID"))
	ClassID=RSQL(request("ClassID"))
	ID=ChkNumeric(request("ID"))
  	Set rs=actcms.actexe("Select top 10 * From Comment_act Where Locked=1 And ModeID=" & ModeID & " And ClassID='" & ClassID & "' And acticleID=" & ID & " Order By Y desc ,N Desc  ,ID Desc  ,AddDate Desc")
	If Not rs.eof Then 
	Do While Not rs.eof 	
	names=ActCMS.UserM(rs("UserID"))
	If names=false Then names="[匿名]"
 

	CommentStr=CommentStr &("document.write(""<div class='pinglun_nr_info'><var>" & RS("AddDate") & "</var> <span id='badfb"&RS("id")&"' ><a  href=### onclick=postBadGood('"&actcms.actsys&"','2',"&RS("id")&")>反对</a>["&RS("n")&"]</span><span  id='goodfb"&RS("id")&"'> <a  href=### onclick=postBadGood('"&actcms.actsys&"','1',"&RS("id")&")>支持</a>["&RS("y")&"]</span>"&names&"</div><div class='pinglun_nr_content'>" & RS("Content") & "</div>"");")&vbcrlf&vbcrlf





	rs.movenext
	loop
  End If 
	echo  CommentStr
 %>