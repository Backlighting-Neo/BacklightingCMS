 <!--#include file="../act_inc/ACT.User.asp"-->
<%
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim UserHS,A
Set UserHS = New ACT_User
 A=request("A")

 Select Case A
	Case "JS"
		Call js()
	Case "Html"	
		Call Html()
	Case Else
		Call Html()
 End Select 
 
 Sub Html()
%>


<link rel="stylesheet" rev="stylesheet" type="text/css" href="../images/actcms.css"/>

			<%
		 
			If CBool(UserHS.UserLoginChecked) = False Then
			%>
			<table align="center" width="98%" border="0" cellspacing="0" cellpadding="0">
			 <form name="myform" action="<%= ACTCMS.ActSys %>User/Checklogin.asp?Action=LoginCheck" method="post" onSubmit="return(CheckForm())">
			  <tr>
				<td height="25">用户名：
				<input name="UserName" type="text" class="textbox" id="UserName" size="13"></td>
			  </tr>
			  <tr>
				<td height="25">密　码：
				<input name="PassWord" type="PassWord" class="textbox" id="Password" size="15"></td>
			  </tr>
			  <% if ACTCMS.ActCMS_Sys(15) = 0 Then%>
			  <tr>
				<td height="25">验证码：
				<input name="Code" type="text" class="textbox" id="Code" size="6">
				<img src="../act_inc/code.asp?s='+Math.random();" alt="验证码" title="看不清楚? 换一张！" style="cursor:hand;" onClick="src='../act_inc/code.asp?s='+Math.random()"/> 
				</td>
			  </tr>
			  	<%end if %>
			  <tr>
				<td height="25"><div align="center">  <a href="<%= ACTCMS.ActCMSDM %>user/GetPass.asp" target="_parent">忘记密码</a>   <a href="<%= ACTCMS.ActCMSDM %>User/Reg.asp" target="_parent">新会员注册</a>    </div></td>
			  </tr>
			  <tr>
				<td height="25"><div align="center">
				  <input type="submit" name="Submit" class="inputButton" value="登录">
				  <input name="CookieDate" type="checkbox" id="CookieDate" value="checkbox"><label for="CookieDate">永久登录</label></div></td>
			  </tr>
			  </form>
            </table><%
			Else
 			%>
			<table align="center" width="80%" border="0" cellspacing="0" cellpadding="0">
			<tr><td align="center"><font color=red><%=UserHS.UserName%></font>,
           <%
			If (Hour(Now) < 6) Then
            Response.Write "<font color=##0066FF>凌晨好!</font>"
			ElseIf (Hour(Now) < 9) Then
				Response.Write "<font color=##000099>早上好!</font>"
			ElseIf (Hour(Now) < 12) Then
				Response.Write "<font color=##FF6699>上午好!</font>"
			ElseIf (Hour(Now) < 14) Then
				Response.Write "<font color=##FF6600>中午好!</font>"
			ElseIf (Hour(Now) < 17) Then
				Response.Write "<font color=##FF00FF>下午好!</font>"
			ElseIf (Hour(Now) < 18) Then
				Response.Write "<font color=##0033FF>傍晚好!</font>"
			Else
				Response.Write "<font color=##ff0000>晚上好!</font>"
			End If
			%>&nbsp;&nbsp;&nbsp;</td></tr>
		 
 			<tr><td>登录次数： <strong><%=UserHS.LoginNumber%></strong> 次</td></tr>
            <tr><td nowrap="nowrap">【<a href="<%=ACTCMS.ActCMSDM%>User/index.asp" target="_blank">会员中心</a>】【<a href="<%=ACTCMS.ActCMSDM%>User/Checklogin.asp?Action=LoginOut">退出登录</a>】</td></tr>
			</table>
<%End IF
  End Sub  
 Sub js()


 If CBool(UserHS.UserLoginChecked) = False Then
 %>
document.writeln("              <form name=\"userlogin\" action=\"\<%= ACTCMS.ActSys %>User\/Checklogin.asp?Action=LoginCheck\" method=\"POST\">");
document.writeln("				<div class=\"login_kk\">帐号：<input name=\"UserName\" type=\"text\" class=\"login_input\" \/><\/div>");
document.writeln("				<div class=\"login_kk\">密码：<input name=\"PassWord\" type=\"PassWord\" class=\"login_input\" \/><\/div>");
   
  <% if ACTCMS.ActCMS_Sys(15) = 0 Then%>
document.writeln("				<div class=\"login_kk_yzm\"><samp><img src=\"..\/act_inc\/code.asp?s=\'+Math.random();\" alt=\"验证码\" title=\"看不清楚? 换一张！\" style=\"cursor:hand;\" onClick=\"src=\'..\/act_inc\/code.asp?s=\'+Math.random()\"\/><\/samp>验证码：<input name=\"Code\" type=\"text\" class=\"login_yzm\" \/><\/div>");
 <%end if %>
document.writeln("				<div class=\"login_kk\"><input name=\"\" type=\"image\" src=\"images\/login.gif\" \/>");
document.writeln(" 				&nbsp;&nbsp;&nbsp;&nbsp;<a href=\"\<%= ACTCMS.ActCMSDM %>User\/Reg.asp\"><img   type=\"image\" src=\"images\/reg.gif\" \/><\/a><\/div>");
document.writeln("                <\/form>");
<%
Else




	Dim face
	If Trim(UserHS.myface)<>"" Then 
		face= UserHS.myface
	Else 
		face= ACTCMS.ActSys&"user\/images\/nophoto.gif" 
	End If 
%>

document.writeln("<div class=\"userinfo\">");
document.writeln("    <div class=\"welcome\">你好：<strong><%=UserHS.UserName%><\/strong>，欢迎登录 <\/div>");
document.writeln("    <div class=\"userface\">");
document.writeln("        <a href=\"\<%= ACTCMS.ActSys %>User\/index.asp\"><img src=\"\<%=face%>\" width=\"52\" height=\"52\" \/><\/a>");
document.writeln("    <\/div>");
document.writeln("    <div class=\"mylink\">");
document.writeln("        <ul>");
 document.writeln("            <li><a href=\"\<%= ACTCMS.ActSys %>User\/ACT.manage.asp\">发表文章<\/a><\/li>");
document.writeln("            <li><a href=\"\<%= ACTCMS.ActSys %>User\/Friend.asp\">好友管理<\/a><\/li>");
document.writeln("            <li><a href=\"\<%= ACTCMS.ActSys %>User\/Comment.asp\">我的评论<\/a><\/li>");
document.writeln("            <li><a href=\"\<%= ACTCMS.ActSys %>User\/search.asp\">查找好友<\/a><\/li>");
document.writeln("        <\/ul>");
document.writeln("    <\/div>");
document.writeln("    <div class=\"uclink\">");
document.writeln("        <a href=\"\<%= ACTCMS.ActSys %>User\/index.asp\">会员中心<\/a> | ");
document.writeln("        <a href=\"\<%= ACTCMS.ActSys %>User\/edit.asp\">资料<\/a> | ");
document.writeln("        <a href=\"<%= ACTCMS.ActSys %>space\/?<%=actcms.ACT_U(UserHS.UModeID,5)%>-<%=UserHS.userid%>\">空间<\/a> | ");
document.writeln("        <a href=\"\<%= ACTCMS.ActSys %>User\/Checklogin.asp?Action=LoginOut\">退出登录<\/a> ");
document.writeln("    <\/div>");
document.writeln("<\/div>")
<%

End If 

End Sub %>