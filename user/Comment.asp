<!--#include file="../act_inc/ACT.User.asp"-->
 <!--#include file="../ACT_INC/cls_pageview.asp"-->
<% 
	 dim  UserHS
    	Set UserHS = New ACT_User
	IF Cbool(UserHS.UserLoginChecked)=false then
	  Response.Write "<script>top.location.href ='login.asp' ;</script>"
	  Response.end
	End If	
 %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>会员中心</title>
<script language="JavaScript" src="main.js"></script>
 <link href="images/css.css" rel="stylesheet" type="text/css" />
 </head>
<body style="background-color:#fff">
<div id="head">
  <div id="logo"><a href="index.asp" alt=""><img src="images/logo_member.gif" alt="actcms"></a></div><div id="banner"></div>
</div>
<div id="membermenu">
<!--#include file="menu.asp"-->

</div>
 
<div id="main">
 <div id="left">
  <div id="treemenu">
    <h5>基本设置</h5>
    <div style="text-align:center;">
    <img src="<%If Trim(UserHS.myface)<>"" Then 
		response.write ACTCMS.ActSys&"UpFiles/User/"&UserHS.UModeID&"/"&UserHS.Userid&"."&UserHS.myface
	Else 
		response.write "images/nophoto.gif" 
	End If 
	
	%> " alt="actcms" height="150" width="150"/>
	</div>
    <table cellpadding="0" cellspacing="0" class="member_info">
    <tr>
      <th>用户名：</th><td><%=UserHS.username%></td>
    </tr>
    <tr>
      <th>用户组：</th><td><%=UserHS.G_Name%></td>
    </tr>
    <tr>
      <th valign="top">等　级：</th>
      <td></td></tr>
    </table>
    <ul>
       <li><a href="edit.asp">修改资料</a></li>
      <li><a href="editpwd.asp">修改密码</a></li>
    </ul>
  </div>
  <ol>
    <li class="local"><a href="<%= actcms.ActCMSDM%>">返回网站首页</a></li>
    <li class="exit"><a href="Checklogin.asp?Action=LoginOut">退出登录</a></li>
  </ol>
</div>
  <div id="right">

<p id="position"><strong>当前位置：</strong><a href="index.asp">会员中心</a> 查看评论 </p>
<div class="clear"></div>
<div class="clear"></div>

<table cellpadding="0" cellspacing="1" class="table_list">
<tr>
  <td bgcolor="#F7FCFF"><strong>评论选项：</strong><strong><a href="?">所有评论</a> ┆ <a href="?Type=Lock">已审核</a>┆ <a href="?Type=UnLock">未审核</a></strong></td>
</tr>
</table>
 <table cellpadding="0" cellspacing="1" class="table_list">
  <caption>管理信息</caption>
<tr>
<th width="45">ID</th>
<th width="460">评论内容</th>
<th width="70">发表时间</th>
<th width="120">审核与否</th>
</tr>
<% 	Dim strLocalUrl
	strLocalUrl = request.ServerVariables("SCRIPT_NAME")
	
	Dim intPageNow
	intPageNow = request.QueryString("page")
	
	Dim intPageSize, strPageInfo
	intPageSize = 20
	
	Dim arrRecordInfo, i
	Dim sql, sqlCount,Sqls,pages
	pages = "Type="&Request("Type")&"&page"
	Select Case Request.QueryString("Type")
		Case "Lock"
			Sqls = " Where  Locked = 1 and  UserID="& UserHS.UserID 
		Case "UnLock"
			Sqls = " Where  Locked = 0  and  UserID="& UserHS.UserID 
		Case Else
			Sqls = " Where  UserID="& UserHS.UserID 
	End Select
	sql = "SELECT [ID], [ModeID], [Content], [AddDate],[Locked],[UserIP],[ClassID],[acticleID]" & _
		" FROM [Comment_Act] " &Sqls& _
		"   ORDER BY [ID] DESC"
 	sqlCount = "SELECT Count([ID])" & _
			" FROM [Comment_Act]  " &Sqls
		Dim clsRecordInfo
		Set clsRecordInfo = New Cls_PageView
			clsRecordInfo.intRecordCount = 2816
			clsRecordInfo.strSqlCount = sqlCount
			clsRecordInfo.strSql = sql
			clsRecordInfo.intPageSize = intPageSize
			clsRecordInfo.intPageNow = intPageNow
			clsRecordInfo.strPageUrl = strLocalUrl
			clsRecordInfo.strPageVar = pages
		clsRecordInfo.objConn = Conn		
		arrRecordInfo = clsRecordInfo.arrRecordInfo
		strPageInfo = clsRecordInfo.strPageInfo
		Set clsRecordInfo = nothing
 		If IsArray(arrRecordInfo) Then
			For i = 0 to UBound(arrRecordInfo, 2)
 	%>
    <tr>
    <td align="center"><%= arrRecordInfo(1,i) %></td>
<td class="align_left"> 
<%= arrRecordInfo(2,i) %> </td>
<td class="align_c"><%= arrRecordInfo(3,i) %></td>
<td class="align_c">
<% IF arrRecordInfo(4,i) = 1 Then Response.Write "<font color=red>&nbsp;&nbsp;已审核&nbsp;&nbsp;</font>" Else Response.Write "<font color=#0000FF>&nbsp;&nbsp;未审核&nbsp;&nbsp;</font>"%>	
          </td>
</tr>

<% 
	Next
	End If
	%>
     </table>
    <div class="button_box">
        </div>
 <div id="pages">
<%= strPageInfo%>
 </div>


  </div>
</div>
 
<!--#include file="foot.asp"-->
</body>
</html>