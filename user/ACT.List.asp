<!--#include file="../act_inc/ACT.User.asp"-->
 <!--#include file="../ACT_INC/cls_pageview.asp"-->
<% 
	 dim  ClassID,sh,shname,UserHS,ModeID,shset,ID
	 ClassID=RSQL(ACTCMS.G("ClassID"))
	 sh=ChkNumeric(actcms.s("sh"))
 	ModeID=ChkNumeric(actcms.ACT_L(ClassID,10))
	if ModeID="0" then response.Redirect "ACT.manage.asp"
	Set UserHS = New ACT_User
	IF Cbool(UserHS.UserLoginChecked)=false then
	  Response.Write "<script>top.location.href ='login.asp' ;</script>"
	  Response.end
	End If	
	   Select Case sh
	  			Case 0
					shname="<font color=green>审核通过</font>"	
				Case 1
					shname= "<font color=red>草稿</font>"
				Case 2
					shname="<font color=red>待审核</font>"	
				Case 3
					shname= "<font color=red>退稿</font>"
				case 999
					shname= "<font color=red>已删除</font>"
		 End Select	
 	if sh="999" then shset=" delif=1  " else shset=" isAccept="&sh&" and delif=0 "
	
  if request("A")="Del" then 
    ID=ChkNumeric(actcms.s("ID"))
	If ID="0" Then Call ACTCMS.Alert("你没有选中要删除的内容!",""):Response.End
	Conn.Execute("Update  "&ACTCMS.ACT_C(ModeID,2)&" set delif=1 Where isAccept<>0 and UserID="& UserHS.UserID &"   and ID =" & ID )
	Response.Redirect "?ClassID="& ClassID &"&sh="& sh&""
  End if
	
	
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
      <h5>信息管理</h5>
      <ul>
        <li><a href="ACT.Add.asp?action=add&ClassID=<%= ClassID %>">发布信息</a></li>
        <li><a href="ACT.List.asp?ClassID=<%= ClassID %>&sh=0">审核通过</a></li>
        <li><a href="ACT.List.asp?ClassID=<%= ClassID %>&sh=1">草稿</a></li>
        <li><a href="ACT.List.asp?ClassID=<%= ClassID %>&sh=2">待审核</a></li>
        <li><a href="ACT.List.asp?ClassID=<%= ClassID %>&sh=3">被退稿</a></li>
        <li><a href="ACT.List.asp?ClassID=<%= ClassID %>&sh=999">已删除</a></li>
      </ul>
    </div>
    <ol>
    <li class="local"><a href="<%= actcms.ActCMSDM%>">返回网站首页</a></li>
    <li class="exit"><a href="Checklogin.asp?Action=LoginOut">退出登录</a></li>
    </ol>
  </div>
  <div id="right">

<p id="position"><strong>当前位置：</strong><a href="index.asp">会员中心</a><a href="ACT.manage.asp">信息管理</a><a href="?ClassID=<%= ClassID %>&sh=<%= sh %>"><%= actcms.ACT_L(ClassID,2) %></a>查看 <%= shname %> 的信息</p>
<div class="clear"></div>
<div class="clear"></div>


 <table cellpadding="0" cellspacing="1" class="table_list">
  <caption>管理信息</caption>
<tr>
<th width="45">ID</th>
<th width="460">标题</th>
<th width="70">更新时间</th>
<th width="120">管理操作</th>
</tr>
<% 	


Dim strLocalUrl
	strLocalUrl = request.ServerVariables("SCRIPT_NAME")
	
	Dim intPageNow
	intPageNow = request.QueryString("page")
	
	Dim intPageSize, strPageInfo
	intPageSize = 20
	
	Dim arrRecordInfo, i
	Dim sql, sqlCount,Sqls,pages
	pages = "ClassID="&ClassID&"&sh="&sh&"&page"
    Sqls = " where "&shset&" and userid="& UserHS.userid &" "
  	sql = "SELECT [ClassID],[ID],[actlink],[FileName],[InfoPurview],[ReadPoint],[title],[updatetime]" & _
		" FROM ["&ACTCMS.ACT_C(ModeID,2)&"]" &Sqls& _
		"ORDER BY [ID] DESC"
 	sqlCount = "SELECT Count([ID])" & _
			" FROM ["&ACTCMS.ACT_C(ModeID,2)&"]"&Sqls
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
<td class="align_left"><a  href="<%= ACTCMS.GetInfoUrl(ModeID,arrRecordInfo(0,i),arrRecordInfo(1,i),arrRecordInfo(2,i),arrRecordInfo(3,i),arrRecordInfo(4,i),arrRecordInfo(5,i)) %>" target="_blank">
<%= arrRecordInfo(6,i) %></a></td>
<td class="align_c"><%= arrRecordInfo(7,i) %></td>
<td class="align_c">
<% if sh="0" then  %>
    <a href="<%=ACTCMS.GetInfoUrl(ModeID,arrRecordInfo(0,i),arrRecordInfo(1,i),arrRecordInfo(2,i),arrRecordInfo(3,i),arrRecordInfo(4,i),arrRecordInfo(5,i)) %>" target="_blank">预览</a> 

 <% else  %>
  <a href="ACT.Add.asp?action=edit&ClassID=<%= arrRecordInfo(0,i) %>&ID=<%= arrRecordInfo(1,i) %>">修改</a> 
    | <a href="?A=Del&ID=<%= arrRecordInfo(1,i) %>&ClassID=<%= ClassID %>&sh=<%= sh %>"  onClick="return confirm('确认删除此文章吗?此操作不可恢复!')">删除</a> 
        
    <% end if  %>
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