<!--#include file="../act_inc/ACT.User.asp"-->
 <!--#include file="../ACT_INC/cls_pageview.asp"-->
<% 
	 dim  UserHS
    	Set UserHS = New ACT_User
	IF Cbool(UserHS.UserLoginChecked)=false then
	  Response.Write "<script>top.location.href ='login.asp' ;</script>"
	  Response.end
	End If	
	
		select case request("A")
		case "Del"
			call DelFriend()
		case "flag"
			call flag()
		Case "Add"
			Call addf()
   	end select
	
	sub addf()
			Dim UM,U,rs
			UM=ChkNumeric(request("UM"))
			U=ChkNumeric(request("U"))
			Set rs=actcms.actexe("select userid from User_ACT where   userid<>"&userhs.userid&"  and userid="&u)
			If rs.eof Then call actcms.Alert("没有找到该用户","Friend.asp"):response.End
			Set rs=actcms.actexe("select id from Friend_ACT where userid="&u&"    and u="&userhs.userid&"  ")
			If Not rs.eof Then call actcms.Alert("该用户已经是您的好友","Friend.asp"):response.End
 		Dim sql, rs2
		sql = "Select  * from Friend_ACT"
		Set rs2 = Server.CreateObject("Adodb.RecordSet")
		rs2.Open sql, Conn, 1, 3
		rs2.AddNew
		rs2("userid") = u
 		rs2("um") = userhs.umodeid
		rs2("u") = userhs.userid
		rs2("AddDate") =now
 		rs2.Update
		rs2.Close:Set rs2 = Nothing
		Call actcms.Alert("添加好友成功","Friend.asp")
 	end sub  
	
	sub flag()
 		Dim TG_ID:TG_ID =Request("ID")
 		IF TG_ID = "" Then
			response.Write "请先选定好友"
			response.End
		End IF		
		 TG_ID = Split(TG_ID,",")
		 For I = LBound(TG_ID) To UBound(TG_ID)
 				Conn.execute("Update Friend_ACT set flag=2 where U="& UserHS.UserID &"   and ID = "&ChkNumeric(TG_ID(i))&"")
 		 Next
		set conn=nothing
		response.Redirect("?Type=1")
	end sub  
	sub DelFriend()
		Dim TG_ID:TG_ID =Request("ID")
 		IF TG_ID = "" Then
			response.Write "请先选定好友"
			response.End
		End IF		
 		 TG_ID = Split(TG_ID,",")
		 For I = LBound(TG_ID) To UBound(TG_ID)
 				 Conn.execute("Delete from  Friend_ACT   where U="& UserHS.UserID &"   and  ID = "&ChkNumeric(TG_ID(i))&"")
 		 Next
		set conn=nothing
		response.Redirect("?")
 	end sub 
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
		response.write UserHS.myface
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

<p id="position"><strong>当前位置：</strong><a href="index.asp">会员中心</a> 好友管理 </p>
<div class="clear"></div>
<div class="clear"></div>

<table cellpadding="0" cellspacing="1" class="table_list">
<tr>
  <td bgcolor="#F7FCFF"><strong>评论选项：</strong><strong><a href="?Type=1">好友</a>┆ <a href="?Type=2">黑名单</a></strong>┆ <a href="search.asp">查找好友</a></strong></td>
</tr>
</table>
 <table cellpadding="0" cellspacing="1" class="table_list">
  <caption>管理信息</caption>
<tr>
<th width="45">ID</th>
<th width="200">用户名</th>
<th width="200">真实姓名</th>
<th width="260">操 作</th>
<th width="200">加入时间</th>
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
	pages = "Type="&Request("Type")&"&page"
	Select Case Request.QueryString("Type")
		Case "1"
			Sqls = " Where  flag = 1 and    U="& UserHS.UserID
		Case "2"
			Sqls = " Where  flag = 2  and     U="& UserHS.UserID
		Case Else
			Sqls = " Where       U="& UserHS.UserID
	End Select
	sql = "SELECT [ID], [Userid], [AddDate]" & _
		" FROM [Friend_ACT] " &Sqls& _
		"   ORDER BY [ID] DESC"
   	sqlCount = "SELECT Count([ID])" & _
			" FROM [Friend_ACT]  " &Sqls
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
   <form action="?" method="post" name="actcmsform" >
 <tr>
    <td align="center"> 
      <input type="checkbox" name="ID" id="checkbox" value=<%=arrRecordInfo(0,i)  %> /></td>
    <td class="align_left"><%
	
	dim rs,t:t=false
	set rs=actcms.actexe("select top 1 username,RealName from user_act where userid="&arrRecordInfo(1,i)) 
	if not rs.eof then response.Write ACTCMS.UserM(arrRecordInfo(1,i)):t=true
	
	
	%></td>
    <td class="align_left"><%if t=false then response.write "该用户已经不存在" else response.Write  rs("RealName") %></td>
    <td align="center" ><a href="send.asp?s=send&Touser=<%= rs("username") %>&U=<%= arrRecordInfo(1,i) %>">发送短信</a> <a href="?A=Del&ID=<%= arrRecordInfo(0,i) %>" onClick="return(confirm('确定删除该位好友吗？'))">移除</a></td>
<td class="align_c"><%= arrRecordInfo(2,i) %></td>
</tr>

<% 
	Next
	End If
	%>
  
  
   <tr>
    <td colspan="5">
    <label for="chkAll"><input name="ChkAll" type="checkbox" id="ChkAll" onClick="CheckAll(this.form)" value="checkbox">
		选中所有显示记录</label>&nbsp;
    
  <button class="button_style" onclick="ConvSta();" type="button" >加入黑名单</button>&nbsp;

     
                <button class="button_style" onclick="DelSel();" type="button" >删除</button>
     
    </td>
    </tr></form>
     </table>
<div class="button_box">
        </div>
 <div id="pages">
<%= strPageInfo%>
 </div>


  </div>
</div>
 <script type="text/javascript">
function CheckAll(form)
		  {  
		 for (var i=0;i<form.elements.length;i++)  
			{  
			   var e = actcmsform.elements[i];  
			   if (e.name != 'ChkAll'&&e.type=="checkbox")  
			   e.checked = actcmsform.ChkAll.checked;  
		   }  
	  }
		
function GetCheckfolderItem()
{
	var allSel='';
	if(document.actcmsform.ID.value) return document.actcmsform.ID.value;
	for(i=0;i<document.actcmsform.ID.length;i++)
	{
		if(document.actcmsform.ID[i].checked)
		{
			if(allSel=='')
			allSel=document.actcmsform.ID[i].value;
			else
			allSel=allSel+","+document.actcmsform.ID[i].value;
		}
	}
	return allSel;
}

function DelSel(ftype)
{
	var ID = GetCheckfolderItem();
	if(ID=='') {
		alert("你没选中任何好友！");
		return false;
	}
	if(window.confirm("你确定要删除这些好友么？"))
	{
		location = "?A=Del&ID="+ID+"";
	}
}

function ConvSta()
{
	var ID = GetCheckfolderItem();
	if(ID=='') {
		alert("你没选中任何好友！");
		return false;
	}
 	if(window.confirm("你确定要把这些好友么？"))
	{
		location = "?A=flag&ID="+ID+"&flag=2";
	}
}
</script>

<!--#include file="foot.asp"-->
 
</body>
</html>