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
    <h5>信息管理</h5>
    <ul>
        <li><a href="logmoney.asp">资金明细</a></li>
		<li><a href="logpoint.asp">点券明细</a></li>
		<li><a href="logedays.asp">有效期明细</a></li>
 		<li><a href="card.asp">充值卡充值 </a></li>
		<li><a href="exchange.asp?Action=Point">兑换<%=actcms.ActCMS_Sys(24)%></a></li>
		<li><a href="exchange.asp?Action=Edays">兑换有效期</a></li>
		<li><a href="exchange.asp?Action=Money"><%=actcms.ActCMS_Sys(24)%>兑换账户资金</a></li>
     </ul>
  </div>
  <ol>
    <li class="local"><a href="<%= actcms.ActCMSDM%>">返回网站首页</a></li>
    <li class="exit"><a href="Checklogin.asp?Action=LoginOut">退出登录</a></li>
  </ol>
</div>
  <div id="right">

<p id="position"><strong>当前位置：</strong><a href="index.asp">会员中心</a> 有效期明细  </p>
<div class="clear"></div>
<div class="clear"></div>

<table cellpadding="0" cellspacing="1" class="table_list">
<tr>
  <td bgcolor="#F7FCFF"><a href="logmoney.asp">资金明细</a>
				<a href="logpoint.asp">点券明细</a>
				<a href="logedays.asp">有效期明细</a>
 				</td>
</tr>
</table>
 <table cellpadding="0" cellspacing="1" class="table_list">
  <caption>查询我的有效期明细 &nbsp;<a href='?'><font color=red>·所有记录</font></a> ·<a href='?Flag=1'>收入记录[<%=conn.execute("select count(id) from Edays_ACT where Flag=1 and UserID=" & UserHS.UserID & "")(0)%>]</a> ·<a href='?Flag=2'>支出记录[<%=conn.execute("select count(id) from Edays_ACT where Flag=2 and UserID=" & UserHS.UserID & "")(0)%>]</a>
		  </caption>
<tr>
	<th width="80" height="25" align="center"><strong> 用户名</strong></th>
	<th width="180" height="25" align="center"><strong>消费时间</strong></th>
	<th width="113" align="center"><strong>IP地址</strong></th>
	<th width="71" height="25" align="center"><strong>收入天数</strong></th>
	<th width="74" align="center"><strong>支出天数</strong></th>
	<th width="55" height="25" align="center"><strong>摘要</strong></th>
	<th width="75" height="25" align="center"><strong> 操作员</strong></th>
	<th width="239" align="center"><strong>备注</strong></th>

</tr>
<% 
 
 	Dim strLocalUrl,InEdays,InPoint,OutEdays
	strLocalUrl = request.ServerVariables("SCRIPT_NAME")
	
	Dim intPageNow
	intPageNow = request.QueryString("page")
	
	Dim intPageSize, strPageInfo
	intPageSize = 20
	
	Dim arrRecordInfo, i
	Dim sql, sqlCount,Sqls,pages
	pages = "page"
 
 
 If ChkNumeric(ACTCMS.S("Flag"))=1 Or ChkNumeric(ACTCMS.S("Flag"))=2 Then
 	sql="Select  ID,UserID,AddDate,IP,Edays,Flag,Userlog,Descript From Edays_ACT Where Flag=" & ChkNumeric(ACTCMS.S("Flag")) & " And  UserID=" & UserHS.UserID & " order by id desc"
	Else
	  sql="Select ID,UserID,AddDate,IP,Edays,Flag,Userlog,Descript From Edays_ACT Where UserID=" & UserHS.UserID & " order by id desc"
	End if
 
     	sqlCount = "SELECT Count([ID])" & _
			" FROM [Edays_ACT]  Where UserID=" & UserHS.UserID
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
 

  <td class="splittd" width="80" align="center"><%=arrRecordInfo(1,i)%></td>
    <td class="splittd" align="center"><%=arrRecordInfo(2,i)%></td>
    <td class="splittd"align="center"><%=arrRecordInfo(3,i)%></td>
    <td class="splittd" align="right"><%if arrRecordInfo(5,I)=1 Then Response.Write arrRecordInfo(4,I) & "天":InEdays=InEdays+arrRecordInfo(4,I) ELSE Response.Write "-"%></td>
    <td class="splittd" align="right"><%if arrRecordInfo(5,I)=2 Then Response.Write arrRecordInfo(4,I) & "天":OutEdays=OutEdays+arrRecordInfo(4,I) ELSE Response.Write "-"%></td>
    <td class="splittd" align="center"><%if arrRecordInfo(5,I)=1 Then Response.Write "<font color=red>收入</font>" Else Response.Write "支出"%></td>
    <td class="splittd" align="center"><%=arrRecordInfo(6,i)%></td>
	<td class="splittd"><%=arrRecordInfo(7,i)%></td>
</tr>

<% 
	Next
	End If
	%>
  
  <tr class='tdbg' onMouseOut="this.className='tdbg'" onMouseOver="this.className='tdbgmouseover'">    <td class="splittd" colspan='3' align='right'>本页合计：</td>    <td class="splittd" align='right'><%=InEdays%>天</td>    <td class="splittd" align='right'><%=ChkNumeric(OutEdays)%>天</td>    <td class="splittd" colspan='4'>&nbsp;</td>  </tr> 
  <% Dim totalinEdays:totalinEdays=conn.execute("Select sum(Edays) From Edays_ACT where UserID=" & UserHS.UserID & " AND Flag=1")(0)
     Dim TotalOutEdays:TotalOutEdays=conn.execute("Select sum(Edays) From Edays_ACT where UserID=" & UserHS.UserID & " AND  Flag=2")(0)
	 If ChkNumeric(totalInEdays)=0 Then totalInEdays=0
	 If ChkNumeric(TotalOutEdays)=0 Then TotalOutEdays=0
  %>
    <tr class='tdbg' onMouseOut="this.className='tdbg'" onMouseOver="this.className='tdbgmouseover'">    <td class="splittd" colspan='3' align='right'>所有合计：</td>    <td class="splittd" align='right'><%=ChkNumeric(totalInEdays)%>天</td>    <td align='right'><%=ChkNumeric(totalOutEdays)%>天</td>    <td class="splittd" colspan='4' align='center'>累计还剩：<%=totalInEdays-totalOutEdays%>天</td>  </tr> 

   
   <tr>
    <td colspan="12">
<div id="pages">
<%= strPageInfo%>
 </div>     
    </td>
    </tr>   
    
    </form>
     </table>
 
<span id="toggle_pannel" style="display:none;"></span>
<div class="clear"></div>
</div>
  </div>  
<!--#include file="foot.asp"-->
 
</body>
</html>