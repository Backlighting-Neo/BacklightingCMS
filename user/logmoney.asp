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

<p id="position"><strong>当前位置：</strong><a href="index.asp">会员中心</a> 资金明细 </p>
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
  <caption>查询我的资金明细&nbsp;<a href='?'><font color=red>所有记录</font></a> <a href='?IncomeFlag=1'>收入记录[<%=conn.execute("select count(id) from Money_Log_ACT where IncomeFlag=1 and UserID=" & UserHS.UserID & "")(0)%>]</a> <a href='?IncomeFlag=2'>支出记录[<%=conn.execute("select count(id) from Money_Log_ACT where IncomeFlag=2 and UserID=" & UserHS.UserID & "")(0)%>]</a>
		  </caption>
<tr>
<th width=160 height="25">交易时间</th>
<th width=108>用户名</th>
<th width=85>客户姓名</th>
<th width=95>交易方式</th>
<th width=55>币种</th>
<th width=65>收入金额</th>
<th width=65>支出金额</th>
<th width=45>摘要</th>
<th width=45>余额</th>
<th width="170">备注/说明</th>
</tr>
<% 
 
 	Dim strLocalUrl,intotalmoney,outtotalmoney
	strLocalUrl = request.ServerVariables("SCRIPT_NAME")
	
	Dim intPageNow
	intPageNow = request.QueryString("page")
	
	Dim intPageSize, strPageInfo
	intPageSize = 20
	
	Dim arrRecordInfo, i
	Dim sql, sqlCount,Sqls,pages
	pages = "page"
 
 
 If ChkNumeric(ACTCMS.S("IncomeFlag"))=1 Or ChkNumeric(ACTCMS.S("IncomeFlag"))=2 Then
 	sql="Select ID,LogTime,UserID,clientname,MoneyType,[money],IncomeFlag,CurrMoney,Remark From Money_Log_ACT Where IncomeFlag=" & ChkNumeric(ACTCMS.S("IncomeFlag")) & " And  UserID=" & UserHS.UserID & " order by id desc"
	Else
	  sql="Select ID,LogTime,UserID,clientname,MoneyType,[money],IncomeFlag,CurrMoney,Remark From Money_Log_ACT Where UserID=" & UserHS.UserID & " order by id desc"
	End if
 
 
    	sqlCount = "SELECT Count([ID])" & _
			" FROM [Money_Log_ACT]  Where UserID=" & UserHS.UserID
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
      <%= arrRecordInfo(1,i) %>
      
     </td>
    <td align="center"> 
       <%= ACTCMS.UserM(arrRecordInfo(2,i)) %>
      
     </td>
    <td align="center"> 
      <%= arrRecordInfo(3,i) %>
      
     </td>
    <td align="center"> 
      <%
	  Select Case arrRecordInfo(4,i)
	      Case 1:Response.WRite "现金"
		  Case 2:Response.Write "银行汇款"
		  Case 3:Response.Write "在线支付"
		  Case 4:Response.Write "资金余额"
		 End Select %>
      
     </td>
    <td align="center"> 
     
      人民币
     </td>
    <td align="center"> 
      <% 
	  
	  If arrRecordInfo(6,i) =1 Then
	     Response.Write formatnumber(arrRecordInfo(5,i) ,2,-1)
		 intotalmoney=intotalmoney+arrRecordInfo(5,i) 
	    End If
	  
	  %>
      
     </td>
    <td align="center"> 
         <% 
	  
	  If arrRecordInfo(6,i) =2 Then
	     Response.Write formatnumber(arrRecordInfo(5,i) ,2,-1)
		 outtotalmoney=outtotalmoney+arrRecordInfo(5,i) 
	    End If
	  
	  %>
      
     </td>
    <td align="center"> 
      
       <% If arrRecordInfo(6,i)=1 Then
	      Response.Write "<font color=red>收入</font>"
		 Else
		  Response.Write "<font color=green>支出</font>"
		 End If
		 %>
     </td>
    <td align="center"> 
      <%=formatnumber(arrRecordInfo(7,i),2,-1)%>
      
     </td>
    <td align="center"> 
      <%= arrRecordInfo(8,i) %>
      
     </td>
</tr>

<% 
	Next
	End If
	%>
  
  
 <tr>
      <td   align=right colSpan=5>本页合计：</td>
      <td  align=right><%=formatnumber(intotalmoney,2,-1)%></td>
      <td align=right><%=formatnumber(outtotalmoney,2,-1)%></td>
      <td colSpan=3>&nbsp;</td>
    </tr>
    
  
 <tr class=tdbg onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
      <td class="splittd" align=right colSpan=5>总计金额：</td>
	  	  <%intotalmoney=Conn.execute("Select Sum(Money) From Money_Log_ACT Where UserID=" & UserHS.UserID & " And IncomeFlag=1")(0)
	    outtotalmoney=Conn.execute("Select Sum(Money) From Money_Log_ACT Where UserID=" & UserHS.UserID & " And IncomeFlag=2")(0)
	    if not isnumeric(intotalmoney) then intotalmoney=0
		if not isnumeric(outtotalmoney) then outtotalmoney=0
	  %>
      <td class="splittd" align=right><%=formatnumber(intotalmoney,2,-1)%></td>
      <td class="splittd" align=right><%=formatnumber(outtotalmoney,2,-1)%></td>
      <td class="splittd" align=middle colSpan=3>资金余额：<%=formatnumber(UserHS.Money,2,-1)%></td>

    </tr>  
  
  
   <tr>
    <td colspan="11">
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