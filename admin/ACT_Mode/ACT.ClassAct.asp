<!--#include file="../ACT.Function.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>内容添加 By Act</title>
<link href="../Images/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
 Dim Action,obj_news_rs,ShowErr,ModeID
 ModeID = ChkNumeric(Request("ModeID"))
 if ModeID=0 or ModeID="" Then ModeID=1
 Action = Request("Action")
 Select Case Action
		Case "one"
			Call OrderOne()
		Case "Order_one"
			Call UpdateOrderID()
 End Select 
 Sub OrderOne()
%>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
	<tr class="td_bg">
		<td class="bg_tr">栏目管理┆<a href="#" target="_blank" style="cursor:help;'" ><strong class="bg_tr">帮助</strong></a></td>
	</tr>
	<tr>
		<td height="18" ><a href="ACT.Class.asp?ModeID=<%=ModeID %>">管理首页</a>┆<a href="ACT.ClassAdd.asp?ModeID=<%=ModeID %>&Action=add">添加根栏目</a>┆<a href="ACT.ClassAct.asp?Action=one&ModeID=<%=ModeID %>">栏目排序</a></td>
	</tr>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr class="td_bg"> 
    <td height="22" class="bg_tr">栏目名称</td>
    <td height="22" class="bg_tr"><div align="center">ID</div></td>
    <td class="bg_tr"><div align="center">操作</div></td>
  </tr>
  <%
	Set obj_news_rs = server.CreateObject("Adodb.Recordset")
	Dim Classid:Classid=request("classid")
	If Classid = "" then
	obj_news_rs.Open "Select Orderid,id,ClassID,ParentID,ClassName from Class_ACT where Parentid  = '0'    Order by Orderid asc,ID asc",Conn,1,3
	Else
	obj_news_rs.Open "Select Orderid,id,ClassID,ParentID,ClassName from Class_ACT where     Parentid= '"&Classid&"' Order by Orderid asc,ID asc",Conn,1,3
	End if
	Do while Not obj_news_rs.eof 
	%>
  <form name="ClassForm" method="post" action="ACT.ClassAct.asp?ModeID=<%=ModeID%>">
    <tr class="hback"> 
      <td width="39%" height="31" class="td_bg"><img src="../Images/+.gif" width="15" height="15" /> 
      <a href= "?Action=one&ModeID=<%=ModeID%>&classid=<% = obj_news_rs("Classid") %>"> <b><% = obj_news_rs("ClassName") %></b></a> </td>
      <td width="21%" class="td_bg"><div align="center"> 
          <% = obj_news_rs("ID") %>
      </div></td>
      <td width="40%" class="td_bg"><div align="center"> 
          <input name="OrderID" type="text" id="OrderID" value="<% = obj_news_rs("OrderID") %>" size="4" maxlength="3">
          <input name="ClassID" type="hidden" id="ClassID" value="<% = obj_news_rs("ClassID") %>">
          <input name="Action" type="hidden" id="ClassID" value="Order_one">
          <input type="submit" Class="ACT_BTN" name="Submit" value="更新权重(排列序号)">
      </div></td>
    </tr>
  </form>
  <%
		obj_news_rs.MoveNext
	loop
%>
</table>
<table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" class="table">
  <tr>
    <td class="td_bg">说明：权重(排列序号)数字越小排得越靠前.如果权重(排列序号)数字相同，就根据ID来排列</td>
  </tr>
</table>
<%
obj_news_rs.close
set obj_news_rs =nothing
End Sub
Sub  UpdateOrderID()
		Dim ClassID,OrderID
		ClassID = Request.Form("ClassID")
		OrderID = Request.Form("OrderID")
		if ClassID="" then
 			Call Actcms.ActErr("错误参数","","")
 			Response.end
		else
			ClassID=ClassID
		end if
		if isnumeric(OrderID)= false then
 			Call Actcms.ActErr("错误参数:排列序号请填写正确的数字","","")
 			Response.end
		End if
		if OrderID="" then
  			Call Actcms.ActErr("错误参数:OrderID","","")
			Response.end
		end if
		Conn.execute "update Class_ACT set OrderID=" & OrderID & " where ClassID='" & ClassID &"'"
		Response.Redirect "ACT.ClassAct.asp?Action=one&ModeID="&ModeID
End sub
%>
