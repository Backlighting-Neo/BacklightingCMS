<!--#include file="../act_inc/ACT.User.asp"-->
 <!--#include file="../ACT_INC/cls_pageview.asp"-->
<% 	Dim UserHS
	Set UserHS = New ACT_User
	IF Cbool(UserHS.UserLoginChecked)=false then
	  Response.Write "<script>top.location.href ='login.asp' ;</script>"
	  Response.end
	End If	
	
	if request("a")="del" then 
 		Dim TG_ID:TG_ID =Request("ID")
 		IF TG_ID = "" Then
			response.Write "请先选定消息"
			response.End
		End IF		
 		 TG_ID = Split(TG_ID,",")
		 For I = LBound(TG_ID) To UBound(TG_ID)
 				 Conn.execute("Delete from  Message_Act   where UserID="& UserHS.UserID &"  and  ID = "&ChkNumeric(TG_ID(i))&"")
 		 Next
		set conn=nothing
	end if 
 %>
 <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>会员中心</title>
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
      <h5>短消息</h5>
      <ul>   	
		<li><a href="send.asp">发送短消息</a></li>
        <li><a href="inbox.asp">收件箱</a></li>
        <li><a href="outbox.asp">发件箱</a></li>
      </ul>
    </div>
    <ol>
    <li class="local"><a href="<%= actcms.ActCMSDM%>">返回网站首页</a></li>
    <li class="exit"><a href="Checklogin.asp?Action=LoginOut">退出登录</a></li>
    </ol>
  </div>
  <div id="right">
<p id="position"><strong>当前位置：</strong><a href="inbox.asp">短消息</a>收件箱</p>
        <form name="actform"   method="post" action="?a=del">
    	 <table cellpadding="0" cellspacing="1" class="table_list">
         <caption>收件箱</caption>
        <tr>
        	<th width="5%">选中</th>
         	<th width="*">标题</th>
<th width="10%">发件人</th>
           	<th width="10%">大小</th>
<th width="20%">发送时间</th>
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
	pages = "page"
   	sql = "SELECT [id],[title],[U],[content],[SendTime]" & _
		" FROM [Message_Act] where Flag=1 and   UserID="& UserHS.UserID & _
		" order by flag,SendTime desc"
 	sqlCount = "SELECT Count([ID])" & _
			" FROM [Message_Act]  where  Flag=1 and    UserID="& UserHS.UserID
		
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
        	<td>
            <input type="checkbox" name="ID" value="<%=arrRecordInfo(0,i)%>">
            </td>
            <td  class="align_left">
            <a href="read.asp?id=<%=arrRecordInfo(0,i)%>"><%=server.htmlencode(arrRecordInfo(1,i))%></a></td>
<td align="center"><%=ACTCMS.UserM(arrRecordInfo(2,i))%></td>
            <td align="center"><%=len(arrRecordInfo(3,i))%>Byte</td>
<td align="center"><%=formatdatetime(arrRecordInfo(4,i),2)%></td>
        </tr>
        
       <% Next
	End If %> 
        
        <tr>
<td colspan="5" style="text-align:left;">
 
<label for=chk>
		<input id="chk" type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form)">&nbsp;选择全部</label>
 &nbsp;
&nbsp;
<input class="button_style" type=submit name=action onClick="{if(confirm('确认批量删除这些消息吗?')){this.document.actform.submit();return true;}return false;}" value="删除">&nbsp
 </td>
</tr>
</table>
</form>
<script   language="JavaScript">

 	function CheckAll(form)  
  {  
 for (var i=0;i<form.elements.length;i++)  
    {  
    var e = actform.elements[i];  
   if (e.name != 'chkall')  
      e.checked = actform.chkall.checked;  
   }  
  }

</script>


<div id="pages">
<%= strPageInfo%></div>
</div>
</div>
<div class="clear"></div>
<div id="toogle_panel"></div>
<!--#include file="foot.asp"-->
</body>
</html> 