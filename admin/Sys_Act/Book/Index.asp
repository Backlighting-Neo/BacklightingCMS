<!--#include file="../../ACT.Function.asp"-->
<!--#include file="../../../act_inc/cls_pageview.asp"-->
<style>
.tb {word-break:break-all}
#notice {
	BORDER-RIGHT: #f4d738 2px solid;
	PADDING-RIGHT: 0px;
	BORDER-TOP: #f4d738 2px solid;
	PADDING-LEFT: 35px;
	FONT-WEIGHT: bold;
	PADDING-BOTTOM: 0px;
	BORDER-LEFT: #f4d738 2px solid;
	COLOR: #c00;
	LINE-HEIGHT: 30px;
	PADDING-TOP: 0px;
	BORDER-BOTTOM: #f4d738 2px solid;
	HEIGHT: 30px;
	background-color: #fffdaa;
	background-repeat: no-repeat;
	background-position: 8px 50%;
	width: 694px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 5px;
	margin-left: 0px;
}
</style>
<%
Server.ScriptTimeOut=999
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache"
response.Charset = "utf-8"
ConnectionDatabase
If Not ACTCMS.ACTCMS_QXYZ(0,"lyxt_ACT","") Then   Call Actcms.Alert("对不起，你没有操作权限！","") 
Dim Rs1,ActCMS_BookSetting
	Set Rs1=Conn.Execute("Select ActCMS_SysSetting From Config_Act")
	ActCMS_BookSetting=Split(Rs1(0),"^@$@^")%>
<head>
<title>网站留言本首页-By ActCMS.Com</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="keywords" content="ACT内容管理系统、免费开源的ASP网站新闻发布系统、免费网站管理系统、免费CMS系统、网站设计、搜索引擎排名研究、网站经营理念">
<meta name="description" content="开源CMS,网站管理系统,cms,新闻发布系统,内容管理系统,免费网站管理系统,免费CMS系统,ASP网站开发,网站优化,ASP网站管理系统 ">
<SCRIPT>function showimage(){document.images.tus.src="../../../Plus/Book/face/"+document.form.xq.options[document.form.xq.selectedIndex].value+".gif";}</SCRIPT>
<link href="../../../Plus/Book/face/style.css" rel="stylesheet" type="text/css">
</head>
<body>
  <script language="JavaScript" type="text/JavaScript">
function chk(a)
{
a.submit.value="提交数据中,请稍等...";
a.submit.disabled=true;
a.Submit.disabled=true;
}
</script>
<TABLE style="BORDER-COLLAPSE: collapse" borderColor=#111111 cellSpacing=0 
cellPadding=0 width=691 bgColor=#ffffff border=0>
  <TBODY>
    <TR> 
      <TD bgColor=#ebebeb colSpan=2 height=2></TD>
    </TR>
    <TR> 
      <TD height=23 colspan="2" vAlign=bottom noWrap>&nbsp; 您的位置：<A 
      href="<%= Actcms.ActCMSDM %>" target=_blank class="t999"><%= Actcms.ActCMS_Sys(0) %></A>&gt;&gt; 
        <A 
      href="<%= Actcms.ActCMSDM %>Plus/Book/index.asp" 
      target=_self class="t999">留言本</A>&gt;&gt;</TD>
    </TR>
  <CENTER>
    <TR> 
      <TD>&nbsp; </TD>
    </TR>
  </center>
</TABLE>
<%
Dim intDateStart
	Dim strLocalUrl
	strLocalUrl = request.ServerVariables("SCRIPT_NAME")
	Dim intPageNow
	intPageNow = request.QueryString("page")
	Dim intPageSize, strPageInfo
	intPageSize =20
	Dim arrRecordInfo, i,sql,sqlCount
	sql = "SELECT [show],[name],[qq],[mail],[url],[xq],[nr],[hf],[ip],[addtime],[id],[sh]" & _
		" FROM [Book_ACT] " & _
		 " ORDER BY [addtime] deSC"
	sqlCount = "SELECT Count([ID])" & _
			" FROM [Book_ACT]"
		Dim clsRecordInfo
		Set clsRecordInfo = New Cls_PageView
			clsRecordInfo.intRecordCount = 2816
			clsRecordInfo.strSqlCount = sqlCount
			clsRecordInfo.strSql = sql
			clsRecordInfo.intPageSize = intPageSize
			clsRecordInfo.intPageNow = intPageNow
			clsRecordInfo.strPageUrl = strLocalUrl
			clsRecordInfo.strPageVar = "page"
			clsRecordInfo.objConn = Conn		
			arrRecordInfo = clsRecordInfo.arrRecordInfo
			strPageInfo = clsRecordInfo.strPageInfo
			Set clsRecordInfo = nothing
			conn.close			
				If IsArray(arrRecordInfo) Then
					For i = 0 to UBound(arrRecordInfo, 2)		
					%><table width="694" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#9FBCE3" >
  <tr> 
    <td width="80" align="center" valign="top"  class="tb"><img src="../../../Plus/Book/face/<%= arrRecordInfo(5,i) %>.gif" width="80" height="110"><br><%= arrRecordInfo(1,i)%></td>
    <td height="127">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="10">&nbsp;</td>
          <td><table width="100%" height="26%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolor="#9FBCE3" style="border-collapse:collapse">
              <tr> 
                <td height="126"> <form name="form1" method="post" action="Rel_Act.asp?action=hf&id=<%= arrRecordInfo(10,i) %>">
                  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" class="tb">
                    <tr> 
                      <td width="15">&nbsp;</td>
                      <td width="579" height="20">第<font color="#993333"><%= arrRecordInfo(10,i) %></font>条留言&nbsp;&nbsp;
					 QQ：
					   <% if arrRecordInfo(2,i) <> "" then %>
					  <%= arrRecordInfo(2,i) %>
					 <% 
					 else
					 response.Write("无")
					  End If %>
                        &nbsp;
                        <% if arrRecordInfo(3,i) <> "" then %><a title="<%= "点击这里给"&arrRecordInfo(1,i)&"发送邮件" %>" href="mailto:<%= arrRecordInfo(3,i) %> "><img src="../../../Plus/Book/face/email.gif"  border=0></a>
					 <% 
					 else%>
					<img src="../../../Plus/Book/face/email1.gif"  border=0>
					<% End If %> &nbsp;
					<% if arrRecordInfo(4,i) <> "" then %><a target="_blank" title="请浏览我的主页" href="<%= arrRecordInfo(4,i) %> "><img src="../../../Plus/Book/face/home.gif"  border=0></a>
					 <% 
					 else%>
					<img src="../../../Plus/Book/face/home.gif"  border=0>
					<% End If %> 
					&nbsp;&nbsp;<a  href="Rel_Act.asp?action=sh&id=<%=arrRecordInfo(10,i) %> ">审核</a>
					 <a  href="Rel_Act.asp?action=del&id=<%=arrRecordInfo(10,i) %> " onClick="return confirm('是否确认删除,该操作不可恢复!')">删除</a>
					 <% if arrRecordInfo(11,i)=1 then response.Write "<font color=""#ff8000"">该留言处于审核状态...</font>" %>
					 </td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td valign="top" > 
                        <%
					  if arrRecordInfo(0,i)="0" then
							 response.Write  "<br>&nbsp;&nbsp;&nbsp;&nbsp;"&arrRecordInfo(6,i)
					  else
					  response.Write "<font color=""#ff8000"">给管理员的悄悄话...</font><br>"&arrRecordInfo(6,i)
					  end if %>
                        <br>
                        <br>
                        <% if arrRecordInfo(7,i) <> "" then%><IMG 
            src="../../../Plus/Book/face/dot.gif" width="21" height="10" border=0><font color="red">管理员回复</font><FONT color=#ff8000 >：<%= arrRecordInfo(7,i) %></FONT> 
                  <% End If %>                      </td>
                    </tr>
                    <tr>
                      <td>&nbsp;</td>
                      <td height="20">
<div align="right">IP:<%=arrRecordInfo(8,i)%> 　　<font color="#CCCCCC"><%=arrRecordInfo(9,i)%> </font></div></td>
                    </tr>
                    <tr>
                      <td>&nbsp;</td>
                      <td height="20">
						<div align="center">
                            <textarea name="hf" cols="72" rows="5" class="gbinput" id="hf" ><%=arrRecordInfo(7,i)%></textarea><br>
                            <input name="Submit" type="submit" class="input2" value=" 提交 ">
                            &nbsp;&nbsp;<input name="Submit2" type="reset" class="input2" value=" 重置 ">
                            支持HTML回复</div>
					  </td>
                    </tr>
                  </table></form></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
</table> <br>
 <%Next
		else
		response.Write "还没有任何留言<br>"
 End IF	
 response.Write strPageInfo
   %> <br>
</body>
</html>
