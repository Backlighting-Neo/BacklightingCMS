<!--#include file="../../ACT.Function.asp"-->
<%
Response.Expires = 0
Response.Expiresabsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
  %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../../Images/style.css" rel="stylesheet" type="text/css">
<title>自定义字段</title>
 <script language="JavaScript" type="text/javascript" src="../../../ACT_inc/js/lhgcore/lhgcore.min.js"></script>
 	<script type="text/javascript">
		var DG = frameElement.lhgDG;
    	</script>
 </head>
<body>

<table  width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
   <tr>
    <td colspan="3" align="left" class="bg_tr">您现在的位置：自定义字段显示</td>
  </tr>
<% 	
  dim rs,ModeID,momdeids
	  ModeID = ChkNumeric(Request("ModeID"))
	  If modeid=0 Then modeid=1
	  if modeid=0 then momdeids=" " else momdeids=" and modeid="&modeid&" "
	  Set Rs =ACTCMS.ACTEXE("SELECT * FROM Table_ACT Where actcms=1  " & momdeids & " order by OrderID desc,ID Desc")
	 If Rs.EOF  Then
	 	Response.Write	"<tr><td colspan=""8"" align=""center"">没有记录</td></tr>"
	 Else
		Do While Not Rs.EOF	
			 %>

  <tr>
    <td width="200" height="23" align="right" class="tdclass"><%= actcms.act_c(rs("ModeID"),1)&"模型-"&Rs("title") %>：</td>
    <td height="23" colspan="2" align="left"  class="tdclass">
	 
 <input name="IntactTitle"  type="text"  class="Ainput" id="IntactTitle" value="#<%= Rs("FieldName") %>" size="20"  onclick="this.focus();this.select();window.clipboardData.setData('Text',this.value);document.getElementById('<%= Rs("id") %>').innerHTML='<font color=green>代码已复制到剪贴板</font>';return true;"  onblur="javascript:document.getElementById('<%= Rs("id") %>').innerHTML='请直接复制放到标签里';" /><span id="<%= Rs("id") %>">请直接复制放到标签里</span>
 </td>
  </tr>


  <% 
		Rs.movenext
		Loop
	End if	 %>

 
  


</table>
</body>
</html>
