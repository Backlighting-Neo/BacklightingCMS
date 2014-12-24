<!--#include file="../../ACT.Function.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head><meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ACTCMS来源</title>
<link href="../../Images/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td class="bg_tr">
	您现在的位置：系统中心 &gt;&gt; 插件管理 &gt;&gt;<a href="?"> 留言本管理</a></td>
  </tr>
</table>
<%	If Not ACTCMS.ACTCMS_QXYZ(0,"digg_act","") Then   Call Actcms.Alert("对不起，你没有操作权限！","") 
	Dim Rs,A,ActCMS_Digg,I,BookSetting
	Set Rs=server.CreateObject("adodb.recordset") 
	Rs.OPen "Select PlusConfig from Plus_ACT  where PlusID='digg_act'",Conn,1,3
	If request("A") ="save" Then 
		For I=0 To 2
			BookSetting=BookSetting& Replace(request.Form("ActCMS_Digg" & I &""),"^@$@^","") & "^@$@^"
		Next
		Rs("PlusConfig")=BookSetting
		Rs.Update:Application.Contents.RemoveAll
		Rs.Close:Set Rs = Nothing	
		Call Actcms.ActErr("修改成功","Sys_Act/Digg/Sys.asp","")
 	Else
	ActCMS_Digg=Split(Rs("PlusConfig"),"^@$@^")
%>
<table width="100%" border="0" cellpadding="1" cellspacing="1" class="table">
<form name="form1" method="post" action="?A=save">
	<tr>
      <td width="34%" height="30" align="left" >是否允许游客DIGG：</td>
	  <td width="66%" height="30" ><input <% IF ActCMS_Digg(0)  = "0" Then Response.Write "Checked" %>  type="radio"  id="radio3" name="ActCMS_Digg0" value="0">
	      <label for="radio3"><font color="green">允许 &nbsp;</font></label>
          <input  <% IF  ActCMS_Digg(0)  = "1" Then Response.Write "Checked" %>  id="radio4" type="radio" name="ActCMS_Digg0" value="1">
	    <label for="radio4"><font color="red">禁止 &nbsp;</font></label></td>
    </tr>
 
  <tr>
    <td colspan="2" align="center"><input type=button onclick=CheckForm() class="ACT_btn"  name=Submit1 value="  保存  " />
        &nbsp;&nbsp;&nbsp;&nbsp;<input name="Submit2" type="reset" class="ACT_btn" value="  重置  ">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
    </tr>
</form>
</table>

<%	End If 
CloseConn %>
<script language="javascript">
function OpenWindowAndSetValue(Url,Width,Height,WindowObj,SetObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:yes;help:no;scroll:yes;');
	if (ReturnStr!='') SetObj.value=ReturnStr;
	return ReturnStr;
}	
function CheckForm()
{ var form=document.form1;
		form.Submit1.value="正在提交数据,请稍等...";
		form.Submit1.disabled=true;
		form.Submit2.disabled=true;		
	    form.submit();
        return true;
	}</script> 

</body>
</html>
