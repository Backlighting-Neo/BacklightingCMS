<!--#include file="../../ACT.Function.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ACTCMS来源</title>
<link href="../../Images/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td class="bg_tr">您现在的位置：系统中心 &gt;&gt; 插件管理 &gt;&gt;<a href="?">留言本管理</a></td>
  </tr>
</table>
<%
	If Not ACTCMS.ACTCMS_QXYZ(0,"lyxt_ACT","") Then   Call Actcms.Alert("对不起，你没有操作权限！","") 


Dim Rs,A,ActCMS_Book,I,BookSetting
Set Rs=server.CreateObject("adodb.recordset") 
Rs.OPen "Select PlusConfig from Plus_ACT  where PlusID='lyxt_ACT'",Conn,1,3

	If request("A") ="save" Then 
		For I=0 To 7
			BookSetting=BookSetting& Replace(request.Form("ActCMS_Book" & I &""),"^@$@^","") & "^@$@^"
		Next

		Rs("PlusConfig")=BookSetting
		Rs.Update:Application.Contents.RemoveAll
		Rs.Close:Set Rs = Nothing	
		Call Actcms.ActErr("修改成功","Sys_Act/Book/Sys.asp","")
 	Else
	ActCMS_Book=Split(Rs("PlusConfig"),"^@$@^")
%>
<table width="100%" border="0" cellpadding="1" cellspacing="1" class="table">
<form name="form1" method="post" action="?A=save">

	<tr>
      <td align="left" >留言本状态：</td>
	  <td ><input <% IF  ActCMS_Book(0)   = "0" Then Response.Write "Checked" %>  type="radio"  id="radio" name="ActCMS_Book0" value="0">
	      <label for="radio"><font color="green">正常 &nbsp;</font></label>
          <input  <% IF  ActCMS_Book(0)  = "1" Then Response.Write "Checked" %>  id="radio2" type="radio" name="ActCMS_Book0" value="1">
	    <label for="radio2"><font color="red">关闭 &nbsp;</font></label></td>
    </tr>
	<tr>
      <td align="left" >留言脏话过滤：</td>
	  <td ><input <% IF ActCMS_Book(1)  = "0" Then Response.Write "Checked" %>  type="radio"  id="radio3" name="ActCMS_Book1" value="0">
	      <label for="radio3"><font color="green">启用 &nbsp;</font></label>
          <input  <% IF  ActCMS_Book(1)  = "1" Then Response.Write "Checked" %>  id="radio4" type="radio" name="ActCMS_Book1" value="1">
	    <label for="radio4"><font color="red">禁用 &nbsp;</font></label></td>
    </tr>
	<tr>
      <td align="left" >留言审核功能：</td>
	  <td ><input <% IF ActCMS_Book(2)  = "0" Then Response.Write "Checked" %>  type="radio"  id="radio5" name="ActCMS_Book2" value="0">
	      <label for="radio5"><font color="green">启用 &nbsp;</font></label>
          <input  <% IF ActCMS_Book(2) = "1" Then Response.Write "Checked" %>  id="radio6" type="radio" name="ActCMS_Book2" value="1">
	    <label for="radio6"><font color="red">禁用 &nbsp;</font></label>
	    如禁用留言审核功能，则新留言将直接发表</td>
    </tr>
	<tr>
      <td align="left" >留言验证码：</td>
	  <td ><input <% IF ActCMS_Book(3)  = "0" Then Response.Write "Checked" %>  type="radio"  id="radio7" name="ActCMS_Book3" value="0">
	      <label for="radio7"><font color="green">启用 &nbsp;</font></label>
          <input  <% IF ActCMS_Book(3) = "1" Then Response.Write "Checked" %>  id="radio8" type="radio" name="ActCMS_Book3" value="1">
	    <label for="radio8"><font color="red">禁用 &nbsp;</font></label></td>
    </tr>
	<tr>
      <td align="left" >分页页数：</td>
	  <td ><input name="ActCMS_Book4" type="text"  class="ainput" value="<%= ActCMS_Book(4) %>" size="20">
	    大于0小于100</td>
    </tr>
	<tr>
      <td align="left" >留言字数控制：</td>
	  <td ><input name="ActCMS_Book5" type="text"  class="ainput"   value="<%= ActCMS_Book(5) %>" size="20">
	    0为不限制</td>
    </tr>
  
  
      <tr>
        <td >留言本模板：</td>
        <td >
		<input name="ActCMS_Book7" type="text"  class="ainput"   value="<%= ActCMS_Book(7) %>" size="40">
          <input class="ACT_btn" type="button"  onClick="OpenWindowAndSetValue('../../include/print/SelectPic.asp?CurrPath=<%=ACTCMS.ActSys%><%=actcms.SysThemePath&"/"&actcms.NowTheme%>',500,320,window,document.form1.ActCMS_Book7);" value="选择模板...">			</td>
      </tr>

	 
	 <tr>
        <td >脏话过滤：</td>
        <td ><textarea name="ActCMS_Book6" cols="50" rows="3"><%=ActCMS_Book(6)%></textarea>
            <br>
          用户发表的评论带有此类文字将被替换,用英文逗号  , 符号分隔</td>
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
	
	 if (form.ActCMS_Book4.value=='')
		{ alert("请输入分页页数!");   
		  form.ActCMS_Book4.focus();    
		   return false;
		} 
		form.Submit1.value="正在提交数据,请稍等...";
		form.Submit1.disabled=true;
		form.Submit2.disabled=true;		
	    form.submit();
        return true;
	}</script> 

</body>
</html>
