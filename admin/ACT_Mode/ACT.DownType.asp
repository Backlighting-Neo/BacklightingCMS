<!--#include file="../ACT.Function.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>模型管理</title>
<link href="../Images/style.css" rel="stylesheet" type="text/css">
 </head>
<body>
<% 	 
  	dim DownName,DownPath,IsOuter,isDisp,DownPoint,UserGroup,act,a,id,rs,rs1,sqlstr,rootid
	ID=ChkNumeric(Request("ID"))
	A=Request("A")
	rootid=ChkNumeric(Request("rootid"))
	select case a
		case "A","E","ED"
			call add()
		case "D"
			call del()
		case "serveradd","serversave"
			call saveserver()
		case else
			call main()
	
	end select
	
	sub del()
		Set rs = server.CreateObject("adodb.recordset")'删除图片
		rs.open "Select * from DownType_ACT Where ID=" & ID & " ", conn, 1, 3
		If Not  rs.eof  Then
			if rs("rootid")="0" then 
   			Conn.execute("Delete from DownType_ACT  Where rootid= "&rs("rootid"))
  			Conn.execute("Delete from DownType_ACT  Where id= "&id)
			else
  			Conn.execute("Delete from DownType_ACT  Where id= "&id)
			end if 
 		End If 
 		 Call Actcms.ActErr("操作成功.请返回继续","ACT_Mode/ACT.DownType.asp?","")
 	end sub 
	
	sub saveserver()
 	DownName=actcms.s("DownName")
	DownPath=actcms.s("DownPath")
	if DownName="" then 
		Call Actcms.ActErr("请输入服务器名称","","1")
 	end if 
	if DownPath="" then 
		Call Actcms.ActErr("请输入服务器路径","","1")
 	end if 
	UserGroup=actcms.s("UserGroup")
  	isDisp=ChkNumeric(Request("isDisp"))
	DownPoint=ChkNumeric(Request("DownPoint"))
	IsOuter=ChkNumeric(Request("IsOuter"))
 	
	IF A = "serveradd" Then
		 Set rs = Server.CreateObject("adodb.recordset")
		  sqlstr = "select * from DownType_ACT"
		  rs.Open sqlstr, Conn, 1, 3
		  rs.AddNew
		  rs("DownName") = DownName
		  rs("DownPath") = DownPath
		  rs("UserGroup") = UserGroup
		  rs("isDisp") = isDisp
		  rs("DownPoint") = DownPoint
		  rs("IsOuter") = IsOuter
		  rs.Update
 		 Call Actcms.ActErr("操作成功.请返回继续","ACT_Mode/ACT.DownType.asp?","")
		ElseIF A = "serversave" Then
		  Set rs = Server.CreateObject("adodb.recordset")
		  sqlstr = "select * from DownType_ACT Where ID="&ID
		  rs.Open sqlstr, Conn, 1, 3
		  rs("DownName") = DownName
		  rs("DownPath") = DownPath
		  rs("UserGroup") = UserGroup
		  rs("isDisp") = isDisp
		  rs("DownPoint") = DownPoint
		  rs("IsOuter") = IsOuter
		  rs.Update
 		 Call Actcms.ActErr("操作成功.请返回继续","ACT_Mode/ACT.DownType.asp?","")
		End If
 	end sub
	
	sub main() %>
<form id="form1" name="form1" method="post" action="">
  <table width="100%" border="0" cellpadding="2" cellspacing="1" class="table">
    <tr>
      <td colspan="3" class="bg_tr">您现在的位置：
      <a href="ACT.DownType.asp?A=A">添加下载服务器</a> | 
      <a href="?">管理下载服务器</a> 
       </td>
    </tr>
	<tr>
	  <td width="40%" align="center"><strong>服务器分类</strong></td>
	  <td width="40%" align="center" ><strong>操 作</strong></td>
	  <td width="10%" align="center"><strong>下载数</strong></td>
	</tr>			
 	<%
 	Set rs=actcms.actexe("select * from DownType_ACT where rootid=0")
	
	If Not rs.eof Then 
		Do While Not rs.eof 
	%>
	<tr  onMouseOver=overColor(this) onMouseOut=outColor(this)>
      <td width="30%"><img src="../images/-.gif" /><b><%=rs("downname")%></b></td>
      <td align="right" >&nbsp;<a href="?A=A&ID=<%=rs("id")%>">添加下载服务器路径</a> | <a href="?A=ED&ID=<%=rs("id")%>">服务器设置</a> | <a href="?A=D&ID=<%=rs("id")%>"  onClick="return confirm('删除该下载服务器路径,会将该服务器下所有下载路径删除,此操作不可恢复,是否确认删除?')">删除</a>&nbsp;</td>
      <td align="center">4111</td>
	</tr>
	<%Set rs1=actcms.actexe("select * from DownType_ACT where rootid<>0  and     rootid="&rs("rootid"))
	If Not rs1.eof Then 
		Do While Not rs1.eof %>
				<tr  onMouseOver=overColor(this) onMouseOut=outColor(this)>
				  <td>&nbsp;<img src="../images/L.gif" /><%=rs1("downname")%>
				  &nbsp;<%If rs1("IFLock")="1" Then  response.write "<img src=""../images/Lock.gif"" />"%></td>
				  <td align="right" >
                  <% If rs1("IFLock")="1" Then  %>
                  <a href="?A=free&ID=<%=rs1("id")%>"><font color=green>解除</font></a>  |
                  <% else %>
                  &nbsp;<a href="?A=Lock&ID=<%=rs1("id")%>"><font color=red>锁定</font></a> |
                  <% end if  %>
                 	<a href="?A=E&ID=<%=rs1("id")%>&rootid=<%=rs1("rootid")%>">服务器设置</a>  |
                   <a href="?A=D&ID=<%=rs1("id")%>"  onClick="return confirm('删除将包括该服务器的所有信息，确定删除吗?')">删除</a>&nbsp;
                   </td>
				  <td align="center">4</td>
				</tr>			
			
		<%rs1.movenext
		Loop
	End If 
	
	rs.movenext
	loop
	
	
	End If %>
    <tr>
      <td colspan="6">&nbsp;</td>
    </tr>
  </table>
</form>
<script language="JavaScript" type="text/JavaScript">
function overColor(Obj)
{
	var elements=Obj.childNodes;
	for(var i=0;i<elements.length;i++)
	{
		elements[i].className="tdbg1"
		Obj.bgColor="";
	}
	
}

function outColor(Obj)
{
	var elements=Obj.childNodes;
	for(var i=0;i<elements.length;i++)
	{
		elements[i].className="tdbg";
		Obj.bgColor="";
	}
}
 


</script>
<% end sub  

	sub add()
	set rs=actcms.actexe("select * from DownType_ACT where id="&id&"")
	if a="ED" then 
		DownName=rs("DownName")
	    DownPath=rs("DownPath")
	    IsOuter=rs("IsOuter")
	    isDisp=rs("isDisp")
	    DownPoint=rs("DownPoint")
	    UserGroup=rs("UserGroup")
		act="save"
	else
		isDisp=0
	    IsOuter=0
		act="add"
	end if
%>
<form id="form1" name="form1" method="post" action="?A=server<%= act %>&ID=<%= ID %>">

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="table" >
  <tr>
    <th colspan="2" class="bg_tr">添加新的服务器</th>
  </tr>
  <tr>
    <td width="30%" ><U>服务器名称</U></td>
    <td width="70%" ><input type="text"  class="Ainput"name="DownName" size="60" value="<%= DownName %>">    </td>
  </tr>
  <tr>
    <td ><U>服务器路径</U></td>
    <td ><input type="text"  class="Ainput"name="DownPath" size="60" value="<%= DownPath %>">    </td>
  </tr>
  <tr>
    <td ><U>所属类别</U></td>
    <td ><select name="servers">
       <option value="0" selected>做为服务器分类</option>
    <% dim rs2
	Set rs2=actcms.actexe("select * from DownType_ACT where rootid=0")
	
	If Not rs2.eof Then 
		if a="ED" then 
		 %>     
          <option value="<%= rs2("id") %>" <% if rs2("id")=id then response.Write "selected"  %>><%= rs2("DownName") %></option>

		<% else
		
		Do While Not rs2.eof  %> 
      <option value="<%= rs2("id") %>" <% if rs2("rootid")=rootid then response.Write "selected"  %>><%= rs2("DownName") %></option>
   
   
   <% rs2.movenext
   loop
   end if
   end if %>
    </select></td>
  </tr>
  <tr>
    <td ><U>使用下载服务器的权限</U></td>
    <td ><%= actcms.GetGroup_CheckBox("UserGroup",UserGroup,5)  %>	   </td>
  </tr>
  <tr>
    <td ><U>下载所需点数</U></td>
    <td ><input type="text"  class="Ainput"name="DownPoint" size="10" value='<%= DownPoint %> '>    </td>
  </tr>
  <tr>
    <td ><U>是否直接显示下载地址</U></td>
    <td ><input type="radio" name="isDisp" value="0"  <% if isDisp="0" then response.Write "checked"  %>>
      否&nbsp;&nbsp;
      <input type="radio" name="isDisp" value="1" <% if isDisp="1" then response.Write "checked"  %>>
      是 </td>
  </tr>
  <tr>
    <td ><U>是否外部连接</U></td>
    <td ><input type=radio name="IsOuter" value="0" <% if IsOuter="0" then response.Write "checked"  %>>
      否&nbsp;&nbsp;
      <input type=radio name="IsOuter" value="1" <% if IsOuter="1" then response.Write "checked"  %>>
      是&nbsp;&nbsp;
      <input type=radio name="IsOuter" value="2" <% if IsOuter="2" then response.Write "checked"  %>>
      迅雷专用下载地址&nbsp;&nbsp;
      <input type=radio name="IsOuter" value="3" <% if IsOuter="3" then response.Write "checked"  %>>
      快车专用下载地址 <br>
      <font color="red">注意：如果是外部连接，请在“服务器路径”中输入要转向的URL；<br>
        &nbsp;&nbsp;&nbsp;&nbsp;如果选择“迅雷或快车专用下载地址”，请先注册<a href="http://union.xunlei.com/" target="_blank"><font color="blue">迅雷联盟</font></a>|<a href="http://union.flashget.com/" target="_blank"><font color="blue">快车联盟</font></a>，然后在<a href="../sys/admin_setting.asp?action=plus"><font color="blue">联盟插件设置</font></a>中输入相应的联盟ID</font></td>
  </tr>
  <tr>
    <td >&nbsp;</td>
    <td ><input type="submit" name="Submit" class="ACT_btn" value=" 保存 ">    </td>
  </tr>
</table>
</form>
<% end sub  %>
</body>
</html>