<!--#include file="../ACT.Function.asp"-->
<!--#include file="../../ACT_inc/cls_pageview.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ACTCMS_Admin</title>
<link href="../Images/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/lhgcore.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/lhgdialog.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/Main.js"></script>

</head>

<body>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td  
	class="bg_tr"><strong>您现在的位置：系统设置 >>插件管理</strong></td>
  </tr>
    <tr>
      <td ><A href="?A=A"><b>添加新插件</b></a></td>
    </tr>

</table>
<% 
 If Not ACTCMS.ChkAdmin() Then '超级管理员检测
	Call Actcms.ACTCMSErr("")
 End If 
	Dim ShowErr,PlusName,PlusIntro,IsUse,PlusID,PlusUrl,Action,OrderID
	Action = ACTCMS.S("A")
	Select Case Action
			Case "A","E"
				Call EditAdd()
			Case "AddSave","EditSave"
				Call SavePlus()
			Case "D"
				Call DelPlus()
			Case Else 
				Call main()
	End Select 

	Sub DelPlus()
		Dim id 
		ID=ChkNumeric(ACTCMS.S("ID"))
		ACTCMS.ACTEXE("Delete From Plus_ACT Where ID=" & ID)
		Call Actcms.ActErr("删除插件成功","include/ACT.Plus.asp","")		
 	End Sub 
	Sub SavePlus()
		Dim PlusRS,PlusSql,ID
		 PlusName=ACTCMS.S("PlusName")
		 PlusIntro=ACTCMS.S("PlusIntro")
		 IsUse=ChkNumeric(ACTCMS.S("IsUse"))
		 OrderID=ChkNumeric(ACTCMS.S("OrderID"))
		 ID=ChkNumeric(ACTCMS.S("ID"))
		 PlusID=ACTCMS.S("PlusID")
		 PlusUrl=ACTCMS.S("PlusUrl")
		 IF ACTCMS.S("PlusName") = "" Then
			Call ACTCMS.Alert("请输入插件名称!",""):Exit Sub
		 End if
		 IF ACTCMS.S("PlusUrl") = "" Then
			Call ACTCMS.Alert("请输入管理地址!",""):Exit Sub
		 End if
		If Action="AddSave" Then 
			 IF ACTCMS.S("PlusID") = "" Then
				Call ACTCMS.Alert("请输入插件标识符!",""):Exit Sub
			 End if
			 If Not ACTCMS.ACTEXE("SELECT PlusName FROM Plus_ACT Where PlusName='" & PlusName & "' order by ID desc").eof Then
				Call ACTCMS.Alert("系统已存在该插件名称!",""):Exit Sub
			 End if	
			 If Not ACTCMS.ACTEXE("SELECT PlusID FROM Plus_ACT Where PlusID='" & PlusID & "' order by ID desc").eof Then
				Call ACTCMS.Alert("系统已存在该插件标识符!",""):Exit Sub
			 End if	
			 Set PlusRS = Server.CreateObject("adodb.recordset")
			  PlusSql = "select * from Plus_ACT"
			  PlusRS.Open PlusSql, Conn, 1, 3
			  PlusRS.AddNew
		 	  PlusRS("PlusName") = PlusName
		 	  PlusRS("PlusIntro") = PlusIntro
		 	  PlusRS("IsUse") = IsUse
		 	  PlusRS("PlusUrl") = PlusUrl
			  PlusRS("PlusID")=ACTCMS.S("PlusID")
			  PlusRS("OrderID") = OrderID
			  PlusRS.Update
			  PlusRS.Close:Set PlusRS = Nothing	
			  Call Actcms.ActErr("添加成功","include/ACT.Plus.asp","")		
 		Else
		 	If Not ACTCMS.ACTEXE("SELECT PlusName FROM Plus_ACT Where ID <>" & ID & " AND  PlusName='" & PlusName & "' order by ID desc").eof Then
				Call ACTCMS.Alert("系统已存在该插件名称!",""):Exit Sub
			 End if	
			 Set PlusRS = Server.CreateObject("adodb.recordset")
			  PlusSql = "select * from Plus_ACT Where ID="&ID
			  PlusRS.Open PlusSql, Conn, 1, 3
		 	  PlusRS("PlusName") = PlusName
		 	  PlusRS("PlusIntro") = PlusIntro
		 	  PlusRS("IsUse") = IsUse
		 	  PlusRS("PlusUrl") = PlusUrl
			  PlusRS("OrderID") = OrderID
			  PlusRS.Update
			  PlusRS.Close:Set PlusRS = Nothing	
			  Call Actcms.ActErr("修改成功","include/ACT.Plus.asp","")		
 		End If 
	End Sub 

	Sub main
	Dim strLocalUrl
	strLocalUrl = request.ServerVariables("SCRIPT_NAME")
	
	Dim intPageNow
	intPageNow = request.QueryString("page")
	
	Dim intPageSize, strPageInfo
	intPageSize = 30
	
	Dim arrRecordInfo, i
	Dim sql, sqlCount,Sqls,pages
	sql = "SELECT [ID], [PlusName],[PlusIntro], [PlusID],[IsUse],[OrderID]" & _
		" FROM [Plus_ACT]" & _
		" Order by [OrderID] asc,[ID] asc"
	sqlCount = "SELECT Count([ID])" & _
			" FROM [Plus_ACT]"

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
	 %>

  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
    <tr>
      <td colspan="6" class="bg_tr">您现在的位置：系统设置 &gt;&gt; <a href="?"><font class="bg_tr">插件管理</font></a> </td>
    </tr>
    <tr>
      <td width="136" align="center">插件名称</td>
      <td width="123" align="center">插件说明</td>
      <td width="245" align="center">是否启用</td>
      <td width="182" align="center">管理操作</td>
    </tr>
	 <%
		Dim bgColor
		If IsArray(arrRecordInfo) Then
			For i = 0 to UBound(arrRecordInfo, 2)
		
	%>
    <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
      <td align="center" nowrap><%= arrRecordInfo(1,i) %></td>
      <td align="center" nowrap><%= arrRecordInfo(2,i) %></td>
      <td align="center" nowrap><% IF arrRecordInfo(4,i) = 0 Then Response.Write "<font color=green>正常</font>" else  Response.Write "<font color=red>禁用</font>" %></td>
      <td align="center" nowrap>
	  <a href="?A=E&ID=<%= arrRecordInfo(0,i) %>">修改</a>┆
	  <a href="?A=D&ID=<%= arrRecordInfo(0,i) %>"  onClick="return confirm('确认删除此插件吗?')">删除</a>
	  </td>
    </tr>
	<% 
	Next
	End If
	%>
    <tr >
      <td height="25" colspan="6" align="center"><%= strPageInfo%></td>
    </tr>
  </table>
<%End Sub 

	Sub EditAdd()
	If Action ="E" Then 
		Dim Rs,ID,A
		id = ChkNumeric(Request.QueryString("id"))
		Set Rs=actcms.actexe("select * from Plus_ACT Where id="&id&"")
		If rs.eof Then
			Call actcms.alert("未知错误","")
		Else
			PlusName=Rs("PlusName")
			PlusIntro=Rs("PlusIntro")
			PlusID=Rs("PlusID")
			PlusUrl=Rs("PlusUrl")
			IsUse=Rs("IsUse")
			Id=Rs("Id")
			OrderID=Rs("OrderID")
		End If
		A="EditSave"
	Else
		A="AddSave"	
		OrderID=10
	End If 
%>
    <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <form name="Plus_ACT" method="post" action="?A=<%= A %>&ID=<%= ID %>">
      <tr>
        <td align="right"><strong>插件名称：</strong></td>
        <td><input name="PlusName" size=40 type="text"  class="Ainput"  id="PlusName" value="<%= PlusName %>">
		<span class="h" style="cursor:help;"  onclick="dohelp('ACTplus_cjmc')"  id="ACTplus_cjmc">帮助</span>多个管理名称请用 - 号分割</td>
      </tr>
   



	  <tr>
        <td align="right"><strong>插件说明：</strong></td>
        <td><textarea name="PlusIntro" cols="50" rows="8" id="PlusIntro"><%= PlusIntro %></textarea>
		<span class="h" style="cursor:help;"  onclick="dohelp('ACTplus_cjsm')"  id="ACTplus_cjsm">帮助</span></td>
      </tr>


	   <tr>
        <td align="right"><strong>排序：</strong></td>
        <td><input name="OrderID"   class="Ainput"  type="text" id="OrderID" value="<%= OrderID %>">
		<span class="h" style="cursor:help;"  onclick="dohelp('ACTplus_cjpx')"  id="ACTplus_cjpx">帮助</span>
数字越小排得越靠前.如果权重(排列序号)数字相同，就根据ID来排列		</td>
      </tr>



      <tr>
        <td align="right"><strong>插件标识符：</strong></td>
        <td><input name="PlusID"  class="Ainput"   <% if A="EditSave" then response.Write "disabled" %>  type="text" id="PlusID" value="<%= PlusID %>">
		<span class="h" style="cursor:help;"  onclick="dohelp('ACTplus_cjbsf')"  id="ACTplus_cjbsf">帮助</span><font color=red>这是你插件的唯一的标识，注意不能有重复的</font></td>
      </tr>
      <tr>
        <td align="right"><strong>是否启用：</strong>：</td>
        <td><input <% IF IsUse = 0 Then Response.Write "Checked" %>  type="radio" id="IsUse1" name="IsUse" value="0">
		<label for="IsUse1">启用</label>
        <input <% IF IsUse = 1 Then Response.Write "Checked" %>  type="radio" id="IsUse2" name="IsUse" value="1">
		<label for="IsUse2">关闭</label>
		<span class="h" style="cursor:help;"  onclick="dohelp('ACTplus_sfqy')"  id="ACTplus_sfqy">帮助</span></td>
      </tr>
      <tr>
        <td align="right"><strong>访问地址：</strong></td>
        <td><input name="PlusUrl" size="40" class="Ainput"  type="text" id="PlusUrl" value="<%= PlusUrl %>">
		<span class="h" style="cursor:help;"  onclick="dohelp('ACTplus_fwdz')" id="ACTplus_fwdz">帮助</span>多个管理地址请用 - 号分割(和上面要相当应)</td>
      </tr>
      <tr>
        <td colspan="2" align="center">
		<input type=button onclick=CheckForm() class="ACT_btn"  name=Submit1 value="  保存  " />
        &nbsp;&nbsp;&nbsp;&nbsp;<input name="Submit2" type="reset" class="ACT_btn" value="  重置  "></td>
      </tr>
  </form>
	</table>

<script language="javascript">
function CheckForm()
{ var form=document.Plus_ACT;
	
	 if (form.PlusName.value=='')
		{ alert("请输入插件名称!");   
		  form.PlusName.focus();    
		   return false;
		} 
		 <% if A="AddSave" then %>
	 if (form.PlusID.value=='')
		{ alert("请输入插件标识符!");   
		  form.PlusID.focus();    
		   return false;
		} 
		<%end if %>
	if (form.PlusUrl.value=='')
		{ alert("请输入管理地址!");   
		  form.PlusUrl.focus();    
		   return false;
		} 
		form.Submit1.value="正在提交数据,请稍等...";
		form.Submit1.disabled=true;
		form.Submit2.disabled=true;		
	    form.submit();
        return true;
	}</script>  <%End Sub %>
	<script language="javascript">


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
}</script> 
<% CloseConn %>
</body>
</html>
