<!--#include file="../../ACT.Function.asp"-->
<!--#include file="../../../ACT_inc/cls_pageview.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ACTCMS_标签目录</title>
<link href="../../Images/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td class="bg_tr">您现在的位置：后台管理 &gt;&gt; 标签目录管理</td>
  </tr>
  <tr>
    <td>
	<a href="?Action=add">
	<strong>添加标签目录</strong></a>┆
	<a href="?"><strong>查看标签目录</strong></a>
	</a>
	</td>
  </tr>
</table>
<% 
If Not ACTCMS.ACTCMS_QXYZ(0,"bqxt","") Then   Call Actcms.Alert("对不起，你没有操作权限！","")

	Dim sql, sqlCount,Sqls,intPageSize, strPageInfo,arrRecordInfo, i,pages,intPageNow,strLocalUrl,Action,Foldername,Field1
	Action=Request("Action")
	Dim ShowErr
		IF Request.QueryString("Action") = "del" Then
		Dim ID:ID = Request("ID")
			IF ID = "" Then
				Call Actcms.ActErr("请指定要删除的标签目录","","1")
				Response.end
			End IF
		If instr(ID,",")>0 then
			ID=replace(ID," ","")
			Sql="delete from ACT_LabelFolder where ID in (" & ID & ")"
		Else
			Sql="delete from ACT_LabelFolder where ID=" &  ChkNumeric(ID) & ""
		End if
		Conn.Execute sql:Set Conn=nothing
		Call Actcms.ActErr("标签目录删除成功","include/Label/ACT.LabelFolder.asp","")
 	  End IF
	  
	  Select Case Action
	  		Case "edit","add"
				call edit()
			Case "AddSave","EditSave"
				Call Saves()
			Case Else
				call main()
		end select
		sub main()
	 Dim ACT_TypeDiY,TypeDiY,Manage
	strLocalUrl = request.ServerVariables("SCRIPT_NAME")
	intPageNow = request.QueryString("page")
	intPageSize =20
	sql = "SELECT [ID], [Foldername]" & _
		" FROM [ACT_LabelFolder] " & _
		" ORDER BY [ID] asc"
	sqlCount = "SELECT Count([ID])" & _
			" FROM [ACT_LabelFolder] "
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
  <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="table">
    <form name="Article" method="post" action="?Action=">
    <tr>
      <td  align="center" class="bg_tr" nowarp>选中 </td>
      <td  align="center" class="bg_tr" nowarp>ID</td>
      <td width="60%" align="center" class="bg_tr">标签目录</td>
      <td  colspan="2" align="center" class="bg_tr" nowarp>管理操作</td>
    </tr>
	 <%
		Dim bgColor
		If IsArray(arrRecordInfo) Then
			For i = 0 to UBound(arrRecordInfo, 2)
			bgColor="#FFFFFF"
			if i mod 2=0 then bgColor="#DFEFFF"
	%>
    <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
      <td align="center" >
	  <input type="checkbox" name="ID" value="<%= arrRecordInfo(0,i) %>">	  </td>
      <td align="center" ><%= arrRecordInfo(0,i) %></td>
      <td align="center" ><%= arrRecordInfo(1,i) %></td>
      <td colspan="2" align="center">
	  <a href="?Action=edit&id=<%= arrRecordInfo(0,i) %>">修改</a>┆
	  <a href="?Action=del&ID=<%= arrRecordInfo(0,i) %>" onClick="return confirm('确认删除此标签目录吗?')">删除</a>	  </td>
    </tr>
	<% 
	Next
	End If
	%>
    <tr >
      <td height="30" colspan="6">
	 <label for=chk>
		<input id="chk" type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form)">选择全部</label>
	  
	  <input type="button" class="ACT_btn"  name="yd" value="批量删除" onClick="delpost()"></td>
    </tr>
    <tr >
      <td height="25" colspan="6" align="center"><%= strPageInfo%></td>
    </tr></form>
</table>

<p>
  <script language="javascript">
function CheckAll(form)  
  {  
 for (var i=0;i<form.elements.length;i++)  
    {  
    var e = Article.elements[i];  
   if (e.name != 'chkall')  
      e.checked = Article.chkall.checked;  
   }  
  }
//CSS背景控制
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
 


function delpost(){
    document.Article.method="post";
    document.Article.action="?Action=del";
{
	if(confirm('确认要删除选中的标签目录吗?')){
	this.document.Article.submit();
	return true;}return false;
}
	}
</script>
  <% end sub
Sub edit() 
	If Action ="edit" Then 
		Dim Rs,ID,A
		id = ChkNumeric(Request.QueryString("id"))
		Set Rs=actcms.actexe("select * from ACT_LabelFolder Where id="&id&"")
		If rs.eof Then
			Call actcms.alert("未知错误","")
		Else
			Foldername=Rs("Foldername")
			Id=Rs("Id")
		End If
		A="EditSave"
	Else
		A="AddSave"	
	End If 
%>
<table width="100%" border="0" cellpadding="1" cellspacing="1" class="table">
<form name="form1" method="post" action="?action=<%= A %>&ID=<%= ID %>">
  <tr>
    <td align="right">标签目录名称：</td>
    <td><input name="Foldername" type="text" value="<%= Foldername %>" size="50"></td>
  </tr>
  <tr>
    <td colspan="2" align="center"><input type=button onclick=CheckForm() class="ACT_btn"  name=Submit1 value="  保存  " />
        &nbsp;&nbsp;&nbsp;&nbsp;<input name="Submit2" type="reset" class="ACT_btn" value="  重置  ">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
    </tr>
</form>
</table>

<%end sub 


sub saves()
		dim Rs,RsSql
		 Foldername=ACTCMS.S("Foldername")
		 ID=ChkNumeric(ACTCMS.S("ID"))
		 IF ACTCMS.S("Foldername") = "" Then
			Call ACTCMS.Alert("请输入标签目录名称!",""):Exit Sub
		 End if
		If Action="AddSave" Then 
			 If Not ACTCMS.ACTEXE("SELECT Foldername FROM ACT_LabelFolder Where Foldername='" & Foldername & "' order by ID desc").eof Then
				Call ACTCMS.Alert("系统已存在该标签目录!",""):Exit Sub
			 End if
			 Set Rs = Server.CreateObject("adodb.recordset")
			  RsSql = "select * from ACT_LabelFolder"
			  Rs.Open RsSql, Conn, 1, 3
			  Rs.AddNew
		 	  Rs("Foldername") = Foldername
			  Rs.Update
			  Rs.Close:Set Rs = Nothing			
			  Call Actcms.ActErr("添加成功","include/Label/ACT.LabelFolder.asp","")
 		Else
		 	If Not ACTCMS.ACTEXE("SELECT Foldername FROM ACT_LabelFolder Where ID <>" & ID & " AND  Foldername='" & Foldername & "' order by ID desc").eof Then
				Call ACTCMS.Alert("系统已存在该联系邮箱!",""):Exit Sub
			 End if	
			 Set Rs = Server.CreateObject("adodb.recordset")
			  RsSql = "select * from ACT_LabelFolder Where ID="&ID
			  Rs.Open RsSql, Conn, 1, 3
		 	  Rs("Foldername") = Foldername
			  Rs.Update
			  Rs.Close:Set Rs = Nothing			
			  Call Actcms.ActErr("修改成功","include/Label/ACT.LabelFolder.asp","")
 		End If 
end sub
CloseConn %>
<script language="javascript">
function CheckForm()
{ var form=document.form1;
	
	 if (form.Foldername.value=='')
		{ alert("请输入标签目录名称!");   
		  form.Foldername.focus();    
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
