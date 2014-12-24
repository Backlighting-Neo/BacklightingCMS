<!--#include file="../ACT.Function.asp"-->
<!--#include file="../../ACT_inc/cls_pageview.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ACTCMS自定义属性</title>
<link href="../Images/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/lhgcore.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/lhgdialog.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/Main.js"></script>

</head>
<body>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td class="bg_tr">您现在的位置：系统中心 &gt;&gt; 自定义管理</td>
  </tr>
  <tr>
    <td>
	<a href="ACT.DiyArticle.ASP?Action=add">
	<strong>添加自定义属性</strong></a>&nbsp;
	<a href="ACT.DiyArticle.ASP"><strong>查看自定义属性</strong></a>

	┆
	
	<a href="ACT.author.ASP?Action=add"><strong>添加作者</strong></a>&nbsp;
	
	<a href="ACT.author.ASP"><strong>查看作者</strong></a>
	┆
	
 
	
	<a href="ACT.CopyFrom.ASP?Action=add"><strong>添加来源</strong></a>
	
	<a href="ACT.CopyFrom.ASP"><strong>查看来源</strong></a>&nbsp;
┆
	<a href="ACT.SiteLink.ASP?Action=add"><strong>添加网站关键字</strong></a>
	<a href="ACT.SiteLink.ASP"><strong>查看网站关键字</strong></a>
	</td>
  </tr>
</table>
<% If Not ACTCMS.ChkAdmin() Then  Call Actcms.Alert("对不起，您没有操作权限！","")
	Dim sql, sqlCount,Sqls,intPageSize, strPageInfo,arrRecordInfo, i,pages,intPageNow,strLocalUrl,Action,Aname,AID
	Action=Request("Action")
	Dim ShowErr
		IF Request.QueryString("Action") = "del" Then
		Dim ID:ID = Request("ID")
			IF ID = "" Then
				Call Actcms.ActErr("请指定要删除的自定义属性","","1")
 				Response.end
			End IF
		If instr(ID,",")>0 then
			ID=replace(ID," ","")
			Sql="delete from ATT_ACT where ID in (" & ID & ")"
		Else
			Sql="delete from ATT_ACT where ID=" &  ChkNumeric(ID) & ""
		End if
		Conn.Execute sql:Set Conn=nothing
		Call Actcms.ActErr("自定义属性删除成功","include/ACT.DiyArticle.ASp","")
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
	sql = "SELECT [ID], [AID], [Aname]" & _
		" FROM [ATT_ACT]" & _
		" ORDER BY [ID] DESC"
	sqlCount = "SELECT Count([ID])" & _
			" FROM [ATT_ACT]"
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
      <td width="30" align="center" class="bg_tr">选中 </td>
      <td width="50" align="center" class="bg_tr">ID</td>
      <td width="161" align="center" class="bg_tr">属性ID</td>
      <td align="center" class="bg_tr">属性名称</td>
      <td width="141" colspan="2" align="center" class="bg_tr">管理操作</td>
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
      <td align="center" ><%= arrRecordInfo(2,i) %></td>
      <td colspan="2" align="center">
	  <a href="?Action=edit&id=<%= arrRecordInfo(0,i) %>">修改</a>┆
	  <a href="?Action=del&ID=<%= arrRecordInfo(0,i) %>" onClick="return confirm('确认删除此自定义属性吗?')">删除</a>	  </td>
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
	if(confirm('确认要删除选中的Tags吗?')){
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
		Set Rs=actcms.actexe("select * from ATT_ACT Where id="&id&"")
		If rs.eof Then
			Call actcms.alert("未知错误","")
		Else
			Aname=Rs("Aname")
			AID=Rs("AID")
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
    <td width="37%" align="right">属性ID：</td>
    <td width="63%"><input name="AID" type="text" class="Ainput" value="<%= AID %>">
	<span class="h" style="cursor:help;"  onclick="dohelp('ACTDiyArticle_sxid')"  id="ACTDiyArticle_sxid">帮助</span></td>
  </tr>
  <tr>
    <td align="right">属性名称：</td>
    <td><input name="Aname" type="text" class="Ainput" value="<%= Aname %>">
	<span class="h" style="cursor:help;"  onclick="dohelp('ACTDiyArticle_sxmc')"  id="ACTDiyArticle_sxmc">帮助</span></td>
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
		 AID=ACTCMS.S("AID")
		 Aname=ACTCMS.S("Aname")
		 ID=ChkNumeric(ACTCMS.S("ID"))
		 IF ACTCMS.S("AID") = "" Then
			Call ACTCMS.Alert("请输入属性ID!",""):Exit Sub
		 End if
		 IF ACTCMS.S("Aname") = "" Then
			Call ACTCMS.Alert("请输入属性名称!",""):Exit Sub
		 End if
		If Action="AddSave" Then 
			 If Not ACTCMS.ACTEXE("SELECT Aname FROM ATT_ACT Where Aname='" & Aname & "' order by ID desc").eof Then
				Call ACTCMS.Alert("系统已存在该属性名称!",""):Exit Sub
			 End if
			 If Not ACTCMS.ACTEXE("SELECT AID FROM ATT_ACT Where AID=" & AID & " order by ID desc").eof Then
				Call ACTCMS.Alert("系统已存在该属性ID!",""):Exit Sub
			 End if	
			 Set Rs = Server.CreateObject("adodb.recordset")
			  RsSql = "select * from ATT_ACT"
			  Rs.Open RsSql, Conn, 1, 3
			  Rs.AddNew
		 	  Rs("Aname") = Aname
		 	  Rs("AID") = AID
			  Rs.Update
			  Rs.Close:Set Rs = Nothing			
			  Call Actcms.ActErr("添加成功","include/ACT.DiyArticle.ASP","")
 		Else
		 	If Not ACTCMS.ACTEXE("SELECT Aname FROM ATT_ACT Where ID <>" & ID & " AND  Aname='" & Aname & "' order by ID desc").eof Then
				Call ACTCMS.Alert("系统已存在该属性名称!",""):Exit Sub
			 End if	
			 Set Rs = Server.CreateObject("adodb.recordset")
			  RsSql = "select * from ATT_ACT Where ID="&ID
			  Rs.Open RsSql, Conn, 1, 3
		 	  Rs("Aname") = Aname
		 	  Rs("AID") = AID
			  Rs.Update
			  Rs.Close:Set Rs = Nothing			
			  Call Actcms.ActErr("操作成功","include/ACT.DiyArticle.ASP","")
 		End If 
end sub
CloseConn %>
<script language="javascript">
function CheckForm()
{ var form=document.form1;
	
	 if (form.AID.value=='')
		{ alert("请输入属性ID!");   
		  form.AID.focus();    
		   return false;
		} 
	if (isNaN(form.AID.value))
		{ alert("属性ID只能为数字!");   
		  form.AID.focus();    
		   return false;
		} 
	if (form.Aname.value=='')
		{ alert("请输入属性名称!");   
		  form.Aname.focus();    
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
