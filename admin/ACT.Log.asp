<!--#include file="ACT.Function.asp"-->
<!--#include file="../ACT_inc/cls_pageview.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ACTCMS_Admin</title>
<link href="Images/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript" src="../ACT_inc/js/lhgcore/lhgcore.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../ACT_inc/js/lhgcore/lhgdialog.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../ACT_inc/js/lhgcore/Main.js"></script>
</head>

<body>
<% 
 If Not ACTCMS.ChkAdmin() Then '超级管理员检测
	Call Actcms.ACTCMSErr("")
 End If 
	Dim ShowErr
	IF Request.QueryString("Action") = "Del" Then
		Dim DatAllowDate ,Sqllog
		DatAllowDate = DateAdd("d", -2, Date) 
		Sqllog = "Delete From Log_ACT Where Times<=#" & DatAllowDate & " 23:59:59#"
		Conn.Execute (Sqllog)		
 		Call Actcms.ActErr("日志删除成功,注意!!!两天内的日志不会被删除","ACT.Log.asp","")
 		Response.end
		
	End IF
	
	Dim strLocalUrl
	strLocalUrl = request.ServerVariables("SCRIPT_NAME")
	
	Dim intPageNow
	intPageNow = request.QueryString("page")
	
	Dim intPageSize, strPageInfo
	intPageSize = 30
	
	Dim arrRecordInfo, i
	Dim sql, sqlCount,Sqls,pages
	pages = "Type="&Request("Type")&"&page"
	Select Case Request.QueryString("Type")
		Case "1"
			Sqls = "  where act=1  "
		Case "2"
			Sqls = "  where act=2  "
		Case "3"
			Sqls = "  where act=3  "
		Case "4"
			Sqls = "     "
		Case Else
			Sqls = ""
	End Select
	sql = "SELECT [ID], [UserName], [act], [Times], [LoginIP], [ACTError], [gethttp]" & _
		" FROM [Log_ACT]" &Sqls& _
		"ORDER BY [ID] DESC"
	sqlCount = "SELECT Count([ID])" & _
			" FROM [Log_ACT]"

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
	 %><form name="SysLog" method="post" action="?Action=">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td  class="bg_tr"><strong>系统设置----日志管理首页<a href="#" target="_blank" style="cursor:help;'" class="Help"></a></strong></td>
  </tr>
  <tr>
    <td><strong><a href="?">首页</a> ┆ <a href="?Type=1">系统登陆 </a>┆ <a href="?Type=2">系统操作 </a>┆ <a href="?Type=3">会员操作 </a>┆<a href="?Type=4"> 全部日志 </a></strong></td>
    </tr>
</table>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
    <tr>
      <td colspan="7" class="bg_tr">您现在的位置：系统设置 &gt;&gt; <a href="?"><font class="bg_tr">日志管理</font></a> </td>
    </tr>
    <tr>
      <td width="136" align="center">操作者</td>
      <td width="123" align="center">动作</td>
      <td width="245" align="center">时间</td>
      <td width="182" align="center">IP地址</td>
      <td width="261" align="center">提示信息</td>
      <td width="261" align="center">详细操作</td>
    </tr>
	 <%
		Dim bgColor
		If IsArray(arrRecordInfo) Then
			For i = 0 to UBound(arrRecordInfo, 2)
			
		
	%>
    <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
      <td align="center" ><%= arrRecordInfo(1,i) %></td>
      <td align="center" ><%
 		Select Case  arrRecordInfo(2,i)
			Case "1"
				response.Write "登录"
			Case "2"
				response.Write "系统操作"
			case "3"
			    response.Write "会员操作"
			case else
				response.Write "未知错误"
		End Select 
	  %></td>
      <td align="center"><%= arrRecordInfo(3,i) %></td>
      <td align="center" > 
	  <a  href="#" id="ip<%= arrRecordInfo(0,i) %>" onClick="javascript:upload('<%= arrRecordInfo(0,i) %>','<%= arrRecordInfo(4,i) %>');"><%= arrRecordInfo(4,i) %></a></td>
      <td align="center"><%= arrRecordInfo(5,i) %></td>
      <td align="center" onClick=show("daima<%= arrRecordInfo(0,i) %>") >查看</td>
    </tr>
	
	<tr  id="daima<%= arrRecordInfo(0,i) %>" style="display:none;">
      <td height="30" colspan="7">
	  <%= arrRecordInfo(6,i) %></td>
    </tr>
	
	<% 
	Next
	End If
	%>
    <tr >
      <td height="30" colspan="7" align="right">
	   <input name="Action" type="hidden" id="Action">
	    <input type="button" class="ACT_btn" value="删除所有日志，只能删除最近两天以前的日志"  onClick="delpost()">		</td>
    </tr>
    <tr >
      <td height="25" colspan="7" align="center"><%= strPageInfo%></td>
    </tr>
  </table>
</form>
<script language="javascript">

 function upload(id,ip) 
{
  ( new J.dialog({ id:'ip'+id ,title:'查看IP', loadingText:'网页加载中...',  link:true,page: 'http://www.actcms.com/ip/?q='+ip+ "&" + Math.random(), width:700, height:240 })).ShowDialog();
  }
function show(id)
{
	if(document.all(id).style.display=='none')
{
	document.all(id).style.display='block';
}
else
{
	document.all(id).style.display='none';
}
}
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
    document.SysLog.method="post";
    document.SysLog.action="?Action=Del";
{
	if(confirm('确定清空所有日志吗?注意!两天内的日志将不会被清空!')){
	this.document.SysLog.submit();
	return true;}return false;
}
	}


</script>
<% CloseConn %>
</body>
</html>
