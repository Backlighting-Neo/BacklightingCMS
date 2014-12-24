<!--#include file="../ACT.Function.asp"-->
<!--#include file="../../ACT_inc/cls_pageview.asp"-->
<!--#include file="../include/ACT.F.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Act内容管理系统</title>
<link href="../Images/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/lhgcore.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/lhgdialog.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/Main.js"></script>
</head>

<body>
<% 
		Dim ShowErr,ID,ModeID
	 ModeID = ChkNumeric(Request("ModeID"))
	 if ModeID=0 or ModeID="" Then ModeID=1
		If Not ACTCMS.ACTCMS_QXYZ(ModeID,"","") Then   Call Actcms.Alert("对不起，您没有"&ACTCMS.ACT_C(ModeID,1)&"系该项操作权限！","")

	ID = ChkNumeric(Request.QueryString("ID"))
		IF Request.QueryString("Action") = "UnLock" Then'置顶
			Conn.execute("Update Comment_Act set Locked=0 where ID ="&ID&"")
			set conn=nothing
			Response.Redirect("?")
		End IF
		
		IF Request.QueryString("Action") = "Lock" Then'置顶
			Conn.execute("Update Comment_Act set Locked=1 where ID ="&ID&"")
			set conn=nothing
			Response.Redirect("?")
		End IF
		IF Request.QueryString("Action") = "del" Then'置顶
			dim Sqlbook,rs
			Set rs=actcms.actexe("select acticleID,ModeID from Comment_Act Where ID=" & ID)
			If Not rs.eof Then 
				actcms.ACTEXE("Update "&ACTCMS.ACT_C(rs("ModeID"),2)&"  Set commentscount=commentscount-1 Where ID=" & rs("acticleID") & "")
				Conn.Execute ("Delete from Comment_Act Where ID=" & ID)		
				set conn=Nothing
			End If 
  			 Call Actcms.ActErr("评论已经删除！","ACT_Mode/ACT.Book.asp","")
		End IF
	


		IF Request.QueryString("Action") = "alldel" Then'单一删除
			ID = Request.Form("ID")
			IF ID = "" Then
				ShowErr = "请先选定评论！"
				Call Actcms.ActErr(ShowErr,"","1")
				Response.End
			End IF		
			ID = Split(ID,",")
			 For I = LBound(ID) To UBound(ID)
				Conn.execute("Delete from Comment_Act  Where ID=" & ID(I) & "")
			Next
				set conn=nothing	
				ShowErr = "评论已经删除！"
  				 Call Actcms.ActErr(ShowErr,"ACT_Mode/ACT.Book.asp","")
	  End IF

	Dim strLocalUrl
	strLocalUrl = request.ServerVariables("SCRIPT_NAME")
	
	Dim intPageNow
	intPageNow = request.QueryString("page")
	
	Dim intPageSize, strPageInfo
	intPageSize = 20
	
	Dim arrRecordInfo, i
	Dim sql, sqlCount,Sqls,pages
	pages = "Type="&Request("Type")&"&ModeID="&ModeID&"&page"
	Select Case Request.QueryString("Type")
		Case "Lock"
			Sqls = " Where ModeID = "&ModeID&" And Locked = 1 "
		Case "UnLock"
			Sqls = " Where ModeID = "&ModeID&" And Locked = 0 "
		Case Else
			Sqls = " Where ModeID = "&ModeID&" "
	End Select
	sql = "SELECT [ID], [ModeID], [Content], [AddDate],[Locked],[UserIP],[ClassID],[acticleID]" & _
		" FROM [Comment_Act]" &Sqls& _
		"ORDER BY [ID] DESC"
 	sqlCount = "SELECT Count([ID])" & _
			" FROM [Comment_Act]  " &Sqls

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
	 %><form name="Article" method="post" action="?Action=">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td  class="bg_tr"><strong>您现在的位置：文章中心管理 &gt;&gt; 评论管理
  选择模型:<select name='ModeID' style='width:110px' onChange="location=this.value;">
  <%=AF.ACT_URL_Mode(ModeID,"")%>
  </select>


</strong></td>
  </tr>
  <tr>
    <td class="td_bg"><strong>评论选项：</strong><strong><a href="?">所有评论</a> ┆ <a href="?Type=Lock">已审核</a>┆ <a href="?Type=UnLock">未审核</a>┆</strong>&nbsp;</td>
    </tr>
</table>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  
    <tr>
      <td width="40" align="center" class="bg_tr">选中 </td>
      <td width="45" align="center" class="bg_tr">ID</td>
      <td width="300" align="center" class="bg_tr">评论内容</td>
      <td width="157" align="center" class="bg_tr">发表时间</td>
      <td width="125" align="center" class="bg_tr">IP</td>
      <td width="84" align="center" class="bg_tr">审核与否</td>
      <td width="154" colspan="2" align="center" class="bg_tr">常规管理操作</td>
    </tr>
	 <%
		Dim bgColor
		If IsArray(arrRecordInfo) Then
			For i = 0 to UBound(arrRecordInfo, 2)
	%>
    <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
      <td align="center" class="td_bg" >
	  <input type="checkbox" name="ID" value="<%= arrRecordInfo(0,i) %>">	  </td>
      <td align="center" class="td_bg" ><%= arrRecordInfo(0,i) %></td>
      <td class="td_bg" ><a target="_blank" href="<%=AcTCMS.ActCMSDM&"plus/Comment/index.asp?ModeID="&arrRecordInfo(1,i)&"&ClassID="&arrRecordInfo(6,i)&"&ID="&arrRecordInfo(7,i)%>"><%=arrRecordInfo(2,i)%></a></td>
      <td class="td_bg" ><%= arrRecordInfo(3,i) %></td>
      <td align="center" class="td_bg" ><a  href="#"  id="ip<%= arrRecordInfo(0,i) %>" onClick="javascript:lookip('<%= arrRecordInfo(0,i) %>','<%= arrRecordInfo(5,i) %>');"><font color="green"><%= arrRecordInfo(5,i) %></font></a></td>
      <td align="center" class="td_bg" >
	 	<% IF arrRecordInfo(4,i) = 1 Then Response.Write "<font color=red>&nbsp;&nbsp;已审核&nbsp;&nbsp;</font>" Else Response.Write "<font color=#0000FF>&nbsp;&nbsp;未审核&nbsp;&nbsp;</font>"%>	  </td>
      <td  align="center" class="td_bg">
	  <a href="?Action=del&id=<%= arrRecordInfo(0,i) %>" onClick="return confirm('确认删除此评论吗?')">删除</a>
	┆<% if arrRecordInfo(4,i) = 0 Then %>
	  <a href="?Action=Lock&id=<%= arrRecordInfo(0,i) %>">通过审核</a> 
   <% ELse
   	%> <a title="解除置顶" href="?Action=UnLock&id=<%= arrRecordInfo(0,i) %>">取消审核</a> 
	<%
	 end if%>	  </td>
    </tr>
	<% 
	Next
	End If
	%>
    <tr >
      <td height="30" colspan="8" class="td_bg">
	 <label for=chk>
		<input id="chk" type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form)">选择全部评论</label>
	&nbsp;&nbsp;&nbsp;  <input type="button" Class="ACT_BTN" name="yd" value="批量删除" onClick="delpost()"></td>
    </tr>
    <tr >
      <td height="25" colspan="8" align="center" class="td_bg"><%= strPageInfo%></td>
    </tr>
  </table>
</form>
<script language="javascript">
 function lookip(id,ip)  
{
  ( new J.dialog({ id:'ip'+id ,title:'查看IP', loadingText:'网页加载中...',  link:true,page: 'http://www.actcms.com/ip/?q='+ip+ "&" + Math.random(), width:700, height:240 })).ShowDialog();
 }
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
    document.Article.action="?Action=alldel";
{
	if(confirm('确认要删除选中的评论吗?')){
	this.document.Article.submit();
	return true;}return false;
}
	}




</script>
<% CloseConn %>
</body>
</html>
