<!--#include file="../../ACT.Function.asp"-->
<!--#include file="../../../ACT_inc/cls_pageview.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>广告管理</title>
<SCRIPT LANGUAGE="JavaScript">
function delad(){
if (confirm("确定要删除这则广告么?删除后不可以再恢复哦!?")){return true;}
return false;
}
</SCRIPT>
<link href="../../Images/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%

	dim did,Sqlup
	if  Request.QueryString("Action")="del" then
	did = Request.QueryString("ID")
	Sqlup = "Delete from ads Where ID="&did
	Conn.Execute (Sqlup)		
	 Call Actcms.ActErr("操作成功","Sys_Act/gg/Index.asp","")
 	Response.end
        end if

        Dim sqls
        Select Case Request.QueryString("Action")	
		   Case "stop"
			Sqls = " where ( ADStopViews <> 0 and ADViews > ADStopViews) or ( ADStopHits <> 0 and ADHits > ADStopHits) or ( DateDiff('d',Now(),ADStopDate)<1 ) "
				   Case Else
				 Sqls = ""
		End Select

	Dim strLocalUrl
        Dim id,ADID,ADViews,ADHits,ADType,ADSrc,ADCode,ADLink,ADAlt,ADWidth,ADHeight,ADNote,ADStopViews,ADStopHits,ADStopDate
	strLocalUrl = request.ServerVariables("SCRIPT_NAME")
	
	Dim intPageNow
	intPageNow = request.QueryString("page")
	
	Dim intPageSize, strPageInfo
	intPageSize = 30
	
	Dim arrRecordInfo, i
	Dim sql, sqlCount,pages
	pages = "Type="&Request("Type")&"&page"
	sql = "SELECT [id], [ADID], [ADViews], [ADHits], [ADType], [ADSrc], [ADCode], [ADLink], [ADAlt], [ADWidth], [ADHeight], [ADNote], [ADStopViews], [ADStopHits], [ADStopDate]" & _
		" FROM [ads]" &Sqls& _
		"ORDER BY [ID] DESC"
	sqlCount = "SELECT Count([ID])" & _
			" FROM [ads]"&Sqls

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
%>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td  class="bg_tr"><strong>广告管理----广告管理首页</strong></td>
  </tr>
  <tr>
    <td class="td_bg"><strong><a href="?">首页</a> ┆ <a href="index.asp">广告列表 </a>┆<a href="add.asp">添加广告 </a>┆<a href="http://www.actcms.com/sys/ad.asp">常用广告代码 </a>┆</strong></td>
    </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
    <tr>
      <td colspan="7" class="bg_tr">您现在的位置：广告设置 &gt;&gt; <a href="?"><font class="bg_tr">广告管理</font></a> </td>
    </tr>
    <tr>
      <td width="7%" align="center" class="td_bg">ID序号</td>
      <td width="23%" align="center" class="td_bg">名称</td>
      <td width="8%" align="center" class="td_bg">点击</td>
      <td width="12%" align="center" class="td_bg">显示</td>
      <td width="18%" align="center" class="td_bg">广告类型</td>
      <td width="10%" align="center" class="td_bg">是/否过期</td>
      <td width="32%" align="center" class="td_bg">常规操作</td>
    </tr>
	<%
		Dim bgColor
		If IsArray(arrRecordInfo) Then
			For i = 0 to UBound(arrRecordInfo, 2)
	%>
    <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
      <td align="center" class="td_bg"><%= arrRecordInfo(0,i) %></td>
      <td align="center" class="td_bg"><a href="editjs.asp?id=js_c/<%= arrRecordInfo(1,i)%>.js"><%= arrRecordInfo(1,i)%></a></td>
      <td align="center" class="td_bg"><%= arrRecordInfo(3,i) %></td>
      <td align="center" class="td_bg"><%= arrRecordInfo(2,i) %></td>
      <td align="center" class="td_bg"><%= ShowAdType(arrRecordInfo(4,i),arrRecordInfo(5,i))%></td>
      <td align="center" class="td_bg">
      <%
        If IsStop(arrRecordInfo(2,i),arrRecordInfo(12,i),arrRecordInfo(13,i),arrRecordInfo(3,i),arrRecordInfo(14,i)) Then 
           Response.Write(" <font color=red>已过期</font>")
        Else
           Response.Write(" <font color=red>正常</font>")
        End If
      %>
      </td>
      <td align="center" class="td_bg">
      <a href="edit.asp?id=<%=arrRecordInfo(0,i)%>">修改 </a>｜
      <a href="createjs.asp?id=<%=arrRecordInfo(1,i)%>">生成js </a>｜
      <a href="?Action=del&ID=<%=arrRecordInfo(0,i) %>" onclick='return delad();'>删除 </a>
      </td>
    </tr>
	<% 
	Next
	End If
	%>
    <tr >
      <td height="25" colspan="7" align="center" class="td_bg"><%= strPageInfo%></td>
    </tr>
  </table>
<SCRIPT language=javascript>
<!--
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
//-->
</SCRIPT>
</body>
</html>
<%

'检测是否过期
function IsStop(ADViews,ADStopViews,ADStopHits,ADHits,ADStopDate)
	IsStop=false
	If ( ADStopViews <> 0 and ADViews > ADStopViews) Then 
		IsStop=true
		Exit function
	ElseIf ( ADStopHits <> 0 and ADHits > ADStopHits) Then
		IsStop=true
		Exit function
	ElseIf ( DateDiff("d",Now(),ADStopDate)<1 ) Then	
		IsStop=true
		Exit function
	End If
end function

'判断广告类型
function ShowAdType(ADType,ADSrc)
	Dim ADExt
	ADExt="图片"
	If InStr(1,ADSrc,".swf",1)>0 Then ADExt="FLASH"
	Select Case ADType
		Case 1
			ShowAdType="普通"&ADExt
		Case 2
			ShowAdType="全屏浮动"&ADExt
		Case 3
			ShowAdType="上下浮动 - 右"&ADExt
		Case 4
			ShowAdType="上下浮动 - 左"&ADExt
		Case 5
			ShowAdType="渐隐消失"&ADExt
		Case 6
			ShowAdType="网页对话框"
		Case 7
			ShowAdType="移动透明对话框"
		Case 8
			ShowAdType="打开新窗口"
		Case 9
			ShowAdType="弹出新窗口"
		Case 10 
			ShowAdType="对联式广告"
		Case 11 
			ShowAdType="联盟广告"
		Case else
			ShowAdType="<font color=red><b>错误!将不能正确显示</b>"
	End Select
end function
%>