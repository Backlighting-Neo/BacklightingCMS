<!--#include file="ACT.Function.asp"-->
<!--#include file="../ACT_inc/cls_pageview.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ACTCMS_Label</title>
<link href="Images/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<% 	Dim Types,ShowErr,Sqls,LabelFlag
	Dim strLocalUrl,i,LabelFolderName,rst
	If Not ACTCMS.ACTCMS_QXYZ(0,"bqxt","") Then   Call Actcms.Alert("对不起，你没有操作权限！","")
	Types = ChkNumeric(ACTCMS.S("Type"))
	LabelFlag = request.QueryString("LabelFlag")
	strLocalUrl = request.ServerVariables("SCRIPT_NAME")
	IF Request.QueryString("Action") = "Del" Then
		IF ChkNumeric(Request.QueryString("ID")) = "" Then Response.Write "错误参数":response.end
		actcms.ACTExe("Delete from Label_Act Where ID=" & ChkNumeric(Request.QueryString("ID")))		
		Set conn=nothing
 		Call Actcms.ActErr("标签删除成功","Label_Admin.asp?Type="&Types&"","")
   		Response.End
	End IF
	Dim intPageNow
	intPageNow = request.QueryString("page")
	IF LabelFlag <> "" Then LabelFlag = "And LabelFlag ="& LabelFlag & " "
	Dim intPageSize, strPageInfo
	intPageSize = 20
	Dim arrRecordInfo,pages
	Dim sql, sqlCount
	pages = "LabelFlag="&request.QueryString("LabelFlag")&"&key=" & server.URLEncode(request("key")) & "&Type="&Types&"&page"
	If trim(Request("key"))<>"" Then 
		Sqls = " where LabelType=1 and LabelName Like '%" & RSQL(request("key")) & "%' "
		Types=1
	Else 
		Sqls = " Where LabelType = "&Types&" "& LabelFlag
	End If 


				
	Sql = "SELECT [ID], [LabelName], [AddDate],[LabelFlag],[LabelContent]" & _
		" FROM [Label_Act]" &Sqls& _
		"ORDER BY [ID] deSC"
	SqlCount = "SELECT Count([ID])" & _
			" FROM [Label_Act]"&Sqls

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
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr>
    <td class="bg_tr">您现在的位置：后台管理中心 >> 标签库</td>
  </tr>
  <tr>
    <td ><a href="?Type=<%=request.QueryString("Type")%>">全部标签</a>
	<%Set rst=actcms.actexe("Select ID,Foldername From ACT_LabelFolder")
		Do While Not rst.Eof
			response.write "<a href=""?Type=1&LabelFlag="&rst("id")&""">"&rst("Foldername")&"</a> ┆"
		rst.MoveNext
		Loop
	  rst.Close:Set rst=Nothing
	%>
	<a href="Label_Admin.asp?Type=2"><strong>自定义静态标签</strong></a>&nbsp; <a href="Include/AddLabel.asp?Action=Add" ><strong>创建标签</strong></a><strong>
	┆<a href="Include/StaticLabel.asp?Action=Add">创建自定义标签</a></strong>
	┆<strong><a href="Include/FreeLabel.asp?Action=Add">创建自由标签</a></strong>
	┆<strong><a href="Include/ACT.LabelinOut.asp?A=Out">标签导出</a></strong>
	┆<strong><a href="Include/ACT.LabelinOut.asp?A=in">标签导入</a></strong>
	</td>
  </tr>
</table>


<table width='98%'  border='0' align="center" cellpadding='0' cellspacing='1' class="table">
   <tr> 
          <form name='form3' Action='?Type=1' method='post'>
		 

          <td>
          	<table width='720' border='0' cellpadding='0' cellspacing='0'>
          	<tr>
          	<td width='120' align="right">
          标签搜索 关键字：        </td>
          <td width='160'>
          	<input type='text' name='key' value='<%=request("key")%>' class="Ainput" style='width:150'>  </td>
          <td>
            <input  name="Submit"  type="submit" class="ACT_btn" value="  搜索  ">		  </td>
          </tr>
        </table>
          </td>
        </form>
  </tr>
</table>



<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
  <tr>
    <td width="5%" align="center" class="bg_tr">序号</td>
    <td  align="center" class="bg_tr">标签名称</td>
	<% IF Types = "1" Then%>
	<td  align="center" class="bg_tr">标签属性</td>
	<%End If %>
	<td align="center" class="bg_tr">标签目录</td>
	<td align="center" class="bg_tr">所属模型</td>

    <td  align="center" class="bg_tr">时间</td>
    <td align="center" class="bg_tr" nowrap>描述/操作</td>
  </tr>
  <%
		If IsArray(arrRecordInfo) Then
			For i = 0 to UBound(arrRecordInfo, 2)
%>
  <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
    <td align="center" ><%= arrRecordInfo(0,i) %></td>
    <td >
	<% IF Types = "1" Then%>
	<a href="include/EditLabel.asp?ID=<%= arrRecordInfo(0,i) %>"><%= arrRecordInfo(1,i) %></a>
	<% ElseIF  Types = "2" Then%>
	<a href="include/StaticLabel.asp?Action=EditLabel&ID=<%= arrRecordInfo(0,i) %>" ><%= arrRecordInfo(1,i) %></a>
	<% Else%>
	<a href="include/FreeLabel.asp?Action=EditLabel&ID=<%= arrRecordInfo(0,i) %>" ><%="{ACTSQL_" & Replace(Replace(arrRecordInfo(1,i), "{ACTSQL_", ""), "}", "") & "()}"  %></a>
	<% End IF %></td>
	 
	 <%
	 IF Types = "1" Then
	 response.write "<td align=""center"">"
	  Dim str,FileNames,ModeID
	  Str=mid(arrRecordInfo(4,i), InStrrev(arrRecordInfo(4,i), "("))
  	  FileNames= Replace(Split(Split(arrRecordInfo(4,i),"§")(0),"(")(0),"{$","")
 	  Select Case FileNames
				Case "GetArticleList"
					FileNames="栏目文章列表"
					'ModeID=Split(Str,"§")(22)
				Case "GetArticlePic"
					FileNames="图片文章列表"
				Case "GetSlide"
					FileNames="幻灯片文章"
				Case "GetLastArticleList"
					FileNames="分页文章列表"
					'ModeID=Split(Str,"§")(20)
				Case "GetClassForArticleList"
					FileNames="循环栏目文章"
				Case "CorrelationArticleList"
					FileNames="相关文章列表"
					ModeID=0
				Case "GetNavigation"
					FileNames="网站位置导航"
					ModeID=0
				Case "GetLinkList"
					FileNames="友情链接列表"
				Case "GetSpecial"
					FileNames="专题列表标签"
 				Case "GetClassNavigation"
					FileNames="网站总导航"
					ModeID=0
 	  End Select 
		response.write  FileNames&"标签 </td>"
	End If 
	 %>
	
	 <td align="center" >
	 <%
	 Set LabelFolderName = ACTCMS.ACTEXE("Select ID,Foldername From ACT_LabelFolder where id="&arrRecordInfo(3,i)&" ")
	 If Not LabelFolderName.eof Then
		response.write "<font color=green>"&LabelFolderName("Foldername")&"</font>"
	 Else
		response.write "<font color=green>系统默认</font>"
	 End If 
	 %>
	 </td>
    <td align="center" ><%If ModeID="0" Then response.write "<font color=green>通用模型</font>":Else response.write ACTCMS.ACT_C(modeid,1)&"模型" %></td>

    <td align="center" ><%= arrRecordInfo(2,i) %></td>
	<td align="center" >
    <% 	If Types=1 Then 
	 if FileNames<>"分页文章列表" then  %>
    <a  onClick=show("daima<%= arrRecordInfo(0,i) %>")  href="#"><font color="red">JS调用</font></a>
    <%
	else %>
     <a href="#" disabled="disabled"> JS调用 </a>
	<% end if 
	end if %>
	<% IF Types = "1" Then%>
	<a href="include/ACT.LabelCopy.Asp?A=C&N=<%= arrRecordInfo(1,i) %>&Type=<%= Types %>&ID=<%= arrRecordInfo(0,i) %>">复制</a> ┆
	<a href="include/EditLabel.asp?ID=<%= arrRecordInfo(0,i) %>" >修改</a>
	<% ElseIF  Types = "2" Then%>
	<a href="include/StaticLabel.asp?Action=EditLabel&ID=<%= arrRecordInfo(0,i) %>" >修改</a>
	<% Else%>
	<a href="include/FreeLabel.asp?Action=EditLabel&ID=<%= arrRecordInfo(0,i) %>" >修改</a>
	<% End IF %>
	┆<a href="?Action=Del&Type=<%= Types %>&ID=<%= arrRecordInfo(0,i) %>" onClick="return confirm('确认删除此标签吗?')"> 删除</a> </td>
  </tr>
    <tr id="daima<%= arrRecordInfo(0,i) %>" style="display:none;"><td colspan="7" align="center">
	<input type="text" class="Ainput" size="100%" value="<script language='javascript' src='<%= actcms.ActCMSDM %>plus/JS.asp?LID=<%= arrRecordInfo(0,i) %>&ClassID={$ClassID}&ModeID={$ModeID}&ID={$ID}'></script>"> 
	请将此段代码加到要显示的位置
</td></tr>
  <% 
	Next
	End If
	%>
	<tr >
	<td height="25" colspan="7" align="center" ><%= strPageInfo%></td></tr>
</table>
</body>
</html>
<script language="javascript">
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
function OpenWidndows(Url,Width,Height,WindowObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;status:0;help:0;scroll:0;');
	return ReturnStr;
}
function OpenWindows(url, width, height){
var Win = window.open(url,"openScript",'width=' + width + ',height=' + height + 
',resizable=1,scrollbars=1,menubar=0,status=yes');
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
</script>