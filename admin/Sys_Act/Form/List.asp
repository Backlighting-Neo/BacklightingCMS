<!--#include file="../../ACT.Function.asp"-->
<!--#include file="../../../ACT_inc/cls_pageview.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>自定义表单管理 By ACTCMS.COM</title>
<link href="../../Images/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<% 	Dim ModeID,TableName,ModeName,Action,Rs,id
	 ModeID = ChkNumeric(Request("ModeID"))
	 ID = ChkNumeric(Request("ID"))
	 if ModeID=0 or ModeID="" Then ModeID=1
	 if id=0 or id="" Then id=1
	 If Not ACTCMS.ACTCMS_QXYZ(0,"form_ACT","") Then   Call Actcms.Alert("没有权限","") 
 	set rs=ACTCMS.actexe("select * from ModeForm_ACT where ModeID="&ModeID)
	if  rs.eof then  response.write "错误":response.end
	TableName=rs("ModeTable")
	ModeName=rs("ModeName")
	

 %><table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td class="bg_tr">您现在的位置：后台管理 >> <a href="index.asp">模型列表</a> >> <a href="?A=L&ModeID=<%= ModeID %>">字段列表</a> >> [<%= ModeName%>]模型  查看反馈</td>
  </tr>
  <tr>
    <td>当前表单： <a href="index.asp?A=Add"><b>添加自定义表单</b></a> </td>
  </tr>
</table>
<%
 	Action = Request("A")
	Select Case Action
			Case "D"
				Call Del()
			case "List"
				call list()
			Case Else
				Call Main()
	End Select
	Sub del()
		ACTCMS.ACTEXE("Delete From  "&TableName&" Where ID=" & ChkNumeric(Request("ID")))
		Call Actcms.ActErr("删除表单成功","Sys_Act/Form/list.asp?ModeID="&ModeID&"","")		
 	End Sub 

	sub list()
 	dim rs,MX_Arr,i,k,rs1,rs2,mx
	set rs2=actcms.actexe("SELECT * FROM Table_ACT Where actcms=3 and  ModeID=" & ModeID & " order by OrderID desc,ID asc ")
    If  rs2.eof Then response.Write "没有找到这条记录,请返回":response.end
	do while not rs2.eof
		mx=mx&rs2("FieldName")&","
 	rs2.movenext
	loop
	 mx=Left(mx, Len(mx) - 1)
  	set rs=actcms.actexe("select "&mx&" ,* from "&TableName&" where id="&id&" order by id desc")
     If  rs.eof Then response.Write "没有找到这条记录,请返回":response.end
 
 
 
	 %>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td width="10%" align="center" class="bg_tr" nowrap>标题</td>
    <td width="90%" align="center" class="bg_tr">提交内容</td>
  </tr>
  
  
  <tr  onMouseOver=overColor(this) onMouseOut=outColor(this)>
    <td align="right" >用户：</td>
    <td align="left">

<%  If ActCMS.UserM(rs("UserID"))=false Then 
    response.Write "匿名"
 	Else
    response.Write ActCMS.UserM(rs("UserID"))
 	End If  %>

</td>
  </tr>
   <tr  onMouseOver=overColor(this) onMouseOut=outColor(this)>
    <td align="right" >提交时间：</td>
    <td align="left">

<%=rs("UpdateTime") %>

</td>
  </tr> 
  
     <tr  onMouseOver=overColor(this) onMouseOut=outColor(this)>
    <td align="right" >用户IP：</td>
    <td align="left">

<%=rs("UserIP") %>

</td>
  </tr> 
  
  
  <% =ACT_MXEdit(ModeID,ID) 
    %>
</table>
<% 



end sub 
	
	Public Function ACT_MXEdit(ModeID,ID)'表现方式.输出模型
	 Dim RS
	  Set RS=ACTCMS.ACTEXE("Select * from Table_ACT  Where ModeID=" & ModeID & "  and actcms=3 order by OrderID desc,ID asc")
	  	Do While Not RS.Eof
			ACT_MXEdit=ACT_MXEdit &"<tr>"&vbCrLf&"<td width=""13%"" align=""right"">"&RS("Title")&"：</td>"&vbCrLf&"<td>"&EditField(RS,ModeID,ID)&"</td>"&vbCrLf&"</tr>"&vbCrLf
			
		RS.MoveNext
		Loop
	  RS.Close:Set RS=Nothing
	 ACT_MXEdit=vbCrLf&ACT_MXEdit& vbCrLf 
	End function


	Function EditField(RSObj,ModeID,id)
		Dim i,IsNotNull,TitleTypeArr,checked,rs1,FieldName
		Dim arrtitle,arrvalue,titles
	  Set RS1=ACTCMS.ACTEXE("Select * from "&TableName&"  Where id="&id&"")
	  FieldName= RSObj("FieldName")
	
		If rsobj("IsNotNull")="0" Then 
			IsNotNull="  <font color=red title=""必填"">*</font>  "&rsobj("Description")
		Else
			IsNotNull="  "&rsobj("Description")
		End If 
		 
	    EditField= RS1(FieldName)& vbCrLf 
	  RS1.Close:Set RS1=Nothing
	End Function 




	Sub Main()

	Dim strLocalUrl
	strLocalUrl = request.ServerVariables("SCRIPT_NAME")
	Dim intPageNow
	intPageNow = request.QueryString("page")
	Dim intPageSize, strPageInfo
	intPageSize = 30
	Dim arrRecordInfo, i
	Dim sql, sqlCount,pages
	 pages = "ModeID="&ModeID&"&page"
	sql = "SELECT [ModeID], [UserID], [UpdateTime], [UserIP],[ID]" & _
		" FROM ["&TableName&"] where ModeID="&ModeID&" " & _
		"ORDER BY [id] desc"
	sqlCount = "SELECT Count([ModeID])" & _
			" FROM ["&TableName&"]  where ModeID="&ModeID&" "

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
    <td width="5%" align="center" class="bg_tr" nowrap>ID号</td>
    <td width="8%" align="center" class="bg_tr">提交者</td>
    <td align="center" class="bg_tr">查看</td>
    <td align="center" class="bg_tr">时间</td>
    <td align="center" class="bg_tr">IP</td>
    <td width="10%" align="center" class="bg_tr" nowrap>管理操作</td>
  </tr>
  <%  	Dim bgColor
		If IsArray(arrRecordInfo) Then
			For i = 0 to UBound(arrRecordInfo, 2)
  %>
  <tr  onMouseOver=overColor(this) onMouseOut=outColor(this)>
    <td align="center" ><%= arrRecordInfo(4,i) %></td>
    <td align="center"><%
	 					 If ActCMS.UserM(arrRecordInfo(1,i))=false Then 
							  response.Write "匿名"
							Else
 							    response.Write ActCMS.UserM(arrRecordInfo(1,i))
							End If 
 %></td>
    <td align="center"><a href="?A=List&ModeID=<%= ModeID %>&ID=<%= arrRecordInfo(4,i) %>">查看</a></td>
    <td align="center"><%= arrRecordInfo(2,i) %></td>
    <td align="center"><%= arrRecordInfo(3,i) %></td>
    <td align="center" ><a href="?A=D&ModeID=<%= ModeID %>&ID=<%= arrRecordInfo(4,i) %>"  onClick="{if(confirm('确定删除吗?')){return true;}return false;}">删除</a></td>
  </tr><% 
		Next
	End If %>
	    <tr >
      <td height="25" colspan="7" align="center" class="td_bg"><%= strPageInfo%></td>
    </tr>
</table>
<%
		 

	End Sub 
%>

<script language="JavaScript" type="text/javascript">

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


</body>
</html>