<!--#include file="../../ACT.Function.asp"-->
<!--#include file="../../../ACT_inc/cls_pageview.asp"-->
<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Digg管理 By ActCMS</title>
<link href="../../Images/style.css" rel="stylesheet" type="text/css"></head>
<body>
<%
	If Not ACTCMS.ACTCMS_QXYZ(0,"digg_act","") Then   Call Actcms.Alert("对不起，你没有操作权限！","") 
	With Response
	Dim ShowErr,ModeID,ModeName
	ModeID = ChkNumeric(Request("ModeID"))
	if ModeID=0 or ModeID="" Then ModeID=1
	ModeName= ACTCMS.ACT_C(ModeID,1) 
	Dim strLocalUrl
	strLocalUrl = request.ServerVariables("SCRIPT_NAME")
	Dim intPageNow
	intPageNow = request("page")
	Dim intPageSize, strPageInfo
	intPageSize = 20
	Dim arrRecordInfo, i
	Dim sql, sqlCount,Sqls,pages
	pages = "ModeID="&ModeID&"&page"
	IF Request.QueryString("Action") = "del" Then
	Dim ID:ID = Request("ID")
		IF ID = "" Then
			Call Actcms.ActErr("请指定要删除的ID","","1")
			Response.end
		End If
	 Dim digge,Diggs
	ID = Split(ID,",")
	For I = LBound(ID) To UBound(ID)
		Set Digge=ACTCMS.ACTEXE("Select digg,NewsID from Digg_ACT where ID="&ID(I)&"")
		If Not Digge.eof Then 
			If digge("digg")="1" Then Diggs="digg" Else diggs="down"
			ACTCMS.ACTEXE("Update "&ACTCMS.ACT_C(ModeID,2)&" set "&Diggs&"="&Diggs&"-1 where ID = "&digge("NewsID")&"")
	    End If 
		ACTCMS.ACTEXE("delete from Digg_ACT where ID="&ID(I)&"")
	Next 
		Call Actcms.ActErr("删除成功","Sys_Act/Digg/Index.asp","")
 		Response.end
	End IF
	sql = "SELECT [ID], [IP], [NewsID], [DiggTime], [Digg], [users],  [ModeID]" & _
		" FROM [Digg_ACT] Where ModeID="&clng(ModeID)&Sqls& _
		" ORDER BY [ID] DESC"
	sqlCount = "SELECT Count([ID])" & _
			" FROM [Digg_ACT] Where ModeID="&clng(ModeID)&Sqls
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
<table width="99%" border="0" align="center"  class="table">
  <tr>
    <td  class="bg_tr"><strong>您现在的位置：<%= ModeName %>系统管理 &gt;&gt; <%= ModeName %>管理</strong></td>
  </tr>

  <tr>
    <td >查看选项：
	<%
	Dim MX_Sys,ii
	MX_Sys=ACTCMS.Act_MX_Sys_Arr()
	If IsArray(MX_Sys) Then
		For iI=0 To Ubound(MX_Sys,2)
		response.write "<a href=""?ModeID="&MX_Sys(0,Ii)&""">"&MX_Sys(1,Ii)&"系统</a> ┆" 
		Next
	End If
	%>
</td>
  </tr>
</table>
  <table width="99%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
<form name="Article" method="post" action="?Action=">
    <tr>
      <td width="28" align="center" class="bg_tr">选中 </td>
      <td width="24" align="center" class="bg_tr">ID</td>
      <td width="150" align="center" class="bg_tr">文章标题</td>
	    <td width="80" align="center" class="bg_tr">DIGG用户</td> 
	    <td  align="center" class="bg_tr" nowrap>Digg时间</td> 
	     <td width="60" align="center" class="bg_tr">Digg行为</td> 
	    <td width="50" align="center" class="bg_tr">IP</td> 
	  
    </tr>
	 <%
		If IsArray(arrRecordInfo) Then
			For i = 0 to UBound(arrRecordInfo, 2)
	%>
    <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
      <td align="center"  >
	  <input type="checkbox" name="ID" value="<%= arrRecordInfo(0,i) %>">	  </td>
      <td align="center"  ><%= arrRecordInfo(0,i) %></td>
      <td  ><a target="_blank" href="<%
	 .write  ACTCMS.actsys&"List.asp?C-"&ModeID&"-"&arrRecordInfo(2,i)
	%>"><%
	
	Dim AName
	Set AName=ACTCMS.ACTEXE("Select title from  "&ACTCMS.ACT_C(ModeID,2)&"  where ID = "& arrRecordInfo(2,i))
	If Not Aname.eof Then 
		response.write Aname("title")
	Else
		response.write "<font color=red>该文章已被删除</font>"
	End If 
 %></a>&nbsp;<%
	  %></td>
      
	  <td align="center"  >
	  <%
		If Trim(arrRecordInfo(5,i))<>"" Then 
			response.write "<font color=green>"&arrRecordInfo(5,i)&"</font>"
		Else
			response.write "<font color=red>游客</font>" 
		End If 
%></td>
  
      <td align="center" ><%= arrRecordInfo(3,i) %></td>
	 
      <td align="center" ><%If arrRecordInfo(4,i)="1" Then response.write "<font color=green>支持</font>" Else response.write "<font color=red>反对</font>"%></td>
	   
      <td align="center" ><%=arrRecordInfo(1,i)%></td>
    </tr>
	<% 
	Next
	End If
	%>
    <tr >
      <td height="30" colspan="11" >
<input name="ChkAll" type="checkbox" id="ChkAll" onClick="CheckAll(this.form)" value="checkbox">
		<label for="chkAll">&nbsp;选中本页显示的所有文章</label>
	  <input type="button" class="act_btn" name="Submit" value="批量删除"  onClick="delpost()"> 
	  删除的同时会减少文章的相关DIGG记录</td>
    </tr>
    <tr >
      <td height="25" colspan="11" align="center" ><%= strPageInfo%></td>
    </tr></form>

  </table>

<script language="javascript">
function SelectIterm(form,sign){
	for (var i=0; i<form.elements.length;i++ ){
		if (form.elements[i].type == "checkbox"){
				var e=form.elements[i];
					if (sign==0) e.checked= true;
					if (sign==1) e.checked= !e.checked;
					if (sign==2) e.checked= false;
		}
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

function CheckAll(form)
		  {  
		 for (var i=0;i<form.elements.length;i++)  
			{  
			   var e = Article.elements[i];  
			   if (e.name != 'ChkAll'&&e.type=="checkbox")  
			   e.checked = Article.ChkAll.checked;  
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
    document.Article.action="Index.asp?ModeID=<%=ModeID%>&Action=del";
{
	if(confirm('确认要删除选中的Digg吗?注意删除并不会删除文章')){
	this.document.Article.submit();
	return true;}return false;
}
	}
</script>


<% End With

CloseConn 
%>
</body>
</html>
