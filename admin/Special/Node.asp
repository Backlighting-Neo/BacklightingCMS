<!--#include file="../ACT.Function.asp"-->
 <!--#include file="../../ACT_inc/cls_pageview.asp"-->
 <!--#include file="../include/ACT.F.asp"-->

 <html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>专题节点系统管理-By ACTCMS</title>
<link href="../Images/style.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="../../../ACT_INC/js/swfobject.js"></script>
 <script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/lhgcore.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/lhgdialog.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/Main.js"></script>
 <SCRIPT LANGUAGE='JavaScript'>
 var U="<%=ACTCMS.strToAsc(RSQL(Trim(Request.Cookies(AcTCMSN)("AdminName"))))%>";
var P="<%=ACTCMS.strToAsc(RSQL(Trim(Request.Cookies(AcTCMSN)("AdminPassword"))))%>";
 </SCRIPT>
 <style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
 </style>
 </head>
<body>
<%
dim SID,title
SID= ChkNumeric(Request.QueryString("SID")) 
if SID<>0 then title=Conn.Execute("Select title from Special_ACT Where ID=" & SID)(0)


%>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td  class="bg_tr"><strong>您现在的位置：系统中心 &gt;&gt; 插件管理 &gt;&gt;<a href="?">专题管理</a></strong></td>
  </tr>
  <tr>
    <td  class="tdclass">
	<strong>专题选项：</strong>
	
	<strong><a href="Index.asp">查看专题</a> </strong>┆
	
	<strong><a href="?Action=add">新建通用节点</a></strong>
	 <% if sid<>0 then %>
	<strong><a href="?Action=add&SID=<%= SID %>">新建[<%=title%>]专题节点</a></strong>
	<% end if  %>
  ┆<strong><a href="Node.asp?SID=<%= SID %>">查看专题节点</a></strong> 	
  ┆<strong><a href="Node.asp">查看所有节点</a></strong> 	
  
  </td>
  </tr>
</table>
<% 
 
 dim Action,ID,ShowErr,notename,arcid,isauto,keywords,ClassID,DiyContent,ModeID,TitleLen,DateForm,ContentLen,ListNumber,str
	Action = Request("Action")
	ID = Request("ID")
	Select Case Action
			Case "add","edit" 
				Call edit()
			Case "saveadd"
				Call saveadd()
			Case "saveedit"
				Call saveedit()
			Case "del"
				Call del()
			Case Else
				Call Main()
	End Select
	
	Sub Del()
		Conn.Execute ("Delete from Node_ACT Where ID=" & ChkNumeric(Request.QueryString("ID")))		
		Set conn=nothing
		Call Actcms.ActErr("删除成功","Special/Node.asp?SID="&SID&"","")
 		Response.End
    End Sub
	Sub saveadd()
	    Dim Rs,content 
		 If Not ACTCMS.ACTEXE("SELECT notename FROM Node_ACT Where notename='" & RSQL(Request.Form("notename")) & "' order by ID desc").eof Then
			Call ACTCMS.Alert("系统已存在该节点名称!",""):Exit Sub
		 End if	
 		Set Rs=Server.CreateObject("Adodb.RecordSet")
		Rs.Open "Select * from Node_ACT",Conn,1,3
		Rs.addnew
 		
		Rs("notename")= Request.Form("notename")
 		Rs("arcid")= trim(Request.Form("arcid"))
		Rs("isauto")= Request.Form("isauto")
 		Rs("keywords")= Request.Form("keywords")
		Rs("classid")= Request.Form("classid")
		Rs("DiyContent")= Request.Form("DiyContent")
 		Rs("SID")=SID
		Rs("ModeID")= ChkNumeric(Request.Form("ModeID"))
		Rs("TitleLen")= ChkNumeric(Request.Form("TitleLen"))
		Rs("DateForm")= ChkNumeric(Request.Form("DateForm"))
		Rs("ContentLen")= ChkNumeric(Request.Form("ContentLen"))
		Rs("ListNumber")= ChkNumeric(Request.Form("ListNumber"))
		Rs.update
		Rs.Close:Set Rs=Nothing
 		Call Actcms.ActErr("添加成功","Special/Node.asp?SID="&SID&"","")
 		Response.End
	End Sub	


	Sub saveedit()
 	Dim DiyPath,Rs,content,sql
	
		 If Not ACTCMS.ACTEXE("SELECT notename FROM Node_ACT Where id<>"&id&" and notename='" & RSQL(Request.Form("notename")) & "' order by ID desc").eof Then
			Call ACTCMS.Alert("系统已存在该节点名称!",""):Exit Sub
		 End if	
		SQL = "Select * From Node_ACT Where ID="&ID
		Set Rs=Server.CreateObject("Adodb.RecordSet")
		Rs.Open SQL,Conn,1,3
		Rs("notename")= Request.Form("notename")
 		Rs("arcid")= trim(Request.Form("arcid"))
		Rs("isauto")= Request.Form("isauto")
 		Rs("keywords")= Request.Form("keywords")
		Rs("classid")= Request.Form("classid")
		Rs("DiyContent")= Request.Form("DiyContent")
		Rs("ModeID")= ChkNumeric(Request.Form("ModeID"))
		Rs("TitleLen")= ChkNumeric(Request.Form("TitleLen"))
		Rs("DateForm")= ChkNumeric(Request.Form("DateForm"))
		Rs("ContentLen")= ChkNumeric(Request.Form("ContentLen"))
		Rs("ListNumber")= ChkNumeric(Request.Form("ListNumber"))
  		Rs.update
		Rs.Close:Set Rs=Nothing
 		Call Actcms.ActErr("修改成功","Special/Node.asp?SID="&SID&"","")
 		Response.End
	End Sub	

	sub Main()
	Dim ShowErr
	ConnectionDatabase
	Dim strLocalUrl,rs
	strLocalUrl = request.ServerVariables("SCRIPT_NAME")
	Dim intPageNow
	intPageNow = request.QueryString("page")
	Dim intPageSize, strPageInfo
	intPageSize = 20
 	Dim arrRecordInfo, i
	Dim sql, sqlCount,Sqls
	if sid<>0 then Sqls="where sid="&sid
	sql = "SELECT [ID],[notename],[sid]" & _
		" FROM [Node_ACT] "&Sqls&"" & _
		" ORDER BY [ID] DESC"
	sqlCount = "SELECT Count([ID])" & _
			" FROM [Node_ACT]  "&Sqls&" "
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
   <form name="Article" method="post" action="?Action="> <tr>
      <td width="26" align="center" class="bg_tr">ID</td>
      <td width="524" align="center" class="bg_tr">节点名称</td>
      <td width="524" align="center" class="bg_tr">所属专题</td>
      <td width="202" align="center" class="bg_tr">节点调用标签</td>
      <td width="202" align="center" class="bg_tr">状态</td>
      <td width="200" colspan="2" align="center" class="bg_tr">常规管理操作</td>
    </tr>
	 <%
		Dim bgColor
		If IsArray(arrRecordInfo) Then
			For i = 0 to UBound(arrRecordInfo, 2)
	%>
    <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
      <td align="center"   class="tdclass"><%= arrRecordInfo(0,i) %></td>
      <td align="center"  class="tdclass" ><%= arrRecordInfo(1,i) %> </td>
      <td align="center"  class="tdclass" ><%
	   set rs=Conn.Execute("Select title from Special_ACT Where ID=" & arrRecordInfo(2,i)) 
	   if not rs.eof then 
 	   	response.Write rs("title")
	   else 
	   	response.Write "<font color=green>全部专题</font>"
	   end if
	   
	   %></td>
      <td align="center"  class="tdclass" >{$node_<%= arrRecordInfo(1,i) %>}</td>
      <td align="center"  class="tdclass" >
	  <%
	  if arrRecordInfo(2,i)="0" then 
	  	response.Write "<font color=green>通用节点</font>"
	  else 
	  	response.Write "<font color=red>绑定节点</font>"
	  end if
	  
	  
	  %>	  </td>
      <td  align="center" class="tdclass" >
	 
<a href="?Action=edit&ID=<%= arrRecordInfo(0,i) %>">编辑&nbsp;</a>┆
	 <a href="?Action=del&ID=<%= arrRecordInfo(0,i) %> " onClick="return confirm('确认删除此吗?此操作不可恢复!')">删除</a></td>
    </tr>
	<% 
	Next
	End If
	%>
  
    <tr >
      <td height="25" colspan="7" align="center"  class="tdclass" ><%= strPageInfo%></td>
    </tr>
 </form> </table>

<p>
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
	sub edit()
	Dim DiyPath,Rs,content,pagename,frmAction,tempurl
	IF Action = "edit" Then
		Set Rs=Conn.Execute("Select * from Node_ACT Where ID=" & ID)
			if Rs.Bof And Rs.EOF then
				 Response.Write "不存在！"
				Exit Sub
			End IF	
	 
		notename=rs("notename")
		arcid=rs("arcid")
		isauto=rs("isauto")
		keywords=rs("keywords")
		classid=rs("classid")
		DiyContent=rs("DiyContent")
 	  	sid=rs("sid")
		ModeID=rs("ModeID")
		TitleLen=rs("TitleLen")
		DateForm=rs("DateForm")
		ContentLen=rs("ContentLen")
		ListNumber=rs("ListNumber")
   		frmAction = "edit"
	Else
		isauto=1
		classid=0
		DateForm=0	
		ListNumber=10
		TitleLen=10
		ContentLen=10
		frmAction="add"
	End IF
    %>
	
	
	<% if SID=0 then  %>
<table width="98%" border="0" align="center" >
<div class="tip">请注意!当前添加的节点是通用节点,在任何专题模版里都可以调用</div>
</table>
<%
  end if  %>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <form name="form1" method="post" action="Node.asp?SID=<%= SID %>">


	<!--节点分隔符-->

	<tr>
      <td colspan="2"  class="bg_tr"><div align="center">节点  </div></td>
    </tr>
 	<tr>
      <td  class="tdclass">节点 1 名称：</td>
      <td  class="tdclass"> 
      <input name="notename" type="text"  class="Ainput" id="notename" value="<%= notename %>" size="30"></td>
 	</tr>
	
	
	<tr>
      <td  class="tdclass">栏目ID：</td>
      <td  class="tdclass"> 
 
<INPUT id="classid"   size="30" name="classid"  value="<%= classid %>" class="Ainput">0表示不指定,多个请用 <span class="STYLE1">'栏目ID1','栏目ID2'</span>
 模型ID：
<INPUT id="ModeID"   size="10" name="ModeID"  value="<%= ModeID %>" class="Ainput">0表示不指定</td>
    </tr>	
	
	
	
	
     <tr>
      <td  class="tdclass">节点文章列表：</td>
      <td  class="tdclass"><font color=red>内置标签</font> 
<a href="###" onClick='SetDiyContent(DiyContent,"#ID")'>文章ID</a>&nbsp;
<a href="###" onClick='SetDiyContent(DiyContent,"#Link")'>文章链接</a>&nbsp;
<a href="###" onClick='SetDiyContent(DiyContent,"#Title")'>文章标题</a>&nbsp;
<a href="###" onClick='SetDiyContent(DiyContent,"#CTitle")'>文章标题(过滤HTML)</a>&nbsp;
<a href="###" onClick='SetDiyContent(DiyContent,"#KeyWords")'>关键字</a>&nbsp;
<a href="###" onClick='SetDiyContent(DiyContent,"#Thumb")'>缩略图</a>&nbsp;
<a href="###" onClick='SetDiyContent(DiyContent,"#PicUrl")'>图片地址</a>&nbsp;
<a href="###" onClick='SetDiyContent(DiyContent,"#Intro")'>文章导读</a>&nbsp;
<a href="###" onClick='SetDiyContent(DiyContent,"#ClassName")'>栏目名称</a>&nbsp;
<a href="###" onClick='SetDiyContent(DiyContent,"#ClassLink")'>栏目链接</a>&nbsp;<br />
<a href="###" onClick='SetDiyContent(DiyContent,"#Time")'>时间</a>&nbsp;
<a href="###" onClick='SetDiyContent(DiyContent,"#Hits")'>点击数</a>&nbsp;
<a href="###" onClick='SetDiyContent(DiyContent,"#CopyFrom")'>文章来源</a>&nbsp;
<a href="###" onClick='SetDiyContent(DiyContent,"#Author")'>文章作者</a>&nbsp;
<a href="###" onClick='SetDiyContent(DiyContent,"#AutoID")'>自增长ID</a>&nbsp;
<a href="###" onClick='SetDiyContent(DiyContent,"#ModID")'>间歇ID</a>&nbsp;
<a href="###" onClick='SetDiyContent(DiyContent,"#Path")'>系统路径</a>&nbsp;
 <br />
<textarea   onfocus="this.className='colorfocus';" onBlur="this.className='colorblur';" name="DiyContent" id="DiyContent"  cols="95%" rows="10"><%=DiyContent%></textarea></td>
    </tr>
  	<tr>
      <td  class="tdclass">标题字数：</td>
      <td  class="tdclass"> 
        <input name="TitleLen" type="text"  class="Ainput" id="TitleLen" value="<%= TitleLen %>" size="10">
      数量：
<INPUT id="ListNumber" size="16" name="ListNumber"  class="Ainput" value="<%= ListNumber %>">
内容字数：<input name="ContentLen" type="text"  class="Ainput" id="ContentLen" value="<%= ContentLen %>" size="10">
	  日期：<select  style="width:120;" name="DateForm" id="select2">
		 <%= AF.ACT_DateStr(DateForm) %>
        </select></td>
    </tr>
	
	
	<tr>
      <td  class="tdclass">节点文章列表：</td>
      <td  class="tdclass"><textarea name="arcid"  id="arcids"   cols="50"><%=arcid  %></textarea>
       <input   type="button" class="ACT_btn"  onClick="javascript:list('arcid');"  value="选择节点文章" /></td>
    </tr>
 
 <tr>
      <td  class="tdclass">属性：</td>
      <td  class="tdclass">
	  <label for="isauto"><INPUT id="isauto" value="1" <%if isauto="1" then response.Write "CHECKED" %> type="radio" name="isauto">
按节点文章列表</label>
  <label for="isauto2"><INPUT id="isauto2" value="2" type="radio" name="isauto" <%if isauto="2" then response.Write "CHECKED" %>>
自动获取文档</label>   
  &nbsp; 关键字：
<INPUT id="keywords" size="16" name="keywords"  class="Ainput" value="<%= keywords %>">
</td>
    </tr>

<!--节点分隔符-->
  	
	 
    <tr>
      <td colspan="2" align="center"  class="tdclass" >
	<input type=button onclick=CheckForm() class="ACT_btn"  name=Submit1 value="  保存  " />
        &nbsp;&nbsp;&nbsp;&nbsp;<input name="Submit2" type="reset" class="ACT_btn" value="  重置  ">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		 <input name="Action" type="hidden" id="Action" value="save<%=frmAction%>">
		
		 <input name="ID" type="hidden" id="ID" value="<%=id%>"></td>
  </tr></form>
</table>


<script language="javascript">
function OpenWindowAndSetValue(Url,Width,Height,WindowObj,SetObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:yes;help:no;scroll:yes;');
	if (ReturnStr!='') SetObj.value=ReturnStr;
	return ReturnStr;
}	

 function  
  SetDiyContent(oTextarea,strText){   
  oTextarea.focus();   
  document.selection.createRange().text+=strText;   
  oTextarea.blur();   
  }   


function insertHTMLToEditor(name){
	document.getElementById("arcids").value=name;
 }

function list(iname) 
{
		 (new J.dialog({ id: 'zxscs', title: '专题文章选择',width: '860',height: '600',cancelBtnTxt:'确定', page: '<%=actcms.adminPath%>special/list.asp?A=add&iname='+iname+ "&" + Math.random() })).ShowDialog(); 
 }

function CheckForm()
{ var form=document.form1;
	 if (form.title.value=='')
		{ alert("请输入专题名称!");   
		 form.title.focus();    
		   return false;
		} 
	 
		form.Submit1.value="正在提交数据,请稍等...";
		form.Submit1.disabled=true;
		form.Submit2.disabled=true;		
	    form.submit();
        return true;
	}</script>  
	<%end sub
CloseConn %>
</body>
</html>

