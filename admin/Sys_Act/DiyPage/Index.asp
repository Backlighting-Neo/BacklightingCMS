<!--#include file="../../ACT.Function.asp"-->
 <!--#include file="../../../ACT_inc/cls_pageview.asp"-->
 <!--#include file="../../actcms.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>单页系统管理-By ACTCMS</title>
<link href="../../Images/editorstyle.css" rel="stylesheet" type="text/css">
<meta http-equiv="X-UA-Compatible" content="IE=8" />
<script type="text/javascript" src="../../../ACT_INC/js/swfobject.js"></script>
 <SCRIPT LANGUAGE='JavaScript'>
 var U="<%=ACTCMS.strToAsc(RSQL(Trim(Request.Cookies(AcTCMSN)("AdminName"))))%>";
var P="<%=ACTCMS.strToAsc(RSQL(Trim(Request.Cookies(AcTCMSN)("AdminPassword"))))%>";

</SCRIPT>
</head>
<body>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td  class="bg_tr"><strong>您现在的位置：系统中心 &gt;&gt; 插件管理 &gt;&gt;<a href="?">单页管理</a></strong></td>
  </tr>
  <tr>
    <td class="tdclass">
	<strong><a href="?Action=add">单页</a>选项：</strong><strong><a href="?Action=add">新建单页</a></strong>

	┆<strong><a href="?">查看单页</a></strong>┆<strong><a href="ACT.diy.asp?RefreshFlag=All">生成全部单页</a></strong>	</td>
  </tr>
</table>
<% 
	If Not ACTCMS.ACTCMS_QXYZ(0,"dyxt_ACT","") Then   Call Actcms.Alert("对不起，你没有操作权限！","") 

 dim Action,ID,ShowErr
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
		Conn.Execute ("Delete from DiyPage_ACT Where ID=" & ChkNumeric(Request.QueryString("ID")))		
		Set conn=nothing
		Call Actcms.ActErr("删除成功","Sys_Act/DiyPage/Index.asp","")
 		Response.End
    End Sub
	Sub saveadd()
	%>

	<%Dim DiyPath,Rs,content,pagename

		DiyPath=Request.Form("DiyPath")
		If InStr(DiyPath, "//") > 0   Then
			DiyPath = Replace(DiyPath, "//","/")
		End If 

		Set Rs=Server.CreateObject("Adodb.RecordSet")
		Rs.Open "Select * from DiyPage_ACT",Conn,1,3
		Rs.addnew
		Rs("DiyPath")= DiyPath
		Rs("content")= Request.Form("content")
		Rs("pagename")= Request.Form("pagename")
 		Rs("tempurl")= Request.Form("tempurl")
		Rs.update
		Rs.Close:Set Rs=Nothing
		Response.Write ("<script>parent.frame.cols=""198,*"";</script>")
		Call Actcms.ActErr("添加成功","Sys_Act/DiyPage/Index.asp","")
 		Response.End
	End Sub	


	Sub saveedit()
	%> 

	<%Dim DiyPath,Rs,content,pagename,sql
		SQL = "Select * From DiyPage_ACT Where ID="&ID
		DiyPath=Request.Form("DiyPath")
		If InStr(DiyPath, "//") > 0   Then
			DiyPath = Replace(DiyPath, "//","/")
		End If 
		Set Rs=Server.CreateObject("Adodb.RecordSet")
		Rs.Open SQL,Conn,1,3
		Rs("DiyPath")= DiyPath
		Rs("content")= Request.Form("content")
		Rs("pagename")= Request.Form("pagename")
		Rs("tempurl")= Request.Form("tempurl")
 		Rs.update
		Rs.Close:Set Rs=Nothing
		Response.Write ("<script>parent.frame.cols=""198,*"";</script>")
		Call Actcms.ActErr("修改成功","Sys_Act/DiyPage/Index.asp","")
 		Response.End
	End Sub	

	sub Main()
	Dim ShowErr
	ConnectionDatabase
	Dim strLocalUrl
	strLocalUrl = request.ServerVariables("SCRIPT_NAME")
	Dim intPageNow
	intPageNow = request.QueryString("page")
	Dim intPageSize, strPageInfo
	intPageSize = 20
	Response.Write ("<script>parent.frame.cols=""198,*"";</script>")
	Dim arrRecordInfo, i
	Dim sql, sqlCount,Sqls,ChannelID
	sql = "SELECT [ID],[pageName],[DiyPath]" & _
		" FROM [DiyPage_ACT]" & _
		"ORDER BY [ID] DESC"
	sqlCount = "SELECT Count([ID])" & _
			" FROM [DiyPage_ACT] "
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
      <td width="524" align="center" class="bg_tr">模板名称</td>
      <td width="202" align="center" class="bg_tr">生成路径</td>
      <td width="200" colspan="2" align="center" class="bg_tr">常规管理操作</td>
    </tr>
	 <%
		Dim bgColor
		If IsArray(arrRecordInfo) Then
			For i = 0 to UBound(arrRecordInfo, 2)
	%>
    <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
      <td align="center" class="tdclass" ><%= arrRecordInfo(0,i) %></td>
      <td align="center"  class="tdclass"><a target="_blank" href="<%=ACTCMS.actsys&"plus/page.asp?ID="&arrRecordInfo(0,i)%>"><%= arrRecordInfo(1,i) %></a> </td>
      <td align="center" class="tdclass" ><%= arrRecordInfo(2,i) %></td>
      <td  align="center" class="tdclass">
	 <a href="ACT.diy.asp?RefreshFlag=ID&ID=<%= arrRecordInfo(0,i) %>">生成&nbsp;</a>┆
<a href="?Action=edit&ID=<%= arrRecordInfo(0,i) %>">编辑&nbsp;</a>┆
	 <a href="?Action=del&ID=<%= arrRecordInfo(0,i) %> " onClick="return confirm('确认删除此模板吗?此操作不可恢复!')">删除</a></td>
    </tr>
	<% 
	Next
	End If
	%>
  
    <tr >
      <td height="25" colspan="5" align="center" class="tdclass"><%= strPageInfo%></td>
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
		Set Rs=Conn.Execute("Select * from DiyPage_ACT Where ID=" & ID)
			if Rs.Bof And Rs.EOF then
				 Response.Write "不存在！"
				Exit Sub
			End IF	
		DiyPath = Rs("DiyPath")	
		pagename = Rs("pagename")
		If RS("Content") <> "" Then Content=Server.HTMLEncode(RS("Content"))
		tempurl=Rs("tempurl")
 		frmAction		= "edit"
	Else
		DiyPath = "html/"
		pagename = "actcms.htm"
		content = ""
		pagename=""
		frmAction = "add"
	End IF
  %>
  <script>parent.frame.cols="0,*";</script>
<table width="98%" border="0" align="center" >
<div class="tip">请注意!为使编辑模板有更大的空间.左边栏已经隐藏.请不要担心.保存或者点击其他操作会恢复原样</div>
</table>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
    <form name="form1" method="post" action="Index.asp">
	
	<tr>
      <td width="19%"  class="tdclass">单页名称:</td>
      <td width="81%"  class="tdclass"><input name="pagename" type="text"  class="Ainput"id="pagename" value="<%= pagename %>" size="50"></td>
    </tr>
   

	<tr>
      <td width="19%"  class="tdclass">单页模板地址:</td>
      <td width="81%"  class="tdclass">
	  <input name="tempurl" type="text"  class="Ainput"id="tempurl" value="<%= tempurl %>" size="50">
          <input class="ACT_btn" type="button"  onClick="OpenWindowAndSetValue('../../include/print/SelectPic.asp?CurrPath=<%=ACTCMS.ActSys%><%=actcms.SysThemePath&"/"&actcms.NowTheme%>',500,320,window,document.form1.tempurl);" value="选择模板..."> 
	  </td>
    </tr>
	
	<tr>
      <td  class="tdclass">单页路径:</td>
      <td  class="tdclass"><input name="DiyPath" type="text"  class="Ainput"id="DiyPath" value="<%= DiyPath %>" size="50">
	  不能以/开始,不熟悉系统.不建议修改.以免覆盖系统文件</td>
    </tr>


 <tr>
<td  height="23" align="right"   class="tdclass">批量上传文件：</td>
<td class="tdclass">
 

<div id="sapload">
    
    </div>
 
 <script type="text/javascript">
// <![CDATA[
var so = new SWFObject("<%=ACTCMS.ACTSYS%>act_inc/sapload.swf", "sapload", "450", "25", "9", "#ffffff");
so.addVariable('types','<%=Replace(ACTCMS.ActCMS_Sys(11),"/",";")%>');
so.addVariable('isGet','1');
so.addVariable('args','myid=Upload;ModeID=999;U='+U+";"+';P='+P+";"+'Yname=content1');
so.addVariable('upUrl','<%=ACTCMS.ACTSYS%><%=ACTCMS.ActCMS_Sys(8)%>/include/Upload.asp');
so.addVariable('fileName','Filedata');
so.addVariable('maxNum','110');
so.addVariable('maxSize','<%=ACTCMS.ActCMS_Sys(10)/1024%>');
so.addVariable('etmsg','1');
so.addVariable('ltmsg','1');
so.addParam('wmode','transparent');
so.write("sapload");
function sapLoadMsg(t){
var actup=t.split('|');
 {
  	   KE.insertHtml(actup[0], actup[1]);
}
}

// ]]>
</script> 
 
 
 
</td>
</tr>
	
	<tr >
      <td colspan="2"  class="tdclass">
	  <script charset="utf-8"  language="JavaScript" type="text/javascript" src="../../../editor/kindeditor/kindeditor.js" ></script>
 		<script>
			KE.show({
				id : 'content1'
  			});
		</script>
	 
	   <textarea id="content1" name="content"  style="width:98%;height:300px;visibility:hidden;">
<%=content%>
</textarea>


 
	  </td>
    </tr>
    <tr>
      <td colspan="2" align="center"   class="tdclass">
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
function CheckForm()
{ var form=document.form1;
	 if (form.pagename.value=='')
		{ alert("请输入单页名称!");   
		  form.pagename.focus();    
		   return false;
		} 
	 if (form.DiyPath.value=='')
		{ alert("请输入单页路径!");   
		  form.DiyPath.focus();    
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

