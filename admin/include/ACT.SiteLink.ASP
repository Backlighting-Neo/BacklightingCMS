<!--#include file="../ACT.Function.asp"-->
<!--#include file="../../ACT_inc/cls_pageview.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ACTCMS网站关键字</title>
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
	 
	<a href="ACT.SiteLink.ASP?Action=add"><strong>添加网站关键字</strong></a>
	<a href="ACT.SiteLink.ASP"><strong>查看网站关键字</strong></a>

	</td>
  </tr>
</table>
<% If Not ACTCMS.ChkAdmin() Then  Call Actcms.Alert("对不起，您没有操作权限！","")
	Dim sql, sqlCount,Sqls,intPageSize, strPageInfo,arrRecordInfo, i,pages,intPageNow,strLocalUrl,Action,Title,url,ifs,OpenType
	dim Num,description,OrderID,repset,repcontent
 	Action=Request("Action")
	Dim ShowErr
		IF Request.QueryString("Action") = "del" Then
		Dim ID:ID = Request("ID")
			IF ID = "" Then
				Call Actcms.ActErr("请指定要删除的网站关键字","1")
				Response.end
			End IF
		If instr(ID,",")>0 then
			ID=replace(ID," ","")
			Sql="delete from Sitelink_ACT where ID in (" & ID & ")"
		Else
			Sql="delete from Sitelink_ACT where ID=" &  ChkNumeric(ID) & ""
		End if
		Conn.Execute sql:Set Conn=Nothing
		Call actcms.DelCahe("ReplaceSitelink")
			Call Actcms.ActErr("网站关键字删除成功","include/ACT.SiteLink.ASP","")
 	  End IF
	  
	  Select Case Action
	  		Case "edit","add"
				call edit()
			Case "AddSave","EditSave"
				Call Saves()
			Case "px"
				Call px()
			Case Else
				call main()
		end select
		
		
		
	Sub px()
			Dim i,OrderID
			ID = Split(actcms.s("ID"),","):OrderID = Split(actcms.s("OrderID"),",")
			 For I = LBound(ID) To UBound(ID)
 				 	Conn.execute("Update Sitelink_ACT set OrderID="&ChkNumeric(OrderID(I))&"  where  ID = "&ChkNumeric(ID(I))&"")
			Next 
			set conn=Nothing
			Call Actcms.ActErr("操作成功","include/ACT.SiteLink.ASP","")
	End sub
		
		
		sub main()
	 Dim ACT_TypeDiY,TypeDiY,Manage
	strLocalUrl = request.ServerVariables("SCRIPT_NAME")
	intPageNow = request.QueryString("page")
	intPageSize =20
	sql = "SELECT [ID], [Title], [Url], [IFS], [OpenType],[OrderID],[Num]" & _
		" FROM [Sitelink_ACT]" & _
		" ORDER BY OrderID Asc,[ID] DESC"
	sqlCount = "SELECT Count([ID])" & _
			" FROM [Sitelink_ACT]"
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
      <td  align="center" class="bg_tr" nowrap>关键字名称</td>
      <td align="center" class="bg_tr" nowrap>关键字链接</td>
      <td  align="center" class="bg_tr" nowrap>排序</td>
      <td  align="center" nowrap class="bg_tr">替换次数</td>
	  <td  align="center" class="bg_tr" nowrap>是否启用</td>
      <td  colspan="2" align="center" class="bg_tr" nowrap> 管理操作</td>
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
      <td align="center" ><%= arrRecordInfo(1,i) %></td>
      <td align="center" ><%= arrRecordInfo(2,i) %></td>
      <td align="center" >
	  
	  <input name="OrderID"  type="text" class="Ainput" id="OrderID" value="<%=arrRecordInfo(5,i)%>" size="6" maxlength="20" />
	  
 		  <input name="ID" type="hidden"  value="<%=arrRecordInfo(0,i)%>">	  </td>
      <td align="center" ><%= arrRecordInfo(6,i) %></td>
		  
      <td align="center" ><%If  arrRecordInfo(3,i)="1" Then response.write "<font color=green>正常</font>":Else response.write "<font color=red>禁用</font>"  %></td>
      <td colspan="2" align="center">
	  <a href="?Action=edit&id=<%= arrRecordInfo(0,i) %>">修改</a>┆
	  <a href="?Action=del&ID=<%= arrRecordInfo(0,i) %>" onClick="return confirm('确认删除此网站关键字吗?')">删除</a>	  </td>
    </tr>
	<% 
	Next
	End If
	%>
    <tr >
      <td height="30" colspan="8">
	 <label for=chk>
		<input id="chk" type="checkbox" name="chkall" value="checkbox" onClick="CheckAll(this.form)">选择全部</label>
	  
	  <input type="button" class="ACT_btn"  name="yd" value="批量删除" onClick="delpost()">
	  
	  
	  <input type="button" Class="ACT_btn"  name="px" value=" 批量更新排序 "  onClick="pxpost()">批量排序不需要选中
	  
	  </td>
    </tr>
    <tr >
      <td height="25" colspan="8" align="center"><%= strPageInfo%></td>
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
 
 
 function pxpost(){
    document.Article.method="post";
    document.Article.action="?Action=px";
	this.document.Article.submit();
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
		Set Rs=actcms.actexe("select * from Sitelink_ACT Where id="&id&"")
		If rs.eof Then
			Call actcms.alert("未知错误","")
		Else
			Title=Rs("Title")
			url=Rs("url")
			ifs=Rs("ifs")
			Id=Rs("Id")
			OpenType=Rs("OpenType")
			Num=Rs("Num")
			description=Rs("description")
			OrderID=Rs("OrderID")
			repset=Rs("repset")
 			If RS("repcontent") <> "" Then repcontent=Server.HTMLEncode(RS("repcontent"))
		End If
		A="EditSave"
	Else
		ifs=1:repset=1
		A="AddSave"	
		repcontent="<h1>{$content}</h1>"
	End If 
%>
<table width="100%" border="0" cellpadding="1" cellspacing="1" class="table">
<form name="form1" method="post" action="?action=<%= A %>&ID=<%= ID %>">
 

  <tr>
    <td height="35" align="right">替换属性：</td>
    <td height="35">
	<input type="radio"  onClick=repsets(1) id="repset22" name="repset" value="1"<%If repset="1" Then response.write "checked"%>><label for="repset22">替换关键字</label>            
	<input type="radio" onClick=repsets(0)  id="repset11" name="repset" value="0" <%If repset="0" Then response.write "checked"%>><label for="repset11">替换指定代码</label>
	<span class="h" style="cursor:help;"  onclick="dohelp('ACTDiyArticle_repset')"  id="ACTDiyArticle_repset">帮助</span></td>
  </tr>  
  

  <tr>
    <td width="30%" height="35" align="right">关键字名称：</td>
    <td width="63%" height="35"><input name="Title" type="text" class="Ainput" value="<%= Title %>" size=40>
	<span class="h" style="cursor:help;"  onclick="dohelp('ACTDiyArticle_Title')"  id="ACTDiyArticle_Title">帮助</span></td>
  </tr>
 

<tr id="repset1"
	<%If repset=1 Then response.write "style=""DISPLAY: none"""%>
	>
<td width="30%" height="35" align="right">替换代码：</td>
    <td width="63%" height="35"><textarea name="repcontent" cols="40" rows="5" id="repcontent"><%= repcontent %></textarea>
	要替换的代码请用 <font color=red><b>{$content}</b></font> 标签替换
	<span class="h" style="cursor:help;"  onclick="dohelp('ACTDiyArticle_repcontent')"  id="ACTDiyArticle_repcontent">帮助</span></td>
  </tr>

  
   <tr>
    <td height="35" align="right">替换次数：</td>
    <td height="35"><input name="Num" type="text" class="Ainput" value="<%= Num %>" size=40>最少是一次</td>
  </tr>

    <tr id="repset2">
    <td height="35" align="right">关键字链接：</td>
    <td height="35"><input name="url" type="text" class="Ainput" value="<%= url %>" size=40>
	<span class="h" style="cursor:help;"  onclick="dohelp('ACTDiyArticle_url')"  id="ACTDiyArticle_url">帮助</span></td>
  </tr>
  
   <tr >
    <td height="35" align="right">替换顺序：</td>
    <td height="35"><input name="OrderID" type="text" class="Ainput" value="<%= OrderID %>" size=40>数字越小,越先替换</td>
  </tr>
  
 
<tr id="repset3">
    <td height="35" align="right">链接描述：</td>
    <td height="35"><textarea name="description" cols="40" rows="5" id="description"><%= description %></textarea>
      不支持HTML</td>
  </tr>
  
  
  
<tr id="repset4">
    <td height="35" colspan="2" align="right"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="30%">&nbsp;</td>
        <td width="71%">	<%dim iftrue:iftrue=false%>
      <input name="OpenType" id="OpenType" type="text" class="Ainput" value="<%=OpenType%>" size="8">
      <select   name="OpenTypes"  onchange="document.form1.OpenType.value=this.value">
        <option value="_blank"   style="color:green"  <%If OpenType="_blank" Then response.write "selected":iftrue=true%>>新窗口打开</option>
        <option value="_parent" <%If OpenType="_parent" Then response.write "selected":iftrue=true%>>父窗口打开</option>
        <option value="_self" <%If OpenType="_self" Then response.write "selected":iftrue=true%>>本窗口打开</option>
        <option value="_top" <%If OpenType="_top" Then response.write "selected":iftrue=true%>>主窗口打开</option>
        <option value='' style="color:red"  <%If iftrue=false Then response.write "selected"%>>自定义</option>
      </select>
      <span class="h" style="cursor:help;"  onclick="dohelp('ACTDiyArticle_OpenType')"  id="ACTDiyArticle_OpenType">帮助</span>
      
</td>
      </tr>
    </table></td>
    </tr>

<tr id="repset5">
    <td height="35" align="right">是否开启：</td>
    <td height="35">
	<input type="radio" id="ifs1" name="ifs" value="1" <%If ifs=1 Then response.write "checked"%>><label for="ifs1">开启</label>
	<input type="radio" id="ifs2" name="ifs" value="0"<%If ifs=0 Then response.write "checked"%>><label for="ifs2">关闭</label>            
	<span class="h" style="cursor:help;"  onclick="dohelp('ACTDiyArticle_ifs')"  id="ACTDiyArticle_ifs">帮助</span></td>
  </tr>

  <tr>
    <td height="35" colspan="2" align="center"><input type=button onclick=CheckForm() class="ACT_btn"  name=Submit1 value="  保存  " />
        &nbsp;&nbsp;&nbsp;&nbsp;<input name="Submit2" type="reset" class="ACT_btn" value="  重置  ">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
    </tr>
</form>
</table>
<script type="text/javascript">
<!--
	function repsets(n){
	if (n == 1){
 		repset1.style.display='none';
 		repset2.style.display='';
 		repset3.style.display='';
 		repset4.style.display='';
 		repset5.style.display='';
  	}
	else{
		repset1.style.display='';
 		repset2.style.display='none';
 		repset3.style.display='none';
 		repset4.style.display='none';
 		repset5.style.display='none';
  	}
} 
//-->
</script>
<script language="javascript">repsets("<%=repset%>");</script>

<%end sub 


sub saves()
		dim Rs,RsSql
		 url=ACTCMS.S("url")
		 Title=ACTCMS.S("Title")
		 description=ACTCMS.S("description")
 		 ifs=ChkNumeric(ACTCMS.S("ifs"))
		 ID=ChkNumeric(ACTCMS.S("ID"))
		 Num=ChkNumeric(ACTCMS.S("Num"))
		 OrderID=ChkNumeric(ACTCMS.S("OrderID"))
			
		 repset=ChkNumeric(ACTCMS.S("repset"))
		 repcontent=ACTCMS.S("repcontent")
			
 		 OpenType=ACTCMS.S("OpenType")

		  
		 IF ACTCMS.S("Title") = "" Then
			Call ACTCMS.Alert("请添加关键字名称",""):Exit Sub
		 End If
		 
		If Action="AddSave" Then 
			 If Not ACTCMS.ACTEXE("SELECT Title FROM Sitelink_ACT Where Title='" & Title & "' order by ID desc").eof Then
				Call ACTCMS.Alert("系统已存在该关键字名称称!",""):Exit Sub
			 End if

			 Set Rs = Server.CreateObject("adodb.recordset")
			  RsSql = "select * from Sitelink_ACT"
			  Rs.Open RsSql, Conn, 1, 3
			  Rs.AddNew
		 	  Rs("Title") = Title
		 	  Rs("url") = url
			  Rs("ifs") =ifs
			  Rs("OrderID") =OrderID
			  Rs("description") =description
			  Rs("Num") =Num
			  Rs("OpenType") =OpenType
			  Rs("repset") =repset
			  Rs("repcontent") =repcontent
  
			  Rs.Update
			  Rs.Close:Set Rs = Nothing			
			  Call actcms.DelCahe("ReplaceSitelink")
			  Call Actcms.ActErr("操作成功","include/ACT.SiteLink.ASP","")
 		Else
		 	If Not ACTCMS.ACTEXE("SELECT Title FROM Sitelink_ACT Where ID <>" & ID & " AND  Title='" & Title & "' order by ID desc").eof Then
				Call ACTCMS.Alert("系统已存在该关键字名称!",""):Exit Sub
			 End if	
			 Set Rs = Server.CreateObject("adodb.recordset")
			  RsSql = "select * from Sitelink_ACT Where ID="&ID
			  Rs.Open RsSql, Conn, 1, 3
		 	  Rs("Title") = Title
		 	  Rs("url") = url
			  Rs("ifs") =ifs
			  Rs("OrderID") =OrderID
			  Rs("description") =description
			  Rs("Num") =Num
			  Rs("OpenType") =OpenType
			  Rs("repset") =repset
			  Rs("repcontent") =repcontent
			  Rs.Update
			  Rs.Close:Set Rs = Nothing		
			  Call actcms.DelCahe("ReplaceSitelink")
			  Call Actcms.ActErr("操作成功","include/ACT.SiteLink.ASP","")
 		End If 
end sub
CloseConn %>
<script language="javascript">
function CheckForm()
{ var form=document.form1;
	
	 if (form.Title.value=='')
		{ alert("请输入关键字!");   
		  form.url.focus();    
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
