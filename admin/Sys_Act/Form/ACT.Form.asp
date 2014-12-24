<!--#include file="../../ACT.Function.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>模型管理</title>
<link href="../../Images/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript" src="../../../ACT_inc/js/lhgcore/lhgcore.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../../ACT_inc/js/lhgcore/lhgdialog.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../../ACT_inc/js/lhgcore/Main.js"></script>

</head>
<%  dim Rs,ModeID,FieldSql,Action,FieldName,FieldType,ColumnType,TableName,Title,Text_Default,TitleSize,fun,ModeName
Dim IsNotNull,MultipleTextType_Width,MultipleTextType_Height,ID,Type_Default,FieldRS,Description,Type_Type,check
Dim NType_Default,RadioType_Content,ListBoxType_Content,Content,RadioType_Type,ListBoxType_Type,ISType,YHtml
Dim SupportHtmlType_Width,SupportHtmlType_Heigh,OrderID,RadioPic_Type,regex,regError,SearchIF,ValueOnly,savepic
Dim IsEditor
 ModeID = ChkNumeric(Request("ModeID"))
 if ModeID=0 or ModeID="" Then ModeID=1
 If Not ACTCMS.ChkAdmin() Then   Call Actcms.Alert("对不起，你没有操作权限！","")
 TableName=ACTCMS.actexe("select ModeTable from ModeForm_ACT Where ModeID=" & ModeID & " ")(0)
 ModeName=ACTCMS.actexe("select ModeName from ModeForm_ACT Where ModeID=" & ModeID & " ")(0)
 ID=ChkNumeric(Request("ID"))
	Action = Request.QueryString("A") 
	Select Case Action
		   Case "AddSave"
		   		Call AddSave()
			Case "ESave"
				Call Esave()
			Case "A","E"
				Call AddEdit()
			Case "px"
				Call px()
			Case Else
				Call Main()
	End Select



	IF Action = "Del" Then
		Dim FieldN
		FieldN=ACTCMS.ACTEXE("Select FieldName From Table_ACT Where actcms=3 and  ID=" & ID)(0)
		ACTCMS.ACTEXE("Delete from Table_ACT  Where actcms=3 and  ID=" & ID & "")
		 ACTCMS.ACTEXE("Alter Table "& TableName &" Drop column "& FieldN &"")
		Call Actcms.ActErr("删除字段成功","Sys_Act/Form/ACT.Form.asp?A=L&ModeID="&ModeID&"","")
 	End IF
	

	
	
	Sub px()
			Dim i
			ID = Split(actcms.s("ID"),","):OrderID = Split(actcms.s("OrderID"),",")
			 For I = LBound(ID) To UBound(ID)
					Conn.execute("Update Table_ACT set OrderID="&OrderID(I)&" where actcms=3 and  ID = "&ID(I)&"")
			Next 
			set conn=Nothing
		 Response.Redirect "?A=L&ModeID="&ModeID
	End sub

	Sub AddSave()
		FieldType=ACTCMS.S("FieldType")
		Title=ACTCMS.S("Title")
		IsNotNull=ChkNumeric(ACTCMS.S("IsNotNull"))
		ISType=ChkNumeric(ACTCMS.S("ISType"))
		OrderID=ChkNumeric(ACTCMS.S("OrderID"))
		Type_Default=ACTCMS.S("Type_Default")
		Description=ACTCMS.S("Description")
		savepic=ChkNumeric(ACTCMS.S("savepic"))
		YHtml=ChkNumeric(ACTCMS.S("YHtml"))
		IF ACTCMS.Chkchars(Request.Form("FieldName")) = False  Then
			Call Actcms.ActErr("英文名称只能为英文、数字及下划线","Sys_Act/Form/ACT.Form.asp?A=L&ModeID="&ModeID&"","")
 		Else
			FieldName=ACTCMS.S("FieldName")&"_ACT"
		End if
		Dim ActMode_Width,ActMode_Height
		'长度.宽度.
		Select Case FieldType
			Case "TextType"'单行文本
				ActMode_Width =  ChkNumeric(ACTCMS.S("TitleSize"))'文本框长度
				ColumnType="varchar(255)"
			Case "MultipleTextType"'多行文本(不支持Html
				ActMode_Width =  ChkNumeric(ACTCMS.S("MultipleTextType_Width"))
				ActMode_Height =  ChkNumeric(ACTCMS.S("MultipleTextType_Height"))
				Type_Type=YHtml
				 ColumnType="text"
			Case "MultipleHtmlType"'多行文本(支持Html)
				Content = ACTCMS.S("IsEditor")'编辑器属性放入Content字段
				ActMode_Width =  ChkNumeric(ACTCMS.S("SupportHtmlType_Width"))
				ActMode_Height =  ChkNumeric(ACTCMS.S("SupportHtmlType_Heigh"))
				Type_Type=savepic
				ColumnType="text"
			Case "RadioType"'单选项
			    Content = ACTCMS.S("RadioType_Content")
				Type_Type =  ChkNumeric(ACTCMS.S("RadioType_Type"))'显示方式
				ColumnType="varchar(255)"
			Case "ListBoxType"'多选项
			    Content = ACTCMS.S("ListBoxType_Content")
				Type_Type =  ChkNumeric(ACTCMS.S("ListBoxType_Type"))
				ColumnType="text"
			Case "NumberType"'数字
			    ActMode_Width =  ChkNumeric(ACTCMS.S("NumberType_TitleSize"))'数字的宽度放入总宽度字段名称中
				ColumnType="int"'
		   Case "DateType"
				 ColumnType="datetime"'Response.write "日期时间"
		   Case "NumberType"
				ColumnType="int"'Response.write "数字"
		  case "PicType"
				 Type_Type =ChkNumeric(ACTCMS.S("RadioPic_Type"))
				 				ColumnType="text"

		   Case else
		     ColumnType="varchar(255)"
		End Select 
	

		 Set FieldRS = Server.CreateObject("ADODB.RECORDSET")
		 FieldSql = "Select * From [Table_ACT] Where  FieldName='" & FieldName & "' And ModeID=" & ModeID
		 FieldRS.Open FieldSql, conn, 3, 3
 		 If FieldRS.EOF And FieldRS.BOF Then
			FieldRS.AddNew
			FieldRS("FieldName") = FieldName
			FieldRS("FieldType") = FieldType
			FieldRS("ModeID") = ModeID
			FieldRS("Title") = Title
			FieldRS("IsNotNull") = IsNotNull
			FieldRS("Width") = ActMode_Width
			FieldRS("Height") = ActMode_Height
			FieldRS("Type_Default") = Type_Default
			FieldRS("Description") = Description
			FieldRS("Type_Type") = Type_Type
			FieldRS("Content") = Content
			FieldRS("ISType") = ISType
			FieldRS("OrderID") = OrderID
			FieldRS("actcms") = 3
 			FieldRS("check") = ChkNumeric(ACTCMS.S("check"))		
			if ChkNumeric(ACTCMS.S("check"))	="1" then 
			FieldRS("regex") = ACTCMS.S("fun")
			else 
			FieldRS("regex") = ACTCMS.S("regex")
			FieldRS("regError") = ACTCMS.S("regError")
			end if 
 			FieldRS("SearchIF") = ChkNumeric(ACTCMS.S("SearchIF"))		
		    FieldRS("ValueOnly") = ChkNumeric(ACTCMS.S("ValueOnly"))		
		    FieldRS.Update
		 Conn.Execute("Alter Table "&TableName&" Add "&FieldName&" "&ColumnType&"")
		 Response.Write ("<Script> if (confirm('字段增加成功,继续添加吗?')) { location.href='?A=A&ModeID=" & ModeID& "';} else{location.href='?A=L&ModeID=" & ModeID&"';}</script>")
		 Else
		   Call ACTCMS.Alert("数据库中已存在该字段名称!", "")
		   Exit Sub
		 End If
	End  Sub 


	Sub Esave()
		Title=ACTCMS.S("Title")
		IsNotNull=ACTCMS.S("IsNotNull")
		Type_Default=ACTCMS.S("Type_Default")
		Description=ACTCMS.S("Description")
		OrderID=ChkNumeric(ACTCMS.S("OrderID"))
		savepic=ChkNumeric(ACTCMS.S("savepic"))
		YHtml=ChkNumeric(ACTCMS.S("YHtml"))
		If TitleSize=0 Then TitleSize=40
		 Set FieldRS = Server.CreateObject("ADODB.RECORDSET")
		 FieldSql = "Select * From [Table_ACT] Where ID=" & ID
		 FieldRS.Open FieldSql, conn,1, 3
			FieldRS("ModeID") = ModeID
			FieldRS("Title") = Title
			FieldRS("IsNotNull") = IsNotNull
			Select Case FieldRS("FieldType")
				Case "TextType"'单行文本
					FieldRS("Width") =  ChkNumeric(ACTCMS.S("TitleSize"))'文本框长度
				Case "MultipleTextType"'多行文本(不支持Html
					FieldRS("Width") =  ChkNumeric(ACTCMS.S("MultipleTextType_Width"))
					FieldRS("Height") =  ChkNumeric(ACTCMS.S("MultipleTextType_Height"))
					FieldRS("Type_Type") =YHtml
				Case "MultipleHtmlType"'多行文本(支持Html)
					FieldRS("Content") = ACTCMS.S("IsEditor")'编辑器属性放入Content字段
					FieldRS("Width") =  ChkNumeric(ACTCMS.S("SupportHtmlType_Width"))
					FieldRS("Height") =  ChkNumeric(ACTCMS.S("SupportHtmlType_Heigh"))
					FieldRS("Type_Type") =savepic
				Case "RadioType"'单选项
					FieldRS("Content") = ACTCMS.S("RadioType_Content")
					FieldRS("Type_Type") =  ChkNumeric(ACTCMS.S("RadioType_Type"))'显示方式
				Case "ListBoxType"'多选项
					FieldRS("Content") = ACTCMS.S("ListBoxType_Content")
					FieldRS("Type_Type") =  ChkNumeric(ACTCMS.S("ListBoxType_Type"))		
				Case "NumberType"'数字
					FieldRS("Width") =  ChkNumeric(ACTCMS.S("NumberType_TitleSize"))'数字的宽度放入总宽度字段名称中
				case "PicType"
					 FieldRS("Type_Type") = ChkNumeric(ACTCMS.S("RadioPic_Type"))
 			End Select 
			FieldRS("Description") = Description
			FieldRS("Type_Default") = Type_Default
			FieldRS("ISType")= ChkNumeric(ACTCMS.S("ISType"))
			FieldRS("OrderID") = OrderID
 			FieldRS("check") = ChkNumeric(ACTCMS.S("check"))		
			if ChkNumeric(ACTCMS.S("check"))	="1" then 
			FieldRS("regex") = ACTCMS.S("fun")
			else 
			FieldRS("regex") = ACTCMS.S("regex")
			FieldRS("regError") = ACTCMS.S("regError")
			end if 
			FieldRS("actcms") = 3
			FieldRS("SearchIF") = ChkNumeric(ACTCMS.S("SearchIF"))		
 		    FieldRS("ValueOnly") = ChkNumeric(ACTCMS.S("ValueOnly"))		

			 FieldRS.Update
			 Call Actcms.ActErr("字段修改成功","Sys_Act/Form/ACT.Form.asp?A=L&ModeID=" & ModeID& "","")
 	End  Sub 

 Sub Main() %>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td class="bg_tr">您现在的位置：后台管理 >> <a href="index.asp">模型列表</a> >> <a href="?A=L&ModeID=<%= ModeID %>">字段列表</a> >> 浏览[<%= ModeName %>]字段 </td>
  </tr>
  <tr>
    <td>当前模型：[<%= ModeName %>]&nbsp;&nbsp; <a href="?A=A&ModeID=<%= ModeID %>">添加字段</a> |  
 </td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
   <form name="Form1" method="post" action="?A=px&ModeID=<%= ModeID %>">
 <tr>
    <td align="center" class="bg_tr">字段名称</td>
    <td align="center" class="bg_tr">字段别名</td>
	<td align="center" class="bg_tr">调用标签</td>
    <td align="center" class="bg_tr">字段类型</td>
    <td align="center" class="bg_tr">是否必填</td>
 	<td align="center" class="bg_tr">排序</td>
    <td width="100" align="center" class="bg_tr">管理操作</td>
  </tr>
<% 
	  Set Rs =ACTCMS.ACTEXE("SELECT * FROM Table_ACT Where actcms=3 and  ModeID=" & ModeID & " order by OrderID desc,ID Desc")
	 If Rs.EOF  Then
	 	Response.Write	"<tr><td colspan=""8"" align=""center"">没有记录</td></tr>"
	 Else
		Do While Not Rs.EOF	
			 %>
  <tr   onMouseOver=overColor(this) onMouseOut=outColor(this)>
    <td align="center"><%= Rs("FieldName") %></td>
    <td align="center"><%= Rs("Title") %></td>
	<td align="center"><font color=red>{$<%= Rs("FieldName") %>} </font></td>
    <td align="center"><%
	
		 Select Case Rs("FieldType")
		   Case "TextType"
				Response.write "单行文本"
		   Case "MultipleTextType"
				Response.write "多行文本(不支持Html)"
		   Case "MultipleHtmlType"
				Response.write "多行文本(支持Html)"
		   Case "RadioType"
				Response.write "单选项"
		   Case "ListBoxType"
				Response.write "多选项"
		   Case "DateType"
				Response.write "日期时间"
		   Case "PicType"
				Response.write "图片"
		   Case "FileType"
				Response.write "文件"
		   Case "NumberType"
				Response.write "数字"
		   Case "RadomType"
				Response.write "随机数"
		   Case else
				Response.write "<font color=red>该字段错误</font>"
		 End Select
%></td>
    <td align="center"><%if Rs("IsNotNull")=0 Then Response.Write "是" Else Response.Write "否" %></td>
 	<td align="center">
          <input name="OrderID" type="text"  class="Ainput"  id="OrderID" value="<%=rs("OrderID")%>" size="4" maxlength="3">
          <input name="A" type="hidden" id="A" value="N">
		  <input name="ID" type="hidden" id="ID" value="<%=rs("ID")%>">	
           
	</td>
	<td align="center"><a href="?A=E&ID=<%=Rs("ID")%>&ModeID=<%=Rs("ModeID")%> " >修改</a> ┆ 

<a href="?A=Del&ID=<%=Rs("ID")%>&ModeID=<%=Rs("ModeID")%> " onClick="{if(confirm('确定删除该字段吗?')){return true;}return false;}">删除</a>  </td>
  </tr>
  <% 
		Rs.movenext
		Loop
	End if	 %>

 <tr >
    <td colspan="8"><input type="submit" Class="ACT_BTN" name="Submit" value=" 批量更新排序 "></td>
  </tr>
</form>
</table>
<% End Sub 
   Sub AddEdit()
   Dim A
	if Action="A" Then
		A="AddSave"
		TitleSize="40"
		ISType="1"
		RadioPic_Type="0":check=3
		RadioType_Content="名称-值"& vbCrLf &"名称-值"& vbCrLf &"名称-值"
		ListBoxType_Content="名称-值"& vbCrLf &"名称-值"& vbCrLf &"名称-值"
		SupportHtmlType_Width=620
		SupportHtmlType_heigh=400
		MultipleTextType_Width=300
		MultipleTextType_Height=100
		IsEditor="Simple"
		IsNotNull=1
		OrderID=10
		YHtml=0
		savepic=0
	Else
		Set FieldRS=server.CreateObject("adodb.recordset") 
		FieldRS.OPen "Select * from Table_ACT Where ID = "&ID&" order by ID desc",Conn,1,1
		FieldName = FieldRS("FieldName")
		FieldType = FieldRS("FieldType")  
		ModeID = FieldRS("ModeID") 
		Title = FieldRS("Title")  
		Type_Default = FieldRS("Type_Default") 
		Description = FieldRS("Description") 
		OrderID = FieldRS("OrderID") 
		Select Case FieldType
			Case "TextType"'单行文本
				TitleSize=FieldRS("Width") 
			Case "MultipleTextType"'多行文本(不支持Html
				MultipleTextType_Width=FieldRS("Width") 
				MultipleTextType_Height=FieldRS("Height")
				YHtml=FieldRS("Type_Type")
 			Case "MultipleHtmlType"'多行文本(支持Html)
				IsEditor=FieldRS("Content")
				SupportHtmlType_Width=FieldRS("Width") 
				SupportHtmlType_Heigh=FieldRS("Height")
				savepic=FieldRS("Type_Type")
			Case "RadioType"'单选项
				RadioType_Content=FieldRS("Content")
				RadioType_Type = FieldRS("Type_Type")
			case "PicType"
			    RadioPic_Type = FieldRS("Type_Type")
 			Case "ListBoxType"'多选项
				ListBoxType_Content=FieldRS("Content")
				ListBoxType_Type = FieldRS("Type_Type")
			Case "NumberType"'数字
				TitleSize=FieldRS("Width") 
		End Select 
		ISType=FieldRS("ISType")
		IsNotNull=FieldRS("IsNotNull")
	    check=FieldRS("check")

		if  FieldRS("check")="1" then 
 		fun=FieldRS("regex")
		else
 		regex=FieldRS("regex")
		end if 
		regError=FieldRS("regError")
		SearchIF=FieldRS("SearchIF")
	    ValueOnly=FieldRS("ValueOnly")

		A="ESave"
 	End  If 
    %><body  onload="SelectModelType()">

	<form  name="tcjdxr" method="post" action="?A=<%= A %>&ModeID=<%= Request.QueryString("ModeID") %>&ID=<%= Request.QueryString("ID") %>">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td class="bg_tr">您现在的位置：后台管理 > <a href="index.asp">模型列表</a> > [<%= ModeName  %>] > <a href="?A=L&ModeID=<%= ModeID %>">字段列表</a>  > 添加字段 </td>
  </tr>
</table>
<table class="table" cellspacing="1" cellpadding="2" width="98%" border="0" align="center">
  <tr >
    <td width="15%" align="right"> 字段别名： </td>
    <td><input name="Title" value="<%=Title%>"  type="text"  class="Ainput"  maxlength="20" id="Title" />
    <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('mode_zdbm')"  id="mode_zdbm">帮助</span><font color="#ff0066">*</font> 如：文章标题</td>
  </tr>
  <tr >
    <td align="right"> 字段名称： </td>
    <td><input name="FieldName" value="<%=FieldName%>" <%If action="E" Then Response.write " disabled=""disabled"" "%>  type="text"  class="Ainput"  maxlength="50" id="FieldName" />
       <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('mode_zdmc')">帮助</span>
	   <%If action="E" Then Response.write " 该字段在 <"&ModeName&"> 模型内容页调用标签 <font color=red>{$"&FieldName&"}</font>"%> <br> <font color="#ff0066">*</font>为了和系统字段区分，系统创建字段时会自动以“_ACT”结尾,在模板中可以通过“{$字段名称_ACT}”进行调用 </td>
  </tr>
  <tr >
    <td align="right"> 字段描述： </td>
    <td><textarea name="Description" rows="6" cols="40" id="Description"><%=Description%></textarea>
	<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('mode_zdms')"  id="mode_zdms">帮助</span></td>
  </tr>
  <tr >


  <td align="right"> 是否必填： </td>
    <td><table id="IsNotNull" border="0">
      <tr>
        <td><input id="IsNotNull_0" <% IF IsNotNull = "0" Then Response.Write "Checked" %> type="radio" name="IsNotNull" value="0" /><label for="IsNotNull_0">是</label></td>
        <td><input id="IsNotNull_1" <% IF IsNotNull = "1" Then Response.Write "Checked" %> type="radio" name="IsNotNull" value="1" /><label for="IsNotNull_1">否</label>
			  <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('mode_sfbt')"  id="mode_sfbt">帮助</span></td>
      </tr>
    </table></td>
  </tr>
<!--   <tr >
    <td align="right"> 会员中心调用：</td>
    <td><table id="ISType"  border="0">
      <tr>
        <td><input id="ISType_0" <% IF ISType = "1" Then Response.Write "Checked" %>   type="radio" name="ISType" value="1" /><label for="ISType_0">是</label></td>
        <td><input id="ISType_1" type="radio" <% IF ISType = "0" Then Response.Write "Checked" %> name="ISType" value="0" /><label for="ISType_1">否</label>
			  <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('mode_hyzxdy')">帮助</span>
			  
			  </td>
      </tr>
    </table></td>
  </tr> -->
   <tr >
    <td align="right"> 字段排序： </td>
    <td><input name="OrderID" value="<%=OrderID%>"  type="text"  class="Ainput"  maxlength="20" id="OrderID" />
   数字越大,排的越前</td>
  </tr>
 
 
 
   <tr >
    <td align="right">数据校验规则：</td>
    <td>
      <label for="check3"><input type="radio" name="check"   <% IF check = "3" Then Response.Write "Checked" %>  id="check3"  onClick=chk(3)  value="3">
      默认 </label>
    <label for="check1"><input type="radio" name="check"  <% IF check = "1" Then Response.Write "Checked" %>  id="check1"  onClick=chk(1)  value="1">
      函数 </label>
       <label for="check2"> <input type="radio" name="check"  <% IF check = "2" Then Response.Write "Checked" %>  id="check2" onClick=chk(2)  value="2">
      正则</label><span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('mode_sjjygz')"  id="mode_sjjygz">帮助</span></td>
  </tr> 
 
 
   <tr id="checks3">
    <td align="right">函数名称：</td>
    <td><input name="fun" type="text"  class="Ainput"  id="fun" size="40" value="<%= fun %>">请在Field.asp中自己增加
	<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('mode_hsmc')"  id="mode_hsmc">帮助</span>
  	</td>
  </tr> 
  
   <tr id="checks1">
    <td align="right">数据校验正则：</td>
    <td><input name="regEx" type="text"  class="Ainput"  id="regEx" size="40" value="<%= regEx %>">
 	<select name="select"   onchange="document.tcjdxr.regEx.value=this.value">
          <option selected>-- 常用正则 --</option>
			<option value="^[A-Za-z]+$">英文</option> 
			<option value="^[\u0391-\uFFE5]+$">中文</option> 
<!--			<option value="^[a-z]\w{2,19}$">中英文</option> 
-->			<option value="^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$">email</option> 
			<option value="^[-\+]?\d+$">整型</option> 
          <option value="^\d+$">数字</option> 
			<option value="^[-\+]?\d+(\.\d+)?$">double</option> 
			<option value="^[1-9]\d{4,9}$">qq</option> 
			<option value="^((\(\d{2,3}\))|(\d{3}\-))?(\(0\d{2,3}\)|0\d{2,3}-)?[1-9]\d{6,7}(\-\d{1,4})?$">phone</option> 
			<option value="^((\(\d{2,3}\))|(\d{3}\-))?(1[35][0-9]|189)\d{8}$">mobile</option> 
			<option value="^(http|https|ftp):\/\/[A-Za-z0-9]+\.[A-Za-z0-9]+[\/=\?%\-&_~`@[\]\':+!]*([^<>\])*$">网址</option> 
			<option value="^[A-Za-z0-9\-]+\.([A-Za-z]{2,4}|[A-Za-z]{2,4}\.[A-Za-z]{2})$">域名</option> 
			<option value="^(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5]).(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5]).(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5]).(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5])$">ip</option> 
               </select>
 	<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('mode_zz')"  id="mode_zz">帮助</span></td>
  </tr>
     <tr   id="checks2">
    <td align="right">检验错误提示信息：</td>
    <td><input name="regError" type="text"  class="Ainput"  id="regError" size="40"  value="<%= regError %>">
	
	<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('mode_errorhelp')"  id="mode_errorhelp">帮助</span>
	
	</td>
  </tr>



  <tr >
    <td align="right">作为搜索条件：</td>
    <td>
	
<input id="SearchIF_0" <% IF SearchIF = "0" Then Response.Write "Checked" %>   type="radio" name="SearchIF" value="0" />
<label for="SearchIF_0">是</label>	

<input id="SearchIF_1" <% IF SearchIF = "1" Then Response.Write "Checked" %>   type="radio" name="SearchIF" value="1" />
<label for="SearchIF_1">否</label>	
	<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('mode_ss')"  id="mode_ss">帮助</span>
	</td>
  </tr>


  <tr >
    <td align="right">数据唯一：</td>
    <td>
	
<input id="ValueOnly_0" <% IF ValueOnly = "0" Then Response.Write "Checked" %>   type="radio" name="ValueOnly" value="0" />
<label for="ValueOnly_0">是</label>	

<input id="ValueOnly_1" <% IF ValueOnly = "1" Then Response.Write "Checked" %>   type="radio" name="ValueOnly" value="1" />
<label for="ValueOnly_1">否</label>	
	<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('mode_sjwy')"  id="mode_sjwy">帮助</span>
	</td>
  </tr>

  
  
  <tr >
    <td align="right"> 字段类型： </td>
    <td>	
	
<table id="FieldType" <%If action="E" Then Response.write " disabled=""disabled"" "%> onClick="SelectModelType()" border="0">
	<tr>
		<td><input id="Type_0" <% IF FieldType = "TextType" Then Response.Write "Checked" %> type="radio" name="FieldType" value="TextType" checked="checked" />
		<label for="Type_0">单行文本</label></td>
		<td><input id="Type_1" <% IF FieldType = "MultipleTextType" Then Response.Write "Checked" %>  type="radio" name="FieldType" value="MultipleTextType" />
		<label for="Type_1">多行文本(不支持Html)</label></td><td>
		<input id="Type_2" <% IF FieldType = "MultipleHtmlType" Then Response.Write "Checked" %>   type="radio" name="FieldType" value="MultipleHtmlType" />
		<label for="Type_2" <% IF FieldType = "SupportHtmlType" Then Response.Write "Checked" %> >多行文本(支持Html)</label></td><td>
		<input id="Type_3" <% IF FieldType = "RadioType" Then Response.Write "Checked" %>  type="radio" name="FieldType" value="RadioType" /><label for="Type_3">单选项</label>
		</td>
		<td><input id="Type_4" <% IF FieldType = "ListBoxType" Then Response.Write "Checked" %>  type="radio" name="FieldType" value="ListBoxType" /><label for="Type_4">多选项</label></td>
	</tr><tr>
		<td><input id="Type_5" <% IF FieldType = "DateType" Then Response.Write "Checked" %>  type="radio" name="FieldType" value="DateType" />
		<label for="Type_5">日期时间</label></td><td>
		<input id="Type_6"  <% IF FieldType = "PicType" Then Response.Write "Checked" %> type="radio" name="FieldType" value="PicType" />
		<label for="Type_6">上传</label></td><td><input <% IF FieldType = "FileType" Then Response.Write "Checked" %>  id="Type_7" type="radio" name="FieldType" value="FileType" />
		<label for="Type_7">下载</label></td><td><input <% IF FieldType = "NumberType" Then Response.Write "Checked" %>  id="Type_8" type="radio" name="FieldType" value="NumberType" />
		<label for="Type_8">数字</label></td><td><input <% IF FieldType = "RadomType" Then Response.Write "Checked" %>  id="Type_9" type="radio" name="FieldType" value="RadomType" />
		<label for="Type_9">随机数</label>
		<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('mode_zdlx')"  id="mode_zdlx">帮助</span></td>
	</tr>
</table>	</td>
  </tr>
  <tbody id="DivTextType">
    <tr>
      <td align="right">文本框长度：</td>
      <td><input name="TitleSize" value="<%=TitleSize%>"  type="text"  class="Ainput"   maxlength="4" size="10" id="TitleSize" /></td>
    </tr>
  </tbody>
  <tbody id="DivMultipleTextType" style="display:none">
    <tr>
      <td align="right">显示的宽度：</td>
      <td><input name="MultipleTextType_Width" type="text"  class="Ainput"  value="<%=MultipleTextType_Width%>" maxlength="4" size="10" id="MultipleTextType_Width" />
        px</td>
    </tr>
    <tr>
      <td align="right">显示的高度：</td>
      <td><input name="MultipleTextType_Height" type="text"  class="Ainput"  value="<%=MultipleTextType_Height%>" maxlength="4" size="10" id="MultipleTextType_Height" />
        px</td>
    </tr>
    
    <tr>
      <td align="right">是否允许HTML：</td>
      <td><input id="YHtml_0" <% IF YHtml = "0" Then Response.Write "Checked" %>   type="radio" name="YHtml" value="0" />
<label for="YHtml_0">是</label>	

<input id="YHtml_1" <% IF YHtml = "1" Then Response.Write "Checked" %>   type="radio" name="YHtml" value="1" />
<label for="YHtml_1">否</label>	
<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('mode_yxhtml')"  id="mode_yxhtml">帮助</span>
</td>
    </tr>    
    
  </tbody>
  <tbody id="DivMultipleHtmlType" style="display:none">
    <tr>
      <td align="right">编辑器菜单名称：</td>
      <td>       
<input name="IsEditor" type="text"  class="Ainput"  value="<%=IsEditor%>" maxlength="4" size="10" id="IsEditor" />
<select name="select" onChange="FormatTitle(this, tcjdxr.IsEditor, '')">
          <option selected>-- 请选择 --</option>
          <option value="1">简洁</option>
          <option value="2">超简洁</option>
          <option value="3">全部</option>
        </select>如需自己定义菜单名称,请手工配置editor/fckeditor/fckconfig.js 文件</td>
    </tr>
    <tr>
      <td align="right">显示的宽度：</td>
      <td><input name="SupportHtmlType_Width" type="text"  class="Ainput"  value="<%=SupportHtmlType_Width%>" maxlength="4" size="10" id="SupportHtmlType_Width" />
        px</td>
    </tr>
    <tr>
      <td align="right">显示的高度：</td>
      <td><input name="SupportHtmlType_Heigh" type="text"  class="Ainput"  value="<%=SupportHtmlType_Heigh%>" maxlength="4" size="10" id="SupportHtmlType_Heigh" />
        px</td>
    </tr>
    
    <tr>
      <td align="right">是否保存远程图片：</td>
      <td><input id="savepic_0" <% IF savepic = "0" Then Response.Write "Checked" %>   type="radio" name="savepic" value="0" />
<label for="savepic_0">是</label>	

<input id="savepic_1" <% IF savepic = "1" Then Response.Write "Checked" %>   type="radio" name="savepic" value="1" />
<label for="savepic_1">否</label>	</td>
    </tr>      
  </tbody>
  <tbody id="DivRadioType" style="display:none">
    <tr>
      <td align="right">分行键入每个选项：</td>
      <td><textarea name="RadioType_Content" rows="6" cols="40" id="RadioType_Content"><%=RadioType_Content%></textarea>
	  <font color=red>注意 要按照格式书写 名称-值, 以 - 隔开,列:合肥-HeFei</font></td></tr>
    <tr>
      <td align="right">显示选项：</td>
      <td><table id="RadioType_Type" border="0">
        <tr>
          <td>
		  <input id="RadioType_Type_0"  <% IF RadioType_Type = "0" Then Response.Write "Checked" %>  type="radio" name="RadioType_Type" value="0" checked="checked" />
                <label for="RadioType_Type_0">单选下拉列表框</label></td>
        </tr>
        <tr>
          <td><input id="RadioType_Type_1" <% IF RadioType_Type = "1" Then Response.Write "Checked" %>  type="radio" name="RadioType_Type" value="1" />
                <label for="RadioType_Type_1">单选按钮</label></td>
        </tr>
      </table></td>
    </tr>
  </tbody>
  <tbody id="DivListBoxType" style="display:none">
    <tr>
      <td align="right">分行键入每个选项：</td>
      <td><textarea name="ListBoxType_Content" rows="6" cols="40" id="ListBoxType_Content"><%=ListBoxType_Content%></textarea></td></tr>
    <tr>
      <td align="right">显示选项：</td>
      <td><table id="ListBoxType_Type" border="0">
        <tr>
          <td><input id="ListBoxType_Type_0"  <% IF ListBoxType_Type = "0" Then Response.Write "Checked" %>  type="radio" name="ListBoxType_Type" value="0" checked="checked" />
                <label for="ListBoxType_Type_0">复选框</label></td>
        </tr>
        <tr>
          <td><input id="ListBoxType_Type_1"  <% IF ListBoxType_Type = "1" Then Response.Write "Checked" %>  type="radio" name="ListBoxType_Type" value="1" />
                <label for="ListBoxType_Type_1">多选列表框</label></td>
        </tr>
      </table></td>
    </tr>
  </tbody>
  <tbody id="DivDateType" style="display:none">
  </tbody>
  <tbody id="DivPicType" style="display:none">
    <tr>
      <td align="right">显示选项：</td>
      <td><table id="RadioPic_Type" border="0">
        <tr>
          <td>
		  
     <input id="RadioPic_Type_0"    type="radio" name="RadioPic_Type" value="0"  <% IF RadioPic_Type = "0"  Then Response.Write "Checked" %>    />
                <label for="RadioPic_Type_0">单个文件上传</label></td>
        </tr>
        <tr>
          <td>
          <input id="RadioPic_Type_1"   type="radio" name="RadioPic_Type" value="1"   <% IF RadioPic_Type = "1" Then Response.Write "Checked" %>  />
                <label for="RadioPic_Type_1">多个文件上传</label></td>
        </tr>
      </table></td>
    </tr>  
  
  
  </tbody>
  <tbody id="DivRadomType" style="display: none">
  </tbody>
  <tbody id="DivFileType" style="display:none">
  </tbody>
  <tbody id="DivNumberType" style="display:none">
    <tr>
      <td align="right">文本框长度：</td>
      <td><input name="NumberType_TitleSize" type="text"  class="Ainput"  value="40" maxlength="4" size="10" id="NumberType_TitleSize" /></td>
    </tr>
  </tbody>
  <tr>
	<Td align="right">默认值：</td>
      <td><input name="type_Default" type="text"  class="Ainput"  value="<%=type_Default%>" size="10" id="NType_Default" />
	  <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('mode_mrz')"  id="mode_mrz">帮助</span>
	  注：没有数据录入的默认值，与前台显示无关.</td>
  </tr>

  <tr>
    <td></td>
    <td height="50">
	  <input type=button onclick=CheckForm() class="ACT_btn"  name=Submit1 value="  保存字段  " />
      &nbsp;&nbsp; <input type="reset" class="ACT_btn" name="Submit2" value="  重置  ">	</td>
  </tr>
</table>
	</form>
<script language="JavaScript" type="text/javascript">
function FormatTitle(obj, obj2, def_value)
{
    var FormatFlag = obj.options[obj.selectedIndex].value;
    var tmp_Title = FilterHtmlStr(obj2.value);
    switch(FormatFlag)
    {
        case "1" :
            obj2.value = "UserMode";
            break;
        case "2" :
            obj2.value = "Simple";
            break;
        case "3" :
            obj2.value = "Default";
            break;
    }
    obj.selectedIndex = 0;
}
function FilterHtmlStr(str)
{
    str = str.replace(/<.*?>/ig, "");
    return str;
}
function SelectModelType()
{
    var TypeCount=document.getElementsByName("FieldType"); 
    
    for(var i=1;i<TypeCount.length;i++)
    { 
        var DivType=eval("Div"+TypeCount[i].value);
        
        if(TypeCount[i].checked)
        {
            DivType.style.display="";
        }
        else
        {
            DivType.style.display="none";
        }
    }
}


function chk(n){
	if (n == "1"){
		checks1.style.display='none';
		checks2.style.display='none';
		checks3.style.display='';
 	}
	else  if (n == "2"){
		checks1.style.display='';
		checks2.style.display='';
		checks3.style.display='none';
  	}
	else
	{
		checks1.style.display='none';
		checks2.style.display='none';
		checks3.style.display='none';
 	}
}




	function CheckForm()
	{ var form=document.tcjdxr;
	  
	 if (form.Title.value=='')
		{ alert("字段别名不能够为空！");   
		  form.Title.focus();    
		   return false;
		}
	 if (form.FieldName.value=='')
		{ alert("字段名称不能够为空！!");   
		  form.FieldName.focus();    
		   return false;
		}
	    form.Submit1.value="正在提交数据,请稍等...";
		form.Submit1.disabled=true;	
		form.Submit2.disabled=true;	
	    form.submit();
        return true;
	}	
</script>
    <script language="javascript">chk("<%= check %>");</script>


<%  End Sub



%>


<SCRIPT LANGUAGE="JavaScript">
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
