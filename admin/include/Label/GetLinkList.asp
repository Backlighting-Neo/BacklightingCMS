<!--#include file="../../ACT.Function.asp"-->
<!--#include file="../ACT.F.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>1</title>
<link href="../../Images/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript" src="../../../ACT_inc/js/lhgcore/lhgcore.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../../ACT_inc/js/lhgcore/lhgdialog.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../../ACT_inc/js/lhgcore/Main.js"></script>
</head>
<body>
<%
Dim Action,ID,LabelRS,LabelName,Descript,LabelContent,LabelFlag,LabelContentArr,ClassID,Rs,pages
Dim ListNumber,TitleLen,ColNumber,LogoHeight,LogoWidth,LinkType,TypeStyle,ActF,LinkContent,iftrue,LinkSort
If Not ACTCMS.ACTCMS_QXYZ(0,"bqxt","") Then   Call Actcms.Alert("对不起，你没有操作权限！","")
Action = Request.QueryString("Action")
ID =  ChkNumeric(Request.QueryString("ID"))
IF Action = "Add" Then
	ClassID = 0
	ListNumber = 10
	ColNumber = 7
	TitleLen=30
	 ActF=1
 	LinkType=2:TypeStyle=1:LogoHeight=31:LogoWidth=88
	pages = "新建友情链接列表标签"
Else
	  
	  Set LabelRS = Server.CreateObject("Adodb.Recordset")
	  LabelRS.Open "Select * From Label_Act Where ID=" & ID & "", Conn, 1, 1
	  If LabelRS.EOF And LabelRS.BOF Then
		 LabelRS.Close
		 Conn.Close:Set Conn = Nothing
		 Set LabelRS = Nothing
		 Response.Write "参数传递出错!":Response.End
	  End If
		LabelName = Replace(Replace(LabelRS("LabelName"), "{ACTCMS_", ""), "}", "")
		Descript = LabelRS("Description")
		LabelContent = LabelRS("LabelContent")
		LabelFlag = Clng(LabelRS("LabelFlag"))
		LabelRS.Close:Set LabelRS = Nothing
		LabelContent = Replace(Replace(LabelContent, "{$GetLinkList(", ""), ")}", "")
		LabelContent = Replace(LabelContent, """", "") 
		LabelContentArr = Split(LabelContent, "§")
		ClassID = LabelContentArr(0)
		LinkType = LabelContentArr(1)'是否包含子栏目
		TypeStyle = LabelContentArr(2)'列出条数
		LogoWidth = LabelContentArr(3)'链接目标
		LogoHeight = LabelContentArr(4)
		ListNumber= LabelContentArr(5)
		TitleLen = LabelContentArr(6)
		ColNumber= LabelContentArr(7)
		ActF= LabelContentArr(8)
		LinkSort= LabelContentArr(9)
		LinkContent= LabelContentArr(10)
		pages = "修改友情链接列表标签"
End IF
 %>
<form id="myform" name="myform" method="post" action="AddLabelSave.asp">
 <input type="hidden" name="LabelContent"> 
  <input type="hidden" name="Action" value="<%= Action %>"> 
  <input type="hidden" name="ID" value="<%= ID %>"> 
 <input type="hidden" name="FileUrl" value="GetLinkList.asp">  
 <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
    <tr>
      <td colspan="2" align="center" class="bg_tr"><%= pages %>&nbsp;</td>
    </tr>
    <tr>
      <td width="50%" >标签名称
      <input name="LabelName"   type="text"  class="Ainput" id="LabelName" value="<%= LabelName %>"></td>
      <td width="50%" ><font color="red">* 调用格式"{ACTCMS_标签名称}"</font></td>
    </tr>
   <tr>
      <td >标签目录      
        <select name="LabelFlag" id="select">
          <option value="0">系统默认</option>
			 <%=AF.ACT_LabelFolder(CInt(LabelFlag))%>
        </select>&nbsp;&nbsp;<a href="ACT.LabelFolder.asp"><font color=red><b>新建存放目录</b></font></a></td>
	 <td width="50%" height="24"  ><font color=green>标签存放目录,日后方便管理标签</font></td>
    </tr>
    <tr>
      <td >
	  链接类别    <Select Name="ClassID" >  
	  <% 			 response.Write ("<option Value=""0"">-列出所有分类的站点-</option>")
					Dim LinkRs
					Set LinkRs = Conn.Execute("Select id,ClassLinkName From ClassLink_Act")
					 Do While Not LinkRs.EOF
					   If ClassID = CStr(LinkRs(0)) Then
						response.Write ("<Option value=" & LinkRs(0) & " selected>" & LinkRs(1) & "</OPTION>")
					   Else
						response.Write ("<Option value=" & LinkRs(0) & ">" & LinkRs(1) & "</OPTION>")
					   End If
					   LinkRs.MoveNext
					 Loop
					 LinkRs.Close
					 Set LinkRs = Nothing
					 %></Select>
                  <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_linkClassID')"  id="Label_linkClassID">帮助</span>     
                     
                     </td>
      <td >
      
      
排序方法  
	<input name="LinkSort" id="LinkSort"  type="text"  class="Ainput" value="<%=LinkSort%>" size="20">

	  
	<select   name="LinkSorts"  onchange="document.myform.LinkSort.value=this.value">
		<option value='ID Desc' <%If LinkSort="ID Desc" Then Response.write "selected":iftrue=true%>>链接ID(降序)</option>
		<option value='ID Asc' <%If LinkSort="ID Asc" Then Response.write "selected":iftrue=true%>>链接ID(升序)</option>
		<option value='AddDate Asc' <%If LinkSort="AddDate Asc" Then Response.write "selected":iftrue=true%>>更新时间(升序)</option>
		<option value='AddDate Desc' <%If LinkSort="AddDate Desc" Then Response.write "selected":iftrue=true%>>更新时间(降序)</option>
 		<option value='' style="color:red"  <%If iftrue=false Then Response.write "selected"%> >自定义</option>	
		</select>       
      
       <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_LinkSort')"  id="Label_LinkSort">帮助</span>	  </td>
    </tr>
 
    <tr>
      <td > 
      
  输出模式
		 <select  style='width:40%' name="ActF" id="ActF" onChange="SetActF(this.options[this.selectedIndex].value);"> 
	 <option value="1" <% IF ActF = 1 Then Response.Write("selected") %>>普通模式</option>
  <option value="2" <% IF ActF = 2 Then Response.Write("selected") %>>代码模式</option>
  </select>  <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_ActF')"  id="Label_ActF">帮助</span>	    </td>
      <td > 显示数目
      <input name="ListNumber"  type="text"  class="Ainput" id="ListNumber" value="<%= ListNumber %>">
      <font color="red">设置为0时将列出所有友情链接站点</font>
      <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_linkListNumber')"  id="Label_linkListNumber">帮助</span> 
      </td>
    </tr> 
 
 
  <tr id=ActFs ><td  colspan="2" >

<font color=red>内置标签</font> 
<a href="#" onClick='SetLinkContent(LinkContent,"#Link")'>链接地址</a>&nbsp;
<a href="#" onClick='SetLinkContent(LinkContent,"#Title")'>链接名称</a>&nbsp;
<a href="#" onClick='SetLinkContent(LinkContent,"#Logo")'>链接LOGO</a>&nbsp;
 <a href="#" onClick='SetLinkContent(LinkContent,"#Description")'>网站简介</a>&nbsp;

 <br />
<textarea onFocus="this.className='colorfocus';" onBlur="this.className='colorblur';"  name="LinkContent"  id="DiyContent"  cols="95%" rows="10"><%=Server.HTMLEncode(LinkContent)%></textarea>
<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_LinkContent')"  id="Label_LinkContent">帮助</span> </td>
	</tr> 
  
    <tr>
      <td >链接类型
        <INPUT onClick="SetLogo(2)" <% IF LinkType = 2 Then  Response.Write "Checked" %> type="radio" value="2" name="LinkType" id="LinkType1"><label for="LinkType1">全部链接</label>
        <INPUT onClick="SetLogo(0)" <% IF LinkType = 0 Then  Response.Write "Checked" %> type="radio" value="0" name="LinkType" id="LinkType2"><label for="LinkType2">文本链接</label>
        <INPUT onClick="SetLogo(1)" <% IF LinkType = 1 Then  Response.Write "Checked" %> type="radio" value="1" name="LinkType" id="LinkType3"><label for="LinkType3">LOGO链接</label>  
   <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_LinkType')"  id="Label_LinkType">帮助</span>      
        
            </td>
      <td >显示方式
        <INPUT onClick="SetType(0)" <% IF TypeStyle = 0 Then  Response.Write "Checked" %>  type="radio" value="0" name="TypeStyle"  id="TypeStyle1"><label for="TypeStyle1">向上滚动</label>
		<INPUT onClick="SetType(1)" <% IF TypeStyle = 1 Then  Response.Write "Checked" %>  type="radio"  value="1" name="TypeStyle"  id="TypeStyle2"><label for="TypeStyle2">横向列表</label>
		<INPUT onClick="SetType(2)" <% IF TypeStyle = 2 Then  Response.Write "Checked" %>  type="radio" value="2" name="TypeStyle"  id="TypeStyle3"><label for="TypeStyle3">下拉列表</label>
         <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_TypeStyle')"  id="Label_TypeStyle">帮助</span>
        
        </td>
    </tr>
    <tr>
      <td >Logo宽度
      <input name="LogoWidth"  type="text"  class="Ainput" id="LogoWidth" value="<%= LogoWidth %>">
      <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_LogoWidth')"  id="Label_LogoWidth">帮助</span>
      </td>
      <td >Logo高度
        <input name="LogoHeight"  type="text"  class="Ainput" id="LogoHeight" value="<%= LogoHeight %>"> 
        <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_LogoHeight')"  id="Label_LogoHeight">帮助</span>     </td>
    </tr>
    <tr>
      <td >标题字数
      <input name="TitleLen"  type="text"  class="Ainput" id="TitleLen" value="<%= TitleLen %>">
      <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_TitleLen')"  id="Label_TitleLen">帮助</span>
      
      </td>
      <td >显示列数
      <input name="ColNumber"  type="text"  class="Ainput" id="ColNumber" value="<%= ColNumber %>">
      <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_ColNumber')"  id="Label_ColNumber">帮助</span>
      </td>
    </tr>
    <tr>
      <td colspan="2" align="center"  >
       <input name="SubmitBtn" class="ACT_btn" type="button"  onClick="InsertScriptFun()"  id="SubmitBtn"  value=" 确 定 ">    
      &nbsp;&nbsp; <input type="reset" class="ACT_btn" name="Submit" value="  重置  "></td>
    </tr>
    </table>
  
</form>
<script language="javascript" >
		function SetActF(Val)
		{
		 if(Val==1)	
			{
			 ActFs.style.display="none";
 			  document.myform.TypeStyle1.disabled=false;
			  document.myform.TypeStyle2.disabled=false;
			  document.myform.TypeStyle3.disabled=false;
			  document.myform.LogoWidth.disabled=false;
			  document.myform.LogoHeight.disabled=false;
			  document.myform.TitleLen.disabled=false;
			  document.myform.ColNumber.disabled=false;
			}
		 if(Val==2)	
			{
			 ActFs.style.display="";
 			  document.myform.TypeStyle1.disabled=true;
			  document.myform.TypeStyle2.disabled=true;
			  document.myform.TypeStyle3.disabled=true;
			  document.myform.LogoWidth.disabled=true;
			  document.myform.LogoHeight.disabled=true;
			  document.myform.TitleLen.disabled=true;
			  document.myform.ColNumber.disabled=true;
 			}
		}

 function  
  SetLinkContent(oTextarea,strText){   
  oTextarea.focus();   
  document.selection.createRange().text+=strText;   
  oTextarea.blur();   
  }   

	
function OpenWindow(Url,Width,Height,WindowObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;');
	return ReturnStr;
}	

	function SetLogo(Num)
		{
		if (Num==0)
		{
		 document.myform.LogoWidth.disabled=true;
		 document.myform.LogoHeight.disabled=true;
		}
		else
		{
		 document.myform.LogoWidth.disabled=false;
		 document.myform.LogoHeight.disabled=false;
		}
		}
function SetType(Num)
		{
		 if (Num==0||Num==2)
		  {
		  document.myform.ColNumber.disabled=true;
		  }
		 else
		  {
		  document.myform.ColNumber.disabled=false;
		  }
		}		function OpenWindowAndSetValue(Url,Width,Height,WindowObj,SetObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:yes;help:no;scroll:yes;');
	if (ReturnStr!='') SetObj.value=ReturnStr;
	return ReturnStr;
}	
	function InsertScriptFun(Obj)
		{   if (document.myform.LabelName.value=='')
			 {
			  alert('请输入标签名称');
			  document.myform.LabelName.focus(); 
			  return false
			  }
			var LinkType,TypeStyle;
			var ClassID=document.myform.ClassID.value;
			var LogoWidth=document.myform.LogoWidth.value;
			var LogoHeight=document.myform.LogoHeight.value;
			var ListNumber=document.myform.ListNumber.value;
			var TitleLen=document.myform.TitleLen.value;
			var ColNumber=document.myform.ColNumber.value;
			var ActF=document.myform.ActF.value;
			var LinkSort=document.myform.LinkSort.value;
			var LinkContent=document.myform.LinkContent.value;
 			for (var i=0;i<document.myform.LinkType.length;i++){
			 var TCJ = document.myform.LinkType[i];
			if (TCJ.checked==true)	   
				LinkType = TCJ.value
			}
			for (var i=0;i<document.myform.TypeStyle.length;i++){
			 var TCJ = document.myform.TypeStyle[i];
			if (TCJ.checked==true)	   
				TypeStyle = TCJ.value
			}
			if  (document.myform.LinkSort.value=='') LinkSort="ID Desc";
			document.myform.LabelContent.value=	'{$GetLinkList('+ClassID+'§'+LinkType+'§'+TypeStyle+'§'+LogoWidth+'§'+LogoHeight+'§'+ListNumber+'§'+TitleLen+'§'+ColNumber+'§'+ActF+'§'+LinkSort+'§'+LinkContent+')}';
			document.myform.SubmitBtn.value="正在提交数据,请稍等...";
			document.myform.SubmitBtn.disabled=true;
			document.myform.Submit.disabled=true;	
			document.myform.submit();
		}
</script>
<script language="javascript">SetActF(<%= ActF %>);</script>
</body>
</html>
<%		response.Write "<script>"
		response.Write "SetLogo(" & LinkType & ");"
		response.Write "SetType(" & TypeStyle & ");"
		response.Write "</script>"
 %>