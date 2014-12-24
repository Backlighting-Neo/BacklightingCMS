<!--#include file="../../ACT.Function.asp"-->
<!--#include file="../ACT.F.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ACT_标签管理</title>
<link href="../../Images/style.css" rel="stylesheet" type="text/css">
</head>
<SCRIPT src="../../../ACT_inc/dtreeFunction.js"></SCRIPT>
<LINK href="../../../ACT_inc/dtree.css" type=text/css rel=StyleSheet>
<SCRIPT src="../../../ACT_inc/dtree.js" type=text/javascript></SCRIPT>
<script language="JavaScript" type="text/javascript" src="../../../ACT_inc/js/lhgcore/lhgcore.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../../ACT_inc/js/lhgcore/lhgdialog.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../../ACT_inc/js/lhgcore/Main.js"></script>
<body>
<%
Dim Action,ID,LabelRS,LabelName,Descript,LabelContent,LabelFlag,LabelContentArr,ClassID,ClassName,Rs,SubClass,pages
Dim ArticleSort,TitleLen,RowHeight,ColNumber,TitleCss,DateCss,DateForm,DateAlign,NavType,TypeWordPic,IntroNumber
Dim sysdir,TypeClassName,OpenType,Division,TypeNew,TypeHot,NavPic,NavWord,DiyContent
dim pagestyle,pagenumber,iftrue
Dim ActF,Str,ACTIF,ModeID
If Not ACTCMS.ACTCMS_QXYZ(0,"bqxt","") Then   Call Actcms.Alert("对不起，你没有操作权限！","")
sysdir=actcms.ActSys
Action = Request.QueryString("Action")
ID =  ChkNumeric(Request.QueryString("ID"))
IF Action = "Add" Then
	ClassID = 1
	TitleLen =30
	ColNumber = 1
	ActF=1
	RowHeight = 22
	PageStyle = 1
	PageNumber = 20
	pages = "新建分页文章列表标签"
	ModeID=1
	IntroNumber=50
	ArticleSort="ID Desc"
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
		LabelContent = Replace(Replace(LabelContent, "{$GetLastArticleList(", ""), ")}", "")
		LabelContentArr = Split(LabelContent, "§")

		ActF=LabelContentArr(0)
		ClassID=LabelContentArr(1)
		PageStyle=LabelContentArr(2)
		ArticleSort = LabelContentArr(3)'排序方法
		OpenType=LabelContentArr(4)
		PageNumber=LabelContentArr(5)
		RowHeight=LabelContentArr(6)
		TitleLen = LabelContentArr(7)
		ColNumber = LabelContentArr(8)
		TypeClassName = LabelContentArr(9)
		TypeNew = LabelContentArr(10)	
		NavType= LabelContentArr(11)
		Division = LabelContentArr(13)'分隔图片
		DateForm = LabelContentArr(14)'日期格式
		DateAlign = LabelContentArr(15)'日期对齐
		TitleCss = LabelContentArr(16)'标题样式
		DateCss  = LabelContentArr(17)'日期样式


		ACTIF=LabelContentArr(18)
		DiyContent=LabelContentArr(19)
		ModeID=LabelContentArr(20)
		SubClass=LabelContentArr(21)
		IntroNumber=LabelContentArr(22)


		IF NavType = 0 Then 
			NavWord = LabelContentArr(12)
			Navpic = ""
		Else
			NavWord = ""
			Navpic = LabelContentArr(12)
		End IF
		pages = "修改分页文章列表标签"
		NavWord=server.HTMLEncode(NavWord)
End IF
 With Response 
%>
<form id="myform" name="myform" method="post" action="AddLabelSave.asp">
 <input type="hidden" name="LabelContent"> 
  <input type="hidden" name="LabelType" value="1">
  <input type="hidden" name="Action" value="<%= Action %>"> 
  <input type="hidden" name="ID" value="<%= ID %>"> 
 <input type="hidden" name="FileUrl" value="GetLastArticleList.asp">  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
    <tr>
      <td colspan="2" align="center" class="bg_tr"><%= pages %>&nbsp;</td>
    </tr>
    <tr>
      <td width="50%" class="td_bg">标签名称
      <input name="LabelName"  type="text" class="Ainput" id="LabelName" value="<%= LabelName %>"></td>
      <td width="50%" class="td_bg">&nbsp;</td>
    </tr>
    <tr>
      <td class="td_bg">标签目录      
        <select name="LabelFlag" id="select">
          <option value="0">系统默认</option>
			 <%=AF.ACT_LabelFolder(CInt(LabelFlag))%>
        </select>&nbsp;&nbsp;<a href="ACT.LabelFolder.asp"><font color=red><b>新建存放目录</b></font></a>
		&nbsp;<font color=green>标签存放目录,方便管理标签</font>
			
		</td>

	 <td width="50%" height="24"  class="td_bg"> 
	 所属模型
	 <select name="ModeID" id="ModeID">
          <option value="0" style="color:green">模型通用</option>
          <%=AF.ACT_L_Mode(CInt(ModeID))%>
        </select>
		<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_ModeID')"  id="Label_ModeID">帮助</span>
        </td>
    </tr>

	 <tr>
      <td width="50%" >
	 	 输出模式
		 <select  style='width:40%' name="ActF" id="ActF" onChange="SetActF(this.options[this.selectedIndex].value);"> 
	 <option value="1" <% IF ActF = 1 Then Response.Write("selected") %>>普通模式</option>
  <option value="2" <% IF ActF = 2 Then Response.Write("selected") %>>代码模式</option>
  </select>   
			<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_Code')"  id="Label_Code">帮助</span> </td>
      <td width="50%" ><%=actcms.ReturnPageStyle(PageStyle) %>
	  </td>
    </tr>
  <tr id=ActFs ><td  colspan="2" >

<font color=red>内置标签</font> 
<a href="#" onClick='SetDiyContent(DiyContent,"#ID")'>文章ID</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#Link")'>文章链接</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#Title")'>文章标题</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#CTitle")'>文章标题(过滤HTML)</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#KeyWords")'>关键字</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#Thumb")'>缩略图</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#PicUrl")'>图片地址</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#Intro")'>文章导读</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#ClassName")'>栏目名称</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#ClassLink")'>栏目链接</a>&nbsp;<br />
<a href="#" onClick='SetDiyContent(DiyContent,"#Time")'>时间</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#Hits")'>点击数</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#CopyFrom")'>文章来源</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#Author")'>文章作者</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#AutoID")'>自增长ID</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#ModID")'>间歇ID</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#Path")'>系统路径</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#ClassSeo")'>栏目SEO标题</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#ClassPicUrl")'>栏目缩略图</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#ClassPicFile")'>栏目缩略图地址</a>&nbsp;
<a style="cursor:pointer;" onClick="javascript:upload('uploadzd');" id="uploadzd"><font color="#FF0000">[自定义字段]</font></a>
 
<font color=red>导读字数</font>
<input name="IntroNumber" id="IntroNumber" type="text" class="Ainput" value="<%=IntroNumber%>" size="10">

<br />
<textarea onFocus="this.className='colorfocus';" onBlur="this.className='colorblur';"  name="DiyContent"  id="DiyContent"   cols="90" rows="10"><%=Server.HTMLEncode(DiyContent)%></textarea>


<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_DiyContent')"  id="Label_DiyContent">帮助</span> 
  </td>
	</tr>

   <tr>
      <td >
	  所属栏目
	   <input name="ClassID" type="text" class="Ainput" id="ClassID" value="<%= ClassID %>" readonly disabled=true>
	    <select name="select1" onChange="SelectClass();">
    <option value="0" <% IF ClassID = "0" Then  Response.Write "selected" %>>不指定栏目</option>
    <option value="1"  style="color:red" <% IF ClassID = "1" Then  Response.Write "selected" %>>当前栏目通用</option>
	<option value="2" <% IF ClassID <> "0" And ClassID <> "1"  Then  Response.Write "selected" %>>指定栏目</option>
	</select><a href="#" onClick="SelectClass()">快速打开</a>
      <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_ClassID')" id="Label_ClassID">帮助</span> </td>
      <td ><input id="SubClass22" type="checkbox" value="true" name="SubClass" <%If InStr(ClassID,",") > "1"  Then  .write "disabled=true  "%> <%If CBool(SubClass)=true  Then  .write "Checked  "%>>
	  <label for="SubClass22">允许包含子栏目</label>
          <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_SubClass')"  id="Label_SubClass">帮助</span> </td>
    </tr>



    <tr>
      <td class="td_bg">
	  
		排序方法  
	<input name="ArticleSort" id="ArticleSort" type="text" class="Ainput" value="<%=ArticleSort%>" size="20">

	  
	<select   name="ArticleSorts"  onchange="document.myform.ArticleSort.value=this.value">
		<option value='ID Desc' <%If ArticleSort="ID Desc" Then .write "selected":iftrue=true%>>文章ID(降序)</option>
		<option value='ID Asc' <%If ArticleSort="ID Asc" Then .write "selected":iftrue=true%>>文章ID(升序)</option>
		<option value='UpdateTime Asc' <%If ArticleSort="UpdateTime Asc" Then .write "selected":iftrue=true%>>更新时间(升序)</option>
		<option value='UpdateTime Desc' <%If ArticleSort="UpdateTime Desc" Then .write "selected":iftrue=true%>>更新时间(降序)</option>
		<option value='Hits Asc' <%If ArticleSort="Hits Asc" Then .write "selected":iftrue=true%>>点击数(升序)</option>
		<option value='Hits Desc' <%If ArticleSort="Hits Desc" Then .write "selected":iftrue=true%>>点击数(降序)</option>
 		<option value='commentscount Asc' <%If ArticleSort="commentscount Asc" Then .write "selected":iftrue=true%>>评论数(升序)</option>
		<option value='commentscount Desc' <%If ArticleSort="commentscount Desc" Then .write "selected":iftrue=true%>>评论数(降序)</option>
		<option value='Digg Desc' <%If ArticleSort="Digg Desc" Then .write "selected":iftrue=true%>>digg支持(降序)</option>
		<option value='down Desc' <%If ArticleSort="down Desc" Then .write "selected":iftrue=true%>>digg反对(降序)</option>	
		<option value='' style="color:red"  <%If iftrue=false Then .write "selected"%> >自定义</option>	
		</select> 
 		  <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_ArticleSort')"  id="Label_ArticleSort">帮助</span>
 		</td>
      <td class="td_bg">

	  
链接目标  
	<input name="OpenType" id="OpenType" type="text" class="Ainput" value="<%=OpenType%>" size="8">

	  <%iftrue=false%>
	<select   name="OpenTypes"  onchange="document.myform.OpenType.value=this.value">
          <option value="_blank"   style="color:green"  <%If OpenType="_blank" Then .write "selected":iftrue=true%>>新窗口打开</option>
          <option value="_parent" <%If OpenType="_parent" Then .write "selected":iftrue=true%>>父窗口打开</option>
          <option value="_self" <%If OpenType="_self" Then .write "selected":iftrue=true%>>本窗口打开</option>
          <option value="_top" <%If OpenType="_top" Then .write "selected":iftrue=true%>>主窗口打开</option>
		<option value='' style="color:red"  <%If iftrue=false Then .write "selected"%>>自定义</option>	
		</select>	  
	 <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_OpenType')"  id="Label_OpenType">帮助</span>  
	</td>
    </tr>
    <tr>
      <td class="td_bg">
	  每页数量
	   	  <input name="PageNumber" type="text" class="Ainput" id="PageNumber" value="<%= PageNumber %>" onKeyUp="value=value.replace(/[^\d]/g,'') "onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))">
	<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_PageNumber')"  id="Label_PageNumber">帮助</span>   
	  </td>
      <td class="td_bg">文章行距
      <input name="RowHeight"  type="text" class="Ainput" id="RowHeight2"    style="width:70%;" value="<%= RowHeight %>">
      <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_RowHeight')"  id="Label_RowHeight">帮助</span>  
      </td>
    </tr>
    <tr>
      <td class="td_bg">标题字数
      <input name="TitleLen"   type="text" class="Ainput"    style="width:70%;" value="<%= TitleLen %>">
     <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_TitleLen')"  id="Label_TitleLen">帮助</span>   
      </td>
      <td class="td_bg">排列列数
        <input type="text" class="Ainput"    size="4" value="<%= ColNumber %>" name="ColNumber">
        <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_ColNumber')"  id="Label_ColNumber">帮助</span>  
        </td>
    </tr>
    <tr>
      <td colspan="1" >附加显示
        <input name="TypeClassName" type="checkbox" id="TypeClassName"  value="<%= TypeClassName %>"  <% IF TypeClassName = "true" Then Response.Write("Checked") %>><label for="TypeClassName">显示栏目</label>&nbsp;&nbsp;&nbsp;     
		<% 	
		  .Write "&nbsp;&nbsp;&nbsp;"
		 If  cbool(TypeNew) = True Then
		  .Write ("<input id=""TypeNew1"" type=""checkbox"" value=""true"" name=""TypeNew"" checked><label for=""TypeNew1"">最新文章标志</label>")
		 Else
		  .Write ("<input id=""TypeNew2"" type=""checkbox"" value=""true"" name=""TypeNew""><label for=""TypeNew2"">最新文章标志</label>")
		 End If

 %>      <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_TypeClassName')"  id="Label_TypeClassName">帮助</span>        </td>
 <td>SQL判断条件<input type="text" class="Ainput"  style="width:60%;"  size="4" value="<%= ACTIF %>" name="ACTIF">
 <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_ACTIF')"  id="Label_ACTIF">帮助</span> 
 </td>
    </tr>
    <tr >
      <td class="td_bg">导航类型
        
      <label for="NavType1">  <input id="NavType1"  <% IF NavType = 0 Then Response.Write("Checked") %> type="radio" name="NavType" value="0" onClick="SetNavStatus(1);">
        文字导航</label>
       <label for="NavType2">  <input id="NavType2" <% IF NavType = 1 Then Response.Write("Checked") %>   type="radio" name="NavType" value="1" onClick="SetNavStatus(2);">
        图片导航</label>	
        
         <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_NavType')"  id="Label_NavType">帮助</span> 	</td>
    <td width="50%" height="24" class="td_bg"  id=SetNavStatus1 
	<% if NavType=1 then %>
	style="DISPLAY: none" 
	<% end if %>>
 <input name="NavWord" type="text" class="Ainput"  id="NavWord" style="width:70%;" value="<%= NavWord %>"> 
 支持HTML语法<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_NavWord')"  id="Label_NavWord">帮助</span></td>
   <td width="50%" height="24" class="td_bg"  id=SetNavStatus2 
	<% if NavType=0 then %>
	style="DISPLAY: none"
	SetNavStatus1.style.display=''
	<% end if %> >
<input name="NavPic" type="text" class="Ainput"  id="NavPic" style="width:250;" value="<%= NavPic %>" readonly>
<input class="ACT_btn" type="button"  onClick="OpenWindowAndSetValue('../print/SelectPic.asp?CurrPath=<%=sysdir%>',500,320,window,document.myform.NavPic);" value="选择图片...">&nbsp;<span style="cursor:hand;color:green;" onClick="javascript:document.myform.NavPic.value='';" onMouseOver="this.style.color='red'" onMouseOut="this.style.color='green'">清除</span>
<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_NavPic')"  id="Label_NavPic">帮助</span></td>  
    </tr>
    <tr>
      <td colspan="2" class="td_bg">分隔图片
      <input name="Division"  type="text" class="Ainput" id="Division" style="width:61%;" value="<%= Division %>" readonly>
	  <input class="ACT_btn" type="button" id="Division1"  onClick="OpenWindowAndSetValue('../print/SelectPic.asp?CurrPath=<%=sysdir%>',500,320,window,document.myform.Division);" name="Submit3" value="选择图片...">
	 &nbsp;<span style="cursor:hand;color:green;" onClick="javascript:document.myform.Division.value='';" onMouseOver="this.style.color='red'" onMouseOut="this.style.color='green'">清除</span>
	 <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_Division')"  id="Label_Division">帮助</span> </td>
    </tr>
    <tr>
      <td class="td_bg">日期格式
        <select  style="width:70%;" name="DateForm" id="select2">
        <% 
		.Write AF.ACT_DateStr(DateForm)
			 %>
        </select>  <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_DateForm')"  id="Label_DateForm">帮助</span>    </td>
      <td class="td_bg">日期对齐

<% 		.Write "  <select class=""textbox"" name=""DateAlign"" id=""select3"" style=""width:70%;"">"
							
					If ID = "" Or CStr(DateAlign) = "left" Then
					 Str = " selected"
					Else
					 Str = ""
					End If
					 .Write ("<option value=""left""" & Str & ">左对齐</option>")
					If CStr(DateAlign) = "center" Then
					 Str = " selected"
					Else
					 Str = ""
					End If
					 .Write ("<option value=""center""" & Str & ">居中对齐</option>")
					If CStr(DateAlign) = "right" Then
					 Str = " selected"
					Else
					 Str = ""
					End If
					 .Write ("<option value=""right""" & Str & ">右对齐</option>")
					 
				.Write "                  </select>"
		End With %> <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_DateAlign')"  id="Label_DateAlign">帮助</span>   </td>
    </tr>
    <tr>
      <td class="td_bg">标题样式
      <input name="TitleCss" type="text" class="Ainput"  id="TitleCss" style="width:70%;" value="<%= TitleCss %>">
      <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_TitleCss')"  id="Label_TitleCss">帮助</span> 
      
      </td>
      <td class="td_bg">日期样式
      <input name="DateCss"  type="text" class="Ainput" id="DateCss" style="width:70%;" value="<%= DateCss %>">
      <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_DateCss')"  id="Label_DateCss">帮助</span> 
      </td>
    </tr>
    <tr>
      <td colspan="2" align="center" class="td_bg">
	   <input name="SubmitBtn" class="ACT_btn" type="button"  onClick="InsertScriptFun()"  id="SubmitBtn"  value=" 确 定 ">	
	    &nbsp;&nbsp; <input type="reset" class="ACT_btn" name="Submit" value="  重置  ">  </td>
    </tr>
    </table>
  
</form>
<script language="javascript" >

 function upload(iname) 
{
   var cid=document.myform.ModeID.value
   J('#'+iname).dialog({ id:'actcmssc'+iname+'s' ,title:'自定义字段', loadingText:'上传加载中...', page: '<%=actcms.actsys&actcms.adminurl%>/include/Label/diyField.asp?ModeID='+cid+ '&instrname='+iname+ "&" + Math.random(),fixed:true, left:300, top:300 ,  width:720, height:240 });
}

function SelectClass()
{
 if(document.myform.select1.value==0 )	
	{
	document.all.ClassID.value=0
	document.myform.SubClass.disabled=false;
	}

 if(document.myform.select1.value==1 )	
	{
	document.all.ClassID.value=1
	document.myform.SubClass.disabled=false;
	}

 if(document.myform.select1.value==2 )	
	{
	var cid=document.myform.ModeID.value
	if (cid==0)
	{ 
	document.all.ClassID.value=0;
	alert("模型通用不能选择栏目");
	return false;
	}
	var result = Selector(2, document.all.ClassID.value,cid);
	if(!result) return false;
	var val = "";
	for(var i=0; i<result.length; i++)
	{
		if(val == "")
		{
			val += result[i].id;
		}else{
			val += "," + result[i].id;
		}
	}
	document.all.ClassID.value = val;
		var str=val;
		var i = str.indexOf(",");
		if (i==-1)
		{
 		document.myform.SubClass.disabled=false;
		}
		else
		{
		document.myform.SubClass.disabled=true;

		}
	
	}
	}
	
	function SetNavStatus(n){
			if (n==1){
			SetNavStatus1.style.display='';
			SetNavStatus2.style.display='none';	
			}
		  else{
			SetNavStatus1.style.display='none';
			SetNavStatus2.style.display='';	
		}
}

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
		function SetActF(Val)
		{
		 if(Val==1)	
			{
			 ActFs.style.display="none";
			  document.myform.OpenType.disabled=false;
			  document.myform.OpenTypes.disabled=false;
			  document.myform.NavType1.disabled=false;
 			  document.myform.NavType2.disabled=false;
 			  document.myform.NavWord.disabled=false;
			  document.myform.NavPic.disabled=false;
			  document.myform.Division1.disabled=false;
 			  document.myform.TitleCss.disabled=false;
			  document.myform.DateCss.disabled=false;
			  document.myform.ColNumber.disabled=false;
			  document.myform.TypeClassName.disabled=false;
			  document.myform.TypeNew.disabled=false;
			  
			  document.myform.DateAlign.disabled=false;
			}
		 if(Val==2)	
			{
			 ActFs.style.display="";
			  document.myform.OpenType.disabled=true;
			  document.myform.OpenTypes.disabled=true;
			  document.myform.NavType1.disabled=true;
			  document.myform.NavType2.disabled=true;
 			  document.myform.NavWord.disabled=true;
			  document.myform.NavPic.disabled=true;
			  document.myform.Division1.disabled=true;
 			  document.myform.TitleCss.disabled=true;
			  document.myform.DateCss.disabled=true;
			  document.myform.ColNumber.disabled=true;
			  document.myform.TypeClassName.disabled=true;
			  document.myform.TypeNew.disabled=true;
			  
			  document.myform.DateAlign.disabled=true;
			}
		}
 
		function InsertScriptFun()
		{   if (document.myform.LabelName.value=='')
			 {
			  alert('请输入标签名称');
			  document.myform.LabelName.focus(); 
			  return false
			  }
			var SubClass=false,NavType=1;
			var ClassID=document.myform.ClassID.value;
			var TypeClassName,TypeWordPic,TypeNew,TypeHot;
			var PageNumber=document.myform.PageNumber.value;
			var PageStyle=document.myform.PageStyle.value;
			var RowHeight=document.myform.RowHeight.value;
			var TitleLen=document.myform.TitleLen.value;
			var ColNumber=document.myform.ColNumber.value;
			var Nav,NavType=document.myform.NavType.value;
			var DateForm=document.myform.DateForm.value;
		
			var OpenType=document.myform.OpenType.value;
			var ArticleSort=document.myform.ArticleSort.value;
			var Division=document.myform.Division.value;
			var DateAlign=document.myform.DateAlign.value;
			var TitleCss=document.myform.TitleCss.value;
			var DateCss=document.myform.DateCss.value;
			
			
			var ActF=document.myform.ActF.value;

			var DiyContent=document.myform.DiyContent.value;
			var ACTIF=document.myform.ACTIF.value;
 			var ModeID=document.myform.ModeID.value;
			var IntroNumber=document.myform.IntroNumber.value;

			if (document.myform.SubClass.checked) SubClass=true;
			if (document.myform.TypeClassName.checked)
				TypeClassName = true
			else
			    TypeClassName =false;

			if (document.myform.TypeNew.checked)
			   TypeNew= true
			else
			   TypeNew=false;

			if (RowHeight=='') RowHeight=20
			if  (TitleLen=='') TitleLen=30;
			if  (ColNumber=='') ColNumber=1;
			if  (IntroNumber=='') IntroNumber=50;
			if  (PageNumber=='') PageNumber=10;

			for (var i=0;i<document.myform.NavType.length;i++){
			 var TCJ = document.myform.NavType[i];
			if (TCJ.checked==true)	   
				NavType = TCJ.value
			if  (NavType==0) Nav=''+document.myform.NavWord.value+''
			 else  Nav=''+document.myform.NavPic.value+'';
			}
			if  (document.myform.ArticleSort.value=='') ArticleSort="ID Desc";
			document.myform.LabelContent.value='{$GetLastArticleList('+ActF+'§'+ClassID+'§'+PageStyle+'§'+ArticleSort+'§'+OpenType+'§'+PageNumber+'§'+RowHeight+'§'+TitleLen+'§'+ColNumber+'§'+TypeClassName+'§'+TypeNew+'§'+NavType+'§'+Nav+'§'+Division+'§'+DateForm+'§'+DateAlign+'§'+TitleCss+'§'+DateCss+'§'+ACTIF+'§'+DiyContent+'§'+ModeID+'§'+SubClass+'§'+IntroNumber+')}';
			document.myform.SubmitBtn.value="正在提交数据,请稍等...";
			document.myform.SubmitBtn.disabled=true;
			document.myform.Submit.disabled=true;	
			document.myform.submit();
		}
	
</script><script language="javascript">SetActF(<%= ActF %>);</script>
</body>
</html>
