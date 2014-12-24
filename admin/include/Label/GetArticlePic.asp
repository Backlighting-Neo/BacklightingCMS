<!--#include file="../../ACT.Function.asp"-->
<!--#include file="../ACT.F.asp"-->

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ACT图片标签管理</title>
<link href="../../Images/style.css" rel="stylesheet" type="text/css">
<SCRIPT src="../../../ACT_inc/dtreeFunction.js"></SCRIPT>
<LINK href="../../../ACT_inc/dtree.css" type=text/css rel=StyleSheet>
<SCRIPT src="../../../ACT_inc/dtree.js" type=text/javascript></SCRIPT>
<script language="JavaScript" type="text/javascript" src="../../../ACT_inc/js/lhgcore/lhgcore.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../../ACT_inc/js/lhgcore/lhgdialog.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../../ACT_inc/js/lhgcore/Main.js"></script>
</head>
<body>
<%
Dim Action,ID,LabelRS,LabelName,Descript,LabelContent,LabelFlag,LabelContentArr,ClassID,ClassName,Rs,pages
Dim ArticleSort,ListNumber,TitleLen,ColNumber,TitleCss,ContentLen,PicHeight,PicWidth,SubClass
Dim OpenType,PicStyle,TypeTitle,piccss,ModeID,iftrue,DiyContent
Dim ActF,Str,ACTIF,ATT
If Not ACTCMS.ACTCMS_QXYZ(0,"bqxt","") Then   Call Actcms.Alert("对不起，你没有操作权限！","")
Action = Request.QueryString("Action")
ID =  ChkNumeric(Request.QueryString("ID"))
IF Action = "Add" Then
	ClassID = 0
	ListNumber = 10
	TitleLen =30
	ATT = 0
	ActF=1
	ColNumber = 1
	PicWidth = 130
	PicHeight = 90
	PicStyle = 1
	TitleLen = 30
	LabelFlag = 0
	ContentLen = 50
	TypeTitle = True
	ModeID =  1
	ClassName = "不指定栏目"
	pages = "新建图片文章列表标签"
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
		LabelContent = Replace(Replace(LabelContent, "{$GetArticlePic(", ""), ")}", "")
		LabelContentArr = Split(LabelContent, "§")
		ClassID=LabelContentArr(0)
		ActF = LabelContentArr(1)
		ATT = LabelContentArr(2)
		ArticleSort = LabelContentArr(3)
		OpenType = LabelContentArr(4)
		ListNumber = LabelContentArr(5)
		ColNumber = LabelContentArr(6)
		TitleLen = LabelContentArr(7)
		Titlecss= LabelContentArr(8)
		piccss = LabelContentArr(9)
		PicWidth = LabelContentArr(10)
		PicHeight = LabelContentArr(11)
		ContentLen = LabelContentArr(12)
		PicStyle = LabelContentArr(13)
		TypeTitle = LabelContentArr(14)
		ACTIF= LabelContentArr(15)
		DiyContent=LabelContentArr(16)
		ModeID=LabelContentArr(17)
		SubClass=LabelContentArr(18)
		pages = "修改图片文章列表标签"
End IF
 With Response 

%>
<form id="myform" name="myform" method="post" action="AddLabelSave.asp">
 <input type="hidden" name="LabelContent"> 
  <input type="hidden" name="LabelType" value="1">
  <input type="hidden" name="Action" value="<%= Action %>"> 
  <input type="hidden" name="ID" value="<%= ID %>"> 
 <input type="hidden" name="FileUrl" value="GetArticlePic.asp">  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
    <tr>
      <td colspan="2" align="center" class="bg_tr"><%= pages %>&nbsp;</td>
    </tr>
    <tr>
      <td width="50%">标签名称
      <input name="LabelName" type="text" class="Ainput" id="LabelName" value="<%= LabelName %>"></td>
      <td width="50%"><font color="red">* 调用格式"{ACTCMS_标签名称}"</font></td>
    </tr>
    <tr>
      <td>标签目录      
        <select name="LabelFlag" id="select">
         <option value="0">系统默认</option>
			 <%=AF.ACT_LabelFolder(CInt(LabelFlag))%>
        </select>&nbsp;&nbsp;<a href="ACT.LabelFolder.asp"><font color=red><b>新建存放目录</b></font></a>
		&nbsp;<font color=green>标签存放目录,方便管理标签</font></td>
	   <td width="50%" height="24" >
所属模型
	
	 <select name="ModeID" id="ModeID">
          <option value="0" style="color:green">模型通用</option>
          <%=AF.ACT_L_Mode(CInt(ModeID))%>
        </select>	
		<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_ModeID')"  id="Label_ModeID">帮助</span>  </td>
    </tr>
		 <tr>
      <td width="50%">
	 输出模式
		 <select  style='width:40%' name="ActF" id="ActF" onChange="SetActF(this.options[this.selectedIndex].value);"> 
	 <option value="1" <% IF ActF = 1 Then Response.Write("selected") %>>普通模式</option>
  <option value="2" <% IF ActF = 2 Then Response.Write("selected") %>>代码模式</option>
  </select><span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_Code')"  id="Label_Code">帮助</span> 
		 </td>
      <td width="50%">
文章属性
        <select name="ATT" id="ATT">
          <option value="0" style="color:green">普通文章</option>
          <%=ACTCMS.ACT_ATT(CInt(ATT))%>
        </select><span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_Att')"  id="Label_Att">帮助</span> 
	  </td>
    </tr>
	
  <tr id=ActFs><td  colspan="2" >

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
<a href="#" onClick='SetDiyContent(DiyContent,"#New")'>New图标</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#ClassSeo")'>栏目SEO标题</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#ClassPicUrl")'>栏目缩略图</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#ClassPicFile")'>栏目缩略图地址</a>&nbsp;
<a style="cursor:pointer;" onClick="javascript:upload('uploadzd');" id="uploadzd"><font color="#FF0000">[自定义字段]</font></a>
<br />
<textarea onFocus="this.className='colorfocus';" onBlur="this.className='colorblur';"  name="DiyContent"  id="DiyContent"   cols="80" rows="10"><%=Server.HTMLEncode(DiyContent)%></textarea>
  <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_DiyContent')"  id="Label_DiyContent">帮助</span> 
  </td>
	</tr>
	
	
	<tr>
          <td>
	  所属栏目
	   <input name="ClassID" type="text" class="Ainput" id="ClassID" value="<%= ClassID %>" readonly disabled=true>
	    <select name="select1" onChange="SelectClass();">
    <option value="0" <% IF ClassID = "0" Then  Response.Write "selected" %>>不指定栏目</option>
    <option value="1"  style="color:red" <% IF ClassID = "1" Then  Response.Write "selected" %>>当前栏目通用</option>
	<option value="2" <% IF ClassID <> "0" And ClassID <> "1"  Then  Response.Write "selected" %>>指定栏目</option>
	</select><a href="#" onClick="SelectClass()">快速打开</a>
      <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_ClassID')"  id="Label_ClassID">帮助</span>  </td>
      <td><input id="SubClass22" type="checkbox" value="true" name="SubClass" <%If InStr(ClassID,",") > "1"  Then  .write "disabled=true  "%> <%If CBool(SubClass)=true  Then  .write "Checked  "%>>
	  <label for="SubClass22">允许包含子栏目</label> <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_SubClass')"  id="Label_SubClass">帮助</span>  </td>
    </tr>
    <tr>
      <td>
	  
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
      <td>

	  
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
      <td>文章数量
      <input name="ListNumber" type="text" class="Ainput" id="ListNumber2"    style="width:70%;"  value="<%= ListNumber %>">
       <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_ListNumber')"  id="Label_ListNumber">帮助</span>
      </td>
      <td>每行数量
        <input type="text" class="Ainput"   size="4" value="<%= ColNumber %>" name="ColNumber">
        <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_ColNumber')"  id="Label_ColNumber">帮助</span>
        </td>
    </tr>
    <tr>
      <td>标题字数
      <input name="TitleLen"  type="text" class="Ainput"    style="width:70%;" value="<%= TitleLen %>">
      <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_TitleLen')"  id="Label_TitleLen">帮助</span>
      </td>
      <td>标题样式
      <input name="TitleCss" type="text" class="Ainput" id="TitleCss" style="width:70%;" value="<%= TitleCss %>">
      <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_TitleCss')"  id="Label_TitleCss">帮助</span></td>
    </tr>
    <tr>
      <td>显示标题
      <label for="TypeTitle1"> <INPUT <% IF Cbool(TypeTitle) = True Then  Response.Write "Checked" %> type="radio" id="TypeTitle1" value="true" name="TypeTitle">
        是</label>
        <label for="TypeTitle2">  <INPUT <% IF Cbool(TypeTitle) = false Then  Response.Write "Checked" %> type="radio" id="TypeTitle2" value="false" name="TypeTitle">
      否</label>
      <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_TypeTitle')"  id="Label_TypeTitle">帮助</span>
      </td>
      <td>图片样式
      <input name="piccss" type="text" class="Ainput" id="piccss" style="width:70%;" value="<%= piccss %>">
      
       <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_piccss')"  id="Label_piccss">帮助</span>
      
      </td>
    </tr>
    <tr>
      <td>图片设置 宽
        <input name="PicWidth" type="text" class="Ainput" id="PicWidth2" value="<%= PicWidth %>" size="6">
像素 高
<input name="PicHeight" type="text" class="Ainput" id="PicHeight2" value="<%=PicHeight  %>" size="6" >
像素<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_PicWidth')"  id="Label_PicWidth">帮助</span></td>
      <td>内容字数 
	   <input name="ContentLen" type="text" class="Ainput" id="ContentLenArea"    style="width:70%;" value="<%=ContentLen%>">     
       <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_ContentLen')"  id="Label_ContentLen">帮助</span> </td>
    </tr>
    <tr>
      <td>  选择样式 
	  <select name="PicStyle" onChange="SelectPicStyle(this.value)">
	  <%				 If PicStyle = "1" Then
							Response.Write ("<option value=""1"" selected>①:仅显示缩略图</option>")
						  Else
							Response.Write ("<option value=""1"">①:仅显示缩略图</option>")
						  End If
						  If PicStyle = "2" Then
							Response.Write ("<option value=""2"" selected>②:缩略图+名称:上下</option>")
						  Else
							Response.Write ("<option value=""2"">②:缩略图+名称:上下</option>")
						  End If
						  If PicStyle = "3" Then
							Response.Write ("<option value=""3"" selected>③:缩略图+(名称+简介:上下):左右</option>")
						  Else
							Response.Write ("<option value=""3"">③:缩略图+(名称+简介:上下):左右</option>")
						  End If
						  If PicStyle = "4" Then
							Response.Write ("<option value=""4"" selected>④:(名称+简介:上下)+缩略图:左右</option>")
						  Else
							Response.Write ("<option value=""4"">④:(名称+简介:上下)+缩略图:左右</option>")
						  End If
				  
 End With%></select><span id="PicStyles">&nbsp;</span>
 <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_PicStyle')"  id="Label_PicStyle">帮助</span></td>
      <td valign="top">SQL判断条件
      <input type="text" class="Ainput"  style="width:60%;"  size="4" value="<%= ACTIF %>" name="ACTIF">
       <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_ACTIF')"  id="Label_ACTIF">帮助</span>
      </td>
    </tr>
	    <tr>
      <td colspan="2" align="center" >
       <input name="SubmitBtn" class="ACT_btn" type="button"  onClick="InsertScriptFun()"  id="SubmitBtn"  value=" 确 定 ">       &nbsp;&nbsp;
       <input type="reset" class="ACT_btn" name="Submit" value="  重置  "></td>
    </tr>
    
    
    </table>
  
</form>
<script language="javascript" >
 function upload(iname) 
{
	
  
var cid=document.myform.ModeID.value;
(new J.dialog({ id:'actcmssc'+iname+'s' ,title:'自定义字段', loadingText:'加载中...', page: '/admin/include/Label/diyField.asp?ModeID='+cid+ '&instrname='+iname+ "&" + Math.random(),fixed:true, width:720, height:240 })).ShowDialog();

//J('#'+iname).dialog({ id:'actcmssc'+iname+'s' ,title:'自定义字段', loadingText:'上传加载中...', page: '/admin/include/Label/diyField.asp?ModeID='+cid+ '&instrname='+iname+ "&" + Math.random(),fixed:true, left:300, top:300 ,  width:720, height:240 });
  }








function SelectClass()
{

 if(document.myform.select1.value==0)	
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
		function SetActF(Val)
		{
		 if(Val==1)	
			{
			 ActFs.style.display="none";
			  document.myform.OpenType.disabled=false;
			  document.myform.OpenTypes.disabled=false;
 			  document.myform.TitleCss.disabled=false;
 			  document.myform.ColNumber.disabled=false;
 			  document.myform.piccss.disabled=false;
 			  document.myform.TypeTitle1.disabled=false;
 			  document.myform.TypeTitle2.disabled=false;
 			  document.myform.PicWidth2.disabled=false;
 			  document.myform.PicHeight2.disabled=false;
 			  document.myform.PicStyle.disabled=false;
 			  document.myform.PicWidth2.disabled=false;
 			  document.myform.PicWidth2.disabled=false;
			  
 			}
		 if(Val==2)	
			{
			 ActFs.style.display="";
			  document.myform.OpenType.disabled=true;
			  document.myform.OpenTypes.disabled=true;
 			  document.myform.TitleCss.disabled=true;
 			  document.myform.ColNumber.disabled=true;
 			  document.myform.piccss.disabled=true;
 			  document.myform.TypeTitle1.disabled=true;
 			  document.myform.TypeTitle2.disabled=true;
 			  document.myform.PicWidth2.disabled=true;
 			  document.myform.PicHeight2.disabled=true;
 			  document.myform.PicStyle.disabled=true;
 			  document.myform.PicWidth2.disabled=true;
 			  document.myform.PicWidth2.disabled=true;
			  
			   
			}
		}


 function  
  SetDiyContent(oTextarea,strText){   
  oTextarea.focus();   
  document.selection.createRange().text+=strText;   
  oTextarea.blur();   
  }   
function SelectPicStyle(ObjValue)
		{
				document.all.PicStyles.innerHTML='<img src="../../Images/share/Act'+ObjValue+'.gif" height="130" width="240" border="0">';
				if (ObjValue==1)
				document.all.ContentLenArea.style.display="";
				else
				 document.all.ContentLenArea.style.display="";

		}		
		
		
		function InsertScriptFun(Obj)
		{  
		  if (document.myform.LabelName.value=='')
			 {
			  alert('请输入标签名称');
			  document.myform.LabelName.focus(); 
			  return false
			  }
			  if (document.myform.PicStyle.value!=1 && document.myform.ContentLen.value=='')
			 {
				alert('请输入显示内容的字数!');
				  document.myform.ContentLen.focus(); 
				  return false;
			 }
			var ModeID=document.myform.ModeID.value;
			var ClassID=document.myform.ClassID.value;
			var ListNumber=document.myform.ListNumber.value;
			var TitleLen=document.myform.TitleLen.value;
			var ColNumber=document.myform.ColNumber.value;
			var PicWidth=document.myform.PicWidth.value;
			var SubClass=false,TypeTitle
			var PicHeight=document.myform.PicHeight.value;
			var ContentLen=document.myform.ContentLen.value;
			var PicStyle=document.myform.PicStyle.value;
			var ArticleSort=document.myform.ArticleSort.value;
			var OpenType=document.myform.OpenType.value;
			var TitleCss=document.myform.TitleCss.value;
			var piccss=document.myform.piccss.value;
			var ACTIF=document.myform.ACTIF.value;
			var ActF=document.myform.ActF.value;
			var ATT=document.myform.ATT.value;
			var DiyContent=document.myform.DiyContent.value;
		 	
 			if (document.myform.SubClass.checked) SubClass=true;
			var ActF=document.myform.ActF.value;
			for (var i=0;i<document.myform.TypeTitle.length;i++){
			 var TCJ = document.myform.TypeTitle[i];
			if (TCJ.checked==true)	   
				TypeTitle = TCJ.value
			}
			if  (ListNumber=='')  ListNumber=10;
			if  (TitleLen=='') TitleLen=30;
			if  (ColNumber=='') ColNumber=1;
			if  (document.myform.ArticleSort.value=='') ArticleSort="ID Desc";
			document.myform.LabelContent.value=	'{$GetArticlePic('+ClassID+'§'+document.myform.ActF.value+'§'+ATT+'§'+ArticleSort+'§'+OpenType+'§'+ListNumber+'§'+ColNumber+'§'+TitleLen+'§'+TitleCss+'§'+piccss+'§'+PicWidth+'§'+PicHeight+'§'+ContentLen+'§'+PicStyle+'§'+TypeTitle+'§'+ACTIF+'§'+DiyContent+'§'+ModeID+'§'+SubClass+')}';
			document.myform.SubmitBtn.value="正在提交数据,请稍等...";
			document.myform.SubmitBtn.disabled=true;
			document.myform.Submit.disabled=true;	
			document.myform.submit();
		}
	
</script>
<script language="javascript">SetActF(<%= ActF %>);</script>
<script language="javascript">SelectPicStyle(<%= """"&PicStyle&"""" %>);</script>
</body>
</html>
