<!--#include file="../../ACT.Function.asp"-->
<!--#include file="../ACT.F.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ACT_标签管理</title>
<link href="../../Images/editorstyle.css" rel="stylesheet" type="text/css">
<script charset="utf-8"  language="JavaScript" type="text/javascript" src="../../../editor/kindeditor/kindeditor.js" ></script>
<script language="JavaScript" type="text/javascript" src="../../../ACT_inc/js/lhgcore/lhgcore.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../../ACT_inc/js/lhgcore/lhgdialog.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../../ACT_inc/js/lhgcore/Main.js"></script>
</head>
<body>
<%
Dim Action,ID,LabelRS,LabelName,Descript,LabelContent,LabelFlag,LabelContentArr,ClassID,ClassName,Rs,pages
Dim ArticleSort,TitleLen,RowHeight,ColNumber,TitleCss,DateCss,DateForm,DateAlign,NavType,ListNumber,ACTIF
Dim sysdir,TypeClassName,OpenType,MoreLinkType,MoreLinkWord,MoreLinkpic,Division,TypeNew,NavPic,NavWord,DiyContent,iftrue
If Not ACTCMS.ACTCMS_QXYZ(0,"bqxt","") Then   Call Actcms.Alert("对不起，你没有操作权限！","")
sysdir=actcms.ActSys&"upfiles"
Action = Request.QueryString("Action")
ID =  ChkNumeric(Request.QueryString("ID"))
Dim ActF,Str
IF Action = "Add" Then
	ClassID = 0
	TitleLen =30
	ColNumber = 1
	RowHeight = 20
	ListNumber = 10
	ActF=1
	pages = "新建相关文章列表标签"
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
		LabelContent = Replace(Replace(LabelContent, "{$CorrelationArticleList(", ""), ")}", "")
		LabelContentArr = Split(LabelContent, "§")
		ActF = LabelContentArr(0)
		ArticleSort = LabelContentArr(1)'排序方法
		ListNumber = LabelContentArr(2)'文章数量
		OpenType =  LabelContentArr(3)'链接目标
		RowHeight = LabelContentArr(4)'文章行距
		TitleLen = LabelContentArr(5)'标题字数
		ColNumber = LabelContentArr(6)'排列列数
		TypeClassName = LabelContentArr(7)'是否显示栏目
		TypeNew = LabelContentArr(8)'最新文章标志
		NavType = LabelContentArr(9)'导航类型
		
		Division = LabelContentArr(11)'分隔图片
		DateForm = LabelContentArr(12)'日期格式
		DateAlign = LabelContentArr(13)'日期对齐
		TitleCss = LabelContentArr(14)'标题样式
		DateCss  = LabelContentArr(15)'日期样式
		
		'ActF=LabelContentArr(16)
		ACTIF=LabelContentArr(16)
		DiyContent=LabelContentArr(17)
		IF NavType = 0 Then 
			NavWord = LabelContentArr(10)
			Navpic = ""
		Else
			NavWord = ""
			Navpic = LabelContentArr(10)
		End If
		pages = "修改相关栏目文章列表标签"
End IF
 With Response 
%>
<form id="myform" name="myform" method="post" action="AddLabelSave.asp">
 <input type="hidden" name="LabelContent"> 
  <input type="hidden" name="LabelType" value="1">
  <input type="hidden" name="Action" value="<%= Action %>"> 
  <input type="hidden" name="ID" value="<%= ID %>"> 
 <input type="hidden" name="FileUrl" value="CorrelationArticleList.asp">  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
    <tr>
      <td colspan="2" align="center" class="bg_tr"><%= pages %>&nbsp;</td>
    </tr>
    <tr>
      <td width="50%" class="tdclass" >标签名称
      <input name="LabelName" type="text" class="Ainput" id="LabelName" value="<%= LabelName %>"></td>
      <td width="50%" class="tdclass">&nbsp;</td>
    </tr>
    <tr>
      <td class="tdclass">标签目录      
        <select name="LabelFlag" id="select">
          <option value="0">系统默认</option>
			 <%=AF.ACT_LabelFolder(CInt(LabelFlag))%>
        </select>&nbsp;&nbsp;<a href="ACT.LabelFolder.asp"><font color=red><b>新建存放目录</b></font></a></td>
	 <td width="50%" height="24" class="tdclass" ><font color=green>标签存放目录,日后方便管理标签</font></td>
    </tr>
  
 <tr>
      <td width="50%" class="tdclass">
 输出模式
		 <select  style='width:40%' name="ActF" id="ActF" onChange="SetActF(this.options[this.selectedIndex].value);"> 
	 <option value="1" <% IF ActF = 1 Then Response.Write("selected") %>>普通模式</option>
  <option value="2" <% IF ActF = 2 Then Response.Write("selected") %>>代码模式</option>
  </select>  
		 </td>
      <td width="50%" class="tdclass"><font color=red>请选择系统支持的输出格式</font>
	  </td>
    </tr>
  <tr id=ActFs ><td  colspan="2" class="tdclass">
  

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
<textarea   onfocus="this.className='colorfocus';" onBlur="this.className='colorblur';" name="DiyContent" id="DiyContents"  cols="95%" rows="10"><%=Server.HTMLEncode(DiyContent)%></textarea>


  
  </td>
	</tr>
   

    <tr>
      <td class="tdclass">排序方法
	
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
		
		</td>
      <td class="tdclass">链接目标  
	<input name="OpenType" id="OpenType" type="text" class="Ainput" value="<%=OpenType%>" size="8">

	  <%iftrue=false%>
	<select   name="OpenTypes"  onchange="document.myform.OpenType.value=this.value">
          <option value="_blank"   style="color:green"  <%If OpenType="_blank" Then .write "selected":iftrue=true%>>新窗口打开</option>
          <option value="_parent" <%If OpenType="_parent" Then .write "selected":iftrue=true%>>父窗口打开</option>
          <option value="_self" <%If OpenType="_self" Then .write "selected":iftrue=true%>>本窗口打开</option>
          <option value="_top" <%If OpenType="_top" Then .write "selected":iftrue=true%>>主窗口打开</option>
		<option value='' style="color:red"  <%If iftrue=false Then .write "selected"%>>自定义</option>	
		</select>	</td>
    </tr>
    <tr>
      <td class="tdclass">
	  文章数量
	   	  <input name="ListNumber" type="text" class="Ainput" id="ListNumber" value="<%= ListNumber %>" onKeyUp="value=value.replace(/[^\d]/g,'') "onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))">	  </td>
      <td class="tdclass">文章行距
      <input name="RowHeight" type="text" class="Ainput" id="RowHeight2"    style="width:70%;" value="<%= RowHeight %>"></td>
    </tr>
    <tr>
      <td class="tdclass">标题字数
      <input name="TitleLen"  type="text" class="Ainput"    style="width:70%;" value="<%= TitleLen %>"></td>
      <td class="tdclass">排列列数
        <input type="text" class="Ainput"   size="4" value="<%= ColNumber %>" name="ColNumber"></td>
    </tr>
    <tr>
      <td colspan="1" class="tdclass">附加显示
        <input name="TypeClassName" type="checkbox" id="TypeClassName"  value="<%= TypeClassName %>"  <% IF TypeClassName = "true" Then Response.Write("Checked") %>><label for="TypeClassName">显示栏目</label>&nbsp;&nbsp;&nbsp;     
		<% 	
		  .Write "&nbsp;&nbsp;&nbsp;"
		 If  cbool(TypeNew) = True Then
		  .Write ("<input id=""TypeNew1"" type=""checkbox"" value=""true"" name=""TypeNew"" checked><label for=""TypeNew1"">最新文章标志</label>")
		 Else
		  .Write ("<input id=""TypeNew2"" type=""checkbox"" value=""true"" name=""TypeNew""><label for=""TypeNew2"">最新文章标志</label>")
		 End If

 %>             </td>
 <td class="tdclass">SQL判断条件<input type="text" class="Ainput"  style="width:60%;"  size="4" value="<%= ACTIF %>" name="ACTIF"><font color=red>格式</font></td>
    </tr>
    <tr >
      <td class="tdclass">导航类型
        <input id="NavType1"  <% IF NavType = 0 Then Response.Write("Checked") %> type="radio" name="NavType" value="0" onClick="SetNavStatus(1);"><label for="NavType1">文字导航</label>
         <input id="NavType2" <% IF NavType = 1 Then Response.Write("Checked") %>   type="radio" name="NavType" value="1" onClick="SetNavStatus(2);"><label for="NavType2">图片导航</label>		</td>
    <td class="tdclass" width="50%" height="24"   id=SetNavStatus1 
	<% if NavType=1 then %>
	style="DISPLAY: none" 
	<% end if %>>
 <input name="NavWord" type="text" class="Ainput" id="NavWord" style="width:70%;" value="<%= NavWord %>"> 
 支持HTML语法</td>
   <td class="tdclass" width="50%" height="24"   id=SetNavStatus2 
	<% if NavType=0 then %>
	style="DISPLAY: none"
	SetNavStatus1.style.display=''
	<% end if %> >
<input name="NavPic" type="text" class="Ainput" id="NavPic" style="width:250;" value="<%= NavPic %>" readonly>
<input class="ACT_btn" type="button"  onClick="OpenWindowAndSetValue('../print/SelectPic.asp?CurrPath=<%=sysdir%>',500,320,window,document.myform.NavPic);" value="选择图片...">&nbsp;<span style="cursor:hand;color:green;" onClick="javascript:document.myform.NavPic.value='';" onMouseOver="this.style.color='red'" onMouseOut="this.style.color='green'">清除</span></td>  
    </tr>
    <tr>
      <td colspan="2" class="tdclass">分隔图片
      <input name="Division" type="text" class="Ainput" id="Division" style="width:61%;" value="<%= Division %>" readonly>
	  <input class="ACT_btn" type="button"  id="Division1" onClick="OpenWindowAndSetValue('../print/SelectPic.asp?CurrPath=<%=sysdir%>',500,320,window,document.myform.Division);" name="Submit3" value="选择图片...">
	 &nbsp;<span style="cursor:hand;color:green;" onClick="javascript:document.myform.Division.value='';" onMouseOver="this.style.color='red'" onMouseOut="this.style.color='green'">清除</span>	  </td>
    </tr>
    <tr>
      <td class="tdclass">日期格式
        <select  style="width:70%;" name="DateForm" id="select2">
    <% 
		.Write AF.ACT_DateStr(DateForm)
			 %>
        </select>      </td>
      <td class="tdclass">日期对齐

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
		End With %>    </td>
    </tr>
    <tr>
      <td class="tdclass">标题样式
      <input name="TitleCss" type="text" class="Ainput" id="TitleCss" style="width:70%;" value="<%= TitleCss %>"></td>
      <td class="tdclass">日期样式
      <input name="DateCss" type="text" class="Ainput" id="DateCss" style="width:70%;" value="<%= DateCss %>"></td>
    </tr>
    <tr>
      <td colspan="2" align="center" class="tdclass">
	   <input name="SubmitBtn" class="ACT_btn" type="button"  onClick="InsertScriptFun()"  id="SubmitBtn"  value="   保 存   ">	
	    &nbsp;&nbsp; <input type="reset" class="ACT_btn" name="Submit" value="   重  置   ">  </td>
    </tr>
    </table>
  
</form>
<script language="javascript" >
function upload(iname) 
{

  J('#'+iname).dialog({ id:'actcmssc'+iname+'s' ,title:'自定义字段', loadingText:'上传加载中...', page: '<%=actcms.actsys&actcms.adminurl%>/include/Label/diyField.asp?ModeID=0&instrname='+iname+ "&" + Math.random(),fixed:true, left:300, top:300 ,  width:720, height:240 });
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


		function SetActF(Val)
		{
		 if(Val==1)	
			{
			 ActFs.style.display="none";
			  document.myform.OpenType.disabled=false;
			  document.myform.OpenTypes.disabled=false;
			  document.myform.NavType1.disabled=false;
 			  document.myform.NavType2.disabled=false;
 			  document.myform.Division.disabled=false;
			  document.myform.Division1.disabled=false;
			  document.myform.NavWord.disabled=false;
			  document.myform.NavPic.disabled=false;
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
 			  document.myform.Division.disabled=true;
			  document.myform.Division1.disabled=true;
			  document.myform.NavWord.disabled=true;
			  document.myform.NavPic.disabled=true;
 			  document.myform.TitleCss.disabled=true;
			  document.myform.DateCss.disabled=true;
			  document.myform.ColNumber.disabled=true;
			  document.myform.TypeClassName.disabled=true;
			  document.myform.TypeNew.disabled=true;
			  
			  document.myform.DateAlign.disabled=true;
 			}
		} 
		
		
function   SetDiyContent(oTextarea,strText){   
  oTextarea.focus();   
  document.selection.createRange().text+=strText;   
  oTextarea.blur();   
  }   

 function OpenWindowAndSetValue(Url,Width,Height,WindowObj,SetObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:yes;help:no;scroll:yes;');
	if (ReturnStr!='') SetObj.value=ReturnStr;
	return ReturnStr;
}	
		
		function InsertScriptFun()
		{   if (document.myform.LabelName.value=='')
			 {
			  alert('请输入标签名称');
			  document.myform.LabelName.focus(); 
			  return false
			  }
			var NavType=1;
			var TypeClassName,TypeWordPic,TypeNew,TypeHot;
			var RowHeight=document.myform.RowHeight.value;
			var TitleLen=document.myform.TitleLen.value;
			var ListNumber=document.myform.ListNumber.value;
			var ColNumber=document.myform.ColNumber.value;
			var Nav,NavType=document.myform.NavType.value;
			var DateForm=document.myform.DateForm.value;
			var ActF=document.myform.ActF.value;


			
			var OpenType=document.myform.OpenType.value;
			var ArticleSort=document.myform.ArticleSort.value;
			var Division=document.myform.Division.value;
			var DateAlign=document.myform.DateAlign.value;
			var TitleCss=document.myform.TitleCss.value;
			var DateCss=document.myform.DateCss.value;



			var ACTIF=document.myform.ACTIF.value;
			var DiyContent=document.myform.DiyContent.value;

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
			for (var i=0;i<document.myform.NavType.length;i++){
			 var TCJ = document.myform.NavType[i];
			if (TCJ.checked==true)	   
				NavType = TCJ.value
			if  (NavType==0) Nav=document.myform.NavWord.value
			 else  Nav=document.myform.NavPic.value;
			}
			if  (document.myform.ArticleSort.value=='') ArticleSort="ID Desc";
			
			document.myform.LabelContent.value=	'{$CorrelationArticleList('+ActF+'§'+ArticleSort+'§'+ListNumber+'§'+OpenType+'§'+RowHeight+'§'+TitleLen+'§'+ColNumber+'§'+TypeClassName+'§'+TypeNew+'§'+NavType+'§'+Nav+'§'+Division+'§'+DateForm+'§'+DateAlign+'§'+TitleCss+'§'+DateCss+'§'+ACTIF+'§'+DiyContent+')}';
			document.myform.SubmitBtn.value="正在提交数据,请稍等...";
			document.myform.SubmitBtn.disabled=true;
			document.myform.Submit.disabled=true;	
			document.myform.submit();
		}
	
</script><script language="javascript">SetActF(<%= ActF %>);</script>
</body>
</html>
