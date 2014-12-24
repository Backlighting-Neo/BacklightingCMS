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
Dim Action,ID,LabelRS,LabelName,Descript,LabelContent,LabelFlag,LabelContentArr,ClassID,Rs,pages
Dim ArticleSort,ListNumber,TitleLen,RowHeight,ColNumber,TitleCss,DateCss,DateForm,DateAlign,NavType
Dim TypeClassName,OpenType,MoreLinkType,MoreLinkWord,MoreLinkpic,Division,TypeNew,NavPic,NavWord
dim menubg,menubgType,menubgWord,menubgpic,SubColNumber,sysdir,iftrue,ModeID,SubClass,outerfor
Dim ActF,Str,ACTIF,ATT,Boxclass,Boxid,DiyContent
Dim MainTitleCss,PicA,PicNum,PicContentNum,ForClassContent
If Not ACTCMS.ACTCMS_QXYZ(0,"bqxt","") Then   Call Actcms.Alert("对不起，你没有操作权限！","")
sysdir=actcms.ActSys&"upfiles"
Action = Request.QueryString("Action")
ID =  ChkNumeric(Request.QueryString("ID"))
IF Action = "Add" Then
	ClassID = 1
	ListNumber = 10
	TitleLen =30
	MoreLinkType = 0
	ColNumber = 2
	ATT = 0
	ActF=1
	RowHeight = 22
	menubgType =0
	SubColNumber =2
	pages = "新建循环栏目文章标签"
	ArticleSort="ID Desc"
	PicA=0
	PicNum=1
	PicContentNum=20
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
		LabelContent = Replace(Replace(LabelContent, "{$GetClassForArticleList(", ""), ")}", "")
		LabelContentArr = Split(LabelContent, "§")
		ClassID = LabelContentArr(0)
		ActF = LabelContentArr(1)'是否包含子栏目
		ATT = LabelContentArr(2)'文章属性
		ArticleSort = LabelContentArr(3)'排序方法

		OpenType =  LabelContentArr(4)'链接目标
		ListNumber =  LabelContentArr(5)'文章数量
		RowHeight = LabelContentArr(6)'文章行距
		TitleLen = LabelContentArr(7)'标题字数
		ColNumber = LabelContentArr(8)'排列列数
		TypeClassName = LabelContentArr(9)'是否显示栏目
		TypeNew = LabelContentArr(10)'图文标志
		ACTIF = LabelContentArr(11)'最新文章标志
		NavType = LabelContentArr(12)
		MoreLinkType=LabelContentArr(14)
		Division = LabelContentArr(16)
		DateForm = LabelContentArr(17)
		DateAlign = LabelContentArr(18)
		TitleCss = LabelContentArr(19)
		DateCss  = LabelContentArr(20)
		SubColNumber = LabelContentArr(21)
		
		
		
		outerfor=LabelContentArr(22)
		DiyContent=LabelContentArr(23)
  



		MainTitleCss=LabelContentArr(24)
		PicA=LabelContentArr(25)
		PicNum=LabelContentArr(26)
		PicContentNum=LabelContentArr(27)
		ForClassContent=LabelContentArr(28)
		SubClass=LabelContentArr(29)
		ModeID=LabelContentArr(30)
		IF NavType = 0 Then 
			NavWord = LabelContentArr(13)
			Navpic = ""
		Else
			NavWord = ""
			Navpic = LabelContentArr(13)
		End IF
		
	IF MoreLinkType = 0 Then 
			MoreLinkWord = LabelContentArr(15)
			MoreLinkpic = ""
		Else
			MoreLinkWord = ""
			MoreLinkpic = LabelContentArr(15)
		End If
		pages = "修改循环栏目文章标签"
	NavWord=server.HTMLEncode(NavWord)
 End IF
 With Response 

%>
<form id="myform" name="myform" method="post" action="AddLabelSave.asp">
 <input type="hidden" name="LabelContent"> 
  <input type="hidden" name="LabelType" value="1">
  <input type="hidden" name="Action" value="<%= Action %>"> 
  <input type="hidden" name="ID" value="<%= ID %>"> 
 <input type="hidden" name="FileUrl" value="GetArticleList.asp"> 
 <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
    <tr>
      <td colspan="2" align="center" class="bg_tr"><%= pages %>&nbsp;</td>
    </tr>
    <tr>
      <td width="50%" >标签名称
      <input name="LabelName" type="text" class="Ainput" id="LabelName" value="<%= LabelName %>"></td>
      <td width="50%" >&nbsp;</td>
    </tr>
   <tr>
      <td class="td_bg">标签目录      
        <select name="LabelFlag" id="select">
          <option value="0">系统默认</option>
			 <%=AF.ACT_LabelFolder(CInt(LabelFlag))%>
        </select>&nbsp;&nbsp;<a href="ACT.LabelFolder.asp"><font color=red><b>新建存放目录</b></font></a>
		&nbsp;<font color=green>标签存放目录,方便管理标签</font></td>
	 <td width="50%" height="24"  class="td_bg"> 所属模型
	 <select name="ModeID" id="ModeID">
          <option value="0" style="color:green">所有模型</option>
          <%=AF.ACT_L_Mode(CInt(ModeID))%>
        </select>	
        <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_ModeID')"  id="Label_ModeID">帮助</span
        	></td>
    </tr>

 <tr>
      <td width="50%" >
	  输出格式
		 <select  style='width:40%' name="ActF" id="ActF" onChange="SetActF(this.options[this.selectedIndex].value);"> 
	 <option value="1" <% IF ActF = 1 Then Response.Write("selected") %>>普通模式</option>
  <option value="2" <% IF ActF = 2 Then Response.Write("selected") %>>代码模式</option>
  </select> <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_ActF')"  id="Label_ActF">帮助</span>
     	 </td>
      <td width="50%" >文章属性
        <select name="ATT" id="ATT">
          <option value="0" style="color:green">普通文章</option>
          <%=ACTCMS.ACT_ATT(CInt(ATT))%>
          </select>
	 <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_ATT')"  id="Label_ATT">帮助</span> </td>
    </tr>
  <tr id=ActFs ><td  colspan="2" >
 <font color=red>外部循环</font> 
 
<a href="#" onClick='SetDiyContent(outerfor,"#outerfor")'>内部循环</a>&nbsp;
<a href="#" onClick='SetDiyContent(outerfor,"#subpic")'>内嵌图文</a>&nbsp;
<a href="#" onClick='SetDiyContent(outerfor,"#ClassName")'>栏目名称</a>&nbsp;
<a href="#" onClick='SetDiyContent(outerfor,"#ClassID")'>栏目ID</a>&nbsp;
<a href="#" onClick='SetDiyContent(outerfor,"#ClassLink")'>栏目链接</a>&nbsp;
<a href="#" onClick='SetDiyContent(outerfor,"#AutoID")'>自增长ID</a>&nbsp;
<br />

<a href="#" onClick='SetDiyContent(outerfor,"#ForClassSeo")'>栏目SEO标题</a>&nbsp;
<a href="#" onClick='SetDiyContent(outerfor,"#ForClassPicUrl")'>栏目缩略图</a>&nbsp;
<a href="#" onClick='SetDiyContent(outerfor,"#ForClassPicFile")'>栏目缩略图地址</a>&nbsp;
<a href="#" onClick='SetDiyContent(outerfor,"#ForClassKeywords")'>栏目META关键字</a>&nbsp;
<a href="#" onClick='SetDiyContent(outerfor,"#ForClassDescription")'>栏目META描述</a>&nbsp;
 
 
 
 
 
 
 内嵌图文<label for="PicA1"> 
<input id="PicA1"  <%If PicA="1" Then response.write "Checked"%>  type="radio" name="PicA" value="1" onClick="SetPicA(1);">
        是</label>
  <label for="PicA2"> 
 <input id="PicA2"  <%If PicA="0" Then response.write "Checked"%>   type="radio" name="PicA" value="0" onClick="SetPicA(2);">
        否</label>		
	<span <%If PicA="0" Then response.write "style='display:none'"%>  id='PicNum'>
	<font color=green>循环次数</font> <input Name="PicNum" class="Ainput" value="<%=PicNum%>"  size=3>	
	<font color=green>内容字数</font> <input Name="PicContentNum" class="Ainput" value="<%=PicContentNum%>"  size=3>
	</span>	
   
  <br />
<textarea  onfocus="this.className='colorfocus';" onBlur="this.className='colorblur';"  name="outerfor"  cols="95%" rows="7"><%=Server.HTMLEncode(outerfor)%></textarea>   <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_outerfor')"  id="Label_outerfor">帮助</span>  
<br /><font color=red>内部循环</font> 
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

<br />
<textarea onFocus="this.className='colorfocus';" onBlur="this.className='colorblur';"  name="DiyContent"  id="DiyContent" cols="95%" rows="10"><%=Server.HTMLEncode(DiyContent)%></textarea>  <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_DiyContent')"  id="Label_DiyContent">帮助</span> 
 <br />
 
 <span id="PicAa" <%If PicA="0" Then response.write "style='display:none'"%>>
 <font color=red>内嵌图文</font> 
<input  type="button"  class="ACT_btn" onClick='SetDiyContent(ForClassContent,"#Link")' value="文章链接"> &nbsp;
<input  type="button"  class="ACT_btn" onClick='SetDiyContent(ForClassContent,"#Title")' value="文章标题"> &nbsp;
<input  type="button"  class="ACT_btn" onClick='SetDiyContent(ForClassContent,"#PicUrl")' value="图片地址"> &nbsp;
<input  type="button"  class="ACT_btn" onClick='SetDiyContent(ForClassContent,"#Intro")' value="文章内容">&nbsp;图文样式&nbsp;<br />
<textarea  onfocus="this.className='colorfocus';" onBlur="this.className='colorblur';"  name="ForClassContent" cols="95%" rows="10"><%=Server.HTMLEncode(ForClassContent)%></textarea>
  </span>
<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_PicAa')"  id="Label_PicAa">帮助</span> 
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
   <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_ClassID')"  id="Label_ClassID">帮助</span>  
     </td>
	
	</td>
	  
	       
      <td >
	 <input id="SubClass22" type="checkbox" value="true" name="SubClass"  <%If InStr(ClassID,",") > "1"  Then  .write "disabled=true  "%> <%If CBool(SubClass)=true  Then  .write "Checked  "%>>
	  <label for="SubClass22">允许包含子栏目</label>
      <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_SubClass')"  id="Label_SubClass">帮助</span> </td>
    </tr>
    <tr>
      <td >表头CSS样式名称	 <input  type="text"  class="ainput"  name="MainTitleCss" value="<%=MainTitleCss%>" >
	  <span class="h" style="cursor:help;"  onclick="dohelp('Label__MainTitleCss')"  id="Label__MainTitleCss">帮助</span></td>
    <td width="50%" height="24" >

	</td>
    </tr>
 
    <tr>
      <td >
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
    <span class="h" style="cursor:help;"  onclick="dohelp('Label__ArticleSort')"  id="Label__ArticleSort">帮助</span>    
        	</td>
      <td >

	  
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
   <span class="h" style="cursor:help;"  onclick="dohelp('Label__OpenType')"  id="Label__OpenType">帮助</span>      
        
        	</td>
    </tr>
    <tr>
      <td >文章数量
      <input name="ListNumber" type="text" class="Ainput" id="ListNumber2"    style="width:70%;"  value="<%= ListNumber %>">
       <span class="h" style="cursor:help;"  onclick="dohelp('Label__ListNumber')"  id="Label__ListNumber">帮助</span> </td>
      <td >文章行距
      <input name="RowHeight" type="text" class="Ainput" id="RowHeight2"    style="width:70%;" value="<%= RowHeight %>">
      <span class="h" style="cursor:help;"  onclick="dohelp('Label__RowHeight')"  id="Label__RowHeight">帮助</span>
      </td>
    </tr>
    <tr>
      <td >标题字数
      <input name="TitleLen"  type="text" class="Ainput"    style="width:70%;" value="<%= TitleLen %>">
       <span class="h" style="cursor:help;"  onclick="dohelp('Label__TitleLen')"  id="Label__TitleLen">帮助</span></td>
      <td >排列列数
        <input type="text" class="Ainput"   size="4" value="<%= ColNumber %>" name="ColNumber"> 
        子栏目排列列数 
        <input name="SubColNumber" type="text" class="Ainput" id="SubColNumber" value="<%= SubColNumber %>"   size="4">
       <span class="h" style="cursor:help;"  onclick="dohelp('Label__ColNumber_SubColNumber')"  id="Label__ColNumber_SubColNumber">帮助</span> 
        </td>
    </tr>
    <tr>
      <td colspan="1" >附加显示
        <label for="TypeClassName"><input name="TypeClassName" type="checkbox" id="TypeClassName"  value="<%= TypeClassName %>"  <% IF TypeClassName = "true" Then Response.Write("Checked") %>>
        显示栏目</label>&nbsp;&nbsp;&nbsp;
		<% 	
		  .Write "&nbsp;&nbsp;&nbsp;"
		 If  cbool(TypeNew) = True Then
		  .Write ("<label for=""TypeNew1""><input id=""TypeNew1"" type=""checkbox"" value=""true"" name=""TypeNew"" checked>最新文章标志</label>")
		 Else
		  .Write ("<label for=""TypeNew2""><input id=""TypeNew2"" type=""checkbox"" value=""true"" name=""TypeNew"">最新文章标志</label>")
		 End If
		 .Write "&nbsp;&nbsp;&nbsp;"

 %>   <span class="h" style="cursor:help;"  onclick="dohelp('Label__TypeClassName')"  id="Label__TypeClassName">帮助</span>
            </td>
 <td>SQL判断条件<input type="text" class="Ainput"  style="width:60%;"  size="4" value="<%= ACTIF %>" name="ACTIF"><font color=red>格式</font>
 <span class="h" style="cursor:help;"  onclick="dohelp('Label__ACTIF')"  id="Label__ACTIF">帮助</span>
 </td>
    </tr>
    <tr >
      <td >导航类型
       <label for="NavType1">  <input id="NavType1"  <% IF NavType = 0 Then Response.Write("Checked") %> type="radio" name="NavType" value="0" onClick="SetNavStatus(1);">
        文字导航</label>
       <label for="NavType2">  <input id="NavType2" <% IF NavType = 1 Then Response.Write("Checked") %>   type="radio" name="NavType" value="1" onClick="SetNavStatus(2);">
        图片导航</label>	
         <span class="h" style="cursor:help;"  onclick="dohelp('Label__NavType')"  id="Label__NavType">帮助</span>
         	</td>
    <td width="50%" height="24"   id=SetNavStatus1 
	<% if NavType=1 then %>
	style="DISPLAY: none" 
	<% end if %>>
 <input name="NavWord" type="text" class="Ainput" id="NavWord" style="width:60%;" value="<%= NavWord %>"> 
 支持HTML语法</td>
   <td width="50%" height="24"   id=SetNavStatus2 
	<% if NavType=0 then %>
	style="DISPLAY: none"
	SetNavStatus1.style.display=''
	<% end if %> >
<input name="NavPic" type="text" class="Ainput" id="NavPic" style="width:60%;" value="<%= NavPic %>" readonly>
<input class="ACT_btn" type="button"  onClick="OpenWindowAndSetValue('../print/SelectPic.asp?CurrPath=<%=sysdir%>',500,320,window,document.myform.NavPic);" value="选择图片...">&nbsp;<span style="cursor:hand;color:green;" onClick="javascript:document.myform.NavPic.value='';" onMouseOver="this.style.color='red'" onMouseOut="this.style.color='green'">清除</span>
<span class="h" style="cursor:help;"  onclick="dohelp('Label__NavPic')"  id="Label__NavPic">帮助</span>
</td>  
    </tr>
    <tr id="MoreLinkArea">
      <td >更多链接
         <label for="MoreLinkType1">
        <input id="MoreLinkType1"  <% IF MoreLinkType = 0 Then Response.Write("Checked") %> type="radio" name="MoreLinkType" value="0" onClick=SetMoreLinkType(1)>
文字链接</label>
        <label for="MoreLinkType2">
        <input id="MoreLinkType2" <% IF MoreLinkType = 1 Then Response.Write("Checked") %>   type="radio" name="MoreLinkType" value="1" onClick=SetMoreLinkType(2)>
图片链接</label>
<span class="h" style="cursor:help;"  onclick="dohelp('Label__MoreLinkType')"  id="Label__MoreLinkType">帮助</span>
 </td>
      <td width="50%" height="24"   id=SetMoreLinkType1 
	<% if MoreLinkType=1 then %>
	style="DISPLAY: none" 
	<% end if %>>
 <input type="text" class="Ainput" name="MoreLinkWord" style="width:60%;" value="<%= MoreLinkWord %>"> 支持HTML语法 </td>
   <td width="50%" height="24"   id=SetMoreLinkType2 
	<% if MoreLinkType=0 then %>
	style="DISPLAY: none"
	SetMoreLinkType1.style.display=''
	<% end if %> >
	<input name="MoreLinkPic" type="text" class="Ainput" id="MoreLinkpic" style="width:60%;" value="<%= MoreLinkpic %>" readonly>
	<input class="ACT_btn" type="button" onClick="OpenWindowAndSetValue('../print/SelectPic.asp?CurrPath=<%=sysdir%>',500,320,window,document.myform.MoreLinkpic);" name="Submit3" value="选择图片...">
	&nbsp;<span style="cursor:hand;color:green;" onClick="javascript:document.myform.MoreLinkpic.value='';" onMouseOver="this.style.color='red'" onMouseOut="this.style.color='green'">清除</span>
    <span class="h" style="cursor:help;"  onclick="dohelp('Label__MoreLinkPic')"  id="Label__MoreLinkPic">帮助</span>
     </td>
    </tr>
    <tr>
      <td colspan="2" >分隔图片
      <input name="Division" type="text" class="Ainput" id="Division" style="width:60%;" value="<%= Division %>" readonly>
	  <input class="ACT_btn" type="button" id="Division1" onClick="OpenWindowAndSetValue('../print/SelectPic.asp?CurrPath=<%=sysdir%>',500,320,window,document.myform.Division);" name="Submit3" value="选择图片...">
	  &nbsp;<span style="cursor:hand;color:green;" onClick="javascript:document.myform.Division.value='';" onMouseOver="this.style.color='red'" onMouseOut="this.style.color='green'">清除</span>
<span class="h" style="cursor:help;"  onclick="dohelp('Label__Division')"  id="Label__Division">帮助</span> 
     </td>
    </tr>
    <tr>
      <td >日期格式
        <select  style="width:70%;" name="DateForm" id="select2">
 <% 
		.Write AF.ACT_DateStr(DateForm)
			 %>        </select>      </td>
      <td >日期对齐

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
		End With %> 
<span class="h" style="cursor:help;"  onclick="dohelp('Label__DateForm')"  id="Label__DateForm">帮助</span> 
           </td>
    </tr>
    <tr>
      <td >标题样式
      <input name="TitleCss" type="text" class="Ainput" id="TitleCss" style="width:70%;" value="<%= TitleCss %>">
      <span class="h" style="cursor:help;"  onclick="dohelp('Label__TitleCss')"  id="Label__TitleCss">帮助</span></td>
      <td >日期样式
      <input name="DateCss" type="text" class="Ainput" id="DateCss" style="width:70%;" value="<%= DateCss %>">
      <span class="h" style="cursor:help;"  onclick="dohelp('Label__DateCss')"  id="Label__DateCss">帮助</span>
      </td>
    </tr>
    <tr>
      <td colspan="2" align="center" >
<input name="SubmitBtn" class="ACT_btn" type="button"  onClick="InsertScriptFun()"  id="SubmitBtn"  value=" 保 存 ">       &nbsp;&nbsp;
       <input type="reset" class="ACT_btn" name="Submit" value="  重置  ">	     </td>
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
		MoreLinkArea.style.display='';
 		document.myform.SubClass.disabled=false;
		}
		else
		{
		MoreLinkArea.style.display='none';
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
			  document.myform.NavType1.disabled=false;
 			  document.myform.NavType2.disabled=false;
			  document.myform.RowHeight.disabled=false;
 			  document.myform.Division.disabled=false;
			  document.myform.Division1.disabled=false;
			  document.myform.NavWord.disabled=false;
			  document.myform.NavPic.disabled=false;
			  document.myform.MoreLinkWord.disabled=false;
			  document.myform.MoreLinkPic.disabled=false;
			  document.myform.TitleCss.disabled=false;
			  document.myform.DateCss.disabled=false;
			  document.myform.ColNumber.disabled=false;
			  document.myform.TypeClassName.disabled=false;
			  document.myform.TypeNew.disabled=false;
			  
			  document.myform.DateAlign.disabled=false;
 			  document.myform.MoreLinkType1.disabled=false;
			  document.myform.MoreLinkType2.disabled=false;
			}
		 if(Val==2)	
			{
			 ActFs.style.display="";
			  document.myform.OpenType.disabled=true;
			  document.myform.OpenTypes.disabled=true;
 			  document.myform.MoreLinkType1.disabled=true;
			  document.myform.MoreLinkType2.disabled=true;
			  document.myform.NavType1.disabled=true;
			  document.myform.NavType2.disabled=true;
 			  document.myform.Division.disabled=true;
			  document.myform.Division1.disabled=true;
			  document.myform.NavWord.disabled=true;
			  document.myform.NavPic.disabled=true;
			  document.myform.MoreLinkWord.disabled=true;
			  document.myform.MoreLinkPic.disabled=true;
			  document.myform.TitleCss.disabled=true;
			  document.myform.DateCss.disabled=true;
			  document.myform.ColNumber.disabled=true;
			  document.myform.TypeClassName.disabled=true;
			  document.myform.TypeNew.disabled=true;
			  document.myform.RowHeight.disabled=true;
			  
			  document.myform.DateAlign.disabled=true;
  			}
		}

function OpenWindow(Url,Width,Height,WindowObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;');
	return ReturnStr;
}	
//Open Modal Window
function OpenWindowAndSetValue(Url,Width,Height,WindowObj,SetObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:yes;help:no;scroll:yes;');
	if (ReturnStr!='') SetObj.value=ReturnStr;
	return ReturnStr;
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
function Getcolor(img_val,Url,input_val){
	var arr = showModalDialog(Url, "", "dialogWidth:18.5em; dialogHeight:17.5em; status:0; help:0");
	if (arr != null){
		document.getElementById(input_val).value = arr;
		img_val.style.backgroundColor = arr;
		}
}

function SetMoreLinkType(n){
			if (n==1){
			SetMoreLinkType1.style.display='';
			SetMoreLinkType2.style.display='none';	
			}
		  else{
			SetMoreLinkType1.style.display='none';
			SetMoreLinkType2.style.display='';	
		}
}



function SetPicA(n){
			if (n==1){
			PicAa.style.display='';
			PicNum.style.display='';

			}
		  else{
			PicAa.style.display='none';
			PicNum.style.display='none';
		}
}

 

  function 
  SetDiyContent(oTextarea,strText){   
   oTextarea.focus();   
  document.selection.createRange().text+=strText;   
  oTextarea.blur();   
  } 
		function InsertScriptFun()
		{   if (document.myform.LabelName.value=='')
			 {
			  alert('请输入标签名称');
			  document.myform.LabelName.focus(); 
			  return false
			  }
			var ClassID=document.myform.ClassID.value;
			var SubClass=false,NavType=1;MoreLinkType;
			var TypeClassName,TypeNew,PicA;
			var OpenType=document.myform.OpenType.value;
		
			
			var ArticleSort=document.myform.ArticleSort.value;
			var Division=document.myform.Division.value;
			var DateAlign=document.myform.DateAlign.value;
			var TitleCss=document.myform.TitleCss.value;
			var DateCss=document.myform.DateCss.value;
			var ACTIF=document.myform.ACTIF.value;
			var ListNumber=document.myform.ListNumber.value;
			var ListNumber=document.myform.ListNumber.value;
			
			
			var ListNumber=document.myform.ListNumber.value;
			var RowHeight=document.myform.RowHeight.value;
			var TitleLen=document.myform.TitleLen.value;
			var ColNumber=document.myform.ColNumber.value;
			var SubColNumber=document.myform.SubColNumber.value;
			var Nav,NavType=document.myform.NavType.value;
			var MoreLink,MoreLinkType=document.myform.MoreLinkType.value;
			var DateForm=document.myform.DateForm.value;
			var ActF=document.myform.ActF.value;
			var ATT=document.myform.ATT.value;
			var MainTitleCss=document.myform.MainTitleCss.value;
			var PicContentNum=document.myform.PicContentNum.value;
			var ForClassContent=document.myform.ForClassContent.value;
			var PicNum=document.myform.PicNum.value;
  
 			var outerfor=document.myform.outerfor.value;

  			var DiyContent=document.myform.DiyContent.value;

			
			
			
			
			if (document.myform.SubClass.checked) SubClass=true;
			if (document.myform.TypeClassName.checked)
				TypeClassName = true
			else
			    TypeClassName =false;

			if (document.myform.TypeNew.checked)
			   TypeNew= true
			else
			   TypeNew=false;
		
			
			
	
						
		  for (var i=0;i<document.myform.PicA.length;i++){
			 var TCJ = document.myform.PicA[i];
			if (TCJ.checked==true)	   
				PicA = TCJ.value
			}
			
			
			
			for (var i=0;i<document.myform.NavType.length;i++){
			 var TCJ = document.myform.NavType[i];
			if (TCJ.checked==true)	   
				NavType = TCJ.value
			if  (NavType==0) Nav=''+document.myform.NavWord.value+''
			 else  Nav=''+document.myform.NavPic.value+'';
			}
			
			for (var i=0;i<document.myform.MoreLinkType.length;i++){
			 var TCJ = document.myform.MoreLinkType[i];
			if (TCJ.checked==true)	   
				MoreLinkType = TCJ.value
			if  (MoreLinkType==0) MoreLink=''+document.myform.MoreLinkWord.value+''
			 else  MoreLink=''+document.myform.MoreLinkPic.value+'';
			}
			
			if  (ListNumber=='')  ListNumber=10;
			if (RowHeight=='') RowHeight=20
			if  (TitleLen=='') TitleLen=30;
			if  (ColNumber=='') ColNumber=2;
			if  (SubColNumber=='') SubColNumber=2;
			if  (PicContentNum=='') PicContentNum=20;
			if  (PicNum==0) PicNum=1;
			
			if  (isNaN(PicNum)) PicNum=1;

			var ModeID=document.myform.ModeID.value;
			//-----------------------------
			if  (document.myform.ArticleSort.value=='') ArticleSort="ID Desc";
			document.myform.LabelContent.value='{$GetClassForArticleList('+ClassID+'§'+ActF+'§'+ATT+'§'+ArticleSort+'§'+OpenType+'§'+ListNumber+'§'+RowHeight+'§'+TitleLen+'§'+ColNumber+'§'+TypeClassName+'§'+TypeNew+'§'+ACTIF+'§'+NavType+'§'+Nav+'§'+MoreLinkType+'§'+MoreLink+'§'+Division+'§'+DateForm+'§'+DateAlign+'§'+TitleCss+'§'+DateCss+'§'+SubColNumber+'§'+outerfor+'§'+DiyContent+'§'+MainTitleCss+'§'+PicA+'§'+PicNum+'§'+PicContentNum+'§'+ForClassContent+'§'+SubClass+'§'+ModeID+')}';
			document.myform.SubmitBtn.value="正在提交数据,请稍等...";
			document.myform.SubmitBtn.disabled=true;
			document.myform.Submit.disabled=true;	
			document.myform.submit();
		}
	
</script><script language="javascript">SetActF(<%= ActF %>);</script>
<%Call CloseConn%>
</body>
</html>
