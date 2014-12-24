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
Dim Action,ID,LabelRS,LabelName,Descript,LabelContent,LabelFlag,LabelContentArr,ClassID,ClassName,Rs,pages
Dim ListNumber,TitleLen,ColNumber,TitleCss,PicHeight,PicWidth,TypeSlide,ModeID,iftrue,ActF,DiyContent
Dim OpenType,sysdir,TypeTitle,piccss,IntroNumber
If Not ACTCMS.ACTCMS_QXYZ(0,"bqxt","") Then   Call Actcms.Alert("对不起，你没有操作权限！","")

sysdir=actcms.ActSys&"upfiles"
Action = Request.QueryString("Action")
ID =  ChkNumeric(Request.QueryString("ID"))
IF Action = "Add" Then
	ModeID =  ChkNumeric(Request.QueryString("ModeID"))
 	ClassID = 0
	ListNumber = 10
	ColNumber = 1
 	TitleLen = 30
 	ActF=1
	LabelFlag=1
	TypeTitle = True
	pages = "新建幻灯片文章标签"
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
		LabelFlag = ChkNumeric(LabelRS("LabelFlag"))
 		LabelRS.Close:Set LabelRS = Nothing
		LabelContent = Replace(Replace(LabelContent, "{$GetSlide(", ""), ")}", "")
		LabelContent = Replace(LabelContent, """", "") 
		LabelContentArr = Split(LabelContent, "§")
		ClassID = LabelContentArr(0)
		ListNumber = LabelContentArr(1)'列出条数
		TitleLen = LabelContentArr(2)'链接目标
		ModeID = LabelContentArr(3)
 		DiyContent = LabelContentArr(4)
		IntroNumber= LabelContentArr(5)
 		pages = "修改新建幻灯片文章标签"
End IF
 %>
<form id="myform" name="myform" method="post" action="AddLabelSave.asp">
 <input type="hidden" name="LabelContent"> 
  <input type="hidden" name="Action" value="<%= Action %>"> 
  <input type="hidden" name="ID" value="<%= ID %>"> 
 <input type="hidden" name="FileUrl" value="GetSlide.asp">  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
    <tr>
      <td colspan="2" align="center" class="bg_tr"><%= pages %>&nbsp;</td>
    </tr>
    <tr>
      <td width="32%" >标签名称
      <input name="LabelName" class="Ainput" type="text" id="LabelName" value="<%= LabelName %>"></td>
      <td width="68%" ><font color="red">      标签目录      
      <select name="LabelFlag" id="select">
            <option value="0">系统默认</option>
            <%=AF.ACT_LabelFolder(CInt(LabelFlag))%>
          </select>
      &nbsp;&nbsp;<a href="ACT.LabelFolder.asp"><font color=red><b>新建存放目录</b></font></a> &nbsp;<font color=green>标签存放目录,方便管理标签</font> * 调用格式"{ACTCMS_标签名称}"</font></td>
    </tr>
   <tr>
      <td width="32%" ><font color="red">所属模型
      <select name="ModeID" id="ModeID">
            <option value="0" style="color:green">模型通用</option>
            <%=AF.ACT_L_Mode(CInt(ModeID))%>
          </select>
      </font><span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_ModeID')"  id="Label_ModeID">帮助</span> </td>
	 <td width="68%" height="24"><font color="red">
	 所属栏目
        <input name="ClassID" type="text"  class="Ainput" id="ClassID" value="<%= ClassID %>" readonly  disabled=true>
        <select name="select1" onChange="SelectClass();">
          <option value="0" <% IF ClassID = "0" Then  Response.Write "selected" %>>不指定栏目</option>
          <option value="1"  style="color:red" <% IF ClassID = "1" Then  Response.Write "selected" %>>当前栏目通用</option>
          <option value="2" <% IF ClassID <> "0" And ClassID <> "1"  Then  Response.Write "selected" %>>指定栏目</option>
        </select>
        <a href="#" onClick="SelectClass()">快速打开</a> <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_ClassID')"  id="Label_ClassID">帮助</span> </font></td>
    </tr>
    
  <tr><td  colspan="2" >

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
<font color=red>导读字数</font>
<input name="IntroNumber" id="IntroNumber" type="text" class="Ainput" value="<%=IntroNumber%>" size="10">
<br />
<textarea onFocus="this.className='colorfocus';" onBlur="this.className='colorblur';"  name="DiyContent"  id="DiyContent" cols="95%" rows="10"><%=Server.HTMLEncode(DiyContent)%></textarea>
<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('GetSlide_DiyContent')"  id="GetSlide_DiyContent">帮助</span>
</td>
	</tr>    
    <tr>
      <td >列出条数
      <input name="ListNumber" type="text" class="Ainput" id="ListNumber2"     value="<%= ListNumber %>" size="30">
      <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_ListNumber')"  id="Label_ListNumber">帮助</span>
      
      </td>
      <td >标题字数
        <input name="TitleLen"  type="text" class="Ainput"     value="<%= TitleLen %>"  size="30">	
        <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_TitleLen')"  id="Label_TitleLen">帮助</span> </td>
    </tr>
    <tr>
      <td colspan="2" align="center"  >
       <input name="SubmitBtn" class="ACT_btn" type="button"  onClick="InsertScriptFun()"  id="SubmitBtn"  value=" 确 定 ">    
      &nbsp;&nbsp; <input type="reset" class="ACT_btn" name="Submit" value="  重置  "></td>
    </tr>
    
    
    </table>
  
</form>
<script language="javascript" >

function SelectClass()
{

 if(document.myform.select1.value==0)	
	{
	document.all.ClassID.value=0
	}

 if(document.myform.select1.value==1 )	
	{
	document.all.ClassID.value=1
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
	}
	}
 

 function  
  SetDiyContent(oTextarea,strText){   
  oTextarea.focus();   
  document.selection.createRange().text+=strText;   
  oTextarea.blur();   
  }   
	
function OpenWindow(Url,Width,Height,WindowObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;');
	return ReturnStr;
}	


function OpenWindowAndSetValue(Url,Width,Height,WindowObj,SetObj)
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
			var DiyContent=document.myform.DiyContent.value;
  			var ClassID=document.myform.ClassID.value;
 			var ListNumber=document.myform.ListNumber.value;
 			var TitleLen=document.myform.TitleLen.value;
		    var IntroNumber=document.myform.IntroNumber.value;
  			var ModeID=document.myform.ModeID.value;
 			if  (ListNumber=='')  ListNumber=10;
			if  (TitleLen=='') TitleLen=30;
			if  (IntroNumber=='') IntroNumber=30;
			
			document.myform.LabelContent.value=	'{$GetSlide('+ClassID+'§'+ListNumber+'§'+TitleLen+'§'+ModeID+'§'+DiyContent+'§'+IntroNumber+')}';
			document.myform.SubmitBtn.value="正在提交数据,请稍等...";
			document.myform.SubmitBtn.disabled=true;
			document.myform.Submit.disabled=true;	
			document.myform.submit();
		}
</script> 
</body>
</html>
