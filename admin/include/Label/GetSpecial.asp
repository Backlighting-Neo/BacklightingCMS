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
Dim ListNumber,ContentLen,ColNumber,TitleCss,PicHeight,PicWidth,TypeSlide,ModeID,iftrue,ActF,DiyContent
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
 	ContentLen = 30
 	ActF=1
	LabelFlag=1
	TypeTitle = True
	pages = "新建幻灯片专题标签"
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
		LabelContent = Replace(Replace(LabelContent, "{$GetSpecial(", ""), ")}", "")
		LabelContent = Replace(LabelContent, """", "") 
		LabelContentArr = Split(LabelContent, "§")
 		ListNumber = LabelContentArr(0)'列出条数
		ContentLen = LabelContentArr(1)'链接目标
  		DiyContent = LabelContentArr(2)
  		pages = "修改新建幻灯片专题标签"
End IF
 %>
<form id="myform" name="myform" method="post" action="AddLabelSave.asp">
 <input type="hidden" name="LabelContent"> 
  <input type="hidden" name="Action" value="<%= Action %>"> 
  <input type="hidden" name="ID" value="<%= ID %>"> 
 <input type="hidden" name="FileUrl" value="GetSpecial.asp">  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
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
    
  <tr><td  colspan="2" >

<font color=red>内置标签</font> 
 <a href="#" onClick='SetDiyContent(DiyContent,"#ID")'>专题ID</a>&nbsp;
 <a href="#" onClick='SetDiyContent(DiyContent,"#Link")'>专题链接</a>&nbsp;
 <a href="#" onClick='SetDiyContent(DiyContent,"#Title")'>专题标题</a>&nbsp;
  <a href="#" onClick='SetDiyContent(DiyContent,"#Thumb")'>缩略图</a>&nbsp;
 <a href="#" onClick='SetDiyContent(DiyContent,"#Content")'>专题导读</a>&nbsp;
 <a href="#" onClick='SetDiyContent(DiyContent,"#Time")'>时间</a>&nbsp;
  <a href="#" onClick='SetDiyContent(DiyContent,"#Writer")'>责任编辑</a>&nbsp;
 <a href="#" onClick='SetDiyContent(DiyContent,"#Hits")'>点击数</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#New")'>New图标</a>&nbsp;
 <br>

<textarea onFocus="this.className='colorfocus';" onBlur="this.className='colorblur';"  name="DiyContent"  id="DiyContent" cols="95%" rows="10"><%=Server.HTMLEncode(DiyContent)%></textarea>
<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('GetSpecial_DiyContent')"  id="GetSpecial_DiyContent">帮助</span>
</td>
	</tr>    
    <tr>
      <td >列出条数
      <input name="ListNumber" type="text" class="Ainput" id="ListNumber2"     value="<%= ListNumber %>" size="30">
      <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_ListNumber')"  id="Label_ListNumber">帮助</span>      </td>
      <td >导读字数
        <input name="ContentLen"  type="text" class="Ainput"     value="<%= ContentLen %>"  size="30">	
        <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_ContentLen')"  id="Label_ContentLen">帮助</span> </td>
    </tr>
    <tr>
      <td colspan="2" align="center"  >
       <input name="SubmitBtn" class="ACT_btn" type="button"  onClick="InsertScriptFun()"  id="SubmitBtn"  value=" 确 定 ">    
      &nbsp;&nbsp; <input type="reset" class="ACT_btn" name="Submit" value="  重置  "></td>
    </tr>
    
    
    </table>
  
</form>
<script language="javascript" >

 
 

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
  			var ListNumber=document.myform.ListNumber.value;
  			var ContentLen=document.myform.ContentLen.value;
			if  (ContentLen=='')  ListNumber=30;
			if  (ListNumber=='')  ListNumber=10;
  			
			document.myform.LabelContent.value=	'{$GetSpecial('+ListNumber+'§'+ContentLen+'§'+DiyContent+')}';
			document.myform.SubmitBtn.value="正在提交数据,请稍等...";
			document.myform.SubmitBtn.disabled=true;
			document.myform.Submit.disabled=true;	
			document.myform.submit();
		}
</script> 
</body>
</html>
