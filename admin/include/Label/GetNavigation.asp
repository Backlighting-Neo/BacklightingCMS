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
Dim Action,ID,LabelRS,LabelName,LabelContent,LabelFlag,LabelContentArr,Rs,pages
Dim TitleCss,OpenType,NavWord,Navpic,NavType,sysdir,TypeModeName,iftrue
If Not ACTCMS.ACTCMS_QXYZ(0,"bqxt","") Then   Call Actcms.Alert("对不起，你没有操作权限！","")
sysdir=actcms.ActSys&"upfiles"
Action = Request.QueryString("Action")
ID =  ChkNumeric(Request.QueryString("ID"))
IF Action = "Add" Then
	NavType = 0
	TypeModeName=1
	pages = "新建网站位置导航标签"
Else
	  ConnectionDatabase	
	  Set LabelRS = Server.CreateObject("Adodb.Recordset")
	  LabelRS.Open "Select LabelName,LabelContent,LabelFlag From Label_Act Where ID=" & ID & "", Conn, 1, 1
	  If LabelRS.EOF And LabelRS.BOF Then
		 LabelRS.Close
		 Conn.Close:Set Conn = Nothing
		 Set LabelRS = Nothing
		 Response.Write "参数传递出错!":Response.End
	  End If
		LabelName = Replace(Replace(LabelRS("LabelName"), "{ACTCMS_", ""), "}", "")
		LabelContent = LabelRS("LabelContent")
		LabelFlag = Clng(LabelRS("LabelFlag"))
		LabelRS.Close:Set LabelRS = Nothing
		LabelContent = Replace(Replace(LabelContent, "{$GetNavigation(", ""), ")}", "")
		LabelContent = Replace(LabelContent, """", "") 
		LabelContentArr = Split(LabelContent, "§")
		TitleCss = LabelContentArr(0)
		OpenType =  LabelContentArr(1)
		NavType =  LabelContentArr(2)
		TypeModeName=  LabelContentArr(4)
		IF NavType = 0 Then 
			NavWord = LabelContentArr(3)
			Navpic = ""
		Else
			NavWord = ""
			Navpic = LabelContentArr(3)
		End IF
		pages = "修改网站位置导航标签"
End IF
%>
<form id="myform" name="myform" method="post" action="AddLabelSave.asp">
 <input type="hidden" name="LabelContent"> 
  <input type="hidden" name="LabelType" value="1">
  <input type="hidden" name="Action" value="<%= Action %>"> 
  <input type="hidden" name="ID" value="<%= ID %>"> 
 <input type="hidden" name="FileUrl" value="GetNavigation.asp"> 
  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
    <tr>
      <td colspan="2" align="center" class="bg_tr"><%= pages %>&nbsp;</td>
    </tr>
    <tr>
      <td width="50%" >标签名称
      <input name="LabelName"  type="text" class="Ainput"  id="LabelName" value="<%= LabelName %>"></td>
      <td width="50%" >&nbsp;</td>
    </tr>
    <tr>
      <td >标签目录      
        <select name="LabelFlag" id="select">
          <option value="0">系统默认</option>
			 <%=AF.ACT_LabelFolder(CInt(LabelFlag))%>
        </select>&nbsp;&nbsp;<a href="ACT.LabelFolder.asp"><font color=red><b><b>新建存放目录</b></b></font></a></td>
	 <td width="50%" height="24"  ><font color=green>标签存放目录,日后方便管理标签</font></td>
    </tr>

    <tr>
      <td >
	  文字样式
	  <input name="TitleCss" type="text" class="Ainput"  style="width:70%;" id="TitleCss" value="<%= TitleCss %>">
      <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_TitleCss')"  id="Label_TitleCss">帮助</span></td>
      <td >链接目标  
	<input name="OpenType" id="OpenType" type="text" class="Ainput" value="<%=OpenType%>" size="8">

	  <%iftrue=false%>
	<select   name="OpenTypes"  onchange="document.myform.OpenType.value=this.value">
          <option value="_blank"   style="color:green"  <%If OpenType="_blank" Then response.write "selected":iftrue=true%>>新窗口打开</option>
          <option value="_parent" <%If OpenType="_parent" Then response.write "selected":iftrue=true%>>父窗口打开</option>
          <option value="_self" <%If OpenType="_self" Then response.write "selected":iftrue=true%>>本窗口打开</option>
          <option value="_top" <%If OpenType="_top" Then response.write "selected":iftrue=true%>>主窗口打开</option>
		<option value='' style="color:red"  <%If iftrue=false Then response.write "selected"%>>自定义</option>	
		</select>	  
	 <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_OpenType')"  id="Label_OpenType">帮助</span> </td>
    </tr>
   
    <tr >
      <td >导航类型
        
      <label for="NavType1">  <input   id="NavType1"  <% IF NavType = 0 Then Response.Write("Checked") %> type="radio" name="NavType" value="0" onClick="SetNavStatus(1);">
        文字导航</label>
       <label for="NavType2">  <input   id="NavType2" <% IF NavType = 1 Then Response.Write("Checked") %> type="radio" name="NavType" value="1" onClick="SetNavStatus(2);">
        图片导航</label>	
        <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_NavType')"  id="Label_NavType">帮助</span>	</td>
    <td width="50%" height="24"   id=SetNavStatus1 
	<% if NavType=1 then %>
	style="DISPLAY: none" 
	<% end if %>>
 <input name="NavWord" type="text" class="Ainput"   id="NavWord" style="width:70%;" value="<%= server.HTMLEncode(NavWord) %>"> 
 支持HTML语法
  <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_NavWord')"  id="Label_NavWord">帮助</span>
 </td>
   <td width="50%" height="24"   id=SetNavStatus2 
	<% if NavType=0 then %>
	style="DISPLAY: none"
	SetNavStatus1.style.display=''
	<% end if %> >
<input name="NavPic" type="text" class="Ainput"   id="NavPic" style="width:250;" value="<%= NavPic %>" readonly>
<input class="ACT_btn" type="button"  onClick="OpenWindowAndSetValue('../print/SelectPic.asp?CurrPath=<%=sysdir%>',500,320,window,document.myform.NavPic);" value="选择图片...">&nbsp;<span style="cursor:hand;color:green;" onClick="javascript:document.myform.NavPic.value='';" onMouseOver="this.style.color='red'" onMouseOut="this.style.color='green'">清除</span>
<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_NavPic')"  id="Label_NavPic">帮助</span>

</td>  
    </tr>


    <tr>
      <td >是否显示模型名称
	   <label for="TypeModeName1">  <input   id="TypeModeName1"  <% IF TypeModeName = 0 Then Response.Write("Checked") %> type="radio" name="TypeModeName" value="0">
        显示</label>
       <label for="TypeModeName2">  <input   id="TypeModeName2" <% IF TypeModeName = 1 Then Response.Write("Checked") %> type="radio" name="TypeModeName" value="1">
        不显示</label>
     <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_TypeModeName')"  id="Label_TypeModeName">帮助</span> </td>
      <td ></td>
    </tr>
   

    <tr>
      <td colspan="2" align="center" >
	   <input name="SubmitBtn" class="ACT_btn" type="button"  onClick="InsertScriptFun()"  id="SubmitBtn"  value=" 确 定 ">	
	    &nbsp;&nbsp; <input type="reset" class="ACT_btn" name="Submit" value="  重置  ">  </td>
    </tr>
    </table>
  
</form>
<script language="javascript" >
	
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

		function InsertScriptFun(Obj)
		{   if (document.myform.LabelName.value=='')
			 {
			  alert('请输入标签名称');
			  document.myform.LabelName.focus(); 
			  return false
			  }
			var NavType=1;
		var OpenType=document.myform.OpenType.value;
			var TitleCss=document.myform.TitleCss.value;
			for (var i=0;i<document.myform.NavType.length;i++){
				 var TCJ = document.myform.NavType[i];
				if (TCJ.checked==true)	   
					NavType = TCJ.value
				if  (NavType==0) Nav=document.myform.NavWord.value
				 else  Nav=document.myform.NavPic.value;
				}

			for (var i=0;i<document.myform.TypeModeName.length;i++){
			 var TCJ = document.myform.TypeModeName[i];
			if (TCJ.checked==true)	   
				TypeModeName = TCJ.value
			}


			document.myform.LabelContent.value=	'{$GetNavigation('+TitleCss+'§'+OpenType+'§'+NavType+'§'+Nav+'§'+TypeModeName+')}';
			document.myform.SubmitBtn.value="正在提交数据,请稍等...";
			document.myform.SubmitBtn.disabled=true;
			document.myform.Submit.disabled=true;	
			document.myform.submit();
		}
</script>
<%Call CloseConn%>
</body>
</html>
