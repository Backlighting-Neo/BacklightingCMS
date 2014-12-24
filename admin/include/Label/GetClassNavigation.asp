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
Dim Action,ID,LabelRS,LabelName,LabelContent,LabelFlag,LabelContentArr,Rs,pages,ModeID,DiyContent
Dim TitleCss,OpenType,NavWord,Navpic,NavType,ClassName,ColNumber,Division,NavHeight,sysdir,ThisCss,iftrue
sysdir=actcms.ActSys
Action = Request.QueryString("Action")
ID =  ChkNumeric(Request.QueryString("ID"))
Dim ActF,divid,divclass,ulid,ulclass,liid,liclass,Str,ACTIF,freetp,ATT
If Not ACTCMS.ACTCMS_QXYZ(0,"bqxt","") Then   Call Actcms.Alert("对不起，你没有操作权限！","")
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

IF Action = "Add" Then
	NavType = 0
	ModeID=  0
	ColNumber =1
	NavHeight =20
	ATT = 0
	ActF=1
	ClassName = "整站导航"
	pages = "新建网站总导航标签"
	ModeID =  ChkNumeric(Request.QueryString("ModeID"))
Else
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
		LabelContent = Replace(Replace(LabelContent, "{$GetClassNavigation(", ""), ")}", "")
		LabelContentArr = Split(LabelContent, "§")
		ModeID = LabelContentArr(0)
		OpenType =  LabelContentArr(1)
		ColNumber =  LabelContentArr(2)
		NavHeight =  LabelContentArr(3)
		TitleCss =  LabelContentArr(4)
		Division =  LabelContentArr(5)
		NavType = LabelContentArr(6)
		ActF= LabelContentArr(8)
		ThisCss= LabelContentArr(9)

		DiyContent=LabelContentArr(10)
		IF NavType = 0 Then 
			NavWord = LabelContentArr(7)
			Navpic = ""
		Else
			NavWord = ""
			Navpic = LabelContentArr(7)
		End IF
		pages = "修改网站总导航标签"
End IF
Function ReturnAllChannel(FolderID)
  Dim ChannelStr:ChannelStr = ""
      ChannelStr = "<select name=""select1"" onChange=""SelectClass();"">"
      ChannelStr = ChannelStr & "<option value=""0""> 整站导航  </option>"
	  if FolderID="888" then
	  ChannelStr = ChannelStr & "<option value=""888"" style=""color:red"" selected>当前频道通用</option>"
	  else
	  ChannelStr = ChannelStr & "<option value=""888"" style=""color:red"">当前频道通用</option>"
	  end if
	 
	 
	  if  InStr(FolderID, ",") > 0  then
	  ChannelStr = ChannelStr & "<option value=""973"" style=""color:blue"" selected>指定栏目</option>"
	  else
	  ChannelStr = ChannelStr & "<option value=""973"" style=""color:blue"">指定栏目</option>"
	  end if
	  
	 'ChannelStr = ChannelStr & "<option value=""973"" >指定栏目</option>"
		ChannelStr = ChannelStr & "<optgroup  label=""指定到模块"">"
		ChannelStr = ChannelStr & ReturnChannelOption(FolderID)
   ChannelStr = ChannelStr & "</Select>"
   ReturnAllChannel = ChannelStr
End Function
	Public Function ReturnChannelOption(SelectModeID)
	  Dim ChannelRS:Set ChannelRS=Server.CreateObject("ADODB.Recordset")
	  Dim ChannelStr:ChannelStr = ""
	   ChannelRS.Open "Select * From Mode_Act Where ModeStatus=0", Conn, 1, 1
	   If ChannelRS.EOF And ChannelRS.BOF Then
		  ChannelRS.Close:Set ChannelRS = Nothing:Exit Function
	  Else
		
	   Do While Not ChannelRS.EOF
		  If Cstr(ChannelRS("ModeID")) = Cstr(SelectModeID) Then
		  ChannelStr = ChannelStr & "<option selected value=" & ChannelRS("ModeID") & ">" & ChannelRS("ModeName") & "</option>"
		 Else
		   ChannelStr = ChannelStr & "<option value=" & ChannelRS("ModeID") & ">" & ChannelRS("ModeName") & "</option>"
		 End If
		ChannelRS.MoveNext
		Loop
	   ChannelRS.Close:Set ChannelRS = Nothing
	  End If
	   ReturnChannelOption = ChannelStr
	End Function
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
      <td width="50%" class="td_bg">标签名称
      <input name="LabelName"  type="text" class="Ainput" id="LabelName" value="<%= LabelName %>"></td>
      <td width="50%" class="td_bg">&nbsp;</td>
    </tr>
    <tr>
      <td class="td_bg">标签目录      
        <select name="LabelFlag" id="select">
          <option value="0">系统默认</option>
			 <%=AF.ACT_LabelFolder(CInt(LabelFlag))%>
        </select>&nbsp;&nbsp;<a href="ACT.LabelFolder.asp"><font color=red>新建存放目录</font></a></td>
	 <td width="50%" height="24"  class="td_bg">导航排列方式
	     <input type="radio" <% IF ColNumber = 989 Then Response.Write("Checked") %> id="Arrangement1" name="Arrangement" value="0" onClick="Arrangements(0);"><label for="Arrangement1">横排</label>
	     <input type="radio" <% IF ColNumber <> 989 Then Response.Write("Checked") %> id="Arrangement2" name="Arrangement" value="1" onClick="Arrangements(1);"><label for="Arrangement2">竖排</label>
	   <br><font color="red" >横排一般应用网站总导航,竖排应用栏目导航</font>
       
       <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_Arrangement')"  id="Label_Arrangement">帮助</span></td>
    </tr>


	 <tr>
      <td width="50%" >
	  输出模式
		 <select  style='width:40%' name="ActF" id="ActF" onChange="SetActF(this.options[this.selectedIndex].value);"> 
	 <option value="1" <% IF ActF = 1 Then Response.Write("selected") %>>普通模式</option>
  <option value="2" <% IF ActF = 2 Then Response.Write("selected") %>>代码模式</option>
  </select>   
		 <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_ActF')"  id="Label_ActF">帮助</span>
          </td>
      <td width="50%" >
	  当前栏目样式 <input name="ThisCss" type="text" class="Ainput" id="ThisCss" value="<%= ThisCss %>" >
       <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_ThisCss')"  id="Label_ThisCss">帮助</span>
	  </td>
    </tr>
  <tr id=ActFs ><td  colspan="2" >
  
<font color=red>内置标签</font> 
<a href="#" onClick='SetDiyContent(DiyContent,"#ClassName")'>栏目名称</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#ClassLink")'>栏目链接</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#Css")'>Css变量</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#AutoID")'>自增长ID</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#ModID")'>间歇ID</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#Path")'>系统路径</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#ClassSeo")'>栏目SEO标题</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#ClassPicUrl")'>栏目缩略图</a>&nbsp;
<a href="#" onClick='SetDiyContent(DiyContent,"#ClassPicFile")'>栏目缩略图地址</a>&nbsp;

<br />
<textarea  onfocus="this.className='colorfocus';" onBlur="this.className='colorblur';"  name="DiyContent"  cols="95%" rows="10"><%=DiyContent%></textarea>
 <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_DiyContent')"  id="Label_DiyContent">帮助</span> 
  </td>
	</tr>
 	
	<tr>
      <td class="td_bg">
	所属栏目   <input name="ModeID" type="text" class="Ainput" id="ModeID" value="<%= ModeID %>" readonly  disabled=true>
 	 <%= ReturnAllChannel(ModeID) %> <a href="#" onClick="SelectClass()">快速打开</a>
   <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_menuModeID')"  id="Label_menuModeID">帮助</span>   
     </td>
      <td class="td_bg">
      
      
       链接目标  
	<input name="OpenType" id="OpenType" type="text" class="Ainput" value="<%=OpenType%>" size="8">
 	  <%iftrue=false%>
	<select   name="OpenTypes"  onchange="document.myform.OpenType.value=this.value">
          <option value="_blank"   style="color:green"  <%If OpenType="_blank" Then response.write "selected":iftrue=true%>>新窗口打开</option>
          <option value="_parent" <%If OpenType="_parent" Then response.write "selected":iftrue=true%>>父窗口打开</option>
          <option value="_self" <%If OpenType="_self" Then response.write "selected":iftrue=true%>>本窗口打开</option>
          <option value="_top" <%If OpenType="_top" Then response.write "selected":iftrue=true%>>主窗口打开</option>
		<option value='' style="color:red"  <%If iftrue=false Then response.write "selected"%>>自定义</option>	
		</select>	  
 <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_OpenType')"  id="Label_OpenType">帮助</span>
      
      </td>
    </tr>
    <tr>
      <td class="td_bg">排列列数
      <input type="text" class="Ainput"    size="4"  <% IF ColNumber = 989 Then Response.Write " disabled='true'" %> value="<%= ColNumber %>" name="ColNumber"> &nbsp;高度
      <input name="NavHeight" type="text" class="Ainput"  id="NavHeight" value="<%= NavHeight %>"   size="4">
      <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_ColNumber')"  id="Label_ColNumber">帮助</span> 
      
      </td>
      <td class="td_bg">标题样式
	  <input name="TitleCss" type="text" class="Ainput" style="width:50%;" id="TitleCss" value="<%= TitleCss %>">
       <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_TitleCss')"  id="Label_TitleCss">帮助</span> 
       </td>
    </tr>
    <tr>
      <td colspan="2" class="td_bg">分隔图片
        <input name="Division"  type="text" class="Ainput" id="Division" style="width:61%;" value="<%= Division %>" readonly>
        <input class="ACT_btn" type="button" id="Division1" onClick="OpenWindowAndSetValue('../print/SelectPic.asp?CurrPath=<%=sysdir%>',500,320,window,document.myform.Division);" name="Submit3" value="选择图片...">
		 &nbsp;<span style="cursor:hand;color:green;" onClick="javascript:document.myform.Division.value='';" onMouseOver="this.style.color='red'" onMouseOut="this.style.color='green'">清除</span>
      <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_Division')"  id="Label_Division">帮助</span>    
         
         </td>
    </tr>
   
    <tr >
      <td class="td_bg">导航类型
        <input id="NavType1"  <% IF NavType = 0 Then Response.Write("Checked") %> type="radio" name="NavType" value="0" onClick="SetNavStatus(1);"><label for="NavType1">文字导航</label>
         <input id="NavType2" <% IF NavType = 1 Then Response.Write("Checked") %> type="radio" name="NavType" value="1" onClick="SetNavStatus(2);"><label for="NavType2">图片导航</label>
   <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_NavType')"  id="Label_NavType">帮助</span> 
         		</td>
    <td width="50%" height="24" class="td_bg"  id=SetNavStatus1 
	<% if NavType=1 then %>
	style="DISPLAY: none" 
	<% end if %>>
 <input name="NavWord" type="text" class="Ainput"  id="NavWord" style="width:70%;" value="<%= NavWord%>"> 
 支持HTML语法
  <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_NavWord')"  id="Label_NavWord">帮助</span>
 </td>
   <td width="50%" height="24" class="td_bg"  id=SetNavStatus2 
	<% if NavType=0 then %>
	style="DISPLAY: none"
	SetNavStatus1.style.display=''
	<% end if %> >
<input name="NavPic" type="text" class="Ainput"  id="NavPic" style="width:250;" value="<%= NavPic %>" readonly>
<input class="ACT_btn" type="button"  onClick="OpenWindowAndSetValue('../print/SelectPic.asp?CurrPath=<%=sysdir%>',500,320,window,document.myform.NavPic);" value="选择图片...">&nbsp;<span style="cursor:hand;color:green;" onClick="javascript:document.myform.NavPic.value='';" onMouseOver="this.style.color='red'" onMouseOut="this.style.color='green'">清除</span>
<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Label_NavPic')"  id="Label_NavPic">帮助</span>
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
function SelectClass()
{

 if(document.myform.select1.value==0)	
	{
	document.all.ModeID.value=0
	}
 if(document.myform.select1.value==1)	
	{
	document.all.ModeID.value=1
	}
 if(document.myform.select1.value==2)	
	{
	document.all.ModeID.value=2
	}
 if(document.myform.select1.value==888 )	
	{
	document.all.ModeID.value=888
	}

 if(document.myform.select1.value==973 )	
	{	
	var result = Selector(2, document.all.ModeID.value);
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
	document.all.ModeID.value = val;
	}
	}
	


function Arrangements(n){
			if (n==0){
document.myform.ColNumber.disabled=true
document.myform.ColNumber.value =989
			}
		  else{
document.myform.ColNumber.disabled=false
document.myform.ColNumber.value =1
		}
}

		function SetActF(Val)
		{
		 if(Val==1)	
			{
			 ActFs.style.display="none";
			  document.myform.OpenType.disabled=false;
			  document.myform.OpenTypes.disabled=false;
			  document.myform.ColNumber.disabled=false;
			  document.myform.NavHeight.disabled=false;
			  document.myform.NavType1.disabled=false;
			  document.myform.NavType2.disabled=false;
 			  document.myform.NavWord.disabled=false;
			  document.myform.NavPic.disabled=false;
			  document.myform.Division1.disabled=false;
			  document.myform.Arrangement1.disabled=false;
			  document.myform.Arrangement2.disabled=false;
			}
		 if(Val==2)	
			{
			 ActFs.style.display="";
			  document.myform.OpenType.disabled=true;
			  document.myform.OpenTypes.disabled=true;
			  document.myform.ColNumber.disabled=true;
			  document.myform.NavHeight.disabled=true;
			  document.myform.NavType1.disabled=true;
			  document.myform.NavType2.disabled=true;
 			  document.myform.NavWord.disabled=true;
			  document.myform.NavPic.disabled=true;
			  document.myform.Division1.disabled=true;
			  document.myform.Arrangement1.disabled=true;
			  document.myform.Arrangement2.disabled=true;
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
		function InsertScriptFun(Obj)
		{   if (document.myform.LabelName.value=='')
			 {
			  alert('请输入标签名称');
			  document.myform.LabelName.focus(); 
			  return false
			  }
			var NavType=1;
			var ModeID=document.myform.ModeID.value;
		
			var OpenType=document.myform.OpenType.value;
			var TitleCss=document.myform.TitleCss.value;
			var Division=document.myform.Division.value;
			
			var ColNumber=document.myform.ColNumber.value;
			var NavHeight=document.myform.NavHeight.value;
			for (var i=0;i<document.myform.NavType.length;i++){
				 var TCJ = document.myform.NavType[i];
				if (TCJ.checked==true)	   
					NavType = TCJ.value
				if  (NavType==0) Nav=''+document.myform.NavWord.value+''
				 else  Nav=''+document.myform.NavPic.value+'';
				}
			var ThisCss=document.myform.ThisCss.value;

			var DiyContent=document.myform.DiyContent.value;
			var ActF=document.myform.ActF.value;

 			if  (ColNumber=='') ColNumber=1;
			if (NavHeight=='') NavHeight=20
			document.myform.LabelContent.value=	'{$GetClassNavigation('+ModeID+'§'+OpenType+'§'+ColNumber+'§'+NavHeight+'§'+TitleCss+'§'+Division+'§'+NavType+'§'+Nav+'§'+ActF+'§'+ThisCss+'§'+DiyContent+')}';
			document.myform.SubmitBtn.value="正在提交数据,请稍等...";
			document.myform.SubmitBtn.disabled=true;
			document.myform.Submit.disabled=true;	
			document.myform.submit();
		}
</script><script language="javascript">SetActF(<%= ActF %>);</script>
</body>
</html>
