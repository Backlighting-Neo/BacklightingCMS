<!--#include file="../ACT.Function.asp"-->
<!--#include file="../include/ACT.F.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../Images/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/lhgcore.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/lhgdialog.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/Main.js"></script>
 <title>个性化定制</title>
</head>
<body>
<%
dim ActCMS_D,ModeID
ModeID = ChkNumeric(Request("ModeID"))
IF Request.QueryString("Action") = "ACTCMS" Then
Dim DIY_Article,i
	For I=0 To 25
		DIY_Article=DIY_Article& Replace(Replace(request.Form("ActCMS_D" & I &""),"§","") & "§",",","")
	Next
	Call  AF.ActCMS_DIY_F(ModeID,2,DIY_Article)
	Response.Redirect("?ModeID="&ModeID&"")
	response.end
Else
		If Trim(AF.ActCMS_DIY_F(ModeID,1,""))="1" Then 
 			ActCMS_D=Split("§0§0-1-0-1§0§actcms§0§§0§§0§§0§Simple§§§0§0§0§1§0§1§0§Class.htm§List.Htm§Content.Htm§0§","§")
		Else
			ActCMS_D=Split(AF.ActCMS_DIY_F(ModeID,1,""),"§")
		End If 
End If 
  %>
<form name="Set_Act_Type_DiY" method="post" action="?Action=ACTCMS&ModeID=<%=ModeID%>">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
    <tr>
      <td colspan="4" class="bg_tr">
	  您现在的位置：后台管理 >> <a href="ACT.MX.asp">模型列表</a> >> 自定义显示</td>
    </tr>
  

    <tr>
      <td height="38" align="right">完整标题 ：</td>
      <td>	  <input name="ActCMS_D0"   type="text"  class="Ainput"   title="在这里填写默认值" value="<%= ActCMS_D(0) %>" size="30">

        默认值 
		<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ActDiY_001')"  id="ActDiY_001">帮助</span> </td>
      <td colspan="2">
<input <% IF  ActCMS_D(1) = "0" Then Response.Write "Checked" %> id="ActCMS_D(1)1" type="radio" name="ActCMS_D1" value="0">
<label for="ActCMS_D(1)1"><font color="green">正常 &nbsp;</font></label>
<input <% IF  ActCMS_D(1) = "1" Then Response.Write "Checked" %> id="ActCMS_D(1)2" type="radio" name="ActCMS_D1" value="1">
<label for="ActCMS_D(1)2"><font color="red">关闭 &nbsp;</font></label>

</td>
    </tr>
    <tr>
      <td height="38" align="right">文章属性 ：</td>
      <td>	  <input name="ActCMS_D2"   type="text"  class="Ainput"   title="在这里填写默认值" value="<%= ActCMS_D(2) %>" size="30">

        默认值 
		<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ActDiY_002')"  id="ActDiY_002">帮助</span></td>
      <td colspan="2">
		
<input  <% IF  ActCMS_D(3) = "0" Then Response.Write "Checked" %>  type="radio"  id="ActCMS_D(3)1" name="ActCMS_D3" value="0">
<label for="ActCMS_D(3)1"><font color="green">正常 &nbsp;</font></label>
<input <% IF  ActCMS_D(3) = "1" Then Response.Write "Checked" %>   id="ActCMS_D(3)2" type="radio" name="ActCMS_D3" value="1">
<label for="ActCMS_D(3)2"><font color="red">关闭 &nbsp;</font></label>

</td>
    </tr>
    <tr>
      <td height="38" align="right">关键字 ：</td>
      <td>	  <input name="ActCMS_D4"   type="text"  class="Ainput"   title="在这里填写默认值" value="<%= ActCMS_D(4) %>" size="30">

        默认值
		<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ActDiY_003')"  id="ActDiY_003">帮助</span></td>
      <td colspan="2">
		
<input  <% IF  ActCMS_D(5) = "0" Then Response.Write "Checked" %>  type="radio"  id="ActCMS_D(5)1" name="ActCMS_D5" value="0">
<label for="ActCMS_D(5)1"><font color="green">正常 &nbsp;</font></label>
<input <% IF  ActCMS_D(5) = "1" Then Response.Write "Checked" %>   id="ActCMS_D(5)2" type="radio" name="ActCMS_D5" value="1">
<label for="ActCMS_D(5)2"><font color="red">关闭 &nbsp;</font></label>		</td>
    </tr>
    <tr>
      <td height="38" align="right">文章作者 ：</td>
      <td>	  <input name="ActCMS_D6"   type="text"  class="Ainput"   title="在这里填写默认值" value="<%= ActCMS_D(6) %>" size="30">

        默认值
		<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ActDiY_004')"  id="ActDiY_004">帮助</span></td>
      <td colspan="2">
<input  <% IF  ActCMS_D(7) = "0" Then Response.Write "Checked" %>  type="radio"  id="ActCMS_D(7)1" name="ActCMS_D7" value="0">
<label for="ActCMS_D(7)1"><font color="green">正常 &nbsp;</font></label>
<input <% IF  ActCMS_D(7) = "1" Then Response.Write "Checked" %>   id="ActCMS_D(7)2" type="radio" name="ActCMS_D7" value="1">
<label for="ActCMS_D(7)2"><font color="red">关闭 &nbsp;</font></label>

		</td>
    </tr>
    <tr>
      <td height="38" align="right">文章来源 ：</td>
      <td>	  <input name="ActCMS_D8"   type="text"  class="Ainput"   title="在这里填写默认值" value="<%= ActCMS_D(8) %>" size="30">

        默认值
		<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ActDiY_005')"  id="ActDiY_005">帮助</span></td>
      <td colspan="2">
<input  <% IF  ActCMS_D(9) = "0" Then Response.Write "Checked" %>  type="radio"  id="ActCMS_D(9)1" name="ActCMS_D9" value="0">
<label for="ActCMS_D(9)1"><font color="green">正常 &nbsp;</font></label>
<input <% IF  ActCMS_D(9) = "1" Then Response.Write "Checked" %>   id="ActCMS_D(9)2" type="radio" name="ActCMS_D9" value="1">
<label for="ActCMS_D(9)2"><font color="red">关闭 &nbsp;</font></label>		
		</td>
    </tr>
    <tr>
      <td height="38" align="right">文章导读 ：</td>
      <td>	  <input name="ActCMS_D10"   type="text"  class="Ainput"   title="在这里填写默认值" value="<%= ActCMS_D(10) %>" size="30">

        默认值
		<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ActDiY_006')"  id="ActDiY_006">帮助</span></td>
      <td colspan="2">
		
<input  <% IF  ActCMS_D(11) = "0" Then Response.Write "Checked" %>  type="radio"  id="ActCMS_D(11)1" name="ActCMS_D11" value="0">
<label for="ActCMS_D(11)1"><font color="green">正常 &nbsp;</font></label>
<input <% IF  ActCMS_D(11) = "1" Then Response.Write "Checked" %>   id="ActCMS_D(11)2" type="radio" name="ActCMS_D11" value="1">
<label for="ActCMS_D(11)2"><font color="red">关闭 &nbsp;</font></label>			</td>
    </tr>
   

    <tr>
      <td height="38" align="right">fckeditor显示菜单 ：</td>
      <td colspan="3">	  <input name="ActCMS_D12" id="ActCMS_D12"   type="text"  class="Ainput"   title="在这里填写默认值" value="<%= ActCMS_D(12) %>" size="30">
<select name="select"   onchange="document.Set_Act_Type_DiY.ActCMS_D12.value=this.value">
          <option selected>-- 请选择 --</option>
          <option value="UserMode">简洁</option>
          <option value="Simple">超简洁</option>
          <option value="Default">默认</option>
        </select><span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ActDiY_12')"  id="ActDiY_12">帮助</span> &nbsp; &nbsp;<a href="http://www.actcms.com/fckmenu.html"  target="_blank"><font color="green" >生成自己的fckeditor菜单</font></a></td>
  
 
    </tr>
   
     
    <tr>
      <td height="38" align="right">系统默认内容 ：</td>
      <td >	  <input name="ActCMS_D13"   type="text"  class="Ainput"   title="在这里填写默认值" value="<%= ActCMS_D(13) %>" size="30">      
		  <input class="ACT_btn" type="button"  onClick="OpenWindowAndSetValue('../include/print/SelectPic.asp?CurrPath=<%=ACTCMS.ActCMS_Sys(3)%>',500,320,window,document.Set_Act_Type_DiY.ActCMS_D13);" value="选择默认内容文件..."> 
<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ActDiY_007')"  id="ActDiY_007">帮助</span>
</td>

    <td colspan="2">
		
<input  <% IF  ActCMS_D(25) = "0" Then Response.Write "Checked" %>  type="radio"  id="ActCMS_D251"  name="ActCMS_D25" value="0">
<label for="ActCMS_D251"><font color="green">正常 &nbsp;</font></label>
<input <% IF  ActCMS_D(25) = "1" Then Response.Write "Checked" %>    type="radio"  id="ActCMS_D252" name="ActCMS_D25" value="1">
<label for="ActCMS_D252"><font color="red">关闭 &nbsp;</font></label>			</td>


    </tr>
 
	
	

    <tr>
      <td height="38" align="right">阅读权限 ：</td>
      <td>
	  <input name="ActCMS_D14"   type="text"  class="Ainput"   title="在这里填写默认值" value="<%= ActCMS_D(14) %>" size="30">
        默认值
		<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ActDiY_008')"  id="ActDiY_008">帮助</span></td>
      <td colspan="2">
		<input  <% IF  ActCMS_D(15) = "0" Then Response.Write "Checked" %>  type="radio"  id="ActCMS_D(15)1" name="ActCMS_D15" value="0">
		<label for="ActCMS_D(15)1"><font color="green">正常 &nbsp;</font></label>
		<input <% IF  ActCMS_D(15) = "1" Then Response.Write "Checked" %>   id="ActCMS_D(15)2" type="radio" name="ActCMS_D15" value="1">
		<label for="ActCMS_D(15)2"><font color="red">关闭 &nbsp;</font></label>		
		</td>
    </tr>

	
    <tr>
      <td height="38" align="right">上传缩略图是否插入到内容 ：</td>
      <td>
       
      <input  <% IF  ActCMS_D(16) = "1" Then Response.Write "Checked" %>  type="radio"  id="ActCMS_D(16)1" name="ActCMS_D16" value="1">
		<label for="ActCMS_D(16)1"><font color="green">正常 &nbsp;</font></label>
		<input <% IF  ActCMS_D(16) = "0" Then Response.Write "Checked" %>   id="ActCMS_D(16)2" type="radio" name="ActCMS_D16" value="0">
		<label for="ActCMS_D(16)2"><font color="red">关闭 &nbsp;</font></label>	
      
      
      
         </td>
      <td colspan="2">
		 
		</td>
    </tr>

    <tr>
      <td height="38" align="right">生成文件名 ：</td>
      <td> 
		<input  <% IF  ActCMS_D(18) = "0" Then Response.Write "Checked" %>  type="radio"  id="ActCMS_D(18)1" name="ActCMS_D18" value="0">
		<label for="ActCMS_D(18)1"><font color="green">正常 &nbsp;</font></label>
		<input <% IF  ActCMS_D(18) = "1" Then Response.Write "Checked" %>   id="ActCMS_D(18)2" type="radio" name="ActCMS_D18" value="1">
		<label for="ActCMS_D(18)2"><font color="red">关闭 &nbsp;</font></label>
		
		<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ActDiY_010')"  id="ActDiY_010">帮助</span></td>
      <td colspan="2">
		</td>
    </tr>

    <tr>
      <td height="38" align="right">除去链接 ：</td>
      <td>
	  <input name="ActCMS_D19"   type="text"  class="Ainput"   title="在这里填写默认值" value="<%= ActCMS_D(19) %>" size="30">
        默认值
		<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ActDiY_011')"  id="ActDiY_011">帮助</span></td>
      <td colspan="2">
		<input  <% IF  ActCMS_D(20) = "0" Then Response.Write "Checked" %>  type="radio"  id="ActCMS_D(20)1" name="ActCMS_D20" value="0">
		<label for="ActCMS_D(20)1"><font color="green">正常 &nbsp;</font></label>
		<input <% IF  ActCMS_D(20) = "1" Then Response.Write "Checked" %>   id="ActCMS_D(20)2" type="radio" name="ActCMS_D20" value="1">
		<label for="ActCMS_D(20)2"><font color="red">关闭 &nbsp;</font></label>		
		</td>
    </tr>

    <tr>
      <td height="38" align="right">上传文件 ：</td>
      <td> 
		<input  <% IF  ActCMS_D(21) = "0" Then Response.Write "Checked" %>  type="radio"  id="ActCMS_D(21)1" name="ActCMS_D21" value="0">
		<label for="ActCMS_D(21)1"><font color="green">正常 &nbsp;</font></label>
		<input <% IF  ActCMS_D(21) = "1" Then Response.Write "Checked" %>   id="ActCMS_D(21)2" type="radio" name="ActCMS_D21" value="1">
		<label for="ActCMS_D(21)2"><font color="red">关闭 &nbsp;</font></label>
		
		<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ActDiY_021')"  id="ActDiY_021">帮助</span></td>
      <td colspan="2">
		</td>
    </tr>




    <tr>
      <td height="38" align="right">栏目页模板预设值 ：</td>
      <td colspan="3">	  <input name="ActCMS_D22"   type="text"  class="Ainput"   title="在这里填写默认值" value="<%= ActCMS_D(22) %>" size="30">      
		  <input class="ACT_btn" type="button"  onClick="OpenWindowAndSetValue('../include/print/SelectPic.asp?CurrPath=<%=ACTCMS.ActCMS_Sys(3)%>',500,320,window,document.Set_Act_Type_DiY.ActCMS_D22);" value="选择默认模板..."> 
<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ActDiY_022')"  id="ActDiY_022">帮助</span>只在添加栏目的时候有效
</td>
    </tr>


    <tr>
      <td height="38" align="right">列表页模板预设值 ：</td>
      <td colspan="3">	  <input name="ActCMS_D23"   type="text"  class="Ainput"   title="在这里填写默认值" value="<%= ActCMS_D(23) %>" size="30">      
		  <input class="ACT_btn" type="button"  onClick="OpenWindowAndSetValue('../include/print/SelectPic.asp?CurrPath=<%=ACTCMS.ActCMS_Sys(3)%>',500,320,window,document.Set_Act_Type_DiY.ActCMS_D23);" value="选择默认模板..."> 
<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ActDiY_023')"  id="ActDiY_023">帮助</span>只在添加栏目的时候有效
</td>
    </tr>

	

	    <tr>
      <td height="38" align="right">内容页模板预设值 ：</td>
      <td colspan="3">	  <input name="ActCMS_D24"   type="text"  class="Ainput"   title="在这里填写默认值" value="<%= ActCMS_D(24) %>" size="30">      
		  <input class="ACT_btn" type="button"  onClick="OpenWindowAndSetValue('../include/print/SelectPic.asp?CurrPath=<%=ACTCMS.ActCMS_Sys(3)%>',500,320,window,document.Set_Act_Type_DiY.ActCMS_D24);" value="选择默认模板..."> 
<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ActDiY_024')"  id="ActDiY_024">帮助</span>只在添加栏目的时候有效
</td>
    </tr>

	
	<tr>
      <td height="38" colspan="4" align="center">
        <input type="submit" class="act_btn" name="Submit2" value="  保存  ">
      </td>
    </tr>
  </table>
  <SCRIPT LANGUAGE="JavaScript">
  <!--
	function OpenWindowAndSetValue(Url,Width,Height,WindowObj,SetObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:yes;help:no;scroll:yes;');
	if (ReturnStr!='') SetObj.value=ReturnStr;
	return ReturnStr;
}	
  //-->
  </SCRIPT>
  <p>&nbsp;</p>
</form></body>
</html>
