<!--#include file="include.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>安装主题</title>
<link href="../../Images/style.css" rel="stylesheet" type="text/css">
<link href="Images/new.css" rel="stylesheet" type="text/css">
 </head>

<body> 
<%	Dim strXmlFile,fso
	ThemeID=actcms.s("ThemeID")
 	If actcms.s("act")="save" Then
		If  Trim(ThemeID)="" Then 
			Call Actcms.ActErr("请选择一个主题","","1")
		End If 
 	 	Set fso = CreateObject(ACTCMS.FsoName)
 		strXmlFile =Tpath & ACTCMS.ActCMS_Sys(8) & "\Sys_Act\Theme\ThemeInstallTemp\" &ThemeID& "\Install.Asp"
   		If fso.FileExists(strXmlFile)  Then
			actcms.actexe("Update Config_act Set ActCMS_Theme='"&ThemeID&"'  ")
			Call actcms.DelCahe("NowTheme")
			response.Redirect "ThemeInstallTemp/"&ThemeID&"/Install.asp?ThemeID="&ThemeID&"&Install="&actcms.s("Install")
  		Else 
			'actcms.actexe("Update Config_act Set ActCMS_Theme='"&ThemeID&"'  ")
			'Call actcms.DelCahe("NowTheme")
 			 Call Actcms.ActErr(ThemeID&"主题安装失败,没有安装程序","","1")
		End If 
 		response.end
  End If 
%>	

<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td class="bg_tr">您现在的位置：在线主题样式 &gt;&gt; 安装主题</td>
  </tr>
  <tr>
    <td>
	
	
	 
 <div class="md-head" >
<div class="zsj"></div>
<a href="Act.Theme.ACTList.asp" class="a10 wrap"  >在线安装主题</a>
<a href="Act.Theme.Reg.asp" class="a10 wrap"  >我的帐号</a>
 <a href="Index.asp" class="a10 wrap"  >主题管理</a>
<a href="Act.Theme.Check.asp" class="a10 wrap"  >查看主题的可用更新</a>
 <div>
 
</td>
  </tr>
</table>
  
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <form method="post" action="?act=save&theme=<%=ThemeID%>">
 

 
    
    <tr>
      <td height="25" align="right">主题名称&nbsp;&nbsp;</td>
      <td height="25"><%=ThemeID%>
 </td>
    </tr>
 
      	<input type="hidden" name="ThemeID" value="<%=ThemeID%>">

     <tr>
      <td height="25" align="right">安装说明&nbsp;&nbsp;</td>
      <td height="25">
	  <%
 	 	Set fso = CreateObject(ACTCMS.FsoName)
		strXmlFile =Tpath & ThemePath & request("ThemeID") & "/" & "theme.xml"
		If fso.FileExists(strXmlFile) Then
    		Set XML=ACTCMS.NoAppGetXMLFromFile(strXmlFile)
			If IsObject(XML) And XML.readyState=4 And XML.parseError.errorCode = 0 Then
				echo XML.documentElement.selectSingleNode("description").text
 				Set XML=Nothing
			End If
 		End If
%>
     </td>
    </tr> 
 
    <tr>
      <td height="25" align="right">标签重名覆盖&nbsp;&nbsp;</td>
      <td height="25"><label for="regss"><input type="checkbox" id="regss" name="regs" value="1"/> 
		<font color="red">* 如果主题有多模型.会同时创建(此操作不可恢复,建议先装一个测试系统来测试主题)</font></label>
     </td>
    </tr>
     

  
 <tr>
      <td   colspan="2"  align="center">
      <input type="submit" class="ACT_btn"  name=Submit1 value="  安装主题  " />
    </td>
    </tr>
 
 
  </table>	
</form>	

 
</body>
</html>
