<!--#include file="include.asp"-->
  <%
 SelectedTheme=Request.QueryString("theme")
ThemeName=Request.QueryString("themename")
 If ThemeName = "" Then ThemeName = SelectedTheme
 echo ThemeID
 %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>插件安装</title>
 <link href="../../Images/style.css" rel="stylesheet" type="text/css">
 <link href="Images/css.css" rel="stylesheet" type="text/css">
 </head>

<body> 
 <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td class="bg_tr">您现在的位置：插件样式修改 &gt;&gt; 浏览</td>
  </tr>
  <tr>
    <td>
	
	
  <div class="md-head" >
<div class="zsj"></div>
<a href="Act.Theme.ACTList.asp" class="a10 wrap"  >在线安装插件</a>
<a href="Act.Theme.Reg.asp" class="a10 wrap"  >我的帐号</a>
<a href="Index.asp" class="a10 wrap cur wid"  >插件管理</a>
<a href="Act.Theme.Check.asp" class="a10 wrap"  >查看插件的可用更新</a>
 </div>

 <div class="top1"><span style=" margin-left:10px" ><strong>提示：</strong>修改 ID 为<%=SelectedTheme%> 的插件的信息文档.&nbsp;&nbsp; Theme.xml 文件, 该文件将位于插件目录内
</span></div>
</td>
  </tr>
</table>
  
 
 
 <%
	echo "<p id=""loading"">正在载入插件信息, 请稍候... 如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
	Response.Flush
 		Dim fso:Set fso = CreateObject(ACTCMS.FsoName)
  		ThemeSource_Name=Empty:ThemeSource_Url=Empty:ThemeID=Empty:ThemeName=Empty:ThemeURL=Empty:ThemeNote=Empty:ThemeModified=Empty
 		If fso.FileExists(Tpath & ThemePath & SelectedTheme & "/" & "theme.xml") Then
   			Set XML=ACTCMS.NoAppGetXMLFromFile(Tpath & ThemePath  & SelectedTheme & "/" & "theme.xml")
			If IsObject(XML) And XML.readyState=4 Then
 				If XML.parseError.errorCode = 0 Then 
 					 ThemeAuthor_Name=GetNode("author/name"):ThemeAuthor_Url=GetNode("author/url"):ThemeSource_Name=GetNode("source/name")
					 ThemeSource_Url=GetNode("source/url"):ThemeID=GetNode("id"):ThemeName=GetNode("name"):ThemeURL=GetNode("url")
					 ThemePubDate=GetNode("pubdate"):ThemeModified=GetNode("modified"):ThemeNote=GetNode("note"):ThemeAdapted=GetNode("adapted")
					 If ThemeAuthor_Name=Empty Then
						ThemeAuthor_Name=ThemeSource_Name:ThemeAuthor_Url=ThemeSource_Url
					 End If
					ThemeDescription=GetNode("description")	
 					If ThemeModified=Empty Then ThemeModified=ThemePubDate
 					ThemeNote=ACTCMS.CloseHtml(ThemeNote)
 				End If 
			Set XML=Nothing
 		Else
 			ThemeSource_Name="unknown":ThemeSource_Url=Empty:ThemeID=f1.name
			ThemeName=f1.name:ThemeURL=Empty:ThemeNote="unknown":ThemeModified="unknown"
 	  End If 
   	  End If
 	If fso.FileExists(Tpath & ThemePath& SelectedTheme & "/" & "preview.png") Then
		ThemeScreenShot="../../../" &ThemePath&SelectedTheme & "/" & "preview.png"
	Else
		ThemeScreenShot="Images/nopreview.jpg"
	End If
   	If fso.FileExists(Tpath & ThemePath & ThemeID & "/" & "verchk.xml") Then
		echo "<p><a class=""notice"" href=""Act.Theme.Install.Asp?act=update&amp;url=" & Server.URLEncode(Update_URL & ThemeID) & """ title=""升级插件""><b >发现该插件的新版本!</b></a></p><br />"&vbCrLf
	ElseIf fso.FileExists(Tpath & ThemePath & ThemeID & "/" & "error.log") Then
		echo "<p><b class=""notice"">该插件不支持在线更新.</b></p><br />"&vbCrLf
	End If

 
 

%>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <form method="post" action="Act.Theme.ThemeInstall.asp?A=PlugInOpen">
 


    <tr>
      <td width="20%" height="35" align="right">插件ID&nbsp;&nbsp;</td>
      <td width="46%" height="35"><% 	
	  If UCase(ThemeID)<>UCase(SelectedTheme) Then
		echo "该插件ID错误, 请 <a href=""Act.Theme.Edit.asp?theme=" & Server.URLEncode(SelectedTheme) & """ title=""编辑插件信息""><font color=""red""><b>[重新编辑插件信息]</b></font></a>"&vbCrLf
	Else
		echo  ThemeID &vbCrLf
	End If
	%>
      (插件ID应为插件文件夹名称, 由编辑器自动完成填写, 不可修改.) </td>
      <td width="34%" rowspan="9">
	    <a href="<%= ThemeScreenShot %>" target="_blank"><img src="<%= ThemeScreenShot %>" border="0" title="<%= ThemeName %>"  /></a> </td>
    </tr>
  
  
  
    <tr>
      <td height="35" align="right">插件名称&nbsp;&nbsp;</td>
      <td height="35"><%=SelectedTheme%> </td>
    </tr>
  
  
  
    <%If ThemeURL<>Empty Then %>
    <tr>
      <td height="35" align="right">发布地址&nbsp;&nbsp;</td>
      <td height="35"><%="<a href=""" & ThemeURL & """ target=""_blank"" title=""插件的发布地址"">" & ThemeURL & "</a>"%>        </td>
    </tr>
    <%End If %>
  
 
 
 

    <tr>
      <td height="35" align="right">插件作者&nbsp;&nbsp;</td>
      <td height="35"><%
	If ThemeAuthor_Url=Empty Then
		echo "" & ThemeAuthor_Name & "</p>"&vbCrLf
	Else
		echo "<a href=""" & ThemeAuthor_Url & """ target=""_blank"" title=""作者主页"">" & ThemeAuthor_Name & "</a>"&vbCrLf
	End If
%>  </td>
    </tr>
 
 
 
 
 
 
 
     <tr>
      <td height="35" align="right">作者邮箱&nbsp;&nbsp;</td>
      <td height="35"><%
	If ThemeAuthor_Email<>Empty Then echo "<a href=""mailto:" & ThemeAuthor_Email & """ title=""作者邮箱"">" & ThemeAuthor_Email &vbCrLf
%>  </td>
    </tr>
 
 <tr>
      <td height="35" align="right">发布日期&nbsp;&nbsp;</td>
      <td height="35"><%=ThemePubDate%>  </td>
    </tr>
 
 <tr>
      <td height="35" align="right">插件简介&nbsp;&nbsp;</td>
      <td height="35"><%=ThemeNote%>  </td>
    </tr>
 
 <tr>
      <td height="35" align="right">适用于&nbsp;&nbsp;</td>
      <td height="35"><%=ThemeAdapted%>  </td>
    </tr>
 
 <tr>
      <td height="35" align="right">插件版本&nbsp;&nbsp;</td>
      <td height="35"><%=ThemeVersion%>  </td>
    </tr>
 
 <tr>
      <td height="35" align="right">修正日期&nbsp;&nbsp;</td>
      <td height="35"><%=ThemeModified%>  </td>
      <td>&nbsp;</td>
 </tr>
 
 <tr>
      <td height="35" align="right">插件源作者&nbsp;&nbsp;</td>
      <td height="35"><%	If ThemeSource_Name<>Empty Then
		If ThemeSource_Url=Empty Then
			echo " " & ThemeSource_Name & vbCrLf
		Else
			echo " <a href=""" & ThemeSource_Url & """ target=""_blank"" title=""源作者主页"">" & ThemeSource_Name & "</a>"&vbCrLf
		End If
		If ThemeSource_Email<>Empty Then echo "<p><b>源作者邮箱:</b> <a href=""mailto:" & ThemeSource_Email & """ title=""源作者邮箱"">" & ThemeSource_Email & "</a></p>"&vbCrLf
	End If%>  </td>
      <td>&nbsp;</td>
 </tr>
 

 <tr>
      <td height="35" align="right">编辑插件信息&nbsp;&nbsp;</td>
      <td height="35"><%="<a href=""Act.Theme.Edit.asp?theme=" & Server.URLEncode(ThemeID) & """ title=""编辑插件信息"">[编辑插件信息]</a>:</b> 此功能可用于生成或编辑该插件的信息文档 Theme.xml"%>  </td>
      <td>&nbsp;</td>
 </tr>
 
  	
 <tr>
      <td height="35" align="right">升级修复插件&nbsp;&nbsp;</td>
      <td height="35"><%="<a href=""Act.Theme.Install.Asp?act=update&amp;url=" & Server.URLEncode(Update_URL & ThemeID) & """ title=""升级修复插件"">[升级修复插件]</a>"%>  </td>
      <td>&nbsp;</td>
 </tr>
 
  	
 <tr>
      <td height="35" align="right">删除插件&nbsp;&nbsp;</td>
      <td height="35"><%="<a href=""Index.asp?act=themedel&amp;theme=" & Server.URLEncode(SelectedTheme) & "&amp;themename=" & Server.URLEncode(ThemeName) & """ title=""删除此插件"" onclick=""return window.confirm('您将删除此插件的所有文件, 确定吗?');"">[删除插件]</a>"%>  </td>
      <td>&nbsp;</td>
 </tr>
 
  	
 <tr>
      <td height="35" align="right">详细说明&nbsp;&nbsp;</td>
      <td height="35"><%= ThemeDescription & " <br />"%>  </td>
      <td>&nbsp;</td>
 </tr>
 
  	<input type="hidden" name="ThemeID" value="<%=ThemeID%>">
 <tr>
      <td   colspan="3"  align="center">
      
	<%
	   	echo "<p><input type=""submit"" class=""ACT_btn"" value="" 应用此插件 ""   /> <input onclick=""self.location.href='Index.asp';"" type=""button"" class=""ACT_btn"" value="" 返回插件管理 ""   /></p>"&vbCrLf
  	echo "</form>"&vbCrLf
 	Set fso = nothing
	Err.Clear
	 echo "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('loading').style.display = 'none';}catch(e){};</script>"
%>  
 	  
	  </td>
    </tr>
 </form>	
  </table>	





   
</body>
</html>
