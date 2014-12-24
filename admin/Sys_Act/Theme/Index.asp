<!--#include file="include.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
 <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>主题安装</title>
 <link href="../../Images/style.css" rel="stylesheet" type="text/css">
 <link href="Images/New.css" rel="stylesheet" type="text/css">
   </head>
 <body> 
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td class="bg_tr">您现在的位置：主题样式 &gt;&gt; 主题管理</td>
  </tr>
  <tr>
    <td>
  <div class="md-head" >
<div class="zsj"></div>
<a href="Act.Theme.ACTList.asp" class="a10 wrap"  >在线安装主题</a>
<a href="Act.Theme.Reg.asp" class="a10 wrap"  >我的帐号</a>
<a href="Index.asp" class="a10 wrap cur wid"  >主题管理</a>
<a href="Act.Theme.Check.asp" class="a10 wrap"  >查看主题的可用更新</a>
 </div>
 
</td>
  </tr>
</table>
 
  <%
Dim title
title=actcms.s("title")
If title<>"" Then 
	echo "<span id=""title""><strong>提示：</strong>"&title&"</span>"
Else 
	echo "<div class=""top1""><strong>提示：</strong>点击可进入查看并对其进行管理操作</div>"
End If 
Action=Request.QueryString("act")
NewVersionExists=False
If Action = "themedel" Then

	SelectedTheme=Request.QueryString("theme")
	SelectedThemeName=Request.QueryString("themename")

	If UCase(SelectedTheme)=Ucase(ACTCMS.NowTheme) Then
		echo "<p class=""red"">您请求的主题正在使用, 无法删除...</p>"
		echo "<script>setTimeout(""self.history.back(1)"",2000);</script>"
		Response.End
	End If

	Dim DelError
	DelError = 0

	If SelectedTheme<>"" Then
		echo "<p class=""red"">正在处理您的请求...</p>"
		Response.Flush

		echo "<p>"
		DelError = DelError + DeleteFolder(Tpath & ThemePath & SelectedTheme)
		echo "</p>"
	Else
		echo "<p class=""red"">请求的参数错误, 正在退出...</p>"
		Response.Flush
		DelError = 13
	End If

	If DelError = 0 Then
		echo "<p><font color=""green""> √ 主题 - " & SelectedThemeName & "  删除成功!</font><p>"
	Else
		echo "<p><font color=""red""> × 主题 - " & SelectedThemeName & "  删除失败! 请手动删除之.</font><p>"
	End If

	echo "<p class=""red"">如果您的浏览器没能自动跳转 请 <a href=""Index.asp"">[点击这里]</a>.<p>"
	echo "<script>setTimeout(""self.location.href='Index.asp'"",1500);</script>"
	response.end
End If

echo "<span id=""loading"">正在载入主题列表, 请稍候... 如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></span>"
	Response.Flush
  	echo "<div class=""top1"" id=""newVersion""  style=""display:none;""><p ><a href=""Act.Theme.Check.asp"" class=""notice"" title=""查看主题的可用更新"">[发现了您安装的某个主题有了新版本, 点此查看现有主题的可用更新]</a>.</p></div>"
	echo "<table width=""98%"" border=""0"" align=""center"" cellpadding=""2"" cellspacing=""1"" class=""table"">"
	echo  "<tr> <td   colspan=""2""  align=""center""> <div class=""zhuti""><ul>"
	Dim objXmlFile,strXmlFile
	Dim fso, f, f1, fc, s, t
	Set fso = CreateObject(ACTCMS.FsoName)
  	Set f = fso.GetFolder(Tpath & ThemePath)
	Set fc = f.SubFolders
	For Each f1 in fc
 		ThemeSource_Name=Empty:ThemeSource_Url=Empty:ThemeID=Empty:ThemeName=Empty:ThemeURL=Empty:ThemeNote=Empty:ThemeModified=Empty
 		If fso.FileExists(Tpath & ThemePath & f1.name & "/" & "theme.xml") Then
   			Set XML=ACTCMS.NoAppGetXMLFromFile(Tpath & ThemePath  & f1.name & "/" & "theme.xml")
			If IsObject(XML) And XML.readyState=4 Then
 				If XML.parseError.errorCode <> 0 Then 
				Else 
					 ThemeAuthor_Name=GetNode("author/name"):ThemeAuthor_Url=GetNode("author/url"):ThemeSource_Name=GetNode("source/name")
					 ThemeSource_Url=GetNode("source/url"):ThemeID=GetNode("id"):ThemeName=GetNode("name"):ThemeURL=GetNode("url")
					 ThemePubDate=GetNode("pubdate"):ThemeModified=GetNode("modified"):ThemeNote=GetNode("note")
					 If ThemeAuthor_Name=Empty Then
						ThemeAuthor_Name=ThemeSource_Name
						ThemeAuthor_Url=ThemeSource_Url
					 End If
					If ThemeModified=Empty Then ThemeModified=ThemePubDate
 					ThemeNote=ACTCMS.CloseHtml(ThemeNote)
					If Len(ThemeNote)>38 then ThemeNote=Left(ThemeNote,31) & "...<a href=""Act.Theme.Type.asp?theme=" & Server.URLEncode(ThemeID) & """>more</a>"
 				End If 
			Set XML=Nothing
 		Else
 			ThemeSource_Name="unknown":ThemeSource_Url=Empty:ThemeID=f1.name
			ThemeName=f1.name:ThemeURL=Empty:ThemeNote="unknown":ThemeModified="unknown"
 	  End If 
   	  End If
  		If fso.FileExists(Tpath & ThemePath & f1.name & "/" & "verchk.xml") Then
			t="<a class=""notice"" href=""Act.Theme.Install.Asp?act=update&amp;url=" & Server.URLEncode(Update_URL & f1.name) & """ title=""升级主题""><b >发现新版本!</b></a>"
 			NewVersionExists=True
		ElseIf fso.FileExists(Tpath & ThemePath& f1.name & "/" & "error.log") Then
			t="<b class=""red"">不支持在线更新.</b>"
		Else
			t=""
		End If

		If fso.FileExists(Tpath & ThemePath & f1.name & "/" & "preview.gif") Then
			ThemeScreenShot="../../../" &ThemePath & f1.name & "/" & "preview.gif"
		Else
			ThemeScreenShot="Images/nopreview.jpg"
		End If
	
	Dim thistitle
	If UCase(ThemeID)=UCase(actcms.NowTheme) Then
		echo "<li class=""thisclass""> "
		thistitle="<font color=red>使用中</font>"
    Else 
		echo "<li > "
		thistitle="使用"
	End If

	If UCase(ThemeID)<>UCase(f1.name) Then
		echo "  <div class=""zhuti_title"">该主题ID错误, 请 <a href=""Act.Theme.Edit.asp?theme=" & Server.URLEncode(f1.name) & """ title=""编辑主题信息""><font color=""red""><b>[重新编辑主题信息]</b></font></a>.<a href=""Index.asp?act=themedel&amp;theme=" & Server.URLEncode(f1.name) & "&amp;themename=" & Server.URLEncode(ThemeName) & """ title=""删除此主题"" onclick=""return window.confirm('您将删除此主题的所有文件, 确定吗?');"">删除</a></div>" 
   	Else
	      
          echo "      <div class=""zhuti_title"">"
          echo t&"      <a href=""Act.Theme.ThemeInstall.asp?Install=yes&ThemeID=" & Server.URLEncode(f1.name) & """  class=""green""  title=""使用当前主题""> ["&thistitle&"]</a>   <a href=""Act.Theme.Edit.asp?theme=" & Server.URLEncode(f1.name) & """ title=""编辑主题信息"">[编辑主题]</a>"
          echo "         <a href=""../../include/TempLate.asp?ShowPath="&actcms.ThisTheme&""" title=""编辑模版"">编辑模版</a>"
          echo " <a href=""Act.Theme.Install.Asp?act=update&amp;url=" & Server.URLEncode(Update_URL & ThemeID) & """ title=""升级修复主题"">升级</a>"
          echo "<a href=""Index.asp?act=themedel&amp;theme=" & Server.URLEncode(f1.name) & "&amp;themename=" & Server.URLEncode(ThemeName) & """ title=""删除此主题"" onclick=""return window.confirm('您将删除此主题的所有文件, 此操作不可恢复,确定吗?');"">删除</a>"
          echo "    </div>  "
          echo "    <a href=""Act.Theme.Type.asp?theme=" & Server.URLEncode(f1.name) & "&amp;themename=" & Server.URLEncode(ThemeName) & """ class=""zhuti_img""><img src=""" & ThemeScreenShot & """ title=""点击查看 " & ThemeName & " 的详细信息!""   width=""160"" height=""140""  border=""0""/><div class=""this""></div></a>"
		   echo "    <div class=""zhuti_js"">"
		If ThemeURL=Empty Then
		  echo "    <span>名称:" & ThemeName & "</span> "     
 		Else
		  echo "    <span>名称:<a href=""" & ThemeURL & """ target=""_blank"" title=""主题发布地址"">" & ThemeName & "</a></span> "     
			
  		End If

         
		  
		 If ThemeAuthor_Url=Empty Then
		  echo "    <span>作者:" & ThemeAuthor_Name & "</span>  "    
 		 Else
		  echo "    <span>作者:<a href=""" & ThemeAuthor_Url & """ target=""_blank"" title=""作者主页"">" & ThemeAuthor_Name & "</a></span>  "    
  		 End If
		  
          echo "   <span>修改时间:" & ThemeModified & "</span>"
          echo "   <span>" & ThemeNote & "</span> "
          echo "    </div>  	"
     
		echo " </li>" 
 	End If 
 	Next
	Set fso = nothing
	Err.Clear
 	%>
     </ul>
    <div style="clear:both"></div>
    </div>
</td>
  </tr>
   </table>
 <%
	If NewVersionExists Then
		echo "<script language=""JavaScript"" type=""text/javascript"">document.getElementById('newVersion').style.display = 'block';</script>"
	End If
	Response.Flush
 	echo "<script language=""JavaScript"" type=""text/javascript"">try{document.getElementById('loading').style.display = 'none';}catch(e){};</script>"
 %>
 </body>
</html>
