<!--#include file="include.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh-CN" lang="zh-CN">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>插件安装</title>
 <link href="../../Images/style.css" rel="stylesheet" type="text/css">
 <link href="Images/New.css" rel="stylesheet" type="text/css">
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
 
</td>
  </tr>
</table>

<%
Dim Install_Error
Install_Error=0
 		echo "<span id=""loading"">正在验证更新, 请稍候...  如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></span>"

Dim Install_Url,Install_ID,Install_Pack,Install_Path,Install_Data 
Install_Url=Request.QueryString("url")
Install_ID=Mid(Install_Url, InStrRev(Install_Url, "theme=")+6)
 If Install_Url="" Then
	echo "<p><font color=""red""> × 插件的下载地址为空.</font></p>"
	Install_Error=Install_Error+1

Else
 	'验证所要安装的插件是否存在
	Action=Request.QueryString("act")
 	If Action<>"confirm" Then
		Response.Flush
			Dim strXmlDataWeb,strXmlDataLocal
 			Dim Item_ID_Web,Item_Name_Web,Item_Url_Web,Item_Version_Web,Item_PubDate_Web,Item_Modified_Web
 			Dim Item_ID_Local,Item_Name_Local,Item_Url_Local,Item_Version_Local,Item_PubDate_Local,Item_Modified_Local,UpdateNeeded
 			Install_Data = getHTTPPage(Resource_URL & Install_ID)
   			Call ActCMS.CreateFolder(actcms.ActSys&actcms.SysPlusPath&"/"&Install_ID&"/")
			Install_Data=Replace(Install_Data,"utf-8","utf-8")
			Install_Pack = Tpath&ThemePath&Install_ID& "\Update.xml"
			Call SaveToFile(Install_Pack,Install_Data)
			Set XML=ACTCMS.NoAppGetXMLFromFile(Tpath&ThemePath &Install_ID& "\Update.xml")
 			Response.Flush
 			If IsObject(XML) And XML.readyState=4  And  XML.parseError.errorCode =0 Then
				Item_ID_Web=GetNode("id")
				Item_Name_Web=GetNode("name")
				Item_Url_Web=GetNode("url")
				Item_Version_Web=GetNode("version")
				Item_PubDate_Web=GetNode("pubdate")
				Item_Modified_Web=GetNode("modified")
				If Item_Version_Web="" Then Item_Version_Web="Undefine"
				If Item_PubDate_Web="" Then Item_PubDate_Web="Undefine"
				If Item_Modified_Web="" Then Item_Modified_Web="Undefine"
		
			End If 
  		    ActCMS.DeleteFile(Tpath&ThemePath&Install_ID& "\Update.xml")
			Set XML=Nothing
 			Set XML=ACTCMS.NoAppGetXMLFromFile(Tpath & ThemePath & Install_ID & "\Theme.xml")
 			Response.Flush
 			If IsObject(XML) And XML.readyState=4  And  XML.parseError.errorCode =0 Then
				Item_ID_Local=GetNode("id")
				Item_Name_Local=GetNode("name")
				Item_Url_Local=GetNode("url")
				Item_Version_Local=GetNode("version")
				Item_PubDate_Local=GetNode("pubdate")
				Item_Modified_Local=GetNode("modified")
				If Item_Version_Local="" Then Item_Version_Local="Undefine"
				If Item_PubDate_Local="" Then Item_PubDate_Local="Undefine"
				If Item_Modified_Local="" Then Item_Modified_Local="Undefine"
			End If 
			If (DateDiff("d",Item_PubDate_Web,Item_PubDate_Local)>0 Or DateDiff("d",Item_Modified_Web,Item_Modified_Local)>0) Then
				UpdateNeeded=False
			ElseIf Item_Version_Web<>Item_Version_Local Or Item_PubDate_Local<>Item_PubDate_Web Or Item_Modified_Local<>Item_Modified_Web Then
				UpdateNeeded=True
			Else
				UpdateNeeded=False
			End If
			
			If (Item_ID_Web<>Item_ID_Local) Then UpdateNeeded=False
		If Action="update" Then
			If (LCase(Item_ID_Web)=LCase(Install_ID)) Then
		 echo "<table width=""98%"" border=""0"" align=""center"" cellpadding=""2"" cellspacing=""1"" class=""table""><tr><td   colspan=""2"" >"
				echo "<p class=""red"">您将对这个插件 <b>("& Item_Name_Web &")</b> 进行修复或升级, 如果继续安装会将其<b>完全覆盖</b>.<br/> 请在继续前确认您对插件中自定义的部分进行了备份.</p>"

				echo "您当前插件版本为: <b>"& Item_Version_Local &"</b>. 发布日期为: <b>"& Item_PubDate_Local &"</b>. 最后修改日期为: <b>"& Item_Modified_Local &"</b>.<br/>"
				echo "将要安装的版本为: <b>"& Item_Version_Web &"</b>. 发布日期为: <b>"& Item_PubDate_Web &"</b>. 最后修改日期为: <b>"& Item_Modified_Web &"</b><br/><br/>"

				If Item_Url_Web<>Empty Then
					echo "<a href="""& Item_Url_Web &""" target=""_blank"" title=""查看插件的发布页面"" class=""notice"">点此查看插件的发布信息!</a><br/><br/>"
				End If
				
				If UpdateNeeded=True Then
					echo "该插件<font color=""red""><b>需要</b></font>升级.<br/><br/>"
				Else
					echo "该插件<font color=""green""><b>不需要</b></font>升级.<br/><br/>"
				End If
 				echo "<a href=""Act.Theme.Install.Asp?act=confirm&amp;url="& Server.URLEncode(Install_Url) &""" title=""继续安装"" class=""notice"">[继续安装]</a> 或 <a href=""javascript:history.back(1);"" title=""返回上一页面""  class=""green"">[取消安装]</a><p>"
			Else
				echo "<div class=""red""><span style="" margin-left:10px"" ><strong>提示：</strong>对不起, 这个插件 <b>("& Item_Name_Local &")</b> 不支持在线安装或升级, 请返回上一页. <a href=""javascript:history.back(1);"" title=""返回上一页面"">[返回上一页]</a></span></div>"
			End If
			Install_Error=Install_Error+1
		Else
 			If (LCase(Item_ID_Local)=LCase(Install_ID)) Then
		 echo "<table width=""98%"" border=""0"" align=""center"" cellpadding=""2"" cellspacing=""1"" class=""table""> <tr><td class=""bg_tr"">查看插件</td></tr><tr><td   colspan=""2"" >"
				echo "<div class=""red""><span style="" margin-left:10px"" >您已安装过这个插件 <b>("& Item_Name_Local &")</b>, 插件文件已存在, 如果继续安装会将其<b>完全覆盖</b>, 这可能会导致您对该插件的个性化修改丢失, 是否继续安装?</span></div>"

				echo "您当前插件版本为: <b>"& Item_Version_Local &"</b>. 发布日期为: <b>"& Item_PubDate_Local &"</b>. 最后修改日期为: <b>"& Item_Modified_Local &"</b>.<br/>"
				echo "将要安装的版本为: <b>"& Item_Version_Web &"</b>. 发布日期为: <b>"& Item_PubDate_Web &"</b>. 最后修改日期为: <b>"& Item_Modified_Web &"</b><br/><br/>"

				If Item_Url_Web<>Empty Then
					echo "<a href="""& Item_Url_Web &""" target=""_blank"" title=""查看插件的发布页面"">点此查看插件的发布信息!</a><br/><br/>"
				End If
 				If UpdateNeeded=True Then
					echo "该插件<font color=""red""><b>需要</b></font>升级.<br/><br/>"
				Else
					echo "该插件<font color=""green""><b>不需要</b></font>升级.<br/><br/>"
				End If

				echo "您还可以跳过安装步骤, 直接查看已安装的插件, 并在该页面中选择应用它. <b>(推荐)</b><br/><br/>"

				echo "<a href=""Act.Theme.Install.Asp?act=confirm&amp;url="& Server.URLEncode(Install_Url) &""" title=""继续安装"" class=""notice"">[继续安装]</a> 或 <a href=""javascript:history.back(1);"" title=""返回上一页面"" class=""green"">[取消安装]</a> 或 <span ><a href=""Act.Theme.Type.asp?theme=" & Server.URLEncode(Install_ID) & """  class=""notice""  title=""直接查看已安装过的("& Install_ID &")插件"">[直接查看已安装过的 """& Install_ID &""" 插件]</a></span><p>"
				Install_Error=Install_Error+1
			End If
		End If
		Set XML=Nothing
		Call display("loading")
		If Install_Error<>0 Then Response.End
	End If
	   Call display("loading")
	
	echo "<span  id=""status"" >正在下载 tpl 插件安装包文件... <img  align=""absmiddle"" src=""images/loading.gif"" /></span>"
	Response.Flush
    echo "<table width=""98%"" border=""0"" align=""center"" cellpadding=""2"" cellspacing=""1"" class=""table""> <tr><td class=""bg_tr"">查看插件</td></tr><tr><td   colspan=""2"" >"
 	Install_Data = getHTTPPage(Install_Url)
  	Install_Data=Replace(Install_Data,"utf-8","utf-8")
 	Install_Data=Replace(Install_Data,"{$sys}",Tpath)
 	Install_Data=Replace(Install_Data,"{$adminsys}",ACTCMS.ActCMS_Sys(8)&"\")
 	Install_Data=Replace(Install_Data,"{$ThemePath}","plugin\")
 	Install_Pack = Tpath&ThemePath& "Theme.tpl"
	Install_Path = Tpath&ThemePath
  	If Install_Data = False Then
		echo "<p><font color=""red""> × tpl 文件下载失败.</font></p>"
		Install_Error=Install_Error+1
	Else
		echo "<p><font color=""green""> √ tpl 文件下载完成.</font></p>"
	End If
 	Call display("status")
	Response.Flush
  	Call SaveToFile(Install_Pack,Install_Data)
	echo "<p><font color=""green""> √ tpl 文件 """&ACTCMS.SysPlusPath&"/Theme.tpl"" 已被保存到您的空间内.</font></p>"
	Response.Flush
 	echo "<p class=""red""> √ tpl 文件 """&ACTCMS.SysPlusPath&"/Theme.tpl"" 正在解包安装...<p>"
	Response.Flush
 		Dim objNodeList,objFSO,objStream,i,j
  		Set XML=ACTCMS.NoAppGetXMLFromFile(Install_Pack)
  		If Not  IsObject(XML) And XML.readyState<>4 Then
 			echo "<p><font color=""red""> × tpl 文件未准备就绪, 无法解包.</font></p>"
			Install_Error=Install_Error+1
		Else
			If XML.parseError.errorCode <> 0 Then
				echo "<p><font color=""red""> × tpl 文件有错误, 无法解包.</font></p>"
				Install_Error=Install_Error+1
 			Else

				Dim Pack_ver,Pack_Type,Pack_For,Pack_ID,Pack_Name
				Pack_Ver = XML.documentElement.SelectSingleNode("//root").getAttributeNode("version").value
				Pack_Type = XML.documentElement.selectSingleNode("//root").getAttributeNode("type").value
				Pack_For = XML.documentElement.selectSingleNode("//root").getAttributeNode("for").value
				Pack_ID = GetNode("id")
				Pack_Name = GetNode("name")
				
				If (CDbl(Pack_Ver) > CDbl(XML_Pack_Ver)) Then
					echo "<p><font color=""red""> × tpl 文件的 XML 版本为 "& Pack_Ver &", 而你的解包器版本为 "& XML_Pack_Ver &", 请升级您的 插件插件, 安装被中止.</font></p>"
					Install_Error=Install_Error+1
				ElseIf (LCase(Pack_Type) <> LCase(XML_Pack_Type)) Then
					echo "<p><font color=""red""> × 不是 tpl 文件, 而可能是 "& Pack_Type &", 安装被中止.</font></p>"
					Install_Error=Install_Error+1
			 	ElseIf (LCase(Pack_For) <> LCase(XML_Pack_Version)) Then
				 	echo "<p><font color=""red""> × tpl 文件版本不符合, 该版本可能是 "& Pack_For &", 安装被中止.</font></p>"
				 	Install_Error=Install_Error+1
				Else
					echo "<div ><textarea   style=""height:150px;width:99%""   cols="" rows="">"
					Set objNodeList = XML.documentElement.selectNodes("//folder/path")
					Set objFSO = CreateObject(ActCMS.FsoName)
						j=objNodeList.length-1
						For i=0 To j
 							If objFSO.FolderExists(objNodeList(i).text)=False Then
								objFSO.CreateFolder(objNodeList(i).text)
							End If
							echo "创建目录" & objNodeList(i).text & vbCrLf
							Response.Flush
						Next
 					Set objFSO = Nothing
					Set objNodeList = Nothing
					Set objNodeList = XML.documentElement.selectNodes("//file/path")
 						j=objNodeList.length-1
						For i=0 To j
							Set objStream = CreateObject("ADODB.Stream")
								With objStream
									.Type = 1
									.Open
									.Write objNodeList(i).nextSibling.nodeTypedvalue
									.SaveToFile  objNodeList(i).text,2
								 	echo "释放文件" & objNodeList(i).text & vbCrLf
									Response.Flush
									.Close
								End With
							Set objStream = Nothing
						Next
					Set objNodeList = Nothing
					echo "</textarea></div>"

				End If

			End If
		End If
		
	Set XML = Nothing
 	If Err.Number<>0 Then Install_Error=Install_Error+1
	Err.Clear
 	echo "<p>"
	Install_Error = Install_Error + DeleteFile(Tpath&ThemePath & "Theme.tpl")
	echo "</p>"
 	If Install_Error = 0 Then 
		echo "<p>"
		Install_Error = Install_Error + DeleteFile(Tpath & ThemePath & Pack_ID & "\verchk.xml")
		echo "</p>"
	End If
 	Response.Flush
 End If

If Install_Error = 0 Then
	echo "<span style=""color:#06F""><br /> √ 插件安装完成. 如果您的浏览器没能自动跳转, 请 <a href=""Act.Theme.Type.asp?theme=" & Server.URLEncode(Pack_ID) & "&amp;themename=" & Server.URLEncode(Pack_Name) & """  class=""green"">[点击这里]<br /><br /></span>"
 Else
	echo "<span style=""color:#06F""><font color=""red""> × 插件安装失败. "
	echo "<a href=""javascript:window.location.reload();"" title=""返回上一个页面""  class=""green""><span>[点此重试]</span></a> 或 <a href=""Index.asp"" title=""返回资源列表页""  class=""green"">[点此返回资源列表页]</span>"
End If
 Call display("loading")
%>
	</div>
</div>
</div>
</td>
  </tr>
   </table>
</body>
</html>
 