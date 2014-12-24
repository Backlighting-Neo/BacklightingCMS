<!--#include file="include.asp"-->
<%Action=Request.QueryString("act")
SelectedTheme=Request.QueryString("theme")
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>主题安装</title>
<link href="../../Images/style.css" rel="stylesheet" type="text/css">
<link href="Images/new.css" rel="stylesheet" type="text/css">
 </head>

<body> 

<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td class="bg_tr">您现在的位置：主题样式修改 &gt;&gt; 浏览</td>
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


If Action="save" Then
 	Echo "<p id=""loading2"">正在写入主题信息, 请稍候...  如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
 	Dim Pack_Error
	Pack_Error=0

	If SelectedTheme="" Then
		Pack_Error= "<font color=""red"">  主题的名称为空.</font>"
  	Else
		Echo "<p ><font color=""Navy"">正在保存XML...</font><p>"
 		Dim ZipPathFile
 		'打包文件目录与生成文件名
		ZipPathFile = Tpath & ThemePath & SelectedTheme & "\Theme.xml"

		'开始打包
		CreateXml(ZipPathFile)
	End If

	If Pack_Error = "0" Then
 		Call Actcms.ActErr("主题信息保存完成","Sys_Act/Theme/Act.Theme.Type.asp?theme="& Server.URLEncode(SelectedTheme) &"","")
 	Else
  		Call Actcms.ActErr("主题信息保存失败"&Pack_Error,"","1")
 	End If
 	Echo "<script language=""JavaScript"" type=""text/javascript"">document.getElementById('loading2').style.display = 'none';</script>"
End If
If Action="" Then
	Echo "<p id=""loading"">正在载入主题信息, 请稍候...  如果长时间停止响应, 请 <a href=""javascript:window.location.reload();"" title=""点此重试"">[点此重试]</a></p>"
	Response.Flush

	Call EditXMLPackInfo()

	Call display("loading")
End If

%>
 
</body>
</html>
<%
 Sub EditXMLPackInfo()
'On Error Resume Next
 	Dim fso,strXmlFile
	Set fso = CreateObject(ACTCMS.FsoName)
 		If fso.FileExists(Tpath & ThemePath & SelectedTheme & "/" & "theme.xml") Then
 			strXmlFile =Tpath & ThemePath & SelectedTheme & "/" & "theme.xml"
    		Set XML=ACTCMS.NoAppGetXMLFromFile(strXmlFile)
			If IsObject(XML) And XML.readyState=4 And XML.parseError.errorCode = 0 Then
 					ThemeID=GetNode("id"):ThemeName=GetNode("name"):ThemeURL=GetNode("url"):ThemeNote=GetNode("note")
 					ThemeAuthor_Name=GetNode("author/name"):ThemeAuthor_Url=GetNode("author/url"):ThemeAuthor_Email=GetNode("author/email")
 					ThemeSource_Name=GetNode("source/name"):ThemeSource_Url=GetNode("source/url"):ThemeSource_Email=GetNode("source/email")
 					ThemeAdapted=GetNode("adapted"):ThemeVersion=GetNode("version"):ThemePubDate=GetNode("pubdate")
					ThemeModified=GetNode("modified"):ThemeDescription=GetNode("description")
 					ThemeAuthor_Name=server.HTMLEncode(ThemeAuthor_Name)
					ThemeSource_Name=server.HTMLEncode(ThemeSource_Name)
					ThemeName=server.HTMLEncode(ThemeName)
					ThemeNote=server.HTMLEncode(ThemeNote)
					ThemeDescription=server.HTMLEncode(ThemeDescription)
 					Set XML=Nothing
			End If
 		Else
 			ThemeID=SelectedTheme:ThemeName=SelectedTheme
 			ThemeAdapted="3.0":ThemePubDate=Date():ThemeModified=Date()
 		End If
	Set fso = nothing
	Err.Clear
 %>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <form method="post" action="Act.Theme.Edit.asp?act=save&theme=<%=SelectedTheme%>">
 


    <tr>
      <td height="25" align="right">主题ID&nbsp;&nbsp;</td>
      <td height="25"><input  class="Ainput"   name="ThemeID" type="text"  value="<%=SelectedTheme%>"  size="30" readonly disabled=true/>
      (主题ID应为主题文件夹名称, 由编辑器自动完成填写, 不可修改.) </td>
    </tr>
    
     
    <tr>
      <td height="25" align="right">主题名称&nbsp;&nbsp;</td>
      <td height="25"><input  class="Ainput"   name="ThemeName" type="text" size="30" value="<%=ThemeName%>" />
       </td>
    </tr>
     
    <tr>
      <td height="25" align="right">主题的发布页面&nbsp;&nbsp;</td>
      <td height="25"><input  class="Ainput"   name="ThemeURL" type="text" size="60" value="<%=ThemeURL%>" />
      (带 http:// 等协议名的页面地址, 以方便使用者获取更多的主题信息) </td>
    </tr>
     
    <tr>
      <td height="25" align="right">主题简介&nbsp;&nbsp;</td>
      <td height="25">
	  
	  <textarea name="ThemeNote" cols="75" rows="4" ><%=ThemeNote%></textarea>
       (可以用 &lt;br /&gt; 换行, 可以使用 html 标签)</td>
    </tr>
     
    <tr>
      <td height="25" align="right">适用的版本&nbsp;&nbsp;</td>
      <td height="25"><input  class="Ainput"   name="ThemeAdapted" type="text" size="30" value="<%=ThemeAdapted%>" />
       (直接写版本号,如果有多个版本,请用,号分隔)</td>
    </tr>
 
    <tr >
      <td height="25"   colspan="2"  class="bg_tr" align="center"> 
       以下信息对查找主题可用更新极为重要, 建议在每次修改主题后更新这些信息</td>
    </tr>
 
    <tr>
      <td height="25" align="right">主题的版本号&nbsp;&nbsp;</td>
      <td height="25"><input  class="Ainput"   name="ThemeVersion" type="text" size="30" value="<%=ThemeVersion%>" />
      </td>
    </tr>
 
    <tr>
      <td height="25" align="right">您的主题的发布日期&nbsp;&nbsp;</td>
      <td height="25"><input  class="Ainput"   name="ThemePubDate" type="text" size="30" value="<%=ThemePubDate%>" />
     (日期标准格式:<%=Date()%>) </td>
    </tr>
 
    <tr>
      <td height="25" align="right">最后修改日期&nbsp;&nbsp;</td>
      <td height="25"><input  class="Ainput"   name="ThemeModified" type="text" size="30" value="<%=ThemeModified%>" />
     最后修改日期: (日期标准格式:<%=Date()%>) </td>
    </tr>
 
    <tr>
      <td height="25" align="right">作者名称&nbsp;&nbsp;</td>
      <td height="25"><input  class="Ainput"   name="ThemeAuthor_Name" type="text" size="30" value="<%=ThemeAuthor_Name%>" />
      </td>
    </tr>
 
    <tr>
      <td height="25" align="right">作者网址&nbsp;&nbsp;</td>
      <td height="25"><input  class="Ainput"   name="ThemeAuthor_URL" type="text" size="30" value="<%=ThemeAuthor_URL%>" />
      </td>
    </tr>
 
    <tr>
      <td height="25" align="right">作者 Email&nbsp;&nbsp;</td>
      <td height="25"><input  class="Ainput"   name="ThemeAuthor_Email" type="text" size="30" value="<%=ThemeAuthor_Email%>" />
      </td>
    </tr>
 
 
    <tr>
      <td height="25" align="right">源作者名称&nbsp;&nbsp;</td>
      <td height="25"><input  class="Ainput"   name="ThemeSource_Name" type="text" size="30" value="<%=ThemeSource_Name%>" />
      </td>
    </tr>
 
 
    <tr>
      <td height="25" align="right">源作者网址&nbsp;&nbsp;</td>
      <td height="25"><input  class="Ainput"   name="ThemeSource_URL" type="text" size="30" value="<%=ThemeSource_URL%>" />
      </td>
    </tr>
 
 
    <tr>
      <td height="25" align="right">源作者 Email&nbsp;&nbsp;</td>
      <td height="25"><input  class="Ainput"   name="ThemeSource_Email" type="text" size="30" value="<%=ThemeSource_Email%>" />
      </td>
    </tr>
 
    <tr>
      <td height="25" align="right">详细说明 (可应用 HTML 代码)&nbsp;&nbsp;</td>
      <td height="25">
	  
	    <textarea name="ThemeDescription" cols="75" rows="6" ><%=ThemeDescription%></textarea>
 
	   </td>
    </tr>
  
 
 <tr>
      <td   colspan="2"  align="center">
      <input type="submit" class="ACT_btn"  name=Submit1 value="  完成编辑并保存信息  " />
      <input type="button" name="Submit2"  onclick="self.location.href='Index.asp';" class="ACT_btn" value="  取消并返回主题管理  " /></td>
    </tr>
 
 
  </table>	
</form>	
<%	 

End Sub


'创建一个空的XML文件，为写入文件作准备
Sub CreateXml(FilePath)
'On Error Resume Next
 	Dim XmlDoc,Root,xRoot
	Set XmlDoc = Server.CreateObject("Microsoft.XMLDOM")
		XmlDoc.async = False
		XmlDoc.ValidateOnParse=False
		Set Root = XmlDoc.createProcessingInstruction("xml","version='1.0' encoding='utf-8' standalone='yes'")
		XmlDoc.appendChild(Root)
		Set xRoot = XmlDoc.appendChild(XmlDoc.CreateElement("theme"))
			xRoot.setAttribute "version",XML_Pack_Ver
		Set xRoot = Nothing

		'写入文件信息

		Dim ThemeAuthor,ThemeSource 
		Dim XMLcdata

		Set ThemeID = XmlDoc.SelectSingleNode("//theme").AppendChild(XmlDoc.CreateElement("id"))
			ThemeID.Text = SelectedTheme
		Set ThemeID=Nothing

		Set ThemeName = XmlDoc.SelectSingleNode("//theme").AppendChild(XmlDoc.CreateElement("name"))
			ThemeName.Text = Request.Form("ThemeName")
		Set ThemeName=Nothing

		Set ThemeURL = XmlDoc.SelectSingleNode("//theme").AppendChild(XmlDoc.CreateElement("url"))
			ThemeURL.Text = Request.Form("ThemeURL")
		Set ThemeURL=Nothing

		Set ThemeNote = XmlDoc.SelectSingleNode("//theme").AppendChild(XmlDoc.CreateElement("note"))
			ThemeNote.Text = Replace(Replace(Request.Form("ThemeNote"),vbCr,""),vbLf,"")
		Set ThemeNote=Nothing


		Set ThemeAuthor = XmlDoc.SelectSingleNode("//theme").AppendChild(XmlDoc.CreateElement("author"))

			Set ThemeAuthor_Name = ThemeAuthor.AppendChild(XmlDoc.CreateElement("name"))
				ThemeAuthor_Name.Text = Request.Form("ThemeAuthor_Name")
			Set ThemeAuthor_Name=Nothing

			Set ThemeAuthor_URL = ThemeAuthor.AppendChild(XmlDoc.CreateElement("url"))
				ThemeAuthor_URL.Text = Request.Form("ThemeAuthor_URL")
			Set ThemeAuthor_URL=Nothing

			Set ThemeAuthor_Email = ThemeAuthor.AppendChild(XmlDoc.CreateElement("email"))
				ThemeAuthor_Email.Text = Request.Form("ThemeAuthor_Email")
			Set ThemeAuthor_Email=Nothing

		Set ThemeAuthor=Nothing


		Set ThemeSource = XmlDoc.SelectSingleNode("//theme").AppendChild(XmlDoc.CreateElement("source"))

			Set ThemeSource_Name = ThemeSource.AppendChild(XmlDoc.CreateElement("name"))
				ThemeSource_Name.Text = Request.Form("ThemeSource_Name")
			Set ThemeSource_Name=Nothing

			Set ThemeSource_URL = ThemeSource.AppendChild(XmlDoc.CreateElement("url"))
				ThemeSource_URL.Text = Request.Form("ThemeSource_URL")
			Set ThemeSource_URL=Nothing

			Set ThemeSource_Email = ThemeSource.AppendChild(XmlDoc.CreateElement("email"))
				ThemeSource_Email.Text = Request.Form("ThemeSource_Email")
			Set ThemeSource_Email=Nothing

		Set ThemeSource=Nothing

 

		Set ThemeAdapted = XmlDoc.SelectSingleNode("//theme").AppendChild(XmlDoc.CreateElement("adapted"))
			ThemeAdapted.Text = Request.Form("ThemeAdapted")
		Set ThemeAdapted=Nothing

		Set ThemeVersion = XmlDoc.SelectSingleNode("//theme").AppendChild(XmlDoc.CreateElement("version"))
			ThemeVersion.Text = Request.Form("ThemeVersion")
		Set ThemeVersion=Nothing

		Set ThemePubDate = XmlDoc.SelectSingleNode("//theme").AppendChild(XmlDoc.CreateElement("pubdate"))
			ThemePubDate.Text = Request.Form("ThemePubDate")
		Set ThemePubDate=Nothing

		Set ThemeModified = XmlDoc.SelectSingleNode("//theme").AppendChild(XmlDoc.CreateElement("modified"))
			ThemeModified.Text = Request.Form("ThemeModified")
		Set ThemeModified=Nothing


		Dim CThemeDescription
		Set ThemeDescription = XmlDoc.SelectSingleNode("//theme").AppendChild(XmlDoc.CreateElement("description"))
			Set XMLcdata = XmlDoc.createNode("cdatasection", "","")
				XMLcdata.NodeValue = Request.Form("ThemeDescription")
			Set CThemeDescription = ThemeDescription.AppendChild(XMLcdata)
			Set CThemeDescription = Nothing
			Set XMLcdata = Nothing
		Set ThemeDescription=Nothing
 		XmlDoc.Save(FilePath)
		Set Root = Nothing
	Set XmlDoc = Nothing

 End Sub

 
%>