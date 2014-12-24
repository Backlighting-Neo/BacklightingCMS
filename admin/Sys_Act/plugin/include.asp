<!--#include file="../../ACT.Function.asp"-->
<% Response.Buffer=True %>
<% Server.ScriptTimeout=99999999 %>
<%
If Not ACTCMS.ChkAdmin() Then '超级管理员检测
	Call Actcms.ACTCMSErr("")
End If 
Dim XML
'插件参考haphic在Z-blog上的插件管理插件
Function GetNode(ID)
	 GetNode=XML.documentElement.selectSingleNode(ID).text
End Function 

Dim ThemeID,ThemeName,ThemeURL,ThemeNote,ThemePubDate,ThemeScreenShot
Dim ThemeAdapted,ThemeVersion,ThemeModified,ThemeDescription,ThemeAdminUrl
 
Dim ThemePlugin_Name,ThemePlugin_Note,ThemePlugin_Type
Dim ThemePlugin_Path,ThemePlugin_Include,ThemePlugin_Level

Dim ThemeSource_Name,ThemeSource_Url,ThemeSource_Email
Dim ThemeAuthor_Name,ThemeAuthor_Url,ThemeAuthor_Email
Dim Action,SelectedTheme,SelectedThemeName,NewVersionExists
Const Theme_URL="http://download.actcms.com/"
Const DownLoad_URL="http://download.actcms.com/plugin/List.asp?"
Const Resource_URL="http://download.actcms.com/plugin/Update.asp?theme="    '注意. Include 文件里还有一同名变量要修改
Const Update_URL="http://themes.actcms.com/plugin/download.asp?theme="

Const XML_Pack_Ver="1.0"
Const XML_Pack_Type="plugin"
Const XML_Pack_Version="ACTCMS_1_1"
Dim Tpath,ThemePath 
 Tpath=server.MapPath(actcms.ActSys)&"\"
ThemePath=actcms.SysPlusPath&"\"
'               Tpath&ThemePath
 '定义超时时间
Const SiteResolve = 5    'UNISON_SiteResolve(Msxml2.ServerXMLHTTP有效)域名分析超时(秒)推荐为"5"	'提示 1秒=1000毫秒
Const SiteConnect = 5    'UNISON_SiteConnect(Msxml2.ServerXMLHTTP有效)连接站点超时(秒)推荐为"5"
Const SiteSend = 4    'UNISON_SiteSend(Msxml2.ServerXMLHTTP有效)发送数据时间超时(秒)推荐为"4"
Const SiteReceive = 10    'UNISON_SiteReceive(Msxml2.ServerXMLHTTP有效)等待反馈时间超时(秒)推荐为"10"

  '*********************************************************
 Function LoadFromFile(strFullName,strCharset)
 	 On Error Resume Next
 	Dim objStream
 	Set objStream = Server.CreateObject("ADODB.Stream")
	With objStream
	.Type = 2
	.Mode = adModeReadWrite
	.Open
	.Charset = strCharset
	.Position = objStream.Size
	.LoadFromFile strFullName
	LoadFromFile=.ReadText
	.Close
	End With
	Set objStream = Nothing

	Err.Clear

End Function
'*********************************************************

Function display(str)
	Echo "<script language=""JavaScript"" type=""text/javascript"">document.getElementById('"&str&"').style.display = 'none';</script>"
End Function 

 '*********************************************************
' 目的：    复制文件
'*********************************************************
Function CopyFile(SFile,DFile)
	 On Error Resume Next
	Dim fso
		CopyFile = True

Set fso = Server.CreateObject(ACTCMS.FsoName)
  		fso.CopyFile SFile, DFile
	If Err.Number = 53 Then
		CopyFile = True
 		Err.Clear
		Set fso=Nothing
		Exit Function
	Elseif Err.Number = 70 Then
		CopyFile = True
 		Err.Clear
		Set fso=Nothing
		Exit Function
	Elseif Err.Number <> 0 Then
 		CopyFile = True
 		Err.Clear
		Set fso=Nothing
		Exit Function
	Else
 		CopyFile = False
	End If
	Set fso=Nothing
End Function

Function SaveToFile(strFullName,strContent)
 	 On Error Resume Next
 	Dim objStream
 	Set objStream = Server.CreateObject("ADODB.Stream")
	With objStream
	.Type = 2
	.Mode = 3
	.Open
	.Charset = Response.Charset
	.Position = objStream.Size
	.WriteText = strContent
	.SaveToFile strFullName,2
	.Close
	End With
	Set objStream = Nothing
End Function
Function RemoveBOM(strFullName)
 	On Error Resume Next
 	Dim objStream
	Dim strContent

	Set objStream = Server.CreateObject("ADODB.Stream")
	With objStream
	.Type = adTypeBinary
	.Mode = adModeReadWrite
	.Open
	.Position = objStream.Size
	.LoadFromFile strFullName
	.Position = 3
	strContent=.Read
	.Close
	End With
	Set objStream = NoThing

	Set objStream = Server.CreateObject("ADODB.Stream")
	With objStream
	.Type = adTypeBinary
	.Mode = adModeReadWrite
	.Open
	.Position = objStream.Size
	.Write = strContent
	.SaveToFile strFullName,adSaveCreateOverWrite
	.Close
	End With
	Set objStream = Nothing

	Err.Clear

End Function
 
'*********************************************************
' 目的：    取得文件扩展名
'*********************************************************
Function GetFileExt(sFileName)
	GetFileExt = LCase(Mid(sFileName,InStrRev (sFileName, ".")+1))
End Function
'*********************************************************
' 目的：    检查某目录下的某文件是否存在
'*********************************************************
Function FileExists(fileName)
On Error Resume Next
Dim objFSO
FileExists = False
Set objFSO = CreateObject(ACTCMS.FsoName)
If objFSO.FileExists(fileName) Then
	FileExists = True
End If
Set objFSO = Nothing
Err.Clear
End Function

'*********************************************************
' 目的：    删除文件
'*********************************************************
Function DeleteFile(FileName)
On Error Resume Next
Dim fso
Set fso = Server.CreateObject(ACTCMS.FsoName)
	fso.DeleteFile(FileName)
	FileName=Replace(Replace(FileName,Tpath,""),"\","/")
 If Err.Number = 53 Then
	DeleteFile = 0
	Response.Write "<font color=""green""> √ 文件 """& FileName &"""不存在!</font>"
	Err.Clear
	Set fso=Nothing
	Exit Function
Elseif Err.Number = 70 Then
	DeleteFile = 70
	Response.Write "<font color=""red""> × 文件 """& FileName &"""为只读, 无法删除!</font>"
	Err.Clear
	Set fso=Nothing
	Exit Function
Elseif Err.Number <> 0 Then
	DeleteFile = Err.Number
	Response.Write "<font color=""red""> × 未知错误，错误编码：" & Err.Number & "</font>"
	Err.Clear
	Set fso=Nothing
	Exit Function
Else
	Response.Write "<font color=""green""> √ 文件 """& FileName &"""删除成功.</font>"
	DeleteFile = 0
End If
Set fso = Nothing
End Function
'*********************************************************
' 目的：    删除文件夹
'*********************************************************
Function DeleteFolder(FolderName)
 on Error Resume Next
Dim fso
Set fso = Server.CreateObject(ACTCMS.FsoName)
 	fso.DeleteFolder(FolderName)
If Err.Number = 76 Then
	DeleteFolder = 0
	Response.Write "<font color=""green""> √ 文件夹 """& Replace(FolderName,Tpath,"") &"""不存在!</font>"
	Err.Clear
	Set fso=Nothing
	Exit Function
Elseif Err.Number = 70 Then
	DeleteFolder = 70
	Response.Write "<font color=""red""> × 文件夹 """& Replace(FolderName,Tpath,"") &"""无法操作!</font>"
	Err.Clear
	Set fso=Nothing
	Exit Function
Elseif Err.Number <> 0 Then
	DeleteFolder = Err.Number
	Response.Write "<font color=""red""> × 未知错误，错误编码：" & Err.Number & "</font>"
	Err.Clear
	Set fso=Nothing
	Exit Function
Else
	Response.Write "<font color=""green""> √ 文件夹 """& Replace(FolderName,Tpath,"") &"""删除成功.</font>"
	DeleteFolder = 0
End If
Set fso = Nothing
End Function

Function getxmluser()
	Set XML=ACTCMS.NoAppGetXMLFromFile(Tpath & ACTCMS.ActCMS_Sys(8) & "\Sys_Act\Config\" & "plustheme.xml")
	If IsObject(XML) And XML.readyState=4 And XML.parseError.errorCode = 0 Then
      getxmluser="&U="&GetNode("UserName")&"&"&"P="&GetNode("PassWord")&"&Domain="&GetNode("Domain")&"&ver="&actcms.Version&"&charset="&response.charset
			Set XML=Nothing
	End If
 End Function 
'*********************************************************
' 目的：    取得目标网页的html代码
'*********************************************************
Function getHTTPPage(url)
 Dim Http,ServerConn
Dim j
For j=0 To 2
	Set Http=server.createobject("Msxml2.ServerXMLHTTP")
	Http.setTimeouts SiteResolve*1000,SiteConnect*1000,SiteSend*1000,SiteReceive*1000
	Http.open "GET",url&getxmluser,False
   	Http.send()
	If Err Then
		Err.Clear
		Set http = Nothing
		ServerConn = False
	else
		ServerConn = true
	End If
	If ServerConn Then
		Exit For
	End If
next
If ServerConn = False Then
	getHTTPPage = False
	Exit Function
End If
 If http.Status=200 Then
	'getHTTPPage=Http.ResponseText
	getHTTPPage=bytesToBSTR(Http.ResponseBody,"utf-8")
Else
	getHTTPPage = False
End If
If Len(getHTTPPage)<60 Then Call display("status"):die getHTTPPage
Set http=Nothing
End Function
'*********************************************************
' 目的：    将目标网页转换为某种编码
'*********************************************************
Function BytesToBstr(strPageContent,strPageCharset)
On Error Resume Next
Dim objstream
Set objstream = Server.CreateObject("adodb.stream")
objstream.Type = 1
objstream.Mode =3
objstream.Open
objstream.Write strPageContent
objstream.Position = 0
objstream.Type = 2
objstream.CharSet = strPageCharset
BytesToBstr = objstream.ReadText
objstream.Close
Set objstream = Nothing
Err.Clear
End Function
'*********************************************************

 '*********************************************************
' 目的：    检查引用
' 输入：    
' 输入：    要替换的字符代号
' 返回：    
'*********************************************************
Function TransferHTML(ByVal source,para)

	Dim objRegExp

	'先换"&"
	If Instr(para,"[&]")>0 Then  source=Replace(source,"&","&amp;")
	If Instr(para,"[<]")>0 Then  source=Replace(source,"<","&lt;")
	If Instr(para,"[>]")>0 Then  source=Replace(source,">","&gt;")
	If Instr(para,"[""]")>0 Then source=Replace(source,"""","&quot;")
	If Instr(para,"[space]")>0 Then source=Replace(source," ","&nbsp;")
	If Instr(para,"[enter]")>0 Then
		source=Replace(source,vbCrLf,"<br/>")
		source=Replace(source,vbLf,"<br/>")
	End If
 
	TransferHTML=source

End Function
'*********************************************************
 '*********************************************************
' 目的：    加载指定目录的文件列表
'*********************************************************
Function LoadIncludeFiles(strDir)

	On Error Resume Next

	Dim aryFileList()
	ReDim aryFileList(0)
 	Dim fso, f, f1, fc, s, i
	Set fso = CreateObject(ACTCMS.FsoName)
	Set f = fso.GetFolder(Tpath & strDir)
	Set fc = f.Files

	i=0

	For Each f1 in fc
		i=i+1
		ReDim Preserve aryFileList(i)
		aryFileList(i)=f1.name 
	Next

	LoadIncludeFiles=aryFileList

	Set fso=nothing

	Err.Clear

End Function
'*********************************************************



%>