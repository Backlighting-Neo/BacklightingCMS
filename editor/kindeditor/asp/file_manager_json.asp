<%@ CODEPAGE=65001 %>
<% Option Explicit %>
<% Response.CodePage=65001 %>
<% Response.Charset="UTF-8" %>
<!--#include file="JSON_2.0.4.asp"-->
<%

' KindEditor ASP
'
' 鏈珹SP绋嬪簭鏄紨绀虹▼搴忥紝寤鸿涓嶈鐩存帴鍦ㄥ疄闄呴」鐩腑浣跨敤銆?
' 濡傛灉鎮ㄧ‘瀹氱洿鎺ヤ娇鐢ㄦ湰绋嬪簭锛屼娇鐢ㄤ箣鍓嶈浠旂粏纭鐩稿叧瀹夊叏璁剧疆銆?
'

Dim aspUrl, rootPath, rootUrl, fileTypes
Dim currentPath, currentUrl, currentDirPath, moveupDirPath
Dim path, order, fso, folder, dir, file, result
Dim fileExt, dirCount, fileCount, orderIndex, i, j
Dim dirList(), fileList(), isDir, hasFile, filesize, isPhoto, filetype, filename, datetime

aspUrl = Request.ServerVariables("SCRIPT_NAME")
aspUrl = left(aspUrl, InStrRev(aspUrl, "/"))
aspUrl=""
'鏍圭洰褰曡矾寰勶紝鍙互鎸囧畾缁濆璺緞锛屾瘮濡?/var/www/attached/
rootPath = "/UpFiles/"
'鏍圭洰褰昒RL锛屽彲浠ユ寚瀹氱粷瀵硅矾寰勶紝姣斿 http://www.yoursite.com/attached/
rootUrl = aspUrl & "/UpFiles/"
'鍥剧墖鎵╁睍鍚?
fileTypes = "gif,jpg,jpeg,png,bmp"

currentPath = ""
currentUrl = ""
currentDirPath = ""
moveupDirPath = ""

'鏍规嵁path鍙傛暟锛岃缃悇璺緞鍜孶RL
path = Request.QueryString("path")
If path = "" Then
	currentPath = Server.MapPath(rootPath) & "\"
	currentUrl = rootUrl
	currentDirPath = ""
	moveupDirPath = ""
Else
	currentPath = Server.MapPath(rootPath & path) & "\"
	currentUrl = rootUrl + path
	currentDirPath = path
	moveupDirPath = RegexReplace(currentDirPath, "(.*?)[^\/]+\/$", "$1")
End If

'鎺掑簭褰㈠紡锛宯ame or size or type
order = lcase(Request.QueryString("order"))
Select Case order
	Case "type" orderIndex = 4
	Case "size" orderIndex = 2
	Case Else  orderIndex = 5
End Select

'涓嶅厑璁镐娇鐢?.绉诲姩鍒颁笂涓€绾х洰褰?
If RegexIsMatch(path, "\.\.") Then
	Response.Write "Access is not allowed."
	Response.End
End If
'鏈€鍚庝竴涓瓧绗︿笉鏄?
If path <> "" And Not RegexIsMatch(path, "\/$") Then
	Response.Write "Parameter is not allowed."
	Response.End
End If
'鐩綍涓嶅瓨鍦ㄦ垨涓嶆槸鐩綍
If Not DirectoryExists(currentPath) Then
	Response.Write "Directory does not exist."
	Response.End
End If

Set result = jsObject()
'鐩稿浜庢牴鐩綍鐨勪笂涓€绾х洰褰?
result("moveup_dir_path") = moveupDirPath
'鐩稿浜庢牴鐩綍鐨勫綋鍓嶇洰褰?
result("current_dir_path") = currentDirPath
'褰撳墠鐩綍鐨刄RL
result("current_url") = currentUrl

Set fso = Server.CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(currentPath)

'鏂囦欢鏁?
dirCount = folder.SubFolders.count
fileCount = folder.Files.count
result("total_count") = dirCount + fileCount

ReDim dirList(dirCount)
i = 0
For Each dir in folder.SubFolders
	isDir = True
	hasFile = (dir.Files.count > 0) or (dir.SubFolders.count>0)
	filesize = 0
	isPhoto = False
	filetype = ""
	filename = dir.name
	datetime = FormatDate(dir.DateLastModified)
	dirList(i) = Array(isDir, hasFile, filesize, isPhoto, filetype, filename, datetime)
	i = i + 1
Next
ReDim fileList(fileCount)
i = 0
For Each file in folder.Files
	fileExt = mid(file.name, InStrRev(file.name, ".") + 1)
	isDir = False
	hasFile = False
	filesize = file.size
	isPhoto = (instr(lcase(fileTypes), fileExt) > 0)
	filetype = fileExt
	filename = file.name
	datetime = FormatDate(file.DateLastModified)
	fileList(i) = Array(isDir, hasFile, filesize, isPhoto, filetype, filename, datetime)
	i = i + 1
Next

'鎺掑簭
Dim minidx, temp
For i = 0 To dirCount - 2
	minidx = i
	For j = i + 1 To dirCount - 1
		If (dirList(minidx)(5) > dirList(j)(5)) Then
			minidx = j
		End If
	Next
	If minidx <> i Then
		temp = dirList(minidx)
		dirList(minidx) = dirList(i)
		dirList(i) = temp
	End If
Next
For i = 0 To fileCount - 2
	minidx = i
	For j = i + 1 To fileCount - 1
		If (fileList(minidx)(orderIndex) > fileList(j)(orderIndex)) Then
			minidx = j
		End If
	Next
	If minidx <> i Then
		temp = fileList(minidx)
		fileList(minidx) = fileList(i)
		fileList(i) = temp
	End If
Next

Set result("file_list") = jsArray()
For i = 0 To dirCount - 1
	Set result("file_list")(Null) = jsObject()
	result("file_list")(Null)("is_dir") = dirList(i)(0)
	result("file_list")(Null)("has_file") = dirList(i)(1)
	result("file_list")(Null)("filesize") = dirList(i)(2)
	result("file_list")(Null)("is_photo") = dirList(i)(3)
	result("file_list")(Null)("filetype") = dirList(i)(4)
	result("file_list")(Null)("filename") = dirList(i)(5)
	result("file_list")(Null)("datetime") = dirList(i)(6)
Next
For i = 0 To fileCount - 1
	Set result("file_list")(Null) = jsObject()
	result("file_list")(Null)("is_dir") = fileList(i)(0)
	result("file_list")(Null)("has_file") = fileList(i)(1)
	result("file_list")(Null)("filesize") = fileList(i)(2)
	result("file_list")(Null)("is_photo") = fileList(i)(3)
	result("file_list")(Null)("filetype") = fileList(i)(4)
	result("file_list")(Null)("filename") = fileList(i)(5)
	result("file_list")(Null)("datetime") = fileList(i)(6)
Next

'杈撳嚭JSON瀛楃涓?
Response.AddHeader "Content-Type", "text/html; charset=UTF-8"
result.Flush
Response.End

'鑷畾涔夊嚱鏁?
Function DirectoryExists(dirPath)
	Dim fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	DirectoryExists = fso.FolderExists(dirPath)
End Function

Function RegexIsMatch(subject, pattern)
	Dim reg
	Set reg = New RegExp
	reg.Global = True
	reg.MultiLine = True
	reg.Pattern = pattern
	RegexIsMatch = reg.Test(subject)
End Function

Function RegexReplace(subject, pattern, replacement)
	Dim reg
	Set reg = New RegExp
	reg.Global = True
	reg.MultiLine = True
	reg.Pattern = pattern
	RegexReplace = reg.Replace(subject, replacement)
End Function

Public Function FormatDate(datetime)
	Dim y, m, d, h, i, s
	y = CStr(Year(datetime))
	m = CStr(Month(datetime))
	If Len(m) = 1 Then m = "0" & m
	d = CStr(Day(datetime))
	If Len(d) = 1 Then d = "0" & d
	h = CStr(Hour(datetime))
	If Len(h) = 1 Then h = "0" & h
	i = CStr(Minute(datetime))
	If Len(i) = 1 Then i = "0" & i
	s = CStr(Second(datetime))
	If Len(s) = 1 Then s = "0" & s
	FormatDate = y & "-" & m & "-" & d & " " & h & ":" & i & ":" & s
End Function
%>
