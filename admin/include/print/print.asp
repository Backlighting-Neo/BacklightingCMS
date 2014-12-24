<% Option Explicit %>
<%
Dim PreviewImagePath,FileExtName,FileIconDic,FileIcon,AvailableShowTypeStr,PicPara
PreviewImagePath = Request("FilePath")
AvailableShowTypeStr = "jpg,gif,bmp,pst,png,ico"
Set FileIconDic = CreateObject("Scripting.Dictionary")
FileIconDic.Add "txt","FileIcon/txt.gif"
FileIconDic.Add "gif","FileIcon/gif.gif"
FileIconDic.Add "exe","FileIcon/exe.gif"
FileIconDic.Add "asp","FileIcon/asp.gif"
FileIconDic.Add "html","FileIcon/html.gif"
FileIconDic.Add "htm","FileIcon/html.gif"
FileIconDic.Add "jpg","FileIcon/jpg.gif"
FileIconDic.Add "jpeg","FileIcon/jpg.gif"
FileIconDic.Add "pl","FileIcon/perl.gif"
FileIconDic.Add "perl","FileIcon/perl.gif"
FileIconDic.Add "zip","FileIcon/zip.gif"
FileIconDic.Add "rar","FileIcon/zip.gif"
FileIconDic.Add "gz","FileIcon/zip.gif"
FileIconDic.Add "doc","FileIcon/doc.gif"
FileIconDic.Add "xml","FileIcon/xml.gif"
FileIconDic.Add "xsl","FileIcon/xml.gif"
FileIconDic.Add "dtd","FileIcon/xml.gif"
FileIconDic.Add "vbs","FileIcon/vbs.gif"
FileIconDic.Add "js","FileIcon/vbs.gif"
FileIconDic.Add "wsh","FileIcon/vbs.gif"
FileIconDic.Add "sql","FileIcon/script.gif"
FileIconDic.Add "bat","FileIcon/script.gif"
FileIconDic.Add "tcl","FileIcon/script.gif"
FileIconDic.Add "eml","FileIcon/mail.gif"
FileIconDic.Add "swf","FileIcon/flash.gif"
if PreviewImagePath = "" then
	PreviewImagePath = "FileIcon/DefaultPreview.gif"
else
	FileExtName = Right(PreviewImagePath,Len(PreviewImagePath)-InStrRev(PreviewImagePath,"."))
	if InStr(AvailableShowTypeStr,FileExtName) = 0 then
		FileIcon = FileIconDic.Item(LCase(FileExtName))
		if FileIcon = "" then
			FileIcon = "FileIcon/unknown.gif"
		end if
		PreviewImagePath = FileIcon
		PicPara = " width=""30"" height=""30"" "
	else
		PicPara = ""
	end if
end if
Set FileIconDic = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<HTML>
<HEAD>
<META http-equiv="Content-Type" content="text/html; charset=utf-8">
<TITLE>图片预览</TITLE>
</HEAD>
<BODY topmargin="0" leftmargin="0">
<TABLE width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <TR>
    <TD align="center" valign="middle" width="100%" height="100%"><div align="center"><IMG <% = PicPara %> src="<% = PreviewImagePath %>"></div></TD>
  </TR>
</TABLE>
</BODY>
</HTML>
