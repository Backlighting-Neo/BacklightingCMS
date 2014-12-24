<!--#include file="../act_inc/ACT.User.asp"-->
<!--#include file="../ACT_inc/UploadClass.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>上传文件</title>
 <style type="text/css">
<!--
body,div,dl,dt,dd,ul,ol,li,h1,h2,h3,h4,h5,h6,pre,   
form,fieldset,input,textarea,p,blockquote,th,td {padding:0; margin:0;}
html, body { height:100%;} /* 同时设置html是为了兼容FF */
body { font-size: 12px; color: black; line-height: 150%; background-color:#fff; text-align: center; height:100%; width:98%;}
.table_list, .table_form, .table_info { margin:0 auto; width:99%; *margin-top:6px; background:#D5EDFD; border:1px solid #99d3fb;}
.table_list caption, .table_form caption, .table_info caption{ border:1px solid #99d3fb; border-bottom-width:0; font-weight:bold; color:#077ac7; background:url(../../user/images/bg_table.jpg) repeat-x 0 0; height:27px; line-height:27px; margin:6px auto 0;font-size:12px; font-family:"宋体"}
h4 { border:1px solid #069; border-width:0 1px 1px 0; margin-top:0; font-size:14px; text-align:left; background:url(../../user/images/bg_admin.jpg) repeat-x 0 -58px; height:28px; line-height:28px; color:#fff;position:absolute; top:0; bottom:0; left:0; z-index:500; width:219px;}
h4 span{ background:url(../../user/images/bg_arrow.jpg) no-repeat 5px -1px; padding-left:30px;}
h4 img{ cursor:pointer;}
.table_form, .table_info {}
.table_form tr,.table_info tr,.table_list tr{ background-color:#fff;}
.table_form td, .table_form th, .table_info td,.table_list td  { line-height:150%; padding:4px;font-size:12px; font-family:"宋体";}
.table_form th strong, .table_info th strong { color:#077ac7;}

-->
</style>
 <script language="JavaScript" type="text/javascript" src="../ACT_inc/js/lhgcore/lhgcore.min.js"></script>
 	<script type="text/javascript">
		var DG = frameElement.lhgDG;
   	</script>
</head>
<body>
 
 <% 
  dim UserHS ,A
 	Set UserHS = New ACT_User
	IF Cbool(UserHS.UserLoginChecked)=false then
	  Response.Write "<script>top.location.href ='login.asp' ;</script>"
	  Response.end
	End If	

  If Request("t")="1" Then
	Upfile_Main()
Else
	Main()
End If


Sub Upfile_Main()
	'-----------------------------------------------------------------------------
	'提交验证
	'-----------------------------------------------------------------------------
 
   	Dim Upload,FilePath,FormName,File,F_FileName,fs,instrs,myid,fp
	Call actcms.CreateFolder(ACTCMS.ActSys&"UpFiles/User/"&UserHS.UModeID&"/")
  	fp ="UpFiles/User/"&UserHS.UModeID&"/"

	Call actcms.CreateFolder(ACTCMS.ActSys&fp)
	FilePath = ACTCMS.ActSys&fp
  	Dim UpFile
	set UpFile = New UpLoadClass
  	UpFile.AutoSave = 2
 	UpFile.MaxSize =  UserHS.G_UpfilesSize* 1024
	UpFile.FileType = "jpg/gif/bmp/png"
	UpFile.SavePath = ACTCMS.ActSys&fp
	UpFile.Open() '# 打开对象
 	If UpFile.Save("Filedata",0) Then
		F_FileName=ACTCMS.ActSys&fp&UpFile.Form("Filedata")
 		echo "<div id=""val"">"&F_FileName&"</div>"
 		echo"<script>J('#myface',DG.curDoc).val( J('#val').html() );</script>" 
	    echo "<script>DG.cancel()</script>"
    Else
		Select Case UpFile.Form("Filedata_Err")
		Case -1 : Response.Write "没有文件上传，请返回重新上传"
		Case 1 : Response.Write "文件大小超出限制，请返回重新上传"
		Case 2 : Response.Write "不允许上传的文件类型，请返回重新上传"
		Case 3 : Response.Write "文件大小超出限制并且是不允许上传的文件类型，请返回重新上传</div>"
		Case Else : Response.Write "未知错误，请返回重新上传</div>"
		End Select
 	End If
	Set UpFile = Nothing
End Sub


  
  Sub main()
 %>
<form name="upload" method="post" action="?t=1"  enctype="multipart/form-data">
 <table cellpadding="2" cellspacing="1" class="table_form">
    <caption>图片上传</caption>
  <tr>
     <td>
              <input name="Filedata" type="file" size="15">
 
             <input type="submit" name="dosubmit" value=" 上传 ">
             		 </td>
   </tr>
  
  <tr>
     <td>
  
			 </td>
   </tr>
</table>
</form>
<%End Sub %></body>
</html>