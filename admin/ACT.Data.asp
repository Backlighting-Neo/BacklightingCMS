<!--#include file="ACT.Function.asp"-->
 <!--#include file="actcms.asp"-->
 <html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>数据库操作 By ACTCMS</title>
<link href="Images/style.css" rel="stylesheet" type="text/css">
</head>
<%
 If Not ACTCMS.ChkAdmin() Then '超级管理员检测
	Call Actcms.ACTCMSErr("")
 End If 
Dim ShowErr
Select Case Request.QueryString("Type")
 		Case "Compress"
			Call Compress()
		Case "CompactDatabase"
			Call CompactDatabase()
		Case Else
			Call Main()
End Select
Sub Main()
	 if request.QueryString("Flag") ="Result"  then
			Response.Write ("<body style=""margin:1;"">")
		 Call ExecuteSql
	else
 %>
 <form name="ExecuteForm" method="post" Action="?Action=ExecSql" onSubmit="return CheckForm()">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td class="bg_tr">您现在的位置：系统设置 &gt;&gt; 数据库维护  <a href="#" target="_blank" style="cursor:help;'" class="Help">帮助</a></td>
  </tr>
  <tr>
    <td class="td_bg">首页 ｜ <a href="?Type=Compress">数据库压缩</a>   ｜ SQL语句查询操作</td>
  </tr>
</table>

  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
    <tr>
      <td class="bg_tr">您现在的位置：系统设置 &gt;&gt; 数据库维护 <a href="#" target="_blank" style="cursor:help;'" class="Help">帮助</a></td>
    </tr>
    <tr>
      <td class="td_bg">说明：注：一次只能执行一条Sql语句。如果你对SQL不熟悉，请尽量不要使用。否则一旦出错，将是致命的。<br>
        建议使用查询语句.如：Select Title From Article order by ArticleID desc,尽量不要使用delete,update等命令</td>
    </tr>
    <tr>
      <td class="td_bg"><textarea name="Sql" rows="5" wrap="OFF" style="width:100%;"></textarea></td>
    </tr>
    <tr>
      <td class="td_bg"><iframe id="ExecuteSQLFrame" scrolling="auto" src="ACT.Data.asp?Action=ExecSql&Flag=Result" style="width:100%;height:255" frameborder=1></iframe></td>
    </tr>
    <tr>
      <td class="td_bg">
     <input type="submit" name="submit1" class="ACT_btn" value="立即执行">
      </td>
    </tr>
  </table>
  </form>
<% end if 
End Sub%>
	<%   
	Sub ExecuteSQL()
	  Dim SelectSQLTF,ExecSQLErrorTF,ExeResultNum,ExeResult,FiledObj
		Dim Sql:Sql =request.querystring("Sql")
	    if SQL="" Then Exit Sub
		If Instr(1,lcase(Sql),"delete from log")<>0 then
			response.Write "error"
				Exit Sub
		End If
	    SelectSQLTF = (LCase(Left(Trim(Sql),6)) = "select")
		Conn.Errors.Clear
		On Error Resume Next
		if SelectSQLTF = True then
			  Set ExeResult = Conn.Execute(Sql,ExeResultNum)
		else
			  Conn.Execute Sql,ExeResultNum
		end if
         
		If Conn.Errors.Count<>0 Then
			  ExecSQLErrorTF = True
			  Set ExeResult = Conn.Errors
		Else
			  ExecSQLErrorTF = False
		End If
		if ExecSQLErrorTF = True then
		%>
		<table width="100%" cellpadding="0" cellspacing="1" class="table">
		  <tr class="bg_tr"> 
			<td height="20" nowrap> 
			  <div align="center">错误号</div></td>
			<td height="20" nowrap> 
			  <div align="center">来源</div></td>
			<td height="20" nowrap> 
			  <div align="center">描述</div></td>
			<td height="20" nowrap> 
			  <div align="center">帮助</div></td>
			<td height="20" nowrap> 
			  <div align="center">帮助文档</div></td>
		  </tr>
		  <tr height="20"  class="td_bg"> 
			<td nowrap> 
			  <% = Err.Number %> </td>
			<td nowrap> 
			  <% = Err.Description %> </td>
			<td nowrap> 
			  <% = Err.Source %> </td>
			<td nowrap> 
			  <% = Err.Helpcontext %> </td>
			<td nowrap> 
			  <% = Err.HelpFile %> </td>
		  </tr>
		</table>
		<%
		else
		%>
		<table border="0" cellpadding="0" cellspacing="1" class="table">
		  <%
			if SelectSQLTF = True then
		%>
		  <tr>
		<%
				For Each FiledObj In ExeResult.Fields
		%>
			<td class="bg_tr" nowrap height="26"><div align="center">
				<strong><% = FiledObj.name %></strong>
			  </div></td>
		<%
				next
		%>
		  </tr>
		<%
				do while Not ExeResult.Eof
		%>
		  <tr height="20" nowrap class="td_bg" >
		<%
					For Each FiledObj In ExeResult.Fields
		%>
			<td> 
			  <div align="center">
				<%
				 if IsNull(FiledObj.value) then
					Response.Write("&nbsp;")
				 else
					Response.Write(FiledObj.value)
				 end if
				 %>
			  </div></td>
		<%
					next
		%>
		  </tr>
		<%
					ExeResult.MoveNext
				loop
			else
		%>
		  <tr>
			<td  height="26">
		<div align="center">执行结果</div></td>
		  </tr>
		  <tr>
			<td height="20"  class="td_bg">
		<div align="center">
				<% = ExeResultNum & "条纪录被影响"%>
			  </div></td>
		  </tr>
		<%
			end if
		%>
		</table>
		<%
		  end if
		 End Sub
 
Sub act_bak()
%>

<table width="100%" height="1" border="0" align=center cellpadding="5" cellspacing="1" class="table">
		<tr>
			<th height=25 class="bg_tr" style="text-align:center;">
			&nbsp;&nbsp;<B>备份论坛数据</B>( 需要FSO支持，FSO相关帮助请看微软网站 )			</th>
		</tr>
		<form method="post" action="?Type=bakdata">
		<tr>
			<td height=100 class="td_bg">&nbsp;&nbsp;
				当前数据库路径(相对路径)：
				  <input type=text size=50 name=DBpath value="<%=MyDbPath&db%>">
				  <BR>				  &nbsp;&nbsp;
				备份数据库目录(相对路径)：
				<input type=text size=50 name=bkfolder value="../Databackup">				&nbsp;如目录不存在，程序将自动创建<BR>				&nbsp;&nbsp;
				备份数据库名称(填写名称)：
				<input type=text size=50 name=bkDBname value="ActCMS_Backup_<%=date%>.mdb">				&nbsp;如备份目录有该文件，将覆盖，如没有，将自动创建<BR>&nbsp;&nbsp;
<input type=submit class="ACT_btn" value="确定">
<br>
				-----------------------------------------------------------------------------------------<br>&nbsp;&nbsp;在上面填写本程序的数据库路径全名，本程序的默认数据库文件为Data_act\actcms3.mdb，<B>请一定不能用默认名称命名备份数据库</B><br>&nbsp;&nbsp;您可以用这个功能来备份您的法规数据，以保证您的数据安全！<br>&nbsp;&nbsp;注意：所有路径都是相对与程序空间根目录的相对路径 </font>			</td>
		</tr>	
		</form>
</table>
	<%End Sub
	Sub Compress()
%>
<table width="100%" height="1" border="0" align=center cellpadding="5" cellspacing="1" class="table">
<form action="?Type=CompactDatabase" method="post">
<tr>
<td class="td_bg" height=25><b>注意：</b><br>
  输入数据库所在相对路径,并且输入数据库名称 </td>
</tr>
<tr>
<td class="td_bg">压缩数据库：<input name="dbpath" type="text" value="<%=actcms.actsys&db%>" size="50">
&nbsp;
<input type="submit" class="button" value="开始压缩"></td>
</tr>

<form>
</table>
	<%End sub
		 Public Function CompactDatabase()
				 dim dbpath
				dbpath = request("dbpath")
				On Error Resume Next
				Dim strTempFile, fso, jro, ver, strCon, strTo, LCID
				Set fso = Server.CreateObject("Scripting.FileSystemObject")
				strTempFile = DBPath
				strTempFile = Left(strTempFile, InStrRev(strTempFile, "\")) & fso.GetTempName
				Set jro = Server.CreateObject("JRO.JetEngine")
				LCID = Conn.Properties("Locale Identifier").Value
				CloseConn
				strTo = "Provider=Microsoft.Jet.OLEDB.4.0; Locale Identifier=" & LCID & "; Data Source=" & Server.MapPath(strTempFile) & "; Jet OLEDB:Engine Type=" & ver
				jro.CompactDatabase ConnStr, strTo
				CompactDatabase = False
				If Err Then
					fso.DeleteFile Server.MapPath(strTempFile)
				Else
					fso.DeleteFile Server.MapPath(DBPath)
					fso.MoveFile Server.MapPath(strTempFile), Server.MapPath(DBPath)
					If Err Then
						fso.DeleteFile Server.MapPath(strTempFile)
					Else
						CompactDatabase = True
					End If
				End If
				Set jro = Nothing
				Set fso = Nothing
				'重新打开数据库
				ConnectionDatabase()
			   if  CompactDatabase=true then
					ShowErr = "数据库压缩和修复成功"
			   else
				 ShowErr = "操作失败"
			   end if
					Call Actcms.ActErr(ShowErr,"","")
					Response.end
		End Function




Function CheckDir(FolderPath)
dim fso1
	folderpath=Server.MapPath(".")&"\"&folderpath
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    If fso1.FolderExists(FolderPath) then
       '存在
       CheckDir = True
    Else
       '不存在
       CheckDir = False
    End if
    Set fso1 = nothing
End Function

Function MakeNewsDir(foldername)
	dim f,fso1
	 MakeNewsDir = False
	   Set fso1 = CreateObject("Scripting.FileSystemObject")
        Set f = fso1.CreateFolder(foldername)
        MakeNewsDir = True
    Set fso1 = nothing
End Function
%>
	<script language="javascript">
	<!--
	 function CheckForm()
	 {if (document.ExecuteForm.Sql.value=='')
	  {
	  alert('请输入SQL查询语句！');
	  document.ExecuteForm.Sql.focus();
	  return false;
	  }
	  ExecuteSQLFrame.location.href="ACT.Data.asp?Action=ExecSql&Flag=Result&SQL="+document.ExecuteForm.Sql.value;
	  return false;
	  }
	-->
	</script>
	
</BODY>
	</HTML>

  
  

