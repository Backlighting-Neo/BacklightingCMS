<!--#include file="../ACT.Function.asp"-->
<!--#include file="../actcms.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>1</title>
<link href="../Images/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
 If Not ACTCMS.ChkAdmin() Then '超级管理员检测
	Call Actcms.ACTCMSErr("")
 End If 
Dim Rs1,path,tempfile,Action,a
tempfile=request("File")
path=request("Dir")
A=request("A")
Action=request("Action")
Select Case  Action
	Case "Add"
		Call Add()
	Case "Addsave"
		Call Addsave()
	Case "Save"
		Call Save()
	Case Else
		Call AddEdit()
End Select 

	Function save()
	Dim Content,FileFSO
		Content=request("content")
		Path=request("Path")
		 If lcase(Right(Path,4))<>"html" And lcase(Right(Path,3))<>"htm" Then Call ACTCMS.Alert("文件名的扩展名必须是html或htm","")
		 Set FileFSO = Server.CreateObject("ADODB.Stream")
			With FileFSO
			.Type = 2
			.Mode = 3
			.Open
			.Charset = "utf-8"
			.Position = FileFSO.Size
			.WriteText  Content
			.SaveToFile Server.MapPath(Path),2
			.Close
			End With
		Set FileFSO = Nothing
		Response.Write ("<script>parent.frame.cols=""180,*"";</script>")
		Call ACTCMS.Alert("模板修改成功!","TempLate.asp"):Exit Function
	End Function 

	Function Addsave()
	Dim Content,FileFSO,File,PhysicalPath,fs
		Content=request("content")
		File=request("File")
		Path=request("Path")
		 If lcase(Right(File,4))<>"html" And lcase(Right(File,3))<>"htm" Then Call ACTCMS.Alert("文件名的扩展名必须是html或htm","")
		 Set fs = Server.CreateObject("scripting.filesystemobject")
		  PhysicalPath=Server.Mappath(Path&File)
	      if fs.FileExists(PhysicalPath) = False Then
			 Set FileFSO = Server.CreateObject("ADODB.Stream")
			With FileFSO
			.Type = 2
			.Mode = 3
			.Open
			.Charset = "utf-8"
			.Position = FileFSO.Size
			.WriteText  Content
			.SaveToFile Server.MapPath(Path&File),2
			.Close
			End With
			Set FileFSO = Nothing
			Response.Write ("<script>parent.frame.cols=""180,*"";</script>")
			Call ACTCMS.Alert("模板增加成功!","TempLate.asp"):Exit Function
		  Else 
			  Call ACTCMS.Alert("文件名已经存在,请另取一个名称!","")
		  End  If 
		
	
	
	End Function 


	Function  LoadTemplate(TempString) 
		On Error Resume Next
		Dim  Str,A_W
		set A_W=server.CreateObject("adodb.Stream")
		A_W.Type=2 
		A_W.mode=3 
		A_W.charset="utf-8"
		A_W.open
		A_W.loadfromfile server.MapPath(TempString)
		If Err.Number<>0 Then Err.Clear:LoadTemplate="模板没有找到 <br> by BackLighting Software":Exit Function
		Str=A_W.readtext
		A_W.Close
		Set  A_W=nothing
		LoadTemplate=Str
	End  function

Function AddEdit()
%>
<script>parent.frame.cols="0,*";</script>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
    <form id="form" name="form" method="post" action="?Action=Save&Path=<%=path&"/"&tempfile%>">
<tr>
      <td colspan="2" class="bg_tr">您现在的位置：文章中心 &gt;&gt; 模板修改</td>
    </tr>
    <tr>
      <td colspan="2">文件名：<STRONG><%=tempfile%></STRONG> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	  完整路径 ：<STRONG><font size=5 color=red><%=path&"/"&tempfile%></font></STRONG>&nbsp;&nbsp;&nbsp;&nbsp;</td>
    </tr>
    <tr>
      <td width="13%">
	  <SELECT id=Lable_Name name=Lable_Name> 
              <OPTION value="" selected>==选择标签==</OPTION> 
                      <%
				Set Rs1=server.CreateObject("adodb.recordset") 
			  set rs1 = Conn.execute("select LabelName From Label_Act order by AddDate desc")
			  do while not rs1.eof
			  	Response.Write "<option value=""" & RS1(0) & """>" & RS1(0) & "</option>"
			  rs1.movenext
			  loop
			  rs1.close:set rs1 = nothing
			  %>
	    </SELECT> </td>
      <td width="87%">
	<INPUT  class="act_btn" onClick="if(this.form.Lable_Name.options[this.form.Lable_Name.selectedIndex].value==''){return false;}else{insert(this.form.Lable_Name.options[this.form.Lable_Name.selectedIndex].value);}" type=button value=" 插入标签 "></td>
    </tr>
    <tr>
      <td colspan="2">
	<div>
	<textarea name="content"  rows="28" style="width:100%"><%=server.HTMLEncode(LoadTemplate(path&"/"&tempfile))  %></textarea>	
						  </div>
     
	  </td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><input type=button onclick=CheckForm() class="act_btn" name=Submit value="  保存  " />
      <input type="reset" name="Submit2"  class="act_btn" value="  重置  " /></td></td>
    </tr></form>
  </table>
  <% End Function 



Function Add()
  Dim tcontent
  tcontent = "<html>" & vbcrlf &"<head>" & vbcrlf & "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">" & vbcrlf
  tcontent = tcontent & "<title>默认文档-By www.actcms.com</title>" & vbcrlf & "</head>" & vbcrlf & "<body>" & vbcrlf &"Hello Word!"& vbcrlf & "</body>" & vbcrlf & "</html>"
%>
<script>parent.frame.cols="0,*";</script>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
    <form id="form" name="form" method="post" action="?Action=Addsave&Path=<%=path&"/"&tempfile%>">
<tr>
      <td colspan="2" class="bg_tr">您现在的位置：文章中心 &gt;&gt; 模板添加</td>
    </tr>
    <tr>
      <td colspan="2"> 
	   文件名：<input name="File" type="text"  size="20" maxlength="255" id="File" value="Untitled.html"> 路径：<STRONG><font size=5 color=red><%=path%></font></STRONG></td>
    </tr>
    <tr>
      <td width="13%">
	  <SELECT id=Lable_Name name=Lable_Name> 
              <OPTION value="" selected>==选择标签==</OPTION> 
                      <%
				Set Rs1=server.CreateObject("adodb.recordset") 
			  set rs1 = Conn.execute("select LabelName From Label_Act order by AddDate desc")
			  do while not rs1.eof
			  	Response.Write "<option value=""" & RS1(0) & """>" & RS1(0) & "</option>"
			  rs1.movenext
			  loop
			  rs1.close:set rs1 = nothing
			  %>
	    </SELECT> </td>
      <td width="87%">
	<INPUT  class="act_btn" onClick="if(this.form.Lable_Name.options[this.form.Lable_Name.selectedIndex].value==''){return false;}else{insert(this.form.Lable_Name.options[this.form.Lable_Name.selectedIndex].value);}" type=button value=" 插入标签 "></td>
    </tr>
    <tr>
      <td colspan="2">
	<div>
	<textarea name="content"  rows="28" style="width:100%"><%=tcontent%></textarea>	
						  </div>
     
	  </td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><input type=button onclick=CheckForm() class="act_btn" name=Submit value="  保存  " />
      <input type="reset" name="Submit2"  class="act_btn" value="  重置  " /></td></td>
    </tr></form>
  </table>
  <% End Function %>
  <script language="javascript">
function insert(returnValue_lable)
{
	obj=document.getElementById("content");
	obj.focus();
	if(document.selection==null)
	{
		var iStart = obj.selectionStart
		var iEnd = obj.selectionEnd;
		obj.value = obj.value.substring(0, iEnd) +returnValue_lable+ obj.value.substring(iEnd, obj.value.length);
	}else
	{
		var range = document.selection.createRange();
		range.text=returnValue_lable;
	}
}

  function CheckForm()
{
  		var form=document.form;
	    form.Submit.value="正在提交数据,请稍等...";
		form.Submit.disabled=true;	
		form.Submit2.disabled=true;	
	    form.submit();
        return true;
	}
	</script>

</body>
</html>