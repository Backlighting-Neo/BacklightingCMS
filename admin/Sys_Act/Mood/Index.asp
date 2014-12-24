<!--#include file="../../ACT.Function.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>心情指数管理 By ACTCMS.COM</title>
<link href="../../Images/style.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="../../../act_inc/js/time/WdatePicker.js"></script>
 <link href="../../../act_inc/js/time/skin/default.css" rel="stylesheet" type="text/css">

</head>
<body>
<% 

Response.Expires = 0
Response.Expiresabsolute = Now() - 1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
Dim Action,ID,Rs,UploadSize,UploadPath,UnlockTime,FormCode,StartTime,UserGroupList,SubmitNum,i,Status ,EndTime
			 Dim  TitleContent,PicContent
ID = ChkNumeric(Request("ID"))
	If Not ACTCMS.ACTCMS_QXYZ(0,"mood_ACT","") Then   Call Actcms.Alert("对不起，你没有操作权限！","") 
	Action = Request.QueryString("A")
	 if ID=0 or ID="" Then ID=1
	Select Case Action
		   Case "AddSave","ESave"
		   		Call AddSave()
			Case "Add","E"
				Call AddEdit()
			Case Else
				Call Main()
	End Select
	
	
	IF Action = "Del" Then
			ACTCMS.ACTEXE("Delete From Mood_Plus_ACT Where ID=" & ID)		
			Call Actcms.ActErr("删除心情成功","Sys_Act/Mood/Index.asp","")
 	End IF
	
	
	Sub AddSave()
		 Dim Title,ModeTable,sql,ChannelRS,ChannelRSSql,ModeNote
		 Title = ACTCMS.S("Title")
 		 UnlockTime = ChkNumeric(ACTCMS.S("UnlockTime"))
		 StartTime = ACTCMS.S("StartTime")
		 EndTime = ACTCMS.S("EndTime")
 		 SubmitNum = ChkNumeric(ACTCMS.S("SubmitNum"))
		 Status = ChkNumeric(ACTCMS.S("Status"))
 		 IF ACTCMS.S("Title") = "" Then
		 	Call ACTCMS.Alert("心情名称为空!",""):Exit Sub
		 End if
		 IF Not IsDate(StartTime) Then StartTime=Now
		 IF Not IsDate(EndTime) Then EndTime=Now+10

		 For i=0 To 14

			TitleContent=TitleContent&actcms.s("Name"&i)&"@&@"

		 Next 


		 For i=0 To 14

			PicContent=PicContent&actcms.s("Pic"&i)&"@&@"

		 Next 

		 
 		 if Action="AddSave" Then
 			 Set ChannelRS = Server.CreateObject("adodb.recordset")
			  ChannelRSSql = "select * from Mood_Plus_ACT"
			  ChannelRS.Open ChannelRSSql, Conn, 1, 3
			  ChannelRS.AddNew
 		 Else
		 	If Not ACTCMS.ACTEXE("SELECT Title FROM Mood_Plus_ACT Where ID <>" & ID & " AND  Title='" & Title & "' order by ID desc").eof Then
				Call ACTCMS.Alert("系统已存在该心情名称-!",""):Exit Sub
			 End if	
			 Set ChannelRS = Server.CreateObject("adodb.recordset")
			  ChannelRSSql = "select * from Mood_Plus_ACT Where ID=" &ID
			  ChannelRS.Open ChannelRSSql, Conn, 1, 3
			  if ChannelRS.eof then Call ACTCMS.Alert("错误!",""):Exit Sub
		 End if 
		 	  ChannelRS("StartTime") = StartTime
		 	  ChannelRS("EndTime") = EndTime
		 	  ChannelRS("SubmitNum") = SubmitNum
		 	  ChannelRS("UnlockTime") = UnlockTime
			  ChannelRS("Title") = Title
			  ChannelRS("TitleContent") = TitleContent
			  ChannelRS("PicContent") = PicContent
 			  ChannelRS("Status") = Status
		  ChannelRS.Update
		  ChannelRS.Close:Set ChannelRS = Nothing	
		  Call Actcms.ActErr("修改成功","Sys_Act/Mood/Index.asp","")
 	End Sub
	Sub Main()
	%>	
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td class="bg_tr">您现在的位置：心情指数管理 &gt;&gt; 浏览</td>
  </tr>
  <tr>
    <td>当前心情： <a href="?A=Add"><b>添加心情指数</b></a> </td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td align="center" class="bg_tr">心情编号</td>
    <td align="center" class="bg_tr">心情名称</td>
    <td align="center" class="bg_tr">是否开始</td>
    <td align="center" class="bg_tr">参与人数</td>
	<td align="center" class="bg_tr" >状态</td>
	<td width="50%" align="center" class="bg_tr" nowrap>管理操作</td>
  </tr>
<% 
	  Set Rs =ACTCMS.ACTEXE("SELECT * FROM Mood_Plus_ACT order by ID desc")
	 If Rs.EOF  Then
	 	Response.Write	"<tr><td colspan=""7"" align=""center"">没有记录</td></tr>"
	 Else
		Do While Not Rs.EOF	
			 %>

  <tr  onMouseOver=overColor(this) onMouseOut=outColor(this)>
    <td align="center"><%= Rs("ID") %></td>
    <td align="center"><%= Rs("Title") %></td>
    <td align="center">
	<%If now >Rs("StartTime") Then 
		response.write "<font color=green title=结束日期是"& Rs("EndTime")&">已经开始</a>"
	Else
		response.write "<font color=red title=开始日期是"& Rs("StartTime")&">还没有开始</font>"
	End if%>
	</td>
	<td align="center"><%=actcms.actexe("select Count(id) from Mood_List_ACT where MDID="&Rs("ID")&"")(0) %></td>
	<td align="center"><% IF Rs("Status") = 0 Then Response.Write "<font color=green>正常</font>" else  Response.Write "<font color=red>禁用</font>" %></td>
	<td align="center" >
  	┆ <a href="?A=E&ID=<%=Rs("ID")  %>" >修改</a> ┆ 
	<a href="?A=Del&ID=<%=Rs("ID")  %>"  onClick="{if(confirm('确定删除该心情吗?')){return true;}return false;}">删除</a></td>
  </tr>
  <% 
		
		Rs.movenext
		Loop
	End if	 %>
</table>	
	
	
<% 	
 
	End Sub
	Sub AddEdit()
	Dim FileFolder,ModeTable,Title,FormCode,ModeNote,A,Template,EndTime,StartTime
	if Action="Add" Then
	A="AddSave"
	StartTime=date()
	EndTime=date()+10
  	Status=0
	SubmitNum=1
	TitleContent=Split("支持@&@高兴@&@震惊@&@愤怒@&@无聊@&@无奈@&@谎言@&@枪稿@&@不解@&@标题党@&@@&@@&@@&@@&@@&@","@&@")
	PicContent=Split("images/Plus/xq1.gif@&@images/Plus/xq2.gif@&@images/Plus/xq3.gif@&@images/Plus/xq4.gif@&@images/Plus/xq5.gif@&@images/Plus/xq6.gif@&@images/Plus/xq7.gif@&@images/Plus/xq8.gif@&@images/Plus/xq9.gif@&@images/Plus/xq10.gif@&@images/Plus/xq11.gif@&@images/Plus/xq12.gif@&@images/Plus/xq13.gif@&@images/Plus/xq14.gif@&@images/Plus/xq15.gif@&@","@&@")
	Else
	Set Rs=server.CreateObject("adodb.recordset") 
	Rs.OPen "Select * from Mood_Plus_ACT Where ID = "&ID&" order by ID desc",Conn,1,1
		  StartTime= Rs("StartTime")
		  EndTime = Rs("EndTime")
		  SubmitNum = Rs("SubmitNum")
		  UnlockTime = Rs("UnlockTime")
		  Title = Rs("Title")
		  TitleContent = Split(Rs("TitleContent"),"@&@")
		  PicContent = Split(Rs("PicContent"),"@&@")
		  Status = Rs("Status")
 	A="ESave"
	end if
  %>
<form id="form1" name="form1" method="post" action="?A=<%= A %>&ID=<%= Request.QueryString("ID") %>">

  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="bg_tr">您现在的位置：<a href="?">心情指数管理</a> &gt;&gt; 添加/修改 </td>
    </tr>
    <tr>
      <td width="22%" align="right" class="td_bg">项目状态&nbsp;&nbsp;</td>
      <td width="78%" class="td_bg">
	  <input <% IF Status = 0 Then Response.Write "Checked" %> id="Status1" type="radio" name="Status" value="0" />
     <label for="Status1"> 正常 </label>
      <input <% IF Status = 1 Then Response.Write "Checked" %>  id="Status2" type="radio" name="Status" value="1" /><label for="Status2">关闭</label>      </td>
    </tr>

	<tr>
      <td height="25" align="right">项目名称：&nbsp;&nbsp;</td>
      <td height="25"><input name="Title"   class="Ainput"  type="text" id="Title" value="<%=Title %>" /></td>
    </tr>
    <tr>
      <td height="25" align="right">选项：&nbsp;&nbsp;</td>
      <td height="25">

<%For i=0 To 14

If i<9 Then echo "0"

%>


<%=i+1%>、名称
<input type="text" value="<%=TitleContent(i)%>" name="Name<%=i%>" class="Ainput">
图片地址
<input type="text" name="Pic<%=i%>" value="<%=PicContent(i)%>" class="Ainput">
<br>
<%Next %>

        </td>
    </tr>

    
 

	   <tr>
      <td height="25" align="right">启用时间限制：&nbsp;&nbsp;</td>
      <td height="25">
	  <input onClick=time(1)     <% IF UnlockTime = 0 Then Response.Write "Checked" %> id="UnlockTime1" type="radio" name="UnlockTime" value="0">
        <label for="UnlockTime1">启用</label>
      <input  onClick=time(2)    <% IF UnlockTime = 1 Then Response.Write "Checked" %> id="UnlockTime2"  type="radio" name="UnlockTime" value="1"> 
     <label for="UnlockTime2">不启用</label></td>
    </tr>
	
	
	    <tr id="times1"
	<%If UnlockTime<>1 Then response.write "style=""DISPLAY: none"""%>
	>
      <td height="25" align="right">开始时间：&nbsp;&nbsp;</td>
      <td height="25"><input  class="Ainput"  size=30 name="StartTime" type="text" id="StartTime" value="<%= StartTime %>"  onClick="WdatePicker();" />
	   </td>
    </tr>

	
	    <tr id="times2"
	<%If UnlockTime<>1 Then response.write "style=""DISPLAY: none"""%>
	>
      <td height="25" align="right">结束时间：&nbsp;&nbsp;</td>
      <td height="25"><input  class="Ainput"   size=30  name="EndTime" type="text" id="EndTime" value="<%= EndTime %>"  onClick="WdatePicker();"/>
	   </td>
    </tr>


	

    <tr>
      <td height="25" align="right">每个用户只允许提交一次：&nbsp;&nbsp;</td>
      <td height="25">
	  <input  <% IF SubmitNum = 1 Then Response.Write "Checked" %> id="SubmitNum1" type="radio" name="SubmitNum" value="1">
        <label for="SubmitNum1">是</label>
      <input  <% IF SubmitNum = 0 Then Response.Write "Checked" %> id="SubmitNum2"  type="radio" name="SubmitNum" value="0"> 
     <label for="SubmitNum2">否</label></td>
    </tr>


	 
 

 <tr>
      <td align="right">&nbsp;</td>
      <td><input type=button onclick=CheckForm() class="ACT_btn"  name=Submit1 value="  保存  " />
      <input type="reset" name="Submit2" class="ACT_btn" value="  重置  " /></td>
    </tr>
  </table>
</form><br>
<script language="javascript">time(<%=""""&UnlockTime&""""%>);</script>

<% end sub  %>

<script language="JavaScript" type="text/javascript">

function overColor(Obj)
{
	var elements=Obj.childNodes;
	for(var i=0;i<elements.length;i++)
	{
		elements[i].className="tdbg1"
		Obj.bgColor="";
	}
	
}
		
function time(n){
	if (n == 2){
		times1.style.display='none';
		times2.style.display='none';
	}
	else{
		times1.style.display='';
		times2.style.display='';
	}
} 



function WinOpenDialog(url,w,h)
{
    var feature = "dialogWidth:"+w+"px;dialogHeight:"+h+"px;center:yes;status:no;help:no";
    showModalDialog(url,window,feature);
}function outColor(Obj)
{
	var elements=Obj.childNodes;
	for(var i=0;i<elements.length;i++)
	{
		elements[i].className="tdbg";
		Obj.bgColor="";
	}
}
function CheckForm()
{ var form=document.form1;
	
	 if (form.Title.value=='')
		{ alert("请输入项目名称!");   
		  form.Title.focus();    
		   return false;
		} 
	    form.Submit1.value="正在提交数据,请稍等...";
		form.Submit1.disabled=true;
		form.Submit2.disabled=true;		
	    form.submit();
        return true;
	}
</script>

</body>
</html>
