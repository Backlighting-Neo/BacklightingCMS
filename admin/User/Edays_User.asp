<!--#include file="../ACT.Function.asp"-->
<!--#include file="../../ACT_inc/cls_pageview.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>1</title>
<link href="../Images/style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.STYLE1 {font-weight: bold}
-->
</style>
</head>
<body>
	<%	ConnectionDatabase	
	dim MaxPerPage,RS,TotalPut,TotalPages,I,CurrentPage,SQL,ComeUrl
		  MaxPerPage=20
			Response.Write"<table width=""98%""  height=""25"" border=""0""  align=""center""  cellspacing=""1"" cellpadding=""0""  class=""table"">"
			Response.Write " <tr>"
			Response.Write"	<td height=""25"" class=""bg_tr"" align=""center"" "
			Response.Write " <strong>会员有效期日志</strong>"
			Response.Write	" </td>"
			Response.Write " </tr>"
			Response.Write"</TABLE>"
		ComeUrl=Cstr(Request.ServerVariables("HTTP_REFERER"))
		if request("Action")="del" then
		  Dim Param
		  Select Case ChkNumeric(request("deltype"))
		   Case 1
			Param="datediff('d',adddate," & NowString & ")>11"
		   Case 2
			 Param="datediff(d,adddate," & NowString & ")>31"
		   Case 3
			Param="datediff('d',adddate," & NowString & ")>61"
		   Case 4
			Param="datediff('d',adddate," & NowString & ")>91"
		   Case 5
			Param="datediff('d',adddate," & NowString & ")>181"
		   Case 6
			Param="datediff('d',adddate," & NowString & ")>366"
		  End Select
		  If Param<>"" Then Conn.Execute("Delete From Edays_ACT Where 1=1 and  " & Param)
		  Call Alert("已按所给的条件，删除了有效期日志的相关记录！",ComeUrl)
		end if
		%>
<table width="98%" style="MARGIN-TOP: 3px" border="0" align="center" cellspacing="1" cellpadding="0" class="table">
  <tr class="title">
    <td width="80" height="25" align="center" class="bg_tr"><strong> 用户名</strong></td>
    <td width="138" height="25" align="center" class="bg_tr"><strong>操作时间</strong></td>
    <td width="111" align="center" class="bg_tr"><strong>IP地址</strong></td>
    <td width="71" height="25" align="center" class="bg_tr"><strong>增加有效期</strong></td>
    <td width="74" align="center" class="bg_tr"><strong>减少有效期</strong></td>
    <td width="59" height="25" align="center" class="bg_tr"><strong>摘要</strong></td>
    <td width="75" height="25" align="center" class="bg_tr"><strong> 操作员</strong></td>
    <td width="239" align="center" class="bg_tr"><strong>备注</strong></td>
  </tr>
  <%
  CurrentPage	= ChkNumeric(request("page"))
  Set RS=Server.CreateObject("ADODB.RecordSet")
    RS.Open "Select ID,UserID,AddDate,IP,Edays,Flag,UserLog,Descript From Edays_ACT order by ID desc",conn,1,1
	If RS.Eof And RS.Bof Then
	 Response.Write "<tr><td colspan=9 align=center height=25 class='tdbg'>找不到相关记录！</td></tr>"
	Else
       TotalPut=rs.recordcount
					if (TotalPut mod MaxPerPage)=0 then
						TotalPages = TotalPut \ MaxPerPage
					else
						TotalPages = TotalPut \ MaxPerPage + 1
					end if
					if CurrentPage > TotalPages then CurrentPage=TotalPages
					if CurrentPage < 1 then CurrentPage=1
					rs.move (CurrentPage-1)*MaxPerPage
					SQL = rs.GetRows(MaxPerPage)
					rs.Close:set rs=Nothing
					ShowContent
   End If
%>		
</table>
<table border="0" style="margin-top:20px" width="90%" align=center>
<tr><td><strong>特别提醒：</strong>
如果有效期日志记录太多，影响了系统性能，可以删除一定时间段前的记录以加快速度。但可能会带来会员在查看以前收过费的信息时重复收费（这样会引发众多消费纠纷问题），无法通过有效期日志记录来真实分析会员的消费习惯等问题。
</td></tr>
<form action="?action=del" method=post onSubmit="return(confirm('确实要删除有关记录吗？一旦删除这些记录，会出现会员查看原来已经付过费的收费信息时重复收费等问题。请慎重!'))">
<tr><td>删除范围：<input name="deltype" type="radio" value=1>
10天前 
    <input name="deltype" type="radio" value="2" />
    1个月前
    <input name="deltype" type="radio" value="3" />
    2个月前
    <input name="deltype" type="radio" value="4" />
    3个月前
    <input name="deltype" type="radio" value="5" />
    6个月前
    <input name="deltype" type="radio" value="6" checked="checked" />
    1年前
    <input type="submit" value="执行删除" class="ACT_btn"></td></tr>
  </form>
</table>
<%
Sub ShowContent
For i=0 To Ubound(SQL,2)
	%>
  <tr height="25"   onMouseOver=overColor(this) onMouseOut=outColor(this)>
    <td width="80" align="center" class="td_bg" ><%=SQL(1,i)%></td>
    <td align="center" class="td_bg" ><%=SQL(2,i)%></td>
    <td align="center" class="td_bg" ><%=SQL(3,i)%></td>
    <td align="center" class="td_bg" ><%if SQL(5,I)=1 Then Response.Write SQL(4,I) ELSE Response.Write "0"%>天</td>
    <td align="center" class="td_bg" ><%if SQL(5,I)=2 Then Response.Write SQL(4,I) ELSE Response.Write "0"%>天</td>
    <td width="59" align="center" class="td_bg" ><%if SQL(5,I)=1 Then Response.Write "<font color=red>收入</font>" Else Response.Write "支出"%></td>
    <td align="center" class="td_bg" ><%=SQL(6,i)%></td>
	<td class="td_bg"  align="center"><%=SQL(7,i)%></td>
  </tr>
  <%Next
  Response.Write "<tr><td colspan=9 align=right class='td_bg'>"
  Call ShowPagePara(totalPut, MaxPerPage, "", True, "条记录", CurrentPage, "")
  Response.Write "</td></tr>"
End Sub
				
	Public Function ShowPagePara(totalnumber, MaxPerPage, FileName, ShowAllPages, strUnit, CurrentPage, ParamterStr)
		  Dim N, I, PageStr
				Const Btn_First = "第一页"'样式定义
				Const Btn_Prev = "上一页" 
				Const Btn_Next = "下一页" 
				Const Btn_Last = "最后一页" 
				  PageStr = ""
					If totalnumber Mod MaxPerPage = 0 Then
						N = totalnumber \ MaxPerPage
					Else
						N = totalnumber \ MaxPerPage + 1
					End If
				If N > 1 Then
					PageStr = PageStr & ("页次：<font color=red>" & CurrentPage & "</font>/" & N & "页 共有:" & totalnumber & strUnit & " 每页:" & MaxPerPage & strUnit & " ")
					If CurrentPage < 2 Then
						PageStr = PageStr & Btn_First & " " & Btn_Prev & " "
					Else
						PageStr = PageStr & ("<a href=" & FileName & "?page=1" & "&" & ParamterStr & ">" & Btn_First & "</a> <a href=" & FileName & "?page=" & CurrentPage - 1 & "&" & ParamterStr & ">" & Btn_Prev & "</a> ")
					End If
					
					If N - CurrentPage < 1 Then
						PageStr = PageStr & " " & Btn_Next & " " & Btn_Last & " "
					Else
						PageStr = PageStr & (" <a href=" & FileName & "?page=" & (CurrentPage + 1) & "&" & ParamterStr & ">" & Btn_Next & "</a> <a href=" & FileName & "?page=" & N & "&" & ParamterStr & ">" & Btn_Last & "</a> ")
					End If
					If ShowAllPages = True Then
						PageStr = PageStr & ("GO:<select  onChange='location.href=this.value;' style='width:55;' name='select'>")
				   For I = 1 To N
					 If Cint(CurrentPage) = I Then
						PageStr = PageStr & ("<option value=" & FileName & "?page=" & I & "&" & ParamterStr & " selected>NO." & I & "</option>")
					 Else
						 PageStr = PageStr & ("<option value=" & FileName & "?page=" & I & "&" & ParamterStr & ">NO." & I & "</option>")
					 End If
				   Next
				  PageStr = PageStr & "</select>"
				  End If
			 End If
			 ShowPagePara = PageStr
			 response.write ShowPagePara
	End Function
%> 
<script language="javascript">

//CSS背景控制
function overColor(Obj)
{
	var elements=Obj.childNodes;
	for(var i=0;i<elements.length;i++)
	{
		elements[i].className="tdbg1"
		Obj.bgColor="";
	}
	
}
function outColor(Obj)
{
	var elements=Obj.childNodes;
	for(var i=0;i<elements.length;i++)
	{
		elements[i].className="tdbg";
		Obj.bgColor="";
	}
}
</script>