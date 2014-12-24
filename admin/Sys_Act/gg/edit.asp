<!--#include file="../../ACT.Function.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>广告管理</title>
<link href="../../Images/style.css" rel="stylesheet" type="text/css">
<SCRIPT LANGUAGE="JavaScript" src="images/js.js"></SCRIPT>
</head>

<body>
<%
	Dim rs,sql,id,ShowErr,softurl
	id=Request.QueryString("id")
If Request.QueryString("id")<> "" Then

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select ADID,ADType,ADSrc,ADCode,ADHeight,ADWidth,ADLink,ADAlt,ADStopViews,ADStopHits,ADStopDate,ADNote,ADViews,ADHits from [ads] where id=" & id
	rs.Open sql,conn,1,3
	'是否更新
	If Request.Form("ChangeAD") <> "" Then
	rs("ADID")=DangerEncode(Request.Form("ADID"))
	rs("ADType")=DangerEncode(Request.Form("ADType"))
	rs("ADSrc")=DangerEncode(Request.Form("ADSrc"))
	rs("ADCode")=DangerEncode(Request.Form("ADCode"))
	rs("ADHeight")=TRIM(Request.Form("ADHeight"))
	rs("ADWidth")=TRIM(Request.Form("ADWidth"))
	rs("ADLink")=DangerEncode(Request.Form("ADLink"))
	rs("ADAlt")=DangerEncode(Request.Form("ADAlt"))
	rs("ADStopViews")=TRIM(Request.Form("ADStopViews"))
	rs("ADStopHits")=TRIM(Request.Form("ADStopHits"))
	rs("ADStopDate")=TRIM(Request.Form("ADStopDate"))
	rs("ADNote")=DangerEncode(Request.Form("ADNote"))
	If Request.Form("ADRESET")="YES" Then
		rs("ADViews")=0
		rs("ADHits")=0
	End If
	rs.Update
	Call Actcms.ActErr("操作成功","Sys_Act/gg/Index.asp","")
     Response.end
End If
If Err <> 0 Then
			Response.Write "<font color=red size=2>错误:"&Err.Description
			Response.End
End If

'softurl=Request.Servervariables("server_name")&replace(Request.Servervariables("url"),"/edit.asp","")
%>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td class="bg_tr"><strong>广告管理----广告管理首页</strong></td>
  </tr>
  <tr>
    <td ><strong><a href="?">首页</a> ┆ <a href="Index.asp">广告列表 </a>┆<a href="add.asp">添加广告 </a>┆</strong></td>
  </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <form name=form action=edit.asp?id=<%=id%>  method=post onSubmit="return chkinput()">
  <tr align="center">
    <td class="bg_tr" height="22" colspan="2" align="left">您现在的位置：广告设置 &gt;&gt; <font class="bg_tr">修改广告</font></td>
  </tr>
  <tr>
    <td width="100" height="25" align="right" >广告名称：</td>  
    <td width="400" ><INPUT name=ADID type="text" class="input1" value="<%=rs("ADID")%>" size=20 maxlength=20> 不能重名</td>
  </tr>
  <tr>
  <td height="25" align="right" >广告类型：</td>
  <td >
  <select name="ADType" size="1" class="input1" onChange="ChangeType(this.options[this.selectedIndex].value)">
    <option <%If rs("ADType")=1 Then Response.Write"selected"%> value="1">普通显示</option>
    <option <%If rs("ADType")=2 Then Response.Write"selected"%> value="2">满屏浮动显示</option>
    <option <%If rs("ADType")=3 Then Response.Write"selected"%> value="3">上下浮动显示 - 右</option>
    <option <%If rs("ADType")=4 Then Response.Write"selected"%> value="4">上下浮动显示 - 左</option>
    <option <%If rs("ADType")=5 Then Response.Write"selected"%> value="5">全屏幕渐隐消失</option>
    <option <%If rs("ADType")=6 Then Response.Write"selected"%> value="6">普通网页对话框 </option>
    <option <%If rs("ADType")=7 Then Response.Write"selected"%> value="7">可移动透明对话框 </option>
    <option <%If rs("ADType")=8 Then Response.Write"selected"%> value="8">打开新窗口</option>
    <option <%If rs("ADType")=9 Then Response.Write"selected"%> value="9">弹出新窗口</option>
    <option <%If rs("ADType")=10 Then Response.Write"selected"%> value="10">对联式广告</option>
    <option <%If rs("ADType")=11 Then Response.Write"selected"%> value="11">联盟广告</option>
  </select>
  </td>                                                         
  </tr>
  <tr>
    <td height="25" align="right"  id="adsrc_text">广告地址：</td>
	<td >
	<%If rs("ADType")= 6 Then
    	Response.Write ("<textarea rows=3 name=ADSrc cols=27 class=input2>"&server.HTMLencode(rs("ADSrc"))&"</textarea>")
	else%>
	<INPUT name="ADSrc" type="text" class="input1" value="<%=rs("ADSrc")%>" size=40>                                                      
	<%end if%>   
    *<a href="#" onclick=openhelp("ext") title="点击查看帮助">图片或FLASH地址？</a></td>   
  </tr>
  <tr>
    <td height="25" align="right"  id="adsrc_text">广告地址：</td>
	<td >
	<%If rs("ADType")= 11 Then
    	Response.Write ("<textarea rows=10 name=ADCode cols=80 class=input2>"&server.HTMLencode(rs("ADCode"))&"</textarea>")
	else%>
	<INPUT name="ADCode" type="text" class="input1" value="<%=rs("ADCode")%>" size=40>                                                      
	<%end if%>   
    *script代码</td>   
  </tr>
  <tr>
    <td height="25" align="right" >广告规格：</td>
    <td ><INPUT name=ADWidth type="text" class="input1" onKeyPress="return Num();" value="<%=rs("ADWidth")%>" size=8 maxlength=4>  
      × <INPUT name=ADHeight type="text" class="input1" onKeyPress="return Num();" value="<%=rs("ADHeight")%>" size=8 maxlength=4></td>                     
  </tr>
  <tr style="visibility:hide;">
    <td height="25" align="right" >链接地址：</td>
    <td ><INPUT name=ADLink type="text" class="input1" value="<%=rs("ADLink")%>" size=40 maxlength=150></td>                      
  </tr>
  <tr>
    <td height="25" align="right" >提示文字：</td>
    <td ><INPUT name=ADAlt type="text" class="input1" value="<%=rs("ADAlt")%>" size=40 maxlength=50></td>                      
  </tr>
  <tr>
    <td height="25" align="right" >投放限制：</td>
    <td ><INPUT name=ADStopViews type="text" class="input1" onKeyPress="return Num();" value="<%=rs("ADStopViews")%>" size=10 maxlength=10>
      ·<INPUT name=ADStopHits type="text" class="input1" onKeyPress="return Num();" value="<%=rs("ADStopHits")%>" size=10 maxlength=10>
      ·<INPUT name=ADStopDate type="text" class="input1" value="<%=rs("ADStopDate")%>" size=18 maxlength=30> <a href="#" onclick=openhelp("stop") title="点击查看帮助">显示·点击·日期？</a></td>
  </tr>
  <tr>
    <td height="25" align="right" >重新统计：</td>
    <td ><input type="checkbox" name="ADRESET" value="YES"> 重置显示和点击次数</td>                     
  </tr>
  <tr>
    <td height="25" align="right" >简单注释：</td>
    <td ><INPUT name=ADNote type="text" class="input1" value="<%=rs("ADNote")%>" size=60 maxlength=100> 备注不显示在广告中</td>
  </tr>
  <tr>
    <td height="22" colspan="2" align="center" >
	   <input type=submit class="ACT_btn" name="ChangeAD" value="  保存  " />
	   &nbsp;&nbsp; <input type="reset" class="ACT_btn" name="Submit2" value="  重置  ">             </td>
  </tr>
</form>
</table>
<table cellpadding="0" cellspacing="0"><tr><td height=5></td></tr></table>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr class="bg_tr">
    <td width="20%" align="right">js版投放代码，复制到投放位置:</td>
    <td width="80%" align="left">
    <textarea name="S1" cols="80" rows="2"><%="<script src="""&ACTCMS.ACTCMSDM&"plus/gg/js_c/"&rs("ADID")&".js""></script>"%></textarea>
    </td>
  </tr>
  <tr class="bg_tr">
    <td width="20%" align="right">asp版投放代码，复制到投放位置:</td>
    <td width="80%" align="left">
    <textarea name="S2" cols="80" rows="2"><%="<script src="""&ACTCMS.ACTCMSDM&"plus/gg/ad.asp?adid="&rs("ADID")&"""></script>"%></textarea>
    </td>
  </tr>
</table>
<%
Response.Write "<Script language=javascript>ChangeType('"&rs("ADType")&"')</Script>"
    rs.Close
    set rs=nothing
    conn.Close
    set conn=nothing
Else
	Response.Write "没有指明要编辑的ID。"
End If

Rem 过滤可能出错误的符号
Function DangerEncode(fString)
If not isnull(fString) Then
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10), "")
	fString = replace(fString, "'", """")
    fString = Trim(fString)
    DangerEncode = fString
End If
End Function
%>
</body>
</html>
