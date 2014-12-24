<!--#include file="../ACT.Function.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>权限管理</title>
<link href="../Images/style.css" rel="stylesheet" type="text/css">
</head>
<body><%
 If Not ACTCMS.ChkAdmin() Then '超级管理员检测
	Call Actcms.ACTCMSErr("")
 End If 
	Dim Purview,Admin_Name
	Public Function ACTCMS_QX(QXLX,QXID,ACT)
	dim arrTemp,i,L,ACT_ClassID
	arrTemp=split(QXLX,",")
	For I=LBound(arrTemp) To Ubound(arrTemp)
		if InStr(arrTemp(I),QXID) > 0 Then
			ACT_ClassID=split(arrTemp(I),"-")
				If  ACT_ClassID(ACT)>0 then
					ACTCMS_QX=True
				Else
					ACTCMS_QX=False
				End if
		End  If 
	next 
	End Function
	
	Dim AdminID,Rs,HQQXLX
	AdminID = ChkNumeric(ACTCMS.S("AdminID"))
	IF AdminID  = 1 Then
		Call Actcms.alert("非法操作,系统创建者禁止修改!" ,"")
	End If 
	Set RS=ACTCMS.ActExe("Select * from Admin_ACT Where Id = "&AdminID&" order by ID desc")
	if rs.eof then
		Call actcms.Alert("参数错误","")
		response.End
	Else
		HQQXLX=rs("ACTCMS_QXLX")
		Purview=rs("Purview")
		Admin_Name=rs("Admin_Name")
	End  If
	Public Function ACTCMS_HQQX(HQQXID,HQACT)
	If Trim(HQQXLX)="" Then Exit Function 
	On Error Resume Next
	Dim HQarrTemp,HQi,HQL,HQACT_ClassID
	HQarrTemp=split(HQQXLX,",")'
		For HQI=LBound(HQarrTemp) To Ubound(HQarrTemp)'遍历
			if InStr(HQarrTemp(HQI),HQQXID) > 0 Then
				HQACT_ClassID=Split(HQarrTemp(HQI),"-")
				If UBound(HQACT_ClassID)>0 Then 
					If HQACT_ClassID(1)=HQACT Then
						ACTCMS_HQQX=True
						Exit Function
					Else	
						ACTCMS_HQQX=False
					End If 
				End if
			End  If 
		Next 
	End Function
  If ACTCMS.S("Action")= "Save" Then
	  Dim AdminRS,AdminSql,CountM,I,Purviews
	  AdminID = AdminID
	  CountM=ACTCMS.ACTEXE("Select Count(ModeID) From Mode_Act")(0)
	For I=1 To CountM
		Purviews=Purviews& Replace(request.Form("Purview-" & I &""),",","") & ","
	Next
	  Set AdminRS = Server.CreateObject("adodb.recordset")
	  AdminSql = "select * from Admin_ACT Where ID="&AdminID
	  AdminRS.Open AdminSql, Conn, 1, 3
	  AdminRS("ACTCMS_QXLX") = Replace(ACTCMS.S("ACTCMS_QXLX") , " ", "")
	  AdminRS("SuperTF") = ChkNumeric(ACTCMS.S("SuperTF"))
	  AdminRS("Purview") = Purviews
	  AdminRS("ACT_Other") = Replace(ACTCMS.S("ACT_Other") , " ", "")
	  AdminRS.Update
	  AdminRS.Close:Set AdminRS = Nothing	
  	  Call Actcms.ActErr("管理员权限修改成功","ACT.Admin.asp","")
 	End If 
%>
<form name="ClassForm" method="post" action="?Action=Save&AdminID=<%=ChkNumeric(ACTCMS.S("AdminID"))%>">
  <table width="100%" border="0" cellpadding="2" cellspacing="1" class="table">
    <tr>
      <td colspan="2" align="center" class="bg_tr">修改 <%=Admin_Name%> 管理权限</td>
    </tr>
    <tr>
      <td width="15%" align="right">用 户 名：</td>
      <td width="85%"><%= rs("Admin_Name") %> 
	  
	  <input <% IF rs("SuperTF") = 1 Then Response.Write " checked=""checked""" %> type="checkbox" id="SuperTF" name="SuperTF" value="1">
      <LABEL for="SuperTF"><font color=red>超级管理员 拥有所有权限</font></LABEL></td>
    </tr>

<% 
Dim ModeName,ModeID,Rs1
Set Rs1 =ACTCMS.ACTEXE("SELECT ModeID, ModeName,ModeTable, ModeStatus, IFmake,ModeNote  FROM Mode_Act where ModeStatus=0 order by ModeID asc")
	 If Rs1.EOF  Then
	 	Response.Write	"<tr><td colspan=""6"" align=""center"">没有记录</td></tr>"
	 Else
		Dim q:q=1
		Do While Not Rs1.EOF	
		ModeName=Rs1("ModeName")
		ModeID=Rs1("ModeID")
			 %>

    <tr>
      <td align="right"><%= ModeName %>系统权限设置：</td>
      <td>
       <input <% IF ACTCMS.FoundInArr(rs("Purview"),"ACT"&Rs1("ModeID")&"-ACT",",") Then Response.Write " checked" %>  type="radio" name="Purview-<%=q%>"   id="Purview1<%=Rs1("ModeID")%>" onClick="SuperTFT<%=Rs1("ModeID")%>.style.display='none';"  value="<%="ACT"&Rs1("ModeID")%>-ACT">
       <LABEL for="Purview1<%=Rs1("ModeID")%>"> 在<%= ModeName %>系统无任何管理权限(屏蔽)</LABEL><br>
          <input <% IF ACTCMS.FoundInArr(rs("Purview"),"TCJ"&Rs1("ModeID")&"-TCJ",",") Then Response.Write " checked" %> type="radio" name="Purview-<%=q%>"  onclick="SuperTFT<%=Rs1("ModeID")%>.style.display='none';"  id="Purview2<%=Rs1("ModeID")%>" value="<%="TCJ"&Rs1("ModeID")%>-TCJ">
     <LABEL for="Purview2<%=Rs1("ModeID")%>">模型管理员：拥有此模型的所有管理权限(相当于对<%= ModeName %>系统没有任何限制)</LABEL>
      <br>
	 <input <% IF ACTCMS.FoundInArr(rs("Purview"),"DXR"&Rs1("ModeID")&"-DXR",",") Then Response.Write " checked" %> type="radio" name="Purview-<%=q%>" value="<%="DXR"&Rs1("ModeID")%>-DXR"  id="Purview3<%=Rs1("ModeID")%>"   onclick="SuperTFT<%=Rs1("ModeID")%>.style.display='';"    >
	 <LABEL for="Purview3<%=Rs1("ModeID")%>">栏目管理员：需要详细指定每个栏目的权限</LABEL>
</td>
    </tr>
    <tr id=SuperTFT<%=Rs1("ModeID")%> 
	<% IF ACTCMS.FoundInArr(rs("Purview"),"DXR"&Rs1("ModeID")&"-DXR",",")=false Then Response.Write " style=""display:none""" %>
	>
      <td height="25"  colspan="2"><%=ClassPurview(Rs1("ModeID"))&"</table>"%>
	  </td>
    </tr>
	 <% 
		q=q+1
		Rs1.movenext
		Loop
	End if	 %>

    <tr >
   
	<td width="49%" colspan="2"><table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
      <tr>
        <td colspan="4" class="bg_tr">其他权限设置</td>
        </tr>
      <tr>
        <td>
		<input <% IF Instr(rs("ACT_Other"),"editmypassword")>0 Then Response.Write " checked=""checked""" %>  type="checkbox" name="ACT_Other" value="editmypassword">
          修改自己密码</td>
        <td><input <% IF Instr(rs("ACT_Other"),"rzgl")>0 Then Response.Write " checked=""checked""" %>  type="checkbox" name="ACT_Other" value="rzgl">
          日志管理</td>
        <td><input <% IF Instr(rs("ACT_Other"),"sjkwh")>0 Then Response.Write " checked=""checked""" %> type="checkbox" name="ACT_Other" value="sjkwh">
          数据库维护</td>
        <td><input <% IF Instr(rs("ACT_Other"),"zxtgl")>0 Then Response.Write " checked=""checked""" %> type="checkbox" name="ACT_Other" value="zxtgl">
          子系统管理</td>
      </tr>
      <tr>
        <td><input <% IF Instr(rs("ACT_Other"),"bqxt")>0 Then Response.Write " checked=""checked""" %> type="checkbox" name="ACT_Other" value="bqxt">
          标签系统</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td colspan="4" class="bg_tr">插件系统</td>
        </tr>
      <tr>
        <td colspan="4" align="center"><%= CheckPlus("ACT_Other",rs("ACT_Other"),3) %></td>
        </tr>
    </table></td>
	
	
    </tr>
    <tr>
      <td colspan="2" align="center"><input name="Submit2" type="submit" class="ACT_btn" value="保存修改结果">
      <input name="Submit3" type="reset" class="ACT_btn" value=" 取 消 "></td>
    </tr>
  </table>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <%Function ClassPurview(ModeID) %>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
	
		<tr class="bg_tr">
			<td width="5%" height="25" class="bg_tr">
				<div align="center">ID</div>			</td>
			<td width="30%" class="bg_tr">
				<div align="center">栏名称</div>			</td>
		  <td align="center" class="bg_tr">权限类型</td>
	    </tr>
		<%
	Dim Article_rs,Article_rs_1,ChangesLink
	Set Article_rs = Server.CreateObject("adodb.recordset")
	Article_rs.Open "Select Orderid,id,ClassID,ClassName,enname,ClassEName, actlink,dh,tg from Class_Act where Parentid  = '0'  And ModeID="&ModeID&"    Order by Orderid asc,ID asc",Conn,1,1
	if Article_rs.eof then
	   Response.Write"<TR  ><TD colspan=""5""   height=""40"">没有记录</TD></TR>"
	else
	
	do while not Article_rs.eof
			if Article_rs.eof Then exit do
	%>
		<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
			<td height="22" ><div align="center"><%= Article_rs("ID") %></div></td>
			<td height="22" >
				<% 
			Response.Write  "<img src=""../images/-.gif""></img>"&Article_rs("ClassName")
		Set Article_rs_1 = server.CreateObject("adodb.recordset")
		%></td>
				<td  align="center">
<input <%If ACTCMS_HQQX(Article_rs("ClassID"),"1")=true Then response.write "Checked "%> name="ACTCMS_QXLX" type="checkbox" id="ACTCMS_QXLX1<%= Article_rs("ID") %>" value="<%=Article_rs("ClassID")&"-1"%>">
<label for="ACTCMS_QXLX1<%=Article_rs("ID") %>">管理</label>  
<input <%If ACTCMS_HQQX(Article_rs("ClassID"),"2")=true Then response.write "Checked "%> name="ACTCMS_QXLX" type="checkbox" id="ACTCMS_QXLX2<%= Article_rs("ID") %>" value="<%=Article_rs("ClassID")&"-2"%>">
<label for="ACTCMS_QXLX2<%= Article_rs("ID") %>">增加</label>    
<input <%If ACTCMS_HQQX(Article_rs("ClassID"),"3")=true Then response.write "Checked "%> name="ACTCMS_QXLX" type="checkbox" id="ACTCMS_QXLX3<%= Article_rs("ID") %>" value="<%=Article_rs("ClassID")&"-3"%>">
<label for="ACTCMS_QXLX3<%= Article_rs("ID") %>">修改</label>    
<input <%If ACTCMS_HQQX(Article_rs("ClassID"),"4")=true Then response.write "Checked "%> name="ACTCMS_QXLX" type="checkbox" id="ACTCMS_QXLX4<%= Article_rs("ID") %>" value="<%=Article_rs("ClassID")&"-4"%>">
<label for="ACTCMS_QXLX4<%= Article_rs("ID") %>">删除</label>   
<input <%If ACTCMS_HQQX(Article_rs("ClassID"),"5")=true Then response.write "Checked "%> name="ACTCMS_QXLX" type="checkbox" id="ACTCMS_QXLX5<%= Article_rs("ID") %>" value="<%=Article_rs("ClassID")&"-5"%>">
<label for="ACTCMS_QXLX5<%= Article_rs("ID") %>">审核</label></td>
		</tr>
		<%
		Response.Write(GetClassValue(Article_rs("ClassID"),""))
		Article_rs.MoveNext
	Loop
	Article_rs.close
	End if
	End Function 
	Public Function GetClassValue(TypeID,CompatStr)  
		Dim tg,dh,ClassRs,TempStr,CMS_isUrl
		Set ClassRs = Conn.Execute("Select id,dh,orderid,ClassName,tg,ClassEName,enname,ClassID,actlink from Class_Act where ParentID='" & TypeID & "'  order by Orderid desc,id desc" )
		TempStr =CompatStr & "<img src=""../images/L.gif""></img>"
		do while Not ClassRs.Eof
	  		GetClassValue = GetClassValue & "<tr onMouseOver=overColor(this) onMouseOut=outColor(this)>"&Chr(13) & Chr(10)
			GetClassValue = GetClassValue & "<td  align=""center"">"& ClassRs("id")&"</td>" & Chr(13) & Chr(10)
			CMS_isUrl = "&nbsp;["& ModeName &"："&Conn.Execute("Select Count(id) From article_act Where Classid='" & ClassRs("ClassID")&"'")(0)&"]"
			GetClassValue = GetClassValue & "<td >"& TempStr &"<Img src=""../images/-.gif""></img>"&ClassRs("ClassName")&" </td>" & Chr(13) & Chr(10) 
			GetClassValue = GetClassValue & "<td  align=""center"">"& vbCrLf
			GetClassValue = GetClassValue & "<input "
			If ACTCMS_HQQX(ClassRs("ClassID"),"1")=true Then GetClassValue = GetClassValue & " Checked "
			GetClassValue = GetClassValue & "name=""ACTCMS_QXLX"" type=""checkbox"""
			GetClassValue = GetClassValue & " id=""ACTCMS_QXLX1"&ClassRs("ID")&""" value="""&ClassRs("ClassID")&"-1"""
			GetClassValue = GetClassValue & ">"&vbCrLf&"<label for=""ACTCMS_QXLX1"&ClassRs("ID")&""">管理</label>"&vbCrLf
			GetClassValue = GetClassValue & "<input "
			If ACTCMS_HQQX(ClassRs("ClassID"),"2")=true Then GetClassValue = GetClassValue & " Checked "
			GetClassValue = GetClassValue & "name=""ACTCMS_QXLX"" type=""checkbox"""
			GetClassValue = GetClassValue & " id=""ACTCMS_QXLX2"&ClassRs("ID")&""" value="""&ClassRs("ClassID")&"-2"""
			GetClassValue = GetClassValue & ">"&vbCrLf&"<label for=""ACTCMS_QXLX2"&ClassRs("ID")&""">增加</label>"&vbCrLf
			GetClassValue = GetClassValue & "<input "
			If ACTCMS_HQQX(ClassRs("ClassID"),"3")=true Then GetClassValue = GetClassValue & " Checked "
			GetClassValue = GetClassValue & "name=""ACTCMS_QXLX"" type=""checkbox"""
			GetClassValue = GetClassValue & " id=""ACTCMS_QXLX3"&ClassRs("ID")&""" value="""&ClassRs("ClassID")&"-3"""
			GetClassValue = GetClassValue & ">"&vbCrLf&"<label for=""ACTCMS_QXLX3"&ClassRs("ID")&""">修改</label>"& vbCrLf 
			GetClassValue = GetClassValue & "<input "
			If ACTCMS_HQQX(ClassRs("ClassID"),"4")=true Then GetClassValue = GetClassValue & " Checked "
			GetClassValue = GetClassValue & "name=""ACTCMS_QXLX"" type=""checkbox"""
			GetClassValue = GetClassValue & " id=""ACTCMS_QXLX4"&ClassRs("ID")&""" value="""&ClassRs("ClassID")&"-4"""
			GetClassValue = GetClassValue & ">"&vbCrLf&"<label for=""ACTCMS_QXLX4"&ClassRs("ID")&""">删除</label>"& vbCrLf
			GetClassValue = GetClassValue & "<input "
			If ACTCMS_HQQX(ClassRs("ClassID"),"5")=true Then GetClassValue = GetClassValue & " Checked "
			GetClassValue = GetClassValue & "name=""ACTCMS_QXLX"" type=""checkbox"""
			GetClassValue = GetClassValue & " id=""ACTCMS_QXLX5"&ClassRs("ID")&""" value="""&ClassRs("ClassID")&"-5"""
			GetClassValue = GetClassValue & ">"&vbCrLf&"<label for=""ACTCMS_QXLX5"&ClassRs("ID")&""">审核</label>"&vbCrLf
			GetClassValue = GetClassValue & "</td>" & Chr(13) & Chr(10)
			GetClassValue = GetClassValue & "</tr>" & Chr(13) & Chr(10)
			GetClassValue = GetClassValue &GetClassValue(ClassRs("ClassID"),TempStr)
			ClassRs.MoveNext
		loop
		ClassRs.Close
		Set ClassRs = Nothing
	End Function



 	Public Function CheckPlus(OptionName,SelectArr,RowNum)
	   Dim n:n=0
	   Dim ACT_Plus:Set ACT_Plus=Server.CreateObject("Adodb.Recordset")
	   IF RowNum<=0 Then RowNum=3
	   ACT_Plus.Open "Select ID,PlusName,PlusID From Plus_ACT",Conn,1,1
	   CheckPlus="<table width=""100%"" align=""center"" border=""0"">"
	   Do While Not ACT_Plus.Eof
	        CheckPlus=CheckPlus & "<TR>"
	     For N=1 To RowNum
		    CheckPlus=CheckPlus & "<TD WIDTH=""" & CInt(100 / CInt(RowNum)) & "%"">"
			If Instr(SelectArr,ACT_Plus(2))<>0 Then
			 CheckPlus=CheckPlus & "<input id="& OptionName&ACT_Plus(0)&" type=""checkbox"" checked name=""" & OptionName & """ value=""" & ACT_Plus(2) & """><label for="& OptionName&ACT_Plus(0) &">" & ACT_Plus(1) & "</label>&nbsp;&nbsp;&nbsp;&nbsp;"
			Else
			 CheckPlus=CheckPlus & "<input id="& OptionName&ACT_Plus(0)&" type=""checkbox"" name=""" & OptionName & """ value=""" & ACT_Plus(2) & """><label for="& OptionName&ACT_Plus(0) &">" & ACT_Plus(1) & "</label>&nbsp;&nbsp;&nbsp;&nbsp;"
			End IF
				 CheckPlus=CheckPlus & "</TD>"
		 		ACT_Plus.MoveNext
				If ACT_Plus.Eof Then Exit For
		Next
		CheckPlus=CheckPlus & "</TR>"
		If ACT_Plus.Eof Then Exit Do
	   Loop
	   CheckPlus=CheckPlus & "</TABLE>"
	   ACT_Plus.Close:Set ACT_Plus=Nothing
	End Function 

%>
		
	
</table>
</form>
<p></p>
</body>
</html>
<%
set ACTCMS = Nothing
%>
<script language="JavaScript" type="text/JavaScript">
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
