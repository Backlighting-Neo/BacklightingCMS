<!--#include file="../ACT.Function.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../Images/style.css" rel="stylesheet" type="text/css">

<title>模型导入导出</title>
<%
	Dim a,Rs
	If Not ACTCMS.ChkAdmin() Then  Call Actcms.Alert("对不起，您没有操作权限！","")
	a=request("A")
	Select Case a
		case "dc"
			call dc()
		case "2"
			call dr()
		case "3"
			call drbc()
		case "dcbc"
			call dcbc()
		Case Else 
			Call main()
	End select

	sub drbc()
		dim ModeName,M_ACT,L_ACT,MX_ACT,i,TCJ_ACT,M_Rs,M_RsSql,Sql,ModeID,TableName
		ModeName= request("ModeName")
		M_ACT = split(ModeName,"|*|")
 		if ubound(M_ACT)<>1 then call actcms.Alert("提交的模型配置参数不正确","")
		MX_ACT = Split(M_ACT(0),"$|$")'模型配置
		TCJ_ACT = Split(M_ACT(1),"|||")
		if ubound(MX_ACT)<>18 then call actcms.Alert("提交的模型配置参数不正确","")
		
		'response.Write MX_ACT(22)'目录
		'response.Write MX_ACT(2)'表
		'response.Write MX_ACT(0)'模型名称
		
			 IF MX_ACT(2) = "" Then
				Call ACTCMS.Alert("数据表为空!",""):Exit Sub
			 End if
			 If Not ACTCMS.ACTEXE("SELECT ModeName FROM Mode_Act Where ModeName='" & MX_ACT(0) & "' order by ModeID desc").eof Then
				Call ACTCMS.Alert("系统已存在该模型名称!",""):Exit Sub
			 End if	

			 If Not ACTCMS.ACTEXE("SELECT ModeTable FROM Mode_Act Where ModeTable='" & MX_ACT(2) & "' order by ModeID desc").eof Then
				Call ACTCMS.Alert("系统已存在该数据表!",""):Exit Sub
			 End if	
 			 Set M_Rs = Server.CreateObject("adodb.recordset")
			  M_RsSql = "select * from Mode_Act"
			  M_Rs.Open M_RsSql, Conn, 1, 3
			  M_Rs.AddNew
			  M_Rs("ModeName") = MX_ACT(0)
			  M_Rs("IFmake") = MX_ACT(1)
		 	  M_Rs("ModeTable") = MX_ACT(2)
			  TableName= MX_ACT(2)
			  M_Rs("AutoPage") = MX_ACT(3)
			  M_Rs("ProjectUnit") = MX_ACT(4)
			  M_Rs("UpFilesDir") = MX_ACT(5)
 			  M_Rs("ContentExtension") = MX_ACT(6)
			  M_Rs("ModeStatus") = MX_ACT(7)
			  M_Rs("RefreshFlag") = MX_ACT(8)
			  M_Rs("RecyleIF") = MX_ACT(9)
			  M_Rs("ACT_DiY") ="§0§0-1-0-1§0§actcms§0§§0§§0§§0§Simple§§§0§0§0§1§0§1§0§"&ACTCMS.ActCMS_Sys(3)&"templets/article/ClassIndex.htm§"&ACTCMS.ActCMS_Sys(3)&"templets/article/Class.Htm§"&ACTCMS.ActCMS_Sys(3)&"templets/article/Content.Htm§"
			  M_Rs("MakeFolderDir") = MX_ACT(11)
			  M_Rs("WriteComment") = MX_ACT(12)
			  M_Rs("CommentCode") = MX_ACT(13)
			  M_Rs("Commentsize") = MX_ACT(14)
			  M_Rs("ModeNote") = MX_ACT(15)
			  M_Rs("CommentTemp") = MX_ACT(16)
			  M_Rs("adminmb") = MX_ACT(17)
			  M_Rs("usermb") = MX_ACT(18)
 
			  M_Rs.Update
			  ModeID=M_Rs("ModeID") 
			  M_Rs.Close:Set M_Rs = Nothing			
				 Sql="CREATE TABLE "&MX_ACT(2)&" ([ID] int IDENTITY (1, 1) NOT NULL CONSTRAINT PrimaryKey PRIMARY KEY,"&_
				"ClassID varchar(20),"&_
				"Title varchar(200),"&_
				"IntactTitle varchar(250),"&_
				"ActLink tinyint,"&_
 				"Intro text,"&_
				"Content text,"&_
				"Hits int Default 0,"&_
				"rev tinyint Default 0,"&_
				"ChargeType tinyint Default 0,"&_
				"InfoPurview tinyint Default 0,"&_
				"arrGroupID varchar(250),"&_
				"ReadPoint  int Default 0,"&_
				"PitchTime  int Default 0,"&_
				"ReadTimes  int Default 0,"&_
				"DividePercent  int Default 0,"&_
				"KeyWords varchar(100),"&_
 				"CopyFrom varchar(250),"&_
				"UpdateTime datetime,"&_
				"TemplateUrl varchar(100),"&_
				"FileName varchar(200),"&_
				"isAccept tinyint,"&_
				"delif tinyint Default 0,"&_
				"UserID  int Default 0,"&_
				"ArticleInput varchar(250),"&_
				"Author varchar(250),"&_
				"Slide tinyint Default 0,"&_
				"PicUrl varchar(200),"&_
				"Ismake tinyint,"&_
				"Digg int Default 0,"&_
				"down int Default 0,"&_
				"ATT SmallInt Default 0,"&_
				"IStop tinyint Default 0"&_
				")"
			ACTCMS.ACTEXE(sql)
			Application.Contents.RemoveAll
			 
			 
			 
			 
		dim FieldType,Title,IsNotNull,ISType,Type_Default,Description,FieldName,ColumnType,OrderID
		Dim width,height,Content,Type_Type,FieldRS,FieldSql
		for i=0 to ubound(TCJ_ACT)-1
		L_ACT = Split(TCJ_ACT(i),"@")
		if ubound(L_ACT)<>11 then call actcms.Alert("提交的模型配置参数不正确","")
		IF ACTCMS.Chkchars(L_ACT(0)) = False  Then
			Call Actcms.ActErr("英文名称只能为英文、数字及下划线","Mode/ACT.ListM.ASP?A=L&ModeID="&ModeID,"")
		End if
			FieldName=L_ACT(0)
			Title=L_ACT(1)
			IsNotNull=ChkNumeric(L_ACT(2))
			OrderID=L_ACT(3)
			Description=L_ACT(4)
			FieldType=L_ACT(5)
			Type_Default=L_ACT(6)
			width=L_ACT(7)
			height=L_ACT(8)
			Content	=L_ACT(9)
			Type_Type=L_ACT(10)
			ISType=ChkNumeric(L_ACT(11))
		
		Select Case FieldType
			Case "TextType"'单行文本
				ColumnType="varchar(255)"
			Case "MultipleTextType"'多行文本(不支持Html
				 ColumnType="text"
			Case "MultipleHtmlType"'多行文本(支持Html)
				ColumnType="text"
			Case "RadioType"'单选项
				ColumnType="varchar(255)"
			Case "ListBoxType"'多选项
				ColumnType="text"
			Case "NumberType"'数字
				ColumnType="int"'
		   Case "DateType"
				 ColumnType="datetime"'Response.write "日期时间"
		   Case "NumberType"
				ColumnType="int"'Response.write "数字"
		   Case else
		     ColumnType="varchar(255)"
		End Select 
		
		 Set FieldRS = Server.CreateObject("ADODB.RECORDSET")
		 FieldSql = "Select * From [Table_ACT] Where FieldName='" & FieldName & "' And  actcms=1 and  ModeID=" & ModeID
		 FieldRS.Open FieldSql, conn, 3, 3
		 If FieldRS.EOF And FieldRS.BOF Then
			FieldRS.AddNew
			FieldRS("FieldName")=L_ACT(0)
			FieldRS("Title")=L_ACT(1)
			FieldRS("IsNotNull")=ChkNumeric(L_ACT(2))
			FieldRS("OrderID")=L_ACT(3)
			FieldRS("Description")=L_ACT(4)
			FieldRS("FieldType")=L_ACT(5)
			FieldRS("Type_Default")=L_ACT(6)
			FieldRS("width")=ChkNumeric(L_ACT(7))
			FieldRS("height")=ChkNumeric(L_ACT(8))
			FieldRS("Content")=L_ACT(9)
			FieldRS("Type_Type")=ChkNumeric(L_ACT(10))
			FieldRS("ISType")=ChkNumeric(L_ACT(11))
			FieldRS("ModeID") = ModeID
		  FieldRS.Update
		 Conn.Execute("Alter Table "&TableName&" Add "&L_ACT(0)&" "&ColumnType&"")
		 
		 Else
		   Call ACTCMS.Alert("数据库中已存在该字段名称!", "")
		   Exit Sub
		 End If
		next 
		Application.Contents.RemoveAll
		Call Actcms.ActErr("导入模型成功","Mode/ACT.MX.asp","")
	end sub
	sub dcbc()
	Dim rs,rs1,M_Act,T_ACT,modeid,ModePath
	modeid=ChkNumeric(ACTCMS.S("modeid"))
    ModePath=ACTCMS.S("ModePath")
 	Set Rs=server.CreateObject("adodb.recordset") 
	Set Rs1=server.CreateObject("adodb.recordset") 
	Rs.OPen "Select * from Mode_Act where ModeID="&modeid&" ",Conn,1,1
	If Not rs.eof Then 
				M_Act=M_Act&Rs("ModeName")&"$|$"
				M_Act=M_Act&Rs("IFmake")&"$|$"
				M_Act=M_Act&Rs("ModeTable")&"$|$"
				M_Act=M_Act&Rs("AutoPage")&"$|$"
				M_Act=M_Act&Rs("ProjectUnit")&"$|$"
				M_Act=M_Act&Rs("UpFilesDir")&"$|$"
				M_Act=M_Act&Rs("ContentExtension")&"$|$"
				M_Act=M_Act&Rs("ModeStatus")&"$|$"
				M_Act=M_Act&Rs("RefreshFlag")&"$|$"
				M_Act=M_Act&Rs("RecyleIF")&"$|$"
				M_Act=M_Act&Rs("ACT_DiY")&"$|$"
				M_Act=M_Act&Rs("MakeFolderDir")&"$|$"
				M_Act=M_Act&Rs("WriteComment")&"$|$"
				M_Act=M_Act&Rs("CommentCode")&"$|$"
				M_Act=M_Act&Rs("Commentsize")&"$|$"
				M_Act=M_Act&Rs("ModeNote")&"$|$"
				M_Act=M_Act&Rs("CommentTemp")&"$|$"
				M_Act=M_Act&Rs("usermb")&"$|$"
				M_Act=M_Act&Rs("adminmb")&"|*|"
 	End If 
	rs.close
	Rs1.OPen "Select * from Table_ACT where ModeID="&modeid&" ",Conn,1,1
	If Not rs1.eof Then 
		Do While Not rs1.eof 
		T_ACT=T_ACT&Rs1("FieldName")&"@"
		T_ACT=T_ACT&Rs1("Title")&"@"
		T_ACT=T_ACT&Rs1("IsNotNull")&"@"
		T_ACT=T_ACT&Rs1("OrderID")&"@"
		T_ACT=T_ACT&Rs1("Description")&"@"
		T_ACT=T_ACT&Rs1("FieldType")&"@"
		T_ACT=T_ACT&Rs1("Type_Default")&"@"
		T_ACT=T_ACT&Rs1("width")&"@"
		T_ACT=T_ACT&Rs1("height")&"@"
		T_ACT=T_ACT&Rs1("Content")&"@"
		T_ACT=T_ACT&Rs1("Type_Type")&"@"
		T_ACT=T_ACT&Rs1("ISType")&"|||"

		rs1.movenext
		loop
	End If 
	rs1.close
	Call FSOSaveFile(M_Act&T_ACT,ModePath)	
	response.write "<br><br><br><div align=center>操作完成!<a href=" & ModePath & ">请点击这里下载</a>(右键目标另存为)  </div><br><br><br><br><br><br><br>"
	end sub 
	
	Function FSOSaveFile(Templetcontent,FileName)
		on error resume next 
		Dim FileFSO,FileType
		 Set FileFSO = Server.CreateObject("ADODB.Stream")
			With FileFSO
			.Type = 2
			.Mode = 3
			.Open
			.Charset = "utf-8"
			.Position = FileFSO.Size
			.WriteText  Templetcontent
			.SaveToFile Server.MapPath(FileName),2
			If Err.Number<>0 Then 
				Err.Clear 
				Exit Function 
			End If 
			.Close
			End With
		Set FileType = nothing
		Set FileFSO = nothing
	End Function
	sub dr()
	
	  %>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
    	<form name="form2" method="post" action="?A=3">
	    <tr>
	      <td class="bg_tr">您现在的位置：<a href="ACT.MX.asp">模型管理</a> &gt;&gt; 导入模型 </td>
	      </tr>
	    <tr>
          <td align="center"><textarea name="ModeName" cols="80" rows="20"></textarea></td>
        </tr>
	    <tr>
	      <td align="center" ><strong>将模型代码粘贴到上面的输入框</strong>
          <input name="Submit2" type="submit" class="ACT_btn" value=" 提交 "></td>
	      </tr>
    	</form>  
	  </table>
<%end sub
	sub dc()
	
	 %>
	
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <form name="form1" method="post" action="?A=dcbc&modeid=<%=request("modeid")%>">
<tr>
    <td align="center" class="bg_tr">
  目标路径地址:<input name="ModePath" type="text" id="ModePath" value="<%=actcms.actsys%>modemdb.Act">
  <input name="Submit" type="submit" class="ACT_btn" value="  保 存  ">	</td>
  </tr></form>

</table>

<% end sub

	Sub Main()
	%>	
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td class="bg_tr">您现在的位置：模型管理 &gt;&gt; 浏览</td>
  </tr>
  <tr>
    <td>当前模型： <a href="ACT.MX.asp?A=Add">添加模型</a> | <a href="ModeList.asp?A=1">导出内容模型</a> | <a href="ModeList.asp?A=2">导入内容模型</a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td align="center" class="bg_tr">模型ID</td>
    <td align="center" class="bg_tr">模型名称</td>
    <td align="center" class="bg_tr">表名</td>
    <td align="center" class="bg_tr">类型</td>
    <td align="center" class="bg_tr">描述</td>
    <td align="center" class="bg_tr">状态</td>
	<td align="center" class="bg_tr">生成Html</td>
    <td  align="center" class="bg_tr" nowrap>操作</td>
  </tr>
<% 
	  Set Rs =ACTCMS.ACTEXE("SELECT ModeID, ModeName,ModeTable, ModeStatus, IFmake,ModeNote  FROM Mode_Act order by ModeID asc")
	 If Rs.EOF  Then
	 	Response.Write	"<tr><td colspan=""6"" align=""center"">没有记录</td></tr>"
	 Else
		Do While Not Rs.EOF	
			 %>

  <tr >
    <td align="center"><%= Rs("ModeID") %></td>
    <td align="center"><%= Rs("ModeName") %></td>
    <td align="center"><%= Rs("ModeTable") %></td>
    <td align="center"><% if Rs("ModeID")<5 Then Response.Write "<font color=red>系统</font>" Else  Response.Write "<font color=blue>自定义</font>"  %></td>
    <td align="center"><%= Rs("ModeNote") %></td>
    <td align="center"><% IF Rs("ModeStatus") = 0 Then Response.Write "<font color=green>正常</font>" else  Response.Write "<font color=red>禁用</font>" %></td>
    <td align="center">
	<% Select Case  Rs("IFmake")
		Case "0" 
			Response.Write "<font color=red>不生成(动态浏览) </font>" 
		Case "1" 
			Response.Write "<font color=green>生成(静态)</font>"
		Case "2"
			Response.Write "<font color=red>伪静态</font>"
	  End Select 
	%>	</td>
	<td align="center"><a href="?A=dc&modeid=<%= Rs("ModeID") %>">导出模型</a></td>
  </tr>
  <% 
		
		Rs.movenext
		Loop
	End if	 %>
</table>	
	
<%End Sub%>