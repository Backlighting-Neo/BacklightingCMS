<!--#include file="../ACT.Function.asp"-->
<!--#include file="ACT.M.ASP"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>模型管理</title>
<link href="../Images/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/lhgcore.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/lhgdialog.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/Main.js"></script>

</head>
<body>
<% 	Dim Action,ModeID,Rs,ProjectUnit,MakeFolderDir,ContentExtension,AutoPage,admintemplatevalue,usertemplatevalue,umr,def
	ModeID = ChkNumeric(Request("ModeID"))
	umr = ChkNumeric(Request("umr"))
	def = ChkNumeric(Request("def"))

	Action = Request.QueryString("A")
	 if ModeID=0 or ModeID="" Then ModeID=1
	Select Case Action
		   Case "AddSave","ESave"
		   		Call AddSave()
			Case "Add","E"
				Call AddEdit()
			Case Else
				Call Main()
	End Select
	
	IF Action = "Del" Then
	If Not ACTCMS.ChkAdmin() Then  Call Actcms.Alert("对不起，您没有操作权限！","")
	Dim rs1
		IF ModeID  > 1 Then
			Set rs1=actcms.actexe("select id,classid,modeid from class_act where modeid="& ModeID)
			Do While Not rs1.eof
						Dim rsp
						Set rs = ACTCMS.actexe("Select * From "&ACTCMS.ACT_C(rs1("modeid"),2)&" Where ClassID ='" & rs1("classid") & "'")
						Do While Not rs.eof
									Set rsp = ACTCMS.actexe("Select * from Upload_Act  Where ArtileID=" & rs("id") & " and modeid="&rs1("modeid")&"")
										If Not  rsp.eof  Then
											Do While Not rsp.eof
											Call ACTCMS.DeleteFile(Rsp("UpfileDir"))
											Conn.execute("Delete from Upload_Act  Where id= "&Rsp("id"))
											rsp.movenext
											loop
										End If 
									Dim Tmps,TmpUs 
									If Right(ACTCMS.ACT_C(rs1("modeid"),10),1)<>"/" Then 
											Call ACTCMS.DeleteFile(ACTCMS.ActSys&ACTCMS.ACT_C(rs1("modeid"),6)&rs("FileName")&ACTCMS.ACT_C(rs1("modeid"),11))
									Else
											Call ACTCMS.DeleteFile(ACTCMS.ActSys&ACTCMS.ACT_C(rs1("modeid"),6)&rs("FileName")&"/Index"&ACTCMS.ACT_C(rs1("modeid"),11))
									End If 

									Conn.execute("Delete from Comment_Act  Where acticleID=" & rs("id") & " and ModeID="&rs1("modeid")&"")
									Conn.execute("Delete from Digg_ACT  Where NewsID=" & rs("id") & " and modeid="&rs1("modeid")&"")
									Conn.execute("Delete from "&ACTCMS.ACT_C(rs1("modeid"),2)&"  Where ID=" & rs("id") & "")
						rs.movenext
						loop

									Conn.execute("Delete from class_act  Where ID=" & rs1("id") )

			rs1.movenext
			loop
			ACTCMS.ACTEXE("Delete From Mode_Act Where ModeID=" & ModeID)	
			Call Actcms.ActErr("删除模型成功","Mode/ACT.MX.asp","")	
 		Else
		 	Call Actcms.ActErr("系统定义的模型不允许删除","Mode/ACT.MX.asp","1")	
 		End IF
	End IF
	
	
	Sub AddSave()
		 Dim ModeName,ModeTable,sql,ChannelRS,ChannelRSSql,ModeNote,ModeStatus,IFmake,RefreshFlag
		 Dim UpfilesDir,RecyleIF,CommentCode,Commentsize,WriteComment,CommentTemp
		 Dim usermb,adminmb
		 ModeName = ACTCMS.S("ModeName")
		 ModeTable = ACTCMS.S("ModeTable")&"_U_ACT"
		 ModeNote = ACTCMS.S("ModeNote")
 		 usermb=ChkNumeric(ACTCMS.S("usermb"))
		 adminmb=ChkNumeric(ACTCMS.S("adminmb"))
		 ModeStatus = ACTCMS.S("ModeStatus")
		 IFmake = ACTCMS.S("IFmake")
		 RefreshFlag = ChkNumeric(ACTCMS.S("RefreshFlag"))
		 UpfilesDir = ACTCMS.S("UpfilesDir")
		 RecyleIF = ACTCMS.S("RecyleIF")
		 ProjectUnit = ACTCMS.S("ProjectUnit")
		 MakeFolderDir = ACTCMS.S("MakeFolderDir")
		 ContentExtension = ACTCMS.S("ContentExtension")
 		 AutoPage = ChkNumeric(ACTCMS.S("AutoPage"))
		 CommentCode = ChkNumeric(ACTCMS.S("CommentCode"))
		 Commentsize = ChkNumeric(ACTCMS.S("Commentsize"))
		 WriteComment = ChkNumeric(ACTCMS.S("WriteComment"))
		 CommentTemp = ACTCMS.S("CommentTemp")
 		 Call Actcms.CreateFolder(actcms.actsys&"act_inc/cache/"&ModeID&"/")
	
		 
		 If adminmb="1" Then 
 			admintemplatevalue=	ACTCMS.FFile(request.form("admintemplatevalue"),actcms.actsys&"act_inc/cache/"&ModeID&"/"&ACTCMS.ACT_C(ModeID,2)&"-mode.inc")
		 End If 

		
		 If usermb="1" Then 
 			usertemplatevalue=	ACTCMS.FFile(request.form("usertemplatevalue"),actcms.actsys&"act_inc/cache/"&ModeID&"/"&ACTCMS.ACT_C(ModeID,2)&"-usermode.inc")
		 End If 



		 IF ACTCMS.S("ModeName") = "" Then
		 	Call ACTCMS.Alert("模型名称不能为空!",""):Exit Sub
		 End If
		 
	 

		 Set ChannelRS = Server.CreateObject("adodb.recordset")
		 if Action="AddSave" Then
			 If Not ACTCMS.ChkAdmin() Then   Call Actcms.Alert("对不起，您没有操作权限！","")
			 IF ACTCMS.S("ModeTable") = "" Then
				Call ACTCMS.Alert("数据表为空!",""):Exit Sub
			 End if
			 If Not ACTCMS.ACTEXE("SELECT ModeName FROM Mode_Act Where ModeName='" & ModeName & "' order by ModeID desc").eof Then
				Call ACTCMS.Alert("系统已存在该模型名称!",""):Exit Sub
			 End if	

			 If Not ACTCMS.ACTEXE("SELECT ModeTable FROM Mode_Act Where ModeTable='" & ModeTable & "' order by ModeID desc").eof Then
				Call ACTCMS.Alert("系统已存在该数据表!",""):Exit Sub
			 End if	

			 

			  ChannelRSSql = "select * from Mode_Act"
			  ChannelRS.Open ChannelRSSql, Conn, 1, 3
			  ChannelRS.AddNew
		 	  ChannelRS("ModeTable") = ModeTable
			  ChannelRS("ModeName") = ModeName
			  ChannelRS("ModeNote") = ModeNote
			  ChannelRS("ModeStatus") = ModeStatus
			  ChannelRS("IFmake") = IFmake
			  ChannelRS("RecyleIF") = RecyleIF
			  ChannelRS("AutoPage") = AutoPage
			  ChannelRS("UpfilesDir") = UpfilesDir
			  ChannelRS("ProjectUnit") = ProjectUnit
			  ChannelRS("MakeFolderDir") = MakeFolderDir
			  ChannelRS("ContentExtension") = ContentExtension
			  ChannelRS("RefreshFlag") = RefreshFlag
			  ChannelRS("CommentCode") = CommentCode
			  ChannelRS("Commentsize") = Commentsize
			  ChannelRS("WriteComment") = WriteComment
			  ChannelRS("CommentTemp") = CommentTemp
  			  ChannelRS("usermb")=usermb
			  ChannelRS("adminmb")=adminmb
			  ChannelRS("ACT_DiY")="§0§0-1-0-1§0§actcms§0§§0§§0§§0§Simple§§§0§0§0§1§0§1§0§Class.htm§List.Htm§Content.Htm§0§"
			  ChannelRS.Update
			  ChannelRS.Close:Set ChannelRS = Nothing			
				 Dim sqlformat:If  DataBaseType="access" Then sqlformat=" CONSTRAINT PrimaryKey PRIMARY KEY"
 				 Sql="CREATE TABLE "&ModeTable&" ([ID] int IDENTITY (1, 1) NOT NULL "&sqlformat&" ,"&_
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
				"OrderID SmallInt Default 0,"&_
				"commentscount SmallInt Default 0,"&_
				"IStop tinyint Default 0"&_
				")"
			ACTCMS.ACTEXE(sql)
			Application.Contents.RemoveAll
			Call Actcms.ActErr("添加成功","Mode/ACT.MX.asp","")
 		 Else
		If Not ACTCMS.ACTCMS_QXYZ(ModeID,"","") Then   Call Actcms.Alert("对不起，您没有"&ACTCMS.ACT_C(ModeID,1)&"系统该项操作权限！","")
		 	If Not ACTCMS.ACTEXE("SELECT ModeName FROM Mode_Act Where ModeID <>" & ModeID & " AND  ModeName='" & ModeName & "' order by ModeID desc").eof Then
				Call ACTCMS.Alert("系统已存在该模型名称!",""):Exit Sub
			 End if	
			  ChannelRSSql = "select * from Mode_Act Where ModeID=" &ModeID
			  ChannelRS.Open ChannelRSSql, Conn, 1, 3
			  if ChannelRS.eof then Call ACTCMS.Alert("错误!",""):Exit Sub
		 End if 
		
		 If ChkNumeric(ACTCMS.S("adminmb"))="1" Then 
 			Call ACTCMS.FFile(request.form("admintemplatevalue"),actcms.actsys&"act_inc/cache/"&ModeID&"/"&ACTCMS.ACT_C(ModeID,2)&"-mode.inc")
		 End If 
	
 

		 If usermb="1" Then 
 			usertemplatevalue=	ACTCMS.FFile(request.form("usertemplatevalue"),actcms.actsys&"act_inc/cache/"&ModeID&"/"&ACTCMS.ACT_C(ModeID,2)&"-usermode.inc")
		 End If 
 		  ChannelRS("ModeName") = ModeName
		  ChannelRS("ModeNote") = ModeNote
		  ChannelRS("ModeStatus") = ModeStatus
		  ChannelRS("IFmake") = IFmake
		  ChannelRS("RecyleIF") = RecyleIF
		  ChannelRS("UpfilesDir") = UpfilesDir
		  ChannelRS("ProjectUnit") = ProjectUnit
		  ChannelRS("MakeFolderDir") = MakeFolderDir
		  ChannelRS("ContentExtension") = ContentExtension
		  ChannelRS("AutoPage") = AutoPage
		  ChannelRS("RefreshFlag") = RefreshFlag
		  ChannelRS("CommentCode") = CommentCode
		  ChannelRS("Commentsize") = Commentsize
		  ChannelRS("WriteComment") = WriteComment
		  ChannelRS("CommentTemp") = CommentTemp
 		  ChannelRS("usermb")=usermb
		  ChannelRS("adminmb")=adminmb
		  ChannelRS.Update
		  ChannelRS.Close:Set ChannelRS = Nothing	
		  Application.Contents.RemoveAll
  		  Call Actcms.ActErr("修改成功&nbsp;&nbsp;<a href=Mode/ACT.MX.asp>点击这里返回管理首页</a>","Mode/ACT.MX.asp","")
	End Sub
	Sub Main()
	%>	
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td class="bg_tr">您现在的位置：模型管理 &gt;&gt; 浏览</td>
  </tr>
  <tr>
    <td>当前模型： <a href="?A=Add">添加模型</a> | <a href="ModeList.asp?A=1">导出内容模型</a> | <a href="ModeList.asp?A=2">导入内容模型</a></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td align="center" class="bg_tr">模型ID</td>
    <td align="center" class="bg_tr">模型名称</td>
    <td align="center" class="bg_tr">表名</td>
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

  <tr  onMouseOver=overColor(this) onMouseOut=outColor(this)>
    <td align="center"><%= Rs("ModeID") %></td>
    <td align="center"><%= Rs("ModeName") %></td>
    <td align="center"><%= Rs("ModeTable") %></td>
    <td align="center"><%= Rs("ModeNote") %></td>
    <td align="center"><% IF Rs("ModeStatus") = 0 Then Response.Write "<font color=green>正常</font>" else  Response.Write "<font color=red>禁用</font>" %></td>
    <td align="center">
	<% Select Case  Rs("IFmake")
		Case "0" 
			Response.Write "<font color=red>不生成(动态浏览) </font>" 
		Case "1" 
			Response.Write "<font color=green>生成(静态)</font>"
	  End Select 
	%>
	</td>
	<td align="center">
	<a href="Act.DiY.asp?ModeID=<%=Rs("ModeID")%>">自定义显示</a> ┆ <a href="ACT.ListM.ASP?A=L&ModeID=<%=Rs("ModeID")  %>">字段列表</a> ┆ <a href="?A=E&ModeID=<%=Rs("ModeID")  %>" >修改</a> ┆ <a href="?A=Del&ModeID=<%=Rs("ModeID")  %>"  onClick="{if(confirm('确定删除该模型吗,注意!!!删除模型同时会删除该模型下的所有栏目和文章,如果文章比较多,可以先删除文章,然后再删除模型')){return true;}return false;}">删除</a></td>
  </tr>
  <% 
		
		Rs.movenext
		Loop
	End if	 %>
</table>	
	
	
<% 	
 
	End Sub
	Sub AddEdit()
	Dim ModeTable,ModeName,IFmake,RecyleIF,UpfilesDir,ModeStatus,RefreshFlag,ModeNote,A,WriteComment,CommentCode,Commentsize
	Dim CommentTemp,usermb,adminmb
	if Action="Add" Then
	UpfilesDir="UpFiles/"
	AutoPage=0
	A="AddSave"
	ContentExtension=".html"
	WriteComment=3
	CommentCode=0
	Commentsize=0
	MakeFolderDir="html/"
 	CommentTemp="plus/Comment.html"
	Else
	Set Rs=server.CreateObject("adodb.recordset") 
	Rs.OPen "Select * from Mode_Act Where ModeID = "&ModeID&" order by ModeID desc",Conn,1,1
	ModeTable = Rs("ModeTable")
	ModeName = Rs("ModeName")
	IFmake = Rs("IFmake")
	RecyleIF = Rs("RecyleIF")
	UpfilesDir=Rs("UpfilesDir")
	ModeStatus=Rs("ModeStatus")
	RefreshFlag=Rs("RefreshFlag")
	ModeNote=Rs("ModeNote")
	ProjectUnit=Rs("ProjectUnit")
	MakeFolderDir=Rs("MakeFolderDir")
	ContentExtension=Rs("ContentExtension")
	AutoPage=Rs("AutoPage")
	WriteComment=Rs("WriteComment")
	CommentCode=Rs("CommentCode")
	Commentsize=Rs("Commentsize")
	CommentTemp=Rs("CommentTemp")
     usermb=Rs("usermb")
	adminmb=Rs("adminmb")
 	If adminmb="1" Then 
 		admintemplatevalue=Server.HTMLEncode(actcms.LTemplate(actcms.ACTSYS&"act_inc/cache/"&ModeID&"/"&ACTCMS.ACT_C(ModeID,2)&"-mode.inc"))
 	End If 
	
	If usermb="1" Then 
 		usertemplatevalue=Server.HTMLEncode(actcms.LTemplate(actcms.ACTSYS&"act_inc/cache/"&ModeID&"/"&ACTCMS.ACT_C(ModeID,2)&"-usermode.inc"))
	End If 
	If request("def")="1" Then 
			admintemplatevalue=Server.HTMLEncode(M.ACT_NoRormMXList(ModeID))
			adminmb=1
 	End If 

	If adminmb=1 Then def=1
	If usermb=1 Then def=1
	
	If request("umr")="1" Then 
			usertemplatevalue=Server.HTMLEncode(M.ACTUser_MXList(ModeID))
			usermb=1
 	End If 

	A="ESave"
	end If
  %>
<form id="form1" name="form1" method="post" action="?A=<%= A %>&ModeID=<%= Request.QueryString("ModeID") %>">

  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="bg_tr">您现在的位置：<a href="?">模型管理</a> &gt;&gt; 添加/修改 </td>
    </tr>
    <tr>
      <td width="24%" align="right" >模型状态：&nbsp;&nbsp;</td>
      <td width="76%" >
	  <input <% IF ModeStatus = 0 Then Response.Write "Checked" %> id="ModeStatus1" type="radio" name="ModeStatus" value="0" />
     <label for="ModeStatus1"><font color=green> 正常 </font></label>
       <input <% IF ModeStatus = 1 Then Response.Write "Checked" %>  id="ModeStatus2" type="radio" name="ModeStatus" value="1" /><label for="ModeStatus2"><font color=red> 关闭 </font></label>
      <span class="h" style="cursor:help;"  onclick="dohelp('ACTmx_mxzt')"  id="ACTmx_mxzt">帮助</span></td>
    </tr>
    <tr>
      <td height="25" align="right" >模型名称：&nbsp;&nbsp;</td>
      <td height="25" ><input name="ModeName" type="text" class="Ainput"  id="ModeName" value="<%=ModeName %>" />
	  <span class="h" style="cursor:help;"  onclick="dohelp('ACTmx_mxmc')"  id="ACTmx_mxmc">帮助</span></td>
    </tr>
    <tr>
      <td height="25" align="right" >数据表名称：&nbsp;&nbsp;</td>
      <td height="25" ><input <% if A="ESave" then response.Write "disabled" %> name="ModeTable" type="text" class="Ainput"  id="ModeTable" value="<%= ModeTable %>" />
        <span class="h" style="cursor:help;"  onclick="dohelp('ACTmx_sjbmc')"  id="ACTmx_sjbmc">帮助</span><% if A<>"ESave" then response.Write "_U_ACT" %></td>
    </tr>

    <tr>
      <td height="25" align="right" >单位：&nbsp;&nbsp;</td>
      <td height="25" >
	  <input name="ProjectUnit" type="text" class="Ainput"  id="ProjectUnit" value="<%= ProjectUnit %>" size="30" maxlength="250" />
      <span class="h" style="cursor:help;"  onclick="dohelp('ACTmx_dw')"  id="ACTmx_dw">帮助</span>*如：篇、个、本等</td>
    </tr>
	<tr>
      <td height="25" align="right" >模型描述：&nbsp;&nbsp;</td>
      <td height="25" ><input name="ModeNote" type="text" class="Ainput"  id="ModeNote" value="<%= ModeNote %>" size="40" maxlength="250" />
        <span class="h" style="cursor:help;"  onclick="dohelp('ACTmx_mxms')"  id="ACTmx_mxms">帮助</span>简单的描述.不能超过250个字符</td>
    </tr>

	<tr>
      <td height="25" align="right" >文件生成存放目录：&nbsp;&nbsp;</td>
      <td height="25" ><input name="MakeFolderDir" type="text" class="Ainput"  id="MakeFolderDir" value="<%= MakeFolderDir %>" size="40" maxlength="250" /> 不能以 / 开始,留空也可以
	  <span class="h" style="cursor:help;"  onclick="dohelp('ACTmx_wjsccfml')"  id="ACTmx_wjsccfml">帮助</span></td>
    </tr>

 

    <tr>
      <td height="25" align="right" >是否生成HTML：&nbsp;&nbsp;</td>
      <td height="25" >
	  
	    <input <% IF IFmake = 1 Then Response.Write "Checked" %> id="IFmake1" type="radio" name="IFmake" value="1"  />
        <label for="IFmake1">生成(静态)</label>
	  <input <% IF IFmake = 0 Then Response.Write "Checked" %> id="IFmake2" type="radio" name="IFmake" value="0"  /><label for="IFmake2">不生成(动态浏览)</label>	  
	
	    <input <% IF IFmake = 2 Then Response.Write "Checked" %> id="IFmake3" type="radio" name="IFmake" value="2"  />
		    <label for="IFmake3"><font color=green>伪静态(需要服务器支持)</font></label>
      <a href="http://sighttp.qq.com/authd?IDKEY=dc9d9487c8c797974fec7124ec2248094ab916ca5fbed43c/" target="_blank"><font color="red">ACTCMS官方伪静态空间</font></a>	  <span class="h" style="cursor:help;"  onclick="dohelp('ACTmx_sc')"  id="ACTmx_sc">帮助</span></td>
    </tr>

  

    <tr>
      <td height="25" align="right" >自动分页：&nbsp;&nbsp;</td>
      <td height="25" >	 
      <input name="AutoPage" type="text" class="Ainput"  id="AutoPage" value="<%= AutoPage %>" size="30">
	  <span class="h" style="cursor:help;"  onclick="dohelp('ACTmx_fy')"  id="ACTmx_fy">帮助</span>为0则不设置自动分页</td>
    </tr>

    <tr>
      <td height="25" align="right" >删除文章：&nbsp;&nbsp;</td>
      <td height="25" >
	  <input  <% IF RecyleIF = 0 Then Response.Write "Checked" %> id="RecyleIF1" type="radio" name="RecyleIF" value="0">
        <label for="RecyleIF1">放入回收站</label>
      <input  <% IF RecyleIF = 1 Then Response.Write "Checked" %> id="RecyleIF2"  type="radio" name="RecyleIF" value="1"> 
     <label for="RecyleIF2">彻底删除</label>
	 <span class="h" style="cursor:help;"  onclick="dohelp('ACTmx_del')"  id="ACTmx_del">帮助</span></td>
    </tr>
    <tr>
      <td height="25" align="right" >后台文件上传目录：&nbsp;&nbsp;</td>
      <td height="25" ><input name="UpfilesDir" type="text" class="Ainput"  id="UpfilesDir" value="<%= UpfilesDir %>" size="30">
	  <span class="h" style="cursor:help;"  onclick="dohelp('ACTmx_scwj')"  id="ACTmx_scwj">帮助</span></td>
    </tr>



    <tr>
      <td height="25" align="right" >内容文件扩展名：&nbsp;&nbsp;</td>
      <td height="25" ><input name="ContentExtension" type="text" class="Ainput"  id="ContentExtension" value="<%= ContentExtension %>" size="50">
	  <span class="h" style="cursor:help;"  onclick="dohelp('ACTmx_nrkzm')"  id="ACTmx_nrkzm">帮助</span></td>
    </tr>


  

 
	<tr>
      <td height="25" align="right" >后台添加文章，同时发布选项：&nbsp;&nbsp;</td>
      <td height="25" >
	  <input id="RefreshFlag1" <% IF RefreshFlag = 1 Then Response.Write "Checked" %> type="radio" name="RefreshFlag" value="1" >
 <label for="RefreshFlag1">仅发布内容页</label> 
  <input id="RefreshFlag2"  <% IF RefreshFlag = 2 Then Response.Write "Checked" %> type="radio" name="RefreshFlag" value="2" >
  <label for="RefreshFlag2">发布栏目页+内容页+首页</label> 
  <input id="RefreshFlag3"  <% IF RefreshFlag = 3 Then Response.Write "Checked" %> type="radio" name="RefreshFlag" value="3" >
  <label for="RefreshFlag3">发布首页+内容页</label>
  <span class="h" style="cursor:help;"  onclick="dohelp('ACTmx_fb')"  id="ACTmx_fb">帮助</span></td>
    </tr>
    <tr>
      <td align="right" >评论选项：</td>
      <td ><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><input <% IF WriteComment = 0 Then Response.Write "Checked" %>  type="radio" value="0" name="WriteComment" id="WriteComment1" />
<label for="WriteComment1"><font color="red">关闭本模型的所有信息评论</font></label><br />
<input <% IF WriteComment = 1 Then Response.Write "Checked" %>  type="radio" value="1" name="WriteComment" id="WriteComment2"  />
<label for="WriteComment2">本模型只允许<font color="green">会员</font>评论，且评论内容需要后台的审核</label>
<br />
<input <% IF WriteComment = 2 Then Response.Write "Checked" %>  type="radio" value="2" name="WriteComment" id="WriteComment3"  />
<label for="WriteComment3">本模型只允许<font color="green">会员</font>评论，且评论内容不需要后台审核</label>
<br />
<input <% IF WriteComment = 3 Then Response.Write "Checked" %>  type="radio" value="3" name="WriteComment" id="WriteComment4"  />
<label for="WriteComment4">本模型允许<font color="green">会员</font>，<font color="red">游客</font>评论，且评论内容需要后台审核</label>
<br />
<input <% IF WriteComment = 4 Then Response.Write "Checked" %>  type="radio" value="4" name="WriteComment" id="WriteComment5"  />
<label for="WriteComment5">本模型允许<font color="green">会员</font>，<font color="red">游客</font>评论，且评论内容不需要后台审核<br>
</label></td>
  </tr>
  <tr>
    <td height="50" style="height:30">评论需要验证码：
      <INPUT <% IF CommentCode = 0 Then Response.Write "Checked" %>   type="radio"  value="0" name="CommentCode" id="CommentCode1">
<label for="CommentCode1">是</label>
<INPUT <% IF CommentCode = 1 Then Response.Write "Checked" %>  type="radio" value="1" name="CommentCode" id="CommentCode2">
<label for="CommentCode2">否</label></td>
  </tr>
  <tr>
    <td height="50">评论字数控制：<input name="Commentsize" type="text" class="Ainput"  id="Commentsize" value="<%=Commentsize%>" size="8" maxlength="5"> 
	不限制请输入&quot;0&quot; </td>
  </tr>
   <tr>
    <td height="50">评论页模板：<input name="CommentTemp" type="text" class="Ainput"  id="CommentTemp" value="<%=CommentTemp%>" size="40" maxlength="59"> 
	<input class="ACT_btn" type="button"  onClick="OpenWindowAndSetValue('../include/print/SelectPic.asp?CurrPath=<%=ACTCMS.ActSys%><%=actcms.SysThemePath&"/"&actcms.NowTheme%>',500,320,window,document.form1.CommentTemp);" value="选择模板...">	</td>
  </tr>
  
</table></td>
    </tr>



<!--  -->
    <tr>
      <td height="25" align="right" >后台录入表单模板</td>
      <td height="25" >
	  <input <% IF adminmb = 0 Then Response.Write "Checked" %> id="adminmb1" type="radio" name="adminmb" value="0"  onclick="adminmbq(0)"  />
        <label for="adminmb1">自动录入表单</label>

        <input  <% IF adminmb = 1 Then Response.Write "Checked" %> id="adminmb2"  type="radio" name="adminmb" value="1"  onclick="adminmbq(1)"  />
        <label for="adminmb2">手动录入表单</label>
		
		 
		&nbsp; <span id="amr"><a href="?A=E&ModeID=<%= Request.QueryString("ModeID") %>&def=1&umr=<%=request("umr")%>">载入默认</a></span>
		
		 
		</td>
    </tr>
    <tr id="adminmbs" 
	<%If adminmb=0 Then response.write "style=""DISPLAY: none"""%>
	>
      <td height="25" colspan="2" align="center" >
	  
        <textarea name="admintemplatevalue" style="width:98%" rows="10"><%=admintemplatevalue%></textarea>	  </td>
    </tr>


    <tr>
      <td height="25" align="right" >前台投稿表单模板：</td>
      <td height="25" > 
	  <input <% IF usermb = 0 Then Response.Write "Checked" %> id="usermb1" type="radio" name="usermb" value="0"  onclick="usermbq(0)"  />
        <label for="usermb1">自动录入表单</label>

        <input  <% IF usermb = 1 Then Response.Write "Checked" %> id="usermb2"  type="radio" name="usermb" value="1"  onclick="usermbq(1)"  />
        <label for="usermb2">手动录入表单</label>
		 
		&nbsp;<span id="umr"><a href="?A=E&ModeID=<%= Request.QueryString("ModeID") %>&umr=1&def=<%=request("def")%>">载入默认</a></span>
		 
		</td>
    </tr>
    <tr id="usermbs" <%If usermb=0 Then response.write "style=""DISPLAY: none"""%>
	>
      <td height="25" colspan="2" align="center" >
	  <textarea name="usertemplatevalue" style="width:98%" rows="10"><%=usertemplatevalue%></textarea>	  </td>
    </tr>

<!--  -->


    <tr>
      <td align="right" >&nbsp;</td>
      <td ><input type=button onclick=CheckForm() class="ACT_btn"  name=Submit1 value="  保存  " />
      <input type="reset" name="Submit2" class="ACT_btn" value="  重置  " /></td>
    </tr>
  </table>
</form><br>
<script language="JavaScript" type="text/javascript">
		function adminmbq(q)
				{ if (q==0)
				  {
					 adminmbs.style.display="none";
					 amr.style.display="none";
				  }
				  else
					{
			    	adminmbs.style.display="";
			    	amr.style.display="";

					 }
				}

		function usermbq(q)
				{ if (q==0)
				  {
					 usermbs.style.display="none";
					 umr.style.display="none";
 				  }
				  else
					{
			    	usermbs.style.display="";
			    	umr.style.display="";

					 }
				}

 
function CheckForm()
{ var form=document.form1;
	
	 if (form.ModeName.value=='')
		{ alert("请输入模型名称!");   
		  form.ModeName.focus();    
		   return false;
		} 
	 if (form.ModeTable.value=='')
		{ alert("请输入数据表名称!");   
		  form.ModeTable.focus();    
		   return false;
		} 	    form.Submit1.value="正在提交数据,请稍等...";
		form.Submit1.disabled=true;
		form.Submit2.disabled=true;		
	    form.submit();
        return true;
	}
	function OpenWindowAndSetValue(Url,Width,Height,WindowObj,SetObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:yes;help:no;scroll:yes;');
	if (ReturnStr!='') SetObj.value=ReturnStr;
	return ReturnStr;
}	

</script>
<script language="javascript">usermbq(<%=umr%>);</script>
<script language="javascript">adminmbq(<%=def%>);</script>

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

</body>
</html>
