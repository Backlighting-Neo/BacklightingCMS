<%@ LANGUAGE = VBScript CodePage = 65001%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ACTCMS安装程序</title>
<link href="css.css" rel="stylesheet" type="text/css" />
<SCRIPT src="../act_inc/js/lhgajax.js" type="text/javascript"></SCRIPT>
<!--#include file="../act_Inc/Md5.asp"-->
</head>
<body>
<%
	response.buffer=true
	session.codepage=65001
	response.charset="utf-8"
 	Dim A,SysSetting,helpjs,Fso
	A=request("A")

	Set Fso = Server.CreateObject("scripting.FileSystemObject")
	'If Fso.FileExists(Server.MapPath("../ACT_inc/lock/Install.lock")) Then Response.Write "您已经安装过ACTCMS,如果需要重新安装，请删除 ACT_inc/lock/Install.lock 文件！" : Response.End
	Set Fso = Nothing
	
  	Select Case A
		Case "1"
			Call main()
		Case "2"
			Call seting()
		Case "3"
			Call save()
		Case Else 
			Call main()
	End Select 
 Sub save()
	If request("webname")="" Then 
		 Response.Write ("<script>alert('错误提示:\n\n网站名称不能为空!');history.back();</script>")
	 	 response.End 
	End If 

	If request("webtitle")="" Then 
		 Response.Write ("<script>alert('错误提示:\n\n网站标题不能为空!');history.back();</script>")
	 	 response.End 
	End If 

	If request("webadmin")="" Then 
		 Response.Write ("<script>alert('错误提示:\n\n站长姓名不能为空!');history.back();</script>")
	 	 response.End 
	End If 

	If request("adminmail")="" Then 
		 Response.Write ("<script>alert('错误提示:\n\n管理员信箱不能为空!');history.back();</script>")
	 	 response.End 
	End If 
	If request("loginname")="" Then 
		 Response.Write ("<script>alert('错误提示:\n\n登录帐号不能为空!');history.back();</script>")
	 	 response.End 
	End If 
	If request("password")="" Then 
		 Response.Write ("<script>alert('错误提示:\n\n登录密码不能为空!');history.back();</script>")
	 	 response.End 
	End If 
	If request("admindir")="" Then 
		 Response.Write ("<script>alert('错误提示:\n\n后台目录不能为空!');history.back();</script>")
	 	 response.End 
	End If 
	
	If request("webhc")="" Then 
		 Response.Write ("<script>alert('错误提示:\n\n缓存名称不能为空!');history.back();</script>")
	 	 response.End 
	End If 
	
	if request("sjk")="mssql" then 
   	 	ConnStr="Provider = Sqloledb; User ID = " & request("sqluser") & "; Password = " &  request("sqlps") & "; Initial Catalog = " &  request("sqlname") & "; Data Source = " &  request("server") & ";"
		SQL = split(LoadTemplate("mssql.sql"),"-")
 	else
		db = "../data_act/"&request("webdb")&"" 
		call Createaccess(db)
		Connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
		SQL = split(LoadTemplate("access.sql"),VBCRLF)
	end if 
	 
 	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Connstr
	If Err Then
		Err.Clear
		Set Conn = Nothing
		Response.Write "数据库连接出错， [<a href=""http://www.actcms.com/install.htm"">Help</a>]"
		Response.End
	End If
 	for i=0 to ubound(sql)
		 conn.execute(sql(i))
 	Next
	
 	SysSetting=""&request("webname")&"^@$@^"&request("webtitle")&"^@$@^"&request("AutoDomain")&"^@$@^"&request("InstallDir")&"^@$@^index.asp^@$@^images/logo.gif^@$@^"&request("webadmin")&"^@$@^"&request("adminmail")&"^@$@^"&request("admindir")&"^@$@^Index.Html^@$@^102411^@$@^jpg/gif/bmp/png/swf/rar^@$@^0^@$@^0^@$@^1^@$@^1^@$@^admin|manager|actcms|user|管理员|网站管理员|左岸^@$@^reg.htm^@$@^reglist.htm^@$@^templets^@$@^0^@$@^0^@$@^0^@$@^0^@$@^金币^@$@^个^@$@^^@$@^^@$@^"
 	Set Rs=server.CreateObject("adodb.recordset") 
	Rs.OPen "Select * from Config_ACT",Conn,1,3
	rs.addnew
	Rs("ActCMS_Theme")=CStr("default")
	Rs("ActCMS_SysSetting")=CStr(SysSetting)
 	Rs("ActCMS_OtherSetting")="版权^@&@^ACTCMS^@&@^^@&@^SMTP服务器地址^@&@^SMTP登录用户名^@&@^fdgdfgd^@&@^UpFiles/server/^@&@^0^@&@^^@&@^1^@&@^Scripting.FileSystemObject^@&@^"
 	Rs("ActCMS_Upfile")="#333333^@*&*@^0^@*&*@^0^@*&*@^999^@*&*@^www.actcms.com^@*&*@^28^@*&*@^#FF0000^@*&*@^Arial^@*&*@^0^@*&*@^WaterMap.gif^@*&*@^0.8 ^@*&*@^110^@*&*@^35^@*&*@^0^@*&*@^420^@*&*@^400^@*&*@^0^@*&*@^999^@*&*@^"
	Rs.Update
	Rs.close 
 	Set Rs=server.CreateObject("adodb.recordset") 
	Rs.OPen "Select * from Admin_ACT Where Id =1",Conn,1,3
	rs.addnew
	Rs("Admin_Name")=request("loginname")
	Rs("PassWord")=md5(request("password"))
	Rs("tel")="13721031387"
	Rs("email")="actcms@gmail.com"
	Rs("adddate")=Now
	Rs("locked")=0
	Rs("SuperTF")=1
	Rs("Description")="ActCMS系统安装分配的超级管理员"
	Rs("Purview")=0
	Rs("ACTCMS_QXLX")=0
	Rs("ACT_Other")=0
 	Rs.Update
	Rs.close
    conn.execute("insert Into DiyMenu_ACT(MenuName,MenuUrl,OpenWay,AdminID) values('<font color=green>添加内容</font>','ACT_Mode/ACT.Add.asp?ModeID=1','main','1')")
  	Rs.OPen "Select * from Mode_Act",Conn,1,3
	rs.addnew
 	rs("ModeName")="文章"
	rs("IFmake")="0"
	rs("ModeTable")="Article_ACT"
	rs("FileFolder")="5"
	rs("AutoPage")="0"
	rs("ProjectUnit")="篇"
	rs("UpFilesDir")="UpFiles/Article/"
	rs("ContentExtension")=".html"
	rs("ModeStatus")="0"
	rs("RefreshFlag")="2"
	rs("RecyleIF")="1"
	rs("ACT_DiY")="§0§0-1-0-1§0§actcms-安川网络§0§§0§§0§§0§Default§§§0§0§0§0§0§0§0§class.html§list.html§content.html§0§"
	rs("MakeFolderDir")="html/"
	rs("WriteComment")="3"
	rs("CommentCode")="1"
	rs("Commentsize")="0"
	rs("ModeNote")="模型描述"
	rs("CommentTemp")="plus/Comment.html"
	rs("adminmb")="0"
	rs("usermb")="0" 
	Rs.Update
	Rs.close
    conn.execute("insert Into ModeUser_Act(ModeName,ModeTable,ModeNote,Template,RegCode,SpaceID) values('普通用户','Field_User_ACT','备注','user','0','1')")
    conn.execute("insert Into Group_Act(DefaultGroup,Description,ChargeType,GroupPoint,ValidDays,ModeID,GroupSetting,GroupName) values(0,'邮件验证会员',1,0,0,1,'0^@$@^邮件验证会员简介^@$@^0^@$@^0^@$@^100^@$@^1000^@$@^10^@$@^0^@$@^Article/^@$@^1024^@$@^jpg/gif/bmp/png^@$@^1^@$@^2^@$@^Simple^@$@^0^@$@^0^@$@^','邮件验证会员')")
	conn.execute("insert Into Group_Act(DefaultGroup,Description,ChargeType,GroupPoint,ValidDays,ModeID,GroupSetting,GroupName) values(0,'后台验证会员',1,100,999,1,'0^@$@^^@$@^0^@$@^0^@$@^100^@$@^1000^@$@^10^@$@^0^@$@^Article/^@$@^1024^@$@^0^@$@^1^@$@^3^@$@^^@$@^0^@$@^0^@$@^','后台验证会员')")
	conn.execute("insert Into Group_Act(DefaultGroup,Description,ChargeType,GroupPoint,ValidDays,ModeID,GroupSetting,GroupName) values(1,'注册会员',1,0,0,1,'0^@$@^注册会员^@$@^0^@$@^0^@$@^100^@$@^1000^@$@^100^@$@^1^@$@^Article/^@$@^1024^@$@^jpg/gif/bmp/png^@$@^1^@$@^1^@$@^Simple^@$@^10^@$@^0^@$@^','注册会员')") 
    conn.execute("insert Into ClassLink_Act(ClassLinkName,Description,AddDate) values('友情链接','','"&Now()&"')")
 	conn.execute("insert Into Plus_ACT(PlusName,PlusID,PlusUrl,IsUse,PlusIntro,OrderID,PlusConfig) values('友情链接','yqlj_ACT','Link/Index.asp','0','By ACTCMS.COM','10','')")
	conn.execute("insert Into Plus_ACT(PlusName,PlusID,PlusUrl,IsUse,PlusIntro,OrderID,PlusConfig) values('留言系统-留言系统管理','lyxt_ACT','Book/Index.asp-Book/Sys.asp','0','By ACTCMS.COM','10','0^@$@^0^@$@^0^@$@^0^@$@^10^@$@^0^@$@^^@$@^plus/Book.html^@$@^')")
	conn.execute("insert Into Plus_ACT(PlusName,PlusID,PlusUrl,IsUse,PlusIntro,OrderID,PlusConfig) values('自定义表单','form_ACT','Form/Index.asp','0','By ACTCMS.COM','1','')")
	conn.execute("insert Into Plus_ACT(PlusName,PlusID,PlusUrl,IsUse,PlusIntro,OrderID,PlusConfig) values('单页系统','dyxt_ACT','diypage/index.asp','0','单页系统 By ACTCMS.COM','10','')")
	
	conn.execute("insert Into Plus_ACT(PlusName,PlusID,PlusUrl,IsUse,PlusIntro,OrderID,PlusConfig) values('主题插件','theme','theme/index.asp','0','By ACTCMS.COM','10','')")
	conn.execute("insert Into Plus_ACT(PlusName,PlusID,PlusUrl,IsUse,PlusIntro,OrderID,PlusConfig) values('在线插件','plugin','plugin/index.asp','0','By ACTCMS.COM','10','')")
	conn.execute("insert Into Plus_ACT(PlusName,PlusID,PlusUrl,IsUse,PlusIntro,OrderID,PlusConfig) values('心情插件','plugin','mood/index.asp','0','By ACTCMS.COM','10','')")
 	
	conn.execute("insert Into Plus_ACT(PlusName,PlusID,PlusUrl,IsUse,PlusIntro,OrderID,PlusConfig) values('广告系统','ggxt_ACT','gg/index.asp','0','by 东风','10','')")
	conn.execute("insert Into Plus_ACT(PlusName,PlusID,PlusUrl,IsUse,PlusIntro,OrderID,PlusConfig) values('Digg浏览-Digg管理','digg_act','Digg/Index.asp-digg/sys.asp','0','By ACTCMS.COM','10','0^@$@^^@$@^^@$@^')")
	conn.execute("insert Into Plus_ACT(PlusName,PlusID,PlusUrl,IsUse,PlusIntro,OrderID,PlusConfig) values('投票系统','vote','vote/Index.asp','0','','10','')")
 	conn.execute("insert Into Vote_act(title,isLock,VoteTime,VoteType,VoteNum,rootid,VoteStart,VoteEnd ) values('你是从哪儿得知本站的？','0','2009-3-15 20:02:24','1','0','0','2009-3-15','2010-3-31')")
	conn.execute("insert Into Vote_act(title,isLock,VoteTime,VoteType,VoteNum,rootid,VoteStart,VoteEnd ) values('朋友介绍','0','2009-3-15 20:02:24','1','0','1','2009-3-15','2009-3-15')")
	conn.execute("insert Into Vote_act(title,isLock,VoteTime,VoteType,VoteNum,rootid,VoteStart,VoteEnd ) values('门户网站的搜索引擎','0','2009-3-15 20:02:24','1','0','1','2009-3-15','2009-3-15')")
	conn.execute("insert Into Vote_act(title,isLock,VoteTime,VoteType,VoteNum,rootid,VoteStart,VoteEnd ) values('google或百度引擎','0','2009-3-15 20:02:24','1','0','1','2009-3-15','2009-3-15')")
	conn.execute("insert Into Vote_act(title,isLock,VoteTime,VoteType,VoteNum,rootid,VoteStart,VoteEnd ) values('其他途径','0','2009-3-15 20:02:24','1','0','1','2009-3-15','2009-3-15')")	
    conn.execute("insert Into ACT_LabelFolder(Foldername) values('文章系统')")
    conn.execute("insert Into ACT_LabelFolder(Foldername) values('图片系统')")
    conn.execute("insert Into ACT_LabelFolder(Foldername) values('会员')")
    conn.execute("insert Into space_ACT(ClassName,UModeID,ClassOrder,ClassTemp,ModeID) values('空间首页','1','21','index.html','1')")
    conn.execute("insert Into space_ACT(ClassName,UModeID,ClassOrder,ClassTemp,ModeID) values('个人文集','1','21','space.html','1')")
    conn.execute("insert Into templets_act(templets,UserSet) values('space','2')")
	conn.execute("insert Into ATT_ACT(aid,Aname) values('1','首页头条')")
    conn.execute("insert Into ATT_ACT(aid,Aname) values('2','文章推荐')")
    conn.execute("insert Into Sitelink_ACT(Title,Url,OpenType,IFS,OrderID,Num,description,repset,repcontent) values('内容测试','','',1,3,1,'关键字替换指定代码',0,'<h1>{$content}</h1>')")
    conn.execute("insert Into Sitelink_ACT(Title,Url,OpenType,IFS,OrderID,Num,description,repset,repcontent) values('合肥网络公司','http://www.actcms.com','_blank',1,1,1,'合肥网络公司',1,'')")
    conn.execute("insert Into Sitelink_ACT(Title,Url,OpenType,IFS,OrderID,Num,description,repset,repcontent) values('合肥网站建设','http://www.web265.com','_blank',1,1,1,'合肥网站建设',1,'')")
    conn.execute("insert Into Mood_Plus_ACT(Title,Status,TitleContent,PicContent,SubmitNum,UnlockTime) values('默认',0,'支持@&@高兴@&@震惊@&@愤怒@&@无聊@&@无奈@&@谎言@&@枪稿@&@不解@&@标题党@&@@&@@&@@&@@&@@&@','images/Plus/xq1.gif@&@images/Plus/xq2.gif@&@images/Plus/xq3.gif@&@images/Plus/xq4.gif@&@images/Plus/xq5.gif@&@images/Plus/xq6.gif@&@images/Plus/xq7.gif@&@images/Plus/xq8.gif@&@images/Plus/xq9.gif@&@images/Plus/xq10.gif@&@images/Plus/xq11.gif@&@images/Plus/xq12.gif@&@images/Plus/xq13.gif@&@images/Plus/xq14.gif@&@images/Plus/xq15.gif@&@',1,1)")
    conn.execute("insert Into Link_Act(ClassLinkID,SiteName,url,linktype,adddate,locked,sh) values(1,'ActCMS','http://www.actcms.com',0,'"&Now()&"',0,1)")
    conn.execute("insert Into Link_Act(ClassLinkID,SiteName,url,linktype,adddate,locked,sh) values(1,'合肥网络公司','http://www.web265.com',0,'"&Now()&"',0,1)")
    conn.execute("insert Into Link_Act(ClassLinkID,SiteName,url,linktype,adddate,locked,sh) values(1,'合肥网站建设','http://www.actcms.cn',0,'"&Now()&"',0,1)")
	Dim DataConn:Set DataConn = Server.CreateObject("ADODB.Connection")
	DataConn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath("label.mdb")
	If Err Then 
	Err.Clear:Set DataConn = Nothing:Response.Write "<tr><td>数据库路径不正确，连接出错</td></tr>":Response.End
	else
	 Dim rs:set rs=server.createobject("adodb.recordset")
	 rs.open "select * from Label_Act",dataconn,1,1
	 Dim rsa:set rsa=server.createobject("adodb.recordset")
	 do while not rs.eof 
	  rsa.open "select * from Label_Act where labelname='" & rs("labelname") & "'",conn,1,3
	  if rsa.eof then
		 rsa.addnew
		 rsa("LabelName")=rs("LabelName")
		 rsa("LabelContent")=rs("LabelContent")
		 rsa("Description")=rs("Description")
		 rsa("LabelType")=rs("LabelType")
		 rsa("LabelFlag")=rs("LabelFlag")
		 rsa("AddDate")=rs("AddDate")
		 n=n+1
		rsa.update
	  end if
	   rsa.close
	  rs.movenext
	 loop
	 rs.close:set rs=nothing
	 set rsa=nothing
	end if
	
 	Dim Connfile
	Connfile = "<" & "%" &"@ LANGUAGE = VBScript CodePage = 65001"& "%" & ">"& vbCrLf
	Connfile = Connfile &  "<" & "%" & vbCrLf
	Connfile = Connfile &  "Option Explicit" & vbCrLf
	Connfile = Connfile &  "response.buffer=true" & vbCrLf
	Connfile = Connfile &  "session.codepage=65001" & vbCrLf
	Connfile = Connfile &  "response.charset=""utf-8""" & vbCrLf
	Connfile = Connfile &  "Dim Conn,db,MyDbPath,actcool,actField,NowString,ConnStr,aspexe" & vbCrLf
	Connfile = Connfile &  "Const isSqlDataBase = 0" & vbCrLf
	Connfile = Connfile &  "Const MsxmlVersion="".3.0""  '系统采用XML版本设置" & vbCrLf
	Connfile = Connfile &  "Const AcTCMSN="""&request("webhc")&"""'系统缓存名称.在一个URL下安装多个ACTCMS请设置不同名称" & vbCrLf
	Connfile = Connfile &  "Const DataBaseType="""&request("sjk")&"""  '' 数据库类型: 值分别为 access   mssql" & vbCrLf
 	if request("sjk")="mssql" then 
 	 Connfile = Connfile &  "Dim DataServer,DataUser,DataBaseName,DataBasePsw" & vbCrLf
	 Connfile = Connfile &  " '如果是SQL数据库，请认真修改好以下数据库选项" & vbCrLf
	 Connfile = Connfile &  "DataServer   = """&request("server")&"""                                  '数据库服务器IP" & vbCrLf
	 Connfile = Connfile &  "DataBaseName = """&request("sqlname")&"""                                  '数据库名称" & vbCrLf
	 Connfile = Connfile &  "DataUser     = """&request("sqluser")&"""                                  '访问数据库用户名" & vbCrLf
	 Connfile = Connfile &  "DataBasePsw  = """&request("sqlps")&"""                                  '访问数据库密码 " & vbCrLf
	 Connfile = Connfile &  "NowString = ""getdate()""" & vbCrLf& vbCrLf& vbCrLf
	 Connfile = Connfile &  "'=============================================================== 以下代码请不要自行修改========================================" & vbCrLf
	 Connfile = Connfile &  "Sub ConnectionDatabase()" & vbCrLf
	 Connfile = Connfile &  "  On Error Resume Next" & vbCrLf
	 Connfile = Connfile &  "	ConnStr=""Provider = Sqloledb; User ID = "" & datauser & ""; Password = "" & databasepsw & ""; Initial Catalog = "" & databasename & ""; Data Source = "" & dataserver & "";""" & vbCrLf
	 Connfile = Connfile &  "	Set conn = Server.CreateObject(""ADODB.Connection"")" & vbCrLf
	 Connfile = Connfile &  "	conn.open ConnStr" & vbCrLf
	 Connfile = Connfile &  "	If Err Then Err.Clear:Set conn = Nothing:Response.Write ""数据库连接出错，请检查Conn.asp文件中的数据库参数设置。"":Response.End" & vbCrLf
	 Connfile = Connfile &  "End Sub" & vbCrLf
	else 
 	Connfile = Connfile &  "NowString = ""Now()""" & vbCrLf
 	Connfile = Connfile &  "MyDbPath ="""&request("InstallDir")&"""'系统安装目录,如在虚拟目录下安装.请填写 /虚拟目录名称/" & vbCrLf
	Connfile = Connfile &  "db = ""data_act/"&request("webdb")&""" 'ACCESS数据库的文件名" & vbCrLf
	Connfile = Connfile &  "Connstr = ""Provider=Microsoft.Jet.OLEDB.4.0;Data Source="" & Server.MapPath(MyDbPath & db)" & vbCrLf
	Connfile = Connfile &  "Sub ConnectionDatabase()" & vbCrLf
	Connfile = Connfile &  "	On Error Resume Next" & vbCrLf
	Connfile = Connfile &  "	Set Conn = Server.CreateObject(""ADODB.Connection"")" & vbCrLf
	Connfile = Connfile &  "	Conn.Open Connstr" & vbCrLf
	Connfile = Connfile &  "	If Err Then" & vbCrLf
	Connfile = Connfile &  "		Err.Clear" & vbCrLf
	Connfile = Connfile &  "		Set Conn = Nothing" & vbCrLf
	Connfile = Connfile &  "		Response.Write ""数据库连接出错，请检查Conn.asp文件中的数据库参数设置 [<a href=http://www.actcms.com/install.htm>Help</a>] &nbsp;[<a href='install/index.asp'>点击安装</a>]""" & vbCrLf
	Connfile = Connfile &  "		Response.End" & vbCrLf
	Connfile = Connfile &  "	End If" & vbCrLf
	Connfile = Connfile &  "End Sub" & vbCrLf& vbCrLf
	end if 
	Connfile = Connfile &  "Sub CloseConn()" & vbCrLf
	Connfile = Connfile &  "	On Error Resume Next" & vbCrLf
	Connfile = Connfile &  "	If IsObject(Conn) Then" & vbCrLf
	Connfile = Connfile &  "		Conn.Close:Set Conn = Nothing" & vbCrLf
	Connfile = Connfile &  "	End If" & vbCrLf
	Connfile = Connfile &  "End Sub" & vbCrLf
	Connfile = Connfile &  "%" & ">" & vbCrLf
	Connfile = Connfile &  "<!--#include file=""ACT_INC/ACT.Common.asp"" -->" 
	Dim adminfile
		adminfile = adminfile &  "<" & "%" & vbCrLf
		adminfile = adminfile &  "Const CheckCode="& request("yzm")&"  '是否启用后台管理验证码 是： True  否： False " & vbCrLf
		adminfile = adminfile &  "Const CheckManageCode="& request("rzm")&"  '是否启用后台管理认证码 是： True  否： False " & vbCrLf
		adminfile = adminfile &  "Const CheckManageCodeContent="""& request("glrzm")&"""  '后台管理认证码，请修改，这样即使有人知道了您的后台用户名和密码也不能登录后台 " & vbCrLf
		adminfile = adminfile &  "Const Security="""& request("Security")&"""  '后台安全码 " & vbCrLf
		adminfile = adminfile &  "%" & ">" 

	Call FSOSaveFile(Connfile,request("InstallDir")&"Conn.asp")	
	Call FSOSaveFile(adminfile,request("InstallDir")&request("admindir")&"/CheckCode.asp")	
	Call FSOSaveFile("Powered By ACTCMS 4.0",request("InstallDir")&"act_inc/lock/Install.lock")	

	If  request("testdata")="1" Then 
		Dim datasql,di
		dataSQL = split(LoadTemplate("data.txt"),VBCRLF)
		for di=0 to ubound(datasql)
			 conn.execute(datasql(di))
		Next
	End If 


	Application.Contents.RemoveAll
   Response.Write ("<script language=""Javascript""> alert('系统安装成功,为安全起见,请删除install目录,点击确定进入后台');location.href='" & request("InstallDir")&request("admindir")&"/Login.asp" & "';</script>")
End Sub 
function Createaccess(dbname)
 	on error resume next
	dim dbX : Set dbX = Server.CreateObject("ADOX.CataLog")
	dbX.Create "Provider=Microsoft.Jet.OLEdb.4.0;Data Source=" & Server.MapPath(dbname)
	If err then response.write "错误," & err.description : response.end
	Set dbX = nothing
end function

Function  LoadTemplate(TempString) 
 	on error resume next
	Dim  Str,A_W
	set A_W=server.CreateObject("adodb.Stream")
	A_W.Type=2 
	A_W.mode=3 
	A_W.charset="utf-8"
	A_W.open
	A_W.loadfromfile server.MapPath(TempString)
	If Err.Number<>0 Then Err.Clear:LoadTemplate="错误.没有找到文件":Exit Function
	Str=A_W.readtext
	A_W.Close
	Set  A_W=nothing
	LoadTemplate=Str
End  function


Function FSOSaveFile(Templetcontent,FileName)
	On Error Resume Next
	Dim FileFSO,FileType
	 Set FileFSO = Server.CreateObject("ADODB.Stream")
		With FileFSO
		.Type = 2
		.Mode = 3
		.Open
		.Charset = "utf-8"
		.Position = FileFSO.Size
		.WriteText  Templetcontent&	 vbcrlf 
		.SaveToFile Server.MapPath(FileName),2
		.Close
		End With
	Set FileType = nothing
	Set FileFSO = nothing
End Function

Sub main
%>
<table width="830" border="0" align="center" cellpadding="0" cellspacing="0" >
  <tr>
    <td align="center">〓 许可协议 〓</td>
  </tr>
  <tr>
    <td>
	<div class="pact">
            <p>本软件是自由软件，遵循 Apache License 3.0 许可协议 &lt;<a href="http://www.apache.org/licenses/LICENSE-2.0" target="_blank">http://www.apache.org/licenses/LICENSE-2.0</a>&gt;<p>
            <p>本软件的版权归 ACTCMS官方 所有，且受《中华人民共和国计算机软件保护条例》等知识产权法律及国际条约与惯例的保护。<p>
            <p>本协议适用且仅适用于 ACTCMS 3.x 版本，ACTCMS官方拥有对本协议的最终解释权。<p>
            <p>无论个人或组织、盈利与否、用途如何（包括以学习和研究为目的），均需仔细阅读本协议，在理解、同意、并遵守本协议的全部条款后，方可开始使用本软件。<p>
            <h4><strong>一、协议许可和限制</strong></h4>
            <ol>
              <p>1、未经作者书面许可，不得衍生出私有软件。<p>
              <p>2、不管首页是否以ACTCMS系统生成，网站首页最下面保留清晰可见的支持信息并链接到ACTCMS站，不得以友情链接等方式代替，若网站性质等因素所限，不适合保留支持信息，请和作者签订《ACTCMS 2.0授权合同》<p>
              <p>3、您拥有使用本软件构建的网站全部内容所有权，并独立承担与这些内容的相关法律义务。<p>
              <p>4、未经官方许可，禁止在 ACTCMS 的整体或任何部分基础上以发展任何派生版本、修改版本或第三方版本用于重新分发<p>
              <p>5、您将本软件应用在商业用途时，需遵守以下几条：
      <p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;1、使用本软件建设网站时，无需支付使用费用，但需保留ACTCMS支持链接信息。<p>
                  <p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2、本源码可以用在商业用途，但不可以更名销售，若有OEM需求，请和作者联系。<p>
                  <p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;3、若网站性质等因素所限，不适合保留支持信息，请与作者联系取得书面授权。<p>
				  <p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;4、使用者所生成的网站，首页要包含软件的版权信息；不得对后台版权进行修改。<p>
              <p>
            <h4><strong>二、有限担保和免责声明</strong></h4>
              <p>1、本软件及所附带的文件是作为不提供任何明确的或隐含的赔偿或担保的形式提供的。<p>
              <p>2、用户出于自愿而使用本软件，您必须了解使用本软件的风险，在尚未购买产品技术服务之前，我们不承诺提供任何形式的技术支持、使用担保，也不承担任何因使用本软件而产生问题的相关责任。<p>
              <p>3、ACTCMS官方不对使用本软件构建的网站中的文章或信息承担责任。<p>
            <br>
            <p>本协议保留作者的版权信息在许可协议文本之内，不得擅自修改其信息。</p>
            <p>2008-3，第1.0版 (ACTCMS官方保留对此许可协议的更新及解释权力)<br />
           &nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;协议著作权所有 &copy; ACTCMS.com&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; 软件版权所有 &copy; ACTCMS.com</p>
            <p>
	</td>
  </tr>
  <tr>
    <td>
<div class="readpact boxcenter">
	<input name="readpact" type="checkbox" id="readpact" value="" /><label for="readpact"><strong>我已经阅读并同意此协议</strong></label>
</div>
<div class="butbox boxcenter">
	<input type="button" class="nextbut" value="" onclick="document.getElementById('readpact').checked ?window.location.href='index.asp?A=2' : alert('您必须同意软件许可协议才能安装！');" />
</div>
	</td>
  </tr>
</table>
<%End Sub 
	Public Function AutoDomain()
		Dim TempPath
		If Request.ServerVariables("SERVER_PORT") = "80" Then
			AutoDomain = Request.ServerVariables("SERVER_NAME")
		Else
			AutoDomain = Request.ServerVariables("SERVER_NAME") & ":" & Request.ServerVariables("SERVER_PORT")
		End If
		 If Instr(UCASE(AutoDomain),"/W3SVC")<>0 Then
			   AutoDomain=Left(AutoDomain,Instr(AutoDomain,"/W3SVC"))
		 End If
		 AutoDomain = "http://" & AutoDomain
	End Function
Sub seting()
	dim strDir,strAdminDir,InstallDir
	strDir=Trim(request.ServerVariables("SCRIPT_NAME"))
	strAdminDir=split(strDir,"/")(Ubound(split(strDir,"/"))-1) & "/"
	InstallDir=left(strDir,instr(lcase(strDir),"/"&Lcase(strAdminDir)))
	
	If Instr(UCASE(InstallDir),"/W3SVC")<>0 Then
	   InstallDir=Left(InstallDir,Instr(InstallDir,"/W3SVC"))
	End If

 %>
<form id="form1" name="form1" method="post" action="index.asp?A=3">
  <table width="700" border="0" align="center" cellpadding="0" cellspacing="0"  class="twbox">
    <tr  onmouseover="overColor(this)" onmouseout="outColor(this)">
      <td colspan="2" align="center">ACTCMS 4.0 正式版-系统安装-请仔细阅读说明,然后进行安装</td>
    </tr>
    <tr onmouseover="overColor(this)" onmouseout="outColor(this)">
      <td width="260" align="right"><strong>网站名称：</strong></td>
      <td width="528"><input name="webname" type="text" id="webname" value="网站名称" class="textipt" style="width:150px" /></td>
    </tr>

    <tr onmouseover="overColor(this)" onmouseout="outColor(this)">
      <td width="260" align="right"><strong>网站标题：</strong></td>
      <td width="528"><input name="webtitle" type="text" id="webtitle" value="网站标题"  class="textipt" style="width:150px" /></td>
    </tr>

    <tr onmouseover="overColor(this)" onmouseout="outColor(this)">
      <td width="260" align="right"><strong>站长姓名：</strong></td>
      <td width="528"><input name="webadmin" type="text" id="webadmin" value="左岸"  class="textipt" style="width:150px" /></td>
    </tr>

    <tr onmouseover="overColor(this)" onmouseout="outColor(this)">
      <td width="260" align="right"><strong>管理员信箱：</strong></td>
      <td><input name="adminmail" type="text" id="adminmail" class="textipt" value="test@actcms.com" style="width:150px" /></td>
    </tr>
    <tr onmouseover="overColor(this)" onmouseout="outColor(this)">
      <td width="260" align="right"><strong>网站网址：</strong></td>
      <td><input name="AutoDomain" type="text" class="textipt" id="AutoDomain" style="width:150px" value="<%=AutoDomain%>" /></td>
    </tr>
    <tr onmouseover="overColor(this)" onmouseout="outColor(this)">
      <td width="260" align="right"><strong>CMS安装目录：</strong></td>
      <td><input name="InstallDir" type="text" class="textipt" id="InstallDir" style="width:150px" value="<%=InstallDir%>" />一般按默认就可以</td>
    </tr>
    <tr onmouseover="overColor(this)" onmouseout="outColor(this)">
      <td width="260" align="right"><strong>后台目录：</strong></td>
      <td><input name="admindir" type="text" class="textipt" id="admindir" style="width:150px" value="admin" />
	  请确认是否有该文件夹</td>
    </tr>

    <tr onmouseover="overColor(this)" onmouseout="outColor(this)">
      <td width="260" align="right"><strong>缓存名称：</strong></td>
      <td><input name="webhc" type="text" class="textipt" id="webhc" style="width:150px" value="ACTCMS<%=Left(UCase(MD5(Now())),5)%>" />
	  系统缓存名称.在一个URL下安装多个ACTCMS请设置不同名称</td>
    </tr>

	
<tr class="td1">
      <td width="260" align="right"><strong>数据库类型：</strong></td>
      <td>
		<input Checked id="sjk1" type="radio" name="sjk" value="access"  onClick=sjklx('access')>
		<label for="sjk1"><font color="green">ACCESS &nbsp;</font></label>
		<input  id="sjk2" type="radio" name="sjk" value="mssql"  onClick=sjklx('mssql')>
		<label for="sjk2"><font color="red">MSSQL &nbsp;</font></label>	 
        <span id="mysql5"><a href="http://www.winmay.com/" target="_blank" title="新窗口打开.请放心">要安装SQL首先要有一个SQL数据库</a></span> </td>
    </tr>    

	<tr id="access" onmouseover="overColor(this)" onmouseout="outColor(this)">
      <td width="260" align="right"><strong>数据库名称：</strong></td>
      <td><input name="webdb" type="text" class="textipt" id="webdb" style="width:150px" value="#ACTCMS<%=Left(UCase(MD5(Now())),5)%>.mdb" />
	  请填写数据库文件名(不要填路径)</td>
    </tr>

	<tr id="mysql1" onmouseover="overColor(this)" onmouseout="outColor(this)">
      <td width="260" align="right"><strong>数据库服务器：</strong></td>
      <td><input name="server" type="text" class="textipt" id="webdb" style="width:150px" value="(local)" /> 
        <a href="http://www.actcms.com/html/Help/ChangJianWenTi/103.html" title="新窗口打开.请放心" target="_blank">SQL安装不明白的请点击这里      </a></td>
    </tr>


	<tr id="mysql2" onmouseover="overColor(this)" onmouseout="outColor(this)">
      <td width="260" align="right"><strong>数据库名称：</strong></td>
      <td><input name="sqlname" type="text" class="textipt" id="webdb" style="width:150px" value="" />
      数据库名称,请先在SQL上创建一个数据库
	  </td>
    </tr>

	<tr id="mysql3" onmouseover="overColor(this)" onmouseout="outColor(this)">
      <td width="260" align="right"><strong>数据库用户名：</strong></td>
      <td><input name="sqluser" type="text" class="textipt" id="webdb" style="width:150px" value="sa" />
      数据库连接账号
	  </td>
    </tr>


	<tr id="mysql4" onmouseover="overColor(this)" onmouseout="outColor(this)">
      <td width="260" align="right"><strong>数据库密码：</strong></td>
      <td><input name="sqlps" type="text" class="textipt" id="webdb" style="width:150px" value="" />
      数据库连接密码
	  </td>
    </tr>


    <tr onmouseover="overColor(this)" onmouseout="outColor(this)">
      <td width="260" align="right"><strong>后台登陆验证码：</strong></td>
      <td>
		<input Checked id="yzm1" type="radio" name="yzm" value="true">
		<label for="yzm1"><font color="green">正常 &nbsp;</font></label>
		<input  id="yzm2" type="radio" name="yzm" value="false">
		<label for="yzm2"><font color="red">关闭 &nbsp;</font></label>	  </td>
    </tr>

	<tr onmouseover="overColor(this)" onmouseout="outColor(this)">
      <td width="260" align="right"><strong>后台登陆认证码：</strong></td>
      <td>
		<input Checked id="rzm1" type="radio" name="rzm" value="true"  onclick="glrzms.style.display='';" >
		<label for="rzm1"><font color="green">正常 &nbsp;</font></label>
		<input  id="rzm2" type="radio" name="rzm" value="false"  onclick="glrzms.style.display='none';">
		<label for="rzm2"><font color="red">关闭 &nbsp;</font></label>	  </td>
    </tr>


    <tr id="glrzms" onmouseover="overColor(this)" onmouseout="outColor(this)">
      <td width="260" align="right"><strong>管理认证码：</strong></td>
      <td>
	  <input name="glrzm" type="text" class="textipt" id="glrzm" style="width:150px" value="actcms" />
	  请修改,这样即使有人知道了后台用户名和密码也不能登录后台</td>
    </tr>
    <tr onmouseover="overColor(this)" onmouseout="outColor(this)">
      <td align="right"><strong>安全码：</strong></td>
      <td><input name="Security" type="text" class="textipt" id="webdb2" style="width:150px" value="" />
        后台进行危险操作的时候.需要使用到安全码,必须填写</td>
    </tr>

	<tr onmouseover="overColor(this)" onmouseout="outColor(this)">
      <td width="260" align="right"><strong>登录帐号：</strong></td>
      <td><input name="loginname" type="text" id="loginname" class="textipt" value="admin"  style="width:150px" /></td>
    </tr>
    <tr onmouseover="overColor(this)" onmouseout="outColor(this)">
      <td align="right"><strong>登录密码：</strong></td>
      <td><input name="password" type="text"  class="textipt" value="" style="width:150px" /></td>
    </tr>

  <tr onmouseover="overColor(this)" onmouseout="outColor(this)">
      <td align="right"><strong>体验数据：</strong></td>
      <td><a href="javascript:GetRemoteDemo()">远程获取</a> <span id="testdatatitle"></span></td>
    </tr>

  <tr onmouseover="overColor(this)" onmouseout="outColor(this)">
      <td align="right"><strong>安装测试体验数据：</strong></td>
      <td>安装体验数据<input id="testdata" name="testdata" type="checkbox" value="1" disabled=true></td>
    </tr>


    <tr onmouseover="overColor(this)" onmouseout="outColor(this)">
      <td colspan="2" align="right">
	  <div class="butbox boxcenter">
	<input type="button" class="backbut" value="" onclick="history.back();" style="margin-right:20px" />
	<input   name=Submit1   type="button" class="setupbut"  onclick=CheckForm()  value="" />
</div></td>
    </tr>
  </table>
</form>
<script language="JavaScript" type="text/javascript">

 function CheckForm()
	{ var form=document.form1;
 		

		
	 if (form.password.value=='')
		{ alert("请输入密码!");   
		  form.Security.focus();    
		   return false;
		}

	 if (form.Security.value=='')
		{ alert("请输入安全码!");   
		  form.Security.focus();    
		   return false;
		}
 	 form.Submit1.disabled=true;	
     form.submit();
        return true;
	}	

function overColor(Obj)
{
	var elements=Obj.childNodes;
	for(var i=0;i<elements.length;i++)
	{
		elements[i].className="td1"
		Obj.bgColor="";
	}
	
}	
	function GetRemoteDemo()
	{
  	var urldata=lhgajax.send("data.asp?m="+Math.random(),"GET");
  	switch (urldata)
	{
		 case "err":
			 alert('下载失败');
 			   document.getElementById("testdatatitle").innerHTML = "<font color=red>不能安装演示数据</font>";
			 break;
		 case "OK":
 			   document.getElementById("testdatatitle").innerHTML = "<font color=green>可以安装演示数据</font>";
			    document.form1.testdata.disabled=false;
				 break;
		  default:
  			   document.getElementById("testdatatitle").innerHTML = "<font color=red>不能安装演示数据</font>";
				alert("未知错误");
 	}
 }


function sjklx(n){
	if (n == "access"){
 		access.style.display='';
 		mysql1.style.display='none';
 		mysql2.style.display='none';
 		mysql3.style.display='none';
 		mysql4.style.display='none';
 		mysql5.style.display='none';
 	}
 	else
	{	 
 		mysql1.style.display='';
 		mysql2.style.display='';
 		mysql3.style.display='';
 		mysql4.style.display='';
 		mysql5.style.display='';
		access.style.display='none';
	
	}
}

function outColor(Obj)
{
	var elements=Obj.childNodes;
	for(var i=0;i<elements.length;i++)
	{
		elements[i].className="td2";
		Obj.bgColor="";
	}
}
</script>
<script language="javascript">sjklx("access");</script>

<%
End Sub 
%>
</body>
</html>
