<!--#include file="ACT.Function.right.asp"-->


<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ActCMS网站内容管理系统</title>
<link href="Images/style.css" rel="stylesheet" type="text/css">
</head>
 <% Dim theInstalledObjects(4)
theInstalledObjects(0) = "Gatherer.VBProcess"
theInstalledObjects(1) = "Scripting.FileSystemObject"
theInstalledObjects(2) = "adodb.connection"

theInstalledObjects(3) = "JMail.SMTPMail"
theInstalledObjects(4) = "CDONTS.NewMail"
Function IsObjInstalled(ByVal strClassString)
	Dim xTestObj,ClsString
	On Error Resume Next
	IsObjInstalled = False
	ClsString = strClassString
	Err = 0
	Set xTestObj = Server.CreateObject(ClsString)
	If Err = 0 Then IsObjInstalled = True
	If Err = -2147352567 Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
	Exit Function
End Function %>


<body>
<table class="table" cellspacing="1" cellpadding="2" width="100%" align="center"
        border="0">
        <tbody>
            <tr  height="25px">
                <td class="bg_tr" align=left>
                    <b>&nbsp;网站管理员公告<span style="color:red">（各位网站管理员，请您留意此处！）</span></b></td>
            </tr>
            <tr >
                <td>
                    <div align="center">
                       <%=ShowStaticContent(25) %>
				  </div>                
				  </td>
            </tr>
            
        </tbody>
</table>


<table class="table" cellspacing="1" cellpadding="2" width="100%" align="center" 
border="0">
  <tbody>
    <tr>
      <th class="bg_tr" align="left" colspan="2" height="25">系统信息统计     </th>
    </tr>
    <tr>
      <td width="50%" height="23">服务器类型：<span class="TableRow2"><%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</span></td>
      <td width="50%">脚本解释引擎：<span class="TableRow1"><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></span></td>
    </tr>
    <tr>
      <td width="50%" height="23">站点物理路径：<span class="TableRow2"><%=Request.ServerVariables("APPL_PHYSICAL_PATH")%></span></td>
      <td width="50%">ADO无组件上传：</td>
    </tr>
    <tr>
      <td width="50%" height="23">FSO文本读写：<span class="TableRow2">
        <%If Not IsObjInstalled(theInstalledObjects(1)) Then%>
        <font color="#FF0066"><b>×</b></font>
        <%else%>
        <b>√</b>
        <%end if%>
      </span></td>
      <td width="50%">数据库使用：<span class="TableRow1">
        <%If Not IsObjInstalled(theInstalledObjects(2)) Then%>
        <font color="#FF0066"><b>×</b></font>
        <%else%>
        <b style="color:blue">
          <%If DataBaseType="mssql" Then%>
          MS SQL
          <%else%>
          ACCESS
          <%end if%>
        </b>
        <%end if%>
      </span></td>
    </tr>
    <tr>
      <td width="50%" height="23">Jmail组件支持：<span class="TableRow2">
        <%If Not IsObjInstalled(theInstalledObjects(3)) Then%>
        <font color="#FF0066"><b>×</b></font>
        <%else%>
        <b>√</b>
        <%end if%>
      </span></td>
      <td width="50%">CDONTS组件支持：<span class="TableRow1">
        <%If Not IsObjInstalled(theInstalledObjects(4)) Then%>
        <font color="#FF0066"><b>×</b></font>
        <%else%>
        <b>√</b>
        <%end if%>
      </span></td>
    </tr>
  </tbody>
</table>
<table class="table" cellspacing="1" cellpadding="2" width="100%" align="center" 
border="0">
  <tbody>
    <tr>
      <th class="bg_tr" align="left" colspan="2" height="25">网站管理系统</th>
    </tr>
    <tr>
      <td width="17%" height="23">当前版本<span class="TableRow2"></span></td>
      <td width="83%"><strong>Backlighting-CMS <%= actcms.Version %></strong>
      
        <% Dim MX_Sys,ii,rs
		MX_Sys=ACTCMS.Act_MX_Sys_Arr()
		If IsArray(MX_Sys) Then
			For iI=0 To Ubound(MX_Sys,2)
  				response.Write "<a class='orange' href='ACT_Mode/ACT.Manage.asp?ModeID="&MX_Sys(0,Ii)&"&Action=ListNoAccept'>"&MX_Sys(1,Ii)&"模型待审信息(<font color=red>"&actcms.actexe("select count(id) from "&MX_Sys(2,Ii)&"  where isAccept<>0 ")(0)&"</font>)</a> &nbsp;"
 			next
		end if 

 %>
      <A class=ACT_F title=ACTCMS官方站  href="include/ACT.Map.asp"   >生成google地图</A>
      </td>
    </tr>
    <tr>
      <td width="17%" height="23">版权声明<span class="TableRow2"></span></td>
      <td width="83%">1、本软件为免费软件,可以免费使用; <br>        
        2、用户自由选择是否使用,在使用中出现任何问题和由此造成的一切损失作者将不承担任何责任; <br>      
        3、您可以对本系统进行修改和美化，但必须保留完整的版权信息;  　<br>
      4、本软件受中华人民共和国《著作权法》《计算机软件保护条例》等相关法律、法规保护，作者保留一切权利。</td>
    </tr>
  </tbody>
</table>
    <table class="table" cellspacing="1" cellpadding="2" width="100%" align="center"
        border="0">
        <tbody>
            <tr  height="25px">
                <td class="bg_tr" colspan="4" align=left>
                    <b>&nbsp;联系我们</b></td>
            </tr>
            <tr >
                <td width="13%" height="20">
                    <div align="center">
                        产品开发</div>                </td>
                <td width="30%" height="20">逆光软件工作室</td>
                <td width="12%">
                    <div align="center">
                        定制开发</div>                </td>
                <td width="45%">
                    QQ：398413222 tjtzc@126.com </td>
            </tr>
            <tr>
                <td colspan="4" height="20">
                    <div align="center">
                &copy;  2013-2015 <a href="http://www.actcms.com"  target="_blank">逆光软件</a>  Inc. All Rights Reserved</div>                </td>
            </tr>
        </tbody>
    </table>

</body>
</html>
