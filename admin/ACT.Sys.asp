<!--#include file="ACT.Function.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Act内容管理系统</title>
<link href="Images/style.css" rel="stylesheet" type="text/css">
 <script language="JavaScript" type="text/javascript" src="../ACT_inc/js/lhgcore/lhgcore.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../ACT_inc/js/lhgcore/lhgdialog.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../ACT_inc/js/lhgcore/Main.js"></script>
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
.STYLE2 {color: #FF6600}
-->
</style></head>
<body>
<%
 If Not ACTCMS.ChkAdmin() Then '超级管理员检测
	Call Actcms.ACTCMSErr("")
 End If 
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
	dim strDir,strAdminDir,InstallDir
	strDir=Trim(request.ServerVariables("SCRIPT_NAME"))
	strAdminDir=split(strDir,"/")(Ubound(split(strDir,"/"))-1) & "/"
	InstallDir=left(strDir,instr(lcase(strDir),"/"&Lcase(strAdminDir)))
	
	If Instr(UCASE(InstallDir),"/W3SVC")<>0 Then
	   InstallDir=Left(InstallDir,Instr(InstallDir,"/W3SVC"))
	End If
	

Set Rs=server.CreateObject("adodb.recordset") 
Dim Rs,ActCMS_SysSetting,ActCMS_OtherSetting,ActCMS_WatermarkSetting,MakeJs,ActCMS_Upfile,UploadSetting
Rs.OPen "Select * from Config_ACT",Conn,1,3
IF Request.QueryString("Action") = "Save" Then
dim i,SysSetting,OtherSetting,WatermarkSetting
For I=0 To 27
	SysSetting=SysSetting& Replace(request.Form("ActCMS_SysSetting" & I &""),"^@$@^","") & "^@$@^"
Next
For I=0 To 10
	OtherSetting=OtherSetting& Replace(request.Form("ActCMS_OtherSetting" & I &""),"^@&@^","") & "^@&@^"
Next
 

For I=0 To 17
	ActCMS_Upfile=ActCMS_Upfile&  Replace(request.Form("UploadSetting(" & I &")"),"^@*&*@^","") & "^@*&*@^"
Next
 

 
IF Ubound(Split(ActCMS_Upfile,"^@*&*@^"))<>18 Or Ubound(Split(SysSetting,"^@$@^"))<>28  Or Ubound(Split(OtherSetting,"^@&@^"))<>11 then
response.Write "数据获取错误.请不要外部提交!"
response.end
End If
		Rs("ActCMS_SysSetting")=SysSetting
		Rs("ActCMS_OtherSetting")=OtherSetting
		Rs("ActCMS_Upfile")=ActCMS_Upfile
		Rs.Update:Application.Contents.RemoveAll
 		Rs.Close:Set Rs = Nothing	
		response.Redirect "ACT.Sys.asp"
  Else
ActCMS_SysSetting=Split(Rs("ActCMS_SysSetting"),"^@$@^")
ActCMS_OtherSetting=Split(Rs("ActCMS_OtherSetting"),"^@&@^")
UploadSetting=Split(Rs("ActCMS_Upfile"),"^@*&*@^")
 %>
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
   <form id="LoveShe" name="LoveShe" method="post" action="?action=Save">
 <tr>
      <td colspan="2" class="bg_tr">您现在的位置：&gt;&gt;网 站 信 息 配 置 </td>
    </tr>
    <tr>
      <td width="340" align="right">网站名称：</td>
      <td width="878"><input name="ActCMS_SysSetting0" type="text" class="ainput" id="SiteName" value="<%= ActCMS_SysSetting(0) %>" size="40">
	  <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_wzmc')" id="ACTSys_wzmc">帮助</span></td>
    </tr>
    <tr>
      <td width="340" align="right"><p>网站标题：</p></td>
      <td><input name="ActCMS_SysSetting1" type="text" class="ainput" id="SiteTitle" value="<%= ActCMS_SysSetting(1) %>" size="40">
	  <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_wzbt')"  id="ACTSys_wzbt">帮助</span></td>
    </tr>
    <tr>
      <td width="340" align="right">网站地址： </td>
      <td><input name="ActCMS_SysSetting2" type="text" class="ainput" id="SiteURL" value="<%=AutoDomain %>" size="40">
	  <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_wzdz')"   id="ACTSys_wzdz">帮助</span>
(请使用http://标识),后面不要带&quot;/&quot;符号 <br><font color=red>系统会自动获得正确的路径，但需要手工保存设置</font>   <font color=green>如果换了域名,请在基本设置里重新保存下</font></td>
    </tr>
    <tr>
      <td width="340" align="right">安装目录：</td>
      <td><input name="ActCMS_SysSetting3" type="text" class="ainput" id="SysDir" value="<%= InstallDir %>" size="40">
	  <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_azml')"   id="ACTSys_azml">帮助</span>
      系统安装的虚拟目录,后面要加上/<br><font color=red>系统会自动获得正确的路径，但需要手工保存设置</font></td>
    </tr>
    <tr>
      <td width="340" align="right">首页文件名：</td>
      <td><input name="ActCMS_SysSetting4" type="text" class="ainput" id="SiteIndex" value="<%= ActCMS_SysSetting(4) %>" size="40">
	  <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_sywjm')"   id="ACTSys_sywjm">帮助</span></td>
    </tr>
    <tr>
      <td width="340" align="right">网站Logo地址：</td>
      <td><input name="ActCMS_SysSetting5" type="text" class="ainput" id="SiteLogo" value="<%= ActCMS_SysSetting(5) %>" size="40">
	  <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_logo')"   id="ACTSys_logo">帮助</span></td>
    </tr>
    <tr>
      <td width="340" align="right">站长姓名：</td>
      <td><input name="ActCMS_SysSetting6" type="text" class="ainput" id="WebmasterName" value="<%= ActCMS_SysSetting(6) %>" size="40">
	  <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_zzxm')"   id="ACTSys_zzxm">帮助</span></td>
    </tr>
    <tr>
      <td width="340" align="right">站长信箱：</td>
      <td><input name="ActCMS_SysSetting7" type="text" class="ainput" id="WebmasterMail" value="<%= ActCMS_SysSetting(7) %>" size="40">
	  <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_zzxx')"   id="ACTSys_zzxx">帮助</span></td>
    </tr>
    <tr>
      <td width="340" align="right">后台目录：</td>
      <td><input name="ActCMS_SysSetting8" type="text" class="ainput"  value="<%= ActCMS_SysSetting(8) %>" size="40">
	  <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_html')"  id="ACTSys_html">帮助</span></td>
    </tr>
    <tr>
      <td width="340" align="right">首页模板：</td>
      <td>
	  
	   <input name="ActCMS_SysSetting9" type="text" class="ainput" id="ActCMS_SysSetting9"  value="<%= ActCMS_SysSetting(9) %>" size="40">
          <input class="ACT_btn" type="button"  onClick="OpenWindowAndSetValue('include/print/SelectPic.asp?CurrPath=<%=ACTCMS.ActSys%><%=actcms.SysThemePath&"/"&actcms.NowTheme%>',500,320,window,document.LoveShe.ActCMS_SysSetting9);" value="选择模板..."> 
			<span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_symb')"  id="ACTSys_symb">帮助</span>
	</td></tr>
    <tr>
      <td width="340" align="right">上传文件大小： </td>
      <td><input name="ActCMS_SysSetting10" type="text" class="ainput" id="ActCMS_SysSetting10"  value="<%= ActCMS_SysSetting(10) %>" size="40">
      <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_scwjdx')"  id="ACTSys_scwjdx">帮助</span>单位(KB)</td>
    </tr>
    <tr>
      <td width="340" align="right">上传文件类型： </td>
      <td><input name="ActCMS_SysSetting11" type="text" class="ainput" id="ActCMS_SysSetting11"  value="<%= ActCMS_SysSetting(11) %>" size="40">
      <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_scwjlx')"  id="ACTSys_scwjlx">帮助</span>请用“/”分开</td>
    </tr>

    <tr>
      <td width="340" align="right">模板文件夹名称： </td>
      <td><input name="ActCMS_SysSetting19" type="text" class="ainput" id="ActCMS_SysSetting19"  value="<%= ActCMS_SysSetting(19) %>" size="40">
      <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_mblj')"  id="ACTSys_mblj">帮助</span></td>
    </tr>


   
    <tr>
      <td width="340" align="right">备案号： </td>
      <td><input name="ActCMS_SysSetting26" type="text" class="ainput" id="ActCMS_SysSetting26"  value="<%= ActCMS_SysSetting(26) %>" size="40">
      <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_bah')"  id="ACTSys_bah">帮助</span></td>
    </tr>
	
	
 <tr>
      <td width="340" align="right">统计代码： </td>
      <td><textarea name="ActCMS_SysSetting27" cols="50" rows="5"><%= ActCMS_SysSetting(27) %></textarea>
      <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_tjdm')"  id="ACTSys_tjdm">帮助</span>添加你的统计代码</td>
    </tr>



      <tr align="right">
        <td colspan="2" align="center" class="bg_tr">会员信息设置</td>
     </tr>
      <tr>
        <td width="340" align="right">会员系统状态：</td>
        <td><input <% IF  ActCMS_SysSetting(12)  = "0" Then Response.Write "Checked" %>  type="radio"  id="Userclose1" name="ActCMS_SysSetting12" value="0">
        <label for="Userclose1"><font color="green">正常 &nbsp;</font></label>
            <input  <% IF ActCMS_SysSetting(12) = "1" Then Response.Write "Checked" %>  id="Userclose2" type="radio" name="ActCMS_SysSetting12" value="1">
<label for="Userclose2"><font color="red">关闭 &nbsp;</font></label>
<span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_hyxtzt')"  id="ACTSys_hyxtzt">帮助</span>       </td>
      </tr>
      <tr>
        <td width="340" align="right">是否允许新会员注册：</td>
        <td><input <% IF ActCMS_SysSetting(13) = "0" Then Response.Write "Checked" %>  type="radio"  id="UserReg1" name="ActCMS_SysSetting13" value="0">
        <label for="UserReg1"><font color="green">正常 &nbsp;</font></label>
            <input  <% IF ActCMS_SysSetting(13) = "1" Then Response.Write "Checked" %>  id="UserReg2" type="radio" name="ActCMS_SysSetting13" value="1">
<label for="UserReg2"><font color="red">关闭 &nbsp;</font></label>
<span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_sfxhyzc')"  id="ACTSys_sfxhyzc">帮助</span> </td>
      </tr>
      <tr>
        <td width="340" align="right">会员注册时是否启用验证码功能：</td>
        <td><input <% IF ActCMS_SysSetting(14) = "0" Then Response.Write "Checked" %>  type="radio"  id="RegCode" name="ActCMS_SysSetting14" value="0"><label for="RegCode">是 &nbsp;</label>
          <input  <% IF ActCMS_SysSetting(14) = "1" Then Response.Write "Checked" %>  id="RegCode2" type="radio" name="ActCMS_SysSetting14" value="1"><label for="RegCode2">否 &nbsp;</label>
		  <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_yzm')"  id="ACTSys_yzm">帮助</span> </td>
      </tr>
      <tr>
        <td width="340" align="right">会员登录时是否启用验证码功能：</td>
        <td><input <% IF ActCMS_SysSetting(15) = "0" Then Response.Write "Checked" %>  type="radio"  id="LoginCode" name="ActCMS_SysSetting15" value="0"><label for="LoginCode">是 &nbsp;</label>
          <input  <% IF ActCMS_SysSetting(15) = "1" Then Response.Write "Checked" %>  id="LoginCode2" type="radio" name="ActCMS_SysSetting15" value="1"><label for="LoginCode2">否 &nbsp;</label>
		  <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_dlyzm')"  id="ACTSys_dlyzm">帮助</span> </td>
      </tr>
    
	  
      <tr>
        <td width="340" align="right">注册协议模板：</td>
        <td>

	  <input name="ActCMS_SysSetting17" type="text" class="ainput" id="ActCMS_SysSetting17"  value="<%= ActCMS_SysSetting(17) %>" size="40">
          <input class="ACT_btn" type="button"  onClick="OpenWindowAndSetValue('include/print/SelectPic.asp?CurrPath=<%=ACTCMS.ActSys%><%=actcms.SysThemePath&"/"&actcms.NowTheme%>',500,320,window,document.LoveShe.ActCMS_SysSetting17);" value="选择模板..."> 

		  <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_zcxymb')"  id="ACTSys_zcxymb">帮助</span> </td>
      </tr>
    
      <tr>
        <td width="340" align="right">注册页模板：</td>
        <td>

	  <input name="ActCMS_SysSetting18" type="text" class="ainput" id="ActCMS_SysSetting18"  value="<%= ActCMS_SysSetting(18) %>" size="40">
          <input class="ACT_btn" type="button"  onClick="OpenWindowAndSetValue('include/print/SelectPic.asp?CurrPath=<%=ACTCMS.ActSys%><%=actcms.SysThemePath&"/"&actcms.NowTheme%>',500,320,window,document.LoveShe.ActCMS_SysSetting18);" value="选择模板..."> 

		  <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_zcymb')"  id="ACTSys_zcymb">帮助</span> </td>
      </tr>
    
	  <tr>
        <td width="340" align="right">禁止注册的用户名： </td>
        <td><textarea name="ActCMS_SysSetting16" cols="50" rows="3" id="ActCMS_SysSetting16"><%=ActCMS_SysSetting(16) %></textarea>
            
         <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_jzzcymh')"  id="ACTSys_jzzcymh">帮助</span>  <br>在上边指定的用户名将被禁止注册，每个用户名请用“|”符号分隔</td>
      </tr>



  <tr>
        <td width="340" align="right">会员的积分与点券的兑换比率： </td>
        <td><input name="ActCMS_SysSetting20" type="text" class="ainput" id="ActCMS_SysSetting20"  value="<%= ActCMS_SysSetting(20) %>" size="40"></td>
      </tr>



	    <tr>
        <td width="340" align="right">会员的积分与有效期的兑换比率： </td>
        <td><input name="ActCMS_SysSetting21" type="text" class="ainput" id="ActCMS_SysSetting21"  value="<%= ActCMS_SysSetting(21) %>" size="40"></td>
      </tr>



	    <tr>
        <td width="340" align="right">会员的资金与点券的兑换比率： </td>
        <td><input name="ActCMS_SysSetting22" type="text" class="ainput" id="ActCMS_SysSetting22"  value="<%= ActCMS_SysSetting(22) %>" size="40"></td>
      </tr>

	    <tr>
        <td width="340" align="right">会员的资金与有效期的兑换比率： </td>
        <td><input name="ActCMS_SysSetting23" type="text" class="ainput" id="ActCMS_SysSetting23"  value="<%= ActCMS_SysSetting(23) %>" size="40"></td>
      </tr>

  <tr>
        <td width="340" align="right">点券设置： </td>
        <td><input name="ActCMS_SysSetting24" type="text" class="ainput" id="ActCMS_SysSetting24"  value="<%= ActCMS_SysSetting(24) %>" size="40">
       </td>
      </tr>

  <tr>
        <td width="340" align="right"><%= ActCMS_SysSetting(24) %>单位： </td>
        <td><input name="ActCMS_SysSetting25" type="text" class="ainput" id="ActCMS_SysSetting25"  value="<%= ActCMS_SysSetting(25) %>" size="40">
       </td>
      </tr>

      <tr align="right">
        <td colspan="2" class="bg_tr">其他信息设置</td>
    </tr>
    <tr>
      <td width="340" align="right">版权信息：</td>
      <td><textarea name="ActCMS_OtherSetting0" cols="50" rows="5"><%=ActCMS_OtherSetting(0)%></textarea>
	  <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_bqxx')"  id="ACTSys_bqxx">帮助</span> </td>
    </tr>
    <tr>
      <td width="340" align="right">网站META关键词： </td>
      <td><textarea name="ActCMS_OtherSetting1" cols="50" rows="5"><%=ActCMS_OtherSetting(1)%></textarea>
	  <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_metagjz')"  id="ACTSys_metagjz">帮助</span> </td>
    </tr>
    <tr>
      <td width="340" align="right">网站META网页描述： </td>
      <td><textarea name="ActCMS_OtherSetting2" cols="50" rows="5"><%=ActCMS_OtherSetting(2)%></textarea>
	  <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_metams')"  id="ACTSys_metams">帮助</span> </td>
    </tr>
    <tr>
      <td width="340" align="right">过滤字符：<br></td>
      <td><textarea name="ActCMS_OtherSetting8" cols="50" rows="5"><%=ActCMS_OtherSetting(8)%></textarea>
	  <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_glzf')"  id="ACTSys_glzf">帮助</span> 
      用户名用“|”符号分隔</td>
    </tr>
    <tr>
      <td width="340" align="right">SMTP服务器地址: </td>
      <td><input name="ActCMS_OtherSetting3" type="text" class="ainput" value="<%=ActCMS_OtherSetting(3)%>" size="18">
      <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_smtpdz')"  id="ACTSys_smtpdz">帮助</span> 
	  用来发送邮件的SMTP服务器如果你不清楚此参数含义，请联系你的空间商</td>
    </tr>
    <tr>
      <td width="340" align="right">SMTP登录用户名： </td>
      <td><input name="ActCMS_OtherSetting4" type="text" class="ainput" value="<%=ActCMS_OtherSetting(4)%>" size="18">
      <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_smtpdlm')"  id="ACTSys_smtpdlm">帮助</span> 当你的服务器需要SMTP身份验证时还需设置此参数</td>
    </tr>
    <tr>
      <td width="340" align="right">SMTP登录密码：</td>
      <td><input name="ActCMS_OtherSetting5" type="PassWord" class="ainput"  value="<%=ActCMS_OtherSetting(5)%>" size="20">
        <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_smtpmm')"  id="ACTSys_smtpmm">帮助</span> 当你的服务器需要SMTP身份验证时还需设置此参数</td>
    </tr>
    <tr>
      <td width="340" align="right">远程图片保存目录：</td>
      <td><input name="ActCMS_OtherSetting6" type="text" class="ainput" id="ActCMS_OtherSetting6" value="<%=ActCMS_OtherSetting(6)%>" size="18">
	  <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_yctp')"  id="ACTSys_yctp">帮助</span> </td>
    </tr>
    <tr>
      <td width="340" align="right">是否以日期保存：</td>
<td><input <% IF ActCMS_OtherSetting(7) = "0" Then Response.Write "Checked" %>  type="radio"  id="ActCMS_OtherSetting7" name="ActCMS_OtherSetting7" value="0"><label for="ActCMS_OtherSetting7">是 &nbsp;</label>
          <input  <% IF ActCMS_OtherSetting(7) = "1" Then Response.Write "Checked" %>  id="ActCMS_OtherSetting7)2" type="radio" name="ActCMS_OtherSetting7" value="1"><label for="ActCMS_OtherSetting7)2">否 &nbsp;</label>
		  <span class="h" style="cursor:help;"  onclick="dohelp('ACTSys_rqbc')"  id="ACTSys_rqbc">帮助</span> </td>
    </tr>

    <tr>
      <td width="340" align="right">生成方式：</td>
<td><input <% IF ActCMS_OtherSetting(9) = "0" Then Response.Write "Checked" %>  type="radio"  id="ActCMS_OtherSetting9" name="ActCMS_OtherSetting9" value="0"><label for="ActCMS_OtherSetting9">绝对路径 &nbsp;</label>
          <input  <% IF ActCMS_OtherSetting(9) = "1" Then Response.Write "Checked" %>  id="ActCMS_OtherSetting9)2" type="radio" name="ActCMS_OtherSetting9" value="1"><label for="ActCMS_OtherSetting9)2">根相对路径 (相对根目录) &nbsp;</label>
		如果绑定了子域名,请选择绝对路径</td>
    </tr>
    <tr>
      <td width="340" align="right">FSO组件的名称：</td>
      <td>
	  <input name="ActCMS_OtherSetting10" type="text" class="ainput" id="ActCMS_OtherSetting10" value="<%=ActCMS_OtherSetting(10)%>" size="30">
	</td>
    </tr>

	
	
	<tr align="right">
      <td colspan="2" align="center"  class="bg_tr">缩略图水印设置</td>
    </tr>
 <div id="Issubport0" style="display:none">请选择EMAIL组件！</div>
<div id="Issubport999" style="display:none"></div>
 <% Dim InstalledObjects(12)
InstalledObjects(1) = "JMail.Message"				'JMail 4.3
InstalledObjects(2) = "CDONTS.NewMail"				'CDONTS
InstalledObjects(3) = "Persits.MailSender"			'ASPEMAIL
'-----------------------
InstalledObjects(4) = "Adodb.Stream"				'Adodb.Stream
InstalledObjects(5) = "Persits.Upload"				'Aspupload3.0
InstalledObjects(6) = "SoftArtisans.FileUp"			'SA-FileUp 4.0
InstalledObjects(7) = "DvFile.Upload"				'DvFile-Up V1.0
'-----------------------
InstalledObjects(9) = "CreatePreviewImage.cGvbox"	'CreatePreviewImage
InstalledObjects(10) = "Persits.Jpeg"				'AspJpeg
InstalledObjects(11) = "SoftArtisans.ImageGen"		'SoftArtisans ImgWriter V1.21
InstalledObjects(12) = "sjCatSoft.Thumbnail"		'sjCatSoft.Thumbnail V2.6

For i=1 to 12
	Response.Write "<div id=""Issubport"&i&""" style=""display:none"">"
	If IsObjInstalled(InstalledObjects(i)) Then Response.Write InstalledObjects(i)&":<font color=red><b>√</b>服务器支持!</font>" Else Response.Write InstalledObjects(i)&"<b>×</b>服务器不支持!" &vbcrlf
	Response.Write "</div>"
Next


Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = actcms.iCreateObject(strClassString)
	If Err = 0 Then IsObjInstalled = True
	If Err = -2147352567 Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
 
 
 %>
 
 
<iframe width="260" height="165" id="colourPalette" src="include/selcolor.htm" style="visibility:hidden; position: absolute; left: 0px; top: 0px;border:1px gray solid" frameborder="0" scrolling="no" ></iframe>
 <SCRIPT LANGUAGE="JavaScript">

var ColorImg;
var ColorValue;
function hideColourPallete() {
	document.getElementById("colourPalette").style.visibility="hidden";
}
function Getcolor(img_val,input_val){
	var obj = document.getElementById("colourPalette");
	ColorImg = img_val;
	ColorValue = document.getElementById(input_val);
	if (obj){
	obj.style.left = getOffsetLeft(ColorImg) + "px";
	obj.style.top = (getOffsetTop(ColorImg) + ColorImg.offsetHeight) + "px";
	if (obj.style.visibility=="hidden")
	{
	obj.style.visibility="visible";
	}else {
	obj.style.visibility="hidden";
	}
	}
}
//Colour pallete top offset
function getOffsetTop(elm) {
	var mOffsetTop = elm.offsetTop;
	var mOffsetParent = elm.offsetParent;
	while(mOffsetParent){
		mOffsetTop += mOffsetParent.offsetTop;
		mOffsetParent = mOffsetParent.offsetParent;
	}
	return mOffsetTop;
}

//Colour pallete left offset
function getOffsetLeft(elm) {
	var mOffsetLeft = elm.offsetLeft;
	var mOffsetParent = elm.offsetParent;
	while(mOffsetParent) {
		mOffsetLeft += mOffsetParent.offsetLeft;
		mOffsetParent = mOffsetParent.offsetParent;
	}
	return mOffsetLeft;
}
function setColor(color)
{
	if (ColorValue){ColorValue.value = color;}
	if (ColorImg){ColorImg.style.backgroundColor = color;}
	document.getElementById("colourPalette").style.visibility="hidden";
}



<!--
function chkselect(s,divid)
{
var divname='Issubport';
var chkreport;
	s=Number(s)
	if (divid=="know1")
	{
	divname=divname+s;
	}
	if (divid=="know2")
	{
	s+=4;
	if (s==1003){s=999;}
	divname=divname+s;
	}
	if (divid=="know3")
	{
	s+=9;
	if (s==1008){s=999;}
	divname=divname+s;
	}

	if (divid=="know4")
	{
	s+=9;
	if (s==1008){s=999;}
	divname=divname+s;
	}


document.getElementById(divid).innerHTML=divname;
 chkreport=document.getElementById(divname).innerHTML;
document.getElementById(divid).innerHTML=chkreport;
}
//-->
</SCRIPT>
   
     
    
 <input type="hidden" name="UploadSetting(2)" id="UploadSetting(2)">
 


<tr> 
	<td align="right" >选取生成缩略图片组件：</td>
	<td > 
	<select name="UploadSetting(3)" id="UploadSetting(3)" onChange="chkselect(options[selectedIndex].value,'know3');">
	<option value="999">关闭
	<option value="0">CreatePreviewImage组件
	<option value="1">AspJpeg组件
	<option value="2">SA-ImgWriter组件
	</select><div id="know3"></div>	</td>
	
</tr>



<tr> 
	<td align="right">生成缩略图片大小设置（宽度|高度）：</td>
	<td>
		宽度：<INPUT type="text" class="ainput" NAME="UploadSetting(14)" size=10 value="<%=UploadSetting(14)%>"> 象素
		高度：<INPUT type="text" class="ainput" NAME="UploadSetting(15)" size=10 value="<%=UploadSetting(15)%>"> 象素	</td>
	
</tr>
<tr> 
	<td align="right" >生成缩略图片大小规则选项：</td>
	<td > 
		<SELECT name="UploadSetting(16)" id="UploadSetting(16)">
		<OPTION value=0>固定</OPTION>
		<OPTION value=1>等比例缩小</OPTION>
		</SELECT>	</td>
	
</tr>
<tr> 
	<td align="right" >图片水印组件：</td>
	<td > 
	<select name="UploadSetting(17)" id="UploadSetting(17)" onChange="chkselect(options[selectedIndex].value,'know4');">
	<option value="999">关闭
 	<option value="1">AspJpeg组件
	<option value="2">SA-ImgWriter组件
	</select><div id="know4"></div>	</td>
	
</tr>

<tr> 
	<td align="right">图片水印设置：</td>
	<td> 
		<SELECT name="UploadSetting(1)" id="UploadSetting(1)">
		<OPTION value="0">关闭水印效果</OPTION>
		<OPTION value="1">水印文字效果</OPTION>
		<OPTION value="2">水印图片效果</OPTION>
		</SELECT>	</td>
	
</tr>
<tr> 
	<td align="right" >上传图片添加水印文字信息（可为空或0）：</td>
	<td > 
	<INPUT type="text" class="ainput" NAME="UploadSetting(4)" size=40 value="<%=UploadSetting(4)%>">	</td>
	
</tr>
<tr> 
	<td align="right">上传添加水印字体大小：</td>
	<td> 
	<INPUT type="text" class="ainput" NAME="UploadSetting(5)" size=10 value="<%=UploadSetting(5)%>"> <b>px</b>	</td>
	
</tr>
<tr> 
	<td align="right" >上传添加水印字体颜色：</td>
	<td >
	<input type="hidden" name="UploadSetting(6)" id="UploadSetting(6)" value="<%=UploadSetting(6)%>">
	<img border=0 src="Images/rect.gif" style="cursor:pointer;background-Color:<%=UploadSetting(6)%>;" onClick="Getcolor(this,'UploadSetting(6)');" title="选取颜色!">	</td>
	
</tr>
<tr> 
	<td align="right">上传添加水印字体名称：</td>
	<td>
	<SELECT name="UploadSetting(7)" id="UploadSetting(7)">
	<option value="宋体">宋体</option>
	<option value="楷体">楷体</option>
	<option value="新宋体">新宋体</option>
	<option value="黑体">黑体</option>
	<option value="隶书">隶书</option>
	<OPTION value="Andale Mono" selected>Andale Mono</OPTION> 
	<OPTION value=Arial>Arial</OPTION> 
	<OPTION value="Arial Black">Arial Black</OPTION> 
	<OPTION value="Book Antiqua">Book Antiqua</OPTION>
	<OPTION value="Century Gothic">Century Gothic</OPTION> 
	<OPTION value="Comic Sans MS">Comic Sans MS</OPTION>
	<OPTION value="Courier New">Courier New</OPTION>
	<OPTION value=Georgia>Georgia</OPTION>
	<OPTION value=Impact>Impact</OPTION>
	<OPTION value=Tahoma>Tahoma</OPTION>
	<OPTION value="Times New Roman" >Times New Roman</OPTION>
	<OPTION value="Trebuchet MS">Trebuchet MS</OPTION>
	<OPTION value="Script MT Bold">Script MT Bold</OPTION>
	<OPTION value=Stencil>Stencil</OPTION>
	<OPTION value=Verdana>Verdana</OPTION>
	<OPTION value="Lucida Console">Lucida Console</OPTION>
	</SELECT>	</td>
	
</tr>
<tr> 
	<td align="right" >上传水印字体是否粗体：</td>
	<td > 
		<SELECT name="UploadSetting(8)" id="UploadSetting(8)">
		<OPTION value=0>否</OPTION>
		<OPTION value=1>是</OPTION>
		</SELECT>	</td>
	
</tr>
<!-- 上传图片添加水印LOGO图片定义 -->
<tr> 
	<td align="right">上传图片添加水印LOGO图片信息（可为空或0）：<br>填写LOGO的图片相对路径</td>
	<td> 
	<INPUT type="text" class="ainput" NAME="UploadSetting(9)" size=40 value="<%=UploadSetting(9)%>">	</td>
	
</tr>
<tr> 
	<td align="right" >上传图片添加水印透明度：</td>
	<td > 
	<INPUT type="text" class="ainput" NAME="UploadSetting(10)" size=10 value="<%=UploadSetting(10)%>"> 如60%请填写0.6	</td>
	
</tr>
<tr> 
	<td align="right" >水印图片去除底色：<br>保留为空则水印图片不去除底色</td>
	<td > 
	<INPUT type="text" class="ainput" NAME="UploadSetting(0)" ID="UploadSetting(0)" size=10 value="<%=UploadSetting(0)%>"> 
	<img border=0 src="Images/rect.gif" style="cursor:pointer;background-Color:<%=UploadSetting(0)%>;" onClick="Getcolor(this,'UploadSetting(0)');" title="选取颜色!">	</td>
	
</tr>
<tr> 
	<td align="right">水印文字或图片的长宽区域定义：<br>如水印图片的宽度和高度</td>
	<td> 
	宽度：<INPUT type="text" class="ainput" NAME="UploadSetting(11)" size=10 value="<%=UploadSetting(11)%>"> 象素
	高度：<INPUT type="text" class="ainput" NAME="UploadSetting(12)" size=10 value="<%=UploadSetting(12)%>"> 象素	</td>
	
</tr>
<tr> 
	<td align="right" >上传图片添加水印LOGO位置坐标：</td>
	<td >
	<SELECT NAME="UploadSetting(13)" id="UploadSetting(13)">
		<option value="0">左上</option>
		<option value="1">左下</option>
		<option value="2">居中</option>
		<option value="3">右上</option>
		<option value="4">右下</option>
	</SELECT>	</td>
	
</tr>
    
     <tr>
      <td width="340" align="right">&nbsp;</td>
      <td><input type=button class="ACT_btn" onclick=CheckForm()  name="Submit1" value="  保存  " />
      <input type="reset" class="ACT_btn" name="Submit2" value="  重置  "></td>
    </tr>   
    
    
 </form>
 </table>
 
 <script type="text/javascript">
function CheckSel(Voption,Value)
{
	var obj = document.getElementById(Voption);
	for (i=0;i<obj.length;i++){
		if (obj.options[i].value==Value){
		obj.options[i].selected=true;
		break;
		}
	}
}
</script>

 
 
 <SCRIPT LANGUAGE="JavaScript">
CheckSel('UploadSetting(0)','<%=UploadSetting(0)%>');
CheckSel('UploadSetting(2)','<%=UploadSetting(2)%>');
CheckSel('UploadSetting(3)','<%=UploadSetting(3)%>');
CheckSel('UploadSetting(7)','<%=UploadSetting(7)%>');
CheckSel('UploadSetting(8)','<%=UploadSetting(8)%>');
CheckSel('UploadSetting(13)','<%=UploadSetting(13)%>');
CheckSel('UploadSetting(16)','<%=UploadSetting(16)%>');
CheckSel('UploadSetting(17)','<%=UploadSetting(17)%>');
CheckSel('UploadSetting(1)','<%=UploadSetting(1)%>');
</script>   
 
 <% end if %>
<SCRIPT language=javascript>
<!--

function CheckForm()
{ var form=document.LoveShe;
   if (form.SiteName.value=='')
    { alert("请输入网站名称名称!");   
	  form.SiteName.focus();    
	   return false;
    }
	
if (form.SiteTitle.value=='')
    { alert("请输入网站标题!");   
	  form.SiteTitle.focus();    
	   return false;
    }	
	
if (form.SiteURL.value=='')
    { alert("请输入网站地址!");   
	  form.SiteURL.focus();    
	   return false;
    }	
		
if (form.SysDir.value=='')
    { alert("请输入安装目录!");   
	  form.SysDir.focus();    
	   return false;
    }	
			
	
if (form.SiteIndex.value=='')
    { alert("请输入首页文件名!");   
	  form.SiteIndex.focus();    
	   return false;
    }	
			
		form.Submit1.value="正在提交数据,请稍等...";
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

//-->
</SCRIPT>

</body>
</html>