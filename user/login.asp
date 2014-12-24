<!--#include file="../act_inc/ACT.User.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>用户管理登录</title>
<link rel="stylesheet" href="Login.css" type="text/css" />
</head>
<body>


<div id="Header">
	<div id="logo" title="用户管理登录">用户管理登录</div>
	<ul id="menu">
		<li>
			<a href="index.asp">首页</a>&nbsp;|&nbsp;
			<a href="reg.asp">注册</a>
		</li>
	</ul>
</div>
<div id="Container">
	<div id="PageBody">
		<div class="Sidebar">
			<form name="ActCMS" method="post" action="Checklogin.asp?Action=LoginCheck">
				<ul>
					<li><label>用户名：<input type="text" name="UserName" id="UserName" onFocus="this.className='input_onFocus'" onBlur="this.className='input_onBlur'" value="" /><input name="act" type="hidden" id="act" value="cool">
</label></li>
					<li><label>密　码：<input type="password" name="PassWord" id="Password" onFocus="this.className='input_onFocus'" onBlur="this.className='input_onBlur'" /></label></li>
<% if ACTCMS.ActCMS_Sys(15)=0 Then%>
  <LI><LABEL>验证码： <input name="Code" id="codestr" type="text" class="put2" size="6" maxlength="4" onFocus="this.className='input_onFocus'" onBlur="this.className='input_onBlur'" />
  <img src="../act_inc/code.asp?s='+Math.random();" alt="验证码" title="看不清楚? 换一张！" style="cursor:hand;" onClick="src='../act_inc/code.asp?s='+Math.random()"/> </LABEL>
<%end if %>
</label></li>
					
					<li class="CookieDate"><label for="CookieDate"><input type="checkbox" name="CookieDate" id="CookieDate" value="3" />保存我的登录信息</label></li>
					<li><input type="hidden" name="fromurl" value=""><input name="Submit" id="Submit" onclick="return CheckForm()" type="submit" value="登　录" /><a href="GetPass.asp">忘记密码？</a></li>
					<li class="hr"></li>
					<li>如果你不是本站会员，请注册</li>
					<li class="regbt"><a href="reg.asp"><img src="images/reg.jpg" /></a></li>
				</ul>
			</form>
			<ul class="help">
				<li>如果你密码丢失或原有用户名登录不了，请试试<a href="GetPass.asp">找回密码</a>。</li>
				
				<li>当你看不清验证码时请点验证码图片刷新。</li>
				
			</ul>
		</div>
		<div class="MainBody">
			<div class="ad">稳定的平台，完善的功能，满意的服务，和谐的环境。</div>
			<dl class="d1">
				<dt>发布网络文章</dt>
				<dd>在网络中用文字记录您的日常生活</dd>
			</dl>
			<dl class="d2">
				<dt>共享您的照片</dt>
				<dd>保存和共享您的照片，用光和影展现您的生活</dd>
			</dl>
			<dl class="d3">
				<dt>展示个性的您</dt>
				<dd>您可自由设置空间，展示一个独一无二的自我</dd>
			</dl>
		</div>
		<div class="clear"></div>
	</div>
	<div class="clear"></div>
</div>
<div id="Footer"><center>Copyright by ActCMS<br>
</center>
</div>
<SCRIPT language=javascript>
<!--
function CheckForm()
{ 
var form=document.ActCMS;
   if (form.UserName.value=='')
    { alert("请输入用户名!");   
	  form.UserName.focus();    
	   return false;
    }
    if (form.PassWord.value=='')  
	  { alert("请输入密码!");   
	    form.PassWord.focus();   
		  return false;   
		 }  
	<% if  ACTCMS.ActCMS_Sys(15)=0 Then%>	
	if (isNaN(form.Code.value) || form.Code.value.length!=4)
	  { alert("请输入正确的验证码!");   
	    form.Code.focus();   
		  return false;   
		 } 
	<% end if %> 
	    form.Submit.value="登　录";
		form.Submit.disabled=true;	
	    form.submit();
        return true;
}
//-->
</SCRIPT>

</body>
</html>
