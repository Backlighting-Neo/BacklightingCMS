<!--#include file="../act_inc/ACT.User.asp"-->
<!--#include file="../ACT_inc/md5.ASP"-->
 <% 
    dim UserHS ,A
 	Set UserHS = New ACT_User
	IF Cbool(UserHS.UserLoginChecked)=false then
	  Response.Write "<script>top.location.href ='login.asp' ;</script>"
	  Response.end
	End If	
	
		if request("A")="PassSave" then 
	
	
	
 		     Dim Oldpassword:Oldpassword=ACTCMS.S("Oldpassword")
			 Dim NewPassWord:NewPassWord=ACTCMS.S("NewPassWord")
			 Dim ReNewPassWord:ReNewPassWord=ACTCMS.S("ReNewPassWord")
			 If Oldpassword = "" Then
				 Response.Write("<script>alert('请输入旧登录密码!');history.back();</script>")
				 Response.End
              End IF
			 If NewPassWord = "" Then
				 Response.Write("<script>alert('请输入登录密码!');history.back();</script>")
				 Response.End
			 ElseIF ReNewPassWord="" Then
				 Response.Write("<script>alert('请输入确认密码');history.back();</script>")
				 Response.End
			 ElseIF NewPassWord<>ReNewPassWord Then
				 Response.Write("<script>alert('两次输入的密码不一致');history.back();</script>")
				 Response.End
			 End If
			 
			 OldPassWord =MD5(OldPassWord)
			 NewPassWord =MD5(NewPassWord)
			 
             Dim RS:Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select PassWord From User_ACT Where UserName='" & UserHS.UserName & "' And PassWord='" & OldPassWord & "'",Conn,1,3
			  IF RS.Eof And RS.Bof Then
			  	 Response.Write("<script>alert('您输入的旧密码有误！');history.back();</script>")
				 Response.End
			  Else
			     RS(0)=NewPassWord
				 RS.Update
				 Response.Cookies(AcTCMSN)("PassWord") = NewPassWord
			  End if
			 RS.Close:Set RS=Nothing
	
	end if 
	
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>会员中心</title>
 <link href="images/css.css" rel="stylesheet" type="text/css" />
  </head>
<body style="background-color:#fff">
<div id="head">
  <div id="logo"><a href="index.asp" alt=""><img src="images/logo_member.gif" alt="actcms"></a></div><div id="banner"></div>
</div>
<div id="membermenu">
<!--#include file="menu.asp"-->
</div>
<div id="main">
 <div id="left">
  <div id="treemenu">
    <h5>基本设置</h5>
   <div style="text-align:center;">
    <img src="<%If Trim(UserHS.myface)<>"" Then 
		response.write UserHS.myface
	Else 
		response.write "images/nophoto.gif" 
	End If 
	
	%> " alt="actcms" height="150" width="150"/>
	</div>    <table cellpadding="0" cellspacing="0" class="member_info">
    <tr>
      <th>用户名：</th><td><%=UserHS.username%></td>
    </tr>
    <tr>
      <th>用户组：</th><td><%=UserHS.G_Name%></td>
    </tr>
     </table>
    <ul>
       <li><a href="edit.asp">修改资料</a></li>
      <li><a href="editpwd.asp">修改密码</a></li>
    </ul>
  </div>
  <ol>
    <li class="local"><a href="<%= actcms.ActCMSDM%>">返回网站首页</a></li>
    <li class="exit"><a href="Checklogin.asp?Action=LoginOut">退出登录</a></li>
  </ol>
</div>

<div id="right">
  <p id="position"> <strong>当前位置：</strong><a href="index.asp">会员中心</a>修改密码</p>
  <form action="?A=PassSave"  method="post" name="tcjdxr" id="tcjdxr" >
    <table cellpadding="0" cellspacing="1" class="table_form">
    <caption>修改密码</caption>
      <tr>
        <th width="20%">用户名：</th>
        <td width="80%"><strong><%=UserHS.username%></strong></td>
      </tr>
      <tr>
        <th>原密码：</th>
        <td><input name="oldpassword" type="password" id="oldpassword" size="25" /></td>
      </tr>
      <tr>
        <th>新密码：</th>
        <td><input name="newpassword" type="password" id="newpassword" size="25" /></td>
      </tr>
      
    
      <tr>
        <th>确认新密码：</th>
        <td><input name="renewpassword" type="password" id="renewpassword" size="25"    /></td>
      </tr>
      <tr>
         <th></th>
        <td colspan="2"><label>
     
          
<input  class="button_style" name="Submit" type="button"  value=" 确 定 "  onclick=CheckForm() />
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input  class="button_style" name="Submit2" type="reset" value=" 重 填 " />            
          
          </label></td>
      </tr>
    </table>
  </form>
</div>
</div>
<!--#include file="foot.asp"-->
</body>
</html><script language="javascript">
 	      function CheckForm() 
 { 
 	var form=document.tcjdxr;
 	 if (form.oldpassword.value=='')
    { alert("请填写您的旧密码!");   
	  form.oldpassword.focus();    
	   return false;
    }

	
	 if (form.newpassword.value=='')
    { alert("请输入您的新密码!");   
	  form.newpassword.focus();    
	   return false;
    }
	
	 if (form.renewpassword.value=='')
    { alert("请输入您的新确认密码!");   
	  form.renewpassword.focus();    
	   return false;
    }
	
	 if (form.newpassword.value!=form.renewpassword.value)
    { alert("两次输入的密码不一致!");   
 	   return false;
    }
	
		
	
	    form.Submit.value="正在提交数据,请稍等...";
		form.Submit.disabled=true;	
		form.Submit2.disabled=true;	
	    form.submit();
        return true;		
		
	}
		
 		
 </script>
