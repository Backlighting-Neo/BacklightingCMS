<!--#include file="ACT.Function.asp"-->
<!--#include file="../act_Inc/Md5.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ACTCMS_Admin</title>
<link href="Images/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript" src="../ACT_inc/js/lhgcore/lhgcore.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../ACT_inc/js/lhgcore/lhgdialog.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../ACT_inc/js/lhgcore/Main.js"></script>

 </head>
<body>
<% 
	Dim Rs,ShowErr
	Set Rs=server.CreateObject("adodb.recordset") 
	  If Request.QueryString("Action") = "Edit" Then	
		Dim AdminID,Admin_Name,Locked,Tel,Email,Description,Sex,RealName,Action
		AdminID= ChkNumeric(ACTCMS.S("AdminID"))
		'AdminID = Cint(Request.Cookies(AcTCMSN)("AdminID"))
		If AdminID =0 Or AdminID="" Then  Call Actcms.ACTCMSErr("")
		 If Not ACTCMS.ChkAdmin() Then '超级管理员检测
			If Cint(Request.Cookies(AcTCMSN)("AdminID"))<>AdminID Then Call Actcms.ACTCMSErr("")
		 End If 
		Rs.OPen "Select * from Admin_ACT Where Id = "&AdminID&" order by ID desc",Conn,1,1
			IF Not Rs.eof Then
				Admin_Name = Rs("Admin_Name")
				Locked = Rs("Locked")
				RealName = Rs("RealName")
				Tel = Rs("Tel")
				Email = Rs("Email")
				Description = Rs("Description")
				Sex = Rs("Sex")
			Else
				Admin_Name = ""
				Locked = 0
				Sex = 0
			End IF
			Rs.Close:Set Rs = Nothing
		End IF
		IF Request.QueryString("Action") = "Edit" Then Action = "Save" Else Action = "AddAdmin"
		Select Case Request.QueryString("Action")	
			   Case "AddAdmin" 
			   		Call Saveadmin()
					Response.End
			   Case "Save"
					Call SaveAdmin()
					Response.End
		End Select
			
		
		Sub Saveadmin()
		Dim Admin_Name ,PassWord,RPassWord,Locked,Tel,Email,Description,Sex,RealName,TempRs,AdminRS,AdminSql,AdminID
			Admin_Name = RSQL(Request.Form("Admin_Name"))
			PassWord = Request.Form("PassWord")
			RPassWord = Request.Form("RPassWord")	
			Sex = ChkNumeric(Request.Form("Sex"))
			Locked = ChkNumeric(Request.Form("Locked"))
			RealName = RSQL(Request.Form("RealName"))	
			Tel = RSQL(Request.Form("Tel"))	
			Email = RSQL(Request.Form("Email"))	
			Description = RSQL(Request.Form("Description"))	
			IF Trim(PassWord) <> Trim(RPassWord) Then
				ShowErr = "<li>2次输入的密码不一致! </li>"
				 Call Actcms.ActErr(ShowErr,"","")
				Response.end
			End IF
			
		IF Request.QueryString("Action") = "AddAdmin" Then
			 If Not ACTCMS.ChkAdmin() Then '超级管理员检测
				Call Actcms.ACTCMSErr("")
			 End If 
			IF Admin_Name <> "" Then
				IF Len(Admin_Name) >= 100 Then
					ShowErr = "<li>管理员名称不能超过50个字符</li>"
 					 Call Actcms.ActErr(ShowErr,"","")
					Response.end
				End IF
			Else
					ShowErr = "<li>请输入管理员名称!</li>"
					 Call Actcms.ActErr(ShowErr,"","")
					Response.end
			End if	
		
			Set TempRs = Conn.Execute("Select Admin_Name from Admin_ACT where Admin_Name='" & Admin_Name & "'")
			IF Not TempRs.Eof Then
					ShowErr = "<li>数据库中已存在该管理员名称!</li>"
					 Call Actcms.ActErr(ShowErr,"","")
					Response.end
			End IF			
			
			 Set AdminRS = Server.CreateObject("adodb.recordset")
				  AdminSql = "select * from Admin_ACT"
				  AdminRS.Open AdminSql, Conn, 1, 3
				  AdminRS.AddNew
				  AdminRS("AddDate") = Now
				  AdminRS("Admin_Name") = Admin_Name
				  AdminRS("PassWord")=MD5(RPassWord)
				  AdminRS("Locked") = Locked
				  AdminRS("RealName") = RealName
				  AdminRS("Sex") = Sex
				  AdminRS("Tel") = Tel
				  AdminRS("Email") = Email
				  AdminRS("Description") = Description
				  AdminRS("SuperTF") = 0
				  AdminRS("LoginTime") = Now
				  AdminRS("LoginIP") = GetIP()
				  AdminRS.Update
				  AdminRS.Close:Set AdminRS = Nothing			
 				  Call Actcms.ActErr("添加管理员成功","ACT.admin.asp","")
 				  Response.end
			ElseIF Request("Action") = "Save" Then
				  AdminID = ChkNumeric(Request("AdminID"))
				 If Not ACTCMS.ChkAdmin() Then '超级管理员检测
					If CStr(Request.Cookies(AcTCMSN)("AdminID"))<>CStr(AdminID) Then Call Actcms.ACTCMSErr("")
					If Not ACTCMS.ACTCMS_QXYZ(0,"editmypassword","") Then   Call Actcms.Alert("对不起，您不能修改自己的密码！","")
				 End If 
				  Set AdminRS = Server.CreateObject("adodb.recordset")
				  AdminSql = "select * from Admin_ACT Where ID="&AdminID
				  AdminRS.Open AdminSql, Conn, 1, 3
				  IF RPassWord <> "" Then  AdminRS("PassWord")=MD5(RPassWord)
				  AdminRS("Locked") = Locked
				  AdminRS("RealName") = RealName
				  AdminRS("Sex") = Sex
				  AdminRS("Tel") = Tel
				  AdminRS("Email") = Email
				  AdminRS("Description") = Description
				  AdminRS.Update
				Response.Cookies(AcTCMSN)("AdminName") = AdminRS("Admin_Name")'更新
				Response.Cookies(AcTCMSN)("AdminPassword") = AdminRS("PassWord")
				Response.Cookies(AcTCMSN)("AdminID") = AdminRS("ID")
				If AdminRS("SuperTF")=1 Then Response.Cookies(AcTCMSN)("SuperTF")=1
				Response.Cookies(AcTCMSN)("Purview") = AdminRS("Purview")
				Response.Cookies(AcTCMSN)("ACT_Other") = AdminRS("ACT_Other")
				Response.Cookies(AcTCMSN)("HQQXLX") = AdminRS("ACTCMS_QXLX")
			    AdminRS.Close:Set AdminRS = Nothing			
			    ShowErr = "<li>操作成功!&nbsp;&nbsp;<a href=ACT.admin.asp>点击这里返回管理首页</a></li>"
				
			     Call Actcms.ActErr(ShowErr,"ACT.admin.asp","")
			    Response.end
			End If
		
		End Sub
		
 %>
<form id="HFCMS" name="HFCMS" method="post" action="?Action=<%=Action %>&AdminID=<%=Request.QueryString("AdminID")%>">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
    <tr>
      <td colspan="2" class="bg_tr">您现在的位置：系统设置 &gt;&gt; 添加\修改管理员管理员 </td>
    </tr>
    <tr>
      <td width="20%" height="25">管理员帐号</td>
      <td width="80%" height="25"><input  class="Ainput"  <%IF Request.QueryString("Action") = "Edit" Then Response.Write "readonly disabled=true" %> name="Admin_Name" type="text" id="Admin_Name" value="<%= Admin_Name %>" size="40" /></td>
    </tr>

    <tr>
      <td height="25">初始密码</td>
      <td height="25"><input   class="Ainput"   type="password" size="42" name="PassWord" /> 
			<span class="h" style="cursor:help;"  onclick="dohelp('ACTAAdd_csmm')" id="ACTAAdd_csmm">帮助</span>
      <% if Request.QueryString("Action") = "Add" Then  Response.Write "<font color=""red"">密码不能少于6位</font>" Else Response.Write "<font color=""red"">密码不能少于6位,不修改请保持为空</font>" %>
	  </td>
    </tr>
     <tr>
      <td height="25">重复密码</td>
      <td height="25"><input  class="Ainput"  name="RPassWord" type="password" id="RPassWord" size="42" maxlength="40" />
      <span class="h" style="cursor:help;"  onclick="dohelp('ACTAAdd_cfmm')"  id="ACTAAdd_cfmm">帮助</span></td>
    </tr>
    <tr>
      <td height="25">是否锁定</td>
      <td height="25">
	  <label for="locked1"><input <%if Locked = 0 then response.Write("Checked")%>  type="radio" checked="checked" value="0" id="Locked1" name="Locked" />&nbsp;&nbsp;正常&nbsp;&nbsp;</label>
      <label for="locked2"><input <%if Locked = 1 then response.Write("Checked")%> type="radio" value="1" id="Locked2" name="Locked" />&nbsp;&nbsp;锁定&nbsp;&nbsp;</label>
<font color="red"> 锁定的用户不能登录后台管理</font>
<span class="h" style="cursor:help;"  onclick="dohelp('ACTAAdd_sd')"  id="ACTAAdd_sd">帮助</span>            </td>
    </tr>
    <tr>
      <td height="25">真实姓名</td>
      <td height="25"><input  class="Ainput"  name="RealName" type="text" id="RealName" value="<%= RealName %>" size="40" />
	  <span class="h" style="cursor:help;"  onclick="dohelp('ACTAAdd_zsxm')"  id="ACTAAdd_zsxm">帮助</span></td>
    </tr>
    <tr>
      <td height="25">性 别</td>
      <td height="25">
	  <label for="Sex1"><input <%if Sex = 0 then response.Write("Checked")%> type="radio" checked="checked" value="0" id="Sex1" name="Sex" />&nbsp;男 &nbsp;&nbsp;</label>
        <label for="Sex2"><input <%if Sex = 1 then response.Write("Checked")%> type="radio" value="1" id="Sex2" name="Sex" /> &nbsp; 女 &nbsp;&nbsp;</label>     </td>
    </tr>
    <tr>
      <td height="25">联系电话</td>
      <td height="25"><input  class="Ainput"  name="Tel" type="text" id="Tel" value="<%= Tel %>" size="40" /></td>
    </tr>
    <tr>
      <td height="25">电子信箱</td>
      <td height="25"><input  class="Ainput"  name="Email" type="text" id="Email" value="<%= Email %>" size="40" /></td>
    </tr>
       <tr>
      <td height="25">简要说明</td>
      <td height="25"><textarea name="Description" cols="50%" rows="6" id="Description"><%= Description %></textarea></td>
    </tr>
    <tr>
      <td height="25">&nbsp;</td>
      <td height="25"><input type=button onclick=CheckForm() class="ACT_btn" name=Submit value="  保存  " />
	  								
      <input type="reset" class="ACT_btn" name="Submit2" value="  重置  " /></td>
    </tr>
  </table>
</form>
<SCRIPT language=javascript>
<!--
function CheckForm()
{ var form=document.HFCMS;
   if (form.Admin_Name.value=='')
    { alert("请输入管理员名称!");   
	  form.Admin_Name.focus();    
	   return false;
    }
  <% if Request.QueryString("Action") = "Add" then %>
    if (form.PassWord.value=='')  
	  {     alert("请输入初始密码!");   
	    form.PassWord.focus();   
		  return false;   
		 }   else if (form.PassWord.value.length<6)  
		   {  alert("初始密码不能少于6位!");   
		     form.PassWord.focus();   
			   return false;   
	}   if (form.RPassWord.value=='')  
	  {     alert("请输入确定密码!");   
	    form.RPassWord.focus();    
		 return false;    
		 }  
		  else if(form.RPassWord.value.length<6) 
		   {    
		    alert("确定密码不能少于6位!");    
			 form.RPassWord.focus(); 
			 return false;    } 
		<%end if%>  
	 if (form.PassWord.value!=form.RPassWord.value)   
	   {    
	    alert("两次输入的密码不一致!");   
		  form.PassWord.focus();   
		  return false;    
		  }
		   if (form.RealName.value=='')
			{
			 alert("请输入真实姓名");
			 form.RealName.focus();
			 return false;
			}
   if (form.Email.value!='')
   if(check(form.Email.value)==false)
      { alert('非法电子邮箱!');
        form.Email.focus();
        return false;
     }  
		form.Submit.value="正在提交数据,请稍等...";
		form.Submit.disabled=true;	
		form.Submit2.disabled=true;	
	    form.submit();
        return true;
}

function check(str)
{ if((str.indexOf("@")==-1)||(str.indexOf(".")==-1)){
	
	return false;
	}
	return true;
}
//-->
</SCRIPT>
</body>
</html>
