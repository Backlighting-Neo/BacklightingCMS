<!--#include file="../act_inc/ACT.User.asp"-->
 <!--#include file="../ACT_inc/md5.ASP"-->
  <!--#include file="../ACT_inc/ACT.U_M.ASP"-->
<!--#include file="../Field.asp"-->
<% 
    dim UserHS ,A
 	Set UserHS = New ACT_User
	IF Cbool(UserHS.UserLoginChecked)=false then
	  Response.Write "<script>top.location.href ='login.asp' ;</script>"
	  Response.end
	End If	
 %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>会员中心</title>
 <link href="images/css.css" rel="stylesheet" type="text/css" />
<script charset="utf-8"  language="JavaScript" type="text/javascript" src="../editor/kindeditor/kindeditor.js" ></script>
<script language="JavaScript" type="text/javascript" src="../ACT_inc/js/lhgcore/lhgcore.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../ACT_inc/js/lhgcore/lhgdialog.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../ACT_inc/js/lhgcore/Main.js"></script>
<script type='text/javascript' src='../ACT_INC/js/time/WdatePicker.js'></script>
<script type='text/javascript' src='../ACT_INC/main.js'></script>
<script type="text/javascript" src="../ACT_INC/js/swfobject.js"></script>
 <script type="text/javascript">
var U="<%=ACTCMS.strToAsc(RSQL(Trim(Request.Cookies(AcTCMSN)("UserName"))))%>";
var P="<%=ACTCMS.strToAsc(RSQL(Trim(Request.Cookies(AcTCMSN)("PassWord"))))%>";
</script>
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
	</div>
    <table cellpadding="0" cellspacing="0" class="member_info">
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
<%  A=request("A")
 	Select Case A
		Case "InfoSave"
			Call InfoSave()
		Case Else
			Call Main()
 	End Select 
    sub InfoSave()
			Dim IF_NULL,i
			IF_NULL=ACTCMS.Act_MX_Arr(Userhs.UModeID,2)
			If IsArray(IF_NULL) Then
			For I=0 To Ubound(IF_NULL,2)
			 If IF_NULL(2,I)=0 And Trim(ACTCMS.S(IF_NULL(0,I)))="" Then  Call  ACTCMS.ALERT(IF_NULL(1,I)&"不能为空","")
			Next
			End If
	
	
			 Dim Email:Email=ACTCMS.S("Email")
			 if ACTCMS.IsValidEmail(Email)=false then
				 Response.Write("<script>alert('请输入正确的电子邮箱!');history.back();</script>")
				 Exit Sub
			 end if
		
			 Dim address:address=ACTCMS.S("address")
			 Dim HomeTel:HomeTel=ACTCMS.S("HomeTel")
			 Dim Mobile:Mobile=ACTCMS.S("Mobile")
			 Dim Sex:Sex=ChkNumeric(ACTCMS.S("Sex"))
			 Dim province:province=ACTCMS.S("prov")
			 Dim city:city=ACTCMS.S("city")
 			 Dim Realname:Realname=ACTCMS.S("Realname")
			 Dim QQ:QQ=ACTCMS.S("QQ")		 
			 Dim MSN:MSN=ACTCMS.S("MSN")		 
             Dim Birthday:Birthday=ACTCMS.S("Birthday")		 
	         Dim postcode:postcode=ACTCMS.S("postcode")		 
	         Dim myface:myface=ACTCMS.S("myface")		 
 			if IsDate(Birthday)=false then Birthday="1900-1-1"
		 
		  Dim RS:Set RS=Server.CreateObject("Adodb.RecordSet")
		  RS.Open "Select * From User_ACT Where UserID=" & UserHS.UserID & "",Conn,1,3
		  IF RS.Eof And RS.Bof Then
			 Response.End
		  Else
		     RS("postcode")=postcode
 			 RS("Birthday")=Birthday
			 RS("Realname")=Realname
			 RS("QQ")=QQ
			 RS("MSN")=MSN
			 RS("address")=address
			 RS("HomeTel")=HomeTel
			 RS("Mobile")=Mobile
			 RS("Sex")=Sex
			 RS("Email")=Email
			 RS("Province")=Province
			 RS("City")=City
			 RS("myface")=myface
			 RS.Update



		  Set RS=Server.CreateObject("Adodb.RecordSet")
		  RS.Open "Select * From "&ACTCMS.ACT_U(UserHS.UModeID,2)&" Where UserID=" & UserHS.UserID & "",Conn,1,3
			 If IsArray(IF_NULL) Then
 				For I=0 To Ubound(IF_NULL,2)
 					If IF_NULL(3,I)="NumberType" Then 
						   If actcms.regexField(ACTCMS.S(IF_NULL(0,I)),"^\d+$")=True Then 
							   rs("" & IF_NULL(0,I) & "" )= ACTCMS.S(IF_NULL(0,I))
						   End If 
					ElseIf IF_NULL(3,I)="DateType" Then 
						If IsDate(ACTCMS.S(IF_NULL(0,I)))=False Then 
							rs("" & IF_NULL(0,I) & "")= Now()
						Else 
							rs("" & IF_NULL(0,I) & "")=ACTCMS.S(IF_NULL(0,I))
						End If
					ElseIf IF_NULL(4,I)="1" Then 
 							 rs("" & IF_NULL(0,I) & "")= actcms.AField(IF_NULL(5,I))
					ElseIf IF_NULL(4,I)="2" Then 
							If actcms.regexField(ACTCMS.S(IF_NULL(0,I)),IF_NULL(5,I))=True Then 
								rs("" & IF_NULL(0,I) & "")=ACTCMS.S(IF_NULL(0,I))
							Else 
								Call Actcms.Alert(IF_NULL(6,I),"")
 							End If 
  					Else 
						rs("" & IF_NULL(0,I) & "")=ACTCMS.S(IF_NULL(0,I))
					End If 
					actField=""
 				Next
			 End If			
		    rs.Update











			 Response.Write "<script>alert('您的联系信息修改成功！');location.href='edit.asp';</script>"
			 Response.End()
		  End if
		RS.Close:Set RS=Nothing


  end sub 
 sub main() %>
<p id="position"> <strong>当前位置：</strong><a href="index.asp">会员中心</a>修改资料</p>
<form name="actcmsfrom" action="?A=InfoSave"   method="post">
  <table cellpadding="0" cellspacing="1" class="table_form">
    <caption>修改资料</caption>
    <tr>
      <th width="20%"><strong>用户名：</strong><br /></th>
      <td width="80%"><strong><%=UserHS.username%></strong></td>
    <tr>
      <th><strong>E-mail：</strong><br /></th>
      <td><input name="Email" type="text" id="Email" size="30"  value="<%=UserHS.Email%>"> 
      </td>
    </tr>
    <tr>
      <th><strong>用户头像：</strong><br /></th>
      <td><input name="myface" type="text" value="<%
If UserHS.myface<>"" Then Response.Write UserHS.myface


%>" size="50" maxlength="100" />  
 	  <input name="button"    onClick="javascript:uploadimg('myface','<%=UserHS.UModeID%>');"  id="myfaces"   type="button"  class="button_style"  value="点击上传图片">  
	  
	  
	  
	  </td>
    </tr>
<tr>
  <th><strong>所在地区：</strong><br /></th>
  <td>
 
                              <select name="prov" onChange="selectcityarea('prov','city','actcmsfrom');" style="width:110">
                                <%if UserHS.Province<>"" then%>
                                <option  value="<%=UserHS.Province%>" selected="selected"><%=UserHS.Province%></option>
                                <%else%>
                                <option  value="" selected="selected">==请选择省份==</option>
                                <%end if%>
                              </select>
                              <select name="city" style="width:110">
                                <%if UserHS.city<>"" then%>
                                <option  value="<%=UserHS.city%>" selected="selected"><%=UserHS.city%></option>
                                <%else%>
                                <option  value="" selected="selected">==请选择城市==</option>
                                <%end if%>
                              </select>
                              选择您所在的省份和城市。
                              <script language="JavaScript" src="../act_inc/Province.js" type="text/javascript"></script>
 
          </td>
</tr>
    <tr>
    	<th><strong>姓名：</strong><br /></th><td><input name="Realname" type="text" id="Realname" size="30"  value="<%=UserHS.Realname%>">  </td>
    </tr>
        <tr>
    	<th><strong>性别：</strong><br /></th><td>
          <select name="Sex" id="Sex" style="width:110">
          <option value="">==请选择性别==</option>
          <option value="0" <%if UserHS.sex="0" then%> selected="selected" <%else%><%end if%>>男</option>
          <option value="1" <%if UserHS.sex="1" then%> selected="selected" <%else%><%end if%>>女</option>
          <option value="2" <%if UserHS.sex="2" then%> selected="selected" <%else%><%end if%>>保密</option>
         </select>
         </td>
    </tr>
        <tr>
    	<th><strong>出生日期：</strong><br /></th>
        <td>
        <input name="Birthday" type="text" id="Birthday" size="30"   onClick="WdatePicker()" value="<%=UserHS.Birthday%>">
        </td>
    </tr>
        <tr>
    	<th><strong>手机：</strong><br /></th><td><input name="Mobile" type="text" id="Mobile" size="30"  value="<%=UserHS.Mobile%>">
         </td>
    </tr>
        <tr>
    	<th><strong>电话：</strong><br /></th><td><input name="HomeTel" type="text" id="HomeTel" size="30"  value="<%=UserHS.HomeTel%>">   </td>
    </tr>
        <tr>
    	<th><strong>QQ：</strong><br /></th><td><input name="QQ" type="text" id="QQ" size="30"  value="<%=UserHS.QQ%>">  </td>
    </tr>
        <tr>
    	<th><strong>MSN：</strong><br /></th><td><input name="MSN" type="text" id="MSN" size="30"  value="<%=UserHS.MSN%>">  </td>
    </tr>
        <tr>
    	<th><strong>地址：</strong><br /></th><td><input name="address" type="text" id="address" size="30"  value="<%=UserHS.address%>"> </td>
    </tr>
    
    
    <tr>
    	<th><strong>邮编：</strong><br /></th><td><input name="postcode" type="text" id="postcode" size="30"  value="<%=UserHS.postcode%>">  </td>
    </tr>
    
    <%= U_M.ACT_MXEdit(UserHS.UModeID,UserHS.UserID) %>
    
    
        <tr>
      <th></th>
      <td><label>
        <input type="submit" name="dosubmit" id="button" class="button_style"  value="确 定" />
　
        <input type="reset" name="button2" id="button2" class="button_style"  value="重 置" />
        </label></td>
    </tr>
  </table>
</form>  
<% end sub %>
<span id="toggle_pannel" style="display:none;"></span>
<div class="clear"></div>
</div>
  </div>
<!--#include file="foot.asp"-->
</body>
</html> 