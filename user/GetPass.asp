<!--#include file="../act_inc/ACT.User.asp"-->
<!--#include file="../act_inc/MD5.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>用户管理中心</title>
<link href="Images/css/css.css" rel="stylesheet" type="text/css">

<%		 ConnectionDatabase
		 Dim Step:Step=ACTCMS.S("Step")
		  IF Step="" Then Step=1
		  IF Step=2 Then
		     Dim RS,TableName,RsUser,ModeID
			 Dim UserName:UserName=RSQL(ACTCMS.S("UserName"))
			 
             Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select Question From  User_ACT Where UserName='" & UserName & "'",Conn,1,1
			  IF RS.Eof And RS.Bof Then
			  	 Response.Write("<script>alert('对不起,您输入的用户名不存在！');history.back();</script>")
				 Response.End
			  Else
		     %>
			 	<script language = "JavaScript">
				function CheckForm()
				{
				if (document.myform.Answer.value=="")
				  {
					alert("请输入问题答案！");
					document.myform.Answer.focus();
					return false;
				  }
				if (document.myform.Code.value=="")
				  {
					alert("请输入验证码！");
					document.myform.Code.focus();
					return false;
				  }
	              return true;
				  }
				</script>
					  <table width="96%"  border="0" align="center" cellpadding="0" cellspacing="1" class="table">
					 	<form name="myform" method="post" action="?Step=3" onSubmit="return CheckForm();">
                        <input type="hidden" value="<%=UserName%>" name="UserName">
						<input name="ModeID" type="hidden" id="ModeID" value="<%=ModeID%>">
                        <tr class="Title">
                            <td height="24" colspan=2 align="center" class="bg_tr">取回密码第二步 回答密码问题 </td>
                        </tr>
                            <tr class="tdbg">
                              <td width="40%" height="30" align="right" class="td_bg"> 密码问题：</td>
                              <td width="60%" class="td_bg"><%=RS(0)%></td>
                            </tr>
                            <tr class="tdbg">
                              <td width="40%" height="30" align="right" class="td_bg"> 您的答案：</td>
                              <td width="60%" class="td_bg"><input name="Answer" type="text" id="Answer" size="20" /></td>
                            </tr>
                            <tr class="tdbg">
                              <td width="40%" height="30" align="right" class="td_bg"> 验证码：</td>
                              <td width="60%" class="td_bg"><input name="Code" type="text" id="Code" size="6" />
							<img src="../act_inc/code.asp?s='+Math.random();" alt="验证码" title="看不清楚? 换一张！" style="cursor:hand;" onclick="src='../act_inc/code.asp?s='+Math.random()"/>  
							  </td>
                            </tr>
                            <tr class="tdbg">
                              <td height="42" colspan=2 align="center" class="td_bg">
							  <input class="Button" name="Submit2" type="submit" value="下一步" />
                              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
                            </tr>
                  </table>
						</form>
                    
		  <%   End IF
		  ElseIF Step=3 Then

             Dim Code:Code=	RSQL(ACTCMS.S("Code"))
			 UserName=RSQL(ACTCMS.S("UserName"))

			 Dim Answer:Answer=RSQL(ACTCMS.S("Answer"))
			IF Trim(Code)<>Cstr(Session("GetCode")) then 
		   	 Response.Write("<script>alert('验证码有误，请重新输入！');history.back();</script>")
		     Response.End
			End If


		 
			Dim RSC
            Set RSC=ACTCMS.ACTEXE("Select Answer From User_ACT  Where UserName='" & UserName & "' and Answer='" & Answer & "'")
			IF RSC.EOF AND RSC.Bof Then
			  	 Response.Write("<script>alert('对不起,您输入的答案不正确！');history.back();</script>")
				 Response.End
			Else
			 %>
			 
			 <script language = "JavaScript">
				function CheckForm()
				{
				if (document.myform.PassWord.value=="")
				  {
					alert("请输入新密码！");
					document.myform.PassWord.focus();
					return false;
				  }
				if (document.myform.RePassWord.value=="")
				  {
					alert("请输入确认密码！");
					document.myform.RePassWord.focus();
					return false;
				  }
				if (document.myform.PassWord.value!=document.myform.RePassWord.value)
				  {
					alert("两次输入的密码不一致！");
					document.myform.PassWord.focus();
					return false;
				  }
	              return true;
				  }
				</script>
                       <table width="96%"  border="0" align="center" cellpadding="0" cellspacing="1" class="table">
							<tr class="Title">
							<td height="24" colspan="2" align="center" class="bg_tr">取回密码第三步 设置新密码 </td>
							</tr>
						 <form name="myform" method="post" action="?Step=4" onSubmit="return CheckForm();">
                         <input type="hidden" value="<%=request("Answer")%>"  id="Answer"  name="Answer">
								<tr class="tdbg">
  
								  <td width="40%" height="30" align="right"> 用户名：</td>
                                  <td width="60%"><input type="text" readonly value="<%=UserName%>" name="UserName"></td>
                                </tr>
                                <tr class="tdbg">
                                  <td width="40%" height="30" align="right"> 新密码：</td>
                                  <td width="60%"><input name="PassWord" type="password" id="PassWord" size="20" /></td>
                                </tr>
                                <tr class="tdbg">
                                  <td width="40%" height="30" align="right"> 确认密码：</td>
                                  <td width="60%"><input name="RePassWord" type="password" id="RePassWord" size="20" /></td>
                                </tr>
                                <tr class="tdbg">
                                  <td height="42" align="center" colspan=2><input  class="Button" name="Submit22" type="submit" value=" 完 成 " />
                                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
                                </tr>
                              </tbody>
                  </table>
						  </form>
					
			<% End IF
		  ElseIF Step=4 Then
			 UserName=RSQL(ACTCMS.S("UserName"))
		  	 Dim PassWord:PassWord=RSQL(ACTCMS.S("PassWord"))
			 Dim RePassWord:RePassWord=ACTCMS.S("RePassWord")
		  	 answer=RSQL(Request("Answer"))
			 If PassWord = "" Then
				 Response.Write("<script>alert('请输入登录密码!');history.back();</script>")
				 Response.End
			 ElseIF RePassWord="" Then
				 Response.Write("<script>alert('请输入确认密码');history.back();</script>")
				 Response.End
			 ElseIF PassWord<>RePassWord Then
				 Response.Write("<script>alert('两次输入的密码不一致');history.back();</script>")
				 Response.End
			 End If
			
		 
			 
			 Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select PassWord From user_act Where UserName='" & UserName & "'  and answer='" & answer &"'",Conn,1,3
			  If Not rs.eof Then 
				 RS(0)=MD5(PassWord)
				 RS.Update
			  Else
				response.write "非法提交"
				response.end
			  End If 
			 RS.Close
			 Set RS=Nothing
		  %>
                  <table width="96%"  border="0" align="center" cellpadding="0" cellspacing="1" class="table">
                          <tr class="Title">
                              <td height="25" align="center" valign="bottom" class="bg_tr">取回密码成功</td>
                          </tr>
                           <tr class="tdbg">
                                  <td height="50" align="center">恭喜你,密码取回成功!您的新密码是:<font color=red><%=PassWord%></font>,请用新密码登录。</td>
                    </tr>
                  </table>
                       
		  <%
           Else
		   %>
		   <script language = "JavaScript">
				function CheckForm()
				{
				if (document.myform.UserName.value=="")
				  {
					alert("请输入用户名！");
					document.myform.UserName.focus();
					return false;
				  }
	              return true;
				  }
				</script>

			 <form name="myform" method="post" action="?Step=2" onSubmit="return CheckForm();">
                 <table width="96%"  border="0" align="center" cellpadding="0" cellspacing="1" class="table">
					  <tr class="Title">
							<td height="24" colspan="2" align="center" class="bg_tr">取回密码第一步 输入用户名 </td>
					  </tr>
						  <TR class="tdbg">
							<TD width="40%" height=25 align="right"> 您的用户名：</TD>
							<TD width="60%">
							<input name="UserName" type="text" id="UserName" size="20">
						
						</TD>
						  </TR>
						  <TR class="tdbg">
							<TD  colspan="2" height=42 align="center"> 
							<input  class="ACT_btn" name="Submit" type="submit" value="下一步">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </TD>
						  </TR>
						</TBODY>
			   </TABLE>
				</form>
		  	 <%End IF%> 			  

