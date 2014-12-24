<!--#include file="../../act_inc/ACT.User.asp"-->
<%Dim CommUser
ConnectionDatabase

	Dim ModeID,ClassID,ID,Action,rs,CommentStr,names
	ModeID=ChkNumeric(request("ModeID"))
	ClassID=RSQL(request("ClassID"))
	ID=ChkNumeric(request("ID"))
	Action=request("Action")
	Set CommUser = New ACT_User
	Action=request("Action")
	Select Case Action
	  Case "WriteSave"
	  Call WriteSave()
	 Case "Support"
	  Call Support()
	 Case "Write"
	  Call Write
	 Case Else
		response.write "error":response.end
	 End Select


  
	Function Write()
   	If Actcms.ACT_C(ModeID,15)="0"  Then 
		response.write ("document.write(""该评论已关闭"");")&vbcrlf&vbcrlf
		response.end
	End If 

	Set rs=actcms.actexe("select rev from  "&ACTCMS.ACT_C(ModeID,2)&"  where id="&ID)
	If rs.eof Then response.write ("document.write(""已经关闭评论"");")&vbcrlf&vbcrlf:response.end
	If rs("rev")="0" Then  response.write ("document.write(""已经关闭评论"");")&vbcrlf&vbcrlf:response.End
	CommentStr=CommentStr &("document.write(""<form action='"&actcms.actsys&"plus/Comment/ACT.Comment.asp?Action=WriteSave' method='post' name='Comment'>"");")&vbcrlf&vbcrlf
    CommentStr = CommentStr &("document.write(""<input type='hidden' value='" & ModeID & "' name='ModeID'><input type='hidden' value='" & ClassID & "' name='ClassID'><input type='hidden' value='" & ID & "' name='ID'>"");")&vbcrlf&vbcrlf
	CommentStr=CommentStr &("document.write(""<div class='act_fbpl_nr'><small>请自觉遵守互联网相关的政策法规，严禁发布色情、暴力、反动的言论。</small></div> "");")&vbcrlf&vbcrlf
 	CommentStr=CommentStr &("document.write(""<textarea cols='60' name='Content' rows='5' class='act_fbpl_k'></textarea> <div class='act_fbpl_info'>"");")&vbcrlf&vbcrlf
 	IF Cbool(CommUser.UserLoginChecked)=false then
		CommentStr=CommentStr &("document.write(""用户名:<input type='text' name='username' size='16' class='act_fbpl_kk' />E-Mail:<input type='text' name='Email' size='16' class='act_fbpl_kk' />"");")&vbcrlf&vbcrlf
 	Else
		CommentStr=CommentStr &("document.write(""用户名:"&CommUser.username&""");")&vbcrlf&vbcrlf
  	End If	
	
	If Actcms.ACT_C(ModeID,13)=0 Then
		CommentStr=CommentStr &("document.write(""&nbsp;验证码：<input type='text' class='act_fbpl_kk' size='4' name='Code' />"");")&vbcrlf&vbcrlf
		CommentStr = CommentStr &("document.write(""<img style='cursor:hand;' src='"&ACTCMS.ActSys&"ACT_INC/Code.asp?s=+Math.random();' id='IMG1' onclick=this.src='"&ACTCMS.ActSys&"ACT_INC/Code.asp?s=+Math.random();' alt='看不清楚? 换一张！'>"");")&vbcrlf&vbcrlf
	End If
 	CommentStr=CommentStr &("document.write(""<input type='checkbox' name='AnounName'  id='AnounNames'/><label for='AnounNames'> 匿名</label>"");")&vbcrlf&vbcrlf
  	CommentStr=CommentStr &("document.write(""<input  type='button' class='dcmp-submit' onclick=CheckForm() name=Submit value='提交发表'/>"");")&vbcrlf&vbcrlf
	CommentStr=CommentStr &("document.write(""</div></form>"");")&vbcrlf&vbcrlf
	 response.write CommentStr
	End Function 

	Sub WriteSave()
 		Dim UserName,Email,Content,Locked,point,Code,AnounName
 		If Actcms.ACT_C(ModeID,15)="0"  Then 
			response.write ("document.write(""已经关闭评论"");")&vbcrlf&vbcrlf
			response.end
		End If 
		Set rs=actcms.actexe("select rev from  "&ACTCMS.ACT_C(ModeID,2)&"  where id="&ID)
		If rs.eof Then response.write ("document.write(""已经关闭评论"");")&vbcrlf&vbcrlf:response.End
		If rs("rev")="0" Then  response.write ("document.write(""已经关闭评论"");")&vbcrlf&vbcrlf:response.End
		UserName=Server.HTMLEncode(actcms.s("UserName"))
 		AnounName=ChkNumeric(request("AnounName"))
		Email=Server.HTMLEncode(actcms.s("Email"))
		Content=request("Content")
		Code=request("Code")
 		IF AnounName="" Or Not IsNumeric(AnounName) Then AnounName=1
		IF Cint(CommUser.UserLoginChecked)=false Then AnounName=1
			Select Case Actcms.ACT_C(ModeID,15)
				   Case 0
						If Cint(CommUser.UserLoginChecked)=True And  CommUser.G_WriteComment="1" Then 
						Else 
							Response.Write("<script>alert('本栏目禁止发表评论!');history.back();</script>")
							Response.End
						End If 
						Locked=0
						If Cint(CommUser.UserLoginChecked)=True And  CommUser.G_SH="1" Then Locked=1
				   Case 1
						If Cint(CommUser.UserLoginChecked)=True And  CommUser.G_WriteComment="1" Then 
							Locked=0
						Else 
							IF Cint(CommUser.UserLoginChecked)=False Then
								 Response.Write("<script>alert('本栏目只允许会员发表评论,请注册后再发表');history.back();</script>")
								 Response.End
							Else
								Locked=0
							End If 
						End If 
						If Cint(CommUser.UserLoginChecked)=True And  CommUser.G_SH="1" Then Locked=1
				   Case 2
						If Cint(CommUser.UserLoginChecked)=True And  CommUser.G_WriteComment="1" Then 
							Locked=1
						Else 
							IF Cint(CommUser.UserLoginChecked)=False Then
								 Response.Write("<script>alert('本栏目只允许会员发表评论,请注册后再发表');history.back();</script>")
								 Response.End
							Else
								Locked=1
							End If 
						End If 
					Case 3
						Locked=0
						If Cint(CommUser.UserLoginChecked)=True And  CommUser.G_SH="1" Then Locked=1
					Case 4
						Locked=1
					Case Else 
						Exit Sub 
			End Select 
 
		If Actcms.ACT_C(ModeID,13)=0 and Trim(Request.Form("Code"))<>Trim(Session("GetCode")) Then
		 Response.Write("<script>alert('验证码有误，请重新输入！!');history.back();</script>")
		 Response.End
		End IF
		IF ID="" Then 
		 Response.Write("<script>alert('参数传递有误!');history.back();</script>")
		 Response.End
		End if
	 
		if Content="" Then 
		 Response.Write("<script>alert('请填写评论内容!');history.back();</script>")
		 Response.End
		End If
		If  AnounName=0 And  Cint(CommUser.Locked)=1 Then 
		 Response.Write("<script>alert('您的账号已被管理员锁定，请与管理员联系或选择游客发表!');history.back();</script>")
		 Response.End
		End If
		If Actcms.ACT_C(ModeID,14)<>0 Then 
			if Len(Content)>Actcms.ACT_C(ModeID,14) Then 
			 Response.Write("<script>alert('评论内容必须在" &Actcms.ACT_C(ModeID,14) & "个字符以内!');history.back();</script>")
			 Response.End
			End if

		End If 
 		Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select * From Comment_Act",Conn,1,3
		RS.AddNew

		If AnounName=1 Then 
 			 If UserName="" Then RS("userid")="0"
 			 RS("Email")=Email
		Else
			 RS("UserID")=CommUser.UserID
  		End If 
		 RS("ModeID")=ModeID
		 RS("ClassID")=ClassID
		 RS("acticleID")=ID
		 RS("Content")=ACTCMS.HTMLEncode(Content)
		 RS("UserIP")=GetIP
		 RS("Locked")=Locked
		 RS("AddDate")=Now
		RS.UpDate
		actcms.ACTEXE("Update "&ACTCMS.ACT_C(ModeID,2)&"  Set commentscount=commentscount+1 Where ID=" & ID & "")
		RS.Close:Set RS=Nothing
		If Locked=1 then
		 Response.Write "<script>alert('你的评论发表成功!');location.href='" & Request.ServerVariables("HTTP_REFERER") & "';</script>"
		Else
		 Response.Write "<script>alert('你的评论发表成功,需要管理员后台审核才能显示!');location.href='" & Request.ServerVariables("HTTP_REFERER") & "';</script>"
		End if

	End Sub 




	Sub Support()
	Dim ID,OpType,RS,rs1
	ID=ChkNumeric(request("id"))
	OpType=request("Type")
	IF Cbool(Request.Cookies(Cstr(ID))("SupportCommentID"))<>true Then
		Set RS=Server.CreateObject("Adodb.Recordset")
		RS.Open "Select * From Comment_Act Where ID=" & ID ,Conn,1,3
		 if not rs.eof then
			  if OpType=1 Then
			 RS("Y")=RS("Y")+1
			  else
			 RS("N")=RS("N")+1
			  end if
		 RS.UpDate
		 RS.Close:Set RS=Nothing
	   end if
		 Response.Cookies(Cstr(ID))("SupportCommentID")=True
 		 Set rs1=actcms.actexe("Select top 1 id,y,n From Comment_Act Where ID="&id)
		 If Not rs1.eof Then 
			if OpType=1 Then
 				response.write "ok|"&rs1("Y")
			 Else 
				 response.write "ok|"&rs1("N")
			End If 
		End If 
  	Else
 		 response.write "err|0"
	End If
 	End Sub











%>