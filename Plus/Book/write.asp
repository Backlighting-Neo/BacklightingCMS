<!--#include file="../../ACT_inc/ACT.User.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<% 
		Server.ScriptTimeOut=999
		Response.Buffer = True
		Response.Expires = -1
		Response.ExpiresAbsolute = Now() - 1
		Response.CacheControl = "no-cache"
		response.Charset = "utf-8"
		ConnectionDatabase
		Dim Rs1,ActCMS_BookSetting,Content,Book_Code,url
		Set Rs1=Conn.Execute("Select PlusConfig,IsUse from Plus_ACT where PlusID='lyxt_ACT'")
		If Rs1("IsUse")=1 Then Call actcms.alert("该系统已经被管理员关闭","")
		ActCMS_BookSetting=Split(Rs1(0),"^@$@^")
		Dim server_v1,server_v2,rs,nr,sh,sms
		server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
		server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
		Content=ACTCMS.HTMLcode(ACTCMS.s("nr"))
		if ActCMS_BookSetting(0)=1 then call ACTCMS.Alert("请不要从外部提交数据.留言功能已经被关闭",""):response.end
		if ActCMS_BookSetting(2)=0 then sh=1 else sh=0
		if  ActCMS_BookSetting(3) = 0 Then
		 Book_Code = Request.Form("Code")
			If Book_Code <>CStr(Session("GetCode")) then
				Call ACTCMS.Alert("验证码有误，重新输入","")		
				Response.End
			End if
		End  If
		If Len(Content)>CLng(ActCMS_BookSetting(5)) and ActCMS_BookSetting(5)<>0 Then
		 Call  ACTCMS.alert("留言内容内容必须在" &ActCMS_BookSetting(5) & "个字符以内!","")
		 Response.End
		End if
		if  mid(server_v1,8,len(server_v2))<>server_v2  then
			call ACTCMS.alert("警告！你正在从外部提交数据！！请立即终止！！","")
			response.end
		end if
		if ACTCMS.s("name") ="" then
			call ACTCMS.alert("请输入您的姓名,按确定返回上一页","")
			response.end
		End  if
		if len(request.Form("name"))>12 then
			call ACTCMS.alert("请输入您的正确姓名,并且不能超过6个字,按确定返回上一页","")
			response.end
		end if
		if len(ACTCMS.s("qq"))>12 and isnumeric(ACTCMS.s("qq")) then
			call ACTCMS.alert("请输入您的正确QQ,按确定返回上一页","")
			response.end
		End  if
		if Content ="" then
			call ACTCMS.alert("请输入内容,按确定返回上一页","")
			response.end
		End  if
		If Trim(url)="http://" Then 
			url=""
		Else
			url= server.HTMLEncode(Request.Form("url"))
		End If 
		set rs=server.CreateObject("adodb.recordset")
		rs.open "Book_ACT",conn,3,3
		rs.addnew
		rs("show")=server.HTMLEncode(Request.Form("show"))
		rs("qq")=server.HTMLEncode(Request.Form("qq"))
		rs("name")=server.HTMLEncode(Request.Form("name"))
		rs("mail")=server.HTMLEncode(Request.Form("mail"))
		rs("url")=url
		rs("xq")=ChkNumeric(Request.Form("xq"))
		rs("sh")=sh
		rs("nr")=replace(server.HTMLEncode(Content),chr(13)&chr(10),"<br>")
		rs("ip")=Request.serverVariables("REMOTE_ADDR")
		rs("addtime")= now
		rs.Update 
		rs.close:set rs =nothing  
		If  sh=1 then sms="您的留言需要管理员审核才可以看到"
		Call  ACTCMS.alert("留言成功,"&sms&",谢谢支持,按确定返回首页","index.asp")
		response.End 

 %>
