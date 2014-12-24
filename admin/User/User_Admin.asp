<!--#include file="../ACT.Function.asp"-->
<!--#include file="../../act_inc/cls_pageview.asp"-->
  <html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>act_cms_会员管理</title>
<link href="../Images/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript" src="../../editor/ckeditor/ckeditor.js"></script>
<script type='text/javascript' src='../../ACT_INC/js/time/WdatePicker.js'></script>
<style type="text/css">
<!--
.STYLE1 {font-weight: bold}
-->
</style>
</head>
<body>
<%
	If Not ACTCMS.ChkAdmin() Then  Call Actcms.Alert("对不起，您没有操作权限！","")

	Dim ShowErr,Action,UserID,Sql,ValidDays,tmpDays,EdaysFlag,ModeID,TableName,IF_NULL,i
	Action = Request("Action")
	UserID = Request("UserID")
 ModeID = ChkNumeric(Request("ModeID"))
 if ModeID=0 or ModeID="" Then ModeID=1
 
%>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td colspan="2" align="center"  class="bg_tr"><strong>会员系统----用户管理</strong></td>
  </tr>
  <tr>
    <td width="9%" align="right"><strong>用户选项：</strong></td>
    <td><span class="STYLE1"><A href="?Action=UserSearch">搜索用户</A>┆<A href="?ModeID=<%=ModeID%>">所有用户</A>┆<A href="?UserSearch=1">被锁住的用户</A>┆<A href="?UserSearch=2">待审批会员</A>┆<A href="?UserSearch=3">待邮件验证</A>┆<A href="?UserSearch=4">24小时内登录</A>┆<A href="?UserSearch=5">24小时内注册</A></span></td>
  </tr>
</table><% 
	 TableName=actcms.ACT_U(ModeID,2) 
  	Select Case Action
 			Case "Del"
				Call Del()
			Case "Lock"
				Call Locked()
			Case "UnLock"
				Call UnLocked()
			Case "Move"
				Call MoveUser()
			Case "UserSearch"
				Call UserSearch()
		    Case "AddZJ"
				Call AddZJ()
		    Case "SaveAddZJ"
				Call SaveAddZJ()
			Case "AddMoney"
				Call AddMoney()
			Case "SaveAddMoney"
				Call SaveAddMoney()
 			Case Else
				Call Main()
	End Select
		
	
 	

		Sub SaveAddMoney()
			dim UserID,ChargeType,Point,Edays,rsUser,sqlUser,Reason,ErrMsg
			Dim Money:Money=Request("Money")
			If Not IsNumeric(Money) Then 
 				Call Actcms.ActErr("减去的资金有误","","1")
				exit sub
			End if
			Action=Trim(request("Action"))
			UserID=ChkNumeric(request("UserID"))
			if UserID=0 then
 				Call Actcms.ActErr("参数不足","","1")
				exit sub
			end if
			ChargeType=Trim(request("ChargeType"))
			Point=ChkNumeric(Trim(request("Point")))
			Edays=ChkNumeric(Trim(request("Edays")))
			Reason=Request("Reason")
		
			if ChargeType="" then
				ChargeType=1
			else
				ChargeType=Clng(ChargeType)
			end if
			
			if ChargeType=1 and Point=0 then
 				Call Actcms.ActErr("请输入要追加的用户点数","","1")
			end if
			if ChargeType=2 and Edays=0 then
 				Call Actcms.ActErr("请输入要追加的天数","","1")
			end if
		    if Reason="" Then
 				Call Actcms.ActErr("请输入操作原因","","1")
 			end if
 			Set rsUser=Server.CreateObject("Adodb.RecordSet")
			sqlUser="select * from User_ACT where UserID=" & UserID
			rsUser.Open sqlUser,Conn,1,3
			if rsUser.bof and rsUser.eof then
 				rsUser.close:set rsUser=Nothing
 				Call Actcms.ActErr("找不到指定的用户","","1")
				exit sub
			end if
			If Round(rsUser("Money"))<Round(Money) Then
			  rsUser.close:set rsUser=Nothing
 				Call Actcms.ActErr("该用户的可用资金不足","","1")
			 exit sub
			End If
			'rsUser("Money")=rsUser("Money")-Money
			if ChargeType<>"1" then
 				ValidDays=rsUser("Edays")
				tmpDays=ValidDays-DateDiff("D",rsUser("BeginDate"),now())
				if tmpDays>0 then
					rsUser("Edays")=rsUser("Edays")+Edays
				else
					rsUser("BeginDate")=now
					rsUser("Edays")=Edays
				end if
			end if
			rsUser.update
			
			'消费记录
			If Money>0 Then
			 if ChargeType=2 Then
			  Call ACTCMS.MoneyInOrOut(rsUser("UserID"),rsUser("RealName"),Money,4,2,now,0,RSQL(Trim(Request.Cookies(AcTCMSN)("AdminName"))),"用于兑换有效天数",0,0)
			 else
			  Call ACTCMS.MoneyInOrOut(rsUser("UserID"),rsUser("RealName"),Money,4,2,now,0,RSQL(Trim(Request.Cookies(AcTCMSN)("AdminName"))),"用于兑换点券",0,0)
			 end if
			end if
			
			if ChargeType=1 then
			 Call ACTCMS.PointInOrOut(0,0,rsUser("UserID"),1,Point,RSQL(Trim(Request.Cookies(AcTCMSN)("AdminName"))),Reason,0)
			else
			 Call ACTCMS.EdaysInOrOut(rsUser("UserID"),1,Edays,RSQL(Trim(Request.Cookies(AcTCMSN)("AdminName"))),Reason)
			end if
			rsUser.Close:set rsUser=Nothing
			Call Actcms.ActErr("操作成功","User/User_Admin.asp","")
 			rsUser.Close:set rsUser=Nothing
 			Response.End
		end sub
		
 	
	
		'添加会员资金
		Sub AddZJ()
		dim rsUser,sqlUser
			UserID=ChkNumeric(UserID)
			if UserID=0 then Response.Write("<script>alert('参数不足！');history.back();</script>")
			Set rsUser=Conn.Execute("select * from User_ACT where UserID=" & UserID)
			if rsUser.bof and rsUser.eof then
				rsUser.close:set rsUser=Nothing
				Call Actcms.ActErr("找不到指定的用户","","1")
				exit sub
			end if
		%>
		<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
		<FORM name="myform" action="?" method="post">
			<TR >
			  <TD colspan="2" align="center" class="bg_tr"><b>用 户 续 费(增加资金)</b></TD>
		   </TR>
			<TR > 
			  <TD width="25%" height="28" align="right" ><b>用户名：</b></TD>
			  <TD width="75%"><%=rsUser("UserName")%></TD>
			</TR>
			<TR > 
			  <TD width="25%" height="28" align="right" ><strong>可用资金：</strong></TD>
			  <TD width="75%"><%=rsUser("Money")%> 元</TD>
			</TR>
			<TR > 
			  <TD width="25%" height="28" align="right" ><strong>用户级别：</strong></TD>
			  <TD width="75%"><%=GetGroupName(rsUser("GroupID")) %></TD>
			</TR>
			<TR  >
			  <TD height="28" align="right" ><strong>资金来源：</strong></TD>
			  <TD><input name="MoneyType" type="radio" id="ChargeType" checked onclick="document.all.Remark.value='银行汇款';" value="2">银行汇款
			      <input name="MoneyType" type="radio" id="ChargeType" onclick="document.all.Remark.value='现金收取';" value="1">其它（如：现金）
		      </TD>
			</TR>
			<TR  >
			  <TD height="28" align="right" ><strong>汇款日期：</strong></TD>
			  <TD><input name="PayTime" type="text" id="PayTime"  onClick="WdatePicker()"  value="<%=formatdatetime(now,2)%>" size="15" class="Ainput"></TD>
			</TR>
			<TR  >
			  <TD height="28" align="right" ><strong>续费金额：</strong></TD>
			  <TD> <input name="Money" type="text" id="Money" value="100" size="15" class="Ainput">
			  元</TD>
			</TR>
			
			
			<TR >
			  <TD height="28" align="right" ><strong>备注：</strong></TD>
			  <TD> <input name="Remark" type="text" id="Remark" value="银行汇款" size="55" class="Ainput"></TD>
			</TR>
			<TR > 
			  <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveAddZJ"> 
			  <input name=Submit  class='button' type=submit id="Submit" value="&nbsp;保存续费结果&nbsp;" > <input name="UserID" type="hidden" id="UserID" value="<%=rsUser("UserID")%>"><input class='button' type='button' value=' 返回 ' onclick='javascript:history.back();'></TD>
			</TR>
		</form>
	    </TABLE>
		<%
			rsUser.close : set rsUser=Nothing
		End Sub	
	
	
	
	
			'保存续费
		sub SaveAddZJ()
			dim UserID,MoneyType,Money,PayTime,Remark,sqlUser,rsUser
			Action=Trim(request("Action"))
			UserID=ChkNumeric(request("UserID"))
			if UserID=0 then
				Call Actcms.ActErr("参数不足","","1")
				exit sub
			end if
			MoneyType=Trim(request("MoneyType"))
			Money=actcms.s("Money")
			PayTime=actcms.s("PayTime")
			Remark=actcms.s("Remark")
            If Not IsDate(PayTime) Then
				Call Actcms.ActErr("汇款日期格式有误","","1")
			end if
			if ChkNumeric(Money)=0 then
				Call Actcms.ActErr("请输入要续费金额","","1")
			end if
		    if Remark="" Then
				Call Actcms.ActErr("请输入备注","","1")
			end if
			Set rsUser=Server.CreateObject("Adodb.RecordSet")
			sqlUser="select * from User_ACT where UserID=" & UserID
			rsUser.Open sqlUser,Conn,1,3
			if rsUser.bof and rsUser.eof then
				Call Actcms.ActErr("找不到指定的用户","","1")
				rsUser.close:set rsUser=Nothing
				exit sub
			end if
				'rsUser("Money")=rsUser("Money")+Money
			rsUser.update
			
			Call actcms.MoneyInOrOut(rsUser("UserID"),rsUser("RealName"),Money,MoneyType,1,PayTime,"0",RSQL(Trim(Request.Cookies(AcTCMSN)("AdminName"))),Remark,0,0)
			Call Actcms.ActErr("操作成功","User/User_Admin.asp","")
		end sub
	
	
	
	
	
	
	
	Sub Del()'删除
		IF UserID = "" Then
			ShowErr = "<li>请指定要删除的用户</li>"
			Call Actcms.ActErr(ShowErr,"","1")
			Response.End
		End IF
		If instr(UserID,",")>0 then
			UserID=replace(UserID," ","")
			Sql="delete from User_ACT where UserID in (" & UserID & ") "
		Else
			Sql="delete from User_ACT where UserID=" & ChkNumeric(UserID) & " "
		End if
		Conn.Execute sql
 		Call Actcms.ActErr("用户删除成功","User/User_Admin.asp","")
 		Response.End
    End Sub
	Sub Locked()'锁定
		IF UserID = "" Then
 			Call Actcms.ActErr("请指定要锁定的用户","","1")
 			Response.End
		End IF
		If instr(UserID,",")>0 then
			UserID=replace(UserID," ","")
			Sql="Update User_ACT set locked=1 where UserID in (" & UserID & ")"
		Else
			Sql="Update User_ACT set locked=1 where UserID=" & ChkNumeric(UserID)
		End if
		Conn.Execute sql:Set Conn=nothing
 		Call Actcms.ActErr("用户锁定成功","User/User_Admin.asp?ModeID="&ModeID&"","")
 		Response.End
    End Sub
	Sub UnLocked()'解锁
		IF UserID = "" Then
			ShowErr = "<li>请指定要解锁的用户</li>"
			Call Actcms.ActErr("请指定要解锁的用户","","")
 			Response.End
		End IF
		If instr(UserID,",")>0 then
			UserID=replace(UserID," ","")
			Sql="Update User_ACT set locked=0 where UserID in (" & UserID & ")"
		Else
			Sql="Update User_ACT set locked=0 where UserID=" & ChkNumeric(UserID)
		End if
		Conn.Execute sql:Set Conn=nothing
		Call Actcms.ActErr("用户解锁成功","User/User_Admin.asp?ModeID="&ModeID&"","")
  		Response.End
    End Sub
	Sub MoveUser()'移动
	Dim GroupID, RsGroup
		GroupID = ChkNumeric(Request("GroupID"))
		IF UserID = "" Then
 			 Call Actcms.ActErr("请指定要移动的用户","","1")
 			Response.End
		End IF
		IF GroupID = 0 Then Response.Write "目标用户组不存在":Response.end
		UserID=replace(UserID," ","")
		Set RsGroup=Conn.Execute("Select GroupSetting,GroupName From Group_ACT Where GroupID="&GroupID&"")
		If Not (RsGroup.Bof and RsGroup.Eof) then
			IF IsSqlDataBase = 1 Then
				Conn.Execute("Update User_ACT set GroupID=" & GroupID & " where UserID in (" & UserID & ")")
			Else
				Conn.Execute("Update User_ACT set GroupID=" & GroupID & " where UserID in (" & UserID & ")")
			End IF
		Else
			ShowErr = "<li>请指定目标用户组</li>"
			Call Actcms.ActErr(ShowErr,"","1")
			Response.End
			Exit Sub
		End if
  		Call Actcms.ActErr("<li>已经成功将选定用户设为“<font color=red>"&RsGroup(1)&"</font>","User/User_Admin.asp?ModeID="&ModeID&"","")
 		RsGroup.Close : Set RsGroup=Nothing
		Response.End
    End Sub
 Sub Main
	Dim strLocalUrl,ValidDays,tmpDays
	strLocalUrl = request.ServerVariables("SCRIPT_NAME")
	Dim intPageNow
	intPageNow = request.QueryString("page")
	Dim intPageSize, strPageInfo
	intPageSize = 20
	Dim arrRecordInfo, i
	Dim sql, sqlCount,Sqls,UserSearch,pages,StrGuide
	UserSearch = Request("UserSearch")
	pages = "UserSearch="&Request("UserSearch")&"&page"
	Select Case UserSearch
		Case 1
			Sqls = "Where Locked = 1 "
			StrGuide = "所有被锁住的用户"
		Case 2
			Sqls = " Where GroupID = 2  "
			StrGuide = "待管理员认证用户"
		Case 3
			Sqls = " Where GroupID = 1  "
			StrGuide = "待邮件验证的用户"
		Case 4
			Sqls = " where datediff('h',LoginTime," & NowString & ")<25 "
			StrGuide = "最近24小时内登录的用户"
		Case 5
			Sqls = " where datediff('h',RegDate," & NowString & ")<25 "
			StrGuide = "最近24小时内注册的用户"
		Case 6
			Sqls = " where GroupID = "&  ChkNumeric(Request.QueryString("GroupID"))&" "
			StrGuide = "查询结果"
			pages = "UserSearch="&Request.QueryString("UserSearch")&"&GroupID="&ChkNumeric(Request.QueryString("GroupID"))&"&page"
		Case 7
			StrGuide = "查询结果"
			pages = "UserSearch="&Request("UserSearch")&"&Email="&Request("Email")&"&usernamechk="&Request("usernamechk")&"&username="&Request("username")&"&GroupID="&ChkNumeric(Request("GroupID"))&"&page"
			UserID =ChkNumeric(UserID)
					if UserID>0 then
						Sqls = " UserID="&UserID&" "
					else 
						Sqls=""
						if request("username")<>"" then
							if request("usernamechk")="yes" then
								Sqls=Sqls & " username='"&request("username")&"' "
							else
								Sqls=Sqls &" username like '%"&request("username")&"%' "
							end if
						end if
						if cint(request("GroupID"))>0 then
							if Sqls="" then
								Sqls=Sqls & " GroupID="&request("GroupID")&" "
							else
								Sqls=Sqls & " and GroupID="&request("GroupID")&" "
							end if
						end if	
				end if	
			
					if request("Email")<>"" then
							if Sqls="" then
								Sqls=Sqls & " Email like '%"&request("Email")&"%'"
							else
								Sqls=Sqls & " and Email like '%"&request("Email")&"%'"
							end if
						end if
				if Sqls <> "" then Sqls = " where "&Sqls&"  "			
		Case Else
			pages = "page"
			StrGuide = "所有用户"
	End Select
	sql = "SELECT [UserID], [GroupID], [UserName],[Locked], [Loginip],[LoginNumber],[LoginTime],ChargeType,[Umodeid]" & _
		" FROM [User_ACT]" &Sqls& _
		"ORDER BY [UserID] DESC"
	sqlCount = "SELECT Count([UserID])" & _
			" FROM [User_ACT]"&Sqls
		Dim clsRecordInfo
		Set clsRecordInfo = New Cls_PageView
			clsRecordInfo.intRecordCount = 2816
			clsRecordInfo.strSqlCount = sqlCount
			clsRecordInfo.strSql = sql
			clsRecordInfo.intPageSize = intPageSize
			clsRecordInfo.intPageNow = intPageNow
			clsRecordInfo.strPageUrl = strLocalUrl
			clsRecordInfo.strPageVar = pages
		clsRecordInfo.objConn = Conn		
		arrRecordInfo = clsRecordInfo.arrRecordInfo
		strPageInfo = clsRecordInfo.strPageInfo
		Set clsRecordInfo = nothing
 %> 
 <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
    <form name="form2" method="post" action="User_Admin.asp"><tr>
      <td align="center" class="bg_tr">快速搜索</td>
    </tr>
    <tr>
      <td><strong>用户名：</strong>
      <input name="username" type="text" class="Ainput" id="username" size="20">
       <label for="usernamechk"><strong>用户名完整匹配</strong>
       <input name="usernamechk" type="checkbox" id="usernamechk" value="yes" checked></label>
       <strong>&nbsp;用户组：</strong>
      <select size="1" name="GroupID">
        <option value="0" selected>全部</option>
        <%=GroupOption(3)%>
      </select>
		   <input name="UserSearch" type="hidden" id="UserSearch" value="7">
		  <input name=Submit class="act_btn" type=button value="  快速搜索   " onclick=CheckInfo1()></td>
    </tr>
 </form> </table>

  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
 <form name="Userform" method="Post" action="User_Admin.asp" onSubmit="return confirm('确定要执行选定的操作吗？');">
	<tr>
      <td colspan="12" class="bg_tr">您现在的位置：<a href="?"  ><font class="bg_tr">管理首页</font>  </a>&gt;&gt; 注册用户管理&gt;&gt;<%= StrGuide %></td>
    </tr>
    <tr>
      <td  align="center" nowrap><STRONG>选中</STRONG></td>
      <td  align="center" nowrap><STRONG>ID</STRONG></td>
      <td  align="center" nowrap><STRONG>所属模型</STRONG></td>
      <td  align="center" nowrap><STRONG>用户名</STRONG></td>
      <td  align="center" ><STRONG>所属用户组</STRONG></td>
      <td align="center" nowrap><STRONG>最后登录IP</STRONG></td>
      <td  align="center"><STRONG>最后登录时间</STRONG></td>
      <td  align="center" nowrap><STRONG>登录次数</STRONG></td>
      <td  align="center" nowrap><STRONG>状态</STRONG></td>
      <td  colspan="2" align="center" nowrap><strong>管理操作</strong></td>
      <td align="center" nowrap><strong>删除</strong></td>
    </tr>
	 <%
		Dim bgColor
		If IsArray(arrRecordInfo) Then
			For i = 0 to UBound(arrRecordInfo, 2)
	%>
    <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
      <td align="center" ><input name="UserID" type="checkbox" id="UserID" value="<%= arrRecordInfo(0,i) %>"></td>
      <td align="center" ><%= arrRecordInfo(0,i) %></td>
      <td align="center" ><%=actcms.ACT_U(arrRecordInfo(8,i),1) %></td>
      <td align="center" ><%=Replace(arrRecordInfo(2,i),request("username"),"<font color=""RED"">"&request("username")&"</font>") %></td>
      <td align="center" ><%=GetGroupName(arrRecordInfo(1,i)) %></td>
      <td align="center" ><%=arrRecordInfo(4,i)%></td>
      <td align="center" ><%=arrRecordInfo(6,i)%></td>
      <td align="center" ><%=arrRecordInfo(5,i)%></td>
      <td align="center" ><% IF arrRecordInfo(3,i) = 0 Then Response.Write "正常" Else Response.Write "<font color=red>已锁定</font>" %></td>
      <td colspan="2"  align="center">
	
	<select name="page" size="1" onChange="javascript:window.location=this.options[this.selectedIndex].value;">
      <option value="" >请选择</option>
      <option value="ACT.E.ASP?Action=Edit&UserID=<%=arrRecordInfo(0,i)  %>" >修改</option>
	 <% IF arrRecordInfo(3,i) = 0 Then %>
	   <option value="?Action=Lock&UserID=<%= arrRecordInfo(0,i) %>">锁定</option>
	 <%  Else%>
	 <option value="?Action=UnLock&UserID=<%= arrRecordInfo(0,i) %>">解锁</option>
	<%End If 
	
	 If arrRecordInfo(7,i)=1 Then
	   Response.Write "<option value='?Action=AddMoney&UserID=" & arrRecordInfo(0,i) & "'>续点数</option>"
	 ElseIf arrRecordInfo(7,I)=2 Then
	   Response.Write "<option value='?Action=AddMoney&UserID=" & arrRecordInfo(0,i) & "'>续天数</option>"
	 End IF
	
	 %>
    
     <option value="?Action=AddZJ&UserID=<%= arrRecordInfo(0,i) %>">续费</option>
     </select></td>
      <td  align="center"><a href="?Action=Del&UserID=<%= arrRecordInfo(0,i) %>" onClick="return confirm('确认删除此用户吗?')">删除</a></td>
    </tr>
	<% 
	Next
	End If
	%>
    <tr >
      <td height="25" colspan="12"><label for="chkAll">&nbsp;
        <input name="ChkAll" type="checkbox" id="ChkAll" onClick="CheckAll(this.form)" value="checkbox">选中本页显示的所有用户</label>
        <strong>操作：</strong> 
					 <label for="Del"><input ID="Del" name="Action" type="radio" value="Del" checked onClick="document.Userform.GroupID.disabled=true">
					  删除&nbsp;&nbsp;&nbsp;&nbsp;</label>
					  <label for="Lock"><input ID="Lock" name="Action" type="radio" value="Lock" onClick="document.Userform.GroupID.disabled=true">
					  锁定 &nbsp;&nbsp;&nbsp;</label>
					  <label for="UnLock"><input ID="UnLock" name="Action" type="radio" value="UnLock" onClick="document.Userform.GroupID.disabled=true">
					  解锁 &nbsp;&nbsp;&nbsp; </label>
					  <label for="Move"><input ID="Move" name="Action" type="radio" value="Move" onClick="document.Userform.GroupID.disabled=false">
					  移动到</label>
					  <select name="GroupID" id="GroupID" disabled>
					  <%=GroupOption(3)%>
					  </select>
				&nbsp;&nbsp;
	  <input type="submit" name="Submit"  class="act_btn"  value=" 执&nbsp;&nbsp;行 " >	  </td>
    </tr>
    <tr >
      <td height="25" colspan="12" align="center"><%= strPageInfo%></td>
    </tr>
  </form></table>

<script language="javascript">

		function CheckAll(form)
		  {  
		 for (var i=0;i<form.elements.length;i++)  
			{  
			   var e = Userform.elements[i];  
			   if (e.name != 'ChkAll'&&e.type=="checkbox")  
			   e.checked = Userform.ChkAll.checked;  
		   }  
	  }
function overColor(Obj)
{
	var elements=Obj.childNodes;
	for(var i=0;i<elements.length;i++)
	{
		elements[i].className="tdbg1"
		Obj.bgColor="";
	}
	
}
function outColor(Obj)
{
	var elements=Obj.childNodes;
	for(var i=0;i<elements.length;i++)
	{
		elements[i].className="tdbg";
		Obj.bgColor="";
	}
}


		</script>
<% 
 End Sub
	Public Function GetGroupName(GroupID)
	 On Error Resume Next
	 Dim Grs
	 Set Grs=ACTCMS.ACTEXE("Select Groupname From Group_ACT Where GroupID=" & GroupID)
	 If Not Grs.eof Then 
		GetGroupName = Grs(0)
	 Else
		GetGroupName="<font color=red>该用户组已被删除</font>"
	 End If 
	End Function
	

	Public Function GroupOption(Selected)
	 Dim RSObj,GroupName:Set RSObj=Server.CreateObject("Adodb.Recordset")
	    RSObj.Open "Select GroupID,GroupSetting,GroupName From Group_ACT",Conn,1,1
	  	Do While Not RSObj.Eof
		   GroupName=RSObj(2)
		   IF Selected=RSObj(0) Then
			GroupOption=GroupOption & "<option value=""" & RSObj(0) & """ Selected>" &GroupName & "</option>"
		   Else
			GroupOption=GroupOption & "<option value=""" & RSObj(0) & """>" & GroupName & "</option>"
		   End If
		RSObj.MoveNext
		Loop
	  RSObj.Close:Set RSObj=Nothing
	End Function	
	
  
	Sub UserSearch() %>
		<form name="form2" method="post" action="User_Admin.asp">
		<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1"  class="table">
		<tr>
			<td height="25" colspan="2" align="center" class="bg_tr"><strong>高级查询</strong></td>
		</tr>
		<tr>
			<td width="20%" align="right"><strong>用户ID：</strong></td>
			<td width="80%"><input name="userid" type="text" class="Ainput" size="45"></td>
		</tr>
		<tr>
			<td width="20%" align="right"><strong>用户名：</strong></td>
			<td width="80%"><input name="username" type="text" class="Ainput" size="45">
		  <label for="usernamechk1">&nbsp;<input type="checkbox" id="usernamechk1" name="usernamechk" value="yes" checked>用户名完整匹配</label></td>
		</tr>
		<tr>
			<td width="20%" align="right"><strong>用户组：</strong></td>
			<td width="80%">
			<select size="1" name="GroupID">
			<option value="0" selected>全部</option>
			<%=GroupOption(3)%>
			</select>		  </td>
		</tr>
		<tr>
			<td width="20%" align="right"><strong>Email包含：</strong></td>
			<td width="80%"><input size="45" name="Email" type=text class="Ainput" ></td>
		</tr>
		<tr>
		  <td>&nbsp;</td>
		  <td>
		  <input name="UserSearch" type="hidden" id="UserSearch" value="7">
		  <input name=Submit  class="act_btn"  type=button value="   搜  索   " onclick=CheckInfo1()></td>
		  </tr>
		</table>
		</form>	
		<%End Sub
 CloseConn %>
 <script language="javascript">	
	function CheckInfo1()
		{
		form2.Submit.value="正在提交数据,请稍等...";
		form2.Submit.disabled=true;	
	    form2.submit();	
		}
	</script>
	<% Sub AddMoney()
		Dim rsUser,sqlUser,tmpDays
		UserID=ChkNumeric(Request.QueryString("UserID"))
		 	 	

		If UserID=0 then  Call Actcms.ActErr("参数不足","User/User_Admin.asp?ModeID="&ModeID&"","")
		Set rsUser=Conn.Execute("select * from User_ACT where UserID=" & UserID)
		If rsUser.bof and rsUser.eof then
			rsUser.close:set rsUser=Nothing
			Call Actcms.ActErr("找不到指定的用户","User/User_Admin.asp?ModeID="&ModeID&"","1")
 			Exit sub
		End if
		if rsUser("ChargeType")=3 Then
			  rsUser.Close:Set rsUser=Nothing
			  Call Actcms.ActErr("无限期用户无需续费操作","User/User_Admin.asp?ModeID="&ModeID&"","1")
			  Exit Sub
		End if
	%>
	
	
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1"  class="table">
		<form name="form3" method="post" action="?">
			<TR  >
			  <TD height="28" colspan="2" align="center"><b>用 户 续 费</b></TD>
		   </TR>
			<TR > 
			  <TD width="25%" height="28" align="right" ><b>用户名：</b></TD>
			  <TD width="75%"><%=rsUser("UserName")%></TD>
			</TR>
			<TR > 
			  <TD width="25%" height="28" align="right" ><strong>用户级别：</strong></TD>
			  <TD width="75%"><%=GetGroupName(rsUser("GroupID"))%></TD>
			</TR>
			<TR  >
			  <TD height="28" align="right" ><strong>计费方式：</strong></TD>
			  <TD><%
			  if rsUser("ChargeType")=1 then
				Response.Write "扣点数"
			  else
				Response.Write "有效期"
			  end if
			  %>
				<input name="ChargeType" type="hidden" id="ChargeType" value="<%=rsUser("ChargeType")%>">			  </TD>
			</TR>
			<TR > 
			  <TD width="25%" height="28" align="right" ><strong>可用资金：</strong></TD>
			  <TD width="75%"><%=rsUser("Money")%>元人民币</TD>
			</TR>
			<%if rsUser("ChargeType")=1 then%>
			<TR  >
			  <TD height="28" align="right" ><strong>目前的用户点数：</strong></TD>
			  <TD><%=rsUser("Point")%> 点</TD>
			</TR>
			<TR  >
			  <TD height="28" align="right" ><strong>追加点数：</strong></TD>
			  <TD> <input name="Point" class="Ainput"  type="text" id="Point" value="100" size="10" maxlength="10">
			  点</TD>
			</TR>
			<%else%>
			<TR >
			  <TD height="28" align="right" ><strong>目前的有效期限信息：</strong></TD>
			  <TD>
			  <%
			  Response.Write "开始计算日期" & FormatDateTime(rsUser("BeginDate"),2) & "&nbsp;&nbsp;&nbsp;&nbsp;有 效 期：" & rsUser("Edays")
			 
				Response.Write "天"
			 
			  Response.Write "<br>"
			  tmpDays=rsUser("Edays")-DateDiff("D",rsUser("BeginDate"),now())
			  if tmpDays>=0 then
				Response.Write "尚有 <font color=blue>" & tmpDays & "</font> 天到期"
			  else
				Response.Write "已经过期 <font color=#ff6600>" & abs(tmpDays) & "</font> 天"
			  end if
			  %>			  </TD>
			</TR>
			<tr >
			  <td height="60" align="right" ><strong>追加天数：</strong><br></td>
			  <td>
			  <input name="Edays" class="Ainput"  type="text" id="Edays" value="100" size="10" maxlength="10">
			  天<br />
			  若目前用户尚未到期，则追加相应天数<br />
若目前用户已经过了有效期，则有效期从续费之日起重新计数。</td>
			</tr>
			<%end if%>
			<tr >
			  <td height="30" align="right" ><strong>同时减去：</strong><br></td>
			  <td>
			  <input name="Money" type="text" class="Ainput"  id="Money" value="100" size="10" maxlength="10"> 元人民币
			  <font color=red>
			  <%if rsUser("ChargeType")=1 then %>
			   资金与点券的默认比率：<%=actcms.ActCMS_Sys(22)%>:1
			  <%else%>
			  资金与有效期的默认比率：<%=actcms.ActCMS_Sys(23)%>:1
			  <%end if%>
			  </font> 不想扣除资金，请输入0
			  </td>
			</tr>
			<TR >
			  <TD height="28" align="right" ><strong>请输入原因：</strong></TD>
			  <TD> <input name="Reason" class="Ainput"  type="text" id="Reason" value="<%If rsUser("ChargeType")=1 Then Response.Write "续点券操作" Else Response.Write "续有效天数操作"%>" size="55"></TD>
			</TR>
			<TR > 
			  <TD height="40" colspan="2" align="center">
          <input name="Action" type="hidden" id="Action" value="SaveAddMoney"> 
 		   <input name="UserID" type="hidden" id="UserID" value="<%= rsUser("UserID") %>">
		   <input name="ModeID" type="hidden" id="ModeID" value="<%= ModeID %>">
		   <input type="submit" name="Submit2" class="act_btn" value=" 保存 ">
	      <input type="reset" name="Submit3"  class="act_btn" value=" 重置 "></TD>
			</TR>
		</form>
	    </TABLE>
		<%
			rsUser.close : set rsUser=Nothing
End Sub 
 %></body>
</html>
