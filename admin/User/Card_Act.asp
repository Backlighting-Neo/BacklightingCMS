<!--#include file="../ACT.Function.asp"-->
<!--#include file="../../ACT_inc/cls_pageview.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>充值卡管理</title>
<link href="../Images/style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.STYLE1 {font-weight: bold}
-->
</style>
</head>
<body>
<%ConnectionDatabase
		dim MaxPerPage,RS,TotalPut,TotalPages,I,CurrentPage,SQL,ComeUrl
		  MaxPerPage=20
			Response.Write"<table width=""98%""  align=""center""   height=""25"" border=""0"" cellspacing=""1"" cellpadding=""2""  class=""table"">"
			Response.Write " <tr>"
			Response.Write"	<td height=""25"" class=""bg_tr""> "
			Response.Write " <strong>操作导航:</strong>&nbsp;<a href=""?"">所有充值卡</a> | <a href=""?status=1"">未使用充值卡</a> | <a href=""?status=2"">已使用充值卡</a> | <a href=""?status=3"">已失效充值卡</a> | <a href=""?status=4"">未失效充值卡</a> | <a href=""?action=Add"">添加充值卡</a> | <a href=""?action=AddMore"">批量生成充值卡</a>"
			Response.Write	" </td>"
			Response.Write " </tr>"
			Response.Write"</TABLE>"
		ComeUrl=Cstr(Request.ServerVariables("HTTP_REFERER"))
		
		Select Case request("Action")
		 Case "Add","Edit"
		  Call Add()
		 Case  "DoAdd"
		  Call DoAdd()
		 Case "AddMore"
		  Call AddMore()
		 Case "DoAddMore"
		  Call DoAddMore()
		 Case "Del"
		  Call Del()
		 Case Else
		  Call CardList()
		End Select
	
	'点卡列表
	Sub CardList()
		%>
<table width="98%" style="MARGIN-TOP: 3px" border="0" align="center" cellspacing="1" cellpadding="2" class="table">
  <tr class="bg_tr">
    <td width="38" height="25" align="center" class="bg_tr"><strong>选中</strong></td>
    <td width="116" align="center" class="bg_tr">充值卡名称</td>
    <td width="116" height="25" align="center" class="bg_tr"><strong>充值卡号</strong></td>
    <td width="88" height="25" align="center" class="bg_tr"><strong>密码</strong></td>
    <td width="75" align="center" class="bg_tr"><strong>面值</strong></td>
    <td width="82" height="25" align="center" class="bg_tr"><strong>点数</strong></td>
    <td width="100" align="center" class="bg_tr"><strong>过期时间</strong></td>
    <td width="80" height="25" align="center" class="bg_tr"><strong>出售情况</strong></td>
    <td width="80" align="center" class="bg_tr"><strong>使用情况</strong></td>
    <td width="100" height="25" align="center" class="bg_tr"><strong>使用者</strong></td>
    <td width="100" height="25" align="center" class="bg_tr"><strong>充值时间</strong></td>
    <td width="130" align="center" class="bg_tr"><strong>操作</strong></td>
  </tr>
  <%
  
  
  
  
  	Dim strLocalUrl
	strLocalUrl = request.ServerVariables("SCRIPT_NAME")
	
	Dim intPageNow
	intPageNow = request.QueryString("page")
	
	Dim intPageSize, strPageInfo
	intPageSize = 30
	
	Dim arrRecordInfo, i
	Dim sql, sqlCount,Sqls,pages
  CurrentPage	= ChkNumeric(request("page"))
  Dim Param:Param=" where 1=1"
  pages ="Status="&ChkNumeric(request("Status"))&"&page"
  Select Case  ChkNumeric(request("Status"))
   Case 1
     Param=Param & " And IsUsed=0"
   Case 2
     Param=Param & " And IsUsed=1"
   Case 3
     Param=Param & " And datediff('d',EndDate,"&NowString&")>0"
   Case 4
     Param=Param & " And datediff('d',EndDate,"&NowString&")<0"
  End Select
  
  	sql = "SELECT ID,CardNum,CardPass,Money,ValidNum,ValidUnit,AddDate,EndDate,UseDate,UserID,IsUsed,IsSale,title" & _
		" FROM [Card_ACT]" &Param& _
		" ORDER BY [ID] DESC"
  	sqlCount = "SELECT Count([ID])" & _
			" FROM [Card_ACT]"&Param
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
			SQL = clsRecordInfo.arrRecordInfo
			strPageInfo = clsRecordInfo.strPageInfo
			Set clsRecordInfo = nothing
 
%>		

<%
  Dim InPoint,OutPoint
 %>
 <form name=selform method=post action=?action=Del>
 <%
If IsArray(SQL) Then
			For i = 0 to UBound(SQL, 2)
	%>
  <tr height="25"  onMouseOver=overColor(this) onMouseOut=outColor(this)>
    <td align="center"  ><input type="checkbox" name="id" value="<%=SQL(0,i)%>"></td>
    <td align="center"  ><%=SQL(12,i)%></td>
    <td align="center"  ><%=SQL(1,i)%></td>
    <td align="center"  ><%=SQL(2,i)%></td>
    <td align="center"  ><%=SQL(3,i)%>元</td>
    <td align="center"  ><%Response.Write SQL(4,I)
	if SQL(5,I)=1 Then 
	 Response.Write "点" 
	ELSEIf SQL(5,I)=2 Then 
	 Response.Write "天" 
	elseif SQL(5,I)=3 Then
	 response.write "元"
	end if%></td>
    <td align="center"  ><%Response.Write formatdatetime(SQL(7,I),2)%></td>
    <td align="center"  >
	<%
	IF SQL(11,I)=1 Then
	 Response.Write "已售出"
	Else
	 Response.Write "<font color=red>未出售</font>" 
	End If
	%></td>
    <td align="center"  >
	<%
	IF SQL(10,I)=1 Then
	 Response.Write "<font color='#a7a7a7'>已使用</font>"
	Else
	 Response.Write "<font color=red>未使用</font>" 
	End If
	%></td>
    <td align="center"  ><%=ACTCMS.UserM(SQL(9,I))%></td>
    <td align="center"  >
	<%if Isdate(Sql(8,i)) then
	   response.write formatdatetime(SQL(8,i),2)
	  end if%></td>
	<td align="center"  >
	<%if SQL(11,I)<>1 and SQL(10,I)<>1 then%>
	<a href="?action=Edit&ID=<%=SQL(0,i)%>">修改</a> <a href="?action=Del&ID=<%=SQL(0,i)%>">删除</a>
	<%end if%>	</td>
  </tr>
  <%Next
	End If
  
  Response.Write "<tr><td height='30' colspan=12  >"
  Response.Write "&nbsp;&nbsp;<input id=""chkAll"" onClick=""CheckAll(this.form)"" type=""checkbox"" value=""checkbox""  name=""chkAll""><label for=""chkAll"">全选&nbsp;&nbsp;</label><input class=act_btn type=""submit"" name=""Submit2"" value="" 删除选中的充值卡 "" onclick=""{if(confirm('此操作不可逆，确定要删除选中的记录吗?')){this.document.selform.submit();return true;}return false;}"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
   Response.Write "</td></tr>"

 %>
 <tr >
      <td height="25" colspan="12" align="center" class="td_bg"><%= strPageInfo%></td>
    </tr>
 </form>
</table>
<table border="0" style="margin-top:20px" width="90%" align=center>
<tr><td><Font color=red><strong>提示：</strong>
已售出或已使用的充值卡，不允许删除，修改等操作。</font>
</td></tr>
</table>
<% End Sub

 '删除充值卡
 Sub Del()
  Dim ID:ID=Replace(request("ID"),"Card_Act.asp?action=Add","Card_Act.asp?action=Add")
  Conn.Execute("Delete From Card_ACT Where ID In(" & ID &") and IsSale=0 and IsUsed=0")
  Response.Write "<script>alert('删除成功！');location.href='" & Request.Servervariables("http_referer") & "';</script>"
 End Sub
		
		'批量添加充值卡
		Sub AddMore()
		%>
		
  <table width='98%' border='0' align='center' cellpadding='2' cellspacing='1' class='table' >
   <form method='post' action='?' name='myform'>
    <tr > 
      <td height='22' colspan='2' class="bg_tr"> <div align='center'><strong>批 量 生 成 充 值 卡</strong></div></td>
    </tr>
   
   <tr>
      <td width='40%'><b>充值卡名称：</b></td>
      <td><input name='Title' type='text' class='Ainput'  id='Title' size='20' value="" maxlength='30'>如推销卡等等
        </td>
    </tr>    
    <tr > 
      <td width='40%'><strong>充值卡数量：</strong></td>
      <td width='60%'><input name='Nums' type='text' class='Ainput'  value='100' size='10' maxlength='10'>
        张</td>
    </tr>
    <tr >
      <td width='40%'><strong>充值卡号码前缀：</strong><br>
        例如：2006,Act2007等固定不变的字母或数字</td>
      <td width='60%'><input name='CardNumPrefix' type='text' class='Ainput'  id='CardNumPrefix' value='ACT2007' size='10' maxlength='10'></td>
    </tr>
    <tr >
      <td width='40%'><strong>充值卡号码位数：</strong><br>请输入包含前缀字符在内的总位数</td>
      <td width='60%'><input name='CardNumLen' type='text' class='Ainput'  id='CardNumLen' value='12' size='10' maxlength='10'>
        <font color='#0000FF'>建议设为10-15位</font></td>
    </tr>
    <tr >
      <td width='40%'><strong>充值卡密码位数：</strong></td>
      <td width='60%'><input name='PasswordLen' type='text' class='Ainput'  id='PasswordLen' value='6' size='10' maxlength='10'>
        <font color='#0000FF'>建议设为6-10位</font></td>
    </tr>
    <tr >
      <td><strong>卡密码构成方式：</strong><br>你可以选择数据或字母的组合</td>
      <td><input type="radio" name="zhtype" value="1" checked>纯数字 <input type="radio" name="zhtype" value="2">数字与字母随机组合 </td>
    </tr>
    <tr >
      <td width='40%'><strong>充值卡面值：</strong><br>
      即购买人需要花费的实际金额</td>
      <td width='60%'><input name='Money' type='text' class='Ainput'  id='Money' value='50' size='10'>
      元</td>
    </tr>
    <tr > 
      <td width='40%'><strong>充值卡点数、资金或有效期：</strong><br>
        购买人可以得到的点数、资金或有效期      </td>
      <td width='60%'><input name='ValidNum' type='text' class='Ainput'  id='ValidNum' value='50' size='10' maxlength='10'>
        <select name='ValidUnit' id='ValidUnit'>
               <option value='1' selected>点</option>
          <option value='2'>天</option>
          <option value='3'>元</option>
          <option value='4'>积分</option>
        </select></td>
    </tr>
    
    
	<tr >
	  <td><strong>允许使用此充值卡的用户组：</strong><br>
不限制请留空或全部选中。 </td>
	  <td ><%= actcms.GetGroup_CheckBox("allgroupid","",5)  %>	</td>
	  </tr>
      
      
<tr >
	  <td><strong>充值后自动归入的用户组：</strong><br>
  </td>
	  <td >
	  
	  <select name="grGroupID" id="grGroupID">
 	  <%= actcms.GetGroup_select("")  %>	
      </select>
      
      </td>
	  </tr>
      
      
      
      <tr >
	  <td><strong>到期后自动归入的用户组：</strong><br>
指用户选择充值卡为账户充值后,当账户里的点券,有效天数或资金用完后(具体根据该卡是点券卡,有数天数卡或资金卡而定)。将过期的用户自动归入低一级的用户级别。 </td>
	  <td >
	   <select name="ExpireGroupID" id="ExpireGroupID">
	  <%= actcms.GetGroup_select("")  %>	
      </select>
      </td>
	  </tr>     
    
    <tr >
      <td width='40%'><strong>充值截止期限：</strong><br>
      购买人必须在此日期前进行充值，否则自动失效</td>
      <td width='60%' ><input name='EndDate' type='text' class='Ainput'  id='EndDate' value='<%=now+365%>' size='10' maxlength='10'></td>
    </tr>
    <tr > 
      <td height='40' colspan='2' align='center'><input name='Action' type='hidden' id='Action' value='DoAddMore'> 
        <input  class="ACT_btn" type='submit' name='Submit' value=' 开始生成 ' style='cursor:hand;'> 
        &nbsp; <input class="ACT_btn" name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="window.location.href='?'" style='cursor:hand;'></td>
    </tr></form>
  </table>

		<%
		End Sub	
		'添加充值卡
		Sub Add()
		  Dim CardNum,PassWord,IsSale,IsUsed,Money,ValidNum,ValidUnit,EndDate,action1,allgroupid,Title,grGroupID,ExpireGroupID
		  Dim ID:ID=ChkNumeric(request("ID"))
		  if request("action")="Edit" then
		    Dim RS:Set RS=Server.CreateObject("adodb.recordset")
			rs.open "select * from Card_ACT where ID=" & ID,conn,1,1
			if rs.bof and rs.eof then
			  rs.close:set rs=nothing
			  Call ACTCMS.Alert("参数传递出错！","Card_Act.asp?action=Add")
			  Exit sub
			end if
			CardNum=rs("CardNum")
			PassWord=rs("CardPass")
			Money=rs("money")
			ValidNum=rs("ValidNum")
			ValidUnit=rs("ValidUnit")
			EndDate=rs("EndDate")
			IsSale=rs("IsSale")
			IsUsed=rs("IsUsed")
			Title=rs("Title")
			allgroupid=rs("allgroupid")
			grGroupID=rs("grGroupID")
			ExpireGroupID=rs("ExpireGroupID")
			action1="Edit"
			
		  else
		   IsSale=0:IsUsed=0:Money=50:ValidNum=50:ValidUnit=1:EndDate=Now+365
		  end if
		%>
  <table width='98%' border='0' align='center' cellpadding='2' cellspacing='1' class='table' >
		<form method='post' action='?action1=<%=action1%>&id=<%=ID%>' name='myform'>
	<tr  class="bg_tr"> 
      <td height='22' colspan='2' class="bg_tr"> <div align='center'><strong>
	  <%IF request("Action")="Edit" then
	   response.write "修 改 充 值 卡"
	    Else
		Response.Write "添 加 充 值 卡"
	    End If
		%></strong></div></td>
    </tr>
      <tr>
      <td width='40%'><b>充值卡名称：</b></td>
      <td><input name='Title' type='text' class='Ainput'  id='Title' size='20' value="<%=Title%>" maxlength='30'>如推销卡等等
        </td>
    </tr> 
    
   <tr <%if request("action")="Edit" Then response.write " style='display:none'"%>> 
      <td width='40%'><strong>添加方式：</strong></td>
      <td width='60%'><input name='AddType' type='radio' value='0' checked onClick="trSingle1.style.display='';trSingle2.style.display='';trBatch.style.display='none';"> 单张充值卡&nbsp;&nbsp;&nbsp;&nbsp;<input name='AddType' type='radio' value='1' onClick="trSingle1.style.display='none';trSingle2.style.display='none';trBatch.style.display='';">批量添加充值卡</td>
    </tr>
    <tr  id='trSingle1'>
      <td width='40%'><b>充值卡卡号：</b></td>
      <td><input name='CardNum' type='text' class='Ainput'  id='CardNum' size='20' value="<%=CardNum%>" maxlength='30'>
        <font color='#0000FF'>建议设为10-15位</font></td>
    </tr>
    <tr  id='trSingle2'>
      <td width='40%'><b>充值卡密码：</b></td>
      <td><input name='Password' type='text' class='Ainput'  id='Password' size='20' value="<%=PassWord%>" maxlength='30'>
        <font color='#0000FF'>建议设为6-10位 </font></td>
    </tr>
    <tr  id='trBatch' style='display:none'>
      <td width='40%'><b>格式文本：</b><br><font color='red'>请按照每行一张卡，每张卡按“卡号＋分隔符＋密码”的格式录入</font><br>
      例：734534759|Actf15f4ag5te（以“|”作为分隔符）</td>
      <td><textarea name='CardList' rows='10' cols='50'></textarea></td>
    </tr>
    <tr >
      <td width='40%'><strong>充值卡面值：</strong><br>
      即购买人需要花费的实际金额</td>
      <td width='60%'><input name='Money' type='text' class='Ainput'  id='Money' value='<%=Money%>' size='10'>
      元</td>
    </tr>
    <tr > 
      <td width='40%'><strong>充值卡点数、资金或有效期：</strong><br>
        购买人可以得到的点数、资金或有效期      </td>
      <td width='60%'><input name='ValidNum' value="<%=ValidNum%>" type='text' class='Ainput'  id='ValidNum'  size='10' maxlength='10'>
        <select name='ValidUnit' id='ValidUnit'>
          <option value='1' <%if ValidUnit="1" then response.write " selected"%>>点</option>
          <option value='2' <%if ValidUnit="2" then response.write " selected"%>>天</option>
          <option value='3' <%if ValidUnit="3" then response.write " selected"%>>元</option>
          <option value='4' <%if ValidUnit="4" then response.write " selected"%>>积分</option>
        </select></td>
    </tr>
    <tr >
      <td width='40%'><strong>充值截止期限：</strong><br>
      购买人必须在此日期前进行充值，否则自动失效</td>
      <td width='60%' ><input name='EndDate' type='text' class='Ainput'  id='EndDate' value='<%=EndDate%>' size='10' maxlength='10'></td>
    </tr>
	<tr >
	  <td><strong>允许使用此充值卡的用户组：</strong><br>
不限制请留空或全部选中。 </td>
	  <td ><%= actcms.GetGroup_CheckBox("allgroupid",allgroupid,5)  %>	</td>
	  </tr>
      
      
<tr >
	  <td><strong>充值后自动归入的用户组：</strong><br>
  </td>
	  <td >
	  
	  <select name="grGroupID" id="grGroupID">
 	  <%= actcms.GetGroup_select(grGroupID)  %>	
      </select>
      
      </td>
	  </tr>
      
      
      
      <tr >
	  <td><strong>到期后自动归入的用户组：</strong><br>
指用户选择充值卡为账户充值后,当账户里的点券,有效天数或资金用完后(具体根据该卡是点券卡,有数天数卡或资金卡而定)。将过期的用户自动归入低一级的用户级别。 </td>
	  <td >
	   <select name="ExpireGroupID" id="ExpireGroupID">
	  <%= actcms.GetGroup_select(ExpireGroupID)  %>	
      </select>
      </td>
	  </tr>      
      
	<tr >
      <td width='40%'><strong>是否出售：</strong><br>
      添加新充值卡，请选项未出售</td>
      <td width='60%' ><input name='issale' type='radio' id='issale' value='0'<%if issale=0 then response.write " checked"%>>未出售 <input name='issale' type='radio' id='issale' value='1'<%if issale=1 then response.write " checked"%>>已出售</td>
    </tr>
	<tr >
      <td width='40%'><strong>是否使用：</strong><br>
      添加新充值卡，请选项未使用</td>
      <td width='60%' ><input name='isused' type='radio' id='isused' value='0'<%if isused=0 then response.write " checked"%>>未使用 <input name='isused' type='radio' id='isused' value='1'<%if isused=1 then response.write " checked"%>>已使用</td>
    </tr>
    <tr > 
      <td height='40' colspan='2' align='center'><input name='Action' type='hidden' id='Action' value='DoAdd'> 
        <input class="ACT_btn" type='submit' name='Submit' value=' <% if request("action")="Edit" then response.write "确定修改" Else Response.write "开始生成" %> ' style='cursor:hand;'> 
        &nbsp; <input class="ACT_btn"  name='Cancel' type='button' id='Cancel' value=' 取 消 ' onClick="window.location.href='?'" style='cursor:hand;'></td>
    </tr></form>
  </table>

<%
		End Sub
		
		'开始生成充值卡
		Sub DoAdd()
		 Dim AddType:AddType=request("AddType")
		 Dim CardNum:CardNum=RSQL(request("CardNum"))
		 Dim Password:Password=request("Password")
		 Dim CardList:CardList=request("CardList")
		 Dim Money:Money=ChkNumeric(request("Money"))
		 Dim ValidNum:ValidNum=ChkNumeric(request("ValidNum"))
		 Dim ValidUnit:ValidUnit=request("ValidUnit")
		 Dim EndDate:EndDate=request("EndDate")
		 Dim IsUsed:IsUsed=request("IsUsed")
		 Dim ISSale:IsSale=ChkNumeric(request("IsSale"))
		 Dim title:title=request("title")
		 
		 Dim allgroupid:allgroupid=request("allgroupid")
		 Dim grGroupID:grGroupID=ChkNumeric(request("grGroupID"))
		 Dim ExpireGroupID:ExpireGroupID=ChkNumeric(request("ExpireGroupID"))
	 
		 
 		 IF Money=0 Then Call ACTCMS.Alert("充值卡面值，必须大于0","Card_Act.asp?action=Add"):exit sub
		 IF ValidNum=0 Then Call ACTCMS.Alert("充值卡点数，必须大于0","Card_Act.asp?action=Add"):exit sub
		 If Not IsDate(EndDate) Then Call ACTCMS.Alert("充值截止期限格式不正确!","Card_Act.asp?action=Add"):exit sub
          If AddType=0 or request("action1")="Edit" then
		    if CardNum="" then call ACTCMS.Alert("你没有输入充值卡号!","Card_Act.asp?action=Add"):exit sub
			if PassWord=" "then call ACTCMS.Alert("你没有输入充值卡密码","Card_Act.asp?action=Add"):exit sub
			
			   Dim RS:Set RS=Server.CreateObject("adodb.recordset")
			    if request("action1")="Edit" then
				 rs.open "select * from Card_ACT where id=" & ChkNumeric(request("id")),conn,1,3
				else
					if not conn.execute("select cardnum from Card_ACT where cardnum='" & cardnum & "'").eof then
					  call ACTCMS.Alert("你输入的充值卡号已存在，请重输!","Card_Act.asp?action=Add"):exit sub
					end if
				   rs.open "select * from Card_ACT",conn,1,3
				   rs.addnew
				   rs("AddDate")=now
			   end if
				 rs("cardnum")=CardNum
				 rs("cardpass")=PassWord
				 rs("money")=money
				 rs("ValidNum")=ValidNum
				 rs("ValidUnit")=ValidUnit
				 rs("enddate")=EndDate
				 rs("isused")=isused
				 rs("isSale")=isSale
				 rs("title")=title
				 rs("allgroupid")=allgroupid
				 rs("grGroupID")=grGroupID
				 rs("ExpireGroupID")=ExpireGroupID
 			   rs.update
			   rs.close:set rs=nothing
		  else 
		    if CardList="" then call ACTCMS.Alert("你没有输入充值卡号!","Card_Act.asp?action=Add"):exit sub
			Dim i,j,CardAndPass,CardArr:CardArr=Split(CardList,vbcrlf)
			For I=0 to Ubound(CardArr)
			   CardAndPass=Split(CardArr(I),"|")
			   if not conn.execute("select cardnum from Card_ACT where cardnum='" & CardAndPass(0) & "'").eof then
					 call ACTCMS.Alert("你输入的充值卡号已存在，请重输!","Card_Act.asp?action=Add"):exit sub
			   else
				   Set RS=Server.CreateObject("adodb.recordset")
				   rs.open "select * from Card_ACT",conn,1,3
				   rs.addnew
					 rs("cardnum")=CardAndPass(0)
					 rs("cardpass")=CardAndPass(1)
					 rs("money")=money
					 rs("ValidNum")=ValidNum
					 rs("ValidUnit")=ValidUnit
					 rs("AddDate")=now
					 rs("enddate")=EndDate
					 rs("isused")=isused
					 rs("isSale")=issale
					 rs("title")=title
				 rs("allgroupid")=allgroupid
				 rs("grGroupID")=grGroupID
				 rs("ExpireGroupID")=ExpireGroupID
				   rs.update
				   rs.close:set rs=nothing
			  end if
			Next
		  end if
		  if request("action1")="Edit" then
			   response.write "<script>alert('修改充值卡成功！');location.href='?';</script>"
		  else
			   response.write "<script>alert('添加充值卡成功！');location.href='?';</script>"
		  end if
		End Sub
		'批量生成充值卡操作
		Sub DoAddMore()
		 Dim Nums:Nums=ChkNumeric(request("Nums"))
		 Dim CardNumPrefix:CardNumPrefix=request("CardNumPrefix")
		 Dim CardNumLen:CardNumLen=ChkNumeric(request("CardNumLen"))
		 Dim PasswordLen:PasswordLen=ChkNumeric(request("PasswordLen"))
		 Dim zhtype:zhtype=request("zhtype")
		 Dim Money:Money=ChkNumeric(request("money"))
		 Dim ValidNum:ValidNum=ChkNumeric(request("ValidNum"))
		 Dim ValidUnit:ValidUnit=request("ValidUnit")
		 Dim EndDate:EndDate=request("EndDate")
		 Dim title:title=request("title")
		 Dim allgroupid:allgroupid=request("allgroupid")
		 Dim grGroupID:grGroupID=ChkNumeric(request("grGroupID"))
		 Dim ExpireGroupID:ExpireGroupID=ChkNumeric(request("ExpireGroupID"))
		 
		 IF Nums=0 Then Call ACTCMS.Alert("生成充值卡数量，必须大于0","Card_Act.asp?action=Add"):exit sub
		 IF CardNumLen=0 Then Call ACTCMS.Alert("充值卡号码长度，必须大于0","Card_Act.asp?action=Add"):exit sub
		 IF PasswordLen=0 Then Call ACTCMS.Alert("充值卡密码长度，必须大于0","Card_Act.asp?action=Add"):exit sub
		 IF Money=0 Then Call ACTCMS.Alert("充值卡面值，必须大于0","Card_Act.asp?action=Add"):exit sub
		 IF ValidNum=0 Then Call ACTCMS.Alert("充值卡点数，必须大于0","Card_Act.asp?action=Add"):exit sub
		 If Not IsDate(EndDate) Then Call ACTCMS.Alert("充值截止期限格式不正确!","Card_Act.asp?action=Add"):exit sub
		 %>
		 		   <br>
			  <table width='300'  border='0' align='center' cellpadding='2' cellspacing='1' class='table'>
				<tr  class="bg_tr">
				  <td colspan='2' align='center'><strong>本次生成的点卡信息如下：</strong></td>
				</tr>
				<tr >
				  <td width='100'>充值卡数量：</td>
				  <td><%=nums%> 张</td>
				</tr>
				<tr >
				  <td width='100'>充值卡面值：</td>
				  <td><%=money%> 元</td>
				</tr>
				<tr >
				  <td width='100'>
				  <% select case ValidUnit
					case 1:response.write "充值卡点数："
					case 2:response.write "充值卡有效天数："
					case 3:response.write "充值卡金额："
					end select
					%></td>
				  <td>
				  <% response.write ValidNum
				  select case validunit
				   case 1:response.write " 点"
				   case 2:response.write " 天"
				   case 3:response.write " 元"
				  end select
				  %>
			      </td>
				</tr>
				<tr >
				  <td width='100'>充值截止日期：</td>
				  <td><%=enddate%></td>
				</tr>
				
</table>
			<br>
			<table width='300' border='0' align='center' cellpadding='2' cellspacing='1' class="table">
		  <tr align='center' class="bg_tr">
			<td  width=150 height='22'><strong> 卡 号 </strong></td>
			<td  width=150 height='22'><strong> 密 码 </strong></td>
		  </tr>
		 <%
		 Dim n,currcard,CurrCardPass
		 For N=1 To Nums
		   CurrCard=ACTCMS.MakeRandom(CardNumLen-len(CardNumPrefix))
		   CurrCard=CardNumPrefix & CurrCard
		   If ZhType=2 then
		     CurrCardPass=ACTCMS.GetRandomize(PasswordLen)
		   Else
		     CurrCardPass=ACTCMS.MakeRandom(PasswordLen)
		   End If
		   Do While not Conn.execute("select CardNum From Card_ACT Where CardNum='" & CurrCard & "'").eof 
			   CurrCard=ACTCMS.MakeRandom(CardNumLen-len(CardNumPrefix))
			   CurrCard=CardNumPrefix & CurrCard
		   loop
		   
		   response.write "<tr align='center' >" & vbcrlf
		   response.write "<td height='22'>" & CurrCard & "</td>" & vbcrlf
		   response.write "<td>" & CurrCardPass & "</td>" & vbcrlf
		   response.write "</tr>" & vbcrlf
		   Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		   rs.open "select * from Card_ACT",conn,1,3
		   rs.addnew
		     rs("cardnum")=CurrCard
			 rs("cardpass")=CurrCardPass
			 rs("money")=money
			 rs("ValidNum")=ValidNum
			 rs("ValidUnit")=ValidUnit
			 rs("AddDate")=now
			 rs("enddate")=EndDate
			 rs("isused")=0
			 rs("isSale")=0
			 rs("title")=title
				 rs("allgroupid")=allgroupid
				 rs("grGroupID")=grGroupID
				 rs("ExpireGroupID")=ExpireGroupID
		   rs.update
		   rs.close:set rs=nothing
		   %>
		   
		   <%
		 Next
		 		   response.write "</table>"

		End SUb	
 		  
%> 
<script language="javascript">

function CheckAll(form)
		  {  
		 for (var i=0;i<form.elements.length;i++)  
			{  
			   var e = selform.elements[i];  
			   if (e.name != 'chkAll'&&e.type=="checkbox")  
			   e.checked = selform.chkAll.checked;  
		   }  
	  }function overColor(Obj)
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
function overColor(Obj)
{
	var elements=Obj.childNodes;
	for(var i=0;i<elements.length;i++)
	{
		elements[i].className="tdbg1"
		Obj.bgColor="";
	}
	
}
</script>