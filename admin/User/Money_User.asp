<!--#include file="../ACT.Function.asp"-->
<!--#include file="../../ACT_inc/cls_pageview.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>1</title>
<link href="../Images/style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.STYLE1 {font-weight: bold}
-->
</style>
</head>
<body>
<%ConnectionDatabase	
		Private totalPut,rs, CurrentPage, MaxPerPage,DomainStr,SearchType,SQLParam
		SearchType=ChkNumeric(request("SearchType"))
		%>
  <table class=table style="margin-top:2px" cellSpacing=1 cellPadding=2 width="98%" align=center border=0>
    <tr class=title>
      <td height=22 colSpan=10 class="bg_tr">
       <B>资 金 明 细 查 询</B>      </td>
    </tr>
    <tr class=tdbg height=30>
<FORM name=form1 action=? method=get>
      <td>快速查找： 
<Select onchange=javascript:submit() size=1 name=SearchType class='textbox'> 
  <Option value=0<%If SearchType="0" Then Response.write " selected"%>>所有资金明细记录</Option> 
  <Option value=1<%If SearchType="1" Then Response.write " selected"%>>最近10天内的新资金明细记录</Option> 
  <Option value=2<%If SearchType="2" Then Response.write " selected"%>>最近一月内的新资金明细记录</Option> 
  <Option value=3<%If SearchType="3" Then Response.write " selected"%>>所有收入记录</Option> 
  <Option value=4<%If SearchType="4" Then Response.write " selected"%>>所有支出记录</Option>
      </Select>&nbsp;&nbsp;&nbsp;&nbsp;<a href="?">资金明细首页</a></td></FORM>
<FORM name=form2 action=? method=post>
      <td>高级查询： 
<Select id=Field name=Field class='textbox'> 
  <Option value=1 selected>客户姓名</Option> 
  <Option value=2>用户名</Option> 
  <Option value=3>交易时间</Option> 
</Select> 
  <Input id=Keyword class='Ainput' maxLength=30 name=Keyword> 
  <Input class="ACT_btn" type=submit value=" 查 询 " name=Submit2> 
        <Input id=SearchType type=hidden value=5 name=SearchType> </td></FORM>
    </tr>
</table>
  <table width="98%" align="center" cellpadding="2" cellspacing="1" class="table">
    <tr>
      <td align=left class="bg_tr">您现在的位置：<a href="?">资金明细记录管理</a>&nbsp;&gt;&gt;&nbsp;
	  <%Dim SearchTypeStr
	    Dim KeyWord:KeyWord=request("KeyWord")
	  Select Case SearchType
	     Case 0 :SearchTypeStr="所有资金明细记录"
		 Case 1 :SearchTypeStr="最近10天内的新资金明细记录"
		 Case 2 :SearchTypeStr="最近一月内的新资金明细记录"
		 Case 3 :SearchTypeStr="所有收入记录"
		 Case 4 :SearchTypeStr="所有支出记录"
		 Case 5 
		    Select Case ChkNumeric(request("Field"))
			  Case 1:SearchTypeStr="客户姓名含有<font color=red>""" & KeyWord & """</font>"
			  Case 2:SearchTypeStr="用户名含有<font color=red>""" & KeyWord & """</font>"
			  Case 3:SearchTypeStr="交易时间含有<font color=red>""" & KeyWord & """</font>"
			End Select
	  End Select
	  Response.Write SearchTypeStr%></td>
    </tr>
</table>
  <table cellpadding="2" cellspacing="1" width="98%" border=0 class="table"  align="center">
    <tr  align=middle>
      <td class="bg_tr" width=150 height="25" >交易时间</td>
      <td class="bg_tr" width=80>用户名</td>
      <td class="bg_tr" width=80>客户姓名</td>
      <td class="bg_tr" width=60>交易方式</td>
      <td class="bg_tr" width=50>币种</td>
      <td class="bg_tr" width=80>收入金额</td>
      <td class="bg_tr" width=80>支出金额</td>
      <td class="bg_tr" width=40>摘要</td>
      <td class="bg_tr">备注/说明</td>
    </tr>
	<%
			MaxPerPage=20
			If request("page") <> "" Then
				  CurrentPage = ChkNumeric(request("page"))
			Else
				  CurrentPage = 1
			End If
			SqlParam="1=1"
            If SearchType<>"0" Then
			  Select Case SearchType
			   Case 1
					SqlParam=SqlParam &" And datediff('d',Logtime," & NowString & ")<=10"
			   Case 2
					SqlParam=SqlParam &" And datediff('d',Logtime," & NowString & ")<=30"
			  Case 3 : SqlParam = SqlParam & "And IncomeFlag=1"
			  Case 4 : SqlParam = SqlParam & "And IncomeFlag=2"
			  Case 5
			      Select Case ChkNumeric(request("Field"))
				   Case 1
				     SqlParam=SqlParam &" And ClientName Like '%" & Keyword & "%'"
				   Case 2
				     SqlParam=SqlParam &" And UserID Like '%" & Keyword & "%'"
				   Case 3
				     SqlParam=SqlParam &" And logtime Like '%" & Keyword & "%'"
				  End Select
			  End Select
			End If
	Set RS=Server.CreateObject("ADODB.RECORDSET")
 	
	RS.Open "Select * From Money_Log_ACT Where " & SqlParam & " Order By ID Desc",Conn,1,1
	If RS.Eof AND RS.Bof Then
	 Response.WRITE "<tr  lass='td_bg' ><td colspan=9 align=center height='25'>找不到" & SearchTypeStr & "!</td></tr>"
   Else
                          totalPut = RS.RecordCount
							If CurrentPage < 1 Then	CurrentPage = 1
							If (CurrentPage - 1) * MaxPerPage > totalPut Then
								If (totalPut Mod MaxPerPage) = 0 Then
									CurrentPage = totalPut \ MaxPerPage
								Else
									CurrentPage = totalPut \ MaxPerPage + 1
								End If
							End If
							If CurrentPage = 1 Then
								Call showContent()
							Else
								If (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
									Call showContent()
								Else
									CurrentPage = 1
									Call showContent()
								End If
							End If
   End If
   RS.Close:Set RS=Nothing
   %>

  <div align="center">
      <%
		   	  '显示分页信息
			  Call ShowPagePara(totalPut, MaxPerPage, "", True, "条记录", CurrentPage, "SearchType=" & SearchType & "&Field=" & request("Field") & "&KeyWord=" & KeyWord)
		
		   %>
     </div>
         <br>
   <table border="0" width="98%" align="center">
    <tr>
	  <td>
     <font color=red>说明：为免引起不必要的纠纷，资金明细仅提供查询功能，不能删除操作！</font>
      </td>
	</tr>
</table>
</body>
<html>
   <%
  
	Public Function ShowPagePara(totalnumber, MaxPerPage, FileName, ShowAllPages, strUnit, CurrentPage, ParamterStr)
		  Dim N, I, PageStr
				Const Btn_First = "第一页"'样式定义
				Const Btn_Prev = "上一页" 
				Const Btn_Next = "下一页" 
				Const Btn_Last = "最后一页" 
				  PageStr = ""
					If totalnumber Mod MaxPerPage = 0 Then
						N = totalnumber \ MaxPerPage
					Else
						N = totalnumber \ MaxPerPage + 1
					End If
				If N > 1 Then
					PageStr = PageStr & ("页次：<font color=red>" & CurrentPage & "</font>/" & N & "页 共有:" & totalnumber & strUnit & " 每页:" & MaxPerPage & strUnit & " ")
					If CurrentPage < 2 Then
						PageStr = PageStr & Btn_First & " " & Btn_Prev & " "
					Else
						PageStr = PageStr & ("<a href=" & FileName & "?page=1" & "&" & ParamterStr & ">" & Btn_First & "</a> <a href=" & FileName & "?page=" & CurrentPage - 1 & "&" & ParamterStr & ">" & Btn_Prev & "</a> ")
					End If
					
					If N - CurrentPage < 1 Then
						PageStr = PageStr & " " & Btn_Next & " " & Btn_Last & " "
					Else
						PageStr = PageStr & (" <a href=" & FileName & "?page=" & (CurrentPage + 1) & "&" & ParamterStr & ">" & Btn_Next & "</a> <a href=" & FileName & "?page=" & N & "&" & ParamterStr & ">" & Btn_Last & "</a> ")
					End If
					If ShowAllPages = True Then
						PageStr = PageStr & ("GO:<select  onChange='location.href=this.value;' style='width:55;' name='select'>")
				   For I = 1 To N
					 If Cint(CurrentPage) = I Then
						PageStr = PageStr & ("<option value=" & FileName & "?page=" & I & "&" & ParamterStr & " selected>NO." & I & "</option>")
					 Else
						 PageStr = PageStr & ("<option value=" & FileName & "?page=" & I & "&" & ParamterStr & ">NO." & I & "</option>")
					 End If
				   Next
				  PageStr = PageStr & "</select>"
				  End If
			 End If
			 ShowPagePara = PageStr
			 response.write ShowPagePara
	End Function
  Sub ShowContent()
     Dim I,intotalmoney,outtotalmoney
     Do While Not rs.eof 
	%>
    <tr class=tdbg  onMouseOver=overColor(this) onMouseOut=outColor(this)>
      <td align=middle width=150 class="td_bg"><%=rs("LogTime")%></td>
      <td align=middle width=80 class="td_bg"><%=ACTCMS.UserM(rs("UserID"))%></td>
	  <td align=middle width=80 class="td_bg"><%=rs("clientname")%></td>
      <td align=middle width=60 class="td_bg">
	  <% Select Case rs("MoneyType")
	      Case 1:Response.WRite "现金"
		  Case 2:Response.Write "银行汇款"
		  Case 3:Response.Write "在线支付"
		  Case 4:Response.Write "资金余额"
		 End Select
	 %>
	  </td>
      <td align=middle width=50 class="td_bg">人民币</td>
      <td width=80 align=right class="td_bg"> 
	  <%If rs("IncomeFlag")=1 Then
	     Response.Write formatnumber(rs("money"),2)
		 intotalmoney=intotalmoney+rs("money")
	    End If
		%></td>
      <td align=right width=80 class="td_bg">
	  <%If rs("IncomeFlag")=2 Then
	     Response.Write formatnumber(rs("money"),2)
		 outtotalmoney=outtotalmoney+rs("money")
	    End If
		%></td>
      <td align=center width=40 class="td_bg">
	  <% If rs("IncomeFlag")=1 Then
	      Response.Write "<font color=red>收入</font>"
		 Else
		  Response.Write "<font color=green>支出</font>"
		 End If
		 %></td>
      <td align=middle class="td_bg"><%=rs("Remark")%></td>
    </tr>
	<%
	            
				I = I + 1
				RS.MoveNext
				If I >= MaxPerPage Then Exit Do

	 loop
	%>
    <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
      <td align=right colSpan=5 class="td_bg">本页合计：</td>
      <td align=right class="td_bg"><%=formatnumber(intotalmoney,2)%></td>
      <td align=right class="td_bg"><%=formatnumber(outtotalmoney,2)%></td>
      <td colSpan=3 class="td_bg">&nbsp;</td>
    </tr>
    <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
      <td align=right colSpan=5 class="td_bg">总计金额：</td>
	  <%intotalmoney=Conn.execute("Select Sum(Money) From Money_Log_ACT Where "& SqlParam & " And IncomeFlag=1")(0)
	    outtotalmoney=Conn.execute("Select Sum(Money) From Money_Log_ACT Where "& SqlParam & " And IncomeFlag=2")(0)
	    if not isnumeric(intotalmoney) then intotalmoney=0
		if not isnumeric(outtotalmoney) then outtotalmoney=0
	  %>
      <td align=right class="td_bg"><%=formatnumber(intotalmoney,2)%></td>
      <td align=right class="td_bg"><%=formatnumber(outtotalmoney,2)%></td>
      <td align=middle colSpan=3 class="td_bg">资金余额：<%=formatnumber(intotalmoney-outtotalmoney,2)%></td>
    </tr>
  </table>
		<%
		End Sub

%> <script language="javascript">

//CSS背景控制
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
