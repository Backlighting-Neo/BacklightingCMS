<!--#include file="../ACT.Function.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>标签拷贝</title>
<link href="../Images/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
	Dim A,ID
	A = Request("A")
	ID = ChkNumeric(ACTCMS.S("ID"))
	Select Case A
		Case "U"
			Call UpLabel()
		Case "C"
		   Call Main()
		Case Else
		  Call main()
	End Select 


	Sub UpLabel()
		Dim Rs,LabelName,LabelRS
		LabelName=RSQL(ACTCMS.S("LabelName"))
		If LabelName="" Then Call ACTCMS.Alert("标签名称不能为空",""):Exit Sub
		LabelName="{ACTCMS_" & LabelName & "}"
		Set Rs = Actcms.Actexe("Select * From Label_Act Where ID=" & ID & "")
		If Not rs.eof Then 
			Set LabelRS = Server.CreateObject("Adodb.RecordSet")
			LabelRS.Open "Select LabelName From Label_Act Where LabelName='" & LabelName & "'", Conn, 1, 1
			If Not LabelRS.EOF Then
			  Call ACTCMS.Alert("标签名称已经存在!",""):Exit Sub
			  Response.End
			Else
				LabelRS.Close
				LabelRS.Open "Select * From Label_Act", Conn, 1, 3
				LabelRS.AddNew
				LabelRS("LabelName") = LabelName
				LabelRS("Description") = RS("Description")
				LabelRS("LabelContent") = RS("LabelContent")
				LabelRS("LabelFlag") = RS("LabelFlag")
				LabelRS("AddDate") = Now
				LabelRS("LabelType") = 1 '指定为系统函数标签
				LabelRS.Update
				LabelRS.Close
			    Call ACTCMS.Alert("标签复制成功!","../Label_Admin.asp?Type="&request("Type")&""):Exit Sub
		   End If 
		End If 

	End Sub 
	

	Sub main() 
	Dim N
	N = Request("N")
	N = "复制"&Replace(Replace(N, "{ACTCMS_", ""), "}", "")
 %>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
 <form name="form1" method="post" action="?A=U&ID=<%=ID%>&Type=<%=request("Type")%>"> <tr>
    <td colspan="2" class="bg_tr">您现在的位置：后台管理中心 >> 标签库</td>
  </tr>
  <tr>
    <td align="right">新标签名称：</td>
    <td><input name="LabelName" size="40" type="text" id="LabelName" value="<%=N%>"></td>
  </tr>
  <tr>
    <td colspan="2" align="center"><input name="Submit" type="submit" class="ACT_btn" value="  提  交  ">
      &nbsp;&nbsp;&nbsp;
      <input name="Submit2" type="reset" class="ACT_btn" value="  重  置  "></td>
  </tr> </form>
</table>
<%End Sub %>

</body></html>