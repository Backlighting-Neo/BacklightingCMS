<!--#include file="../ACT_INC/ACT.User.asp"-->
<%
ConnectionDatabase
 Dim ID,Hits,SqlStr,RS,ModeID
ModeID = ChkNumeric(Request("ModeID"))
If ModeID =0 Then response.End  
ID = ChkNumeric(Request("ID"))
 If ID = "" Then
	Hits = 0
 Else
	SqlStr = "SELECT Hits From "&ACTCMS.ACT_C(ModeID,2)&" Where ID=" & ID &""
	Set RS = Server.CreateObject("ADODB.Recordset")
	RS.Open SqlStr, conn, 1, 3
	If RS.bof And RS.EOF Then
		Hits = 0
	Else
		If request("A")="List" Then 
			Hits = rs(0)
		Else
 			Hits = rs(0) + 1
			rs(0) = Hits
			rs.Update
		End If 
	End If
	rs.Close
	Set rs = Nothing
 End IF
Response.Write "document.write('" & Hits & "');"


%>