<!--#include file="../ACT.Function.asp"-->
<%
 	Dim A:A=request("A")
	Select Case A
		Case "testsource"
			Call testsource()
 	End Select 


	  Sub testsource()
	  on error resume next
	   dim str:str=request("str")
	   If actcms.G("DataType")="1" or actcms.G("DataType")="5" or actcms.G("DataType")="6"  Then str=actcms.GetAbsolutePath(str)
	   dim tconn:Set tconn = Server.CreateObject("ADODB.Connection")
		tconn.open str
		If Err Then 
 		  Set tconn = Nothing
		  Response.Write "0|"&Err.Description
		else
		  Response.Write "1|"&Err.Description
		end if
	  end sub

%>