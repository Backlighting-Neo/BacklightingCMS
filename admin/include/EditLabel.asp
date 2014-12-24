<!--#include file="../ACT.Function.asp"-->
<% ConnectionDatabase
Dim ID,LabelRS,Str,LabelContent,FileName
ID=clng(Request.QueryString("ID"))
Set LabelRS=Server.CreateObject("Adodb.Recordset")
 Str="SELECT LabelContent FROM  Label_Act Where ID=" & ID &""
 LabelRS.Open Str,Conn,1,1
IF LabelRS.Eof and LabelRS.Bof THEN
 LabelRS.Close
 Set LabelRS=Nothing
 Response.Write("<Script>alert('参数传递出错!');window.close();</Script>")
 Response.End
End if
 LabelContent=LabelRS(0)
 LabelRS.Close
  FileName=Replace(Split(Split(LabelContent,"§")(0),"(")(0),"{$","")
 FileName=FileName & ".asp?Action=Edit&ID=" & ID
 Set LabelRS=Nothing
 response.Redirect "Label/" &FILENAME&""
%>?