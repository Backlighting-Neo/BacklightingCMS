<%
	Set Fso = Server.CreateObject("scripting.FileSystemObject")
	If Fso.FileExists(Server.MapPath("../ACT_inc/lock/Install.lock")) Then Response.Write "err" : Response.End
	Call SaveFile()
	Function SaveFile()
		Dim SaveRemoteFile:SaveRemoteFile=True
		dim Ads,Retrieval,GetRemoteData
		Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
		With Retrieval
			.Open "Get", "http://www.actcms.com/data/data.txt", False, "", ""
			.Send
			If .Readystate<>4 then
				SaveRemoteFile=False
				Exit Function
			End If
		GetRemoteData = .ResponseBody
		End With
		Set Retrieval = Nothing
		Set Ads = Server.CreateObject("Adodb.Stream")
		With Ads
			.Type = 1
			.Open
			.Write GetRemoteData
			.SaveToFile server.MapPath("data.txt"),2
			.Cancel()
		If  .size >1000 Then 
			response.write "OK"
		Else 
			response.write "err"
		End If 
		.Close()
		End With
		Set Ads=nothing
	
	End Function
	'

%>