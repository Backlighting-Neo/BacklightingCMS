<%
 
		Sub echo(str)
			response.write str
		End Sub 
		Sub die(str)
			  Response.Write str : Response.End
		End Sub 
  
		Function Rep(ByVal ReContent,ByVal ReKey,ByVal ReStr)
			On Error Resume Next
			IF IsNull(ReContent) Or Len(ReContent)=0 Then ReContent=""
			IF IsNull(ReKey) Or Len(ReKey)=0 Then ReKey=""
			IF IsNull(ReStr) Or Len(ReStr)=0 Then ReStr=""
			Rep=Replace(ReContent,ReKey,ReStr)
		End Function

  		Function closeHTML(strContent)
		  CloseHTML=strContent:Exit Function 
 		    Dim arrTags,i,OpenPos,ClosePos,re,strMatchs,j,Match
			Set re=new RegExp
			re.IgnoreCase =True
			re.Global=True
			arrTags=array("strong","em","strike","b","u","i","font","span","a", "h1","h2","h3","h4","h5","h6","p","li","ol","ul","td","tr","tbody","table","blockquote","pre","cite","div")
			  For i=0 To ubound(arrTags)
				   OpenPos=0:ClosePos=0
				   re.Pattern="\<"+arrTags(i)+"( [^\<\>]+|)\>"
				   Set strMatchs=re.Execute(strContent)
				   For Each Match In strMatchs
					OpenPos=OpenPos+1
				   Next
				   re.Pattern="\</"+arrTags(i)+"\>"
				   Set strMatchs=re.Execute(strContent)
				   For Each Match In strMatchs
					ClosePos=ClosePos+1
				   Next
				   For j=1 To OpenPos-ClosePos
					  strContent=strContent+"</"+arrTags(i)+">"
				   Next
			  Next
			  CloseHTML=strContent
 		End Function 

 
 
   		Public Function ChkNumeric(ByVal CheckID)
			If CheckID <> "" And IsNumeric(CheckID) Then
				CheckID = CLng(CheckID)
				If CheckID < 0 Then CheckID = 0
			Else
				CheckID = 0
			End If
			ChkNumeric = CheckID
		End Function
		'过滤非法的SQL字符
		Public Function RSQL(strChar)
			If strChar = "" Or IsNull(strChar) Then RSQL = "":Exit Function
			Dim strBadChar, arrBadChar, tempChar, I
			strBadChar = "$,#,',%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ""
			arrBadChar = Split(strBadChar, ",")
			tempChar = strChar
			For I = 0 To UBound(arrBadChar)
				tempChar = Replace(tempChar, arrBadChar(I), "")
			Next
			RSQL = tempChar
		End Function


		Public Function GetIP() 
			Dim strIPAddr 
			If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" Or InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then 
				strIPAddr = Request.ServerVariables("REMOTE_ADDR") 
			ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then 
				strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",")-1) 
			ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then 
				strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";")-1)
			Else 
				strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
			End If 
			getIP = Replace(Trim(Mid(strIPAddr, 1, 30)), "'", "")
			getIP = Replace(getIP,";","")
			getIP = Replace(getIP,"-","")
			getIP = Replace(getIP,"(","")
			getIP = Replace(getIP,")","")
			getIP = Replace(getIP,">","")
			getIP = Replace(getIP,"<","")
			getIP = Replace(getIP,"=","")
			getIP = Replace(getIP,"*","")
	   End Function
 
       '检查网站目前是否关闭，逆光于2013年10月30日添加
	   Sub CheckWebSiteOnline()
	       If (ShowStaticContent(22)="0") and (not(Session("Backlighting")="debug-website")) Then
		   response.Redirect("/plus/page.asp?ID=1")
		   End If
	   End Sub
%>