<!--#include file="../Conn.asp"-->
<!--#include file="../act_inc/ACT.Main.asp"-->
<!--#include file="label.asp"-->
 <%
	Dim ACTCMS
	Set ACTCMS = New ACT_Main
 
	Function CLreg()

		Dim regEx,Matches,Match,CTemp
		Set regEx = New RegExp
		regEx.Pattern = "<!--ActClass-->([\s\S]*?)<!--ActClass-->"
		regEx.IgnoreCase = True
		regEx.Global = True
		Set Matches = regEx.Execute(TemplateContent)
 		For Each Match In Matches
			 CTemp=ClName(Match.SubMatches(0))
			 TemplateContent =Replace(TemplateContent, Match.Value,CTemp)
		Next
	End Function 
	
	Function ClName(TContents)
			Dim rs,TContent,Content
 			Set rs=actcms.actexe("select ID,ClassName,UModeID from  space_ACT where UModeID="&U&" order by ClassOrder desc,ID desc")
			If Not rs.eof Then 
				Do While Not rs.eof 
					TContent=TContents
 					If InStr(TContent, "{$ClassName}") > 0  Then
					   TContent = Replace(TContent, "{$ClassName}", rs("ClassName"))
					End If
				
					If InStr(TContent, "{$UModeID}") > 0  Then
					   TContent = Replace(TContent, "{$UModeID}", rs("UModeID"))
					End If
				
					If InStr(TContent, "{$ClassTemp}") > 0  Then
					   TContent = Replace(TContent, "{$ClassTemp}", rs("ClassTemp"))
					End If
					
					If InStr(TContent, "{$C}") > 0  Then
					   TContent = Replace(TContent, "{$C}", rs("ID"))
					End If
					Content=Content&TContent
				rs.movenext
				loop
 			Else
				   Content = Replace(TContents, "{$ClassName}", "")
 			End If 
			ClName=Content
	End Function 

 
 %>