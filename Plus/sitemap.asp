<!--#include file="../act_inc/ACT.User.asp"-->
<%	


	Dim TemplateContent,ACT_L,rs,C
    Set ACT_L = New ACT_Code
	TemplateContent = ACT_L.LoadTemplate("plus/sitemap.html")
	
	Set rs=actcms.actexe("Select ClassID,classname from Class_Act where Parentid  = '0' and actlink=1 Order by Orderid asc,ID asc")
	If Not rs.eof Then 
		Do While Not rs.eof
				C=C&"<div class=""linkbox""><h3><a target='_blank' href='"&ACTCMS.DiyClassName(rs("classid"))&"'>"&rs("classname")&"</a></h3><ul class=""f6"">	"& vbCrLf
				C=C&CL(actcms.TempClassID(rs("classid")))&"</ul></div>"& vbCrLf
 		rs.movenext
		loop
 	End If 

	Function cl(classid)
		Dim cid,i
		classid=Replace(ClassID,"'","")
		cid=Split(classid,",")
		For I = 0 To UBound(cid)
 			CL=cl&"   <li><a href='"&ACTCMS.DiyClassName(cid(I))&"'>"&ACTCMS.ACT_L(cid(I),2)&"</a></li>"& vbCrLf
		Next 
 	End Function 
 	If InStr(TemplateContent, "{$maplist}") > 0  Then
	   TemplateContent = Replace(TemplateContent, "{$maplist}", c)
	End if
 	 TemplateContent = ACT_L.LabelReplaceAll(TemplateContent)
 	 response.write TemplateContent
	Call CloseConn

%>