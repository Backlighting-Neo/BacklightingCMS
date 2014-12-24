	Function newlogin(UModeID)
		Dim rs,myface
		Set rs=actcms.actexe("Select top 6  * from User_act   Order By LoginTime desc")
		If Not rs.eof Then 
		Do While Not rs.eof 
 			If rs("myface")<>"" Then 
				myface=rs("myface")
			Else 
				myface=ACTCMS.ActSys&"user/images/nophoto.gif" 
			End If 
 				 newlogin=newlogin&"<li><a href=""space/?"&ACTCMS.ACT_U(UmodeID,5)&"-"&rs("userid")&"""    target=""_blank""><img src="""&myface&""" alt='"&rs("UserName")&"' width=""52"" height=""52"" />"&rs("UserName")&"</a></li>"
  		rs.movenext
		loop
	  End If 
 	  rs.close:set rs=Nothing
	 actcool=newlogin   '注释 独立函数一定要将结果赋值给actcool
 	End Function 