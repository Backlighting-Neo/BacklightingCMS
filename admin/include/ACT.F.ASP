<%
	Class F
		Public Function ACT_DateStr(DateForm)
			 IF DateForm = 0 Then
		 		Response.Write"<option value=""0"" Selected>-不显示日期-</option>"
			 Else
		 		Response.Write "<option value=""0"">-不显示日期-</option> "
			 End IF
			  If CStr(DateForm) = "1" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""1""" & Str & " >2005-10-1</option>")
			  If CStr(DateForm) = "2" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""2""" & Str & ">2005.10.1</option>")
			  If CStr(DateForm) = "3" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""3""" & Str & ">2005/10/1</option>")
			  If CStr(DateForm) = "4" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""4""" & Str & ">10/1/2005</option>")
			  If CStr(DateForm) = "5" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""5""" & Str & ">2005年10月</option>")
			  If CStr(DateForm) = "6" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""6""" & Str & ">2005年10月1日</option>")
			  If CStr(DateForm) = "7" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""7""" & Str & ">10.1.2005</option>")
			  If CStr(DateForm) = "8" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""8""" & Str & ">10-1-2005</option>")
			  If CStr(DateForm) = "9" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""9""" & Str & ">10/1</option>")
			  If CStr(DateForm) = "10" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""10""" & Str & ">10.1</option>")
			  If CStr(DateForm) = "11" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""11""" & Str & ">10月1日</option>")
			  If CStr(DateForm) = "12" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""12""" & Str & ">1日12时</option>")
			  If CStr(DateForm) = "13" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""13""" & Str & ">1日12点</option>")
			  If CStr(DateForm) = "14" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""14""" & Str & ">12时12分</option>")
			  If CStr(DateForm) = "15" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""15""" & Str & ">12:12</option>")
			  If CStr(DateForm) = "16" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""16""" & Str & ">10-1</option>")
			   If CStr(DateForm) = "17" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""17""" & Str & ">10/1 12:00</option>")
			  
			  ACT_DateStr = ACT_DateStr & ("<optgroup  label=""-----加括号格式-----""></optgroup>")

			  If CStr(DateForm) = "21" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""21""" & Str & " >(2005-10-1)</option>") 
			  If CStr(DateForm) = "22" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""22""" & Str & ">(2005.10.1)</option>")
			  If CStr(DateForm) = "23" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""23""" & Str & ">(2005/10/1)</option>")
			  If CStr(DateForm) = "24" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""24""" & Str & ">(10/1/2005)</option>")
			  If CStr(DateForm) = "25" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""25""" & Str & ">(2005年10月)</option>")
			  If CStr(DateForm) = "26" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""26""" & Str & ">(2005年10月1日)</option>")
			  If CStr(DateForm) = "27" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""27""" & Str & ">(10.1.2005)</option>")
			  If CStr(DateForm) = "28" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""28""" & Str & ">(10-1-2005)</option>")
			  If CStr(DateForm) = "29" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""29""" & Str & ">(10/1)</option>")
			  If CStr(DateForm) = "30" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""30""" & Str & ">(10.1)</option>")
			  If CStr(DateForm) = "31" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""31""" & Str & ">(10月1日)</option>")
			  If CStr(DateForm) = "32" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""32""" & Str & ">(1日12时)</option>")
			  If CStr(DateForm) = "33" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""33""" & Str & ">(1日12点)</option>")
			  If CStr(DateForm) = "34" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""34""" & Str & ">(12时12分)</option>")
			  If CStr(DateForm) = "35" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""35""" & Str & ">(12:12)</option>")
			  If CStr(DateForm) = "36" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""36""" & Str & ">(10-1)</option>")
			  If CStr(DateForm) = "37" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""37""" & Str & ">(10/1 12:00)</option>")


			  ACT_DateStr = ACT_DateStr & ("<optgroup  label=""-----加中括号格式-----""></optgroup>")
			  If CStr(DateForm) = "41" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""41""" & Str & ">[2005-10-1]</option>")
			  If CStr(DateForm) = "42" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""42""" & Str & ">[2005.10.1]</option>")
			  If CStr(DateForm) = "43" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""43""" & Str & ">[2005/10/1]</option>")
			  If CStr(DateForm) = "44" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""44""" & Str & ">[10/1/2005]</option>")
			  If CStr(DateForm) = "45" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""45""" & Str & ">[2005年10月]</option>")
			  If CStr(DateForm) = "46" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""46""" & Str & ">[2005年10月1日]</option>")
			  If CStr(DateForm) = "47" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""47""" & Str & ">[10.1.2005]</option>")
			  If CStr(DateForm) = "48" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""48""" & Str & ">[10-1-2005]</option>")
			  If CStr(DateForm) = "49" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""49""" & Str & ">[10/1]</option>")
			  If CStr(DateForm) = "50" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""50""" & Str & ">[10.1]</option>")
			  If CStr(DateForm) = "51" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""51""" & Str & ">[10月1日]</option>")
			  If CStr(DateForm) = "52" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""52""" & Str & ">[1日12时]</option>")
			  If CStr(DateForm) = "53" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""53""" & Str & ">[1日12点]</option>")
			  If CStr(DateForm) = "54" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""54""" & Str & ">[12时12分]</option>")
			  If CStr(DateForm) = "55" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""55""" & Str & ">[12:12]</option>")
			  If CStr(DateForm) = "56" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""56""" & Str & ">[10-1]</option>")
			  If CStr(DateForm) = "57" Then Str = " Selected" Else Str = ""
			   ACT_DateStr = ACT_DateStr & ("<option value=""57""" & Str & ">[10/1 12:00]</option>")
		End Function
	


		Public Function ACT_L_Mode(Selected)
			
		Dim MX_Sys,i
			MX_Sys=ACTCMS.Act_MX_Sys_Arr()
			If IsArray(MX_Sys) Then
				For I=0 To Ubound(MX_Sys,2)
		  
			   IF Selected=MX_Sys(0,I) Then
				ACT_L_Mode=ACT_L_Mode & "<option value=""" & MX_Sys(0,I) & """ Selected>" & MX_Sys(1,I)&"模型" & "</option>"& vbCrLf
			   Else
				ACT_L_Mode=ACT_L_Mode & "<option value=""" & MX_Sys(0,I) & """>" & MX_Sys(1,I)&"模型" & "</option>"& vbCrLf
			   End If
			
			Next
			End If
		End Function	


		Public Function ACT_URL_Mode(Selected,URL)
			
		Dim MX_Sys,i
			MX_Sys=ACTCMS.Act_MX_Sys_Arr()
			If IsArray(MX_Sys) Then
				For I=0 To Ubound(MX_Sys,2)
		  
			   IF Selected=MX_Sys(0,I) Then
				ACT_URL_Mode=ACT_URL_Mode & "<option value=""?ModeID=" &  MX_Sys(0,I)&URL & """ Selected>" & MX_Sys(1,I)&"模型" & "</option>"& vbCrLf
			   Else
				ACT_URL_Mode=ACT_URL_Mode & "<option value=""?ModeID=" &  MX_Sys(0,I)&URL & """>" & MX_Sys(1,I)&"模型" & "</option>"& vbCrLf
			   End If
			
			Next
			End If
		End Function	

		Public Function ActCMS_DIY_F(ModeID,W,DIY_Article)
		Dim ActCMS__SQL1,ActCMS__SQL2
		Set ActCMS__SQL1 = server.CreateObject("adodb.recordset")
			ActCMS__SQL2= "Select ACT_DiY from Mode_Act where ModeID="& ModeID
			ActCMS__SQL1.Open ActCMS__SQL2,Conn,1,3
			If W=1  Then 
				If Trim(ActCMS__SQL1("ACT_DiY"))<>"" Then
					ActCMS_DIY_F=ActCMS__SQL1("ACT_DiY"):Exit Function 
				Else 
					ActCMS_DIY_F="1":Exit Function 
				End If 
			End If 
			ActCMS__SQL1("ACT_DiY")=DIY_Article
			ActCMS__SQL1.update
		End Function


	Public Function GroupOption(Selected)
	 Dim RSObj:Set RSObj=Server.CreateObject("Adodb.Recordset")
	    RSObj.Open "Select GroupID,GroupName From Group_ACT",Conn,1,1
	  	Do While Not RSObj.Eof
 		   IF Selected=RSObj(0) Then
			GroupOption=GroupOption & "<option value=""" & RSObj(0) & """ Selected>" &RSObj(1) & "</option>"
		   Else
			GroupOption=GroupOption & "<option value=""" & RSObj(0) & """>" & RSObj(1) & "</option>"
		   End If
		RSObj.MoveNext
		Loop
	  RSObj.Close:Set RSObj=Nothing
	End Function	



	Public Function ACT_LabelFolder(Selected)
	 Dim RSObj
	    Set RSObj = ACTCMS.ACTExe("Select ID,Foldername From ACT_LabelFolder")
	  	Do While Not RSObj.Eof
		   IF Selected=RSObj(0) Then
			ACT_LabelFolder=ACT_LabelFolder & "<option value=""" & RSObj(0) & """ Selected>" & RSObj(1) & "</option>"& vbCrLf
		   Else
			ACT_LabelFolder=ACT_LabelFolder & "<option value=""" & RSObj(0) & """>" & RSObj(1) & "</option>"& vbCrLf
		   End If
		RSObj.MoveNext
		Loop
	  RSObj.Close:Set RSObj=Nothing
	End Function	


	Function  ForClasslist(ModeID)
		 Dim Crs,selected,ModeSql
		' If Modeid<>0 Then ModeSql="And ModeID="&ModeID
		 Set Crs = actcms.actexe("Select * from Class_act where ParentID='0' "&ModeSql&"  and ActLink<>2 Order by Orderid desc,ID desc")
		 If Not  Crs.Eof Then
		 do while Not Crs.Eof
				If ClassID=Crs("ClassID") Then selected=" selected=""selected""" Else selected=""
				If   ACTCMS.ACTCMS_QXYZ(ModeID,"1",Crs("ClassID")) Then  
					If Crs("ActLink")=3 Then 
					 ForClasslist=ForClasslist&"<option disabled=""value""  value='"&Crs("ClassID")&"' "&selected&">"& Crs("ClassName") & "</option>"& vbCrLf
					Else 
					 ForClasslist=ForClasslist&"<option  value='"&Crs("ClassID")&"' "&selected&">"& Crs("ClassName") & "</option>"& vbCrLf
					End If 
				Else
				 ForClasslist=ForClasslist&"<option value='' >无权限</option>"& vbCrLf
				End If 
				 ForClasslist=ForClasslist&(GetForClass(Crs("ClassID"),""))
			 Crs.MoveNext
		 Loop
		 End IF
	 End  Function 
	 Function GetForClass(CID,Str)
	       Dim Sql,RSC,TempImageStr,ImageStr,CheckStr,selected
	        TempImageStr = "&nbsp;└"
	        Sql = "Select * from Class_act where ParentID='" & CID & "'  and ActLink<>2"
	        Set RSC = Conn.Execute(Sql)
			ImageStr = Str & "&nbsp;└"
	            do while Not RSC.Eof
					If ClassID=RSC("ClassID") Then selected=" selected=""selected""" Else selected=""
					  If RSC("ModeID")= ModeID   Then 
						If   ACTCMS.ACTCMS_QXYZ(ModeID,"1",RSC("ClassID")) Then  
							If RSC("ActLink")=3 Then 
								GetForClass = GetForClass  & "<option disabled=""value"" value='"&RSC("ClassID")&"' "&selected&">"& ImageStr & TempImageStr &" "& RSC("ClassName")& "</option>"& vbCrLf
							Else 
								GetForClass = GetForClass  & "<option value='"&RSC("ClassID")&"' "&selected&">"& ImageStr & TempImageStr &" "& RSC("ClassName")& "</option>"& vbCrLf
							End If 
 						End If 
					  End If 
					   GetForClass = GetForClass & GetForClass(RSC("ClassID"),ImageStr)
					   RSC.MoveNext
	           loop
	       Set RSC = Nothing
	 End Function


End Class 
	Dim AF
	Set AF = New F
%>