<%
	Class ACTM
 
	Public Function ACT_NoRormMXList(ModeID)'读取系统字段输出
	 Dim RSObj
	  Set RSObj=ACTCMS.ACTEXE("Select * from Table_ACT  Where ModeID=" & ModeID & "   and actcms=2 and ISType = 1  and SearchIF=1 order by OrderID desc,ID asc")
	  	Do While Not RSObj.Eof
			ACT_NoRormMXList=ACT_NoRormMXList &"<tr>"&vbCrLf&"<th width=""20%""  ><strong> "&RSObj("Title")&"：</strong><br /></td>"&vbCrLf&"<td>"&ListField(RSObj)&"</td>"&vbCrLf&"</tr>"&vbCrLf&vbCrLf
		RSObj.MoveNext
		Loop
	  RSObj.Close:Set RSObj=Nothing
	 ACT_NoRormMXList=vbCrLf&ACT_NoRormMXList& vbCrLf 
	End function




	
	Public Function ACT_MXEdit(ModeID,ID)'表现方式.输出模型
	 Dim RS
	  Set RS=ACTCMS.ACTEXE("Select * from Table_ACT  Where ModeID=" & ModeID & "  and actcms=2  and ISType = 1    order by OrderID desc,ID asc")
	  	Do While Not RS.Eof
 			ACT_MXEdit=ACT_MXEdit &"<tr>"&vbCrLf&"<th width=""13%"" align=""right""><strong>"&RS("Title")&"：</strong><br /></td>"&vbCrLf&"<td>"&EditField(RS,ModeID,ID)&"</td>"&vbCrLf&"</tr>"&vbCrLf
 		RS.MoveNext
		Loop
	  RS.Close:Set RS=Nothing
	 ACT_MXEdit=vbCrLf&ACT_MXEdit& vbCrLf 
	End function


 	Function ReplaceFormEdit(ModeID,id,TempletContent)
 		Dim MX_Arr,K
		MX_Arr=ACTCMS.Act_MX_Arr(ModeID)
		If IsArray(MX_Arr) Then
		  For K=0 To Ubound(MX_Arr,2)
 			  TempletContent = Replace(TempletContent,"{$" & MX_Arr(0,K) & "}",PrintModeEdit(ModeID,id,MX_Arr(0,K)))
			  TempletContent = Replace(TempletContent,"{$" & MX_Arr(0,K) & "_title}",tip_act(ModeID,MX_Arr(0,K)))
		  Next
		End If
		ReplaceFormEdit=TempletContent
	End Function


	Function PrintModeEdit(ModeID,id,FieldName)

		 Dim RSObj
		  Set RSObj=ACTCMS.ACTEXE("Select * from Table_ACT  Where ModeID=" & ModeID & " and FieldName='"&FieldName&"' and actcms=2 and ISType = 1    order by OrderID asc,ID Desc")
			If Not rsobj.eof Then 
				PrintModeEdit=EditField(RSObj,ModeID,id)

			Else
				PrintModeEdit=""
			End If 
			
		  RSObj.Close:Set RSObj=Nothing
	End Function 
	



	Function PrintMode(ModeID,FieldName)

		 Dim RSObj
		  Set RSObj=ACTCMS.ACTEXE("Select * from Table_ACT  Where ModeID=" & ModeID & " and FieldName='"&FieldName&"'  and actcms=2 and ISType = 1   order by OrderID asc,ID Desc")
			If Not rsobj.eof Then 
				PrintMode=ListField(RSObj)

			Else
				PrintMode=""
			End If 
			
		  RSObj.Close:Set RSObj=Nothing
	End Function 
	
	Function tip_act(ModeID,FieldName)
		  Dim RSObj
		  Set RSObj=ACTCMS.ACTEXE("Select * from Table_ACT  Where ModeID=" & ModeID & " and FieldName='"&FieldName&"' and actcms=2 and ISType = 1   order by OrderID asc,ID Desc")
			If Not rsobj.eof Then 
				If rsobj("IsNotNull")="0" Then 
					tip_act="  <font color=red title=""必填"">*</font>  "&rsobj("Description")
				Else
					tip_act="  "&rsobj("Description")
				End If 
			Else
				tip_act=""
			End If 
	End Function 

	Function ListField(RSObj)
		Dim i,TitleTypeArr,checked,IsNotNull
		Dim arrtitle,arrvalue,titles

		If rsobj("IsNotNull")="0" Then 
			IsNotNull="  <font color=red title=""必填"">*</font>  "&rsobj("Description")
		Else
			IsNotNull="  "&rsobj("Description")
		End If 
 		 Select Case RSObj("FieldType")
		   Case "TextType"
				ListField= "<input  class=""ainput"" type=""text"" title="""&RSObj("Description")&""" name="""&RSObj("FieldName")&""" size="""&RSObj("width")&""" value="""&RSObj("Type_Default")&""">"&IsNotNull
		   Case "MultipleTextType"
				ListField= "<textarea title="""&RSObj("Description")&""" name="""&RSObj("FieldName")&""" style=""height:"&RSObj("height")&"px;width:"&RSObj("width")&"px;"">"&RSObj("Type_Default")&"</textarea>"&IsNotNull
		   Case "MultipleHtmlType"
 				ListField=ListField&"<textarea rows=""10"" cols=""80"" name="""&RSObj("FieldName")&""" id="""&RSObj("FieldName")&"""   style=""width:"&RSObj("width")&"px;height:"&RSObj("height")&"px;visibility:hidden;"" ></textarea>"& vbCrLf
				ListField=ListField&"<script>"& vbCrLf 
				ListField=ListField&"KE.show({"& vbCrLf 
				ListField=ListField&"id : '"&RSObj("FieldName")&"'"& vbCrLf 
				ListField=ListField&"});"& vbCrLf 
  				ListField=ListField&"</script>"&IsNotNull
		   Case "RadioType"
				TitleTypeArr=Split(RSObj("Content"), vbCrLf)
				If RSObj("Type_Type")=0 Then 
				  ListField= ListField&"<select  name="""&RSObj("FieldName")&""">"
				  For I = 0 To UBound(TitleTypeArr)
					arrtitle=Split(TitleTypeArr(I),"-")
					If  UBound(arrtitle)="1" Then 
						titles=arrtitle(0)
						arrvalue=arrtitle(1)
					ElseIf  TitleTypeArr(I)<>"" Then 
					titles=arrtitle(0)
					arrvalue=arrtitle(0)
					Else
						Exit for
					End If 
					If RSObj("Type_Default")=arrvalue Then checked="selected" Else checked=""
					ListField = ListField & "<option value='" & arrvalue & "' "&checked&">" & titles & "</option>"& vbCrLf
				  Next
					ListField= ListField&" </select>"&IsNotNull
				Else
				  For I = 0 To UBound(TitleTypeArr)
				
					arrtitle=Split(TitleTypeArr(I),"-")
					If  UBound(arrtitle)="1" Then 
						titles=arrtitle(0)
						arrvalue=arrtitle(1)
					ElseIf  TitleTypeArr(I)<>"" Then 
					titles=arrtitle(0)
					arrvalue=arrtitle(0)
					Else
						Exit for
					End If 
					
					If RSObj("Type_Default")=arrvalue Then checked="checked" Else checked=""
					ListField = ListField &"<label for="""&RSObj("FieldName")&i&"""> <input  id="""&RSObj("FieldName")&i&""" type=""radio""  name="""&RSObj("FieldName")&""" value="""&arrvalue&""" "&checked&" />"&titles&"&nbsp;&nbsp;</label>" 
				  Next
				    ListField = ListField&IsNotNull
				End If 
		   Case "ListBoxType"
			
				TitleTypeArr=Split(RSObj("Content"), vbCrLf)
				If RSObj("Type_Type")=0 Then 
				  For I = 0 To UBound(TitleTypeArr)
					arrtitle=Split(TitleTypeArr(I),"-")
					If  UBound(arrtitle)="1" Then 
						titles=arrtitle(0)
						arrvalue=arrtitle(1)
					ElseIf  TitleTypeArr(I)<>"" Then 
					titles=arrtitle(0)
					arrvalue=arrtitle(0)
					Else
						Exit for
					End If 
					If RSObj("Type_Default")=arrvalue Then checked="checked" Else checked=""
					ListField = ListField &"<label for="""&RSObj("FieldName")&i&"""> <input  id="""&RSObj("FieldName")&i&""" type=""checkbox""  name="""&RSObj("FieldName")&""" value="""&arrvalue&""" "&checked&" />"&titles&"&nbsp;&nbsp;</label>"& vbCrLf
				  Next
				  ListField = ListField&IsNotNull
				Else
				  ListField= ListField&"<select  size=""4""   style=""width:300px;height:126px""  name="""&RSObj("FieldName")&""" multiple>"
				  For I = 0 To UBound(TitleTypeArr)
					arrtitle=Split(TitleTypeArr(I),"-")
					If  UBound(arrtitle)="1" Then 
						titles=arrtitle(0)
						arrvalue=arrtitle(1)
					ElseIf  TitleTypeArr(I)<>"" Then 
					titles=arrtitle(0)
					arrvalue=arrtitle(0)
					Else
						Exit for
					End If 
					If RSObj("Type_Default")=arrvalue Then checked="checked" Else checked=""
					ListField = ListField & "<option value=""" & arrvalue & """ "&checked&">" & vbCrLf& titles & "</option>"& vbCrLf
				  Next
					ListField= ListField&" </select>"&IsNotNull
				End If 
		   Case "DateType"
				ListField= ListField&"<input name="""&RSObj("FieldName")&""" type=""text"" id="""&RSObj("FieldName")&""" value="""" onClick=WdatePicker();  >"&IsNotNull
		   Case "PicType"
				ListField="未注册用户禁止上传,请注册成功后去会员中心修改资料"
			Case "FileType"
				ListField= "<input  name="""&RSObj("FieldName")&""" type=""text""  value="""" size=""40""><iframe src=""../Upload_Admin.asp?ModeID=1&instr=1&instrname="&RSObj("FieldName")&"&YNContent=1&file=yes&amp;instrct=content"" name=""image"" width='75%' height=""25"" scrolling=""No"" frameborder=""0"" id=""image""></iframe>"&IsNotNull
		   Case "NumberType"
				ListField= "<input type=""text"" name="""&RSObj("FieldName")&""" size="""&RSObj("width")&""" value="""&RSObj("Type_Default")&""">"&IsNotNull
		   Case "RadomType"
				ListField= "<input type=""text"" name="""&RSObj("FieldName")&""" size=""25""  value="""&ACTCMS.MakeRandom(20)&""">"&IsNotNull
		   Case else
				ListField= "<font color=red>该字段错误</font>"
		   End Select 
	End Function 






	Function EditField(RSObj,ModeID,UserID)
		Dim i,IsNotNull,TitleTypeArr,checked,rs1,FieldName
		Dim arrtitle,arrvalue,titles
	  Set RS1=ACTCMS.ACTEXE("Select * from "&ACTCMS.ACT_U(ModeID,2)&"  Where UserID="&UserID&"")
  	  If rs1.eof Then 
		  Actcms.ACTEXE("insert Into "&ACTCMS.ACT_U(ModeID,2)&"(UserID) values("&UserID&")")
		  Set RS1=Actcms.ACTEXE("Select * from "&ACtCMS.ACT_U(ModeID,2)&"  Where UserID="&UserID&"")
	  End if
	  Set RS1=ACTCMS.ACTEXE("Select * from "&ACTCMS.ACT_U(ModeID,2)&"  Where UserID="&UserID&"")
	  FieldName= RSObj("FieldName")
	
		If rsobj("IsNotNull")="0" Then 
			IsNotNull="  <font color=red title=""必填"">*</font>  "&rsobj("Description")
		Else
			IsNotNull="  "&rsobj("Description")
		End If 
 		 Select Case RSObj("FieldType")
		   Case "TextType"
				EditField= "<input  class=""ainput"" title="""&RSObj("Description")&""" type=""text"" name="""&RSObj("FieldName")&""" size="""&RSObj("width")&""" value="""&RS1(FieldName)&""">"&IsNotNull& vbCrLf 
		   Case "MultipleTextType"
				EditField= "<textarea title="""&RSObj("Description")&""" name="""&RSObj("FieldName")&""" style=""height:"&RSObj("height")&"px;width:"&RSObj("width")&"px;"">"&RS1(FieldName)&"</textarea>"&IsNotNull& vbCrLf 
		   Case "MultipleHtmlType"
				EditField=EditField&"<textarea rows=""10"" cols=""80"" name="""&RSObj("FieldName")&""" id="""&RSObj("FieldName")&"""    style=""width:"&RSObj("width")&"px;height:"&RSObj("height")&"px;visibility:hidden;"" >"&RS1(FieldName)&"</textarea>"& vbCrLf
				EditField=EditField&"<script>"& vbCrLf 
				EditField=EditField&"KE.show({"& vbCrLf 
				EditField=EditField&"id : '"&RSObj("FieldName")&"'"& vbCrLf 
				EditField=EditField&"});"& vbCrLf 
  				EditField=EditField&"</script>"&IsNotNull
		   Case "RadioType"
				TitleTypeArr=Split(RSObj("Content"), vbCrLf)
				If RSObj("Type_Type")=0 Then 
				  EditField= EditField&"<select  name="""&RSObj("FieldName")&""">"
				  For I = 0 To UBound(TitleTypeArr)
					arrtitle=Split(TitleTypeArr(I),"-")
					If  UBound(arrtitle)="1" Then 
						titles=arrtitle(0)
						arrvalue=arrtitle(1)
					ElseIf  TitleTypeArr(I)<>"" Then 
					titles=arrtitle(0)
					arrvalue=arrtitle(0)
					Else
						Exit for
					End If 

					If RS1(FieldName)=arrvalue Then checked="selected" Else checked=""
					EditField = EditField & "<option value='" & arrvalue & "' "&checked&">" & titles & "</option>"& vbCrLf
				  Next
					EditField= EditField&" </select>"&IsNotNull& vbCrLf 
				Else
				  For I = 0 To UBound(TitleTypeArr)
					arrtitle=Split(TitleTypeArr(I),"-")
					If  UBound(arrtitle)="1" Then 
						titles=arrtitle(0)
						arrvalue=arrtitle(1)
					ElseIf  TitleTypeArr(I)<>"" Then 
					titles=arrtitle(0)
					arrvalue=arrtitle(0)
					Else
						Exit for
					End If 
					
					
					If RS1(FieldName)=arrvalue Then checked="checked" Else checked=""
					EditField = EditField &"<label for="""&RSObj("FieldName")&i&"""> <input  id="""&RSObj("FieldName")&i&""" type=""radio""  name="""&RSObj("FieldName")&""" value="""&arrvalue&""" "&checked&" />"&titles&"&nbsp;&nbsp;</label>"& vbCrLf
				  Next
				    EditField = EditField&IsNotNull
				End If 
		   Case "ListBoxType"
				TitleTypeArr=Split(RSObj("Content"), vbCrLf)
				If RSObj("Type_Type")=0 Then 
				  For I = 0 To UBound(TitleTypeArr)
				
					arrtitle=Split(TitleTypeArr(I),"-")
					If  UBound(arrtitle)="1" Then 
						titles=arrtitle(0)
						arrvalue=arrtitle(1)
					ElseIf  TitleTypeArr(I)<>"" Then 
					titles=arrtitle(0)
					arrvalue=arrtitle(0)
					Else
						Exit for
					End If 
					If ACTCMS.FoundInArr(RS1(FieldName),arrvalue,",")=true Then checked="checked" Else checked=""
					EditField = EditField &"<label for="""&RSObj("FieldName")&i&"""> <input  id="""&RSObj("FieldName")&i&""" type=""checkbox""  name="""&RSObj("FieldName")&""" value="""&arrvalue&""" "&checked&" />"&titles&"&nbsp;&nbsp;</label>"& vbCrLf
				  Next
				  EditField = EditField&IsNotNull
				Else
				    EditField= EditField&"<select  size=""4""   style=""width:300px;height:126px""  name="""&RSObj("FieldName")&""" multiple>"
				  For I = 0 To UBound(TitleTypeArr)
					arrtitle=Split(TitleTypeArr(I),"-")
					If  UBound(arrtitle)="1" Then 
						titles=arrtitle(0)
						arrvalue=arrtitle(1)
					ElseIf  TitleTypeArr(I)<>"" Then 
					titles=arrtitle(0)
					arrvalue=arrtitle(0)
					Else
						Exit for
					End If 
					If ACTCMS.FoundInArr(RS1(FieldName),arrvalue,",")=true Then checked="selected" Else checked=""
					EditField = EditField & "<option value=""" & arrvalue & """ "&checked&">" & vbCrLf& titles & "</option>"& vbCrLf
				  Next
					EditField= EditField&" </select>"&IsNotNull
				End If 
		   Case "DateType"
				EditField= EditField&"<input  class=""ainput"" name="""&RSObj("FieldName")&""" type=""text"" id="""&RSObj("FieldName")&""" value="""&RS1(FieldName)&""" onClick=WdatePicker(); >"&IsNotNull
		   Case "PicType"
 			 	If RSObj("Type_Type")="0" Then 
					  EditField=	"<input class=""Ainput""  name="""&RSObj("FieldName")&""" type=""text""    value="""&RS1(FieldName)&"""  size=""40""><input name=""button""  onClick=""J('#"&RSObj("id")&"s').dialog({ id:'actcmsscs' ,page: 'Upload_Admin.asp?A=add&instr=2&ModeID="&RSObj("ModeID")&"&instrname="&RSObj("FieldName")&"',  width:720, height:240 });""   id="""&RSObj("id")&"s""   type=""button""  class=""button_style"" style=""cursor:hand;"" value=""点击上传图片"">"
				Else
					EditField="<div id=""sapload"&RSObj("FieldName")&""">"& vbCrLf 
					EditField=EditField&	"</div>"& vbCrLf 
					EditField=EditField& "<script type=""text/javascript"">"& vbCrLf 
					EditField=EditField&"// <![CDATA["& vbCrLf 
					EditField=EditField&"var so = new SWFObject("""&ACTCMS.ACTSYS&"act_inc/sapload.swf"", ""sapload"&RSObj("FieldName")&""", ""450"", ""25"", ""9"", ""#ffffff"");"& vbCrLf 
					EditField=EditField&"so.addVariable('types','"&Replace(ACTCMS.ActCMS_Sys(11),"/",";")&"');"
					EditField=EditField&"so.addVariable('isGet','1');"& vbCrLf 
					EditField=EditField&"so.addVariable('args','myid=Upload;ModeID="&ModeID&";U='+U+"";""+';P='+P+"";""+'Yname="&RSObj("FieldName")&"');"& vbCrLf 
					EditField=EditField&"so.addVariable('upUrl','"&ACTCMS.ACTSYS&"User/Upload.asp');"& vbCrLf 
					EditField=EditField&"so.addVariable('fileName','Filedata');"& vbCrLf 
					EditField=EditField&"so.addVariable('maxNum','110');"& vbCrLf 
					EditField=EditField&"so.addVariable('maxSize','"&ACTCMS.ActCMS_Sys(10)/1024&"');"& vbCrLf 
					EditField=EditField&"so.addVariable('etmsg','1');"& vbCrLf 
					EditField=EditField&"so.addVariable('ltmsg','1');"& vbCrLf 
					EditField=EditField&"so.addParam('wmode','transparent');"& vbCrLf 
					EditField=EditField&"so.write(""sapload"&RSObj("FieldName")&""");"& vbCrLf 
					EditField=EditField&"// ]]>"& vbCrLf 
					EditField=EditField&"</script>"			& vbCrLf 	
					EditField=EditField&"<textarea rows=""10"" cols=""80"" name="""&RSObj("FieldName")&""" id="""&RSObj("FieldName")&"""    style=""width:"&RSObj("width")&"px;height:"&RSObj("height")&"px;visibility:hidden;"" >"&RS1(FieldName)&"</textarea>"& vbCrLf
					EditField=EditField&"<script>"& vbCrLf 
					EditField=EditField&"KE.show({"& vbCrLf 
					EditField=EditField&"id : '"&RSObj("FieldName")&"'"& vbCrLf 
					EditField=EditField&"});"& vbCrLf 
					EditField=EditField&"</script>"&IsNotNull
				End If 
		   
		   Case "FileType"
			'	EditField= "<input  name="""&RSObj("FieldName")&""" type=""text""  value="""&RS1(FieldName)&""" size=""40""><iframe src=""../Upload_Admin.asp?ModeID=1&instr=1&instrname="&RSObj("FieldName")&"&YNContent=1&file=yes&amp;instrct=content"" name=""image"" width='75%' height=""25"" scrolling=""No"" frameborder=""0"" id=""image""></iframe>"&IsNotNull
		   Case "NumberType"
				EditField= "<input  class=""ainput"" type=""text"" name="""&RSObj("FieldName")&""" size="""&RSObj("width")&""" value="""&RS1(FieldName)&""">"&IsNotNull
		   Case "RadomType"
				EditField= "<input  class=""ainput"" type=""text"" name="""&RSObj("FieldName")&""" size=""25""  value="""&RS1(FieldName)&""">"&IsNotNull
		   Case else
				EditField= "<font color=red>该字段错误</font>"
		   End Select 
	  RS1.Close:Set RS1=Nothing
	End Function 







 

End Class 
	Dim U_M
	Set U_M = New ACTM
%>