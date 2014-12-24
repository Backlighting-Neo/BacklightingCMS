<%'前后台公用调用
	Class ACTM
 
	Public Function ACT_MXEdit(ModeID,ID)'表现方式.输出模型
	 Dim RS
	  Set RS=ACTCMS.ACTEXE("Select * from Table_ACT  Where ModeID=" & ModeID & " and actcms=2    order by OrderID desc,ID Desc")
	  	Do While Not RS.Eof
			ACT_MXEdit=ACT_MXEdit &"<tr><td  width=""10%""  align=""right""  class=""tdclass""><b>"&RS("Title")&"</b>：</td><td colspan=""2""  class=""tdclass""  align=""left"">"&EditField(RS,ModeID,ID)&"</td></tr>"& vbCrLf
		RS.MoveNext
		Loop
	  RS.Close:Set RS=Nothing
 	End function


 
 



	Function EditField(RSObj,ModeID,UserID)
		Dim i,IsNotNull,TitleTypeArr,checked,rs1,FieldName,arrtitle,titles,arrvalue
	  Set RS1=ACTCMS.ACTEXE("Select * from "&ACTCMS.ACT_U(ModeID,2)&"  Where UserID="&UserID&"")
  	  If rs1.eof Then 
		  Actcms.ACTEXE("insert Into "&ACTCMS.ACT_U(ModeID,2)&"(UserID) values("&UserID&")")
		  Set RS1=Actcms.ACTEXE("Select * from "&ACtCMS.ACT_U(ModeID,2)&"  Where UserID="&UserID&"")
	  End if
	  Set RS1=ACTCMS.ACTEXE("Select * from "&ACTCMS.ACT_U(ModeID,2)&"  Where UserID="&UserID&"")
 		If rsobj("IsNotNull")="0" Then 
			IsNotNull="  <font color=red title=""必填"">*</font>  "&rsobj("Description")
		Else
			IsNotNull="  "&rsobj("Description")
		End If 
	  
	  FieldName= RSObj("FieldName")
		 Select Case RSObj("FieldType")
		   Case "TextType"
				EditField= "<input class=""ainput"" title="""&RSObj("Description")&""" type=""text"" name="""&RSObj("FieldName")&""" size="""&RSObj("width")&""" value="""&RS1(FieldName)&""">"&IsNotNull& vbCrLf 
		   Case "MultipleTextType"
				EditField= "<textarea title="""&RSObj("Description")&""" name="""&RSObj("FieldName")&""" style=""height:"&RSObj("height")&"px;width:"&RSObj("width")&"px;"">"&RS1(FieldName)&"</textarea>"&IsNotNull& vbCrLf 
		   Case "MultipleHtmlType"
				EditField=EditField&"<textarea rows=""10"" cols=""80"" name="""&RSObj("FieldName")&""" id="""&RSObj("FieldName")&""" >"&RS1(FieldName)&"</textarea>"& vbCrLf
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
				EditField= EditField&"<input class=""ainput"" name="""&RSObj("FieldName")&""" type=""text"" id="""&RSObj("FieldName")&""" value="""&RS1(FieldName)&""" onClick=""WdatePicker();""  >"&IsNotNull
		   Case "PicType"
 			 	If RSObj("Type_Type")="0" Then 
					
						  EditField=	"<input class=""ainput""  name="""&RSObj("FieldName")&""" type=""text""    value="""&RS1(FieldName)&"""  size=""40""><a style=""cursor:pointer;""  onClick=""javascript:upload(2,"&RSObj("ModeID")&",'"&RSObj("FieldName")&"');""    title=""选择已上传的图片""><font color=""#FF0000"">[点击上传图片]</font></a>"
 				Else
					EditField="<div id=""sapload"&RSObj("FieldName")&""">"& vbCrLf 
					EditField=EditField&	"</div>"& vbCrLf 
					EditField=EditField& "<script type=""text/javascript"">"& vbCrLf 
					EditField=EditField&"// <![CDATA["& vbCrLf 
					EditField=EditField&"var so = new SWFObject("""&ACTCMS.ACTSYS&"act_inc/sapload.swf"", ""sapload"&RSObj("FieldName")&""", ""450"", ""25"", ""9"", ""#ffffff"");"& vbCrLf 
					EditField=EditField&"so.addVariable('types','"&Replace(ACTCMS.ActCMS_Sys(11),"/",";")&"');"
					EditField=EditField&"so.addVariable('isGet','1');"& vbCrLf 
					EditField=EditField&"so.addVariable('args','myid=Upload;ModeID="&ModeID&";U='+U+"";""+';P='+P+"";""+'Yname="&RSObj("FieldName")&"');"& vbCrLf 
					EditField=EditField&"so.addVariable('upUrl','"&ACTCMS.ACTSYS&ACTCMS.ActCMS_Sys(8)&"/include/Upload.asp');"& vbCrLf 
					EditField=EditField&"so.addVariable('fileName','Filedata');"& vbCrLf 
					EditField=EditField&"so.addVariable('maxNum','110');"& vbCrLf 
					EditField=EditField&"so.addVariable('maxSize','"&ACTCMS.ActCMS_Sys(10)/1024&"');"& vbCrLf 
					EditField=EditField&"so.addVariable('etmsg','1');"& vbCrLf 
					EditField=EditField&"so.addVariable('ltmsg','1');"& vbCrLf 
					EditField=EditField&"so.write(""sapload"&RSObj("FieldName")&""");"& vbCrLf 
					EditField=EditField&"// ]]>"& vbCrLf 
					EditField=EditField&"</script>"			& vbCrLf 	
 					EditField=EditField&"<textarea rows=""10"" cols=""80"" name="""&RSObj("FieldName")&""" id="""&RSObj("FieldName")&""" >"&RS1(FieldName)&"</textarea>"& vbCrLf
					EditField=EditField&"<script type=""text/javascript"" language=""JavaScript"">"& vbCrLf 
 					EditField=EditField&"CKEDITOR.replace( '"&RSObj("FieldName")&"',"& vbCrLf 
					EditField=EditField&"			{"& vbCrLf 
					EditField=EditField&"				skin : 'v2',height:""250px"", width:""100%"",toolbar:'Simple'"& vbCrLf 
					EditField=EditField&"			});"& vbCrLf 
 					EditField=EditField&"</script>"&IsNotNull
				End If 
 		   Case "FileType"
				EditField= "<input  name="""&RSObj("FieldName")&""" type=""text""  value="""&RS1(FieldName)&""" size=""40""><iframe src=""../Upload_Admin.asp?ModeID=1&instr=1&instrname="&RSObj("FieldName")&"&YNContent=1&file=yes&amp;instrct=content"" name=""image"" width='75%' height=""25"" scrolling=""No"" frameborder=""0"" id=""image""></iframe>"&IsNotNull
		   Case "NumberType"
				EditField= "<input  class=""ainput"" type=""text"" name="""&RSObj("FieldName")&""" size="""&RSObj("width")&""" value="""&RS1(FieldName)&""">"&IsNotNull
		   Case "RadomType"
				EditField= "<input  class=""ainput""  type=""text"" name="""&RSObj("FieldName")&""" size=""25""  value="""&RS1(FieldName)&""">"&IsNotNull
		   Case else
				EditField= "<font color=red>该字段错误</font>"
		   End Select 
	  RS1.Close:Set RS1=Nothing
	End Function 







 End Class 
	Dim U_M
	Set U_M = New ACTM
%>