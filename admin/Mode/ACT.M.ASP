<%
	Class ACTM

 


	Public Function ACT_NoRormMXList(ModeID)'读取系统字段输出
	 Dim RSObj
	  Set RSObj=ACTCMS.ACTEXE("Select * from Table_ACT  Where ModeID=" & ModeID & " and actcms=1  order by OrderID desc,ID asc")
	  	Do While Not RSObj.Eof
			ACT_NoRormMXList=ACT_NoRormMXList &"<tr>"&vbCrLf&"<td width=""13%"" align=""right"" class=""tdclass"">"&RSObj("Title")&"：</td>"&vbCrLf&"<td class=""tdclass"">"&ListField(RSObj)&"</td>"&vbCrLf&"</tr>"&vbCrLf
		RSObj.MoveNext
		Loop
	  RSObj.Close:Set RSObj=Nothing
	 ACT_NoRormMXList=vbCrLf&ACT_NoRormMXList& vbCrLf 
	End function

 	
	
	Public Function ACT_MXEdit(ModeID,ID)'表现方式.输出模型
	 Dim RS
	  Set RS=ACTCMS.ACTEXE("Select * from Table_ACT  Where ModeID=" & ModeID & "  and actcms=1 order by OrderID desc,ID asc")
	  	Do While Not RS.Eof
			ACT_MXEdit=ACT_MXEdit &"<tr>"&vbCrLf&"<td width=""13%"" align=""right"" class=""tdclass"">"&RS("Title")&"：</td>"&vbCrLf&"<td class=""tdclass"">"&EditField(RS,ModeID,ID)&"</td>"&vbCrLf&"</tr>"&vbCrLf
			
		RS.MoveNext
		Loop
	  RS.Close:Set RS=Nothing
	 ACT_MXEdit=vbCrLf&ACT_MXEdit& vbCrLf 
	End function

 

	Public Function ACTUser_MXList(ModeID)'读取系统字段输出
	 Dim RSObj
	  Set RSObj=ACTCMS.ACTEXE("Select * from Table_ACT  Where ModeID=" & ModeID & " and actcms=1   order by OrderID desc,ID asc")
	  	Do While Not RSObj.Eof
			ACTUser_MXList=ACTUser_MXList &"<tr>"&vbCrLf&"<th width=""20%""  class=""tdclass""><strong> "&RSObj("Title")&"：</strong><br /></td>"&vbCrLf&"<td class=""tdclass"">"&ListField(RSObj)&"</td>"&vbCrLf&"</tr>"&vbCrLf&vbCrLf
		RSObj.MoveNext
		Loop
	  RSObj.Close:Set RSObj=Nothing
	 ACTUser_MXList=vbCrLf&ACTUser_MXList& vbCrLf 
	End function



	Function ReplaceForm(ModeID,TempletContent)
	If Trim(TempletContent)="" Then Exit Function 
 		Dim MX_Arr,K
		MX_Arr=ACTCMS.Act_MX_Arr(ModeID,1)
		If IsArray(MX_Arr) Then
		  For K=0 To Ubound(MX_Arr,2)
 			  TempletContent = Replace(TempletContent,"{$" & MX_Arr(0,K) & "}","")
			  TempletContent = Replace(TempletContent,"{$" & MX_Arr(0,K) & "_title}",tip_act(ModeID,MX_Arr(0,K)))
		  Next
		End If
		ReplaceForm=TempletContent
	End Function


	Function ReplaceFormEdit(ModeID,id,TempletContent)
	On Error Resume Next
	If Trim(TempletContent)="" Then Exit Function 
 		Dim MX_Arr,K
		MX_Arr=ACTCMS.Act_MX_Arr(ModeID,1)
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
		  Set RSObj=ACTCMS.ACTEXE("Select * from Table_ACT  Where ModeID=" & ModeID & " and FieldName='"&FieldName&"'  order by OrderID asc,ID Desc")
			If Not rsobj.eof Then 
				PrintModeEdit=EditFieldlist(RSObj,ModeID,id)

			Else
				PrintModeEdit=""
			End If 
			
		  RSObj.Close:Set RSObj=Nothing
	End Function 
	


	Function EditFieldlist(RSObj,ModeID,id)
	 
		Dim rs1,FieldName
	  Set RS1=ACTCMS.ACTEXE("Select * from "&ACTCMS.ACT_C(ModeID,2)&"  Where id="&id&"")
	 FieldName= RSObj("FieldName")
 	EditFieldlist=RS1(FieldName)
	End Function 



 
	
	Function tip_act(ModeID,FieldName)
		  Dim RSObj
		  Set RSObj=ACTCMS.ACTEXE("Select * from Table_ACT  Where ModeID=" & ModeID & " and FieldName='"&FieldName&"'  order by OrderID asc,ID Desc")
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
				ListField= "<input class=""Ainput"" type=""text"" title="""&RSObj("Description")&""" name="""&RSObj("FieldName")&""" size="""&RSObj("width")&""" value="""&RSObj("Type_Default")&""">"&IsNotNull
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
				ListField= ListField&"<input class=""Ainput"" name="""&RSObj("FieldName")&""" type=""text"" id="""&RSObj("FieldName")&""" value=""""  onClick=WdatePicker();  >"&IsNotNull
		   Case "PicType"
 			 	If RSObj("Type_Type")="0" Then 
						  ListField=	"<input class=""Ainput""  name="""&RSObj("FieldName")&""" type=""text""  value="""" size=""50""><input name=""button""  onClick=""J('#"&RSObj("id")&"s').dialog({ id:'actcmsscs' ,page: '"&actcms.actsys&actcms.adminurl&"/include/Upload_Admin.asp?A=add&instr=2&ModeID="&RSObj("ModeID")&"&instrname="&RSObj("FieldName")&"',  width:720, height:240 });""   id="""&RSObj("id")&"s""   type=""button""  class=""ACT_btn"" style=""cursor:hand;"" value=""点击上传文件"">"
 				Else
					ListField="<div id=""sapload"&RSObj("FieldName")&""">"& vbCrLf 
					ListField=ListField&	"</div>"& vbCrLf 
					ListField=ListField& "<script type=""text/javascript"">"& vbCrLf 
					ListField=ListField&"// <![CDATA["& vbCrLf 
					ListField=ListField&"var so = new SWFObject("""&ACTCMS.ACTSYS&"act_inc/sapload.swf"", ""sapload"&RSObj("FieldName")&""", ""450"", ""25"", ""9"", ""#ffffff"");"& vbCrLf 
					ListField=ListField&"so.addVariable('types','"&Replace(ACTCMS.ActCMS_Sys(11),"/",";")&"');"
					ListField=ListField&"so.addVariable('isGet','1');"& vbCrLf 
					ListField=ListField&"so.addVariable('args','myid=Upload;ModeID="&ModeID&";U='+U+"";""+';P='+P+"";""+'Yname="&RSObj("FieldName")&"');"& vbCrLf 
					ListField=ListField&"so.addVariable('upUrl','"&ACTCMS.ACTSYS&ACTCMS.ActCMS_Sys(8)&"/include/Upload.asp');"& vbCrLf 
					ListField=ListField&"so.addVariable('fileName','Filedata');"& vbCrLf 
					ListField=ListField&"so.addVariable('maxNum','110');"& vbCrLf 
					ListField=ListField&"so.addVariable('maxSize','"&ACTCMS.ActCMS_Sys(10)/1024&"');"& vbCrLf 
					ListField=ListField&"so.addVariable('etmsg','1');"& vbCrLf 
					ListField=ListField&"so.addVariable('ltmsg','1');"& vbCrLf 
					ListField=ListField&"so.addParam('wmode','transparent');"& vbCrLf 
 					ListField=ListField&"so.write(""sapload"&RSObj("FieldName")&""");"& vbCrLf 
					ListField=ListField&"// ]]>"& vbCrLf 
					ListField=ListField&"</script>"			& vbCrLf 	
					ListField=ListField&"<textarea rows=""10"" cols=""80"" name="""&RSObj("FieldName")&""" id="""&RSObj("FieldName")&"""  style=""width:"&RSObj("width")&"px;height:"&RSObj("height")&"px;visibility:hidden;"" ></textarea>"& vbCrLf
					ListField=ListField&"<script>"& vbCrLf 
					ListField=ListField&"KE.show({"& vbCrLf 
					ListField=ListField&"id : '"&RSObj("FieldName")&"'"& vbCrLf 
					ListField=ListField&"});"& vbCrLf 
					ListField=ListField&"</script>"&IsNotNull

				End If 
		   
		   Case "FileType"
				ListField= "<input  name="""&RSObj("FieldName")&""" type=""text""  value="""" size=""40""><iframe src="""&actcms.actsys&actcms.adminurl&"/Upload_Admin.asp?ModeID=1&instr=1&instrname="&RSObj("FieldName")&"&YNContent=1&file=yes&amp;instrct=content"" name=""image"" width='75%' height=""25"" scrolling=""No"" frameborder=""0"" id=""image""></iframe>"&IsNotNull
		   Case "NumberType"
				ListField= "<input class=""Ainput""  type=""text"" name="""&RSObj("FieldName")&""" size="""&RSObj("width")&""" value="""&RSObj("Type_Default")&""">"&IsNotNull
		   Case "RadomType"
				ListField= "<input class=""Ainput""  type=""text"" name="""&RSObj("FieldName")&""" size=""25""  value="""&ACTCMS.MakeRandom(20)&""">"&IsNotNull
		   Case "DownType"

						 ListField="<table  border=""0""   cellpadding=""3"" cellspacing=""1""  >"
						 ListField=ListField&  "<tr ><td width=""12%""   ><b>设置下载数量：</b></td>"

						 ListField=ListField& "<td width=""85%"" colspan=""3"" ><input type=""text"" name=""no"" value=""4"" size=""2"">&nbsp;&nbsp;<input 	"		
						 ListField=ListField& " type=""button"" name=""button"" class=""act_btn"" onclick=""setid();"" value=""添加下载地址数""><font color=""red"">"	
						 ListField=ListField& "如果选择了使用下载服务器，请在下面↓输入文件名称。</font>"
						 ListField=ListField& " <font color=""blue"">下载服务器路径 + 下载文件名称 = 完整下载地址</font><br>"
						 ListField=ListField& "</td></tr><tr><td   ><b>下载地址：</b></td><td colspan=""3"" >"
						 ListField=ListField& " <select name=""downid"" size=""1"">"
						 
						 
						 ListField=ListField& "<option value=""1"" selected>本地软件下载服务器</option><option value=""0"">↓不使用下载服务器↓</option></select>"
						 ListField=ListField& " <input name=""DownFileName"" type=""text"" size=""50"" value='5434'>-<input name=""DownText"" type=""text"" size=""15"" value='下载地址2'> "
						 ListField=ListField& "<br> <span id=""upid""></span>"



						 ListField=ListField& "</td> </tr>"
						 ListField=ListField& " </table>"


		   Case else
				ListField= "<font color=red>该字段错误</font>"
		   End Select 
	End Function 






	Function EditField(RSObj,ModeID,id)
		Dim i,IsNotNull,TitleTypeArr,checked,rs1,FieldName
		Dim arrtitle,arrvalue,titles
	  Set RS1=ACTCMS.ACTEXE("Select * from "&ACTCMS.ACT_C(ModeID,2)&"  Where id="&id&"")
	  FieldName= RSObj("FieldName")
	
		If rsobj("IsNotNull")="0" Then 
			IsNotNull="  <font color=red title=""必填"">*</font>  "&rsobj("Description")
		Else
			IsNotNull="  "&rsobj("Description")
		End If 
		 
		 Select Case RSObj("FieldType")
		   Case "TextType"
				EditField= "<input class=""Ainput"" title="""&RSObj("Description")&""" type=""text"" name="""&RSObj("FieldName")&""" size="""&RSObj("width")&""" value="""&RS1(FieldName)&""">"&IsNotNull& vbCrLf 
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
				EditField= EditField&"<input class=""Ainput"" name="""&RSObj("FieldName")&""" type=""text"" id="""&RSObj("FieldName")&""" value="&RS1(FieldName)&" onClick=WdatePicker();  >"&IsNotNull
		   Case "PicType"
  			 	If RSObj("Type_Type")="0" Then 
					  EditField=	"<input class=""Ainput""  name="""&RSObj("FieldName")&""" type=""text""    value="""&RS1(FieldName)&"""  size=""40""><input name=""button""  onClick=""J('#"&RSObj("id")&"s').dialog({ id:'actcmsscs' ,page: '"&actcms.actsys&actcms.adminurl&"/include/Upload_Admin.asp?A=add&instr=2&ModeID="&RSObj("ModeID")&"&instrname="&RSObj("FieldName")&"',  width:720, height:240 });""   id="""&RSObj("id")&"s""   type=""button""  class=""ACT_btn"" style=""cursor:hand;"" value=""点击上传文件"">"
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
				EditField= "<input  name="""&RSObj("FieldName")&""" type=""text""  value="""&RS1(FieldName)&""" size=""40""><iframe src="""&actcms.actsys&actcms.adminurl&"/Upload_Admin.asp?ModeID=1&instr=1&instrname="&RSObj("FieldName")&"&YNContent=1&file=yes&amp;instrct=content"" name=""image"" width='75%' height=""25"" scrolling=""No"" frameborder=""0"" id=""image""></iframe>"&IsNotNull
		   Case "NumberType"
				EditField= "<input class=""Ainput""  type=""text"" name="""&RSObj("FieldName")&""" size="""&RSObj("width")&""" value="""&RS1(FieldName)&""">"&IsNotNull
		   Case "RadomType"
				EditField= "<input class=""Ainput""  type=""text"" name="""&RSObj("FieldName")&""" size=""25""  value="""&RS1(FieldName)&""">"&IsNotNull
		   Case else
				EditField= "<font color=red>该字段错误</font>"
		   End Select 
	  RS1.Close:Set RS1=Nothing
	End Function 





 


End Class 
	Dim M
	Set M = New ACTM
%>