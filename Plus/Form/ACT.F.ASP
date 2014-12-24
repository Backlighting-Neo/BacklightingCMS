<!--#include file="../../ACT_inc/ACT.User.asp"-->
 <!--#include file="../../Field.asp"-->
<%'前后台公用调用
	
	   Dim Act_Form,ModeID,ModeTable,Rs,ModeName,MailBodyStr
	   ModeID = ChkNumeric(Request("ModeID"))
	   if ModeID=0 or ModeID="" Then ModeID=1
	If ACTCMS.S("A")="Save" Then 
		Dim IF_NULL,A_C	,i,UserHS,Check_F
		IF_NULL=ACTCMS.Act_MX_Arr(ModeID,3)
		Set UserHS = New ACT_User
		Set Check_F=ACTCMS.actexe("select * from ModeForm_ACT where ModeID="&ModeID&"")
		If Not Check_F.eof Then '判断表单属性
			ModeTable=Check_F("ModeTable")
			ModeName=Check_F("ModeName")
			If Check_F("ModeStatus")=1 Then Call ActCms.Alert("对不起,该表单已关闭!","")
			If  Check_F("UnlockTime")=0 Then '时间限制否?
				If Now < Check_F("StartTime") Then Call  ActCms.Alert("对不起,该表单还没有开始!","")
				If Now > Check_F("EndTime") Then Call  ActCms.Alert("对不起,该表单已经结束!","")
			End If 
		
			If  Check_F("UserGroupList") <> "0" Or Check_F("UserGroupList") = "" Then 
				If UserHS.UserLoginChecked = False Then Call  ActCms.Alert("对不起，请登录后才能提交！","")
				If Not ACTCMS.FoundInArr(Check_F("UserGroupList"),UserHS.GroupID,",") Then  Call  ActCms.Alert("对不起，您所在的用户组不能参与该表单的提交！","")
			End If 
		
			If Check_F("FormCode") =0 Then 
				If CStr(Actcms.s("Code")) <>CStr(Session("GetCode")) Then
					 Call  ActCms.Alert("验证码有误，重新输入！","")
				End If 
			End If 

			If Check_F("SubmitNum") = 0 Then 
		
				If UserHS.UserLoginChecked = False Then Call  ActCms.Alert("对不起，请登录后才能提交！","")
			 If Not ACTCMS.ACTEXE("SELECT [UserID] FROM "&ModeTable&"  Where [UserID]=" & UserHS.UserID & " and UModeID="&UserHS.UModeID&"  and ModeID="&ModeID&" order by ID desc").eof Then
				Call ACTCMS.Alert("对不起，您已提交过一次,请不要重复提交！","")
			 End if	
			End If 

			If  Check_F("Moneys") <> 0 Then '对于设置金币不等于0,将强制只能调查一次,已防出现刷金币现象
				If UserHS.UserLoginChecked = False Then Call  ActCms.Alert("对不起，请登录后才能提交！","")
 				 If Not ACTCMS.ACTEXE("SELECT [UserID] FROM "&ModeTable&" Where [UserID]='" & UserHS.UserID & "' and  UModeID="&UserHS.UModeID&"  and ModeID="&ModeID&" order by ID desc").eof Then
					Call ACTCMS.Alert("对不起，您已提交过一次,请不要重复提交！","")
				 End if	
				If UserHS.UserLoginChecked = False Then Call  ActCms.Alert("对不起，该项操作需要登录后才能提交！","")
				ACTCMS.ACTEXE("Update "&UserHS.ModeTable(UserHS.UModeID)&" Set Moneys=Moneys+"&Check_F("Moneys")&" Where UModeID="&UserHS.UModeID&"  and  UserID='" & UserHS.UserID & "'")
			End If 
		Else 
			Call actcms.alert("不存在该表单!","")
		End If 
		If IsArray(IF_NULL) Then
			For I=0 To Ubound(IF_NULL,2)
			 If IF_NULL(2,I)=0 And Trim(ACTCMS.S(IF_NULL(0,I)))="" Then  Call  ACTCMS.ALERT(IF_NULL(1,I)&"不能为空","")
			Next
		End If


		
 	
 
	UserHS.UserLoginChecked
		Set Rs=Server.CreateObject("adodb.recordset")
	    Rs.open "Select * From "&ModeTable&" Where 1=0",conn,1,3
	    Rs.Addnew
	    Rs("ModeID")=ModeID
 		Rs("UserIP")=GetIP()
		Rs("UpdateTime")=now
		If Trim(UserHS.UserID)="" Then 
			Rs("UserID")="0" 
		Else 
			Rs("UserID")=UserHS.UserID
		End If 
		If Trim(UserHS.UModeID)="" Then 
			Rs("UModeID")="0" 
		Else 
			Rs("UModeID")=UserHS.UModeID
		End If 
 	
		
			If IsArray(IF_NULL) Then
 				For I=0 To Ubound(IF_NULL,2)
 					If IF_NULL(3,I)="NumberType" Then 
					   If actcms.regexField(ACTCMS.S(IF_NULL(0,I)),"^\d+$")=True Then 
						   Rs("" & IF_NULL(0,I) & "" )= ACTCMS.S(IF_NULL(0,I))
					   End If 
					ElseIf IF_NULL(3,I)="DateType" Then 
						If IsDate(ACTCMS.S(IF_NULL(0,I)))=False Then 
							Rs("" & IF_NULL(0,I) & "")= Now()
						Else 
							Rs("" & IF_NULL(0,I) & "")=ACTCMS.S(IF_NULL(0,I))
						End If
					ElseIf IF_NULL(4,I)="1" Then 
 							 Rs("" & IF_NULL(0,I) & "")= actcms.AField(IF_NULL(5,I))
					ElseIf IF_NULL(4,I)="2" Then 
							If actcms.regexField(ACTCMS.S(IF_NULL(0,I)),IF_NULL(5,I))=True Then 
								Rs("" & IF_NULL(0,I) & "")=ACTCMS.S(IF_NULL(0,I))
							Else 
								Call Actcms.Alert(IF_NULL(6,I),"")
 							End If 
  					Else 
						Rs("" & IF_NULL(0,I) & "")=ACTCMS.S(IF_NULL(0,I))
					End If 
					actField=""
					MailBodyStr=MailBodyStr&IF_NULL(1,I)&"="&ACTCMS.S(IF_NULL(0,I))&"<br />"
 				Next
			End If			
		
		Rs.Update:Rs.Close
		MailBodyStr=ModeName&"表单提交内容如下<br />"&MailBodyStr
 	    Call ACTCMS.SendMail(ACTCMS.ActCMS_Other(3), ACTCMS.ActCMS_Other(4), ACTCMS.ActCMS_Other(5), AcTCMS.ActCMS_Sys(0) & "-有人提交了表单提交", AcTCMS.ActCMS_Sys(7),ModeName, MailBodyStr,ACTCMS.ActCMS_Other(4))
 	    Set  Rs = Nothing :Set UserHS=Nothing:Set ACTCMS=Nothing
		Response.Write ("<script>alert('你的信息已提交成功！!');history.back();</script>")
		Response.End 
	Else 
		If ACTCMS.S("A")="list" Then 
		   response.write  "<script type='text/javascript' src='act.f.asp?ModeID="&ModeID&"'></script>"
		Else 
			Call ListForm()
		End If 
	End If 
	Sub ListForm()
	 If Not ACTCMS.ACTEXE("SELECT ModeID FROM ModeForm_ACT Where ModeID=" & ModeID & " order by ModeID desc").eof Then
 		   Act_Form=Act_Form & "document.write(""<script type='text/javascript' src='" &ACTCMS.ActCMSDM&"ACT_INC/js/time/WdatePicker.js'></script>"");"& vbCrLf
		   Act_Form=Act_Form & "document.write(""<script type='text/javascript' src='" &ACTCMS.ActCMSDM&"ACT_INC/js/lhgcore/lhgcore.min.js'></script>"");"& vbCrLf
		   Act_Form=Act_Form & "document.write(""<script type='text/javascript' src='" &ACTCMS.ActCMSDM&"ACT_INC/js/lhgcore/lhgdialog.min.js'></script>"");"& vbCrLf
		   Act_Form=Act_Form & "document.write(""<script type='text/javascript' src='" &ACTCMS.ActCMSDM&"ACT_INC/main.js'></script>"");"& vbCrLf
		   Act_Form=Act_Form & "document.write(""<script type='text/javascript' src='" &ACTCMS.ActCMSDM&"ACT_INC/js/swfobject.js'></script>"");"& vbCrLf
 		   Act_Form=Act_Form & "document.write("""
		   Act_Form=Act_Form &"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>"
		   Act_Form=Act_Form & """);"& vbCrLf
		   Act_Form=Act_Form & "document.write("""
		   Act_Form=Act_Form & "<form name='myform' action='" &ACTCMS.ActCMSDM&  "plus/Form/ACT.F.ASP?A=Save&ModeID=" & ModeID & "' method='post'> "
		   Act_Form=Act_Form & """);"& vbCrLf
		   Act_Form=Act_Form& ACT_MXList(ModeID)& vbCrLf
		   Set Rs=ACTCMS.actexe("select FormCode from ModeForm_ACT where ModeID="&ModeID&"")
			if not  rs.eof then
				if Rs("FormCode")=0 then 
					 Act_Form=Act_Form & "document.write("""
					 Act_Form=Act_Form& "<tr><td>验证码：</td><td>"
					 Act_Form=Act_Form & """);"& vbCrLf
				     Act_Form=Act_Form & "document.write("""
					 Act_Form=Act_Form& "<input type='text' size='10' name='Code'> <img style='cursor:hand;'  src='"&ACTCMS.ActCMSDM&"ACT_INC/Code.asp?s=+Math.random();' id='IMG1' onclick=this.src='"&ACTCMS.ActCMSDM&"ACT_INC/Code.asp?s=+Math.random();' alt='看不清楚? 换一张！'>"
					 Act_Form=Act_Form & """);"& vbCrLf
					 Act_Form=Act_Form & "document.write("""
					 Act_Form=Act_Form& "</td></tr>"
					 Act_Form=Act_Form & """);"& vbCrLf
				end if 
			end if  
		   Act_Form=Act_Form & "document.write("""
		   Act_Form=Act_Form& "<tr> <td  colspan='2' align='center'>"
		   Act_Form=Act_Form & """);"& vbCrLf
		   Act_Form=Act_Form & "document.write("""
		   Act_Form=Act_Form&"<input type=submit   name=Submit1 value='  提 交  ' />&nbsp;"
		   Act_Form=Act_Form & """);"& vbCrLf
		   Act_Form=Act_Form & "document.write("""
		   Act_Form=Act_Form& "<input type='reset' name='Submit2'  value='  重 置  ' /></td></tr>"
		   Act_Form=Act_Form&  "</form>"
		   Act_Form=Act_Form&  "</table>"
		   Act_Form=Act_Form & """);"& vbCrLf
		   response.write Act_Form
		 End if	
	End Sub 

	Public Function ACT_MXList(ModeID)'表现方式.输出模型
	 Dim RSObj
	  Set RSObj=ACTCMS.ACTEXE("Select * from Table_ACT  Where ModeID=" & ModeID & " and actcms=3  order by OrderID desc,ID asc")
		If Not rsobj.eof Then 
			Do While Not RSObj.Eof
			    ACT_MXList=ACT_MXList & "document.write("""
				ACT_MXList=ACT_MXList &"<tr><td  width='10%'  align='left'>"&RSObj("Title")&"：</td><td align='left'>"&ListField(RSObj)&"</td></tr>"
				ACT_MXList=ACT_MXList & """);"& vbCrLf
			RSObj.MoveNext
			Loop
		End If 
	  RSObj.Close:Set RSObj=Nothing
	End function


 

 
	Function ListField(RSObj)
		Dim i,TitleTypeArr,checked,IsNotNull
		Dim arrtitle,arrvalue,titles

		If rsobj("IsNotNull")="0" Then 
			IsNotNull="  <font color=red title='必填'>*</font>  "&rsobj("Description")
		Else
			IsNotNull="  "&rsobj("Description")
		End If 
 		 Select Case RSObj("FieldType")
		   Case "TextType"
				ListField= "<input type='text' title='"&RSObj("Description")&"' name='"&RSObj("FieldName")&"' size='"&RSObj("width")&"' value='"&RSObj("Type_Default")&"'>"&IsNotNull
		   Case "MultipleTextType"
				ListField= "<textarea title='"&RSObj("Description")&"' name='"&RSObj("FieldName")&"' style='height:"&RSObj("height")&"px;width:"&RSObj("width")&"px;'>"&RSObj("Type_Default")&"</textarea>"&IsNotNull
		   Case "MultipleHtmlType"
				ListField= "<textarea title='"&RSObj("Description")&"' name='"&RSObj("FieldName")&"' style='height:"&RSObj("height")&"px;width:"&RSObj("width")&"px;'>"&RSObj("Type_Default")&"</textarea>"&IsNotNull
		   Case "RadioType"
				TitleTypeArr=Split(RSObj("Content"), vbCrLf)
				If RSObj("Type_Type")=0 Then 
				  ListField= ListField&"<select  name='"&RSObj("FieldName")&"'>"
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
					ListField = ListField & "<option value='" & arrvalue & "' "&checked&">" & titles & "</option>"
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
					ListField = ListField &"<label for='"&RSObj("FieldName")&i&"'> <input  id='"&RSObj("FieldName")&i&"' type='radio'  name='"&RSObj("FieldName")&"' value='"&arrvalue&"' "&checked&" />"&titles&"&nbsp;&nbsp;</label>" 
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
					ListField = ListField &"<label for='"&RSObj("FieldName")&i&"'> <input  id='"&RSObj("FieldName")&i&"' type='checkbox'  name='"&RSObj("FieldName")&"' value='"&arrvalue&"' "&checked&" />"&titles&"&nbsp;&nbsp;</label>"
				  Next
				  ListField = ListField&IsNotNull
				Else
				  ListField= ListField&"<select  size='4'   style='width:300px;height:126px'  name='"&RSObj("FieldName")&"' multiple>"
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
					ListField = ListField & "<option value='"& arrvalue & "' "&checked&">" & titles & "</option>"
				  Next
					ListField= ListField&" </select>"&IsNotNull
				End If 
		   Case "DateType"
				ListField= ListField&"<input name='"&RSObj("FieldName")&"' type='text' id='"&RSObj("FieldName")&"' value='' onfocus='WdatePicker()'  >"&IsNotNull
		   Case "PicType"
 		
  				  ListField=	"<input  name='"&RSObj("FieldName")&"' type='text'  value='' size='40'> <a style='cursor:pointer;' onClick=javascript:uploadform('"&actcms.ActCMSDM&"plus/Form/','"&RSObj("FieldName")&"','"&RSObj("id")&"s'); id='"&RSObj("id")&"s' title='选择已上传的图片'><font color='#FF0000'>[点击上传图片]</font></a>"
 		   
		   Case "FileType"
  				  ListField=	"<input  name='"&RSObj("FieldName")&"' type='text'  value='' size='40'> <a style='cursor:pointer;' onClick=javascript:uploadform('"&actcms.ActCMSDM&"plus/Form/','"&RSObj("FieldName")&"','"&RSObj("id")&"s'); id='"&RSObj("id")&"s' title='选择已上传的图片'><font color='#FF0000'>[点击上传图片]</font></a>"
		   Case "NumberType"
				ListField= "<input type='text' name='"&RSObj("FieldName")&"' size='"&RSObj("width")&"' value='"&RSObj("Type_Default")&"'>"&IsNotNull
		   Case "RadomType"
				ListField= "<input type='text' name='"&RSObj("FieldName")&"' size='25'  value='"&ACTCMS.MakeRandom(20)&"'>"&IsNotNull
	 

		   Case else
				ListField= "<font color=red>该字段错误</font>"
		   End Select 

 	End Function 

%> 