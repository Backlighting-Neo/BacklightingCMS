<!--#include file="../../ACT.Function.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>自定义表单管理 By ACTCMS.COM</title>
<link href="../../Images/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
	   Dim Act_Form,ModeID,ModeName,Rs
	   ModeID = ChkNumeric(Request("ModeID"))
	   if ModeID=0 or ModeID="" Then ModeID=1
		ModeName=ACTCMS.actexe("select ModeName from ModeForm_ACT where ModeID="&ModeID&"")(0)
		 %>
		<table width="590" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
    <form name="form1" method="post" action="" id="form1">
	<td align="center" class="bg_tr"><%= ModeName %>的HTML调用代码</td>
    </tr>
  <tr>
    <td align="center">
	<%	If ACTCMS.S("A")="Html" Then %>
    <textarea name="textarea" cols="50" rows="24"  style="width:100%;"><% Call ListForm() %></textarea>
	<%Else %><textarea name="textarea" cols="10" rows="4"  style="width:100%;"><%response.write "<script language=""javascript"" type=""text/javascript"" src="""
response.write actcms.ActCMS_Sys(2)&ACTCMS.ActCMSDM&"plus/Form/ACT.F.ASP?ModeID="&ModeID
	response.write """></script>"%></textarea>
	<%End If %>
	</td>
    </tr>
  <tr>
    <td align="center"><input id="Button1" type="button" value=" 复制到剪贴板 " class="ACT_btn" onClick="A_CP('textarea')" />
                                &nbsp;
    <input id="Button2" type="button" value=" 关闭对话框 " class="ACT_btn" onClick="window.close()" /></td>
  </tr></form>
</table>


	<%

	Sub ListForm()
	 If Not ACTCMS.ACTEXE("SELECT ModeID FROM ModeForm_ACT Where ModeID=" & ModeID & " order by ModeID desc").eof Then
 		   Act_Form=Act_Form & "<script type='text/javascript' src='" &ACTCMS.ActCMSDM&"ACT_INC/js/time/WdatePicker.js'></script>"& vbCrLf
		   Act_Form=Act_Form & "<script type='text/javascript' src='" &ACTCMS.ActCMSDM&"editor/ckeditor/ckeditor.js'></script>"& vbCrLf
		   Act_Form=Act_Form & "<script type='text/javascript' src='" &ACTCMS.ActCMSDM&"ACT_INC/js/lhgdialog/lhgcore.js'></script>"& vbCrLf
		   Act_Form=Act_Form & "<script type='text/javascript' src='" &ACTCMS.ActCMSDM&"ACT_INC/js/lhgdialog/lhgdialog.js'></script>"& vbCrLf
		   Act_Form=Act_Form & "<script type='text/javascript' src='" &ACTCMS.ActCMSDM&"ACT_INC/main.js'></script>"& vbCrLf
		   Act_Form=Act_Form & "<script type='text/javascript' src='" &ACTCMS.ActCMSDM&"ACT_INC/js/swfobject.js'></script>"& vbCrLf
 		   Act_Form=Act_Form &"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>"& vbCrLf
  		   Act_Form=Act_Form & "<form name='myform' action='" &ACTCMS.ActCMSDM&  "plus/Form/ACT.F.ASP?A=Save&ModeID=" & ModeID & "' method='post'> "& vbCrLf
 		   Act_Form=Act_Form& ACT_MXList(ModeID)& vbCrLf
		   Set Rs=ACTCMS.actexe("select FormCode from ModeForm_ACT where ModeID="&ModeID&"")
			if not  rs.eof then
				if Rs("FormCode")=0 then 
 					 Act_Form=Act_Form& "<tr><td>验证码：</td><td>"& vbCrLf
  					 Act_Form=Act_Form& "<input type='text' size='10' name='Code'> "& vbCrLf&"<img style='cursor:hand;'  src='"&ACTCMS.ActCMSDM&"ACT_INC/Code.asp?s=+Math.random();' id='IMG1' onclick=this.src='"&ACTCMS.ActCMSDM&"ACT_INC/Code.asp?s=+Math.random();' alt='看不清楚? 换一张！'>"& vbCrLf
  					 Act_Form=Act_Form& "</td></tr>"& vbCrLf
 				end if 
			end if  
 		   Act_Form=Act_Form& "<tr> <td  colspan='2' align='center'>"& vbCrLf
  		   Act_Form=Act_Form&"<input type=submit   name=Submit1 value='  提 交  ' />&nbsp;"& vbCrLf
 		   Act_Form=Act_Form& "<input type='reset' name='Submit2'  value='  重 置  ' /></td></tr>"& vbCrLf
		   Act_Form=Act_Form&  "</form>"& vbCrLf
		   Act_Form=Act_Form&  "</table>"& vbCrLf
 		   response.write server.HTMLEncode(Act_Form)
		 End if	
	End Sub 

	Public Function ACT_MXList(ModeID)'表现方式.输出模型
	 Dim RSObj
	  Set RSObj=ACTCMS.ACTEXE("Select * from Table_ACT  Where ModeID=" & ModeID & " and actcms=3  order by OrderID desc,ID asc")
		If Not rsobj.eof Then 
			Do While Not RSObj.Eof
 				ACT_MXList=ACT_MXList &"<tr><td  width='10%'  align='left'>"&RSObj("Title")&"：</td>"& vbCrLf&"<td align='left'>"&ListField(RSObj)&"</td></tr>"& vbCrLf
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
				ListField="<input type=hidden id="&RSObj("FieldName")&" name="&RSObj("FieldName")&" value="&RSObj("Type_Default")&"><input type=hidden id="&RSObj("FieldName")&"___Config value=><iframe id="&RSObj("FieldName")&"___Frame src="&ACTCMS.ActCMSDM&"editor/fckeditor/editor/fckeditor.html?InstanceName="&RSObj("FieldName")&"&Toolbar="&RSObj("Content")&" width="&RSObj("width")&"px height="&RSObj("height")&"px frameborder=no scrolling=no></iframe>"
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
 			 	If RSObj("Type_Type")="0" Then 
					ListField= "<input  name="""&RSObj("FieldName")&""" type=""text""  value="""" size=""40""><a style='cursor:pointer;' onClick=""javascript:upload('"&actcms.actsys&"plus/Form/','"&RSObj("FieldName")&"','999');"" title='选择已上传的图片'><font color='#FF0000'>[点击上传图片]</font></a>"&IsNotNull
				Else
					ListField="<div id=""sapload"&RSObj("FieldName")&""">"& vbCrLf 
					ListField=ListField&	"</div>"& vbCrLf 
					ListField=ListField& "<script type=""text/javascript"">"& vbCrLf 
					ListField=ListField&"// <![CDATA["& vbCrLf 
					ListField=ListField&"var so = new SWFObject("""&ACTCMS.ACTSYS&"act_inc/sapload.swf"", ""sapload"&RSObj("FieldName")&""", ""450"", ""25"", ""9"", ""#ffffff"");"& vbCrLf 
					ListField=ListField&"so.addVariable('types','"&Replace(ACTCMS.ActCMS_Sys(11),"/",";")&"');"
					ListField=ListField&"so.addVariable('isGet','1');"& vbCrLf 
					ListField=ListField&"so.addVariable('args','myid=Upload;ModeID="&ModeID&";U='+U+"";""+';P='+P+"";""+'Yname="&RSObj("FieldName")&"');"& vbCrLf 
					ListField=ListField&"so.addVariable('upUrl','"&ACTCMS.ACTSYS&"User/Upload.asp');"& vbCrLf 
					ListField=ListField&"so.addVariable('fileName','Filedata');"& vbCrLf 
					ListField=ListField&"so.addVariable('maxNum','110');"& vbCrLf 
					ListField=ListField&"so.addVariable('maxSize','"&ACTCMS.ActCMS_Sys(10)/1024&"');"& vbCrLf 
					ListField=ListField&"so.addVariable('etmsg','1');"& vbCrLf 
					ListField=ListField&"so.addVariable('ltmsg','1');"& vbCrLf 
					ListField=ListField&"so.write(""sapload"&RSObj("FieldName")&""");"& vbCrLf 
					ListField=ListField&"// ]]>"& vbCrLf 
					ListField=ListField&"</script>"			& vbCrLf 	
					ListField=ListField&"<textarea rows=""10"" cols=""80"" name="""&RSObj("FieldName")&""" id="""&RSObj("FieldName")&""" ></textarea>"& vbCrLf
					ListField=ListField&"<script type=""text/javascript"" language=""JavaScript"">"& vbCrLf 
 					ListField=ListField&"CKEDITOR.replace( '"&RSObj("FieldName")&"',"& vbCrLf 
					ListField=ListField&"			{"& vbCrLf 
					ListField=ListField&"				skin : 'v2',height:""250px"", width:""100%"",toolbar:'Simple'"& vbCrLf 
					ListField=ListField&"			});"& vbCrLf 
 					ListField=ListField&"</script>"&IsNotNull
				End If 
		   Case "FileType"
				ListField= "<input  name='"&RSObj("FieldName")&"' type='text'  value='' size='40'><iframe src='../Upload_Admin.asp?ModeID=1&instr=1&instrname="&RSObj("FieldName")&"&YNContent=1&file=yes&amp;instrct=content' name='image' width='75%' height='25' scrolling='No' frameborder='0' id='image'></iframe>"&IsNotNull
		   Case "NumberType"
				ListField= "<input type='text' name='"&RSObj("FieldName")&"' size='"&RSObj("width")&"' value='"&RSObj("Type_Default")&"'>"&IsNotNull
		   Case "RadomType"
				ListField= "<input type='text' name='"&RSObj("FieldName")&"' size='25'  value='"&ACTCMS.MakeRandom(20)&"'>"&IsNotNull
		   Case "DownType"

						 ListField="<table  border='0'   cellpadding='3' cellspacing='1'  >"
						 ListField=ListField&  "<tr ><td width='12%'   ><b>设置下载数量：</b></td>"

						 ListField=ListField& "<td width='85%' colspan='3' ><input type='text' name='no' value='4' size='2'>&nbsp;&nbsp;<input 	"		
						 ListField=ListField& " type='button' name='button' class='act_btn' onclick='setid();' value='添加下载地址数'><font color='red'>"	
						 ListField=ListField& "如果选择了使用下载服务器，请在下面↓输入文件名称。</font>"
						 ListField=ListField& " <font color='blue'>下载服务器路径 + 下载文件名称 = 完整下载地址</font><br>"
						 ListField=ListField& "</td></tr><tr><td   ><b>下载地址：</b></td><td colspan='3' >"
						 ListField=ListField& " <select name='downid' size='1'>"
						 
						 
						 ListField=ListField& "<option value='1' selected>本地软件下载服务器</option><option value='0'>↓不使用下载服务器↓</option></select>"
						 ListField=ListField& " <input name='DownFileName' type='text' size='50' value='5434'>-<input name='DownText' type='text' size='15' value='下载地址2'> "
						 ListField=ListField& "<br> <span id='upid'></span>"



						 ListField=ListField& "</td> </tr>"
						 ListField=ListField& " </table>"


		   Case else
				ListField= "<font color=red>该字段错误</font>"
		   End Select 

 	End Function 



%>		
<script>
			function A_CP(ob)
			{
				var obj=MM_findObj(ob); 
				if (obj) 
				{
					obj.select();js=obj.createTextRange();js.execCommand("Copy");}
					alert('复制成功，粘贴到你要调用的html代码里即可!');
				}
				function MM_findObj(n, d) { //v4.0
			  var p,i,x;
			  if(!d) d=document;
			  if((p=n.indexOf("?"))>0&&parent.frames.length)
			   {
				d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);
			   }
			  if(!(x=d[n])&&d.all) x=d.all[n];
			  for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
			  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
			  if(!x && document.getElementById) x=document.getElementById(n); return x;
			}
  </script>
		  