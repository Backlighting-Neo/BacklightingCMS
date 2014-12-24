 <!--#include file="ACT.F.asp"-->
<!--#include file="../act_inc/cls_pageview.asp"-->
 <% 
 			 Dim  ACT_L,U,rs,rst,rs2,C,C2,SqlStr,TemplateContent,CurrPage,PageStyle,PerPageNumber,ACT_Lable,ModeID,Parameter,ArticleSql,CurrPageStr
 			 Dim Url,urlarr
			 Url=Request.ServerVariables("QUERY_STRING")
			 urlarr=Split(url,"-")
			 C = ChkNumeric(urlarr(0))
			 If  C="0" Then  response.write "参数错误":response.End
 			 Set C2=actcms.actexe("select top 1 ModeID,UModeID,ClassTemp from space_ACT where id="&C)
			 If  C2.eof Then response.write "没有找到指定模板":response.End
			 U=C2("UModeID"):ModeID = C2("ModeID")
			 Set ACT_L = New ACT_Space
			 If  UBound(urlarr)=2 Then CurrPage=ChkNumeric(urlarr(2))
			 If CurrPage<=0 Then CurrPage=CurrPage+1
 			 Set rst=actcms.actexe("select top 1 templetsid from User_ACT where userid="&ACT_L.UID)
			 If  rst.eof Then response.write "用户没有指定模板,请到后台模板栏目进行绑定":response.End
 			 Set rs2=actcms.actexe("select top 1 templets from templets_act where id="&rst("templetsid"))
			 If  rs2.eof Then response.write "没有找到指定模板":response.End
 			 TemplateContent = ACT_L.LoadTemplate(ACTCMS.ACT_U(U,3)&"/"&rs2("templets")&"/"&C2("ClassTemp"))
			 Call CLreg()
			 If InStr(TemplateContent, "{$C}") > 0  Then
			   TemplateContent = Replace(TemplateContent, "{$C}", C)
			 End If
			 TemplateContent = ACTCMS.ReplaceUserContent(TemplateContent,ACT_L.UID)
 			 TemplateContent = ACT_L.LabelReplaceAll(TemplateContent)
		     TemplateContent=Replace(TemplateContent,"{$PageList}" ,ACT_GetPage("?"&C&"-"&ACT_L.UID,Application("PageStyle"),CurrPage,Application("PageNum"),true))
			 Dim PageArr:PageArr=Split(Application(AcTCMSN &"PageParam"),"§")
			 If Ubound(PageArr)>0  Then
		     If PageArr(0)="GetLastArticleList" Then
	         PageStyle=PageArr(3)
 			 Dim ACT_IF
			 If Ucase(Left(Trim(PageArr(4)),2))<>"ID" Then  PageArr(4)=PageArr(4) & ",ID Desc"
 			 If  PageArr(19)<>"" Then ACT_IF = " And "&PageArr(19)
 		     ArticleSql = "SELECT ID FROM "&ACTCMS.ACT_C(ModeID,2)&" Where isAccept=0 AND delif=0 "&ACT_IF&ACT_L.UIDSQL&" order by IsTop Desc," &PageArr(4) 
  			 Set RS=Server.CreateObject("ADODB.RECORDSET")
		     RS.Open ArticleSql, Conn, 1, 1
				If RS.EOF And RS.BOF Then
						TempStr = "<p>此栏目下没有文章</p>"
				Else
						   PerPageNumber=cint(PageArr(6))
						   Dim PageNum, I, J, k, TempStr,totalput,TempIDArr
							TotalPut = RS.recordcount
							if (TotalPut mod PerPageNumber)=0 then
								PageNum = TotalPut \ PerPageNumber
							else
								PageNum = TotalPut \ PerPageNumber + 1
							end if
							If CurrPage = 1 Then
									TempIDArr=IDArr(RS)
							Else
									If (CurrPage - 1) * PerPageNumber < totalPut Then
										RS.Move (CurrPage - 1) * PerPageNumber
										TempIDArr=IDArr(RS)
									Else
										CurrPage = 1
										TempIDArr=IDArr(RS)
									End If
							End If
 							SqlStr = "SELECT ID,Classid,Title,UpdateTime,actlink,FileName,InfoPurview,ReadPoint,Content,Intro,Picurl FROM "&ACTCMS.ACT_C(ModeID,2)&"   Where ID in (" & TempIDArr & ") AND isAccept=0 AND delif=0  "&ACT_IF&ACT_L.UIDSQL&" order by IsTop Desc," &PageArr(4) 
 							TempStr =  ACT_L.ACTCMS_Page_SQL(SqlStr,PageArr(5),PageArr(7),PageArr(8),PageArr(9),PageArr(10),PageArr(11),PageArr(12),PageArr(13),PageArr(14),PageArr(15),PageArr(16),PageArr(17),PageArr(18),PageArr(1),PageArr(20),ModeID,PageArr(22),PageArr(23))
							TempStr = TempStr & AcTCMS.GetPageList(PageStyle,"篇",PageNum,CurrPage,TotalPut,PerPageNumber)& ACT_GetPage("?"&C&"-"&ACT_L.UID,PageStyle,CurrPage,PageNum, True) 
					 End If
 				  RS.Close:Set RS = Nothing
			  End If
			End If
			 TemplateContent=Replace(TemplateContent,Application(AcTCMSN &"PageParam"),TempStr)

	 Function ACT_GetPage(FileName,PageStyle,CurrPage,TotalPage, TypeSelect)
			Dim PageStr, I, J, SelectStr
			 If PageStyle=0 Then PageStyle=1
			 Select Case PageStyle
			  Case 1
			   If CurrPage = 1 And CurrPage <> TotalPage Then
				PageStr = "首页  上一页 <a href=""" & FileName & "-" & CurrPage + 1 & """>下一页</a>  <a href= """ & FileName & "-" & TotalPage & """>尾页</a>"
			   ElseIf CurrPage = 1 And CurrPage = TotalPage Then
				PageStr = "首页  上一页 下一页 尾页"
			   ElseIf CurrPage = TotalPage And CurrPage <> 2 Then  
				 PageStr = "<a href=""" & FileName & """>首页</a>  <a href=""" & FileName & "-" & CurrPage - 1 & """>上一页</a> 下一页  尾页"
			   ElseIf CurrPage = TotalPage And CurrPage = 2 Then
				 PageStr = "<a href=""" & FileName & """>首页</a>  <a href=""" & FileName & """>上一页</a> 下一页  尾页"
			   ElseIf CurrPage = 2 Then
				PageStr = "<a href=""" & FileName & """>首页</a>  <a href=""" & FileName & """>上一页</a> <a href=""" & FileName & "-" & CurrPage + 1 & """>下一页</a>  <a href= """ & FileName & "-" &TotalPage & """>尾页</a>"
			   Else
				PageStr = "<a href=""" & FileName & """>首页</a>  <a href=""" & FileName & "-" & CurrPage - 1 & """>上一页</a> <a href=""" & FileName & "-" & CurrPage + 1 & """>下一页</a>  <a href= """ & FileName & "-" & TotalPage & """>尾页</a>"
			   End If
			 Case 2
			 	If CurrPage=1 Then
			     PageStr="首页 上一页"
				ElseIf CurrPage=2 Then
			     PageStr="<a href=""" & FileName & """ title=""首页"">首页</a> <a href=""" & FileName & """ title=""上一页"">上一页</a>"& vbcrlf
				Else
				 PageStr="<a href=""" & FileName & """ title=""首页"">首页</a> <a href=""" & FileName & "-"&  CurrPage - 1 &""" title=""上一页"">上一页</a> "& vbcrlf
				End If
				 For J=CurrPage To CurrPage+9
				    If J>TotalPage Then Exit For
				    If J= CurrPage Then
				     PageStr=PageStr & " <font color=red>[" & J &"]</font>"& vbcrlf
				    Else
				     PageStr=PageStr & " <a href=""" & FileName & "-" & J&""">[" & J &"]</a>"& vbcrlf
					End If
				 Next
				 If CurrPage=TotalPage Then
				  PageStr=PageStr & " 下一页 尾页"
				 Else
				  PageStr=PageStr & " <a href=""" & FileName & "-" & CurrPage + 1 & """ title=""下一页"">下一页</a> <a href=""" & FileName & "-" & TotalPage & """>尾页</a> "
				 End If
			 Case 3
			 	If CurrPage=1 Then
			     PageStr="<font face=webdings>9</font> <font face=webdings>7</font>"
				ElseIf CurrPage=2 Then
			     PageStr="<a href=""" & FileName & """ title=""首页""><font face=webdings>9</font></a> <a href=""" & FileName & """ title=""上一页""><font face=webdings>7</font></a>"
				Else
				 PageStr="<a href=""" & FileName & """ title=""首页""><font face=webdings>9</font></a> <a href=""" & FileName & "-"&  CurrPage - 1 &""" title=""上一页""><font face=webdings>7</font></a> "
				End If
				 If CurrPage=TotalPage Then
				  PageStr=PageStr & " <font face=webdings>8</font> <font face=webdings>:</font>"
				 Else
				  PageStr=PageStr & " <a href=""" & FileName & "-" & CurrPage + 1 & """ title=""上一页""><font face=webdings>8</font></a> <a href=""" & FileName & "-" & TotalPage & """><font face=webdings>:</font></a> "
				 End If
			 End Select
			   If CBool(TypeSelect) = True Then
				  PageStr = PageStr & " 转到：<select name=""page"" size=""1"" onchange=""javascript:window.location=this.options[this.selectedIndex].value;"">"& vbcrlf
				  For J = 1 To TotalPage
				   If J = CurrPage Then
					 SelectStr = " selected"
				   Else
					 SelectStr = ""
				   End If
				   If J = 1 Then
					 PageStr = PageStr & "<option value=""" & FileName & """" & SelectStr & ">第" & J & "页</option>"& vbcrlf
				   Else
					 PageStr = PageStr & "<option value=""" & FileName & "-" & J & """" & SelectStr & ">第" & J & "页</option>"& vbcrlf
				   End If
			   Next
				  PageStr = PageStr & "</select>"
			   End If
			   	ACT_GetPage=PageStr	&"</div></div>"	   
		End Function
  
   Function IDArr(rs)
	     Dim I
	     Do While Not RS.Eof
		 IDArr = IDArr &RS(0) & ","
		 RS.MoveNext
		 I = I + 1
		  If I >= PerPageNumber Then Exit Do
		 Loop
		 IDArr = Left(IDArr, Len(IDArr) - 1)
	   End Function
 
 %>
