<!--#include file="ACT_inc/ACT.User.asp"-->
<%
Dim ACTCLS,ModeID
 

Dim Url,urlarr,ACT_L,UserHS,ACT_Lable,PerPageNumber,TypeContent,UserID,PayTF,classid
		Dim CurrPage,ID,InfoPurview,ReadPoint,ClassPurview,ClassReadPoint,UserLoginTF,ChargeType,PitchTime,ReadTimes
Url=Request.ServerVariables("QUERY_STRING")
urlarr=Split(Split(url,".")(0),"-")

	
	
	  Set ACT_L = New ACT_Code
	  Set UserHS = New ACT_User
 		Select Case urlarr(0)
			Case "C"
				Call TypeArticle()
			Case "L"
				Call L()
			Case Else
				response.write "error"
		End Select 
		Call CloseConn
  		Sub TypeArticle()
		 Dim  SqlStr,TemplateContent,Rs
		UserLoginTF=Cbool(UserHS.UserLoginChecked)

  			 ID = ChkNumeric(RSQL(urlarr(2)))
 			 If UBound(urlarr)>2 Then CurrPage=ChkNumeric(urlarr(3))
			 ModeID=ChkNumeric(urlarr(1))
 			 If ModeID=0 Then ModeID=1
			 If UBound(urlarr)=4 Then PayTF=urlarr(4)
			 
  			 If CurrPage<=0 Then CurrPage=CurrPage+1
			 If ID = 0 Or ID = "" Then Exit Sub
			 Set Rs=actcms.actexe("Select * From "&ACTCMS.ACT_C(ModeID,2)&" where ID=" & ID)
			 If Rs.Eof And Rs.Bof Then
				Call ACTCMS.Alert("您要查看的文章已删除。或是您非法传递注入参数!!",AcTCMS.ActCMSDM):Response.End
			 ElseIf Rs("actlink") = 1 Then
				 Response.Redirect Rs("FileName")
			End If
		
		    Dim DocXML,Node:Set DocXML=actcms.arrayToXml(Rs.GetRows(1),Rs,"row","root")
			Set Node=DocXml.DocumentElement.SelectSingleNode("row")
			Set ACT_L.Nodes=DocXml.DocumentElement.SelectSingleNode("row")

			TypeContent=ACT_L.GetNodeText("content")
 			 InfoPurview = Cint(ACT_L.GetNodeText("infopurview"))
			 ReadPoint   = Cint(ACT_L.GetNodeText("readpoint"))
			 ChargeType  = Cint(ACT_L.GetNodeText("chargetype"))
			 PitchTime   = Cint(ACT_L.GetNodeText("pitchtime"))
			 ReadTimes   = Cint(ACT_L.GetNodeText("readtimes"))
			 UserID   = ChkNumeric(ACT_L.GetNodeText("userid"))
			 classid   = ACT_L.GetNodeText("classid")
			 ClassPurview= Cint(actcms.ACT_L(ACT_L.GetNodeText("classid"),19))
			ClassReadPoint= Cint(actcms.ACT_L(ACT_L.GetNodeText("classid"),20))
 			
			Dim ClassChargeType,ClassPitchTime,ClassReadTimes
			
			If InfoPurview=2 or ReadPoint>0 Then
				IF UserLoginTF=false Then
					Call GetNoLoginInfo()
				Else 
 					If  ACTCMS.FoundInArr(ACTCMS.ACT_L(ACT_L.GetNodeText("classid"),6),Trim(UserHS.GroupID),",")=False and readpoint=0  Then 
						  Call ACTCMS.Alert("对不起，你所在的用户组没有查看的权限1!",AcTCMS.ActCMSDM)
					 Else
						  Call PayPointProcess()
					End If 
				End If 
		ElseIF InfoPurview=0 And (ClassPurview=1 or ClassPurview=2 Or ClassReadPoint>0) Then 
			  If UserLoginTF=false Then
			    Call GetNoLoginInfo
			  Else     
 			     ReadPoint  = Cint(actcms.ACT_L(ACT_L.GetNodeText("classid"),20))
				 ChargeType = Cint(actcms.ACT_L(ACT_L.GetNodeText("classid"),21))
				 PitchTime  = Cint(actcms.ACT_L(ACT_L.GetNodeText("classid"),22))
				 ReadTimes  = Cint(actcms.ACT_L(ACT_L.GetNodeText("classid"),23))
 				 If ClassPurview=2 Then
					If ACTCMS.FoundInArr(ACTCMS.ACT_L(ACT_L.GetNodeText("classid"),6),Trim(UserHS.GroupID),",")=false Then
						  Call ACTCMS.Alert("对不起，你所在的用户组没有查看的权限!",AcTCMS.ActCMSDM)
					 Else
						Call PayPointProcess()
					 End If
				 Else    
				 Call PayPointProcess()
 				End If
			  End If
		 Else
		   Call PayPointProcess()
		 End If   		
 
			 If ACT_L.GetNodeText("isaccept")<>0 Then
				 If UserHS.UserName<>ACT_L.GetNodeText("articleinput") Then
				   Call ACTCMS.Alert("对不起，该文章还没有通过审核!",AcTCMS.ActCMSDM)
				   Response.End
			     End If 
			 End If
			Application(AcTCMSN & "ACTCMS_TCJ_Type") = "ARTICLECONTENT"
			Application(AcTCMSN & "classid") = ACT_L.GetNodeText("classid")
			Application(AcTCMSN & "modeid")=ModeID
			Application(AcTCMSN & "id")=ACT_L.GetNodeText("id")
			 id = ACT_L.GetNodeText("id")
			 classid=ACT_L.GetNodeText("classid")
			 TemplateContent = ACT_L.LoadTemplate(ACT_L.GetNodeText("templateurl"))
			 TemplateContent = ACT_L.LabelReplaceAll(TemplateContent)
			 Dim ContentArr:ContentArr=Split(ACT_L.GetNodeText("content"),"[NextPage]")
			 Dim TotalPage,N,ArticlePageStr
			 TotalPage = Cint(UBound(ContentArr) + 1)
			   If TotalPage > 1 Then
					   If CurrPage = 1 Then
						 ArticlePageStr = "<p><div Class=""PageCss"" align=center><a href="""&actcms.acturl&"list.asp?C-" & ModeID & "-" & ID & "-" &(CurrPage + 1) & ".Html"">下一页</a><br>"
					   ElseIf CurrPage = TotalPage Then
						 ArticlePageStr = "<p><div Class=""PageCss"" align=center><a href="""&actcms.acturl&"list.asp?C-" & ModeID & "-" & ID & "-" &(CurrPage - 1) & ".Html"">上一页</a><br>"
					   Else
						ArticlePageStr = "<p><div Class=""PageCss"" align=center><a href="""&actcms.acturl&"list.asp?C-" & ModeID & "-" &  ID & "-" &(CurrPage - 1) & ".Html"">上一页</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="""&actcms.acturl&"list.asp?C-" & ModeID & "-" & ID & "-" &(CurrPage + 1) & ".Html"">下一页</a><br>"
					   End If
					   ArticlePageStr = ArticlePageStr & "本文共<b> " & TotalPage & " </b>页,第&nbsp;&nbsp;"
				   For N = 1 To TotalPage
						 If CurrPage = N Then
						  ArticlePageStr = ArticlePageStr & "<b Class=""PageCss"">[" & N & "]</b>&nbsp;"
						 Else
						  ArticlePageStr = ArticlePageStr & "<a Class=""PageCss"" href="""&actcms.acturl&"list.asp?C-" & ModeID & "-" & ID & "-" & N & ".Html"">[" & N & "]</a>&nbsp;"
						 End If
					  If TotalPage > 8 Then
					   If N Mod 8 = 0 Then ArticlePageStr = ArticlePageStr & "<p>"
					  End If
					Next
					ArticlePageStr = ContentArr(CurrPage-1) & ArticlePageStr & "页</div></p>"
				 Else
				  ArticlePageStr = TypeContent
				 End If
				
			TemplateContent= ACT_L.ReplaceArticleContent(ModeID,TemplateContent,ArticlePageStr)
			TemplateContent=ACT_L.actcmsexe(TemplateContent)'自定义函数
			 response.write TemplateContent
		End Sub

 	   Sub GetNoLoginInfo()
		   TypeContent="<div align=center>对不起，您还没有登录，本文至少要求本站的注册会员才可查看!</div><div align=center>如果您还没有注册，请<a href=""" & ACTCMS.ACTCMSDM & "User/Reg.asp""><font color=red>点此注册</font></a>吧!</div><div align=center>如果您已是本站注册会员，赶紧<a href=""" & ACTCMS.ACTCMSDM & "User/login.asp""><font color=red>点此登录</font></a>吧！</div>"
	   End Sub



	   '收费扣点处理过程
	   Sub PayPointProcess()
 	       Dim UserChargeType:UserChargeType=UserHS.ChargeType
	        If (Cint(ReadPoint)>0 or InfoPurview=2 or (InfoPurview=0 And (ClassPurview=1 Or ClassPurview=2))) and UserHS.UserID<>UserID Then
					 
					     If UserChargeType=1 Then
							 Select Case ChargeType
							  Case 0:Call CheckPayTF("1=1")
							  Case 1:Call CheckPayTF("datediff('h',AddDate," & NowString & ")<" & PitchTime)
							  Case 2:Call CheckPayTF("Times<" & ReadTimes)
							  Case 3:Call CheckPayTF("datediff('h',AddDate," & NowString & ")<" & PitchTime & " or Times<" & ReadTimes)
							  Case 4:Call CheckPayTF("datediff('h',AddDate," & NowString & ")<" & PitchTime & " and Times<" & ReadTimes)
							  Case 5:Call PayConfirm()
							  End Select
						Elseif UserChargeType=2 Then
				          If UserHS.GetEdays <=0 Then
						     Content="<div align=center>对不起，你的账户已过期 <font color=red>" & UserHS.GetEdays & "</font> 天,此文需要在有效期内才可以查看，请及时与我们联系！</div>"
 						  End If
 						end if
 					   End IF
	   End Sub


	   '检查是否过期，如果过期要重复扣点券
	   '返回值 过期返回 true,未过期返回false
	   Sub CheckPayTF(Param)
	    Dim SqlStr:SqlStr="Select top 1 Times From Point_Log_ACT Where ModeID=" & ModeID & " And InfoID=" & ID & " And PointFlag=2 and UserID=" & UserHS.UserID & " And (" & Param & ") Order By ID"
   	    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open SqlStr,conn,1,3


		IF RS.Eof And RS.Bof Then
			Call PayConfirm()	
		Else
		       RS.Movelast
			   RS(0)=RS(0)+1
			   RS.Update
		End IF
		 RS.Close:Set RS=nothing
	   End Sub
	   
 
	   Sub PayConfirm()
	     If UserLoginTF=false Then Call GetNoLoginInfo():Exit Sub
		 If ReadPoint<=0 Then Exit Sub

			 If Cint(UserHS.Point)<ReadPoint Then
					 TypeContent="<div style=""text-align:center"">对不起，你的可用" & actcms.ActCMS_Sys(24) & "不足!阅读本文需要 <span style=""color:red"">" & ReadPoint & "</font> " & actcms.ActCMS_Sys(25) & actcms.ActCMS_Sys(24) &",你还有 <span style=""color:green"">" & UserHS.Point & "</span> " & actcms.ActCMS_Sys(25) & actcms.ActCMS_Sys(24) & "</div>,请及时与我们联系！" 
			 Else 
 					If PayTF="1" Then
						Call ACTCMS.PointInOrOut(ModeID,ID,UserHS.UserID,2,ReadPoint,"系统","阅读文档收费",0)
 						 Dim PayPoint:PayPoint=(ReadPoint*ActCMS.ACT_L(classid,24))/100
						 If PayPoint>0 Then
						 
						Call ACTCMS.PointInOrOut(ModeID,ID,UserID,1,PayPoint,"系统","阅读文档收费",0)

 						 End If
 						 
					Else
						TypeContent="<div align=center>阅读本文需要消耗 <font color=red>" & ReadPoint & "</font> " & actcms.ActCMS_Sys(25) & actcms.ActCMS_Sys(24) &",你目前尚有 <font color=green>" & UserHS.Point & "</font> " & actcms.ActCMS_Sys(25) & actcms.ActCMS_Sys(24) &"可用,阅读本文后，您将剩下 <font color=blue>" & UserHS.Point-ReadPoint & "</font> " & actcms.ActCMS_Sys(25) & actcms.ActCMS_Sys(24) &"</div><div align=center>你确实愿意花 <font color=red>" & ReadPoint & "</font> " & actcms.ActCMS_Sys(25) & actcms.ActCMS_Sys(24) & "来阅读此文吗?</div><div>&nbsp;</div><div align=center><a href=""?C-"&ModeID&"-" & ID & "-" & CurrPage &"-1.Html"">我愿意</a>    <a href=""" &AcTCMS.ActCMSDM & """>我不愿意</a></div>"
					End If
			 End If
	   End Sub






'---------------------------------------栏目--------------------------------------------------
 		Public Sub L()
			Dim RsClass,SqlStr,TemplateContent,CurrPage,PageStyle,ACT_Lable,ModeID
			 classid = RSQL(urlarr(1))
 			 If UBound(urlarr)=2 Then CurrPage=ChkNumeric(urlarr(2))
			 If CurrPage<=0 Then CurrPage=CurrPage+1
		     UserHS.UserLoginChecked
			 If  classid = "" Then Exit Sub 
			 Set RsClass=actcms.actexe("Select FolderTemplate,classid,Extension,ParentID,GroupIDClass,ModeID,actlink,content,makehtmlname From Class_ACT Where classid='" & classid & "'")
			 IF RsClass.Eof And RsClass.Bof Then
			  Call ACTCMS.Alert("非法参数!",AcTCMS.ActCMSDM)
			  Exit Sub
			 End If
			 If RsClass("actlink")="2" Then 
				response.Redirect RsClass("makehtmlname")
				response.end
			  End If 
			If ACTCMS.ACT_L(RsClass("classid"),6)<>"" Then
				If  ACTCMS.FoundInArr(ACTCMS.ACT_L(RsClass("classid"),6),UserHS.GroupID,",")=False Then 
 					  Call ACTCMS.Alert("对不起，你所在的用户组没有查看的权限!",AcTCMS.ActCMSDM)
				End If 
 			End If
			 
			Application(AcTCMSN & "classid")=  RsClass("classid")
			Application(AcTCMSN & "modeid")= RsClass("ModeID")
			ModeID= RsClass("ModeID")
			Application(AcTCMSN & "ACTCMS_TCJ_Type")= "Folder"
			Application(AcTCMSN & "Make")="No"
			If Trim(RsClass("ParentID")) = "0" Then	Application(AcTCMSN & "ModeHome") = True	Else Application(AcTCMSN & "ModeHome") = False
			 TemplateContent = ACT_L.LoadTemplate(RsClass("FolderTemplate"))
 			 If RsClass("actlink")="3" Then 
 				TemplateContent=Replace(TemplateContent, "{$GetClassIntro}", RsClass("content"))
			 End If 
 			 TemplateContent = ACT_L.LabelReplaceAll(TemplateContent)
 			 If Application(AcTCMSN & "PageStyle")<>4 Then 
 				 TemplateContent=Replace(TemplateContent,"{$PageList}" ,ACT_GetPage("list.asp?L-" & classid,Application(AcTCMSN & "PageStyle"),CurrPage,Application(AcTCMSN & "PageNum"),true))
 			Else
 				'TemplateContent=Replace(TemplateContent,"{$pagelist}",ACT_DIY_Page("?L-" & classid,Application(AcTCMSN & "PageStyle"),CurrPage,Application(AcTCMSN & "PageNum"), True))
				TemplateContent=Replace(TemplateContent,"{$PageList}","")
			End If 
			Dim PageArr:PageArr=Split(Application(AcTCMSN &"PageParam"),"§")
			If Ubound(PageArr)>0  Then
			  If PageArr(0)="GetLastArticleList" Then
			       PageStyle=PageArr(3)
				   Dim ArticleSql,CurrPageStr,classid
		   Dim Parameter
 			Select Case PageArr(2) 
			    Case "","0":Parameter=""
 				Case "1"
					If Application(AcTCMSN & "classid")<>"0"  Then 
						If  CBool(PageArr(21))=True Then 
							 Parameter="classid In (" & ACTCMS.Tempclassid(Application(AcTCMSN & "classid")) & ") And"
						Else 
							Parameter="classid='" & Application(AcTCMSN & "classid") & "' And" 
						End If 
					End If 
				Case Else
					If InStr(PageArr(2), ",") > 0 Then
						 Parameter="classid In (" & PageArr(2) & ") And"
					Else
						If CBool(PageArr(22))=True Then 
						 Parameter="classid In (" & ACTCMS.Tempclassid(PageArr(2)) & ") And"
						Else 
						 Parameter="classid='" & Replace(PageArr(2),"'","") & "' And"
						End If 
					End If 
			End Select
			Dim ACT_IF
			If Ucase(Left(Trim(PageArr(4)),2))<>"ID" Then  PageArr(4)=PageArr(4) & ",ID Desc"
			If PageArr(21)="1" Then ModeID=Cint(Application(AcTCMSN & "modeid"))
			If  PageArr(19)<>"" Then ACT_IF = " And "&PageArr(19)
 		    ArticleSql = "SELECT ID FROM "&ACTCMS.ACT_C(RsClass("ModeID"),2)&" Where " & Parameter & " isAccept=0 AND delif=0 "&ACT_IF&" order by IsTop Desc," &PageArr(4) 
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		    RS.Open ArticleSql, Conn, 1, 1
				If RS.EOF And RS.BOF Then
					 TempStr = "<p>此栏目下没有文章</p>"
 				Else
 					   PerPageNumber=cint(PageArr(6))
					   Dim PageNum, I, J, k, TempStr,totalput,TempIDArr
						TotalPut = RS.recordcount
 						if (TotalPut mod PerPageNumber)=0 then
							PageNum = TotalPut \ PerPageNumber
						Else 
							PageNum = TotalPut \ PerPageNumber + 1
						End  If 
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
						SqlStr = "SELECT ID,classid,Title,UpdateTime,ActLink,FileName,InfoPurview,ReadPoint,PicUrl,Intro,Content,CopyFrom,Author,KeyWords"&ACTCMS.DIYField(ModeID)&"  FROM "&ACTCMS.ACT_C(RsClass("ModeID"),2)&"   Where ID in (" & TempIDArr & ") AND isAccept=0 AND delif=0  "&ACT_IF&" order by IsTop Desc," &PageArr(4) 
						TempStr =  ACT_L.ACTCMS_Page_SQL(SqlStr,PageArr(5),PageArr(7),PageArr(8),PageArr(9),PageArr(10),PageArr(11),PageArr(12),PageArr(13),PageArr(14),PageArr(15),PageArr(16),PageArr(17),PageArr(18),PageArr(1),PageArr(20),ModeID,PageArr(22),PageArr(23))
 						If PageStyle<>4 Then TempStr = TempStr & AcTCMS.GetPageList(PageStyle,ACTCMS.ACT_C(ModeID,5),PageNum,CurrPage,TotalPut,PerPageNumber)& ACT_GetPage("?L-" & classid,PageStyle,CurrPage,PageNum, True) 
 				 End If
 				 RS.Close:Set RS = Nothing
			  End If
			Else
			   PageNum=Application(AcTCMSN & "PageNum")
			   TotalPut=Application(AcTCMSN & "TotalPut")
			   CurrPage=Application(AcTCMSN & "CurrPage")
			End If
 			 TemplateContent=Replace(TemplateContent,Application(AcTCMSN &"PageParam"),TempStr)
			 If PageStyle=4 Or Application(AcTCMSN & "PageStyle") =4 Then 
				If ACTCMS.ACT_C(ModeID,3)=2 Then 
 				 TemplateContent=Replace(TemplateContent,"{$pagelist}",ACT_DIY_Page(actcms.acturl&"list-" & classid,PageStyle,CurrPage,PageNum, True))
				Else 
 				 TemplateContent=Replace(TemplateContent,"{$pagelist}",ACT_DIY_Page(actcms.acturl&"list.asp?L-" & classid,PageStyle,CurrPage,PageNum, True))
				End If 
				 TemplateContent=Replace(TemplateContent,"{$pagecount}",TotalPut)
				 TemplateContent=Replace(TemplateContent,"{$pagethis}",CurrPage)
				 TemplateContent=Replace(TemplateContent,"{$pagenum}",PageNum)
 			 End If 
			 TemplateContent=ACT_L.actcmsexe(TemplateContent)'自定义函数
			 response.write TemplateContent
		End Sub 
  	    Function ACT_DIY_Page(FileName,PageStyle,CurrPage,TotalPage, TypeSelect)
 			Dim PageStr, I, J, SelectStr
			 If PageStyle=0 Then PageStyle=1
			 If ChkNumeric(TotalPage)=0 Then TotalPage=1
 				If CurrPage=1 Then
			     PageStr=" 首页 上一页"
				ElseIf CurrPage=2 Then
			     PageStr="<a href=""" & FileName & ".Html"" title=""首页"">首页</a> <a href=""" & FileName & ".Html"" title=""上一页"">上一页</a>"& vbcrlf
				Else
				 PageStr="<a href=""" & FileName & ".Html"" title=""首页"">首页</a> <a href=""" & FileName & "-"&  CurrPage - 1 &".Html"" title=""上一页"">上一页</a> "& vbcrlf
				End If
				 For J=CurrPage To CurrPage+9
				    If J>TotalPage Then Exit For
				    If J= CurrPage Then
				     PageStr=PageStr & " <strong>[" & J &"]</strong>"& vbcrlf
				    Else
				     PageStr=PageStr & " <a href=""" & FileName & "-" & J&".Html"">[" & J &"]</a>"& vbcrlf
					End If
				 Next
 				 If CurrPage=TotalPage Then
				  PageStr=PageStr & " 下一页 尾页"
				 Else
				  PageStr=PageStr & " <a href=""" & FileName & "-" & CurrPage + 1 & ".Html"" title=""下一页"">下一页</a> <a href=""" & FileName & "-" & TotalPage & ".Html"">尾页</a> "
				 End If
			   	ACT_DIY_Page=PageStr	
 	 End Function
 	 Function ACT_GetPage(FileName,PageStyle,CurrPage,TotalPage, TypeSelect)
			Dim PageStr, I, J, SelectStr
			 If PageStyle=0 Then PageStyle=1
			 Select Case PageStyle
			  Case 1
			   If CurrPage = 1 And CurrPage <> TotalPage Then
				PageStr = "首页  上一页 <a href=""" & FileName & "-" & CurrPage + 1 & ".Html"">下一页</a>  <a href= """ & FileName & "-" & TotalPage & ".Html"">尾页</a>"
			   ElseIf CurrPage = 1 And CurrPage = TotalPage Then
				PageStr = "首页  上一页 下一页 尾页"
			   ElseIf CurrPage = TotalPage And CurrPage <> 2 Then  
				 PageStr = "<a href=""" & FileName & ".Html"">首页</a>  <a href=""" & FileName & "-" & CurrPage - 1 & ".Html"">上一页</a> 下一页  尾页"
			   ElseIf CurrPage = TotalPage And CurrPage = 2 Then
				 PageStr = "<a href=""" & FileName & ".Html"">首页</a>  <a href=""" & FileName & ".Html"">上一页</a> 下一页  尾页"
			   ElseIf CurrPage = 2 Then
				PageStr = "<a href=""" & FileName & ".Html"">首页</a>  <a href=""" & FileName & ".Html"">上一页</a> <a href=""" & FileName & "-" & CurrPage + 1 & ".Html"">下一页</a>  <a href= """ & FileName & "-" &TotalPage & ".Html"">尾页</a>"
			   Else
				PageStr = "<a href=""" & FileName & ".Html"">首页</a>  <a href=""" & FileName & "-" & CurrPage - 1 & ".Html"">上一页</a> <a href=""" & FileName & "-" & CurrPage + 1 & ".Html"">下一页</a>  <a href= """ & FileName & "-" & TotalPage & ".Html"">尾页</a>"
			   End If
			 Case 2
			 	If CurrPage=1 Then
			     PageStr="首页 上一页"
				ElseIf CurrPage=2 Then
			     PageStr="<a href=""" & FileName & ".Html"" title=""首页"">首页</a> <a href=""" & FileName & ".Html"" title=""上一页"">上一页</a>"& vbcrlf
				Else
				 PageStr="<a href=""" & FileName & ".Html"" title=""首页"">首页</a> <a href=""" & FileName & "-"&  CurrPage - 1 &".Html"" title=""上一页"">上一页</a> "& vbcrlf
				End If
				 For J=CurrPage To CurrPage+9
				    If J>TotalPage Then Exit For
				    If J= CurrPage Then
				     PageStr=PageStr & " <font color=red>[" & J &"]</font>"& vbcrlf
				    Else
				     PageStr=PageStr & " <a href=""" & FileName & "-" & J&".Html"">[" & J &"]</a>"& vbcrlf
					End If
				 Next
				 If CurrPage=TotalPage Then
				  PageStr=PageStr & " 下一页 尾页"
				 Else
				  PageStr=PageStr & " <a href=""" & FileName & "-" & CurrPage + 1 & ".Html"" title=""下一页"">下一页</a> <a href=""" & FileName & "-" & TotalPage & ".Html"">尾页</a> "
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