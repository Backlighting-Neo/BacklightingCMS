<!--#include file="../Conn.asp"-->
<!--#include file="ACT.Main.asp"-->
<!--#include file="ACT.Code.asp"-->
<%
	Dim ACTCMS
	Set ACTCMS = New ACT_Main
	Class ACT_User
	Public Edays,GroupID,Point,UserName,PassWord,Email,Locked,Money,Score,LoginNumber,Realname,sex
	Public Birthday,Privacy,Mobile,HomeTel,Fax,Province,City,QQ,msn,Question,Answer,address,postcode
	Public UserID,UserMID,UModeID,G_SH,G_Name,G_A_SH,G_UserUpFilesTF,myface,ChargeType,BeginDate,GetEdays
	Public G_WriteComment,G_Path,G_UpFileType,G_UpfilesSize,G_Max_sEnd,G_Max_sms,G_Max_Num,G_Simple,G_tgdianshu
		Private Sub Class_Initialize()
		End Sub
		Private Sub Class_Terminate()
		 Set ACTCMS=Nothing
		End Sub
		Public Function UserLoginChecked()
		on error resume next
		Dim UserModeID,Ruser
 			UserLoginChecked = false 
			UserName = RSQL(Trim(Request.Cookies(AcTCMSN)("UserName")))
			PassWord= RSQL(Trim(Request.Cookies(AcTCMSN)("PassWord")))
 			if ACTCMS.ActCMS_Sys(12) = 1   Then UserLoginChecked = false:Exit Function:Response.End
 			Set Ruser=ACTCMS.actexe("Select ModeID from Group_Act where GroupID="&RSQL(ChkNumeric((Request.Cookies(AcTCMSN)("GroupID"))))&"")
				If Not Ruser.eof Then UserMID=Ruser("ModeID"):Ruser.Close:Set Ruser=Nothing
				IF UserName="" Then
				   UserLoginChecked=false
				   Exit Function
				Else
					Dim UserRs:Set UserRS=Server.CreateOBject("ADODB.RECORDSET")
					UserRS.Open "Select * From User_act Where UserName='" & UserName & "' And PassWord='" & PassWord & "'",Conn,2,3
					IF UserRS.Eof And UserRS.Bof Then
					  UserLoginChecked=false
					Else
						UserLoginChecked = True
						UserID=UserRS("UserID")
						GroupID = UserRS("GroupID")
						Money = UserRS("Money")
						ChargeType = UserRS("ChargeType")
						BeginDate = UserRS("BeginDate")
						Edays = UserRS("Edays")
						Point = UserRS("Point")
 				
						Realname = UserRS("Realname")
						Mobile = UserRS("Mobile")
						address = UserRS("address")
						QQ = UserRS("QQ")
						MSN = UserRS("MSN")
						HomeTel = UserRS("HomeTel")
						postcode = UserRS("postcode")
						Mobile = UserRS("Mobile")
						myface = UserRS("myface")
 						Birthday = UserRs("Birthday")
						sex = UserRs("sex"):Realname = UserRs("Realname")
						Score = UserRs("Score"):Email=UserRs("Email"):Locked=UserRs("Locked")
						LoginNumber=UserRs("LoginNumber"):Privacy=UserRs("Privacy")
						Province=UserRs("Province"):City=UserRs("City")
						UModeID=UserRs("UModeID") 
						GetEdays = Edays-DateDiff("D",BeginDate,now())
						  Dim Rs,GroupArr
						  Set Rs=ACTCMS.ACTEXE("Select ModeID,GroupSetting,GroupName from Group_Act  Where GroupID=" & GroupID & " order by GroupID desc")
						  If Not Rs.Eof Then
							GroupArr=Split(Rs("GroupSetting"),"^@$@^")
 							G_Name = rs("GroupName")
							G_WriteComment=GroupArr(2)
							G_SH=GroupArr(3)
							G_Max_Num=GroupArr(4)
							G_Max_sms=GroupArr(5)
							G_Max_sEnd=GroupArr(6)
							G_UserUpFilesTF=GroupArr(7)
							G_Path=GroupArr(8)
							G_UpFileType=GroupArr(10)
							G_UpfilesSize=GroupArr(9)
							G_A_SH=GroupArr(11)
							G_Simple= GroupArr(13)
							G_tgdianshu= GroupArr(15)
						  End If
						  Rs.Close:Set Rs=Nothing
					End if
					UserRS.Close:Set UserRS=Nothing
			   End IF
		End Function 

	 
End  Class 
%>