<!--#include file="../Conn.asp"-->
<!--#include file="../act_inc/ACT.Main.asp"-->
<!--#include file="../act_inc/ACT.Code.asp"-->
<!--#include file="../act_inc/CreateView.asp"-->
<!--#include file="../act_inc/UpLoadClass.asp"-->
 <% 	
	Dim ACTCMS,UserID
	Set ACTCMS = New ACT_Main
	Public Function UserLoginChecked()
	on error resume next
	Dim UserName,UserPassword
		UserLoginChecked = false 
		UserName = RSQL(actcms.toasc(actcms.s("U")))
		UserPassword= RSQL(actcms.toasc(actcms.s("P")))
		IF UserName="" Or UserPassword = "" Then
		   UserLoginChecked=false
		   Exit Function
		Else
			Dim UserRs
			Set Userrs=Actcms.Actexe("Select UserID,UserName,PassWord From User_ACT Where UserName='" & UserName & "' And PassWord='" & UserPassword & "'")
			IF UserRS.Eof And UserRS.Bof Then
				UserLoginChecked=false
			Else
				UserLoginChecked = True
				UserID=userrs("UserID")
			End if
			UserRS.Close:Set UserRS=Nothing
	   End IF
	End Function 
	IF Cbool(UserLoginChecked)=false Then
	  Response.Write "error"
	  Response.end
	End If
 Dim ModeID,Yname,instrs,myid,fp
 Yname=request("Yname")
 myid=Request("myid")
ModeID = ChkNumeric(Request("ModeID"))
If  ModeID=0 or ModeID="" Then ModeID=1
 IF myid="Upload" Then
	If ModeID="999" Then fp=ACTCMS.ActSys&"UpFiles/UserFile" &UserID&"/"  Else fp=ACTCMS.ActSys&ACTCMS.ACT_C(ModeID,8)& "UserFile/"&UserID&"/" 
  	 Call actcms.CreateFolder(fp)
	Dim UpFile
	set UpFile = New UpLoadClass
  	UpFile.AutoSave = 2
	UpFile.MaxSize =  ACTCMS.ActCMS_Sys(10)* 1024
	UpFile.FileType = ACTCMS.ActCMS_Sys(11)
	UpFile.SavePath = fp
	UpFile.Open() '# 打开对象
  	If UpFile.Save("Filedata",0) Then
 		Dim W:Set W = New CreateView
		Call  W.SY(ACTCMS.ActSys&fp&UpFile.Form("Filedata"),UpFile.Form("Filedata_Ext"))
		Call OutUploadScript(UpFile.Form("Filedata_Ext"),actcms.PathDoMain&fp&UpFile.Form("Filedata"),Yname)
	End If
 	Set UpFile = Nothing
End If 

Sub OutUploadScript(sType,strPath,instrct)
	sType = LCase(sType)
	Dim Temps
	If Yname<>"content"  Then 
 		Temps=ACTCMS.LTemplate(actcms.actsys&"Act_inc/T/"&ModeID&"/"&sType&".Html")
		If temps="" Then Temps=ACTCMS.LTemplate(actcms.actsys&"Act_inc/Act_inc/T/1/"&sType&".Html")
	Else
		Temps=ACTCMS.LTemplate(actcms.actsys&"Act_inc/T/"&ModeID&"/C"&sType&".Html")
		If temps="" Then Temps=ACTCMS.LTemplate(actcms.actsys&"Act_inc/T/1/C"&sType&".Html")
	End If 
	If Trim(Temps)="" Then Temps=ACTCMS.LTemplate(actcms.actsys&"Act_inc/T/1/actcms.Html")

	If InStr(Temps, "#FileName") > 0  Then
 		   Temps = Replace(Temps, "#FileName",strPath)
 	End if
	
	If InStr(Temps, "{$InstallDir}") > 0  Then
 		   Temps = Replace(Temps, "{$InstallDir}",AcTCMS.ActCMS_Sys(3))
 	End if
 	Response.Write instrct&"|"&Temps & vbCrLf
End Sub



 %>