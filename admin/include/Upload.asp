<!--#include file="../../Conn.asp"-->
<!--#include file="../../ACT_inc/ACT.Main.asp"-->
<!--#include file="../../ACT_inc/UpLoadClass.asp"-->
<!--#include file="../../ACT_inc/CreateView.asp"-->
 <% 	
	Dim ACTCMS
	Set ACTCMS = New ACT_Main
	Public Function UserLoginChecked()
	on error resume next
	Dim AdminName,PassWord
		UserLoginChecked = false 
   		AdminName = RSQL(actcms.toasc(actcms.s("U")))
		PassWord= RSQL(actcms.toasc(actcms.s("P")))
		IF AdminName="" Or PassWord = "" Then
		   UserLoginChecked=false
		   Exit Function
		Else
			Dim UserRs
			Set Userrs=Actcms.Actexe("Select Admin_Name,PassWord From Admin_ACT Where Admin_Name='" & AdminName & "' And PassWord='" & PassWord & "'")
			IF UserRS.Eof And UserRS.Bof Then
				UserLoginChecked=false
			Else
				UserLoginChecked = true
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
	If ModeID="999" Then fp=ACTCMS.ActSys&"UpFiles/UserFile/Other/" Else fp=ACTCMS.ActSys&ACTCMS.ACT_C(ModeID,8)&year(now) & "/" & month(now)& "/" & Day(now)&"/"
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
 		Temps=ACTCMS.LTemplate("../../Act_inc/T/"&ModeID&"/"&sType&".Html")
		If temps="" Then Temps=ACTCMS.LTemplate("../../Act_inc/T/1/"&sType&".Html")
 	Else
		Temps=ACTCMS.LTemplate("../../Act_inc/T/"&ModeID&"/C"&sType&".Html")
		If temps="" Then Temps=ACTCMS.LTemplate("../../Act_inc/T/1/C"&sType&".Html")
	End If 
	If Trim(Temps)="" Then Temps=ACTCMS.LTemplate("../../Act_inc/T/1/actcms.Html")


	If InStr(Temps, "#FileName") > 0  Then
 		   Temps = Replace(Temps, "#FileName",strPath)
 	End if
	
	If InStr(Temps, "{$InstallDir}") > 0  Then
 		   Temps = Replace(Temps, "{$InstallDir}",AcTCMS.ActCMS_Sys(3))
 	End if
 	Response.Write instrct&"|"&Temps & vbCrLf
End Sub



 %>