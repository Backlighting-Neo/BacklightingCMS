<!--#include file="../../ACT.Function.asp"-->
<%
		If Not ACTCMS.ACTCMS_QXYZ(0,"bqxt","") Then   Call Actcms.Alert("对不起，你没有操作权限！","")
		Dim LabelContent,ShowErr
		Dim ID, LabelRS, SQLStr, LabelName, Descript,   FileUrl, LabelFlag,LabelType
		 FileUrl = Request("FileUrl") 
		Set LabelRS = Server.CreateObject("Adodb.RecordSet")
	Select Case Request.Form("Action")
		Case "Add"
			LabelName = Replace(Replace(Trim(Request.Form("LabelName")), """", ""), "'", "")
			Descript = Replace(Trim(Request.Form("Descript")), "'", "")
			LabelContent = Trim(Request.Form("LabelContent"))
			LabelFlag = Request.Form("LabelFlag")
			If LabelFlag = "" Then LabelFlag = 0
			If LabelName = "" Then
			   Response.Write ("<script>alert('标签名称不能为空!');location.href='" & FileUrl & "?Action=Add';</script>")
			   Response.End
			End If
			If LabelContent = "" Then
			   Response.Write("标签名称不能为空")
			  Response.End
			End If
			LabelName = "{ACTCMS_" & LabelName & "}"
			LabelRS.Open "Select LabelName From Label_Act Where LabelName='" & LabelName & "'", Conn, 1, 1
			If Not LabelRS.EOF Then
			  Response.Write ("<script>alert('标签名称已经存在!');location.href='" & FileUrl & "?Action=Add';</script>")
			  LabelRS.Close:Conn.Close:Set LabelRS = Nothing:Set Conn = Nothing
			  Response.End
			Else
				LabelRS.Close
				LabelRS.Open "Select * From Label_Act", Conn, 1, 3
				LabelRS.AddNew
				 LabelRS("LabelName") = LabelName
				 LabelRS("Description") = Descript
				 LabelRS("LabelContent") = LabelContent
				 LabelRS("LabelFlag") = LabelFlag
				 LabelRS("AddDate") = Now
				 LabelRS("LabelType") = 1 '指定为系统函数标签
				 LabelRS.Update
				 Application.Contents.RemoveAll
				Call Actcms.ActErr("添加标签成功","Label_Admin.asp?Type=1","")
			End If
		Case "Edit"
			
			ID = Trim(Request.Form("ID"))
			LabelName = Replace(Replace(Trim(Request.Form("LabelName")), """", ""), "'", "")
			Descript = Replace(Trim(Request.Form("Descript")), "'", "")
			LabelContent = Trim(Request.Form("LabelContent"))
			LabelFlag = Request.Form("LabelFlag")
			LabelType =  Request.Form("LabelType")
			'If LabelFlag = "" Then LabelFlag = 1
			If LabelName = "" Then
			   Response.Write "null"
			   Response.End
			End If
			If LabelContent = "" Then
			  Response.Write "null"
			  Response.End
			End If
			LabelName = "{ACTCMS_" & LabelName & "}"
			LabelRS.Open "Select LabelName From [Label_Act] Where ID <>" & ID & " AND LabelName='" & LabelName & "'", Conn, 1, 1
			If Not LabelRS.EOF Then
			  Response.Write ("<script>alert('标签名称已经存在!');location.href='" & FileUrl & "?Action=Edit&ID=" & ID & "';</script>")
			  LabelRS.Close
			  Conn.Close
			  Set LabelRS = Nothing
			  Set Conn = Nothing
			  Response.End
			Else
				LabelRS.Close
				LabelRS.Open "Select * From [Label_Act] Where ID=" & ID & "", Conn, 1, 3
				 LabelRS("LabelName") = LabelName
				 LabelRS("Description") = Descript
				 LabelRS("LabelContent") = LabelContent
				 LabelRS("LabelFlag") = LabelFlag
				 LabelRS("AddDate") = Now
				 LabelRS.Update
				 Application.Contents.RemoveAll
				Call Actcms.ActErr("标签修改成功","Label_Admin.asp?Type=1","")
 			End If
	End Select
  %>