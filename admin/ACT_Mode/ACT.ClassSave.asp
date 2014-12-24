<!--#include file="../ACT.Function.asp"-->
<!--#include file="../../act_inc/ACT.Code.asp"-->
<%
	Dim Save_Rs,ClassIDValue,Save_Rs1,GetParentID,Articleadd,EnName,ShowErr,Save_SQL,Save_SQL1,ClassID
	Dim ClassName,Extension,ClassKeywords,ClassDescription,ChangesLinkUrl 
	dim dh,tg,classename,ConTentTemplate,sitepath,moresite,siteurl,ACTlink
	Dim TGGroupID,OrderID,FolderTemplate,ModeID,FilePathName,content,makehtmlname,pageTemplate
	dim ClassPurview,ClassArrGroupID,ClassReadPoint,ClassChargeType ,ClassPitchTime,ClassReadTimes,ClassDividePercent
	Dim SEOtitle,ClassPicUrl,labelfor
	 ModeID = ChkNumeric(Request("ModeID"))
	 labelfor= ChkNumeric(Request("labelfor"))
	 if ModeID=0 or ModeID="" Then ModeID=1
		If Not ACTCMS.ACTCMS_QXYZ(ModeID,"","") Then   Call Actcms.Alert("对不起，您没有"&ACTCMS.ACT_C(ModeID,1)&"系该项操作权限！","")
	ClassPicUrl=Request.Form("ClassPicUrl")
	TGGroupID=Request.Form("TGGroupID")
	ClassID = Request.form("ClassID")
	SEOtitle=Request.form("SEOtitle")
	OrderID = ChkNumeric(Request.form("OrderID"))
	GetParentID = Request.form("ParentID")
	ClassName = Trim(Request.Form("ClassName"))
	Extension = Trim(Request.Form("Extension"))
	ClassDescription = Request.Form("ClassDescription")
	ClassKeywords = Request.Form("ClassKeywords")
	FolderTemplate = Request.Form("FolderTemplate")
	ChangesLinkUrl = Request.Form("ChangesLinkUrl")
	dh = ChkNumeric(Request.Form("dh"))
	tg = ChkNumeric(Request.Form("tg"))
	ACTlink = ChkNumeric(Request.Form("ACTlink"))

	FolderTemplate = Request.Form("FolderTemplate")
	ConTentTemplate = Request.Form("ConTentTemplate")
	FilePathName = ACTCMS.S("FilePathName")

 
	ClassPurview = ChkNumeric(Request.form("ClassPurview"))
	ClassReadPoint = ChkNumeric(Request.form("ClassReadPoint"))
	ClassChargeType = ChkNumeric(Request.form("ClassChargeType"))
	ClassPitchTime = ChkNumeric(Request.form("ClassPitchTime"))
	ClassReadTimes = ChkNumeric(Request.form("ClassReadTimes"))
	ClassDividePercent = ChkNumeric(Request.form("ClassDividePercent"))
	ClassArrGroupID = Request.form("ClassArrGroupID")
 
	content= ACTCMS.S("content")
	makehtmlname= ACTCMS.S("makehtmlname")
	pageTemplate= ACTCMS.S("pageTemplate")

	If ActLink="2"  Then  makehtmlname=ACTCMS.S("LinkUrl")
	If ActLink="3"  Then FolderTemplate=pageTemplate
'----------------------------------------------------------
	moresite = ChkNumeric(Request.Form("moresite"))
	siteurl = Request.Form("siteurl")
	sitepath = Request.Form("sitepath")
'----------------------------------------------------------
	If Trim(TGGroupID)="" Then TGGroupID=0
	IF Trim(ClassName) = ""  Then
		ShowErr = "请填写文章分类名称"
		Call Actcms.ActErr(ShowErr,"","1")
		Response.End
	End If

	IF Trim(FolderTemplate) = ""  Then
	If Not  ActLink="2"  Then 
		ShowErr = "栏目模板地址不能为空"
		Call Actcms.ActErr(ShowErr,"","1")
		Response.End
	  End if
	End If

	IF Trim(ConTentTemplate) = ""  Then
 		ShowErr = "内容页模板地址不能为空"
		Call Actcms.ActErr(ShowErr,"","1")
		Response.End
	 
	End If

	IF ACTCMS.s("ChangesLink") = "1"  Then 
		IF Trim(ChangesLinkUrl) = ""  Then
			ShowErr = "请填写转向链接地址"
			Call Actcms.ActErr(ShowErr,"","1")
			Response.End
		End If
	Else
		ChangesLinkUrl=""
	End IF
 	ClassIDValue = ACTCMS.MakeRandom(10)'随机生成15位字符
	Set Save_SQL = server.CreateObject("adodb.recordset")
 	
		Dim TemplateContent,makenames
		Dim MakePage,namearr,i,namearrs
		Set MakePage =New ACT_Code
	IF Request("Action") = "add" Then
		If  Request.Form("IFPinYin") = "1" Then
			EnName = ACTCMS.GetEn(ACTCMS.PinYin(ClassName))
		Else
			IF ACTCMS.Chkchars(Request.Form("EnName")) = False  Then
				ShowErr = "栏目目录只能为英文、数字及下划线"
				Call Actcms.ActErr(ShowErr,"","1")
				Response.End
			Else
				EnName = Trim(Request.Form("EnName"))
			End if
		End If
		
		
	
		Set Save_Rs = server.CreateObject("adodb.recordset")
		Save_Rs.Open "Select ID from Class_Act where ClassID='"& ClassIDValue &"' order by id desc",Conn,1,3
		if  Not Save_Rs.eof then
				ShowErr = "栏目ClassID意外出现重复，请重新输入"
				Call Actcms.ActErr(ShowErr,"","1")
				Response.end
		End if
		Set Save_Rs = nothing
		
		If ClassID <> "" Then
			Set Save_Rs = server.CreateObject("adodb.recordset")
			Save_Rs.Open "Select ID,ClassEname from Class_Act where ClassID='"& ClassID &"' order by id desc",Conn,1,3
			ClassEname = Save_Rs("ClassEname")&EnName&"/"
			ClassEname = request("onEnName")&EnName&"/"
 		Else 
			ClassEname = EnName&"/"
			Set Save_Rs = nothing
		End IF
		
		Set Save_Rs1 = server.CreateObject("adodb.recordset")
		
		If GetParentID="0" Then 
			Articleadd ="Select ID from Class_Act where EnName='"& trim(EnName) &"' And ParentID='0' and ModeID="&ModeID&""
			'添加根目录
		Else
			Articleadd ="Select ID from Class_Act where ParentID='"&GetParentID&"' and EnName='"& trim(EnName) &"' and ModeID="&ModeID&""
			'下级分类
		End If
		
		Save_Rs1.Open Articleadd,Conn,1,3
		if Not (Save_Rs1.eof and Save_Rs1.bof)  then
			if trim(request("ChangesLink"))<>"0" then
				ShowErr = "栏目英文名称重复，请重新输入"
				Call Actcms.ActErr(ShowErr,"","1")
				Response.end
			End  If 
		End if
		set Save_Rs1 = nothing
		Save_SQL1 = "Select * from Class_Act where 1=2"
		Save_SQL.Open Save_SQL1,Conn,1,3
		Save_SQL.AddNew
		Save_SQL("ClassName") = ClassName
		Save_SQL("ClassEName") = ClassEName
		Save_SQL("EnName") = EnName
		Save_SQL("ModeID") = ModeID
		Save_SQL("ClassID") = ClassIDValue
		Save_SQL("ParentID") = GetParentID
		Save_SQL("OrderID") = OrderID
		Save_SQL("Extension") = Extension
		Save_SQL("ClassKeywords") = ClassKeywords
		Save_SQL("ClassDescription") = ClassDescription
 		Save_SQL("ConTentTemplate") = ConTentTemplate
 		Save_SQL("dh") = dh
 		Save_SQL("SEOtitle") = SEOtitle
		Save_SQL("tg") = tg
		Save_SQL("ModeID") = ModeID
		Save_SQL("TGGroupID") = TGGroupID
		Save_SQL("moresite") = moresite
		Save_SQL("siteurl") = siteurl
		Save_SQL("sitepath") = sitepath
		Save_SQL("FilePathName")=FilePathName
		Save_SQL("content")=content
		Save_SQL("makehtmlname")=makehtmlname
		Save_SQL("ClassPurview")=ClassPurview
		Save_SQL("ClassReadPoint")=ClassReadPoint
		Save_SQL("ClassChargeType")=ClassChargeType
		Save_SQL("ClassPitchTime")=ClassPitchTime
		Save_SQL("ClassReadTimes")=ClassReadTimes
		Save_SQL("ClassDividePercent")=ClassDividePercent
		Save_SQL("ClassArrGroupID")=ClassArrGroupID
		Save_SQL("ClassPicUrl")=ClassPicUrl
		Save_SQL("labelfor")=labelfor
		
		if ActLink=1 then 
			Save_SQL("FolderTemplate")=FolderTemplate
		else  
			Save_SQL("FolderTemplate")=pageTemplate
		end if
 		Save_SQL("ACTlink")=ACTlink
		Save_SQL.update
  		Dim rs:Set rs=ACTCMS.actexe("Select top 1 id,classid from Class_Act order by id desc")
		If Not rs.eof Then ClassID = rs("ClassID")
   		If ActLink=3 Then 
 			makenames=makehtmlname
			Application(AcTCMSN & "ACTCMS_TCJ_Type")= "Folder"
			Application(AcTCMSN & "classid")=  ClassID
			Application(AcTCMSN & "modeid")= ACTCMS.ACT_L(ClassID,10)
 			TemplateContent= MakePage.LoadTemplate(pageTemplate)
		    TemplateContent=MakePage.Loadfile(TemplateContent)
			TemplateContent = MakePage.LabelReplaceAll(TemplateContent)
			TemplateContent=Replace(TemplateContent, "{$GetClassIntro}", content)
  			Call MakePage.FSOSaveFile(TemplateContent,actcms.GetPath(classid,makehtmlname))
	    End If 
   		Application.Contents.RemoveAll
 		Call Actcms.ActErr("栏目添加成功","ACT_Mode/ACT.Class.asp?ModeID="&ModeID&"","")
 	ElseIF  Request("Action") = "edit" Then
  		Save_SQL1 = "Select * from Class_Act where ClassID='"&ClassID &"'"
		Save_SQL.Open Save_SQL1,Conn,1,3
		Save_SQL("ClassName") = ClassName
		Save_SQL("Extension") = Extension
		Save_SQL("ClassKeywords") = ClassKeywords
		Save_SQL("ClassDescription") = ClassDescription
 		Save_SQL("ConTentTemplate") = ConTentTemplate
		Save_SQL("dh") = dh
 		Save_SQL("SEOtitle") = SEOtitle
		Save_SQL("tg") = tg
		Save_SQL("TGGroupID") = TGGroupID
		Save_SQL("OrderID") = OrderID
		Save_SQL("ModeID") = ModeID
		Save_SQL("moresite") = moresite
		Save_SQL("siteurl") = siteurl
		Save_SQL("sitepath") = sitepath
		Save_SQL("FilePathName")=FilePathName
 		Save_SQL("content")=content
		Save_SQL("makehtmlname")=makehtmlname
		Save_SQL("ClassPurview")=ClassPurview
		Save_SQL("ClassReadPoint")=ClassReadPoint
		Save_SQL("ClassChargeType")=ClassChargeType
		Save_SQL("ClassPitchTime")=ClassPitchTime
		Save_SQL("ClassReadTimes")=ClassReadTimes
		Save_SQL("ClassDividePercent")=ClassDividePercent
		Save_SQL("ClassArrGroupID")=ClassArrGroupID
		Save_SQL("ClassPicUrl")=ClassPicUrl
		Save_SQL("labelfor")=labelfor
		if ActLink=1 then 
			Save_SQL("FolderTemplate")=FolderTemplate
		else  
			Save_SQL("FolderTemplate")=pageTemplate
		end if
 		Save_SQL("ACTlink")=ACTlink
		If  ChkNumeric(Request.form("EditEname"))=1 Then 
  			If Right(Request.form("EnName"),1)<>"/" Then 
				Save_SQL("ClassEName")=request("EnName")&"/"
			Else 
				Save_SQL("ClassEName")=request("EnName")
			End If 
  		End if
 		ModeID=Save_SQL("ModeID") 
		Save_SQL.update
		ShowErr = "栏目保存成功"
	End If

	If ActLink=3 Then 
			Application(AcTCMSN & "ACTCMS_TCJ_Type")= "Folder"
			Application(AcTCMSN & "classid")=  ClassID
			Application(AcTCMSN & "modeid")= ACTCMS.ACT_L(ClassID,10)

			TemplateContent= MakePage.LoadTemplate(pageTemplate)
		    TemplateContent=MakePage.Loadfile(TemplateContent)
			TemplateContent = MakePage.LabelReplaceAll(TemplateContent)
			TemplateContent=Replace(TemplateContent, "{$GetClassIntro}", content)
 			Call MakePage.FSOSaveFile(TemplateContent,actcms.GetPath(classid,makehtmlname))
   	End If 
	
	Application.Contents.RemoveAll
	Save_SQL.close:set Save_SQL = Nothing
	Call Actcms.ActErr(ShowErr,"ACT_Mode/ACT.Class.asp?ModeID="&ModeID&"","")
   %>