<%
Class ACT_Main
	Private LocalCacheName, Cache_Data,CacheData
	Public Reloadtime,Version,PathDoMain,dic,Index
	Public ActCMS_Sys,ActCMS_User,ActCMS_Other,ActCMSDM,ActCMSUpfile
  	Private Sub Class_Initialize()
 		Version="4.0 20110521"
 		Reloadtime = 28800
		Call GetConfig()
		ActCMS_Sys=Split(CacheData(0,0),"^@$@^")
		ActCMS_Other=Split(CacheData(1,0),"^@&@^")
		ActCMSUpfile=Split(CacheData(2,0),"^@*&*@^")
 		If ActCMS_Other(9)="0" Then 
			ActCMSDM=Trim(ActCMS_Sys(2) & ActCMS_Sys(3))
			PathDoMain=ActCMS_Sys(2)
		Else 
			ActCMSDM=Trim(ActCMS_Sys(3))
		End If 
		Index=0:Set  dic=server.CreateObject("scripting.dictionary")
 	End Sub
 	Function iCreateObject(str)
		'iis5创建对象方法Server.CreateObject(ObjectName);
		'iis6创建对象方法CreateObject(ObjectName);
		'默认为iis6，如果在iis5中使用，需要改为Server.CreateObject(str);
		Set iCreateObject=CreateObject(str)
	End Function
 	Private Sub Class_Terminate()
		If IsObject(Conn) Then Conn.Close : Set Conn = Nothing
		Call CloseConn()
	End Sub
 	Public Function ACTExe(Command)
		If Not IsObject(Conn) Then ConnectionDatabase	
			on error resume next
			Set ACTExe = Conn.Execute(Command)
			If Err Then
				Set Conn = Nothing
				Response.Write err.description
				Response.Write "<li>查询数据的时候发现错误，请检查您的查询代码是否正确。<br /><li>"
				Response.Write Command
				Response.End
			End If
    End Function
 	Public Property Let Name(ByVal vNewValue)
		LocalCacheName = LCase(vNewValue)
		Cache_Data = Application(AcTCMSN & "_" & LocalCacheName)
	End Property
	Public Property Let Value(ByVal vNewValue)
		If LocalCacheName <> "" Then
			ReDim Cache_Data(2)
			Cache_Data(0) = vNewValue
			Cache_Data(1) = Now()
			Application.Lock
			Application(AcTCMSN & "_" & LocalCacheName) = Cache_Data
			Application.UnLock
		End If
	End Property
	Public Property Get Value()
		If LocalCacheName <> "" Then
			If IsArray(Cache_Data) Then
				Value = Cache_Data(0)
			End If
		End If
	End Property
	Public Function ObjIsEmpty()
		ObjIsEmpty = True
		If Not IsArray(Cache_Data) Then Exit Function
		If Not IsDate(Cache_Data(1)) Then Exit Function
		If DateDiff("s", CDate(Cache_Data(1)), Now()) < (60 * Reloadtime) Then ObjIsEmpty = False
	End Function
	Public Sub DelCahe(MyCaheName)
		Application.Lock
		Application.Contents.Remove (AcTCMSN & "_" &MyCaheName)
		Application.UnLock
	End Sub
 	Public Function GetConfig()'第一次起用系统或者重启IIS的时候加载缓存
		Name = "Config"
		If ObjIsEmpty() Then ReloadConfig
		CacheData = Value
		Name = "Date"
		If ObjIsEmpty() Then
			Value = Date
		Else
			If CStr(Value) <> CStr(Date) Then
				Name = "Config"
				Call ReloadConfig
				CacheData = Value
			End If
		End If
		If Len(CacheData(1, 0)) = 0 Then
			Name = "Config"
			Call ReloadConfig
			CacheData = value
		End If
		End Function

	    Public Sub ReloadConfig()
		   Dim RS
		   Set Rs = ACTExe("SELECT  Top 1 ActCMS_SysSetting,ActCMS_OtherSetting,ActCMS_Upfile,ActCMS_Theme  from [Config_act]")
		   value=RS.GetRows(1)
		   Set RS=Nothing
		End Sub
 	Public Function GetRandomize(CMS_number)'随机字符串
		Randomize
		Dim CMS_Randchar,CMS_Randchararr,CMS_RandLen,CMS_Randomizecode,CMS_iR
		CMS_Randchar="0,1,2,3,4,5,6,7,8,9,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z"
		CMS_Randchararr=split(CMS_Randchar,",") 
		CMS_RandLen=CMS_number 
		For CMS_iR=1 to CMS_RandLen
			CMS_Randomizecode=CMS_Randomizecode&CMS_Randchararr(Int((21*Rnd)))
		Next 
		GetRandomize = CMS_Randomizecode
	End Function

    Public Function Chkchars(Chars)'检测英文名称是否合法
		Dim Charname, i, c
		Charname = Chars
		Chkchars = True
		If Len(Charname) <= 0 Then
			Chkchars = False
			Exit Function
		End If
		For i = 1 To Len(Charname)
		   C = Mid(Charname, i, 1)
			If InStr("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ@,.0123456789|-_", c) <= 0  Then
			   Chkchars = False
			Exit Function
		   End If
	   Next
	End Function
	
	Function GetXMLFromFile(FileName)
		If Not IsObject(Application(AcTCMSN&"_Config"&FileName)) Then
		  Dim objXmlFile:set objXmlFile = iCreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		  objXmlFile.async = false
		  objXmlFile.setProperty "ServerHTTPRequest", true 
		  objXmlFile.load(FileName)
		  Set Application(AcTCMSN&"_Config"&FileName)=objXmlFile
	   End If  
		Set GetXMLFromFile=Application(AcTCMSN&"_Config"&FileName)
	End Function
	Function NoAppGetXMLFromFile(FileName)
		  Dim objXmlFile:set objXmlFile = iCreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		  objXmlFile.async = false
		  objXmlFile.setProperty "ServerHTTPRequest", true 
		  objXmlFile.load(FileName)
		 Set NoAppGetXMLFromFile=objXmlFile
	End Function

	 Public Function NowTheme()
 	    Name="NowTheme"
	    If ObjIsEmpty() Then 
		   Dim rs:Set Rs = ACTExe("SELECT  Top 1 ActCMS_Theme  from [Config_act]")
			If Not Rs.eof Then 
				Value=Rs(0)
				NowTheme=Rs(0)
				Rs.close:set Rs=nothing
			End If 
		Else
			NowTheme=Value
		End If 
	 End Function 

	Public Function ActUrl()
	   ActUrl = Trim(ActCMS_Sys(2) & ActCMS_Sys(3))
	End Function

 	Public Function ActSys()
	   ActSys = Trim(ActCMS_Sys(3))
	End Function

 	Public Function adminurl()
	   adminurl = Trim(ActCMS_Sys(8))
	End Function
	
 	Public Function adminPath()
	   adminPath = Trim(ActCMS_Sys(3))&Trim(ActCMS_Sys(8))&"/"
	End Function
	
	Public Function CheckPlugin(str)
	Dim rs:Set rs=actexe("select * from Plus_ACT where   PlusID='"&rsql(str)&"'")
	If   rs.eof Then CheckPlugin=False :Exit Function 
	If rs("IsUse")=1 Then 
		CheckPlugin=False 
	Else 
		CheckPlugin=True 
	End If 
	End Function 
	
	Public Function ThisTheme
		ThisTheme=actcms.ActSys&actcms.ActCMS_Sys(19)&"/"&NowTheme
	End Function 
 	Public Function SiteUrl()
	   SiteUrl = Trim(ActCMS_Sys(2))
	End Function
 	Public Function SiteName()
	   SiteName = ActCMS_Sys(0)
	End Function

 	Public Function FsoName()
	   FsoName = ActCMS_Other(10)
	End Function

 	Public Function SysThemePath()
	   SysThemePath = Trim(actcms.ActCMS_Sys(19))
	End Function


 	Public Function SysPlusPath()
	   SysPlusPath = "plugin"
	End Function

 	Public Function AEXE(UB,arr)
		On Error Resume Next
 		Dim I,hs
		For I=1 To ub
				hs=hs&""""&arr(i)&""""&","
 		Next 
 		hs=Left(hs, Len(hs) - 1)
		hs=Replace(hs,"}","")
 		execute(LTemplate("/api/"&arr(0)&".inc"))
 		execute("call "&arr(0)&"("&hs&")")
    	If Err Then 
		    AEXE="<font color=red>api函数执行失败,错误原因 -> " & Err.Description & "</font>"
 		    Err.Clear
		Else 
			AEXE= actcool
	    End If 
	End Function 
	
	Public Function AField(UB)
  		  execute("call "&UB&"()")
    	  AField= actField
	End Function 

	Function regexField(ByVal Str, ByVal Pattern)
		If trim(Str)="" Then regexField = False : Exit Function
		Dim Re,Pa
		Set Re = New RegExp
		Re.IgnoreCase = True
		Re.Global = True
		Pa = Pattern'正则代码
		Re.Pattern = Pa
		regexField = Re.Test(CStr(Str))
		Set Re = Nothing
	End Function
	 
	 Public Function SysCount(ModeID)'统计模型文章总数
		Dim CountValue
	    Name="SysCount"&ModeID
	    If ObjIsEmpty() Then 
			Set CountValue=ACTEXE("Select Count(id)  From "&ACT_C(ModeID,2)&"")
			If Not CountValue.eof Then 
				Value=CountValue(0)
				SysCount=CountValue(0)
				CountValue.close:set CountValue=nothing
			End If 
		Else
			SysCount=Value
		End If 
	 End Function 

	 Public Function TodayRenewal(ModeID)'统计模型文章今日更新
	   Dim TodayValue
		Set TodayValue=ACTEXE("Select Count(id)  From "&ACT_C(ModeID,2)&" where DateDiff(""d"",UpdateTime," & NowString & ")=0")
		If Not TodayValue.eof Then  
			TodayRenewal=TodayValue(0)
			TodayValue.close:set TodayValue=nothing
		End If 
	 End Function 

	 Public Function CountClass(ClassID)'统计模型文章今日更新
	   Dim ClassValue
	   Name="CountClass"&ClassID
	   If ObjIsEmpty() Then 
			Set ClassValue=ACTEXE("Select Count(id)  From "&ACT_C(ACT_L(ClassID,10),2)&" where classid='"&ClassID&"'")
			If Not ClassValue.eof Then 
				Value=ClassValue(0)
				CountClass=ClassValue(0)
				ClassValue.close:set ClassValue=nothing
			End If 
		Else
			CountClass=Value
		End If 
	 End Function 

	Public Function ChkAdmin()'检测是否超级管理员
		ChkAdmin = False
		If Request.Cookies(AcTCMSN)("AdminName") = "" Then
			ChkAdmin = False
			Exit Function
		ElseIf Request.Cookies(AcTCMSN)("SuperTF") = "1" Then 
			ChkAdmin = True
			Exit Function
		End If 
	End Function 


	Public Function ACTCMS_QXYZ(ModeID,QXLX,ClassID)'权限验证
			ACTCMS_QXYZ = False
		If Request.Cookies(AcTCMSN)("AdminName") = "" Then
			ACTCMS_QXYZ = False
			Exit Function
		ElseIf Request.Cookies(AcTCMSN)("SuperTF") = "1" Then 
			ACTCMS_QXYZ = True
			Exit Function
		Else 
			If ModeID=0 Then '模型ID=0将进行插件权限检测
				If Instr(Request.Cookies(AcTCMSN)("ACT_Other"),QXLX) >0 Then 
					ACTCMS_QXYZ=True
				Else
					ACTCMS_QXYZ=False 
				End If 
			Else'模块相关权限检测
				If Instr(Request.Cookies(AcTCMSN)("Purview"),"ACT"&ModeID&"-ACT") >0 Then 
					ACTCMS_QXYZ=False 
				ElseIf  Instr(Request.Cookies(AcTCMSN)("Purview"),"TCJ"&ModeID&"-TCJ") >0 Then 
					ACTCMS_QXYZ=True 
				Else 
					If Trim(Classid) ="" Then ACTCMS_QXYZ = False:Exit Function
					ACTCMS_QXYZ=ACTCMS_HQQX(ClassID,QXLX)	
				End If 
			End If 
		End If 
	End Function 

	Public Function ACTCMS_HQQX(HQQXID,HQACT)
		Dim HQarrTemp,HQi,HQL,HQACT_ClassID
		HQarrTemp=split(session("HQQXLX"),",")'
		For HQI=LBound(HQarrTemp) To Ubound(HQarrTemp)'遍历
			if InStr(HQarrTemp(HQI),HQQXID) > 0 Then
				HQACT_ClassID=Split(HQarrTemp(HQI),"-")
				If UBound(HQACT_ClassID)>0 Then 
					If HQACT_ClassID(1)=HQACT Then
						ACTCMS_HQQX=True
						Exit Function
					Else	
						ACTCMS_HQQX=False
					End If 
				End if
			End  If 
		Next 
	End Function
	Function  strToAsc(strValue)
	 Dim  strTemp,i
 	 strTemp=""
	 for i=1 to len(strValue & "")
	 If session.codepage="65001" Then 
		  strTemp=strTemp & ascw(mid(strValue,i,1))&"_"
	  Else 
		  strTemp=strTemp & asc(mid(strValue,i,1))&"_"
	  End If 
	  Next 
	  strToAsc=strTemp
	End  Function  
	 Function toasc(strValue)
		Dim ThisAr,i
		ThisAr=split(strValue,"_") 
		for i=0 to Ubound(ThisAr) 
		if IsNumeric(ThisAr(i)) Then
		  If session.codepage="65001" Then 
			toasc=toasc&chrw(ThisAr(i)) 
		  Else
			toasc=toasc&chr(ThisAr(i)) 
		   End If 
		end if
		next 
	End Function 
	'参  数：RelativePath 数据库连接字段串
	'*********************************************************************************************************
	Function GetAbsolutePath(RelativePath)
		dim Exp_Path,Matches,tempStr
		tempStr=Replace(RelativePath,"\","/")
		if instr(tempStr,":/")>0 then
			GetAbsolutePath=RelativePath
			Exit Function
		End if
		set Exp_Path=new RegExp
		Exp_Path.Pattern="(Data str=|dbq=)(.)*"
		Exp_Path.IgnoreCase=true
		Exp_Path.Global=true
		Set Matches=Exp_Path.Execute(tempStr)
		If instr(LCase(tempStr),"*.xls")<>0 Then
		GetAbsolutePath="driver={microsoft excel driver (*.xls)};dbq="&Server.MapPath(split(Matches(0).value,"=")(1))
		ElseIf Instr(Lcase(tempstr),"*.dbf")<>0 Then
		GetAbsolutePath="driver={microsoft dbase driver (*.dbf)};dbq="&Server.MapPath(split(Matches(0).value,"=")(1))
		Else
		GetAbsolutePath="Provider=Microsoft.Jet.OLEDB.4.0;Data str="&Server.MapPath(split(Matches(0).value,"=")(1))
		End If
	End Function

	Sub InsertLog(UserName,ACT,ACTError,GetHttp)
		Dim sqlLog, rsLog
		sqlLog = "Select  * from Log_ACT where 1=1"
		Set rsLog = Server.CreateObject("Adodb.RecordSet")
		rsLog.Open sqlLog, Conn, 1, 3
		rsLog.AddNew
		rsLog("UserName") = UserName
		rsLog("ACT") = ACT'1 登陆 ,2 会员操作 ,3 ....
		rsLog("Times") = Now()
		rsLog("GetHttp") = GetHttp
		rsLog("LoginIP") = GetIP()
		rsLog("ACTError") = ACTError
		rsLog.Update
		rsLog.Close:Set rsLog = Nothing
	End Sub


	Sub ACTCMSErr(Url)
	   If Url = "" Then
		 Response.Write ("<script>alert('错误提示:\n\n你没有此项操作的权限,请与系统管理员联系!');history.back();</script>")
	   Else
	    Response.Write ("<script>alert('错误提示:\n\n你没有此项操作的权限,请与系统管理员联系!');location.href='" & Url & "';</script>")
	   End If
	   Response.end
	End Sub
	Public Function IsValidEmail(Email)
		Dim names, name, I, c
		IsValidEmail = True
		names = Split(Email, "@")
		If UBound(names) <> 1 Then IsValidEmail = False: Exit Function
		For Each name In names
			If Len(name) <= 0 Then IsValidEmail = False:Exit Function
			For I = 1 To Len(name)
				c = LCase(Mid(name, I, 1))
				If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then IsValidEmail = False:Exit Function
		   Next
		   If Left(name, 1) = "." Or Right(name, 1) = "." Then IsValidEmail = False:Exit Function
		Next
		If InStr(names(1), ".") <= 0 Then IsValidEmail = False:Exit Function
		I = Len(names(1)) - InStrRev(names(1), ".")
		If I <> 2 And I <> 3 Then IsValidEmail = False:Exit Function
		If InStr(Email, "..") > 0 Then IsValidEmail = False
	End Function
	'检查一个数组中所有元素是否包含指定字符串
	Public Function FoundInArr(strArr, strToFind, strSplit)
		Dim arrTemp, i
		FoundInArr = False
		If InStr(strArr, strSplit) > 0 Then
			arrTemp = Split(strArr, strSplit)
			For i = 0 To UBound(arrTemp)
			If LCase(Trim(arrTemp(i))) = LCase(Trim(strToFind)) Then
				FoundInArr = True:Exit For
			End If
			Next
		Else
			If LCase(Trim(strArr)) = LCase(Trim(strToFind)) Then FoundInArr = True
		End If
	End Function


	 Public Function RecordsetToxml(RSObj,row,xmlroot)'该函数参考动网
	  Dim i,node,rs,j,DataArray
	  If xmlroot="" Then xmlroot="xml"
	  If row="" Then row="row"
	  Set RecordsetToxml=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	  RecordsetToxml.appendChild(RecordsetToxml.createElement(xmlroot))
	  If Not RSObj.EOF Then
	   DataArray=RSObj.GetRows(-1)
	   For i=0 To UBound(DataArray,2)
		Set Node=RecordsetToxml.createNode(1,row,"")
		j=0
		For Each rs in RSObj.Fields		   
		   node.attributes.setNamedItem(RecordsetToxml.createNode(2,"ACT"&j,"")).text= DataArray(j,i)& ""
		   j=j+1
		Next
		RecordsetToxml.documentElement.appendChild(Node)
	   Next
	  End If
	  DataArray=Null
	 End Function

	 Function ACT_C(ModeID,RowID)'20
	  on error resume next
	  If not IsObject(Application(ACTCMSN &"_ModeConfig")) Then
		 Application.Lock
		 Dim RS:Set Rs=ACTEXE("select ModeID,ModeName,ModeTable,IFmake,adminmb,ProjectUnit,MakeFolderDir,RecyleIF,UpFilesDir,RefreshFlag,CommentTemp,ContentExtension,AutoPage,CommentCode,Commentsize,WriteComment,usermb From Mode_Act Order by ModeID")
		 Set Application(ACTCMSN &"_ModeConfig")=RecordsetToxml(rs,"Mode","ModeConfig")
		 Set Rs=Nothing
		 Application.unLock
	  End If
		 ACT_C=Application(ACTCMSN &"_ModeConfig").documentElement.selectSingleNode("Mode[@ACT0=" & ModeID & "]/@ACT" & RowID & "").text
	     if err then ACT_C=0:err.Clear
	 End Function
	
	 Function GetTable(ModeTable)
		 Dim RS:Set Rs=ACTEXE("select top 1 ModeID From Mode_Act where ModeTable='"&ModeTable&"' Order by ModeID")
		 If rs.eof Then GetTable="0":Exit Function 
		 GetTable=rs("ModeID"): Set Rs=Nothing
 	 End Function 

	 Function ACT_U(ModeID,RowID)
	  on error resume next
	  If not IsObject(Application(ACTCMSN &"_UserConfig")) Then
		 Application.Lock
		 Dim RS:Set Rs=ACTEXE("select ModeID,ModeName,ModeTable,Template,RegCode,SpaceID  From ModeUser_Act Order by ModeID")
		 Set Application(ACTCMSN &"_UserConfig")=RecordsetToxml(rs,"User","UserConfig")
		 Set Rs=Nothing
		 Application.unLock
	 End If
	     ACT_U=Application(ACTCMSN &"_UserConfig").documentElement.selectSingleNode("User[@ACT0=" & ModeID & "]/@ACT" & RowID & "").text
		 if err then ACT_U="User_ACT":err.Clear
	 End Function


	 Function ACT_L(ClassID,RowID)'18
	  on error resume next
	  If not IsObject(Application(ACTCMSN &"_ClassConfig")) Then
		 Application.Lock
		 Dim RS:Set Rs=ACTEXE("select ClassID,enname,ClassName,ClassEName,FolderTemplate,ConTentTemplate,ClassArrGroupID,Extension,ClassKeywords,ClassDescription,ModeID,ParentID,ActLink,moresite,sitepath,siteurl,FilePathName,makehtmlname,content,ClassPurview,ClassReadPoint,ClassChargeType,ClassPitchTime,ClassReadTimes,ClassDividePercent,seotitle,ClassPicUrl From Class_Act Order by OrderID")
		 Set Application(ACTCMSN &"_ClassConfig")=RecordsetToxml(rs,"Class","ClassConfig")
		 Set Rs=Nothing
		 Application.unLock
	 End If
	     ACT_L=Application(ACTCMSN &"_ClassConfig").documentElement.selectSingleNode("Class[@ACT0=" & ClassID & "]/@ACT" & RowID & "").text
		 if err then ACT_L=0:err.Clear
	 End Function



	 Function ACT_G(GID,RowID)'18
	  on error resume next
	  If not IsObject(Application(ACTCMSN &"_UserGroup")) Then
		 Application.Lock
		 Dim RS:Set Rs=ACTEXE("select GroupID,DefaultGroup,Description,ChargeType,GroupPoint,ValidDays,GroupSetting,GroupName,ModeID From Group_Act Order by GroupID")
		 Set Application(ACTCMSN &"_UserGroup")=RecordsetToxml(rs,"Group","UserGroup")
		 Set Rs=Nothing
		 Application.unLock
	 End If
	     ACT_G=Application(ACTCMSN &"_UserGroup").documentElement.selectSingleNode("Group[@ACT0=" & GID & "]/@ACT" & RowID & "").text
		 if err then ACT_G=0:err.Clear
	 End Function




   Function Act_MX_Arr(ModeID,A)'返回模型数组
	  Dim Rs
	  Set Rs=ACTEXE("Select FieldName,Title,IsNotNull,FieldType,[check],regex,regError from Table_ACT  Where ModeID=" & ModeID & " and actcms="&A&" order by OrderID desc,ID Desc")
	 If Not Rs.Eof Then
	  Act_MX_Arr=Rs.GetRows(-1)
	 Else
	  Act_MX_Arr=""
	 End If
	 Rs.Close:Set Rs=Nothing
   End Function



   Function Act_MX_Sys_Arr()'返回系统模型数组
	 Dim Rs
	  Set Rs =ACTEXE("SELECT ModeID, ModeName,ModeTable, ModeStatus, IFmake,ModeNote  FROM Mode_Act where ModeStatus=0 order by ModeID asc")
	 If Not Rs.Eof Then
	  Act_MX_Sys_Arr=Rs.GetRows(-1)
	 Else
	  Act_MX_Sys_Arr=""
	 End If
	 Rs.Close:Set Rs=Nothing
   End Function


	Public Function DIYField(ModeID)
		Dim MXField,MXtext,i
		name=ModeID&"DIYFieldCache"
		If ObjIsEmpty() Then 
			MXField=Act_MX_Arr(ModeID,1)
			If IsArray(MXField) Then
				For I=0 To Ubound(MXField,2)
					MXtext=MXtext&","&MXField(0,I)
				Next
			Else
				MXtext=""
			End If 
			value=MXtext
		End If 
 		DIYField=value
 	End Function 

 	
	Public Function DIYFieldList(ModeID)
		Dim MXField,MXtext,i
		name=ModeID&"DIYFieldListCache"
		If ObjIsEmpty() Then 
			MXField=Act_MX_Arr(ModeID,1)
			If IsArray(MXField) Then
				For I=0 To Ubound(MXField,2)
					MXtext=MXtext&",["&MXField(0,I)&"]"
				Next
			Else
				MXtext=""
			End If 
			value=MXtext
		End If 
 		DIYFieldList=value
 	End Function 
	'把含有关键字和新生成的a标签加入字典中
	Function  AddToDic(key,strs)
		Dim pattern,reg,matches,m
		pattern="<[^>]*"&key&"[^>]*>|<a[^>]*>[^<]*"&key&"[^<]*<\/a>"
		Set  reg=new RegExp
		reg.Global=true
		reg.IgnoreCase=true
		reg.Pattern=pattern
		set matches=reg.Execute(strs)
		For  each m in matches
			dic.Add "key"&Index,m.value    
			strs=replace(strs,m.value,"$key"&Index&"$")
			Index=Index+1
		Next 
		Set  matches=Nothing
		Set  reg=Nothing
		AddToDic=strs
	End  function
  	Public Function ReplaceSitelink(TempletContent)
		On Error Resume Next
		Dim sqlstr,Sitelink,i,reg,pattern,ky,rs,OpenType
 		Name=AcTCMSN&"ReplaceSitelink"
		If ObjIsEmpty() Then 
 			Set Rs = Actexe("Select Title,Url,Num,OpenType,description,repset,repcontent from Sitelink_ACT  where ifs=1 order by OrderID asc,id desc")
			If Rs.Eof Then Rs.Close : Set Rs = Nothing:ReplaceSitelink=TempletContent:Exit Function
			Value = Rs.GetRows(-1):Rs.Close : Set Rs = Nothing
		End If 
 		Sitelink=Value
		set reg=new RegExp
		reg.Global=true
		reg.IgnoreCase=true
		For  i=0 to ubound(Sitelink,2)
 			ky=trim(Sitelink(0,i))
			TempletContent=AddToDic(ky,TempletContent)
			reg.Pattern=ky
			If Sitelink(5,i) = 1 Then 
 				If Sitelink(3,i) = ""  Then
					OpenType = ""
				Else
					OpenType = " target=""" & Sitelink(3,i) & """"
				End If
 				TempletContent=replace(TempletContent,ky,"<a href="""&Sitelink(1,i)&""" title="""&Sitelink(4,i)& """" & OpenType & ">"&ky&"</a>",1,Sitelink(2,i))
			Else 
				TempletContent=replace(TempletContent,ky,Rep(Sitelink(6,i),"{$content}",ky),1,Sitelink(2,i))
			End If 
			TempletContent=AddToDic(ky,TempletContent)
 		Next 
		Set  reg=nothing
		For  i=0 to Index-1 
			TempletContent=replace(TempletContent,"$key"&i&"$",dic.Item("key"&i))
		Next 
		ReplaceSitelink=TempletContent
 	End Function 


	Public Function CopyFrom(C_Name)
			Dim Rs
			Set Rs = ActExe("Select Field1,Field2 from AC_ACT where Types=0 And  Field1='" & Trim(C_Name) & "'")
			If Rs.Eof Then Rs.Close : Set Rs = Nothing:CopyFrom=C_Name:Exit Function
			CopyFrom = "<a href=""" & Trim(RS("Field2")) & """ target=""_blank"">" & C_Name & "</a>"
			Rs.Close : Set Rs = Nothing
	End Function 
 	 Function GetIndexNavigation(TitleCss,OpenType,StrNav)'首页导航
		  GetIndexNavigation =  StrNav  
	 End Function

	Public Function Author(C_Name)
		Dim Rs
		Set Rs = ActExe("Select Field1,Field2 from AC_ACT where Types=1 And  Field1='" & Trim(C_Name) & "'")
		If Rs.Eof Then Rs.Close : Set Rs = Nothing:Author=C_Name:Exit Function
		Author = "<a href=""" & Trim(RS("Field2")) & """ target=""_blank"">" & C_Name & "</a>"
		Rs.Close : Set Rs = Nothing
	End Function 


 	Function GetParentID(ClassID)
		Dim ClassRS
		Set ClassRS=actexe("Select ParentID,ClassID,Classname from class_ACT where ClassID='"& ClassID &"' order by ID desc")
		If Not ClassRs.eof Then 
			If ClassRS("ParentID")<>"0" Then
				GetParentID = GetParentID(ClassRS("ParentID")) &GetParentID
			End If 
		End if
		GetParentID=GetParentID&ClassID&" , "
		ClassRS.Close:Set ClassRS = Nothing
 	End function


    Function  GetParent(ClassID) 
	   GetParent=Split(GetParentID(ClassID),",")(0)
    End Function 
 
 	 Function GetClassNavigation(TitleCss,OpenType,StrNav,ClassID,TypeMode)'栏目
	    Dim ACT_Nav
		ACT_Nav=GetClassNav(StrNav, OpenType, TitleCss, ClassID)
 		If CBool(Application(AcTCMSN & "ModeHome")) = True Then
		  GetClassNavigation =  ACT_Nav&StrNav 
		Else
		  GetClassNavigation =  ACT_Nav
		End If
	 End Function
 	 Function GetContentNavigation(TitleCss,OpenType,StrNav,ClassID,TypeMode)'内容
  		Dim ClassNavStr:ClassNavStr = GetClassNav(StrNav, OpenType, TitleCss, ClassID)
 		GetContentNavigation =   ClassNavStr 
	 End Function
 	Function GetClassNavStr(ClassID)
		Dim ClassRS
		Set ClassRS=actexe("Select ParentID,ClassID,Classname from class_ACT where ClassID='"& ClassID &"' order by ID desc")
		If Not ClassRs.eof Then 
			If ClassRS("ParentID")<>"0" Then
				GetClassNavStr = GetClassNavStr(ClassRS("ParentID")) &GetClassNavStr
			End If 
		End if
		GetClassNavStr=GetClassNavStr&ClassID&" , "
		ClassRS.Close:Set ClassRS = Nothing
	End function

	Function GetClassNav(StrNav,OpenType, TitleCss, ClassID)
	  Dim TSArr,i,Q
	  ClassID=GetClassNavStr(ClassID)
	  ClassID=Left(Trim(ClassID), Len(Trim(ClassID)) - 1)
	  TSArr = Split(ClassID, ",")
	  For I = 0 To UBound(TSArr)
		If i>0 Then Q=StrNav
		GetClassNav = GetClassNav & Q &"<a "& TitleCss &" href=""" & DiyClassName(Trim(TSArr(I))) & """" &OpenType& ">" & ACT_L(Trim(TSArr(I)), 2) & "</a>"
	  Next
	End Function 
   
 	Function TempClassID(ClassID)
	If ClassID = "" Then Exit Function
	Dim Rs,AllClassID
	Set Rs = ActExe("Select ClassID From Class_Act Where ParentID = '"&ClassID&"' Order By OrderID Desc,ID Desc") 
	If Rs.Eof Then
		AllClassID = "'" & ClassID & "'"
	Else
		AllClassID = ""
		Do While Not Rs.Eof
			AllClassID = AllClassID & "," & TempClassID(Rs(0))
			Rs.MoveNext
		Loop
		AllClassID = "'" & ClassID & "'" & AllClassID
	End If
	TempClassID = AllClassID
	Rs.Close:Set Rs = Nothing
	End Function
 	Public Function ActErr(ShowErr,ErrorUrl,errnum)
		Response.Redirect(ActSys&ActCMS_Sys(8)&"/Error.asp?Title="&Server.URLEncode("<li>"&ShowErr&"</li>")&"&errnum="&errnum&"&___"&ErrorUrl)
		Response.end
	End Function 
 	Public Function GroupArr(GroupID,Row)
 		  Dim Rs
		  Set Rs=ACTEXE("Select ModeID,GroupSetting from Group_Act  Where GroupID=" & GroupID & " order by GroupID desc")
		  If Not Rs.Eof Then
				GroupArr=Split(Rs("GroupSetting"),"^@$@^")(Row)
		  End If 
 	End Function 

	'功能:会员点券明细出入函数	   
   '参数:ModeID-模块ID,InfoID-信息ID，UserName-用户名,PointFlag-操作类型1收入2支出,Point-交易点数,User-操作员,Descript-操作备注
	Public Function PointInOrOut(ModeID,InfoID,UserID,PointFlag,Point,UserLog,Descript,ContributeFlag)
	  If Not IsNumeric(PointFlag) Or Not IsNumeric(Point) Or Point=0 Then PointInOrOut=false:Exit Function
	  Dim PointParam,CurrPoint
	  If PointFlag=1 Then 
	     PointParam="Set Point=Point+" & Point
	  ElseIF PointFlag=2 Then
	     PointParam="Set Point=Point-" & Point
	  Else
	    PointInOrOut=false:Exit Function
	  End If
 
	  If (Conn.Execute("Select top 1 * From Point_Log_ACT Where UserID=" & UserID & " and ModeID=" & ModeID & " and InfoID=" & InfoID & " And PointFlag=" & PointFlag).Eof) Or (ModeID=0 And InfoID=0) or ContributeFlag=0 Then
 		  If UserID<>0 Then 
			  Conn.Execute("Update User_ACT " & PointParam & " Where UserID=" & UserID)
			  CurrPoint=Conn.Execute("Select top 1 Point From User_ACT Where UserID=" & UserID)(0)
		  End If 
 		  Dim RsPoint:Set RsPoint=Server.CreateObject("Adodb.Recordset")
		  RsPoint.Open "Select * From Point_Log_ACT Where 1=1",Conn,1,3
		   RsPoint.AddNew
			 RsPoint("ModeID")=ModeID
 			 RsPoint("UserID")=UserID
 			 RsPoint("InfoID")=InfoID
			 RsPoint("PointFlag")=PointFlag
			 RsPoint("Point")=Point
			 RsPoint("Times")=Now()
			 RsPoint("UserLog")=UserLog
			 RsPoint("Descript")=Descript
			 RsPoint("AddDate")=Now()
			 RsPoint("IP")=getip
			 RsPoint("CurrPoint")=CurrPoint
			 RsPoint("ContributeFlag")=ContributeFlag
		    RsPoint.Update
 
  	  End If
	  IF Err Then PointInOrOut=false Else PointInOrOut=true
	End Function
 	'功能:资金明细出入函数	                 
	'参数:UserName-用户名,ClientName-客户姓名,Money-金钱,MoneyType-类型,IncomeFlag-操作类型1收入2支出,PayTime-汇款日期,OrderID-订单号,Inputer-操作员,Remark-操作备注
	Public Function MoneyInOrOut(UserID,ClientName,Money,MoneyType,IncomeFlag,PayTime,OrderID,Inputer,Remark,ModeID,InfoID)
	  If Not IsNumeric(IncomeFlag) Or Not IsNumeric(Money) Or Money="0" Then MoneyInOrOut=false:Exit Function
	  Dim MoneyParam,CurrMoney
	  If IncomeFlag=1 Then 
	     MoneyParam="Set [Money]=[Money]+" & Money
	  ElseIF IncomeFlag=2 Then
	     MoneyParam="Set [Money]=[Money]-" & Money
	  Else
	    MoneyInOrOut=false:Exit Function
	  End If

   	  If (Conn.Execute("Select top 1 * From Money_Log_ACT Where UserID=" & UserID & " and ModeID=" & ModeID & " and InfoID=" & InfoID & " And IncomeFlag=" & IncomeFlag).Eof) Or (ModeID=0 And InfoID=0) Then
		 ' on error resume next
		  Conn.Execute("Update User_ACT " & MoneyParam & " Where UserID=" & UserID & "")
		  CurrMoney=Conn.Execute("Select top 1 Money From User_ACT Where UserID=" & UserID & "")(0)
	      Conn.Execute("Insert into Money_Log_ACT([UserID],[ClientName],[Money],[MoneyType],[IncomeFlag],[OrderID],[Remark],[PayTime],[LogTime],[Inputer],[IP],[CurrMoney],[ModeID],[InfoID]) values(" & UserID & ",'" & ClientName & "'," & Money & "," & MoneyType & ","& IncomeFlag & ",'" & OrderID & "','" & replace(Remark,"'","""") & "'," & NowString & "," &NowString & ",'" & replace(inputer,"'","""") & "','" & replace(getip,"'","""") & "'," & CurrMoney & "," & ModeID & "," & InfoID & ")")
	  End If
	  IF Err Then MoneyInOrOut=false Else MoneyInOrOut=true
	End Function



	'会员有效期明细出入函数
	'参数:UserName,InOrOutFlag,Edays,User,Descript

	Function EdaysInOrOut(UserID,Flag,Edays,Userlog,Descript)
 		 If Not IsNumeric(Flag) Or Not IsNumeric(Edays) Or Edays=0 Then EdaysInOrOut=false:Exit Function
 		  Conn.Execute("insert into Edays_ACT(UserID,Flag,Edays,[Userlog],descript,adddate,ip) values(" & UserID & "," & Flag & "," & Edays & ",'" & Userlog & "','" & replace(descript,"'","""") & "'," & NowString & ",'" & getip & "')")
		  IF Err Then EdaysInOrOut=false Else EdaysInOrOut=true
	End Function

	Public Sub isAcceptOK(UserID,InfoTitle,ModeID)
	Exit Sub 
	    IF Not IsNumeric(ModeID) Then Exit Sub
	    IF  GroupID=0 Then Exit Sub
	    Dim RSAccept,Tgdianshu:Set RSAccept=Server.CreateObject("ADODB.RECORDSET")
			RSAccept.Open "Select Score,GroupID From "&ACT_U(GroupID,2)&" Where UserID=" & UserID & "",Conn,1,3
				IF Not RSAccept.Eof Then
					Tgdianshu=RSAccept(0)+GroupArr(RSAccept("GroupID"),15)
					If Tgdianshu="0" Or Tgdianshu="" Then Tgdianshu=0
					RSAccept(0)=Tgdianshu
					RSAccept.Update
					Dim Sender:Sender=ActCMS_Sys(0)
					Dim Title:Title="恭喜，您发表的稿件[" & InfoTitle & "]已被审核通过！！！"
					Dim Message:Message="稿件标题：" & InfoTitle &""_
					  & "获得点数：" & GroupArr(RSAccept("GroupID"),15) & ""_
					  & "备注：此信息由系统自动发布，请不要回复！！！"
					Call PointUpdate(ModeID,0,UserName,1,Tgdianshu,"系统","发表搞件[" & InfoTitle & "]所得")  '记录日志          
					Call SendInfo(UserName,Sender,Title,Message)
					ACTEXE("Update "&ACT_U(GroupID,2)&" Set ArticleNum=ArticleNum+1 Where UserID=" & UserID & "")'暂放

			End IF
		RSAccept.Close:Set RSAccept=Nothing
	End Sub
 	Function ReplaceUserContent(TempletContent,userid)
 		Dim MX_Arr,K,rs,i
 		Set rs=actexe("select top 1 * from User_Act where userid="&userid)
		If  rs.eof Then ReplaceUserContent="没有找到这个用户,请返回":response.end
 		For i=0 to rs.Fields.Count-1
			MX_Arr=MX_Arr&(rs.Fields(i).Name &",")
		Next
		MX_Arr = Replace(MX_Arr, ",PassWord","")
		MX_Arr=Split(Left(MX_Arr, Len(MX_Arr) - 1),",")
  		If InStr(TempletContent, "{$G_User}") > 0  Then
			Dim GName
			Set GName=actexe("select top 1 GroupName from Group_ACT where GroupID="&rs("GroupID"))
			If GName.eof Then ReplaceUserContent="没有找到这个组,请返回":response.end
			TempletContent = Replace(TempletContent, "{$G_User}",GName("GroupName"))
		End if
 		If InStr(TempletContent, "{$myface_User}") > 0  Then
			If rs("myface")<>"" Then 
				TempletContent = Replace(TempletContent, "{$myface_User}",rs("myface"))
		    Else
				TempletContent = Replace(TempletContent, "{$myface_User}",ActSys&"user/images/nophoto.gif")
 		    End If 
 		End if
 		If IsArray(MX_Arr) Then
		  For K=0 To Ubound(MX_Arr)
 			  If trim(rs("" &MX_Arr(K) & ""))<>"" Then
  			  TempletContent = Replace(TempletContent,"{$" & MX_Arr(K) & "_User}",rs("" &MX_Arr(K) & ""))
			 Else
			  TempletContent = Replace(TempletContent,"{$" & MX_Arr(K) & "_User}","")
 			 End If
		  Next
		End If
 	    TempletContent=ReplaceUserDiyContent(TempletContent,rs("UModeID"),rs("UserID"))
	 	rs.Close:Set rs=Nothing
		ReplaceUserContent=TempletContent
	End Function
 	Public Function ReplaceUserDiyContent(TempletContent,UM,UID)	
   	Dim IF_NULL,rs1
		IF_NULL=Act_MX_Arr(UM,2) 
 		Set rs1=ActExe("select * from "&ACTCMS.ACT_U(UM,2)&" where userid="&UID)
		If IsArray(IF_NULL) Then
			For I=0 To Ubound(IF_NULL,2)
  			 If trim(rs1(IF_NULL(0,I)))<>"" And InStr(TempletContent,"{$" & IF_NULL(0,I) & "_User}")<>0 Then
  			   TempletContent = Replace(TempletContent,"{$" & IF_NULL(0,I) & "_User}",rs1(IF_NULL(0,I)))
			 Else
			   TempletContent = Replace(TempletContent,"{$" & IF_NULL(0,I) & "_User}","")
 			 End If
  			Next
		End If
		ReplaceUserDiyContent=TempletContent
		rs1.Close:Set rs1=Nothing
 	End Function 
	Public Function UserM(userid)
 		If ChkNumeric(userid)="0"  Then UserM=false:Exit Function 
   		Dim rs3
  		Set rs3=actexe("select userid,[UserName],[UmodeID] from User_Act where userid=" & userid & "")	
		If Not rs3.eof Then 
			UserM="<a  target='_blank' href='"&actsys&"space/?"&ACT_U(rs3("UmodeID"),5)&"-"&rs3("userid")&"'>"&rs3("UserName")&"</a>"
 		rs3.Close:Set rs3=Nothing
		End If 
  	End Function 
 	Public Function PointUpdate(ModeID,ID,UserName,PointFlag,Point,User,Descript)
	  Dim RsPoint:Set RsPoint=Server.CreateObject("Adodb.Recordset")
	  RsPoint.Open "Select * From Point_Log_ACT Where ID is null",Conn,1,3
	   RsPoint.AddNew
	     RsPoint("ModeID")=ModeID
		 RsPoint("ID")=ID
		 RsPoint("UserName")=UserName
		 RsPoint("PointFlag")=PointFlag
		 RsPoint("Point")=Point
		 RsPoint("Times")=1
		 RsPoint("User")=User
		 RsPoint("Descript")=Descript
		 RsPoint("AddDate")=now
		 RsPoint("IP")=Request.ServerVariables("Remote_Addr")
	   RsPoint.Update
	  RsPoint.Close:Set RsPoint=Nothing
	End Function

	Public Sub SendInfo(Incept,Sender,title,Content)
	  ActExe("insert Into Message_Act(Incept,Sender,Title,Content,SendTime,Flag,IsSend,DelR,DelS) values('" & Incept & "','" & Sender & "','" & replace(Title,"'","""") & "','" & replace(Content,"'","""") & "'," & NowString & ",0,1,0,0)")
	End Sub

	'Folder要创建的目录
	 Function CreateFolder(Folder)
		Dim FSO,  SplitFolder, CF, k
		on error resume next
		If Folder = "" Then
		 CreateFolder = False:Exit Function
		End If
	   Folder = Replace(Folder, "\", "/")
	   If Right(Folder, 1) <> "/" Then
		Folder = Folder & "/"
	   End If
	   If Left(Folder, 1) <> "/" Then
		Folder = "/" & Folder
		End If
		 Set FSO = CreateObject(ActCMS_Other(10))
		 If Not FSO.FolderExists(Server.MapPath(Folder)) Then
		   SplitFolder = Split(Folder, "/")
		 For k = 0 To UBound(SplitFolder) - 1
		  If k = 0 Then
		   CF = SplitFolder(k) & "/"
		  Else
		  CF = CF & SplitFolder(k) & "/"
		  End If
		  If (Not FSO.FolderExists(Server.MapPath(CF))) Then
			 FSO.CreateFolder (Server.MapPath(CF))
			 CreateFolder = True
		  End If
		 Next
	   End If
	   Set FSO = Nothing
	   If Err.Number <> 0 Then
	   Err.Clear
	   CreateFolder = False
	   Else
	   CreateFolder = True
	   End If
	 End Function

	Public Function DeleteFile(FileStr)'FSO删除
	   Dim FSO
	   on error resume next
	   Set FSO = CreateObject(ActCMS_Other(10))
		If FSO.FileExists(FileStr) Then
			FSO.DeleteFile FileStr, True
		Else
		DeleteFile = True
		End If
	   Set FSO = Nothing
	   If Err.Number <> 0 Then
	   Err.Clear:DeleteFile = False
	   Else
	   DeleteFile = True
	   End If
	End Function

	Public Function ACT_ATT(Selected)
		 Dim RSObj
	    Set RSObj = ACTExe("Select AID,Aname From ATT_ACT")
	  	Do While Not RSObj.Eof
		   IF Selected=RSObj(0) Then
			ACT_ATT=ACT_ATT & "<option value=""" & RSObj(0) & """ Selected>" & RSObj(1) & "</option>"& vbCrLf
		   Else
			ACT_ATT=ACT_ATT & "<option value=""" & RSObj(0) & """>" & RSObj(1) & "</option>"& vbCrLf
		   End If
		RSObj.MoveNext
		Loop
	  RSObj.Close:Set RSObj=Nothing
	End Function	

	Public Function ReplaceUrl(ReplaceContent, SaveFilePath)
		Dim re, BeyondFile, BFU, SaveFileName, SysDomain
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\){1}((\w)+[.]){1,}(net|com|cn|org|cc|tv|[0-9]{1,3})(\S*\/)((\S)+[.]{1}(gif|jpg|png|bmp)))"
		Set BeyondFile = re.Execute(ReplaceContent)
		Set re = Nothing
		For Each BFU In BeyondFile
		If InStr(BFU, ActCMS_Sys(2)) = 0 Then 
			SaveFileName = Year(Now()) & Month(Now()) & Day(Now()) & MakeRandom(10) & Mid(BFU, InStrRev(BFU, "."))
			 Call SaveFile(SaveFilePath&SaveFileName,BFU)
			 If  ActCMS_Other(9)="0" Then 
				ReplaceContent = Replace(ReplaceContent, BFU,  ACTCMS.PathDoMain&SaveFilePath & SaveFileName)
			 Else
				ReplaceContent = Replace(ReplaceContent, BFU,  SaveFilePath & SaveFileName)
			 End If 
		End If 
		Next
		ReplaceUrl = ReplaceContent
	End Function
	
	Function SaveFile(LocalFileName,RemoteFileUrl)
	    on error resume next
		Dim SaveRemoteFile:SaveRemoteFile=True
		dim Ads,Retrieval,GetRemoteData
		Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
		With Retrieval
			.Open "Get", RemoteFileUrl, False, "", ""
			.Send
			If .Readystate<>4 then
				SaveRemoteFile=False
				Exit Function
			End If
			GetRemoteData = .ResponseBody
		End With
		Set Retrieval = Nothing
		Set Ads = Server.CreateObject("Adodb.Stream")
		With Ads
			.Type = 1
			.Open
			.Write GetRemoteData
			.SaveToFile server.MapPath(LocalFileName),2
			.Cancel()
			.Close()
		End With
		Set Ads=nothing
		SaveFile=SaveRemoteFile
		Dim W:Set W = New CreateView
		Call  W.SY(LocalFileName,LocalFileName)
 		Set W=Nothing
	End Function
	
	'生成指定位数的随机数
	Public Function MakeRandom(ByVal maxLen)
	  Dim strNewPass,whatsNext, upper, lower, intCounter
	  Randomize
	 For intCounter = 1 To maxLen
	   upper = 57:lower = 48:strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
	 Next
	   MakeRandom = strNewPass
	End Function


	'**************************************************
	'函数名：strLength
	'作  用：求字符串长度。汉字算两个字符，英文算一个字符。
	'参  数：str  ----要求长度的字符串
	'返回值：字符串长度
	'**************************************************
	Public Function strLength(Str)
		On Error Resume Next
		Dim WINNT_CHINESE:WINNT_CHINESE = (Len("中国") = 2)
		If WINNT_CHINESE Then
			Dim l, T, c,I
			l = Len(Str)
			T = l
			For I = 1 To l
				c = Asc(Mid(Str, I, 1))
				If c < 0 Then c = c + 65536
				If c > 255 Then
					T = T + 1
				End If
			Next
			strLength = T
		Else
			strLength = Len(Str)
		End If
		If Err.Number <> 0 Then Err.Clear
	End Function


   Public Function GetStrValue(ByVal strs, ByVal strlen)
		If strs = "" Then GetStrValue = "":Exit Function
		If strlen=0 Then GetStrValue=strs:Exit Function
		Dim l, T, c, I, strTemp
		Dim str
		str=CloseHtml(strs)
		l = Len(Str)
		T = 0
		strTemp = Str
		strlen = CLng(strlen)
		For I = 1 To l
 		    If session.codepage="65001" Then 
				c = Abs(Ascw(Mid(Str, I, 1)))
		    Else 
				c = Abs(Asc(Mid(Str, I, 1)))
		    End If 

			If c > 255 Then
				T = T + 2
			Else
				T = T + 1
			End If
			If T >= strlen Then
				strTemp = Left(Str, I)
				Exit For
			End If
		Next
		If strTemp <> Str Then	strTemp = strTemp
		GetStrValue=Replace(strs,str,strTemp)
  End Function

 	Function FFile(Templetcontent,FileName)
  		on error resume next 
		Dim FileFSO,FileType
		 Set FileFSO = Server.CreateObject("ADODB.Stream")
			With FileFSO
			.Type = 2
			.Mode = 3
			.Open
			.Charset = "utf-8"
			.Position = FileFSO.Size
			.WriteText  Templetcontent
			.SaveToFile Server.MapPath(FileName),2
			If Err.Number<>0 Then 
				Err.Clear 
				Exit Function 
			End If 
			.Close
			End With
		Set FileType = nothing
		Set FileFSO = nothing
	End Function


	Function  LTemplate(temppath) 
 		on error resume next
		Dim  Str,A_W
		set A_W=server.CreateObject("adodb.Stream")
		A_W.Type=2 
		A_W.mode=3 
		A_W.charset="utf-8"
		A_W.open
		A_W.loadfromfile server.MapPath(temppath)
		If Err.Number<>0 Then Err.Clear:LTemplate="":Exit Function
		Str=A_W.readtext
		A_W.Close
		Set  A_W=nothing
		LTemplate=Str
	End  function


	Public Function HTMLCode(fString)
		If Not IsNull(fString) then
		fString = replace(fString, "&gt;", ">")
		fString = replace(fString, "&lt;", "<")
		fString = Replace(fString,  "&nbsp;"," ")
		fString = Replace(fString, "&quot;", CHR(34))
		fString = Replace(fString, "&#39;", CHR(39))
		fString = Replace(fString, "</P><P> ",CHR(10) & CHR(10))
		fString = Replace(fString, "<BR> ", CHR(10))
		HTMLCode = fString
		End If
	End Function
	Public Function HTMLEncode(fString)
		If Not IsNull(fString) then
		fString = replace(fString, ">", "&gt;")
		fString = replace(fString, "<", "&lt;")
		fString = replace(fString, "&", "&amp;")
		fString = Replace(fString, CHR(32), "&nbsp;")
		fString = Replace(fString, CHR(9), "&nbsp;")
		fString = Replace(fString, CHR(34), "&quot;")
		fString = Replace(fString, CHR(39), "&#39;")
		fString = Replace(fString, CHR(13), "")
		fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
		fString = Replace(fString, CHR(10), "<BR> ")
		HTMLEncode = fString
		End If
	End Function

	Public Function CloseHtml(ContentStr)
		On Error Resume Next
		Dim TempLoseStr, regEx
		If Trim(ContentStr)="" Then Exit Function
		TempLoseStr = CStr(ContentStr)
		Set regEx = New RegExp
		regEx.Pattern = "<\/*[^<>]*>"
		regEx.IgnoreCase = True
		regEx.Global = True
		TempLoseStr = regEx.Replace(TempLoseStr, "")
		CloseHtml = TempLoseStr
	End Function
	Function japanHtml(str)
		str=Replace(str,"ガ","&#12460;")
		str=Replace(str,"ギ","&#12462;")
		str=Replace(str,"ア","&#12450;")
		str=Replace(str,"ゲ","&#12466;")
		str=Replace(str,"ゴ","&#12468;")
		str=Replace(str,"ザ","&#12470;")
		str=Replace(str,"ジ","&#12472;")
		str=Replace(str,"ズ","&#12474;")
		str=Replace(str,"ゼ","&#12476;")
		str=Replace(str,"ゾ","&#12478;")
		str=Replace(str,"ダ","&#12480;")
		str=Replace(str,"ヂ","&#12482;")
		str=Replace(str,"ヅ","&#12485;")
		str=Replace(str,"デ","&#12487;")
		str=Replace(str,"ド","&#12489;")
		str=Replace(str,"バ","&#12496;")
		str=Replace(str,"パ","&#12497;")
		str=Replace(str,"ビ","&#12499;")
		str=Replace(str,"ピ","&#12500;")
		str=Replace(str,"ブ","&#12502;")
		str=Replace(str,"ブ","&#12502;")
		str=Replace(str,"プ","&#12503;")
		str=Replace(str,"ベ","&#12505;")
		str=Replace(str,"ペ","&#12506;")
		str=Replace(str,"ボ","&#12508;")
		str=Replace(str,"ポ","&#12509;")
		str=Replace(str,"ヴ","&#12532;")
		japanHtml=str
 	End Function 

	Function Htmljapan(str)
		str=Replace(str,"&#12460;","ガ")
		str=Replace(str,"&#12462;","ギ")
		str=Replace(str,"&#12450;","ア")
		str=Replace(str,"&#12466;","ゲ")
		str=Replace(str,"&#12468;","ゴ")
		str=Replace(str,"&#12470;","ザ")
		str=Replace(str,"&#12472;","ジ")
		str=Replace(str,"&#12474;","ズ")
		str=Replace(str,"&#12476;","ゼ")
		str=Replace(str,"&#12478;","ゾ")
		str=Replace(str,"&#12480;","ダ")
		str=Replace(str,"&#12482;","ヂ")
		str=Replace(str,"&#12485;","ヅ")
		str=Replace(str,"&#12487;","デ")
		str=Replace(str,"&#12489;","ド")
		str=Replace(str,"&#12496;","バ")
		str=Replace(str,"&#12497;","パ")
		str=Replace(str,"&#12499;","ビ")
		str=Replace(str,"&#12500;","ピ")
		str=Replace(str,"&#12502;","ブ")
		str=Replace(str,"&#12502;","ブ")
		str=Replace(str,"&#12503;","プ")
		str=Replace(str,"&#12505;","ベ")
		str=Replace(str,"&#12506;","ペ")
		str=Replace(str,"&#12508;","ボ")
		str=Replace(str,"&#12509;","ポ")
		str=Replace(str,"&#12532;","ヴ")
	End Function 

		Function DelSql(Str)
			Dim SplitSqlStr,SplitSqlArr,I
			SplitSqlStr="*|and |exec |insert |select |delete |update |count |master |truncate |declare |and	|exec	|insert	|select	|delete	|update	|count	|master	|truncate	|declare	|char(|mid(|chr("
			SplitSqlArr = Split(SplitSqlStr,"|")
			For I=LBound(SplitSqlArr) To Ubound(SplitSqlArr)
				If Instr(LCase(Str),SplitSqlArr(I))<>0 Then
					Call Alert ("系统警告！\n\n1、您提交的数据有恶意字符;\n2、您的数据已经被记录;\n3、操作日期："&Now&";\n.Com!","")
					Response.End
				End if
			Next
			DelSql = Str
		End Function


		Public Function S(Str)
		 S = Request(Str)
		End Function
		Public Function G(Str)
		 G = Request(Str)
		End Function

		Public Function Alert(SuccessStr, Url)
		 If Url <> "" Then
		  Response.Write ("<script language=""Javascript""> alert('" & SuccessStr & "');location.href='" & Url & "';</script>")
		 Else
		  Response.Write ("<script language=""Javascript""> alert('" & SuccessStr & "');history.back(-1);</script>")
		 End If
		 response.end
		End Function

	
		Public Function ShowPagePara(totalnumber, MaxPerPage, FileName, ShowAllPages, strUnit, CurrentPage, ParamterStr)
				 Dim N, I, PageStr
				Const Btn_First = "<font face='webdings' size='1' title='第一页'>9</font>" '定义第一页按钮显示样式
				Const Btn_Prev = "<font face='webdings' size='1' title='上一页'>3</font>" '定义前一页按钮显示样式
				Const Btn_Next = "<font face='webdings' size='1' title='下一页'>4</font>" '定义下一页按钮显示样式
				Const Btn_Last = "<font face='webdings' size='1' title='最后一页'>:</font>" '定义最后一页按钮显示样式
				  PageStr = ""
					If totalnumber Mod MaxPerPage = 0 Then
						N = totalnumber \ MaxPerPage
					Else
						N = totalnumber \ MaxPerPage + 1
					End If
				If N > 1 Then
					PageStr = PageStr & ("页次：<font color=red>" & CurrentPage & "</font>/" & N & "页 共有:" & totalnumber & strUnit & " 每页:" & MaxPerPage & strUnit & " ")
					If CurrentPage < 2 Then
						PageStr = PageStr & Btn_First & " " & Btn_Prev & " "
					Else
						PageStr = PageStr & ("<a href=" & FileName & "?page=1" & "&" & ParamterStr & ">" & Btn_First & "</a> <a href=" & FileName & "?page=" & CurrentPage - 1 & "&" & ParamterStr & ">" & Btn_Prev & "</a> ")
					End If
					
					If N - CurrentPage < 1 Then
						PageStr = PageStr & " " & Btn_Next & " " & Btn_Last & " "
					Else
						PageStr = PageStr & (" <a href=" & FileName & "?page=" & (CurrentPage + 1) & "&" & ParamterStr & ">" & Btn_Next & "</a> <a href=" & FileName & "?page=" & N & "&" & ParamterStr & ">" & Btn_Last & "</a> ")
					End If
					If ShowAllPages = True Then
						PageStr = PageStr & ("GO:<select  onChange='location.href=this.value;' style='width:55;' name='select'>")
				    For I = 1 To N
					 If Cint(CurrentPage) = I Then
						PageStr = PageStr & ("<option value=" & FileName & "?page=" & I & "&" & ParamterStr & " selected>NO." & I & "</option>")
					 Else
					   PageStr = PageStr & ("<option value=" & FileName & "?page=" & I & "&" & ParamterStr & ">NO." & I & "</option>")
					 End If
				   Next
				   PageStr = PageStr & "</select>"
				  End If
			 End If
			 ShowPagePara = PageStr
		End Function

	   Function DiyClassName(ClassID)
	    on error resume Next
 	    Name = CStr("DiyClassName" & ClassID)
		If ObjIsEmpty() Then
		If  ACT_C(ACT_L(ClassID,10),3)=2 Then 
			DiyClassName=  ActCMSDM & "list-"& ClassID &".html"
		Else 
			IF ACT_L(ClassID,12)="1" Then 
				IF ACT_L(ClassID,6)<>"" Or ACT_C(ACT_L(ClassID,10),3)=0 Then
					DiyClassName=  ActCMSDM & "List.asp?L-"& ClassID &".html"
				Else
					 If ACT_L(GetParent(ClassID),13)="1" Then 
						If ACT_L(ClassID,11)="0" Then 
						  DiyClassName=  ACT_L(ClassID,15)&"/"
						Else 
						  If Trim(ACT_L(ClassID,7))<>"" Then 
							  DiyClassName=  ACT_L(GetParent(ClassID),15)&"/"&ACT_L(ClassID,3)&ACT_L(ClassID,7)
						  Else 
							  DiyClassName=  ACT_L(GetParent(ClassID),15)&"/"&ACT_L(ClassID,3)
						  End If 
						End If 
					 Else 
					    If Trim(ACT_L(ClassID,7))<>"" Then 
						  DiyClassName=  ActCMSDM&ACT_C(ACT_L(ClassID,10),6)&ACT_L(ClassID,3)&ACT_L(ClassID,7)
					    Else 
						  DiyClassName=  ActCMSDM&ACT_C(ACT_L(ClassID,10),6)&ACT_L(ClassID,3)
						End If 
					 End If 
				End If
			ElseIf ACT_L(ClassID,12)="2" Then 
				DiyClassName= ACT_L(ClassID,17)
			Else 
				If ACT_C(ACT_L(ClassID,10),3)=0 Then 
 					DiyClassName= ActCMSDM & "List.asp?L-"& ClassID &".html"
				Else 
					DiyClassName= ActSys&ACT_C(ACT_L(ClassID,10),6)&ACT_L(ClassID,17)
				End If 
			End If
		 End If 
			Value=DiyClassName
	   Else
 			DiyClassName=Value
	   End If 
	   End Function
 	   Function GainClassName(ClassID,opens,TitleCssName)
			on error resume next
			Dim ClassRSArr
			ClassID=Trim(Replace(ClassID,"'",""))
			Name = CStr("Navigation" & ClassID)
			GainClassName= "<a " & TitleCssName & " href=""" & DiyClassName(ClassID) & """" & opens & ">"& ACT_L(ClassID,2) & "</a>"
	   End Function
 	   Sub AddTags(ModeID,Keyword)
			Dim i,Tag,TagRs
			Set TagRs = Server.createobject("Adodb.Recordset")
			Tag = Split(Keyword,",")
			For I = 0 To UBound(Tag)
			   TagRs.Open "Select * From Tags_ACT Where TagsChar ='" & Left(Tag(I),50) & "' And ModeID =" & ModeID,Conn,1,3
			   If TagRs.Eof Then
				 TagRs.AddNew
				 TagRs("TagsChar") = Left(Tag(I),50)
				 TagRs("ModeID") = ModeID
				 TagRs("AddTime") = Now
				 TagRs.Update
				End If:TagRs.Close
			Next
		   Set TagRs = Nothing 
	   End Sub
	   Function AutoSplitPage(StrNewsContent,Page_Split_page,AutoPagesNum)'自动分页
		Dim i,IsCount,OneChar,StrCount,FoundStr,Pages_i_Str,Pages_i_Arr
		AutoPagesNum = Clng(AutoPagesNum)
		Page_Split_page = Cstr(Page_Split_page)
 		If Len(StrNewsContent) < Int(AutoPagesNum+Round(AutoPagesNum/5)) Then AutoSplitPage=StrNewsContent : Exit Function
 		If StrNewsContent<>"" and AutoPagesNum<>0 and InStr(1,StrNewsContent,Page_Split_page)=0 then
			IsCount=True
			Pages_i_Str=""
			For i= 1 To Len(StrNewsContent)
				OneChar=Mid(StrNewsContent,i,1)
				If OneChar="<" Then
					IsCount=False
				ElseIf OneChar=">" Then
					IsCount=True
				Else
					If IsCount=True Then
						If Abs(Asc(OneChar))>255 Then
							StrCount=StrCount+2
						Else
							StrCount=StrCount+1
						End If
						If StrCount>=AutoPagesNum And i<Len(StrNewsContent) Then
							FoundStr=Left(StrNewsContent,i)
							If AllowSplitPage(FoundStr,"table|a|b>|i>|strong|div|span")=true then
								Pages_i_Str=Pages_i_Str & Trim(CStr(i)) & "," 
								StrCount=0
							End If
						End If
					End If
				End If	
			Next
			If Len(Pages_i_Str)>1 Then Pages_i_Str=Left(Pages_i_Str,Len(Pages_i_Str)-1)
			Pages_i_Arr=Split(Pages_i_Str,",")
			For i = UBound(Pages_i_Arr) To LBound(Pages_i_Arr) Step -1
				StrNewsContent=Left(StrNewsContent,Pages_i_Arr(i)) & Page_Split_page & Mid(StrNewsContent,Pages_i_Arr(i)+1)
			Next
		End If
		AutoSplitPage=StrNewsContent
	End Function
	Function AllowSplitPage(TempStr,FindStr)
		Dim Inti,BeginStr,EndStr,BeginStrNum,EndStrNum,ArrStrFind,i
		TempStr=LCase(TempStr)
		FindStr=LCase(FindStr)
		If TempStr<>"" and FindStr<>"" then
			ArrStrFind=split(FindStr,"|")
			For i = 0 to Ubound(ArrStrFind)
				BeginStr="<"&ArrStrFind(i)
				EndStr  ="</"&ArrStrFind(i)
				Inti=0
				do while instr(Inti+1,TempStr,BeginStr)<>0
					Inti=instr(Inti+1,TempStr,BeginStr)
					BeginStrNum=BeginStrNum+1
				Loop
				Inti=0
				do while instr(Inti+1,TempStr,EndStr)<>0
					Inti=instr(Inti+1,TempStr,EndStr)
					EndStrNum=EndStrNum+1
				Loop
				If EndStrNum=BeginStrNum then
					AllowSplitPage=true
				Else
					AllowSplitPage=False
					Exit Function
				End If
			Next
		Else
			AllowSplitPage=False
		End If
	End Function

	Public Function Getrewrite(modeid,id)
		Dim 	rewriteurl
		rewriteurl="{modeid}-{id}.html"
		If Instr(rewriteurl,"{id}") > 0 Then rewriteurl = Replace(rewriteurl,"{id}",id)
		If Instr(rewriteurl,"{modeid}") > 0 Then rewriteurl = Replace(rewriteurl,"{modeid}",modeid)
		Getrewrite=acturl&rewriteurl
	End Function 

	'取得每篇文章、图片链接
	Public Function GetInfoUrl(ByVal ModeID,ByVal ClassID,ByVal ID,ByVal ChangesLink,ByVal FileName,ByVal infopurview,ByVal readpoint)
		IF Not Isnumeric(ModeID) Then GetInfoUrl="#":Exit Function
 		If ChangesLink = "1" Then
			 GetInfoUrl = FileName
		ElseIf ACT_C(ModeID,3)=0  Then '动态
			 GetInfoUrl= ActCMSDM&"List.asp?C-"&ModeID&"-" &ID&".html"
		ElseIf ACT_C(ModeID,3)=2  Then '动态
			 GetInfoUrl= Getrewrite(modeid,id)
		Else
		    If ACT_L(GetParent(ClassID),13)="1" Then 
  				 If Right(ACT_L(ClassID,16),1)<>"/" Then 
						 GetInfoUrl= ACT_L(GetParent(ClassID),15)&"/"&FileName&ACT_C(ModeID,11)
				 Else
						 GetInfoUrl= ACT_L(GetParent(ClassID),15)&"/"&FileName&"/"
				 End If 
  			Else 
				Dim Tmps,TmpUs
 				If Right(ACT_L(ClassID,16),1)<>"/" Then 
						 GetInfoUrl= ActCMSDM&ACT_C(ModeID,6)&FileName&ACT_C(ModeID,11)
				 Else
						 GetInfoUrl= ActCMSDM&ACT_C(ModeID,6)&FileName&"/"
				End If 
			End If 
		End If
	End Function 

	'取得每篇文章、图片链接
	Public Function GetInfoUrlall(ByVal ModeID,ByVal ClassID,ByVal ID,ByVal ChangesLink,ByVal FileName,ByVal infopurview,ByVal readpoint)
		IF Not Isnumeric(ModeID) Then GetInfoUrl="#":Exit Function
 		If ChangesLink = "1" Then
			 GetInfoUrlall = FileName
		ElseIf ACT_C(ModeID,3)=0  Then '动态
			 GetInfoUrlall= acturl&"List.asp?C-"&ModeID&"-" &ID&".html"
		ElseIf ACT_C(ModeID,3)=2  Then '动态
			 GetInfoUrlall= Getrewrite(modeid,id)
		Else
		    If ACT_L(GetParent(ClassID),13)="1" Then 
  				 If Right(ACT_L(ClassID,16),1)<>"/" Then 
						 GetInfoUrlall= acturl&ACT_L(GetParent(ClassID),15)&"/"&FileName&ACT_C(ModeID,11)
				 Else
						 GetInfoUrlall= acturl&ACT_L(GetParent(ClassID),15)&"/"&FileName&"/"
				 End If 
  			Else 
				Dim Tmps,TmpUs
 				If Right(ACT_L(ClassID,16),1)<>"/" Then 
						 GetInfoUrlall= acturl&ACT_C(ModeID,6)&FileName&ACT_C(ModeID,11)
				 Else
						 GetInfoUrlall= acturl&ACT_C(ModeID,6)&FileName&"/"
				End If 
			End If 
		End If
	End Function 

	'显示分页的前部分
	'参数说明:PageStyle-分页样式,ItemUnit-单位,TotalPage-总页数,CurrPage-当前第N页,TotalInfo-总信息数,PerPageNumber-每页显示数
	Function  GetPageList(PageStyle,ItemUnit,TotalPage,CurrPage,TotalInfo,PerPageNumber)
	    Select Case  Cint(PageStyle)
		  Case 1
			GetPageList= "<div class=""pages""><div class=""plist"">" & "共 " & TotalInfo & " " & ItemUnit &"  页次:<font color=red> " & CurrPage & "</font>/" & TotalPage & "页  " & PerPageNumber & " " & ItemUnit &"/页 "
		 Case 2
			GetPageList= "<div class=""pages""><div class=""plist"">第<font color=red>" & CurrPage & "</font>页 共" & TotalPage & "页 "
		 Case 3
			GetPageList= "<div class=""pages""><div class=""plist"">第<font color=red>" & CurrPage & "</font>页 共" & TotalPage & "页 "
	   End Select
	End Function

	Public Function ReturnPageStyle(PageStyle)
		ReturnPageStyle = "         分页样式"
		ReturnPageStyle = ReturnPageStyle & "         <select name=""PageStyle"" style=""width:70%;"" class=""textbox"">"
		ReturnPageStyle = ReturnPageStyle & "          <option value=1"
		If PageStyle=1 Then ReturnPageStyle = ReturnPageStyle & " Selected"
		ReturnPageStyle = ReturnPageStyle & ">①首页 上一页 下一页 尾页</option>"
		ReturnPageStyle = ReturnPageStyle & "          <option value=2"
		If PageStyle=2 Then ReturnPageStyle = ReturnPageStyle & " Selected"
		ReturnPageStyle = ReturnPageStyle & ">②第N页,共N页 [1] [2] [3]</option>"
		ReturnPageStyle = ReturnPageStyle & "          <option value=3"
		If PageStyle=3 Then ReturnPageStyle = ReturnPageStyle & " Selected"
		ReturnPageStyle = ReturnPageStyle & ">③<< <  > >></option>"
		ReturnPageStyle = ReturnPageStyle & "          <option value=4"
		If PageStyle=4 Then ReturnPageStyle = ReturnPageStyle & " Selected"
		ReturnPageStyle = ReturnPageStyle & ">自定义</option>"
		ReturnPageStyle = ReturnPageStyle & "         </select>"
	End Function

	Public  Function  GetEn(EnStr)
		Dim  EnStr4,EnStr3,EnStr2,EnStr1
		Set  EnStr1=new regexp
			EnStr1.ignorecase=true
			EnStr1.global=true
			EnStr1.pattern="[a-zA-Z0-9\- ]"
			Set  EnStr3=EnStr1.execute(EnStr)
				For  each EnStr2 in EnStr3
					EnStr4=EnStr4&EnStr2.value
				Next 
			Set  EnStr3= Nothing 
		Set  EnStr1=nothing
		EnStr4=trim(EnStr4)
		If  len(EnStr4)>0 then EnStr4=replace(EnStr4," ","-")
		While  (instr(EnStr4,"--")>0)
			EnStr4=replace(EnStr4,"--","-")
		Wend 
		GetEn =EnStr4
	End  Function 

	Public Function PinYin(StrChar)
		Dim StrLens,RsStr,StrLen,StrTitle,IFCN,i,Rs
		 On  Error  Resume  Next 
 		Set  RsStr=Server.Createobject("Adodb.Connection")
		RsStr.open  "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath(ActCMS_Sys(3)&"ACT_inc/pinyin.mdb")
		IFCN=true
		For  i=1 to len(StrChar)
			StrTitle=IFCN
			StrLen=mid(StrChar,i,1)
			If  len(trim(StrLen)) = 1 Then 
				set rs=RsStr.execute("select top 1 pinyin from pinyin where content like '%"&StrLen&"%';")
					if not rs.eof and not rs.bof Then 
						StrLen=rs(0)
						IFCN=True 
					Else 
						IFCN=False 
					End  If 
					Rs.Close:Set  Rs = Nothing 
			Else 
				StrLen=" "
			End If 
			If  StrTitle=IFCN Then 
				StrLens=StrLens&StrLen
			Else 
				StrLens=StrLens&" "&StrLen
			End  If 
		Next 
		RsStr.Close:Set RsStr=nothing 
		PinYin=Trim(StrLens)
	End  Function
 	Public Function GetGroup_CheckBox(OptionName,SelectArr,RowNum)
	  On  Error  Resume  Next 
	   Dim n:n=0
	   Dim RSGroup,GroupName:Set RSGroup=Server.CreateObject("Adodb.Recordset")
	   IF RowNum<=0 Then RowNum=3
	   RSGroup.Open "Select GroupID,GroupName From Group_ACT",Conn,1,1
	   GetGroup_CheckBox="<table width=""100%"" align=""center"" border=""0"">"
	   Do While Not RSGroup.Eof
	        GetGroup_CheckBox=GetGroup_CheckBox & "<TR>"
	     For N=1 To RowNum
 		    GetGroup_CheckBox=GetGroup_CheckBox & "<TD class=""tdclass"" WIDTH=""" & CInt(100 / CInt(RowNum)) & "%"">"
			If Instr(SelectArr,RSGroup(0))<>0 Then
			 GetGroup_CheckBox=GetGroup_CheckBox & "<input id="& OptionName&RSGroup(0)&" type=""checkbox"" checked name=""" & OptionName & """ value=""" & RSGroup(0) & """><label for="& OptionName&RSGroup(0) &">" & RSGroup(1) & "</label>&nbsp;&nbsp;&nbsp;&nbsp;"
			Else
			 GetGroup_CheckBox=GetGroup_CheckBox & "<input id="& OptionName&RSGroup(0)&" type=""checkbox"" name=""" & OptionName & """ value=""" & RSGroup(0) & """><label for="& OptionName&RSGroup(0) &">" & RSGroup(1) & "</label>&nbsp;&nbsp;&nbsp;&nbsp;"
			End IF
			GetGroup_CheckBox=GetGroup_CheckBox & "</TD>"
		 	RSGroup.MoveNext
			If RSGroup.Eof Then Exit For
		Next
		GetGroup_CheckBox=GetGroup_CheckBox & "</TR>"
		If RSGroup.Eof Then Exit Do
	   Loop
	   GetGroup_CheckBox=GetGroup_CheckBox & "</TABLE>"
	   RSGroup.Close:Set RSGroup=Nothing
	End Function 




  	Public Function GetGroup_select(SelectArr)
	  On  Error  Resume  Next 
 	   Dim RSGroup,ac:Set RSGroup=Server.CreateObject("Adodb.Recordset")
	   ac=False 
 	   RSGroup.Open "Select GroupID,GroupName From Group_ACT",Conn,1,1
 	   Do While Not RSGroup.Eof
  			If ChkNumeric(SelectArr)=ChkNumeric(RSGroup(0)) Then
			ac=true
			 GetGroup_select=GetGroup_select & "<option  value="""& RSGroup(0)&"""  Selected >" & RSGroup(1) & "</option>"
			Else
			 GetGroup_select=GetGroup_select & "<option value="""& RSGroup(0)&"""> " & RSGroup(1) & "</option>"
			End IF
			GetGroup_select=GetGroup_select 
		 	RSGroup.MoveNext
  	   Loop
	   If ac=True Then 
			ac="<option value='0'>---保持原有用户组---</option>"
 	   Else
			ac="<option value='0' Selected>---保持原有用户组---</option>"
 	   End If 
		GetGroup_select=ac&GetGroup_select
 	   RSGroup.Close:Set RSGroup=Nothing
	End Function 









	'****************************************************
	'参数说明
	  'Subject     : 邮件标题
	  'MailAddress : 发件服务器的地址,如smtp.163.com
	  'LoginName     ----登录用户名(不需要请填写"")
	  'LoginPass     ----用户密码(不需要请填写"")
	  'Email       : 收件人邮件地址
	  'Sender      : 发件人姓名
	  'Content     : 邮件内容
	  'Fromer      : 发件人的邮件地址
	'****************************************************
	  Public Function SendMail(MailAddress, LoginName, LoginPass, Subject, Email, Sender, Content, Fromer)
	   on error resume next
		Dim JMail
		  Set jmail = Server.CreateObject("JMAIL.Message") '建立发送邮件的对象
			jmail.silent = true '屏蔽例外错误，返回FALSE跟TRUE两值j
			jmail.Charset = "utf-8" '邮件的文字编码为国标
			jmail.ContentType = "text/html" '邮件的格式为HTML格式
			jmail.AddRecipient Email '邮件收件人的地址
			jmail.From = Fromer '发件人的E-MAIL地址
			jmail.FromName = Sender
			  If LoginName <> "" And LoginPass <> "" Then
				JMail.MailServerUserName = LoginName '您的邮件服务器登录名
				JMail.MailServerPassword = LoginPass '登录密码
			  End If

			jmail.Subject = Subject '邮件的标题 
			JMail.Body = Content
			JMail.Priority = 1'邮件的紧急程序，1 为最快，5 为最慢， 3 为默认值
			jmail.Send(MailAddress) '执行邮件发送（通过邮件服务器地址）
			jmail.Close() '关闭对象
		Set JMail = Nothing
		If Err Then
			SendMail = Err.Description
			Err.Clear
		Else
			SendMail = "OK"
		End If
	  End Function


	Public Function IsObjInstalled(strClassString)
		on error resume next
		IsObjInstalled = False
		Err = 0
		Dim xTestObj:Set xTestObj = Server.CreateObject(strClassString)
		If 0 = Err Then IsObjInstalled = True
		Set xTestObj = Nothing
		Err = 0
	End Function
	Public Function IsExpired(strClassString)
		on error resume next
		IsExpired = True
		Err = 0
		Dim xTestObj:Set xTestObj = Server.CreateObject(strClassString)
		If 0 = Err Then
			Select Case strClassString
				Case "Persits.Jpeg"
					If xTestObjResponse.Expires > Now Then
						IsExpired = False
					End If
				Case "wsImage.Resize"
					If InStr(xTestObj.errorinfo, "已经过期") = 0 Then
						IsExpired = False
					End If
				Case "SoftArtisans.ImageGen"
					xTestObj.CreateImage 500, 500, RGB(255, 255, 255)
					If Err = 0 Then
						IsExpired = False
					End If
			End Select
		End If
		Set xTestObj = Nothing
		Err = 0
	End Function
	Public Function ExpiredStr(I)
		   Dim ComponentName(3)
			ComponentName(0) = "Persits.Jpeg"
			ComponentName(1) = "wsImage.Resize"
			ComponentName(2) = "SoftArtisans.ImageGen"
			ComponentName(3) = "CreatePreviewImage.cGvbox"
			If IsObjInstalled(ComponentName(I)) Then
				If IsExpired(ComponentName(I)) Then
					ExpiredStr = "，但已过期"
				Else
					ExpiredStr = ""
				End If
			  ExpiredStr = " √支持" & ExpiredStr
			Else
			  ExpiredStr = "×不支持"
			End If
	End Function





	Public Function ArrayToxml(DataArray,Recordset,row,xmlroot)
				Dim i,node,rs,j
				If xmlroot="" Then xmlroot="xml"
				Set ArrayToxml = Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
				ArrayToxml.appendChild(ArrayToxml.createElement(xmlroot))
				If row="" Then row="row"
				For i=0 To UBound(DataArray,2)
					Set Node=ArrayToxml.createNode(1,row,"")
					j=0
					For Each rs in Recordset.Fields
							 node.attributes.setNamedItem(ArrayToxml.createNode(2,LCase(rs.name),"")).text= DataArray(j,i)& ""
							 j=j+1
					Next
					ArrayToxml.documentElement.appendChild(Node)
				Next
		End Function

		Function GetPath(ClassID,Pathstr)
			Dim FolderPath,FileName,namearr,ci,namearrs
			Pathstr=Pathstr
			Pathstr = Replace(Pathstr, "//","/")
			Pathstr = Replace(Pathstr, "\","/")
			If InStr(Pathstr,"/")>0 Then 
				 namearr=Split(Pathstr,"/")
				 For ci=0 To UBound(namearr)-1
					namearrs=namearrs&namearr(ci)&"/"
				 Next
				If Right(Pathstr,1)="/"   Then
					FolderPath=namearrs
					FileName=namearrs&"index.html"
				Else
					If InStr(Pathstr,".")>0 Then 
						FolderPath= namearrs
						FileName=Pathstr 
					Else 
						FolderPath=Pathstr&"/"
						FileName=Pathstr&"/index.html"
					End If 
				End If 
			Else 
				If InStr(Pathstr,".")>0 Then 
					FileName=Pathstr 
				Else 
					FolderPath=Pathstr&"/"
					FileName=Pathstr&"/index.html"
				End If 
			End If 
			Call CreateFolder(actcms.ActSys&ACT_C(ACT_L(ClassID,10),6)&FolderPath)
			GetPath=ActSys&ACT_C(ACT_L(ClassID,10),6)&FileName 
 	    End Function 
 End Class
 
 	'按ID显示静态标签的函数，逆光与2013年10月22日添加
	Function ShowStaticContent(LID)
		Dim Rs
		Set rs=actcms.actexe("select LabelContent from Label_ACT where id="&LID)
		If rs.eof Then response.write "参数错误":response.end
		ShowStaticContent=rs("LabelContent")
	End Function
%>