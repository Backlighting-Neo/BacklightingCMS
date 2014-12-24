<%
	Function GetModeID(ub,num,sttemp)
		Dim tempstr,i
 		For i=0 To ub	
			If i=num Then 
				tempstr=tempstr&newm(sttemp(i))&"§"
			Else 
				If i=ub Then 
					tempstr=tempstr&sttemp(i)
				Else 
					'echo sttemp(0)
					tempstr=tempstr&sttemp(i)&"§"
				End If 
			End If 

		Next 
		GetModeID=tempstr
	End Function 

	'=================自动生成====================
	Function RepModeID(LabelContent)
 		  Dim str,FileNames,ModeID,sttemp
 		  FileNames= Replace(Split(Split(LabelContent,"§")(0),"(")(0),"{$","")
		  Str=mid(LabelContent, InStrrev(LabelContent, "("))
		  sttemp=Split(Str,"§")
		  Select Case FileNames
				Case "GetArticleList"
					FileNames="栏目文章列表"
 					RepModeID="{$GetArticleList"&GetModeID(24,22,sttemp)'循环次数,模型ID下标,数组内容
 				Case "GetArticlePic"
					FileNames="图片文章列表"
  					RepModeID="{$GetArticlePic"&GetModeID(18,17,sttemp)'循环次数,模型ID下标,数组内容
				Case "GetSlide"
					FileNames="幻灯片文章"
					ModeID=Split(Str,"§")(3)
 					RepModeID="{$GetSlide"&GetModeID(5,3,sttemp)'循环次数,模型ID下标,数组内容
				Case "GetLastArticleList"
					FileNames="分页文章列表"
					ModeID=Split(Str,"§")(20)
 					RepModeID="{$GetLastArticleList"&GetModeID(22,20,sttemp)'循环次数,模型ID下标,数组内容
				Case "GetClassForArticleList"
					FileNames="循环栏目文章"
					Str=rep(Str,")}","")
					sttemp=Split(Str,"§")
  					RepModeID="{$GetClassForArticleList"&GetModeID(30,30,sttemp)&")}"'循环次数,模型ID下标,数组内容
				'
					' echo sttemp(30)
					' die ""
				Case "CorrelationArticleList"
					RepModeID=LabelContent
				Case "GetNavigation"
					RepModeID=LabelContent
				Case "GetLinkList"
					RepModeID=LabelContent
				Case "GetSpecial"
					RepModeID=LabelContent
				Case "GetClassNavigation"
 					RepModeID=LabelContent
		  End Select 
		'  echo RepModeID
 	End Function 




	Function GetLabel(ThemeID)
	Dim LabelContenttemp
	Dim n:n=0
	Dim m:m=0
	Dim k:k=0
	Dim LabelMdb:LabelMdb=ACTCMS.S("LabelMdb")
	Dim NewLabelID,regs:regs=ACTCMS.S("regs")
	Dim DataConn:Set DataConn = Server.CreateObject("ADODB.Connection")
	DataConn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(actcms.adminPath&"Sys_Act/Theme/ThemeInstallTemp/"&ThemeID&"/Label.mdb")
	If Err Then 
	Err.Clear:Set DataConn = Nothing:Response.Write "<tr><td>数据库路径不正确，连接出错</td></tr>":Response.End
	else
	 Dim rs:set rs=server.createobject("adodb.recordset")
	 rs.open "select * from Label_Act",dataconn,1,1
	 Dim rsa:set rsa=server.createobject("adodb.recordset")
	 do while not rs.eof 
	  rsa.open "select * from Label_Act where labelname='" & rs("labelname") & "'",conn,1,3
	  
			If rs("LabelType")=1 Then 
				LabelContenttemp=RepModeID(rs("LabelContent"))
			Else 
				LabelContenttemp=rs("LabelContent")
 			End If 
			  if rsa.eof then
			     rsa.addnew
				 rsa("LabelName")=rs("LabelName")
				 rsa("LabelContent")=LabelContenttemp
				 rsa("Description")=rs("Description")
				 rsa("LabelType")=rs("LabelType")
				 rsa("LabelFlag")=rs("LabelFlag")
				 rsa("AddDate")=rs("AddDate")
				 n=n+1
				rsa.update
			  else   '重名处理
			   if regs<>"" then
				 rsa("LabelContent")=LabelContenttemp
				 rsa("Description")=rs("Description")
				 rsa("LabelType")=rs("LabelType")
				 rsa("LabelFlag")=rs("LabelFlag")
				 rsa("AddDate")=rs("AddDate")
				 m=m+1
				rsa.update
			   else
			    k=K+1
			   end if
			  end if
			   rsa.close
			  rs.movenext
			 loop
			 rs.close:set rs=nothing
			 set rsa=nothing
			end if
 
	 
	 Application.Contents.RemoveAll
	response.write "<br><br><br><div align=center>模版安装成功!成功导入了 <font color=red>" & n & "</font> 个标签,覆盖了 <font color=red>" & m & "</font> 个标签,重名跳过了 <font color=red>" & k & "</font> 个标签！  </div><br><br><br><br><br><br><br>"
   dataconn.close:set dataconn=Nothing
   End Function 
 

%>