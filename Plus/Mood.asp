<!--#include file="../act_inc/ACT.User.asp"-->
<%
 Dim id,rs,types,sql,i,ModeID,MdID,Action,SubmitNum
 Dim UnlockTime,StartTime,EndTime,Status
ModeID = ChkNumeric(request("ModeID"))
ID = ChkNumeric(request("ID"))
MdID = ChkNumeric(request("MDID"))
Action=request("Action")
If Action="count" Then 
 	
	Dim rsn,numcount
	numcount=0
 	set rsn=actcms.actexe("Select * From Mood_List_ACT Where ModeID=" & ModeID & " and MDID="& MDID &" And AID=" & ID)
	If Not  rsn.eof then 
	
		 For I=0 To 14 
			 numcount=numcount+rsn("M"&i)
		 Next
	 End If 
 	echo "document.writeln('"&ChkNumeric(numcount)&"')"
	
	
	response.end
End If 

 
 Set Rs=actcms.actexe("select top 1 ID,TitleContent,PicContent,UnlockTime,StartTime,EndTime,Status,SubmitNum from Mood_Plus_ACT where id="&MdID)
 If   rs.eof Then CloseConn:response.end
Status=rs("Status")
UnlockTime=rs("UnlockTime")
StartTime=rs("StartTime")
EndTime=rs("EndTime")
 SubmitNum=rs("SubmitNum")
if Action="submit" Then
	If actcms.S("PrintOut")="js" Then
		Response.Write "MoodPositionBack('" & postmood() & "');"
	Else
		Response.Write postmood()
	End iF
   Response.End()
Else
  Response.Write ReplaceJsBr(main())
End If
Function ReplaceJsBr(Content)
		 Dim i
		 Content=Replace(Content,"""","\""")
		 Dim JsArr:JSArr=Split(Content,Chr(13) & Chr(10))
		 For I=0 To Ubound(JsArr)
		   ReplaceJsBr=ReplaceJsBr & "document.writeln('" & JsArr(I) &"');" & vbcrlf 
		 Next
End Function



Function postmood()
	If SubmitNum=1 And Request.Cookies(AcTCMSN)("mood_"&ModeID&"_"&ID)="ok" Then postmood="standoff":Exit Function
   	If UnlockTime=0 Then '时间限制否?
		If Now < StartTime Then   postmood= ("document.write(对不起,心情指数还没有开始!);"):Exit Function 
		If Now > EndTime Then    postmood=("document.write(对不起,该心情指数已经结束!);"):Exit Function 
	End If 
 	Set Rs=Server.CreateObject("ADODB.RECORDSET")
 	Rs.Open "Select * From Mood_List_ACT Where ModeID=" & ModeID & " and MDID="& MDID &" And AID=" & ID,Conn,1,3
 	If Rs.Eof  Then
	 Rs.AddNew
	 Rs("ModeID")=ModeID
	 Rs("AID")=ID
	 Rs("MDID")=MDID
 	End If
 	 For I=0 To 14 
 	  If Clng(I)= ChkNumeric(request("itemid")) Then
		 Rs("M" & i)=Rs("M"&i)+1
 	  End If
	 Next
 	Rs.Update
	Rs.Close:Set Rs=Nothing
	Response.Cookies(AcTCMSN)("mood_"&ModeID&"_"&ID)="ok"
 	postmood = main
  End Function 
 Function main
 Dim MoodStr,CountMood,MoodFace	
 Set Rs=actcms.actexe("select top 1 ID,TitleContent,PicContent,UnlockTime,StartTime,EndTime from Mood_Plus_ACT where id="&MdID)
 If Not rs.eof Then 
 	If UnlockTime=0 Then '时间限制否?
		If Now < StartTime Then Call  echo ("document.write(对不起,心情指数还没有开始!);"):Exit Function 
		If Now > EndTime Then Call  echo("document.write(对不起,该心情指数已经结束!);"):Exit Function 
	End If 
	Dim TitleContent,PicContent
	TitleContent=Split(rs("TitleContent"),"@&@")
	PicContent=Split(rs("PicContent"),"@&@")
	
	Dim ItemVoteNum(15),TotalVote
	TotalVote=0
 	Set Rs=Server.CreateObject("ADODB.RECORDSET")
 	Rs.Open "Select * From Mood_List_ACT Where ModeID=" & ModeID & " and MDID="& MDID &" And AID=" & ID,Conn,1,1
 	If Not Rs.Eof Then
	 For I=0 To 14
	  TotalVote=TotalVote+Rs("M" & I)
	  ItemVoteNum(i)=Rs("M" & I)
	 Next
	End If
	Rs.Close:Set Rs=Nothing
	
 
	Dim PerVote,Percentage
	For  i=0 To ubound(TitleContent)
		If Trim(TitleContent(i))<>"" And Trim(PicContent(i))<>"" Then 

	        If TotalVote<>0 Then PerVote=Round(ItemVoteNum(I)/TotalVote,4)
		    Percentage=PerVote*100
			if Percentage<1 and Percentage<>0 then	Percentage= "0" & Percentage
 		MoodFace=MoodFace&"<li id=""m1_li""><em>"&ItemVoteNum(i)&"</em><div class=""mood_bar""><div style=""height: "& PerVote*50&"%;"" class=""mood_bar_in"" id=""m1_bar""></div></div></li>"
		MoodStr=MoodStr & "<li><a href=""###""  name=""votebutton""  onclick=""MoodPosition(" & I &");""><img src=""" & actcms.acturl&PicContent(i) & """ /><br />" & TitleContent(i) & "</a><br /><input type=""radio"" onclick=""MoodPosition(" & I &");"" name=""votebutton""></li>"
		
		End If 
	Next
End If 
 MoodStr="<div id=""xinqing""><table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0""><tr><td class=""mood_top"" colspan=""15""><ul class=""xqzt"">"&MoodFace&"</ul>"&MoodStr&"</td></tr></table></div>"
main=MoodStr
End Function %>
  CreateMoodAjax=function(){
	if(window.XMLHttpRequest){
		return new XMLHttpRequest();
	} else if(window.ActiveXObject){
		return new ActiveXObject("Microsoft.XMLHTTP");
	} 
	throw new Error("XMLHttp object could be created.");
}
ajaxReadText=function(file,fun){
	var xmlObj = CreateMoodAjax();
	
	xmlObj.onreadystatechange = function(){
		if(xmlObj.readyState == 4){
			if (xmlObj.status ==200){
				obj = xmlObj.responseText;
				eval(fun);
			}
			else{
				alert("读取文件出错,错误号为 [" + xmlObj.status  + "]");
			}
		}
	}
	try{
	xmlObj.open ('GET', file, true);
	xmlObj.send (null);
	}
	catch(e){
		var head = document.getElementsByTagName("head")[0];        
		var js = document.createElement("script"); 
		js.src = file+"&printout=js"; 
		head.appendChild(js);   
	}
}
MoodPosition=function(itemid){
  ajaxReadText('<%=actcms.acturl%>plus/Mood.asp?Action=submit&ModeID='+<%=ModeID%>+'&ID=<%=ID%>&MDID=<%=MDID%>&itemid='+itemid+"&" + Math.random(),'MoodPositionBack(obj)');
}
MoodPositionBack=function(obj){
 switch(obj){
  case "standoff":
   alert('您已表态过了, 不能重复表态!');
   break;
  case "standoff":
   alert('您已表态过了, 不能重复表态!');
   break;
  case "Status":
   alert('心情指数已关闭!');
   break;
  default:
   alert('恭喜,您已成功表态!');
   document.getElementById('xinqing').innerHTML=obj;
   break;
 }
}