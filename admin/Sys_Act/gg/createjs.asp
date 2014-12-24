<!--#include file="../../ACT.Function.asp"-->
<meta http-equiv="Content-Type" content="text/html; charSet=utf-8">
<%
ConnectionDatabase

	Dim ADID,HostURL,ScrCode,JsCode,Rs,sql,JsFileName,objStream
	Dim AD_Type,AD_Src,AD_Width,AD_Height,AD_Link,AD_Alt,AD_Views,ADViews,AD_Hits,ADHits
	Dim AD_Stop_Views,AD_Stop_Hits,AD_Stop_Date,AD_Code,AD_ADHits,ID
		Response.Expires = 0

		ADID=Request.QueryString("ID")
		'HostURL="http://"&Request.Servervariables("server_name")&":"&Request.ServerVariables("SERVER_PORT")&replace(Request.Servervariables("url"),"/createjs.asp","")
                HostURL=ACTCMS.ACTCMSDM&"plus/"

	Set rs=server.createobject("adodb.recordset")
        sql = "Select ID,ADType,ADSrc,ADCode,ADWidth,ADHeight,ADAlt,ADViews,ADHits,ADStopViews,ADStopHits,ADStopDate from [ads] where ADID='" & ADID & "'"
		rs.open sql,conn,1,3
	If Not (rs.bof or rs.eof) Then
		AD_Type=rs("ADType")
		AD_Src=rs("ADSrc")
		AD_Code=rs("ADCode")
		AD_Width=rs("ADWidth")
		AD_Height=rs("ADHeight")
		AD_Alt=rs("ADAlt")
		ID=rs("ID")
		AD_Views=rs("ADViews")
		AD_ADHits=rs("ADHits")
		AD_Stop_Views=rs("ADStopViews")
		AD_Stop_Hits=rs("ADStopHits")
		AD_Stop_Date=rs("ADStopDate")
		rs("ADViews")= AD_Views + 1
		rs.Update
	Else
		response.write "document.write('<! -- 没有找到您要浏览的广告 -->');"
	End If
		rs.Close
	Set rs=nothing
		conn.Close
	Set conn=nothing

	If ( AD_Stop_Views <> 0 and AD_Views > AD_Stop_Views) Then AD_Type = 0
	If ( AD_Stop_Hits <> 0 and AD_Hits > AD_Stop_Hits) Then AD_Type = 0
	If ( NOW() > AD_Stop_Date) Then AD_Type = 0

	If InStr(1,AD_Src,".swf",1)>0 Then
		ScrCode="<EMBED src="""& AD_Src &""" quality=high WIDTH="""& AD_Width &""" HEIGHT="""& AD_Height &""" TYPE=""application/x-shockwave-flash"" PLUGINSPAGE=""http://www.macromedia.com/go/getflashplayer""></EMBED>"
	Else
		ScrCode="<a href="""& HostURL &"gg/openad.asp?id="& ID &""" target=_blank><img src="""& AD_Src &""" border=0 width="""& AD_Width &""" height="""& AD_Height &""" alt="""& AD_Alt &""" align=top></a>"
	End If

	Select Case AD_Type
	Case 0
		JsCode= "document.write('<!-- 过期广告 ID"&ADID&" -->')"
	Case 1
		JsCode= "document.write('"& ScrCode &"');"
	Case 2
		JsCode= "ns4=(document.layers)?true:false;" & vbCrLf & _
				"ie4=(document.all)?true:false;" & vbCrLf & _
				"if(ns4){document.write('<layer id=DGbanner2 width="& AD_Width &" height="& AD_Height &" onmouseover=stopme(""DGbanner2"") onmouseout=movechip(""DGbanner2"")>"& ScrCode &"</layer>');}" & vbCrLf & _
				"else{document.write('<div id=DGbanner2 style=""position:absolute; width:"& AD_Width &"px; height:"& AD_Height &"px; z-index:9; filter: Alpha(Opacity=90)"" onmouseover=stopme(""DGbanner2"") onmouseout=movechip(""DGbanner2"")>"& ScrCode &"</div>');}" & vbCrLf & _
				"document.write('<script language=javascript src="& HostURL &"gg/js/ad_float_fullscreen.js></script>');"
	Case 3
		JsCode= "if (navigator.appName == 'Netscape')" & vbCrLf & _
				"{document.write('<layer id=DGbanner3 top=150 width="& AD_Width &" height="& AD_Height &">"& ScrCode &"</layer>');}" & vbCrLf & _
				"else{document.write('<div id=DGbanner3 style=""position: absolute;width:"& AD_Width &";top:150;visibility: visible;z-index: 1"">"& ScrCode &"</div>');}" & vbCrLf & _
				"document.write('<script language=javascript src="& HostURL &"gg/js/ad_float_upanddown.js></script>');"
	Case 4
		JsCode= "if (navigator.appName == 'Netscape')" & vbCrLf & _
				"{document.write('<layer id=DGbanner10 top=150 width="& AD_Width &" height="& AD_Height &">"& ScrCode &"</layer>');}" & vbCrLf & _
				"else{document.write('<div id=DGbanner10 style=""position: absolute;width:"& AD_Width &";top:150;visibility: visible;z-index: 1"">"& ScrCode &"</div>');}" & vbCrLf & _
				"document.write('<script language=javascript src="& HostURL &"gg/js/ad_float_upanddown_L.js></script>');"
	Case 5
		JsCode= "ns4=(document.layers)?true:false;" & vbCrLf & _
				"if(ns4){document.write('<layer id=DGbanner4Cont onLoad=""moveToAbsolute(layer1.pageX-160,layer1.pageY);clip.height="& AD_Height &";clip.width="& AD_Width &"; visibility=show;""><layer id=DGbanner4News position:absolute; top:0; left:0>"& ScrCode &"</layer></layer>');}" & vbCrLf & _
				"else{document.write('<div id=DGbanner4 style=""position:absolute;top:0; left:0;""><div id=DGbanner4Cont style=""position:absolute; width:"& AD_Width &"; height:"& AD_Height &";clip:rect(0,"& AD_Width &","& AD_Height &",0)""><div id=DGbanner4News style=""position:absolute;top:0; left:0; right:820"">"& ScrCode &"</div></div></div>');} " & vbCrLf & _
				"document.write('<script language=javascript src="& HostURL &"gg/js/ad_fullscreen.js></script>');"
	Case 6
		JsCode= "window.showModalDialog('"& AD_Src &"','','dialogWidth:"& AD_Width &"px;dialogHeight:"& AD_Height &"px;scroll:no;status:no;help:no');"
	Case 7
		JsCode= "document.write('<script language=javascript src="& HostURL &"gg/js/ad_dialog.js></script>'); " & vbCrLf & _
				"document.write('<div style=""position:absolute;left:300px;top:150px;width:"& AD_Width &"; height:"& AD_Height &";z-index:1;solid;filter:alpha(opacity=90)"" id=DGbanner5 onmousedown=""down1(this)"" onmousemove=""move()"" onmouseup=""down=false""><table cellpadding=0 border=0 cellspacing=1 width="& AD_Width &" height="& AD_Height &" bgcolor=#000000><tr><td height=18 bgcolor=#5A8ACE align=right style=""cursor:move;""><a href=# style=""font-size: 9pt; color: #eeeeee; text-decoration: none"" onClick=clase(""DGbanner5"") >关闭>>><img border=""0"" src="""&HostURL&"images/close_o.gif""></a>&nbsp;</td></tr><tr><td bgcolor=f4f4f4 >&nbsp;"& AD_Src &"</td></tr></table></div>');"
	Case 8
		JsCode= "window.open('"& AD_Src &"','_blank');"
	Case 9
		JsCode= "window.open('"& AD_Src &"','DGBANNER7','width="& AD_Width &",height="& AD_Height &",scrollbars=1')"
	Case 10
		JsCode= "function closeAd(){" & vbCrLf & _
						"huashuolayer2.style.visibility='hidden';" & vbCrLf & _
						"huashuolayer3.style.visibility='hidden';}" & vbCrLf & _
						"function winload()" & vbCrLf & _
						"{" & vbCrLf & _
						"huashuolayer2.style.top=130;" & vbCrLf & _
						"huashuolayer2.style.left=5;" & vbCrLf & _
						"huashuolayer3.style.top=130;" & vbCrLf & _
						"huashuolayer3.style.right=5;" & vbCrLf & _
						"}" & vbCrLf & _
						"if(document.body.offsetWidth>800){" & vbCrLf & _
						"{" & vbCrLf & _
						"document.write('<div id=huashuolayer2 style=""position: absolute;visibility:visible;z-index:1""><table width=100  border=0 cellspacing=0 cellpadding=0><tr><td height=10 align=right bgcolor=666666><a href=javascript:closeAd()><img src="& HostURL &"images/close.gif width=12 height=10 border=0></a></td></tr><tr><td>" & ScrCode & "</td></tr></table></div>'" & vbCrLf & _
						"+'<div id=huashuolayer3 style=""position: absolute;visibility:visible;z-index:1""><table width=100  border=0 cellspacing=0 cellpadding=0><tr><td height=10 align=right bgcolor=666666><a href=javascript:closeAd()><img src="& HostURL &"images/close.gif width=12 height=10 border=0></a></td></tr><tr><td>" & ScrCode & "</td></tr></table></div>');" & vbCrLf & _
						"}" & vbCrLf & _
						"winload()" & vbCrLf & _
						"}"
	Case 11
		JsCode= "document.write('"& AD_Code &"');"
	End Select

	If JsCode<>"" Then
		JsFileName = Server.MapPath(ACTCMS.ACTSYS&"plus/gg/JS_c/"&ADID&".js")
		JsCode = "<!--" & vbCrLf & JsCode & vbCrLf &"//-->"  
	set objStream = Server.CreateObject("ADODB.Stream")
	With objStream
		.Open
		.Type = 2
		.Charset = "utf-8"
		.WriteText JsCode
		.SaveToFile JsFileName, 2
		.Close
	End With
	Set objStream = Nothing

	If Err <> 0 Then
		Response.Write ("<script>alert(' 生成JS代码失败! \n\n "& Err.Description &"');history.back();</script>")
	Else
		Response.Write ("<script>alert(' 生成JS代码成功 !');history.back();</script>")
	End If
	Else
		Response.Write ("<script>alert(' 生成JS代码失败! \n\n 广告类型不正确!');history.back();</script>")
	End If
%>