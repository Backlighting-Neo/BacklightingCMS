<!--#include file="../../act_inc/ACT.User.asp"-->
<%
			Response.Write "<html>" & vbCrLf
			Response.Write "<head>" & vbCrLf
			Response.Write "<title>申请友情链接</title>" & vbCrLf
			Response.Write "<meta http-equiv=""Content-Language"" content=""zh-cn"">" & vbCrLf
			Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">" & vbCrLf
			Response.Write "<link href=""../../user/images/css/css.css"" rel=""stylesheet"" type=""text/css"">" & vbCrLf
			Response.Write "</head>" & vbCrLf
			Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">" & vbCrLf
			'Response.Write "<br>" & vbCrLf
			Response.Write "  <table  width=""778"" border=""0"" align=""center"" cellpadding=""2"" cellspacing=""1"">" & vbCrLf
			Response.Write "    <tr>" & vbCrLf
			Response.Write "      <td align=""center""><br>" & vbCrLf
			Response.Write "         <table border=""0"" cellpadding=""2"" cellspacing=""1"" width=""600""  class=""table"">" & vbCrLf
			Response.Write "         <tr class=""bg_tr"">" & vbCrLf
			Response.Write "           <td colspan=2>本站链接信息</td>" & vbCrLf
			Response.Write "         </tr>" & vbCrLf
			Response.Write "         <tr class=""td_bg"">" & vbCrLf
			Response.Write "           <td colspan=""2"" align=""right""><div align=""center"">申请链接交换，务请先使用下面的代码做好本站链接。</div></td>"
			Response.Write "         </tr>" & vbCrLf
			Response.Write "         <tr class=""td_bg"">" & vbCrLf
			Response.Write "           <td width=""176"" height=""25"" align=""right"" valign=""middle""  ><strong>※本站文字链接代码:</strong></td>"
			Response.Write "           <td width=""313"" height=""25""><div align=""center"">演示:<a href=""" & ACTCMS.ActUrl & """ title=""" & AcTCMS.ActCMS_Sys(0) & """ target=""_blank"">" & AcTCMS.ActCMS_Sys(0) & "</a></div></td>"
			Response.Write "         </tr>" & vbCrLf
			Response.Write "         <tr align=""center"" class=""td_bg"">" & vbCrLf
			Response.Write "           <td height=""25"" colspan=""2""> <textarea name=""textlink"" rows=""2"" onMouseOver=""javascript:this.select();"" style=""width:88%;border-style: solid; border-width: 1""><a href=""" & ACTCMS.ActUrl & """ title=""" & AcTCMS.ActCMS_Sys(0) & """ target=""_blank"">" & AcTCMS.ActCMS_Sys(0) & "</a></textarea></td>"
			Response.Write "         </tr>" & vbCrLf
			Response.Write "         <tr class=""td_bg"">" & vbCrLf
			Response.Write "           <td width=""176"" height=""25"" align=""right""><strong>※本站LOGO链接代码:</strong></td>"
			Response.Write "           <td height=""25""><div align=""center"">演示:<a href=""" & ACTCMS.ActUrl & """ title=""" & AcTCMS.ActCMS_Sys(0) & """ target=""_blank""><img src=""" &AcTCMS.ActSys& Replace(AcTCMS.ActCMS_Sys(5),"{$InstallDir}",AcTCMS.ActSys) & """ width=""88"" height=""31"" border=""0"" align=""absmiddle""></a></div></td>"
			Response.Write "         </tr>" & vbCrLf
			Response.Write "         <tr align=""center"" class=""td_bg"">" & vbCrLf
			Response.Write "           <td height=""25"" colspan=""2""> <textarea name=""logolink"" rows=""2"" onMouseOver=""javascript:this.select();"" style=""width:88%;border-style: solid; border-width: 1""><a href=""" & ACTCMS.ActUrl & """ title=""" & AcTCMS.ActCMS_Sys(0) & """ target=""_blank""><img src=""" & AcTCMS.ActCMS_Sys(5) & """ width=""88"" height=""31"" border=""0""></a></textarea></td>"
			Response.Write "         </tr>" & vbCrLf
			Response.Write "       </table>" & vbCrLf
			Response.Write "         <br>" & vbCrLf
 			Response.Write "  <form action=""LinkRegSave.asp"" name=""LinkForm"" method=""post"">" & vbCrLf
			Response.Write "   <input name=""Action"" type=""hidden"" id=""Action"" value=""AddLink"">" & vbCrLf
			Response.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf
			Response.Write "    <tr>" & vbCrLf
			Response.Write "      <td>" & vbCrLf
			Response.Write "        <table border=""0"" align=""center"" cellpadding=""2"" cellspacing=""1"" width=""500""  class=""table"">" & vbCrLf
			Response.Write "           <tr class=""bg_tr"">" & vbCrLf
			Response.Write "             <td  colspan=""2""  height=""25"">申请友情链接</td>" & vbCrLf
			Response.Write "           </tr>" & vbCrLf
			Response.Write "          <tr class=""td_bg"">" & vbCrLf
			Response.Write "            <td width=""18%"" height=""25"" align=""center"">" & vbCrLf
			Response.Write "              <div align=""center"">网站名称：</div></td>" & vbCrLf
			Response.Write "            <td width=""542"" height=""25"">" & vbCrLf
			Response.Write ("<input name=""SiteName"" class=""textbox"" type=""text"" id=""SiteName"" size=""38"" >")
			Response.Write "              <font color=""red"">*</font></td>" & vbCrLf
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr  class=""td_bg"">" & vbCrLf
			Response.Write "            <td height=""25"" align=""center"">所属类别：</td>" & vbCrLf
			Response.Write "            <td height=""25"">" & vbCrLf
			Response.Write "              <select Name=""ClassLinkID"" >" & vbCrLf
						   
						Dim GRS
						Set GRS = actcms.actexe("Select ID,ClassLinkName From ClassLink_Act Order BY AddDate Desc")
						 Do While Not GRS.EOF
							Response.Write ("<Option value=" & GRS(0) & ">" & GRS(1) & "</OPTION>")
						   GRS.MoveNext
						 Loop
						 GRS.Close
						 Set GRS = Nothing
					   
			 Response.Write "             </Select> </td>" & vbCrLf
			 Response.Write "         </tr>"
			 Response.Write "          <tr  class=""td_bg"">" & vbCrLf
			Response.Write "            <td height=""25"" align=""center"">" & vbCrLf
			Response.Write "              <div align=""center"">网站站长：</div></td>" & vbCrLf
			Response.Write "            <td height=""25"">" & vbCrLf
			Response.Write "              <input name=""Webadmin"" class=""textbox"" type=""text"" size=""38""> <font color=""red"">*</font></td>"
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr  class=""td_bg"">" & vbCrLf
			Response.Write "            <td height=""25"" align=""center"">" & vbCrLf
			Response.Write "              <div align=""center"">站长信箱：</div></td>" & vbCrLf
			Response.Write "            <td height=""25"">" & vbCrLf
			Response.Write "              <input name=""Email"" type=""text"" class=""textbox"" size=""38"" value=""@"" ></td>" & vbCrLf
			Response.Write "          </tr>" & vbCrLf
					  
					 
			
			Response.Write "          <tr  class=""td_bg"">"
			Response.Write "            <td height=""25"" align=""center"">网站地址：</td>" & vbCrLf
			Response.Write "            <td height=""25""><input name=""Url"" class=""textbox"" type=""text""  value=""http://"" id=""Url"" size=""38""> <font color=""red"">*</font></td>"
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr  class=""td_bg"">"
			Response.Write "            <td height=""25"" align=""center"">链接类型：</td>" & vbCrLf
			Response.Write "            <td height=""25"">"
			Response.Write ("<input type=""radio"" name=""LinkType"" onclick=""SetLogoArea('none')"" value=""0"" checked> 文字链接： ")
			Response.Write ("<input type=""radio"" name=""LinkType"" onclick=""SetLogoArea('')"" value=""1"">  LOGO链接： ")
					   
			Response.Write "             </td>" & vbCrLf
			Response.Write "          </tr>"
			Response.Write "         <tr  class=""td_bg"" Style=""display:none"" ID=""LinkArea"">" & vbCrLf
			Response.Write "            <td height=""25"" align=""center"">Logo地址：</td>" & vbCrLf
			Response.Write "            <td height=""25""><input name=""Logo"" class=""textbox"" type=""text""  value=""http://"" id=""Logo"" size=""38""> <font color=""red"">*</font></td>"
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr  class=""td_bg"">" & vbCrLf
			Response.Write "            <td height=""25"" align=""center"">"
			Response.Write "              <div align=""center"">网站简介：</div></td>" & vbCrLf
			Response.Write "            <td height=""25"">" & vbCrLf
			Response.Write "              <textarea name=""Description"" rows=""6"" id=""Description"" style=""width:80%;border-style: solid; border-width: 1""></textarea></td>"
			Response.Write "          </tr>" & vbCrLf
			Response.Write "          <tr  class=""td_bg"">" & vbCrLf
			Response.Write "            <td height=""25"" align=""center"">验 证 码：</td>" & vbCrLf
			Response.Write "            <td height=""25"">" & vbCrLf
			Response.Write "			<input type=""text"" size=""10"" name=""Code""> <img style=""cursor:hand;""  src="""&ACTCMS.ActSys&"ACT_INC/Code.asp?s='+Math.random();"" id=""IMG1"" onclick=this.src="""&ACTCMS.ActSys&"ACT_INC/Code.asp?s='+Math.random();"" alt=""看不清楚? 换一张！"">"
			Response.Write "          </td></tr>" & vbCrLf
			Response.Write "        </table>" & vbCrLf
			Response.Write "       </td>"
			Response.Write "    </tr>" & vbCrLf
			Response.Write "    </table>" & vbCrLf
			Response.Write "  <table width=""100%"" height=""38"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf
			Response.Write "    <tr>" & vbCrLf
			Response.Write "      <td height=""40"" align=""center"">" & vbCrLf
			Response.Write "        <input type=""button""  class=""ACT_btn""  name=""Submit"" Onclick=""CheckForm()"" value="" 确 定 "">" & vbCrLf
			Response.Write "        <input type=""reset""  class=""ACT_btn""  name=""Submit2""  value="" 重 填 "">" & vbCrLf
			Response.Write "      </td>" & vbCrLf
			Response.Write "    </tr>" & vbCrLf
			Response.Write "  </table>" & vbCrLf
			Response.Write "  </form>" & vbCrLf
			Response.Write "<Script Language=""javascript"">" & vbCrLf
			Response.Write "<!--" & vbCrLf
			Response.Write "function is_email(str)" & vbCrLf
			Response.Write "{ if((str.indexOf('@')==-1)||(str.indexOf('.')==-1)){" & vbCrLf
			Response.Write "    return false;" & vbCrLf
			Response.Write "    }" & vbCrLf
			Response.Write "    return true;" & vbCrLf
			Response.Write "}" & vbCrLf
			Response.Write "function SetLogoArea(Value)" & vbCrLf
			Response.Write "{"
			Response.Write "   document.all.LinkArea.style.display=Value;"
			Response.Write "}" & vbCrLf
			Response.Write "function CheckForm()" & vbCrLf
			Response.Write "{ var form=document.LinkForm;" & vbCrLf
			Response.Write "   if (form.SiteName.value=='')" & vbCrLf
			Response.Write "    {"
			Response.Write "     alert(""请输入网站名称!"");" & vbCrLf
			Response.Write "     form.SiteName.focus();" & vbCrLf
			Response.Write "     return false;" & vbCrLf
			Response.Write "    }" & vbCrLf
			Response.Write "   if (form.Webadmin.value=='')" & vbCrLf
			Response.Write "    {"
			Response.Write "     alert(""请输入网站站长!"");" & vbCrLf
			Response.Write "     form.Webadmin.focus();" & vbCrLf
			Response.Write "     return false;" & vbCrLf
			Response.Write "    }" & vbCrLf
			Response.Write "    if ((form.Email.value!='@')&&(is_email(form.Email.value)==false))" & vbCrLf
			Response.Write "    {"
			Response.Write "    alert('非法电子邮箱!');" & vbCrLf
			Response.Write "     form.Email.focus();" & vbCrLf
			Response.Write "     return false;" & vbCrLf
			Response.Write "    }" & vbCrLf
			


			Response.Write "   if (form.Url.value=='' || form.Url.value=='http://')" & vbCrLf
			Response.Write "    {" & vbCrLf
			Response.Write "     alert(""请输入网站地址"");" & vbCrLf
			Response.Write "     form.Url.focus();" & vbCrLf
			Response.Write "     return false;" & vbCrLf
			Response.Write "    }" & vbCrLf
			Response.Write "   if (form.Code.value=='')" & vbCrLf
			Response.Write "    {" & vbCrLf
			Response.Write "     alert(""请输入认证码!"");" & vbCrLf
			Response.Write "     form.Code.focus();" & vbCrLf
			Response.Write "     return false;" & vbCrLf
			Response.Write "    }" & vbCrLf
			Response.Write "    form.submit();" & vbCrLf
			Response.Write "    return true;" & vbCrLf
			Response.Write "}" & vbCrLf
			Response.Write "//-->" & vbCrLf
			Response.Write "</Script>" & vbCrLf
			Response.Write " </td></tr></table>" & vbCrLf
		'	Response.Write "         <br>" & vbCrLf
		'	Response.Write "       </td>" & vbCrLf
		'	Response.Write "     </tr>" & vbCrLf
		'	Response.Write "   </table>" & vbCrLf
			Response.Write " </form>" & vbCrLf
			Response.Write " </body>" & vbCrLf
			Response.Write " </html>" & vbCrLf

%> 
