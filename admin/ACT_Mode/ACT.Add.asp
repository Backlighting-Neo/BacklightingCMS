<!--#include file="../ACT.Function.asp"-->
<!--#include file="../Mode/ACT.M.ASP"-->
<!--#include file="../include/ACT.F.ASP"-->
<!--#include file="../../act_inc/ACT.Code.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=ModeName%>内容添加 By Act</title>
<meta http-equiv="X-UA-Compatible" content="IE=8" />
<link href="../Images/editorstyle.css" rel="stylesheet" type="text/css">
<script charset="utf-8"  language="JavaScript" type="text/javascript" src="../../editor/kindeditor/kindeditor.js" ></script>
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/lhgcore.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/lhgdialog.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/Main.js"></script>
<script type="text/javascript" src="../../ACT_INC/js/swfobject.js"></script>
<script type='text/javascript' src='../../ACT_INC/js/time/WdatePicker.js'></script>
 <SCRIPT LANGUAGE='JavaScript'>
 var U="<%=ACTCMS.strToAsc(RSQL(Trim(Request.Cookies(AcTCMSN)("AdminName"))))%>";
var P="<%=ACTCMS.strToAsc(RSQL(Trim(Request.Cookies(AcTCMSN)("AdminPassword"))))%>";
 

  
</SCRIPT>
</head>
<% 
 Dim FileName,ClassName,ClassID,Title,IntactTitle,Intro,Keywords,CopyFrom,Slide,TemplateUrl
  Dim ATT,IStop,ModeID,ModeName,Action,ID,ActCMS_DIY,SavePic,MakeHtml,KeywordArr
  Dim actlink,rev,Content,Hits,Straction,PicUrl,UpdateTime,ActCMS_DIY2,author
  Dim ChargeType,InfoPurview,arrGroupID,ReadPoint,PitchTime,ReadTimes,DividePercent
  ModeID = ChkNumeric(Request("ModeID"))
  if ModeID=0 or ModeID="" Then ModeID=ACTCMS.ACT_L(Request.QueryString("ClassID"),10)
  if ModeID=0 or ModeID="" Then ModeID=1
  ActCMS_DIY=Split(AF.ActCMS_DIY_F(ModeID,1,""),"§") 
  ModeName= ACTCMS.ACT_C(ModeID,1) 
  Action=Request.QueryString("Action")
  ID=ChkNumeric(Request("ID"))
	KeywordArr=Split(ActCMS_DIY(4), "-")
	Dim TitleTypeList,i,KeywordsArr,AuthorArr,rs3,Score_ACT
	StrAction = request("Action")
	IF Straction = "edit" Then Straction = "edit"  Else Straction = "add"
	Dim Rs,ShowErr
	Set Rs=server.CreateObject("adodb.recordset") 
	IF Action = "edit" Then
	  If UBound(Split(ActCMS_DIY(2),"-"))=3  Then 
			ActCMS_DIY2=Split(ActCMS_DIY(2),"-")
			SavePic=ActCMS_DIY2(2)
			MakeHtml=ActCMS_DIY2(3)
	  Else
			SavePic=0
			MakeHtml=1
	  End If 
 		Rs.OPen "Select * from  "&ACTCMS.ACT_C(ModeID,2)&"  where ID = "& ChkNumeric(Request.QueryString("ID")) &"",Conn,1,1
		IF Rs.eof Then 
			ShowErr = "数据查询错误"
			Call Actcms.ActErr(ShowErr,"","1")
			Response.End
		Else
			actlink=RS("actlink")
			Title=RS("Title")
			rev=RS("rev")
			Slide = Rs("Slide")
			IntactTitle=RS("IntactTitle")
			If RS("Content") <> "" Then Content=Server.HTMLEncode(RS("Content"))
			Intro=RS("Intro")
			Keywords=RS("Keywords")
 			CopyFrom=RS("CopyFrom")
			Hits=RS("Hits")
			FileName=RS("FileName")
			PicUrl=RS("PicUrl")
			'arrGroupID=RS("arrGroupID")
			'Score_ACT=RS("Score_ACT")
			ATT=RS("ATT")
			IStop=RS("IStop")
			TemplateUrl=RS("TemplateUrl")
			ClassID=RS("ClassID")
			UpdateTime=RS("UpdateTime")
			author=rs("author")
			ReadPoint =RS("ReadPoint")
			ChargeType=RS("ChargeType")
			PitchTime =RS("PitchTime")
			ReadTimes =RS("ReadTimes")
			InfoPurview=RS("InfoPurview")
			arrGroupID =RS("arrGroupID")
			DividePercent=RS("DividePercent")
 			Rs.close
		If Not ACTCMS.ACTCMS_QXYZ(ModeID,"3",ClassID) Then   Call Actcms.Alert("对不起，你没有"&ACTCMS.ACT_C(ModeID,1)&"的修改权限","")
		End IF
	Else
	   ATT=0:IStop=0:Slide=0:Hits=0
 	   CopyFrom=session("CopyFrom")
	   UpdateTime=now()
		ReadPoint =0
		ChargeType=0
		PitchTime =0
		ReadTimes =0
		InfoPurview=0
		arrGroupID =0
		DividePercent=0

	  If UBound(Split(ActCMS_DIY(2),"-"))=3  Then 
			ActCMS_DIY2=Split(ActCMS_DIY(2),"-")
			IStop=ActCMS_DIY2(0)
			rev=ActCMS_DIY2(1)
			SavePic=ActCMS_DIY2(2)
			MakeHtml=ActCMS_DIY2(3)
	  Else
			IStop=0
			rev=1
			SavePic=0
			MakeHtml=1
	  End If 
	   IntactTitle=ActCMS_DIY(0)
	   If UBound(KeywordArr)=>0 Then Keywords=KeywordArr(0)
	   author=ActCMS_DIY(6)
	   CopyFrom=ActCMS_DIY(8)
	   Intro=ActCMS_DIY(10)
	   arrGroupID=Replace(ActCMS_DIY(14),"-",",")
	   Score_ACT=ActCMS_DIY(16)
	   Dim ACTCode
	   Set ACTCode =New ACT_Code
	   If Trim(ActCMS_DIY(13)) <> "" Then Content =Server.HTMLEncode(ACTCode.LoadTemplate(ActCMS_DIY(13)))
		IF Request.QueryString("ClassID") <> "" Then
			Rs.OPen "Select ClassName,ClassID,ModeID from Class_Act  where ClassID = '"& Request.QueryString("ClassID") &"' and ActLink =1",Conn,1,1
			If Not rs.eof Then 
				ClassName = Rs("ClassName")
				ClassID = Rs("ClassID")
				ModeID=Rs("ModeID")
			Else
				response.write "程序出现错误,该栏目是外部栏目,或该栏目已被删除"
				response.End 
			End If 
		Else
			IF Session("ClassName") <> "" And Session("ClassID") <> "" Then
				ClassName = Session("ClassName")
				ClassID = Session("ClassID")
			End IF
		End IF
	End If
  With Response 
%>
<body>
<table  width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
<form  name="tcjdxr" id="act" method="post" action="ACT.Save.asp?action=<%=Straction%>&ID=<%=Request("ID")%>&actcms=<%=Request.QueryString("actcms")%>">
<input type="hidden" id="ModeID" name="ModeID" value="<%=ModeID%>">
  <tr>
    <td colspan="3" align="left" class="bg_tr">您现在的位置：<%=ModeName%>管理 &gt;&gt; 添加<%=ModeName%></td>
  </tr>
  <tr>
    <td width="15%" height="23" align="right" class="tdclass">所属栏目：</td>
    <td height="23" colspan="2" align="left"  class="tdclass">

<select name="ClassID">
<option>请选择栏目</option>
      <%= AF.ForClasslist(ModeID)%>
    </select>  
			自定义属性：<select name="ATT" id="ATT">
					<option value="0">普通<%=ModeName%></option>
			<%=ACTCMS.ACT_ATT(ATT)%></select>
			<span class="h" style="cursor:help;" title="点击显示帮助" id="Article_002" onClick="dohelp('Article_002')">帮助</span>
	  <input <%if actlink= 1 Then Response.Write "Checked "%> name="actlink" type="checkbox" id="actlinkss" value="1"  onclick="actlinks()"/>
	<label for="actlinkss"><font color="#FF0000"><b>使用转向链接</b></font></label>
	<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Article_001')"  id="Article_001">帮助</span></td>
  </tr>
  <tr>
    <td height="23" align="right"  class="tdtop" >简短标题：</td>
    <td  height="23" colspan="2" align="left" bgcolor="#FFFDEC" class="tdtop">
      <input name="Title" type="text"  class="title" id="Title" value="<%=Server.HTMLEncode(Title)%>" size="45" maxlength="255"  style="overflow-x:visible;overflow-y:visible;" >
		<span id="msg1"></span><select name="select" onChange="FormatTitle(this, tcjdxr.Title, '')">
          <option selected>-- 修饰 --</option>
          <option value="1">粗体</option>
          <option value="3">红色</option>
          <option value="4">蓝色</option>
          <option value="5">倾斜</option>
          <option value="2">[图文]</option>
          <option value="-1">清除样式</option>
          <option value="-2">清除内容</option>
      </select><font color=green>支持HTML</font>  <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Article_title')"  id="Article_title">帮助</span>
	  </td>
  </tr>
  <%If ActCMS_DIY(1)="0" Then %>
  <tr>
    <td height="23" align="right" class="tdclass">完整标题：</td>
    <td height="23" colspan="2" align="left"  class="tdclass">
    <input name="IntactTitle"  type="text"  class="Ainput" id="IntactTitle" value="<%= IntactTitle %>" size="60" /></td>
  </tr>
  <%
  Else 
	response.write "<input type=""hidden"" id=""IntactTitle"" name=""IntactTitle"" value="""&IntactTitle&""">"
  End If
%>

   <tr>
    <td height="23" align="right" class="tdclass">缩略图：</td>
    <td height="23" colspan="2" align="left"  class="tdclass">
    <input name="PicUrl" type="text"  class="Ainput" id="PicUrl" value="<%= picurl %>" size="40" />
   <input id="PicUrls" class="ACT_btn" onClick="javascript:upload(<%=ActCMS_DIY(16)%>,<%=ModeID%>,'PicUrl');" name="button" type="button"   value="点击上传图片">
    
     <input <% if Slide=1 then Response.Write  "Checked"%>   id="Slide" type="checkbox" value="1" name="Slide" /><label for=Slide>&nbsp;&nbsp;<font color="green">是否幻灯片</font></label>
     		<label for="slt1"> 
<input id="slt1"  <%If Straction="add" Then response.write "Checked"%>   type="radio" name="slt" value="0" onClick="sltA(1);">
        系统自动截图</label>
  <label for="slt2"> 
 <input id="slt2"   type="radio" name="slt" <%If Straction="edit" Then response.write "Checked"%> value="1" onClick="sltA(2);">手工截图</label>	
		<span
		<%If Straction="add" Then %>
		style='display:none'
		<%End If %>
		id='sgjt'>
	 <a href="javascript:" style="color:red;"  onClick="javascript:cutimg();">[手工剪裁图片]</a></span>
 <SCRIPT LANGUAGE="JavaScript">
<!--

function upload(instr,ModeID,iname) 
{
  ( new J.dialog({ id:'actcmssc'+iname+'s' ,title:'ACTCMS上传', loadingText:'上传加载中...', page: '<%=actcms.actsys&actcms.adminurl%>/include/Upload_Admin.asp?A=add&instr='+instr+ "&ModeID="+ModeID+ "&instrname="+iname+ "&" + Math.random(),  width:720, height:160 })).ShowDialog();
 }

 function get_obj(obj){
   return document.getElementById(obj);
}
//-->
</SCRIPT>
 <script language="javascript">
 
  function upfile(url,name){
	document.getElementById(name).value=url;
 }
function cutimg() 
{
var img=get_obj("PicUrl").value;
if(img!=''){
 ( new J.dialog({ id:'actcmscj' ,title:'ACTCMS上传', loadingText:'上传加载中...', page: '<%=actcms.actsys&actcms.adminurl%>/ACT_Mode/eimg/img.asp?url='+img+ "&" + Math.random(), width:820, height:620 })).ShowDialog();

}else{
	get_obj("PicUrl").focus();
	alert('图片地址不存在');
}
}
</script></td>
  </tr>

 <% If ActCMS_DIY(3)="0" Then
  %>
  <tr>
    <td height="23" align="right" class="tdclass"><%=ModeName%>属性：</td>
    <td height="23" colspan="2" align="left"  class="tdclass"><table width="100%" border="0">
                  <tr>
                    <td  class="tdclass">
					<input <%if IStop= "1" Then Response.Write "Checked "%> name="IStop" type="checkbox"  id="IStop" value="1">
					<label for="IStop">置顶&nbsp;&nbsp;</label>
					<input <%if rev= "1" Then Response.Write "Checked "%> id="rev" type="checkbox" value="1" name="rev" />
                      <label for="rev">允许评论&nbsp;&nbsp;</label>
                       <input <%if MakeHtml= "1" Then Response.Write "Checked "%>  id="MakeHtml" type="checkbox"  value="1" name="MakeHtml" />
                      <label for="makehtml">立即生成&nbsp;&nbsp;</label>
                      <input <%if SavePic= "1" Then Response.Write "Checked "%>  id="SavePic" type="checkbox" value="1" name="SavePic" />
                      <label for="savepic">远程存图&nbsp;(是否将文章中的外部图片采集回来,影响速度)</font></label>
                     
					  <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Article_003')"  id="Article_003">帮助</span></td>
                  </tr>
                </table>	  </td>
  </tr>
  <%
  else
  	response.write "<input type=""hidden"" id=""IStop"" name=""IStop"" value="""&IStop&""">"
  	response.write "<input type=""hidden"" id=""rev"" name=""rev"" value="""&rev&""">"
  	response.write "<input type=""hidden"" id=""SavePic"" name=""SavePic"" value="""&SavePic&""">"
  	response.write "<input type=""hidden"" id=""MakeHtml"" name=""MakeHtml"" value="""&MakeHtml&""">"
   End if%>
  <tr class="tdclass" id=ChangesUrl <% 	
  if actlink = 1 Then 
	 .Write "style=""'DISPLAY: none'"""
	else
	 .Write "style=""DISPLAY: none"""
	End if
 %> >
    <td height="23" colspan="3" align="right">
		 <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
			  <tr> 
				  <td width="13%" align="right" class="bg_tr"><div align="center">转向链接地址：</div>			 	 </td>
			    <td width="87%"  class="tdclass"><input name="ChangesLinkUrl" type="text"  class="Ainput" id="ChangesLinkUrl" value="<%=FileName%>" size="50" maxlength="255"> 
			   <font color="#ff0000">如果<%=ModeName%>是转向链接，那么以下各项填写无效,请不要填写!</font></td>
	    </tr></table>   </td>
    </tr> 
	<%  
	If ActCMS_DIY(5)="0" Then
%>
	<tr>
    <td height="23" align="right" class="tdclass">关键字：</td>
    <td height="23" colspan="2" align="left"  class="tdclass">
	<input name="Keywords" type="text"  class="Ainput" size="50" maxlength="255" id="Keywords" value="<%= Keywords %>">
	<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Article_005')"  id="Article_005">帮助</span>
	<input id=Tags type="checkbox" name="Tags" value="1"><label for=Tags>写入Tags</label>
	<%
		  For I = 0 To UBound(KeywordArr)
			.Write  "<a href=""javascript:SetValue('add','Keywords','"&","&KeywordArr(I) &"')"">【<font color=""blue"">" & KeywordArr(I) & "</font>】</a>"
		  Next
	%>	</td>
  </tr>
  <%Else
		response.write "<input type=""hidden"" id=""Keywords"" name=""Keywords"" value="""&Keywords&""">"
	End If 
 	If ActCMS_DIY(7)="0" Then
  %>
  <tr>
    <td height="23" align="right"  class="tdclass"><%=ModeName%>作者：</td>
    <td height="23" colspan="2" align="left"  class="tdclass">
	<input name="author" type="text"  class="Ainput"  value="<%=author %>">
	<select name="" onChange="document.tcjdxr.author.value=this.value">
	<option>常用作者列表</option>
	<%   set rs3=ACTCMS.ActExe("Select Field1,Field2 from  AC_ACT where Types=1")
		 if Not rs3.eof Then
		 Do While Not rs3.eof
			 .Write  "<option value="&rs3("Field1") &">" & rs3("Field1") & "</option>"
			  rs3.movenext
		 Loop
		 rs3.close:set rs3=Nothing
	 End if
 %></select><input type="checkbox" id="addauthor" name="addauthor" value="1" checked>
 <label for="addauthor" title="添加格式(以-分割)： 作者-作者电子信箱">添加至作者列表中</label> 
 <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Article_006')"  id="Article_006">帮助</span></td>
  </tr>
  <%Else
  	response.write "<input type=""hidden"" id=""author"" name=""author"" value="""&author&""">"
  	response.write "<input type=""hidden"" id=""addauthor"" name=""addauthor"" value=""1"">"
    End If
 	If ActCMS_DIY(9)="0" Then
  %>
  <tr>
    <td height="23" align="right" class="tdclass"><%=ModeName%>来源：</td>
    <td height="23" colspan="2" align="left"  class="tdclass">
	<input name="CopyFrom" type="text"  class="Ainput" value="<%= CopyFrom %>">
	<select name="" onChange="document.tcjdxr.CopyFrom.value=this.value">
	<option>常用来源列表</option>
	<% Set  rs3=ACTCMS.ActExe("Select Field1,Field2 from  AC_ACT where Types=0")
	  If Not  rs3.eof Then
		 Do While Not rs3.eof
			.Write  "<option value="&rs3("Field1") &">" & rs3("Field1") & "</option>"
			rs3.movenext
		 loop
		End if
	 %></select><input type="checkbox" id="addCopyFrom" name="addCopyFrom" value="1" checked>
 <label for="addCopyFrom" title="添加格式(以-分割)： 来源-来源网站地址">添加至来源列表中</label>
 <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Article_007')"  id="Article_007">帮助</span></td>
  </tr>
  <%Else
		response.write "<input type=""hidden"" id=""CopyFrom"" name=""CopyFrom"" value="""&CopyFrom&""">"
  		response.write "<input type=""hidden"" id=""addCopyFrom"" name=""addCopyFrom"" value=""1"">"
    End If
    If ActCMS_DIY(11)="0" Then%>
  <tr>
    <td height="23" align="right" class="tdclass"><%=ModeName%>导读：</td>
    <td height="23" colspan="2" align="left"  class="tdclass">
	<textarea name="Intro" cols="75" rows="4" id="Intro" class="tabcontent"><%= Intro %></textarea>
	 <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Article_008')"  id="Article_008">帮助</span></td>
  </tr>
   <% Else
		response.write "<input type=""hidden"" id=""Intro"" name=""Intro"" value="""&Intro&""">"
	  End If
	  If actcms.actexe("select count(id) from Table_ACT where modeid="&modeid)(0)<>"0" Then %>
   
      <% 	 '
If Action="edit" Then 
	If actcms.ACT_C(ModeID,4)="0" Then 
 	    response.write M.ACT_MXEdit(ModeID,ID) 
	Else
		response.write  M.ReplaceFormEdit(ModeID,id,actcms.LTemplate(actcms.ACTSYS&"act_inc/cache/"&ModeID&"/"&ACTCMS.ACT_C(1,2)&"-mode.inc"))
	 
    End If 

   Else 
	If actcms.ACT_C(ModeID,4)="0" Then 

	   response.write   M.ACT_NoRormMXList(ModeID)
	Else
		response.write  M.ReplaceForm(ModeID,actcms.LTemplate(actcms.ACTSYS&"act_inc/cache/"&ModeID&"/"&ACTCMS.ACT_C(1,2)&"-mode.inc"))
	 
    End If 
   End If %> 
    
	<%End If
	
	If ActCMS_DIY(25)="0" Then 
	%>
 <tr>
<td  height="23" align="right"  class="tdclass">附加选项：</td>
<td    class="tdclass">
 
  <%If ActCMS_DIY(20)="0" Then %>
	 <input  <%if ActCMS_DIY(19)= "1" Then Response.Write "Checked "%>  type="checkbox" id="dellink" name="dellink" value="1"><label for="dellink">除去内容中的超级链接</label>
	 <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Article_dellink')"  id="Article_dellink">帮助</span>
	 <%End If %>	
    
</td>
</tr>

<%
End If 
If ActCMS_DIY(21)="0" Then %>
<tr>
<td  height="23" align="right"  class="tdclass">批量上传文件：</td>
<td     class="tdclass">
 

<div id="sapload">
    
    </div>
 
 <script type="text/javascript">
// <![CDATA[
var so = new SWFObject("<%=ACTCMS.ACTSYS%>act_inc/sapload.swf", "sapload", "450", "25", "9", "#ffffff");
so.addVariable('types','<%=Replace(ACTCMS.ActCMS_Sys(11),"/",";")%>');
so.addVariable('isGet','1');
so.addVariable('args','myid=Upload;ModeID=<%=ModeID%>;U='+U+";"+';P='+P+";"+'Yname=content1');
so.addVariable('upUrl','<%=ACTCMS.ACTSYS%><%=ACTCMS.ActCMS_Sys(8)%>/include/Upload.asp');
so.addVariable('fileName','Filedata');
so.addVariable('maxNum','110');
so.addVariable('maxSize','<%=ACTCMS.ActCMS_Sys(10)/1024%>');
so.addVariable('etmsg','1');
so.addVariable('ltmsg','1');
so.addParam('wmode','transparent');
so.write("sapload");
function sapLoadMsg(t){
var actup=t.split('|');
 {
  	   KE.insertHtml(actup[0], actup[1]);
}
}

// ]]>
</script> 
 
 
 
</td>
</tr>

<%End If 

If ActCMS_DIY(25)="0" Then

%> <tr>


    <td height="23" align="right"  class="tdclass"><%=ModeName%>内容：<br>
<input name="button"  type="button"  class="ACT_btn" style="cursor:pointer" onClick="insertHTMLToEditor('[NextPage]','content1');" value="[NextPage]"><br><br>
注：手动分页符标记为：点击插入，注意大小写</td>
    <td height="23" colspan="2"    class="tdclass">
 		<script>
			KE.show({
				id : 'content1'
  			});
		</script>
	 
	   <textarea id="content1" name="content"  style="width:98%;height:300px;visibility:hidden;">
<%=content%>
</textarea>
	</td>
  </tr>

<%  Else
		response.write "<input type=""hidden"" id=""Content"" name=""Content"" value="""&Content&""">"
	  End If

if ActCMS_DIY(18)="0"  Or ActCMS_DIY(20)="0" Then %>
  <tr>
    <td height="23" align="right" class="tdclass">附加选项：</td>
    <td height="23" colspan="2" align="left"  class="tdclass">
	<%If ActCMS_DIY(18)="0" Then %>
	<input <%IF Request.QueryString("Action") = "edit"  Then Response.write " disabled=true "%> name="FileName" type="text"  class="Ainput" id="FileName" value="<% =FileName%>" /> 生成的文件名
      <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Article_009')"  id="Article_009">帮助</span>
	 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
	   <%End If %>
	   </td>
    </tr>
  <tr>
  <%End If %>
    <td height="23" align="right" class="tdclass">添加日期：</td>
    <td height="23" colspan="2" align="left"  class="tdclass">
	<input name="UpdateTime"  type="text"  class="Ainput" id="UpdateTime" value="<%=UpdateTime%>"   onClick="WdatePicker()" />
	  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;初始点击数：
      <input name="Hits" type="text"  class="Ainput" id="Hits" value="<%= Hits %>" />
	  <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Article_010')"  id="Article_010">帮助</span></td>
    </tr>
	<%if ActCMS_DIY(15)="0" Then %>
  
 


   <tr  >
    <td align='right'   class="tdclass">权限：
	 <label for="ydqxs1"><input onClick=ydqxset(0) type="radio"    value="0" name="ydqx"  id="ydqxs1">展开</label>
      <label for="ydqxs2"><input onClick=ydqxset(1) type="radio"  Checked  value="1" name="ydqx" id="ydqxs2">收缩</label></td>
    <td height='30' class="tdclass" >
    <label for="InfoPurview1">
    <input name='InfoPurview' id='InfoPurview1' type='radio' <% if  InfoPurview="0" then response.Write "checked=""checked""" %> value='0'  />
      继承栏目权限（当所属栏目为认证栏目时，建议选择此项）</label><br />
      <label for="InfoPurview2"><input name='InfoPurview' id='InfoPurview2' <% if  InfoPurview="1" then response.Write "checked=""checked""" %>  type='radio' value='1' />
      所有会员（当所属栏目为开放栏目，想单独对某些文章进行阅读权限设置，可以选择此项）</label><br />
     <label for="InfoPurview3"> <input name='InfoPurview' id='InfoPurview3' <% if  InfoPurview="2" then response.Write "checked=""checked""" %>  type='radio' value='2' />
      指定会员组（当所属栏目为开放栏目，想单独对某些文章进行阅读权限设置，可以选择此项,<font color="green">在下面设置相应的会员组权限</font>）</label><br />
      <table border='0' width='90%'>
        <tr>
          <td><%= actcms.GetGroup_CheckBox("arrGroupID",arrGroupID,5)  %></td>
        </tr>
      </table></td>
  </tr>
  <tr   id="ydqx2" style="DISPLAY: none" >
    <td align='right'   height="30"  class="tdclass"><strong>阅读点数： </strong></td>
    <td height='30'  class="tdclass">&nbsp;
        <input  name='ReadPoint' type='text' id='ReadPoint'  value='<%=ReadPoint  %>' size='6' class='Ainput' />
      免费阅读请设为 &quot;<font color="red">0</font>&quot;，否则有权限的会员阅读此文章时将消耗相应点数，游客将无法阅读此文章 </td>
  </tr>
  <tr   id="ydqx3"   style="display:none">
    <td align='right'    height="30"  class="tdclass"><strong>重复收费：</strong><br>
只有当上述计费才有效</td>
    <td height='30'  class="tdclass">
     <label for="ChargeType1">
     <input name='ChargeType'  id='ChargeType1' type='radio' value='0' <% if  ChargeType="0" then response.Write "checked=""checked""" %>  />
      不重复收费(如果需扣点数文章，建议使用)</label><br />
       <label for="ChargeType2"><input name='ChargeType' id='ChargeType2' <% if  ChargeType="1" then response.Write "checked=""checked""" %> type='radio' value='1' />
      距离上次收费时间
      <input name='PitchTime' type='text' class='Ainput' value='<%= PitchTime %>' size='8' maxlength='8'  />
      小时后重新收费</label><br />
     <label for="ChargeType3"> <input name='ChargeType' <% if  ChargeType="2" then response.Write "checked=""checked""" %> id='ChargeType3'  type='radio' value='2' />
      会员重复阅读此文章
      <input name='ReadTimes' type='text' class='Ainput' value='<%= ReadTimes %>' size='8' maxlength='8'  />
      页次后重新收费</label><br />
     <label for="ChargeType4"> <input name='ChargeType' <% if  ChargeType="3" then response.Write "checked=""checked""" %> id='ChargeType4'  type='radio' value='3' />
      上述两者都满足时重新收费</label><br />
      <label for="ChargeType5"><input name='ChargeType' <% if  ChargeType="4" then response.Write "checked=""checked""" %> id='ChargeType5'  type='radio' value='4' />
      上述两者任一个满足时就重新收费</label><br />
      <label for="ChargeType6"><input name='ChargeType' type='radio' <% if  ChargeType="5" then response.Write "checked=""checked""" %>  id='ChargeType6'  value='5' />
      每阅读一页次就重复收费一次（建议不要使用,多页文章将扣多次点数）</label> </td>
  </tr>
  <tr   id="ydqx4"   style="display:none">
    <td align='right'   height="30"  class="tdclass"><strong>分成比例： </strong></td>
    <td height='30'  class="tdclass">&nbsp;
        <input name='DividePercent' type='text' id='DividePercent'  value='<%= DividePercent %>' size='6' class='Ainput' />
      % 　如果比例大于0，则将按比例把向阅读者收取的点数支付给投稿者 </td>
  </tr>
  
  <%
  Else
  		response.write "<input type=""hidden"" id=""arrGroupID"" name=""arrGroupID"" value="""&arrGroupID&""">"
  End If
  %>
  <tr>
    <td height="23" align="right" class="tdclass">模板地址：</td>
    <td height="23" colspan="2" align="left"  class="tdclass">
	<input <% IF TemplateUrl = "" Then Response.Write " checked=""checked""" %> onclick='Templates(this.value);' id="Templates2"   type="radio" name="TemplateUrl`" value="1" /><label for="Templates2">继承栏目设定</label>
      <input <% IF TemplateUrl <> "" Then Response.Write " checked=""checked""" %>  onclick='Templates(this.value);' id="Templates1" type="radio" name="TemplateUrl`" value="2" /><label for="Templates1">自定义</label>
      <div id='Templatefs' 
	   <% 	
  if TemplateUrl <> "" Then 
	 .Write "style=""'DISPLAY: none'"""
	else
	 .Write "style=""DISPLAY: none;text-align: left;"""
	End if
 %>  >
       <input class="Ainput" name='TemplateUrl'  size="30"  value='<%=TemplateUrl%>' />
          &nbsp;
          <input class="ACT_btn" type="button"  onClick="OpenWindowAndSetValue('../include/print/SelectPic.asp?CurrPath=<%=ACTCMS.ActSys%><%=actcms.SysThemePath&"/"&actcms.NowTheme%>',500,320,window,document.tcjdxr.TemplateUrl);" value="选择模板..."> 
	  </div><span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('Article_013')"  id="Article_013">帮助</span></td>
  </tr>

  <tr>
    <td height="23" colspan="3" align="center"  class="tdclass">
	
    <table width="500" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center"> 
<input type=button onclick=CheckForm() class="ACT_btn"  name=Submit1 value="  保存  " />
      &nbsp;&nbsp; <input type="reset" class="ACT_btn" name="Submit2" value="  重置  ">        </td>
  </tr>
</table>
    
    
       
       </td>
    </tr>
</form></table>
<br><br>
<p>

<% end With
	
%>
<script language="JavaScript" type="text/javascript">

 function ydqxset(n){
	if (n == 0){
 		ydqx2.style.display='';
		ydqx3.style.display='';
		ydqx4.style.display='';
	}
	else{
 		ydqx2.style.display='none';
		ydqx3.style.display='none';
		ydqx4.style.display='none';
	}
} 
 function sltA(n){
			if (n==1){
			sgjt.style.display='none';
			}
		  else{
			sgjt.style.display='';
		}
}

function FormatTitle(obj, obj2, def_value)
{
    var FormatFlag = obj.options[obj.selectedIndex].value;
    var tmp_Title = FilterHtmlStr(obj2.value);
    switch(FormatFlag)
    {
        case "1" :
            obj2.value = "<b>" + tmp_Title + "</b>";
            break;
        case "2" :
            obj2.value = tmp_Title + "[图文]";
            break;
        case "3" :
            obj2.value = "<font color=\"red\">" + tmp_Title + "</font>";
            break;
        case "4" :
            obj2.value = "<font color=\"blue\">" + tmp_Title + "</font>";
            break;
        case "5" :
            obj2.value = "<em>" + tmp_Title + "</em>";
            break;
        case "-1" :
            if(confirm("确定要清除样式?"))
            {
                obj2.value = tmp_Title;
            }
            break;
        case "-2" :
			if(confirm("确定要清除内容?"))
            {
            obj2.value = def_value;
            break;
			}
    }
    obj.selectedIndex = 0;
}
function FilterHtmlStr(str)
{
    str = str.replace(/<.*?>/ig, "");
    return str;
}
 

	
			function OpenWindowAndSetValue(Url,Width,Height,WindowObj,SetObj)
		{
			var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:yes;help:no;scroll:yes;');
			if (ReturnStr!='') SetObj.value=ReturnStr;
			return ReturnStr;
		}	
	
	 
		function actlinks()
				{ if (document.tcjdxr.actlink.checked==true)
				  {
				  ChangesUrl.style.display="";
				  }
				  else
					{
					 ChangesUrl.style.display="none";
					 }
				}
			
function SetValue(type,objname,strvalue)
{ 
	var obj=document.getElementById(objname)
	if (type=="add"){
		obj.value=',,,'+obj.value
		obj.value=obj.value.replace(strvalue,'');
		obj.value=obj.value+strvalue;
		obj.value=obj.value.replace(',,,','');
		obj.value=obj.value.replace(',,','');
		}
	else if (type=="+"){obj.value=parseInt(obj.value)+parseInt(strvalue);}
	else{obj.value=strvalue;}
	obj.focus(); 
 return; 
}		
function insertHTMLToEditor(I,codeStr)
{

KE.insertHtml(codeStr, I);

  
}

		function Templates(Templatef)
			{   if (Templatef==2)
			  {
			  Templatefs.style.display="";
			  }
			  else
				{
				 Templatefs.style.display="none";
				 }
			}
						
 function CheckForm()
	{ var form=document.tcjdxr;
 	 if (form.ClassID.value=='')
		{ alert("请选择栏目!");   
		  return false;
		}
		
	 if (form.Title.value=='')
		{ alert("请输入简短标题!");   
		  form.Title.focus();    
		   return false;
		}
 		    form.Submit1.value="正在提交数据,请稍等...";
		form.Submit1.disabled=true;	
    form.submit();
        return true;
	}	
 	
	</script>
  
</body>
</html>