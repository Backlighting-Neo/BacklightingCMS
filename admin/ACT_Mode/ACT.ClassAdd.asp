<!--#include file="../ACT.Function.asp"-->
<!--#include file="../include/ACT.F.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>栏目管理</title>
<meta http-equiv="X-UA-Compatible" content="IE=8" />
<link href="../Images/editorstyle.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/lhgcore.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/lhgdialog.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/Main.js"></script>
<script charset="utf-8"  language="JavaScript" type="text/javascript" src="../../editor/kindeditor/kindeditor.js" ></script>
<script type="text/javascript" src="../../ACT_INC/js/swfobject.js"></script>
 <SCRIPT LANGUAGE='JavaScript'>
 var U="<%=ACTCMS.strToAsc(RSQL(Trim(Request.Cookies(AcTCMSN)("AdminName"))))%>";
var P="<%=ACTCMS.strToAsc(RSQL(Trim(Request.Cookies(AcTCMSN)("AdminPassword"))))%>";
 <!--
//屏蔽js错误

 function ResumeError() {
 return true;
 }
 window.onerror = ResumeError;
 // -->
</SCRIPT>
</head>
<body>
<%
Dim ClassName,EnName,ActLink,dh,tg,ParentID,CMS_Extension,disabled,content,makehtmlname,pageTemplate
Dim ClassID,ClassKeywords,ClassDescription,GetParentID,ConTentTemplate,FolderTemplate,FilePathName,addParentID
Dim GroupIDClass,TGGroupID,OrderID,TiesDomain,ModeName,ModeID,classnames,ActCMS_DIY,moresite,sitepath,siteurl
Dim ClassPurview,ClassArrGroupID,ClassReadPoint,ClassChargeType ,ClassPitchTime,ClassReadTimes,ClassDividePercent
Dim SEOtitle,ClassPicUrl,labelfor
 ClassID = request("ClassID")
 ModeID = ChkNumeric(Request("ModeID"))
 if ModeID=0 or ModeID="" Then ModeID=1
 ModeName= ACTCMS.ACT_C(ModeID,1) 
 If Not ACTCMS.ACTCMS_QXYZ(ModeID,"","") Then   Call Actcms.Alert("对不起，您没有"&ACTCMS.ACT_C(ModeID,1)&"系该项操作权限！","")

With Response
IF Request("Action") = "add" Then
	ActCMS_DIY=Split(AF.ActCMS_DIY_F(ModeID,1,""),"§") 
	ClassPurview=0
 	ClassReadPoint=0
	ClassChargeType =0
	ClassPitchTime=0
	ClassReadTimes=0
	ClassDividePercent=0
	labelfor=1
	dh=1:tg=1:CMS_Extension="Index.Html":OrderID=10:moresite=0:	ActLink=1:FilePathName="{enname}/{id}"
	if ClassID <>"" then
		GetParentID =  ClassID
		FolderTemplate=ActCMS_DIY(23)
	Else
		GetParentID = "0"
		FolderTemplate=ActCMS_DIY(22)
	End if
	IF ClassID <> "" Then 
		Dim Rs,ShowErr
		Set Rs=server.CreateObject("adodb.recordset") 
		Rs.OPen "Select ActLink from Class_Act where ClassID='"& ClassID &"'  order by id desc",Conn,1,1
		If Not  Rs.Eof Then '判断是否为转向链接
			 if Rs("ActLink")="2" then 
 				 Call Actcms.ActErr("转向链接不能添加子类","","1")
			 Else 
				classnames =ACTCMS.ACT_L(ClassID,2)
				EnName=ACTCMS.ACT_L(ClassID,3) 
				addParentID=ACTCMS.ACT_L(ClassID,11) 
			 End if 
		End If
		
	Else
		classnames ="根栏目"
	End If
	ConTentTemplate=ActCMS_DIY(24)
Else
	Set Rs = server.CreateObject("adodb.recordset")
	Rs.open "select * From  Class_Act where ClassID = '"& ClassID &"'",Conn,1,3
	if Not   Rs.eof then
		dh = Rs("dh"):tg = Rs("tg")
		ClassName = Rs("ClassName")
		EnName = Rs("ClasseName") 
		CMS_Extension = Rs("Extension")
		ClassDescription = Rs("ClassDescription")
		ClassKeywords = Rs("ClassKeywords")
		ActLink = Rs("ActLink")
		TGGroupID = Rs("TGGroupID")
		ConTentTemplate=Rs("ConTentTemplate")
		OrderID=Rs("OrderID")
		GroupIDClass=Rs("GroupIDClass")
		ModeID=Rs("ModeID")
		moresite=Rs("moresite")
		sitepath=Rs("sitepath")
		siteurl=Rs("siteurl")
		FilePathName=Rs("FilePathName")
		content=Rs("content")
	    makehtmlname=Rs("makehtmlname")
		FolderTemplate=Rs("FolderTemplate")
		pageTemplate=Rs("FolderTemplate")
 		ClassPurview=Rs("ClassPurview")
		ClassReadPoint=Rs("ClassReadPoint")
		ClassChargeType =Rs("ClassChargeType")
		ClassPitchTime=Rs("ClassPitchTime")
		ClassReadTimes=Rs("ClassReadTimes")
		ClassDividePercent=Rs("ClassDividePercent")
		ClassArrGroupID=Rs("ClassArrGroupID")
		SEOtitle=rs("SEOtitle")
		ClassPicUrl=rs("ClassPicUrl")
		ACTlink=rs("ACTlink")
		labelfor=rs("labelfor")
   		Dim Rs1
		Set Rs1=actcms.actexe("Select ClassName From Class_Act Where ClassID='" &Rs("ParentID")&"' ")
		If Not Rs1.EOF Then classnames=Rs1("ClassName") Else classnames ="根栏目"
		disabled="disabled=true "
	End If
End If
     %>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td class="bg_tr">您现在的位置：<%=ModeName%>管理 &gt;&gt; 栏目管理</td>
  </tr>
  <tr>
		<td height="18" class="tdclass"><a href="ACT.Class.asp?ModeID=<%=ModeID %>">管理首页</a>┆<a href="ACT.ClassAdd.asp?ModeID=<%=ModeID %>&Action=add">添加根栏目</a>┆<a href="ACT.ClassAct.asp?Action=one&ModeID=<%=ModeID %>">栏目排序</a></td>
  </tr>
</table>

  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
    <form id="Article" name="Article" method="post" action="ACT.ClassSave.asp?action=<%= request("Action")%>"><tr>
      <td colspan="2" align="center" class="bg_tr">添加/修改<%=ModeName%>栏目</td>
    </tr>
	<tr>
      <td width="18%" align="right"  class="tdclass"><strong>所属分类：</strong></td>
      <td width="82%"  class="tdclass"><%=classnames %></td>
    </tr>	
	
	<tr>
      <td width="18%" align="right"  class="tdclass"><strong>栏目名称：</strong></td>
      <td  class="tdclass"><input name="ClassName" id="ClassName" class="Ainput"    value="<%= ACTCMS.HTMLCode(ClassName) %>" size="35" /> 
      <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ACTClassAdd_001')"  id="ACTClassAdd_001">帮助</span> 
	  支持HTML语法
		<input name="ClassID" type="hidden" id="ClassID" value="<%= ClassID %>" readonly>
		<input name="ParentID" type="hidden" id="ParentID" value="<%= GetParentID %>" readonly> 
	  </td>
    </tr>

<% If Request("Action")="add" And ClassID<>"" then %>
    <tr>
      <td align="right"  class="tdclass"><strong>上级栏目目录：</strong></td>
      <td  class="tdclass">
        <input name="onEnName" class="Ainput"   id="onEnName" value="<%= EnName %>" size="35" />
       <font color=red>注意:前面不能带 / 符号</font></td>
    </tr>


 <% end if %>
    <tr>
      <td align="right"  class="tdclass"><strong>栏目目录：</strong></td>
      <td  class="tdclass">
        <input name="EnName" class="Ainput"   id="EnName" value="<%=EnName%>"  disabled size="35" />
       <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ACTClassAdd_002')"  id="ACTClassAdd_002">帮助</span>
		<% If Request("Action")="add" then %><label for="IFPinYin">
        <input  id="IFPinYin" value="1" type="checkbox"  name="IFPinYin" checked="checked" onClick="IFPinYins()" />
          栏目拼音</label>
     <font color="red">英文名称</strong>必须是字母,数字,确认后不能修改</font>
	<%Else %>
	 

	 <label for="EditEname">
        <input  id="EditEname" value="1" type="checkbox"  name="EditEname" onClick="EditEnames()" />
          修改</label>
	 
	 <% end if %></td>
    </tr>
   
   
   
      <tr>
      <td height="25" align="right"  class="tdclass"><strong>SEO标题：</strong></td>
      <td height="25"  class="tdclass"><input name="SEOtitle" type="text" class="Ainput"  value="<%= SEOtitle %>" size="35">
	   </td>
    </tr>
 
   
      <tr>
      <td height="25" align="right"  class="tdclass"><strong>栏目缩略图：</strong></td>
      <td height="25"  class="tdclass"><input name="ClassPicUrl" type="text" class="Ainput"  value="<%= ClassPicUrl %>" size="35">
    	  <input name="button"  onClick="J('#scs').dialog({ id:'actcmsscs' ,page: 'include/Upload_Admin.asp?A=0&instr=999&ModeID=1&instrname=ClassPicUrl',  width:720, height:240 });"   id="scs"   type="button"  class="ACT_btn" style="cursor:hand;" value="点击上传图片">
 <font color="#FF0000">[点击上传图片]</font> <SCRIPT LANGUAGE="JavaScript">
<!--

function upload(instr,ModeID,iname) 
{J.dialog.get({ id: 'zxsc', title: '在线上传',width: 720,height: '240', page: '<%=actcms.actsys&actcms.adminurl%>/include/Upload_Admin.asp?A=add&instr='+instr+ "&ModeID="+ModeID+ "&instrname="+iname+ "&" + Math.random() }); 

 }

 function get_obj(obj){
   return document.getElementById(obj);
}
//-->
</SCRIPT>
	   </td>
    </tr>
 
<%
If  (Request("Action") = "add" And classid="") Or  (Request("Action") = "edit" And classid<>"" And addParentID="0" )  And actlink=1 then 

 %>	
	
   <tr>
      <td align="right" class="tdclass"><strong>多站点支持：</strong></td>
      <td  class="tdclass">
      <label for="moresites1"><input onClick=moresiteset(0) type="radio" <%If moresite=0  then response.Write("Checked ")  %>  value="0" name="moresite"  id="moresites1">&nbsp;不启用&nbsp;</label>
      <label for="moresites2"><input onClick=moresiteset(1) type="radio" <%If moresite=1   then response.Write("Checked ")  %>  value="1" name="moresite" id="moresites2">&nbsp;启用&nbsp;</label> <font color=green>说明： 绑名绑定仅需要在顶级栏目设定，下级栏目更改无效。</font>  
		</td>
    </tr>	
 
	<tr id="moresite1"
	<%If moresite<>1 Then response.write "style=""DISPLAY: none"""%>
	>
      <td align="right" class="tdclass"><strong>捆绑域名：</strong></td>
      <td class="tdclass"><input name="siteurl" type="text" class="Ainput"  value="<%= siteurl %>" size="35" />
	  <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ACTClassAdd_003')"  id="ACTClassAdd_003">帮助</span>(需加 http://，一级或二级域名的根网址)&nbsp;<a href="http://www.actcms.com/help/?id=ACTClassAdd_003" target=_blank><font color=green>查看视频教程</font></a></td>
    </tr>

	<tr id="moresite2" 	<%If moresite<>1 Then response.write "style=""DISPLAY: none"""%>
>
      <td align="right" class="tdclass"><strong>站点根目录：</strong></td>
      <td class="tdclass"><input name="sitepath" type="text" class="Ainput"  value="<%= sitepath %>" size="35" />
	   </td>
    </tr>


<% end if %>
 

    <tr>
      <td align="right" class="tdclass"><strong>栏目模板地址：</strong></td>
      <td class="tdclass">
	  <input  name="FolderTemplate" type="text" class="Ainput"  size="35"  value="<%= FolderTemplate %>" />
	 <input class="ACT_btn" type="button"  onClick="OpenWindowAndSetValue('../include/print/SelectPic.asp?CurrPath=<%=ACTCMS.ActSys%><%=actcms.SysThemePath&"/"&actcms.NowTheme%>',500,320,window,document.Article.FolderTemplate);" value="选择模板..."> 
	 <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ACTClassAdd_004')"  id="ACTClassAdd_004">帮助</span></td>
    </tr>
    <tr>
      <td align="right" class="tdclass"><strong>内容页模板地址：</strong></td>
      <td class="tdclass"><input name="ConTentTemplate" type="text" class="Ainput"  size="35" value="<%= ConTentTemplate %>" />
	 <input class="ACT_btn" type="button"  onClick="OpenWindowAndSetValue('../include/print/SelectPic.asp?CurrPath=<%=ACTCMS.ActSys%><%=actcms.SysThemePath&"/"&actcms.NowTheme%>',500,320,window,document.Article.ConTentTemplate);" value="选择模板..."> 
		<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ACTClassAdd_005')"  id="ACTClassAdd_005">帮助</span></td>
    </tr>

    <tr>
      <td height="25" align="right" class="tdclass" ><strong>内容生成规则：</strong></td>
      <td height="25"  class="tdclass"><input name="FilePathName" type="text" class="Ainput"  id="FilePathName" value="<%= FilePathName %>" size="50">
	  <span class="h" style="cursor:help;"  onclick="dohelp('ACTmx_nrsc')"  id="ACTmx_nrsc">帮助</span>
	  全部小写</td>
    </tr>

    <tr>
      <td align="right" class="tdclass"><strong>排列权重：</strong></td>
      <td class="tdclass"><input name="OrderID" type="text" class="Ainput"  value="<%= OrderID %>" />
	  <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ACTClassAdd_006')"  id="ACTClassAdd_006">帮助</span></td>
    </tr>

    <tr>
      <td align="right" class="tdclass"><strong>内容模型：</strong></td>
      <td class="tdclass">
	  <select name="ModeID" id="ModeID">
          <%=AF.ACT_L_Mode(CInt(ModeID))%>
        </select>
	  <input type="button" name="Submit6" class="ACT_btn" value="管理系统模型" onClick="window.open('../Mode/ACT.MX.asp');">
	  
	  </td>
    </tr>
	
	
    <tr>
      <td align="right" class="tdclass"><strong>在导航栏显示：</strong></td>
      <td class="tdclass">
        <label for=dh1>
        <input  type=radio  value=1 <%If dh = 1 then response.Write("Checked ")  %> name=dh id=dh1>
        &nbsp;&nbsp;是&nbsp;&nbsp;</label>
        <label for=dh2>
        <input type=radio  <%If dh = 0 then response.Write("Checked ")  %> value=0 name=dh id=dh2>
        &nbsp;&nbsp;否&nbsp;&nbsp;</label>   
		 <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ACTClassAdd_007')"  id="ACTClassAdd_007">帮助</span></td>
    </tr>


    <tr>
      <td align="right" class="tdclass"><strong><font color=red>是否在循环栏目标签中显示：</font></strong></td>
      <td class="tdclass">
        <label for=labelfor1>
        <input  type=radio  value=1 <%If labelfor = 1 then response.Write("Checked ")  %> name=labelfor id=labelfor1>
        &nbsp;&nbsp;是&nbsp;&nbsp;</label>
        <label for=labelfor2>
        <input type=radio  <%If labelfor = 0 then response.Write("Checked ")  %> value=0 name=labelfor id=labelfor2>
        &nbsp;&nbsp;否&nbsp;&nbsp;</label>   
		 <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ACTClassAdd_labelfor')"  id="ACTClassAdd_labelfor">帮助</span></td>
    </tr>

    <tr>
      <td align="right" class="tdclass"><strong>栏目类型：</strong></td>
      <td class="tdclass">
        <label for=ChangesLink1>
        <input onClick=ClassSetting(1) type=radio   <%If ActLink = "1" then response.Write("Checked ")  %>  value=1 name=ActLink id=ChangesLink1>
        &nbsp;&nbsp;系统栏目&nbsp;&nbsp;</label>
        <label for=ChangesLink2>
        <input onClick=ClassSetting(2) type=radio  <%If ActLink = "2"  then response.Write("Checked ")  %>  value=2 name=ActLink id=ChangesLink2>
        &nbsp;&nbsp;外部链接&nbsp;&nbsp;</label> 
        
         <label for=ChangesLink3>
        <input onClick=ClassSetting(3) type=radio   <%If ActLink = "3"  then response.Write("Checked ")  %>  value=3 name=ActLink id=ChangesLink3>
        &nbsp;&nbsp;单页面&nbsp;&nbsp;</label> 
         
		<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ACTClassAdd_008')"  id="ACTClassAdd_008">帮助</span></td>
    </tr>
    <tr  id=ClassSetting3 >
      <td align="right" class="tdclass"><STRONG>生成文件名：</STRONG></td>
      <td class="tdclass">
        <input name=makehtmlname  class="Ainput" value="<%= makehtmlname %>" size=45>如 about.html,intro.html,help.html等
      </td>
       
    </tr>
    <tr  id=ClassSetting9 >
      <td align="right" class="tdclass"><STRONG>栏目模板：</STRONG></td>
      <td class="tdclass">
        <input name=pageTemplate  class="Ainput" value="<%= pageTemplate %>" size=45><input class="ACT_btn" type="button"  onClick="OpenWindowAndSetValue('../include/print/SelectPic.asp?CurrPath=<%=ACTCMS.ActSys%><%=actcms.SysThemePath&"/"&actcms.NowTheme%>',500,320,window,document.Article.pageTemplate);" value="选择模板..."> 
	  	   

      </td>
     </tr>
       <tr  id=ClassSetting8>
     
<td  height="23" align="right"   class="tdclass">批量上传文件：</td>
<td    class="tdclass">
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
 
 
   
    <tr  id=ClassSetting10 >
      <td align="right" class="tdclass"><STRONG>单页内容：
使用标签{$GetClassIntro}在模板里调用：</STRONG></td>
      <td class="tdclass">
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
    

      <tr  id=ClassSetting1 >
      <td align="right" class="tdclass"><STRONG>转向链接URL：</STRONG></td>
      <td class="tdclass">
        <input name="LinkUrl" type="text" class="Ainput" value="<%= makehtmlname %>"  size="45" />
      </td>
    </tr>
    <tr id=ClassSetting2>
      <td align="right" class="tdclass"><strong>栏目首页文件：</strong></td>
      <td class="tdclass"><input name="Extension" type="text" class="Ainput"  id="Extension" value="<%= CMS_Extension %>" size="10"> 
             <select name="select" onChange="document.Article.Extension.value=this.value">
			 <option value='index.htm'>index.htm</option>
			 <option value='index.html'>index.html</option>
			 <option value='index.asp'>index.asp</option>
			 <option value='index.shtm'>index.shtm</option>
			 <% IF Request("Action") = "edit" Then %>
			 <option value='<%= CMS_Extension %>' selected>default.shtm</option>
			 <% end if  %>
            </select>
           <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ACTClassAdd_009')"  id="ACTClassAdd_009">帮助</span>
		   (如index.html,index.asp等,默认不填写,路径后面不带文件名  如 www.actcms.com/anli)</td>
    </tr>

    <tr id=ClassSetting4  >
      <td align="right" class="tdclass"><strong>栏目是否允许投稿：</strong></td>
      <td class="tdclass">
        <label for=tg1>
        <input type=radio  value=0 name=tg id=tg1 <% if tg = 0 then .Write("Checked") %>>
        &nbsp;&nbsp;否&nbsp;&nbsp;</label>
        <label for=tg2>
        <input type=radio  value=1 name=tg id=tg2 <% if tg = 1 then .Write("Checked") %>>
        &nbsp;&nbsp;是&nbsp;&nbsp;</label>  
	<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ACTClassAdd_010')"  id="ACTClassAdd_010">帮助</span>	</td>
    </tr>
    <tr id=ClassSetting5>
      <td align="right" class="tdclass"><strong>栏目META关键字：</strong></td>
      <td class="tdclass">
        <textarea name="ClassKeywords" style="width:80%" rows="5" id="ClassKeywords"><%= ClassKeywords %></textarea> 
		<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ACTClassAdd_011')"  id="ACTClassAdd_011">帮助</span></td>
    </tr>
    <tr id=ClassSetting6>
      <td align="right" class="tdclass"><strong>栏目META描述：</strong></td>
      <td class="tdclass">
        <textarea name="ClassDescription" style="width:80%" rows="5" id="ClassDescription"><%= ClassDescription %></textarea>
		<span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ACTClassAdd_012')"  id="ACTClassAdd_012">帮助</span></td>
    </tr>
   
     <tr id=ClassSetting7>
      <td align="right" class="tdclass"><STRONG>允许此栏目下投稿的会员组：</STRONG></td>
      <td class="tdclass"><%= actcms.GetGroup_CheckBox("TGGroupID",TGGroupID,5)  %>
	  <span class="h" style="cursor:help;" title="点击显示帮助" onClick="dohelp('ACTClassAdd_014')"  id="ACTClassAdd_014">帮助</span></td>
    </tr>
    
    
    
    

    
  <tr  >
    <td align='right'    class="tdclass">权限设置：
	 </td>
    <td height='30'  class="tdclass">
    <label for="ClassPurview1">
    <input name='ClassPurview' id='ClassPurview1' type='radio' <% if  ClassPurview="0" then response.Write "checked=""checked""" %> value='0'  />
      开放栏目（任何人（包括游客）可以浏览和查看此栏目下的信息）</label><br />
      <label for="ClassPurview2"><input name='ClassPurview' id='ClassPurview2' <% if  ClassPurview="1" then response.Write "checked=""checked""" %>  type='radio' value='1' />
      半开放栏目 （任何人（包括游客）都可以浏览。游客不可查看，其他会员根据会员组的栏目权限设置决定是否可以查看,<font color="green">在下面设置相应的会员组权限</font>）</label><br />
     <label for="ClassPurview3"> <input name='ClassPurview' id='ClassPurview3' <% if  ClassPurview="2" then response.Write "checked=""checked""" %>  type='radio' value='2' />
      认证栏目  游客不能浏览和查看，其他会员根据会员组的栏目权限设置决定是否可以浏览和查看,<font color="green">在下面设置相应的会员组权限</font>）</label><br />
      <table border='0' width='90%'>
        <tr>
          <td><%= actcms.GetGroup_CheckBox("ClassArrGroupID",ClassArrGroupID,5)  %></td>
        </tr>
      </table></td>
  </tr>
  <tr  >
    <td align='right'   height="30" class="tdclass"><strong>阅读点数： </strong></td>
    <td height='30'  class="tdclass">&nbsp;
        <input  name='ClassReadPoint' type='text' id='ClassReadPoint'  value='<%=ClassReadPoint  %>' size='6' class='Ainput' />
      免费阅读请设为 &quot;<font color="red">0</font>&quot;，否则有权限的会员阅读此文章时将消耗相应点数，游客将无法阅读此文章 </td>
  </tr>
  <tr >
    <td align='right'    height="30" class="tdclass"><strong>重复收费：</strong><br>
只有当上述计费才有效</td>
    <td height='30' class="tdclass" >
     <label for="ClassChargeType1">
     <input name='ClassChargeType'  id='ClassChargeType1' type='radio' value='0' <% if  ClassChargeType="0" then response.Write "checked=""checked""" %>  />
      不重复收费(如果需扣点数文章，建议使用)</label><br />
       <label for="ClassChargeType2"><input name='ClassChargeType' id='ClassChargeType2' <% if  ClassChargeType="1" then response.Write "checked=""checked""" %> type='radio' value='1' />
      距离上次收费时间
      <input name='ClassPitchTime' type='text' class='Ainput' value='<%= ClassPitchTime %>' size='8' maxlength='8'  />
      小时后重新收费</label><br />
     <label for="ClassChargeType3"> <input name='ClassChargeType' <% if  ClassChargeType="2" then response.Write "checked=""checked""" %> id='ClassChargeType3'  type='radio' value='2' />
      会员重复阅读此文章
      <input name='ClassReadTimes' type='text' class='Ainput' value='<%= ClassReadTimes %>' size='8' maxlength='8'  />
      页次后重新收费</label><br />
     <label for="ClassChargeType4"> <input name='ClassChargeType' <% if  ClassChargeType="3" then response.Write "checked=""checked""" %> id='ClassChargeType4'  type='radio' value='3' />
      上述两者都满足时重新收费</label><br />
      <label for="ClassChargeType5"><input name='ClassChargeType' <% if  ClassChargeType="4" then response.Write "checked=""checked""" %> id='ClassChargeType5'  type='radio' value='4' />
      上述两者任一个满足时就重新收费</label><br />
      <label for="ClassChargeType6"><input name='ClassChargeType' type='radio' <% if  ClassChargeType="5" then response.Write "checked=""checked""" %>  id='ClassChargeType6'  value='5' />
      每阅读一页次就重复收费一次（建议不要使用,多页文章将扣多次点数）</label> </td>
  </tr>
  <tr >
    <td align='right'   height="30" class="tdclass"><strong>分成比例： </strong></td>
    <td height='30'  class="tdclass">&nbsp;
        <input name='ClassDividePercent' type='text' id='ClassDividePercent'  value='<%= ClassDividePercent %>' size='6' class='Ainput' />
      % 　如果比例大于0，则将按比例把向阅读者收取的点数支付给投稿者 </td>
  </tr>
     
   
   
   
   
    
    <tr>
      <td colspan="2" align="center" class="tdclass">
        <input type=button class="ACT_btn" onclick=CheckForm() name=Submit value="  保存  " />
&nbsp;&nbsp;
<input class="ACT_btn" type="reset" name="Submit2" value="  重置  ">
     </td>
    </tr></form>
  </table>
  <% End With %>
<script language = "JavaScript">
function ClassSetting(n){
	if (n == "3"){
		ClassSetting1.style.display='none';
		ClassSetting2.style.display='none';
		ClassSetting3.style.display='';
		ClassSetting9.style.display='';
		ClassSetting10.style.display='';
		ClassSetting4.style.display='none';
		ClassSetting5.style.display='none';
		ClassSetting6.style.display='none';
		ClassSetting7.style.display='none';
		ClassSetting8.style.display='';
	}
	else if (n == "2"){
		ClassSetting1.style.display='';
		ClassSetting2.style.display='none';
		ClassSetting3.style.display='none';
		ClassSetting9.style.display='none';
		ClassSetting4.style.display='none';
		ClassSetting5.style.display='none';
		ClassSetting6.style.display='none';
		ClassSetting7.style.display='none';
		ClassSetting8.style.display='none';
		ClassSetting10.style.display='none';
	}
	else
	{
		ClassSetting1.style.display='none';
		ClassSetting2.style.display='';
		ClassSetting3.style.display='none';
		ClassSetting4.style.display='';
		ClassSetting5.style.display='';
		ClassSetting6.style.display='';
		ClassSetting7.style.display='';
		ClassSetting8.style.display='none';
		ClassSetting9.style.display='none';
		ClassSetting10.style.display='none';
	
	}
}

function insertHTMLToEditor(codeStr,I)
{
  	var oEditor = CKEDITOR.instances[I];
 	if ( oEditor.mode == 'wysiwyg' )
	{
 		oEditor.insertHtml( codeStr );
	}
	else
		alert( '未定义' );
}

function moresiteset(n){
	if (n == 0){
		moresite1.style.display='none';
		moresite2.style.display='none';
	}
	else{
		moresite1.style.display='';
		moresite2.style.display='';
	}
} 
function IFPinYins()
			{ if (document.Article.IFPinYin.checked==true)
			  {
			  document.Article.EnName.disabled=true;
			  document.Article.EnName.value="不能修改";
			  }
			  else
				{
				document.Article.EnName.checked=true;
				  document.Article.EnName.disabled=false;
				   document.Article.EnName.value="";
				 }
			}
function EditEnames()
			{ if (document.Article.EditEname.checked==false)
			  {
			  document.Article.EnName.disabled=true;
			  }
			  else
				{
				document.Article.EnName.checked=true;
				  document.Article.EnName.disabled=false;
				 }
			}


function CheckForm()
{ var form=document.Article;
   if (form.ClassName.value=='')
    { alert("请输入栏目的分类名称!");   
	  form.ClassName.focus();    
	   return false;
    }
		form.Submit.value="正在提交数据,请稍等...";
		form.Submit.disabled=true;	
		form.Submit2.disabled=true;	
	    form.submit();
        return true;
}
function OpenWindowAndSetValue(Url,Width,Height,WindowObj,SetObj)
{
	var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:yes;help:no;scroll:yes;');
	if (ReturnStr!='') SetObj.value=ReturnStr;
	return ReturnStr;
}	</script>
<script language="javascript">ClassSetting(<%=""""&ActLink&""""%>);</script>
	<% If Request("Action")="add" then %>
		<script language="javascript">IFPinYins();</script>
		<% End if%>
</body>
</html>
