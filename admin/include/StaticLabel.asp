<!--#include file="../ACT.Function.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ACTCMS自定义标签管理</title>
<link href="../Images/editorstyle.css" rel="stylesheet" type="text/css">
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
</SCRIPT></head>
<body>
<% 		Dim ID,Action,LabelRS,LabelName,LabelContent,SQLStr,Description,ShowErr
		Action = Request.QueryString("Action")
	  Set LabelRS = Server.CreateObject("Adodb.Recordset")
		If Action = "EditLabel" Then
			ID = ChkNumeric(Request.QueryString("ID"))
			Set LabelRS = Server.CreateObject("Adodb.Recordset")
			SQLStr = "SELECT * FROM Label_ACT Where ID=" & ID & ""
			LabelRS.Open SQLStr, Conn, 1, 1
			LabelName = Replace(Replace(LabelRS("LabelName"), "{ACT_", ""), "}", "")
			Description = LabelRS("Description")
			LabelContent = Server.HTMLEncode(LabelRS("LabelContent"))
			LabelRS.Close
		Else
		  If LabelContent="" Then LabelContent="请输入您自定义的html代码"
		End If
		
		
		
		Select Case Request.Form("Action")
		 Case "AddNewSubmit"
			LabelName = Replace(Replace(Trim(Request.Form("LabelName")), """", ""), "'", "")
			Description = Replace(Trim(Request.Form("Description")), "'", "")
			LabelContent = Trim(Request.Form("LabelContent"))
			If LabelName = "" Then
			   Call Actcms.ActErr(ShowErr,"","1")
			  Response.End
			End If
		
			
			If LabelContent = "" Then Call Actcms.ActErr(ShowErr,"","1"): Response.End
			 
 			LabelName = "{ACT_" & LabelName & "}"
			LabelRS.Open "Select LabelName From Label_ACT Where LabelName='" & LabelName & "'", Conn, 1, 1
			If Not LabelRS.EOF Then
			  Call Actcms.ActErr("标签名称已经存在","","1")
			  LabelRS.Close
			  Conn.Close
			  Set LabelRS = Nothing
			  Set Conn = Nothing
 			  Response.End
			Else
				LabelRS.Close
				LabelRS.Open "Select * From Label_ACT", Conn, 1, 3
				LabelRS.AddNew
				 'LabelRS("ID") = ID
				 LabelRS("LabelName") = LabelName
				 LabelRS("LabelContent") = LabelContent
				 LabelRS("Description") = Description
				 LabelRS("AddDate") = Now
				 LabelRS("LabelFlag") = 1
				 LabelRS("LabelType") = 2
				 LabelRS.Update
				 Application.Contents.RemoveAll
 			     Call Actcms.ActErr("添加标签成功","Label_Admin.asp?Type=2","")
			End If
		Case "EditSubmit"
			ID = ChkNumeric(Trim(Request.Form("ID")))
			LabelName = Replace(Replace(Trim(Request.Form("LabelName")), """", ""), "'", "")
			Description = Replace(Trim(Request.Form("Description")), "'", "")
			LabelContent = Trim(Request.Form("LabelContent"))
			If LabelName = "" Then
			   Call Actcms.ActErr(ShowErr,"","1")
			  Response.End
			End If
			If LabelContent = "" Then Call Actcms.ActErr(ShowErr,"","1"):Response.End
			  
 			LabelName = "{ACT_" & LabelName & "}"
			LabelRS.Open "Select LabelName From Label_ACT Where ID <>" & ID & " AND LabelName='" & LabelName & "'", Conn, 1, 1
			If Not LabelRS.EOF Then
			  Call Actcms.ActErr("标签名称已经存在","","1")
			  Response.End
			Else
				LabelRS.Close
				LabelRS.Open "Select * From Label_ACT Where ID=" & ID & "", Conn, 1, 3
				 LabelRS("LabelName") = LabelName
				 LabelRS("LabelContent") = LabelContent
				 LabelRS("Description") = Description
				 LabelRS("AddDate") = Now
				 LabelRS.Update
				 Application.Contents.RemoveAll
 				   Call Actcms.ActErr("标签修改成功","Label_Admin.asp?Type=2","")
			End If
		End Select
		
		
		
 %>
<form id="ahhfchhs" name="ahhfchhs" method="post" action="">
<% 
			If Action = "Add" Or Action = "" Then Response.Write "<input type='hidden' name='Action' value='AddNewSubmit'>"
			If Action = "EditLabel" Then Response.Write "<input type='hidden' name='Action' value='EditSubmit'>"
			
 %>
 <input type="hidden" name="ID" value="<%= ID %>"> 
  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table">
    <tr>
      <td colspan="2" align="center" class="bg_tr"><div align="center">新建自定义静态标签</div></td>
    </tr>
    <tr>
      <td width="15%" align="right" class="tdclass"><strong>标签名称：</strong></td>
      <td width="85%" class="td_bg"><input value="<%= LabelName %>" name="LabelName" style="width:200;" /></td>
    </tr>
    <tr>
      <td align="right" class="tdclass"><strong>标签目录：</strong></td>
      <td class="td_bg"><select name="LabelFlag" id="select">
        <option value="1">静态标签</option>
   
      </select></td>
    </tr>
    <tr>
      <td align="right" class="tdclass"><strong>标签简介：</strong></td>
      <td class="td_bg"><textarea name="Description" rows="3" id="Description" style="width:100%;"><%= Description %></textarea></td>
    </tr>
    <tr>
      <td colspan="2" align="center" class="bg_tr"><strong>自 定 义 静 态 标 签 内 容</strong></td>
    </tr>


  <tr>
      <td align="right" class="tdclass"><strong>批量上传：</strong></td>
      <td class="td_bg">

<div id="sapload">
    
    </div>
 
 <script type="text/javascript">
// <![CDATA[
var so = new SWFObject("<%=ACTCMS.ACTSYS%>act_inc/sapload.swf", "sapload", "450", "25", "9", "#ffffff");
so.addVariable('types','<%=Replace(ACTCMS.ActCMS_Sys(11),"/",";")%>');
so.addVariable('isGet','1');
so.addVariable('args','myid=Upload;ModeID=999;U='+U+";"+';P='+P+";"+'Yname=content1');
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


    <tr>
      <td colspan="2" class="td_bg"> 
	  
	 

  
	<script>
			KE.show({
				id : 'content1'
  			});
		</script>
	 
	   <textarea id="content1" name="LabelContent"  style="width:98%;height:300px;visibility:hidden;">
<%=LabelContent%>
</textarea>  
	  
	  </td>
    </tr>
    <tr>
      <td class="td_bg">&nbsp;</td>
      <td class="td_bg"><input type=button class="ACT_btn" onclick=CheckInfo()  name=Submit value=" 保 存 " />
	  					<input type="reset" class="ACT_btn" name="Submit2" value="  重置  " /></td>
    </tr>
  </table>
</form>
<p>
  <SCRIPT language=javascript>
		function CheckInfo()
		{
		  if(document.ahhfchhs.LabelName.value=="")
			{
			  alert("标签名称不能为空！");
			  document.ahhfchhs.LabelName.focus();
			  return false;
			}
			if(document.ahhfchhs.LabelContent.value=="")
			{
			  alert("标签内容不能为空！");
			  document.ahhfchhs.LabelContent.focus();
			  return false;
			}
		ahhfchhs.Submit.value="正在提交数据,请稍等...";
		ahhfchhs.Submit.disabled=true;	
		ahhfchhs.Submit2.disabled=true;	
	    ahhfchhs.submit();
        return true;
		}
		</script>
</p>
<p>&nbsp; </p>
</body>
</html>