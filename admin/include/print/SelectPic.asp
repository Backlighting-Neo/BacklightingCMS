<!--#include file="../../ACT.Function.asp"--><%
Dim TypeSql,RsTypeObj,LableSql,RsLableObj
 
Dim LimitUpFileFlag,CurrPath,ShowVirtualPath
LimitUpFileFlag = Request("LimitUpFileFlag")
CurrPath = Request("CurrPath")
ShowVirtualPath = Request("ShowVirtualPath")
Session("TempPicDir")=CurrPath
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<HTML>
<HEAD>
<META http-equiv="Content-Type" content="text/html; charset=utf-8">
<TITLE>选择图片</TITLE>
<link href="../../Images/style.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY leftmargin="0" topmargin="0">
<TABLE width="99%" border="0" align="center" cellpadding="1" cellspacing="0">
  <TR> 
    <TD height="25"><SELECT onChange="ChangeFolder(this.value);" id="FolderSelectList" style="width:100%;" name="select">
			<OPTION selected value="<% = CurrPath %>"><% = CurrPath %></OPTION>
      </SELECT></TD>
    <TD rowspan="2" align="center" valign="middle"><IFRAME id="PreviewArea" width="100%" height="330" frameborder="1" src="Print.asp"></IFRAME></TD>
  </TR>
  <TR> 
    <TD width="70%" align="center"> <IFRAME id="FolderList" width="100%" height="290" frameborder="1" src="Fso.asp?CurrPath=<% = CurrPath %>&ShowVirtualPath=<% = ShowVirtualPath %>"></IFRAME></TD>
  </TR>

  <TR> 
    <TD height="10" colspan="2"> 
      <TABLE width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
        <TR> 
          <TD width="80" height="10"> </TD>
          <TD>Url地址:<INPUT style="width:40%" type="text" name="UserUrl" id="UserUrl"> 
            <INPUT type="button" onClick="SetUserUrl();"  class="ACT_btn" name="Submit" value=" 确 定 ">
            <INPUT onClick="window.close();"  class="ACT_btn" type="button" name="Submit" value=" 取 消 "> 
          </TD>
        </TR>
        <TR> 
          <TD height="10" colspan="2" align="center"><span class="tx">在空白处点鼠标右键可以进行文件类操作,双击文件选择</span></TD>
        </TR>
      </TABLE></TD>
  </TR>
</TABLE>
</BODY>
</HTML>
<SCRIPT language="JavaScript">
function ChangeFolder(FolderName)
{
	frames["FolderList"].location='fso.asp?CurrPath='+FolderName;
}

function SetUserUrl()
{
	if (document.all.UserUrl.value=='') alert('请填写Url地址');
	else
	{
		var templatedir=document.all.UserUrl.value;
 		templatedir=templatedir.replace('<%=actcms.ActSys%>','');
		templatedir=templatedir.replace('<%=actcms.SysThemePath&"/"&actcms.NowTheme%>/','');
 		window.returnValue=templatedir;
  	 	window.close();
	}
}
window.onunload=CheckReturnValue;
function CheckReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}

</SCRIPT>