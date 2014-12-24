<!--#include file="../../ACT.Function.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>添加广告</title>
<link href="../../Images/style.css" rel="stylesheet" type="text/css">
<SCRIPT LANGUAGE="JavaScript" src="images/js.js"></SCRIPT>
<script language="JavaScript" type="text/javascript" src="../../../ACT_inc/js/lhgcore/lhgcore.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../../ACT_inc/js/lhgcore/lhgdialog.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../../ACT_inc/js/lhgcore/Main.js"></script>
</head>
<body>
<%
Dim rs,sql,ShowErr
ConnectionDatabase
if Request.Form("AddAD")<> "" then

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from [ads] where ADID='"&DangerEncode(Request.Form("ADID"))&"'"
rs.Open sql,conn,1,3
If not rs.eof Then
	Response.Write ("<script>alert(' 操作错误! \n\n 广告ID重复,请使用其他ID !');history.back();</script>")
	Response.end
End iF
rs.AddNew
rs("ADID")=DangerEncode(Request.Form("ADID"))
rs("ADType")=TRIM(Request.Form("ADType"))
rs("ADSrc")=DangerEncode(Request.Form("ADSrc"))
rs("ADCode")=DangerEncode(Request.Form("ADCode"))
rs("ADHeight")=TRIM(Request.Form("ADHeight"))
rs("ADWidth")=TRIM(Request.Form("ADWidth"))
rs("ADLink")=DangerEncode(Request.Form("ADLink"))
rs("ADAlt")=DangerEncode(Request.Form("ADAlt"))
rs("ADStopViews")=TRIM(Request.Form("ADStopViews"))
rs("ADStopHits")=TRIM(Request.Form("ADStopHits"))
rs("ADStopDate")=TRIM(Request.Form("ADStopDate"))
rs("ADNote")=TRIM(Request.Form("ADNote"))
rs.Update
rs.Close
Call Actcms.ActErr("操作成功","Sys_Act/gg/Index.asp","")
 Response.end

set rs=nothing
conn.Close
set conn=nothing
end if

Rem 过滤可能出错误的符号

Function DangerEncode(fString)
If not isnull(fString) Then
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10), "")
	fString = replace(fString, "'", """")
    fString = Trim(fString)
    DangerEncode = fString
End If
End Function
%>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td class="bg_tr"><strong>广告管理----广告管理首页</strong></td>
  </tr>
  <tr>
    <td ><strong><a href="?">首页</a> ┆ <a href="index.asp">广告列表 </a>┆<a href="add.asp">添加广告 </a>┆</strong></td>
    </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <form name=form action=add.asp  method=post onSubmit="return chkinput()">
  <tr align="center">
    <td class="bg_tr" height="22" colspan="2" align="left">您现在的位置：广告设置 &gt;&gt; <a href="?"><font class="bg_tr">添加广告</font></a></td>
  </tr>
  <tr>
    <td width="100" align="right" >广告名称：</td>
    <td width="400" ><INPUT name=ADID type="text" class="Ainput"   size="20" maxlength="20"> *不能重名</td>
  </tr>
  <tr>
          <td align="right" >广告类型：</td>
          <td >
              <select size="1" name="ADType"  onChange="ChangeType(this.options[this.selectedIndex].value)">
                <option selected value="1">普通显示</option>
                <option value="2">满屏浮动显示</option>
                <option value="3">上下浮动显示-右</option>
                <option value="4">上下浮动显示-左</option>
                <option value="5">全屏幕渐隐消失</option>
                <option value="6">普通网页对话框 </option>
                <option value="7">可移动透明对话框 </option>
                <option value="8">打开新窗口</option>
                <option value="9">弹出新窗口</option>
		<option value="10">对联式广告</option>
		<option value="11">联盟广告</option>
              </select></td>
   </tr>
    <SCRIPT LANGUAGE="JavaScript">
<!--

function upload(instr,ModeID,iname) 
{J.dialog.get({ id: 'zxsc', title: '在线上传',width: 720,height: '240', page: '../../include/Upload_Admin.asp?A=add&instr='+instr+ "&ModeID="+ModeID+ "&instrname="+iname+ "&" + Math.random() }); 

 }

 function get_obj(obj){
   return document.getElementById(obj);
}
//-->
</SCRIPT>
   
   
   <tr>
        <td align="right"  id="adsrc_text">广告地址：</td>
        <td height="17" >
        <INPUT name="ADSrc" type="text" class="Ainput"  size="40"  value="http://" > <a href="#" onclick=openhelp("ext")>图片或FLASH地址？</a>
 
<input name="button"  onClick="J('#13s').dialog({ title:'文件上传',id:'actcmsscs' ,page: 'include/Upload_Admin.asp?A=add&instr=2&ModeID=999&instrname=ADSrc',  width:720, height:240 });"   id="13s"   type="button"  class="ACT_btn" style="cursor:hand;" value="点击上传图片">


    <font color="#FF0000">[点击上传图片]</font></a>
        
    
		 </td>
   </tr>
        <tr>
          <td align="right"  id="adsrc_text">广告地址：</td>
          <td height="17" >
              <INPUT name=ADCode type="text" class="Ainput"   value="" size=40 maxlength=150> script代码</a></td>
        </tr>
        <tr>
          <td align="right" >广告规格：</td>
          <td height="17" >
              <INPUT name=ADWidth type="text" class="Ainput"   onKeyPress="return Num();" value="468" size=17 maxlength=4>
              ×
              <INPUT name=ADHeight type="text" class="Ainput"   onKeyPress="return Num();" value="60" size=18 maxlength=4></td>
        </tr>
        <tr>
          <td align="right" >链接地址：</td>
          <td height="17" >
              <INPUT name=ADLink type="text" class="Ainput"   size=40 maxlength=100></td>
        </tr>
        <tr>
          <td align="right" >提示文字：</td>
          <td height="17" >
              <INPUT name=ADAlt type="text" class="Ainput"   size=40 maxlength=50></td>
        </tr>
        <tr>
          <td align="right" >投放限制：</td>
          <td height="18" >
            <INPUT name=ADStopViews type="text" class="Ainput"   onKeyPress="return Num();" value="0" size=8 maxlength=10>
            ·
            <INPUT name=ADStopHits type="text" class="Ainput"   onKeyPress="return Num();" value="0" size=8 maxlength=10>
            ·
            <INPUT name=ADStopDate type="text" class="Ainput"   value="<%=Now()+15%>" size=18 maxlength=30> <a href="#" onclick=openhelp("stop")>显示·点击·日期？</a>默认是半个月有效期</td>
        </tr>
        <tr>
          <td align="right" >简单注释：</td>
          <td height="17" >
              <INPUT name=ADNote type="text" class="Ainput"   size=60 maxlength=100>
              备注不显示</td>
        </tr>
        <tr>
          <td height="22" colspan="2" align="center" >
	   <input type=submit class="ACT_btn" name="AddAD" value="  保存  " />
	   &nbsp;&nbsp; <input type="reset" class="ACT_btn" name="Submit2" value="  重置  ">             </td>
        </tr>
      </form>
</table>

<SCRIPT LANGUAGE="JavaScript">
<!--
	  function upfile(url,name){
	document.getElementById(name).value=url;
 }
//-->
</SCRIPT>
</body>
</html>