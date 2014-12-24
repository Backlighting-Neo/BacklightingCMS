<!--#include file="../../ACT.Function.asp"-->
<html>
<head>
<title>广告管理</title>
<meta http-equiv="Content-Type" content="text/html; charSet=utf-8">
<link href="../../Images/style.css" rel="stylesheet" type="text/css">
<SCRIPT LANGUAGE="JavaScript" src="images/js.js"></SCRIPT>
</head>
<body>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td class="bg_tr"><strong>广告管理----广告管理首页<a href="#" target="_blank" style="cursor:help;'" class="Help"></a></strong></td>
  </tr>
  <tr>
    <td class="td_bg"><strong><a href="?">首页</a> ┆ <a href="list.asp">广告列表 </a>┆<a href="add.asp">添加广告 </a>┆</strong></td>
  </tr>
  </table>

  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr align="center">
    <td class="bg_tr" height="22" colspan="2" align="left">您现在的位置：广告设置 &gt;&gt; <a href="?"><font class="bg_tr">广告设置</font></a></td>
  </tr>
  <tr>
    <td class="td_bg">
      <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
        <form method=post>
          <tr align="center" class="td_bg">
            <td colspan="2">
              <div class="bar2">广告设置</div></td>
          </tr>
          <tr>
            <td height="250" align="center">如果您的空间不支持<font color="#FF0000">FSO</font>，请直接编辑该文件！</font><br><br>
            <%
            dim id,objFSO,objCountFile,fdata,Rs
            id=request("id")
            Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
            if request("save")="" then
                Set objCountFile = objFSO.OpenTextFile(Server.MapPath(ACTCMS.ACTSYS&"plus/gg/"&id),1,True)
                If Not objCountFile.AtEndOfStream Then fdata = objCountFile.ReadAll
            else
                fdata=request("fdata")
                Set objCountFile=objFSO.CreateTextFile(Server.MapPath(ACTCMS.ACTSYS&"plus/gg/"&id),True)
                objCountFile.Write fdata
                    response.write"<script>alert('保存成功！');window.navigate('editjs.asp?id="&id&"');</script>"
            end if
            objCountFile.Close
            Set objCountFile=Nothing
            Set objFSO = Nothing
            %>
            <textarea name="fdata" name="S1" cols="120" rows="20" class="input2"><%=fdata%></textarea>
            </td>
          </tr>
          <tr>
            <td height="22" colspan="2" align="center" class="td_bg">
	   <input type=submit class="ACT_btn" name="save" value="  保存  " />
	   &nbsp;&nbsp; <input type="reset" class="ACT_btn" name="Submit2" value="  重置  ">             </td>
          </tr>
        </form>
      </table>
    </td>
  </tr>
</table>
<%set rs=nothing%>
</body>
</html>