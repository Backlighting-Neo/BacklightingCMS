<!--#include file="../ACT.Function.asp"-->
 <html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>标签显示</title>
<link href="../Images/style.css" rel="stylesheet" type="text/css">
<%
response.Expires = -1
response.ExpiresAbsolute = Now() - 1
response.Expires = 0
response.CacheControl = "no-cache"
Response.CodePage=65001
Response.Charset="utf-8"
 Dim fieldname, num, dbname, dbtype, isknow,isidarr,isid,datasourcetype
 fieldname = Trim(Request("fieldname"))
dbname = Trim(Request("dbname"))
datasourcetype=request("datasourcetype")
isidarr=split(fieldname,".")
isid=false
if ubound(isidarr)=1 then
  if lcase(isidarr(1))="id" and datasourcetype="0" then
    isid=true
  end if
end if

If dbname = "" Then dbname = 0
dbtype = Trim(Request("dbtype"))
If dbtype = "" Then dbtype = 0
isknow = False
%>
 <script language = 'JavaScript'>
function changemode(){
    var dbname=document.myform.ftype.value;
    if(dbname=='Text'){
    input1.style.display='';
    }else{
    input1.style.display='none';
    }
    if(dbname=='Num'){
    input2.style.display='';
    }else{
    input2.style.display='none';
    }
    if(dbname=='Date'){
    input3.style.display='';
    }else{
    input3.style.display='none';
    }
    if(dbname=='GetInfoUrl'|dbname=='GetClassUrl'){
    input5.style.display='';
    }else{
    input5.style.display='none';
    }
}
function changeDate(){
    var dbname=document.myform.Datetype.value;
    if(dbname=='3'){
    document.myform.Datemb.value="2";
    }else{
        document.myform.Datemb.value="YY-MM-DD hh:mm:ss";
    }
}
function submitdate(){
    var dbname=document.myform.ftype.value;
    if(dbname=='Text'){
        dbname="{$Field(" + document.myform.FieldName.value + "," + dbname + "," + document.myform.CatNum.value + "," + document.myform.CatType.value + "," + document.myform.OutSplit.value + ","+document.myform.NullChar.value+")}";
    }
    if(dbname=='Num'){
	    for (var i=0;i<document.myform.OutType.length;i++){
            if (document.myform.OutType[i].checked){
                var OutType=document.myform.OutType[i].value;
        }
        }
       dbname="{$Field(" + document.myform.FieldName.value + "," + dbname + "," + OutType + "," + document.myform.XiaoShu.value + ")}";
    }
    if(dbname=='Date'){
    dbname="{$Field(" + document.myform.FieldName.value + "," + dbname + "," + document.myform.Datemb.value + ")}";
    }
    if(dbname=='GetInfoUrl'||dbname=='GetClassUrl'){
	    for (var i=0;i<document.myform.outype.length;i++){
            if (document.myform.outype[i].checked){
                var outype=document.myform.outype[i].value;
        }
        }
        dbname="{$Field(" + document.myform.FieldName.value + "," + dbname + "," + document.myform.ModeID.value + "," + outype + ")}";
    }
   
    window.returnValue=dbname;
    window.close();
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
		if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>
</head>
<body>
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="table">
<form method='post' action='' name='myform'>
    <tr class="tdbg"><td>字段名称：
    <input name='FieldName' type='text' class='ainput'  id='FieldName' size='20' value="<% =fieldname %>" ></td>
    </tr>
    <tr class="tdbg"><td>输出类型：
      <select name="ftype" style="width:150" onChange="changemode()">
	<option value='Text'>文本型</option>
<%
If (dbtype > 1 And dbtype < 7) Or dbtype = 131 Or dbtype=17 Then
    response.write "<option value='Num' selected>数字型</option>"
    isknow = True
Else
    response.write "<option value='Num'>数字型</option>"
End If
If dbtype = 7 Then
    response.write "<option value='Date' selected>时间型</option>"
    isknow = True
Else
    response.write "<option value='Date'>时间型</option>"
End If

	If   LCase(fieldname) = "id" and datasourcetype="0" Then
        response.write "<option value='GetInfoUrl' selected>对象URL型(系统内置)</option>"
        isknow = True
    Else
       ' response.write "<option value='GetInfoUrl'>对象URL(系统内置)</option>"
    End If

    If Lcase(FieldName)="classid"  Then
        response.write "<option value='GetClassUrl' selected>栏目URL(系统内置)</option>"
        isknow = True
    Else
       ' response.write "<option value='GetClassUrl'>栏目|频道URL(系统内置)</option>"
    End If
%>
</select></td>
    </tr>
<%
If isknow = False Then
    response.write "<tbody id='input1' style='display:'>"
Else
    response.write "<tbody id='input1' style='display:none'>"
End If
%>
    <tr class="tdbg"><td>输出长度：
      <input name='CatNum' type='text' class='ainput'  id='gotopic' size='6' value=0>
    &nbsp;&nbsp;&nbsp;<font color='#FF0000'>为0则不截断</font></td>
    </tr>
	<tr class="tdbg"><td>截断显示：
	  <Input name='CatType' type='text' class='ainput'  value='...' size="6">
	  &nbsp;&nbsp;&nbsp;<font color='#FF0000'>为0则不显示任何标志</font></td>
	</tr>
    <tr class="tdbg"><td>过滤处理：
    <select name='OutSplit'><option value='0' selected>解析HTML标记</option><option value='1'>不解析HTML标记</option><option value='2'>过滤HTML标记</option></select></td>
    </tr>
	    <tr class="tdbg"><td>字段值为空时输出：
    <input title='(如图片值为空，则输出一张默认的图片 "/upfiles/defaule.gif")' name='NullChar' type='text' class='ainput'  id='NullChar' size='20' value=""></td>
    </tr>

</tbody>

<%
If ((dbtype > 1 And dbtype < 7) Or dbtype = 131 Or dbtype=17) And Not (LCase(fieldname) = "id") And Not (LCase(fieldname) = "classid")  and not isid   Then
    response.write "<tbody id='input2' style='display:'>"
Else
    response.write "<tbody id='input2' style='display:none'>"
End If
%>
    <tr class="tdbg"><td>输出方式：<Input type='radio' name='OutType' value='0' checked onClick="input21.style.display='';input22.style.display='none'">
    原数 
        <Input type='radio' name='OutType' value='1' onClick="input21.style.display='none';input22.style.display=''">小数 <Input type='radio' name='OutType' value='2' onClick="input21.style.display='none';input22.style.display='none'">百分数</td></tr>
<%
        If ((dbtype > 1 And dbtype < 7) Or dbtype = 131 Or dbtype=17) And Not (LCase(fieldname) = "id") Then
        response.write "<tbody id='input21' style='display:'>"
        Else
        response.write "<tbody id='input21' style='display:none'>"
        End If
%>
</tbody>
    <tbody id='input22' style='display:none'><tr class="tdbg"><td>小数位数：
      <input name='XiaoShu' type='text' class='ainput'  id='XiaoShu' size='4' value=2></td>
    </tr></tbody>
</tbody>


<%
If dbtype = 7 Or dbtype = 135 Then
    response.write "<tbody id='input3' style='display:'>"
Else
    response.write "<tbody id='input3' style='display:none'>"
End If
%>
    
    <tr class="tdbg">
      <td>输出格式：
        <input name='Datemb' type='text' class='ainput'  id='Datemb' size='28' value="YYYY-MM-DD">
		<br>
		<font color=red>YYYY:年(4位) YY:年(2位) 　MM:月 　DD:日<br>
		hh:时　 mm:分　 ss:秒</font></td>
    </tr>
</tbody>


<%
If dbtype = 11 Then
    response.write "<tbody id='input4' style='display:'>"
Else
    response.write "<tbody id='input4' style='display:none'>"
End If
%>
    
</tbody>


<%
 If LCase(fieldname) = "id" or LCase(fieldname) = "classid"  and datasourcetype="0" Then
    response.write "<tbody id='input5' style='display:'>"
Else
    response.write "<tbody id='input5' style='display:none'>"
End If

%>
<tr class="tdbg"><td>输出方式：
<Input type='radio' name='outype' value=0>
混合 
<Input type='radio' name='outype' value='1' checked>
对象Url 
<Input type='radio' name='outype' value='2'> 
字段值  </td>
 </tr>
 
 <% If LCase(fieldname) = "id" then  %>
<tr class="tdbg"><td>所属模型：<select name="ModeID" id="ModeID">
<% 	    Dim MX_Sys,ii
		MX_Sys=ACTCMS.Act_MX_Sys_Arr()
		If IsArray(MX_Sys) Then
			For iI=0 To Ubound(MX_Sys,2) %>
          <option value="<%=MX_Sys(0,Ii) %>"><%= MX_Sys(1,Ii) %>模型</option>
           <% 
	    Next
		End If %>
        </select></td>
 </tr> 
 <% end if  %>
 
 
</tbody>

<tr class="tdbg" align="center"><td><input type='button' class="ACT_btn" onClick="submitdate();" value=" 插入 " >&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' class="ACT_btn" onClick="window.close();" value=" 取消 " ></td></tr>
<tr class="tdbg" height="100%"><td>&nbsp;<input name='Fieldnum' id='Fieldnum' value="<% =num %>" type='hidden'><br>&nbsp;<br>&nbsp;</td></tr>
</form>
</table>
</body>
</html>
 
