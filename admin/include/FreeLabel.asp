<!--#include file="../ACT.Function.asp"-->
<!--#include file="ACT.F.asp"-->
 <html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>1</title>
<link href="../Images/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgajax.js"></script>
 <script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/lhgcore.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/lhgdialog.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/Main.js"></script>

</head>
<body>
<% 	
	Dim ID,Action,LabelRS,LabelName,LabelContent,SQLStr,Description,ShowErr,LabelC,SQL,LabelContentArr
	dim labeltype, PageStyle,ProjectUnit,LabelFlag,datasourcetype,datasourcestr

		Action = Request.QueryString("Action")
	  Set LabelRS = Server.CreateObject("Adodb.Recordset")
		If Action = "EditLabel" Then
			ID = ChkNumeric(Request.QueryString("ID"))
			Set LabelRS = Server.CreateObject("Adodb.Recordset")
			SQLStr = "SELECT * FROM Label_ACT Where ID=" & ID & ""
			LabelRS.Open SQLStr, Conn, 1, 1
			LabelName = Replace(Replace(LabelRS("LabelName"), "{ACTSQL_", ""), "}", "")
			Description = LabelRS("Description")
			LabelContent = Replace(Replace(LabelRS("LabelContent"), "{$ACTSQL(", ""), ")}", "")
 			LabelContent = Replace(LabelContent, """", "") 
			LabelContentArr = Split(LabelContent, "§")
			labeltype=LabelContentArr(0)
			ProjectUnit=LabelContentArr(1)
			PageStyle=LabelContentArr(2)
 		 	datasourcetype=LabelContentArr(3)
			if datasourcetype="0" then 
				datasourcestr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=数据库.mdb"
			else 
		 		datasourcestr=LabelContentArr(4)
			end if
			SQL=LabelContentArr(5)
			LabelContent=LabelRS("Description")
			LabelRS.Close
 		Else
		  If LabelContent="" Then LabelContent="[loop=10]请在此输入循环内容[/loop]"
			ProjectUnit="篇"
			labeltype=0
			datasourcestr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=数据库.mdb"
			sql="Select TOP 10 ID,ClassID,Title,UpdateTime,ActLink,FileName,Content,Intro From  article_act  Where  isAccept=0 AND delif=0 "
		End If
		
	 
		
		Select Case Request.Form("Action")
		 Case "AddNewSubmit"
			LabelName = Replace(Replace(Trim(Request.Form("LabelName")), """", ""), "'", "")
			labeltype=Request.Form("labeltype")
			SQL = Trim(Request.Form("Sql"))
			ProjectUnit=Request.Form("ProjectUnit")
			datasourcetype=Request.Form("datasourcetype")
			datasourcestr=Request.Form("datasourcestr")
			PageStyle=ChkNumeric(Request("PageStyle"))
			LabelC=labeltype&"§"&ProjectUnit&"§"&PageStyle&"§"&datasourcetype&"§"&datasourcestr&"§"&SQL
			LabelContent=Request.Form("LabelContent")
			If LabelName = "" Then
 			   Call Actcms.ActErr(ShowErr,"","1")
			  Response.End
			End If
			If LabelContent = "" Then Call Actcms.ActErr(ShowErr,"","1"): Response.End
			LabelName = "{ACTSQL_" & LabelName & "}"
			LabelRS.Open "Select LabelName From Label_ACT Where LabelName='" & LabelName & "'", Conn, 1, 1
			If Not LabelRS.EOF Then
			  Call Actcms.ActErr("标签名称已经存在","","")
 			  LabelRS.Close
			  Conn.Close
			  Set LabelRS = Nothing
			  Set Conn = Nothing
			  Set ClsMain = Nothing
			  Response.End
			Else
				LabelRS.Close
				LabelRS.Open "Select * From Label_ACT", Conn, 1, 3
				LabelRS.AddNew
				 LabelRS("LabelName") = LabelName
				 LabelRS("LabelContent") = LabelC
				 LabelRS("Description") = LabelContent
				 LabelRS("AddDate") = Now
				 LabelRS("LabelFlag") = 1
				 LabelRS("LabelType") = 3
				 LabelRS.Update
				 Application.Contents.RemoveAll
 				 Call Actcms.ActErr("添加标签成功","Label_Admin.asp?Type=3","")
 			End If
		Case "EditSubmit"
			ID = ChkNumeric(Trim(Request.Form("ID")))
			LabelName = Replace(Replace(Trim(Request.Form("LabelName")), """", ""), "'", "")
			labeltype=Request.Form("labeltype")
			SQL = Trim(Request.Form("Sql"))
			ProjectUnit=Request.Form("ProjectUnit")
			datasourcetype=Request.Form("datasourcetype")
			datasourcestr=Request.Form("datasourcestr")
 			PageStyle=ChkNumeric(Request("PageStyle"))
			LabelC=labeltype&"§"&ProjectUnit&"§"&PageStyle&"§"&datasourcetype&"§"&datasourcestr&"§"&SQL
			LabelContent=Request.Form("LabelContent")
			If LabelName = "" Then
			   Call Actcms.ActErr(ShowErr,"","1")
			  Response.End
			End If
			If LabelContent = "" Then Call Actcms.ActErr(ShowErr,"","1"): Response.End
			LabelName = "{ACTSQL_" & LabelName & "}"
			LabelRS.Open "Select LabelName From Label_ACT Where ID <>" & ID & " AND LabelName='" & LabelName & "'", Conn, 1, 1
			If Not LabelRS.EOF Then
  			  Call Actcms.ActErr("标签名称已经存在","","1")
			  Response.End
			Else
				LabelRS.Close
				LabelRS.Open "Select * From Label_ACT Where ID=" & ID & "", Conn, 1, 3
				 LabelRS("LabelName") = LabelName
				 LabelRS("LabelContent") = LabelC
				 LabelRS("Description") = LabelContent
				 LabelRS("AddDate") = Now
				 LabelRS.Update
				 Application.Contents.RemoveAll
				  Call Actcms.ActErr("标签修改成功","Label_Admin.asp?Type=3","")
 			End If
		End Select
		
		

 %>
 
<form name="ahhfchhs" id="ahhfchhs" method="post" action="">
<% 
			If Action = "Add" Or Action = "" Then Response.Write "<input type='hidden' name='Action' value='AddNewSubmit'>"
			If Action = "EditLabel" Then Response.Write "<input type='hidden' name='Action' value='EditSubmit'>"
			
 %>
 <input type="hidden" name="ID" value="<%= ID %>"> 
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
    <tr>
      <td colspan="2" align="center" class="bg_tr">创建自由标签</td>
    </tr>
    <tr>
      <td width="16%" align="right"><strong>标签名称：</strong></td>
      <td width="84%">
        <input value="<%= LabelName %>" class="Ainput" name="LabelName" style="width:200;" /> 
        例如标签名称："推荐文章列表"，则在模板中调用："{ACTSQL_推荐文章列表(参数1,参数2...)}"。     </td>
    </tr>
  
    <tr >
      <td height="30"  align='right'><strong>标签目录：</strong></td>
	   <td>  <select name="LabelFlag" id="select">
		  <option value="0">系统默认</option>
			 <%=AF.ACT_LabelFolder(CInt(LabelFlag))%>
        </select>&nbsp;&nbsp;<a href="ACT.LabelFolder.asp"><font color=red><b>新建存放目录</b></font></a>
		&nbsp;<font color=green>标签存放目录,方便管理标签</font></td>
    </tr>
    <tr >
      <td height="30"  align='right'><strong>数 据 源：</strong></td>
	   <td>
	     <select name="datasourcetype" style="width:200px"  onChange="changeconnstr(this.options[this.selectedIndex].value);"  >
		   <option value="0" selected>ACTCMS主数据库</option>
		   <option value="1">Access数据源</option>
		   <option value="2">MS SQL数据源</option>
		   <option value="3">ODBC数据源</option>
		   <option value="4">Oracle数据源</option>
		   <option value="5">Excel数据源</option>
		   <option value="6">Dbase数据源</option>
	     </select>
         
         <span class="h" style="cursor:help;"  onclick="dohelp('label_datasourcetype')"  id="label_datasourcetype">帮助</span>
                </td>
    </tr>
    <tr >
      <td height="25"  align='right'><strong>连接字符串：</strong></td>
		   <td><textarea  disabled name="datasourcestr" cols="80" rows="5"><%=datasourcestr  %></textarea>
		     &nbsp;<input class='button' name="testbutton"  disabled type='button' value='测试' onclick='ajaxcheck();'>
			 <br><font color=green>说明:外部Access数据源支持相对路径,如Provider=Microsoft.Jet.OLEDB.4.0;Data Source=/数据库.mdb</font>
             </td>
    </tr>
  
  
  
  
    <tr>
      <td align="right"><strong>标签类型：</strong></td>
      <td>
	  
<input  <% IF labeltype = 0 Then Response.Write "Checked" %>  onclick='labeltypeS(this.value);'  type="radio" id="labeltype1"  name="labeltype" value="0">
        <label for="labeltype1">普通标签</label>
        <input  <% IF labeltype = 1 Then Response.Write "Checked" %>  onclick='labeltypeS(this.value);' type="radio" id="labeltype2"  name="labeltype" value="1"> 
      <label for="labeltype2">分页标签</label>
			
	  <table border="0" id="pagearea"
	  <% IF labeltype = 0 Then Response.Write "style=""display:none""" %>
	  >
			 <tr><td>
			 分页项目单位：<input type="text" class="Ainput" value="<%= ProjectUnit %>"name="ProjectUnit" size="6"> 
			 如：篇、组、个、部等</td><td width="250"> <%=actcms.ReturnPageStyle(PageStyle) %>	  </td>
			 </tr>
			 </table>	  </td>
    </tr>
    <tr>
      <td align="right"><strong>查询语句</strong>：</td>
      <td><textarea name="Sql" cols="90" rows="5" id="Sql"><%= Sql %></textarea>
      <span class="h" style="cursor:help;"  onclick="dohelp('label_free_sql')"  id="label_free_sql"><font color="red">点击这里查看语句帮助</font></span></td>
    </tr>
    <tr>
      <td align="right"><strong>标签内容</strong>：</td>
      <td><script language="javascript">
		function FieldInsertCode(fieldname,dbtype,dbname)
		{ 
		  document.ahhfchhs.LabelContent.focus();
		  var link="A.free.asp?fieldname=" + fieldname + "&dbtype="+ dbtype + "&dbname=" + dbname+"&datasourcetype=0&actcms="+ Math.random();
          var Val=showModalDialog(link,'','dialogWidth:300px; dialogHeight:260px; help: no; scroll: no; status: no');
		  var str = document.selection.createRange();
		  str.text = Val;
 		}
		function FieldInsertCode1(Val)
		{ 
		  if (Val!=''){
		   document.ahhfchhs.LabelContent.focus();
		  var str = document.selection.createRange();
		  str.text = Val;
		  SetEditorValue();
		   }
		}
		</script>

<table  width="80%" border="0" cellpadding='2' cellspacing='1' class="table">
  <tr  height='20'>
    <td align="center" style="cursor:hand;" onClick="FieldInsertCode('ID',3,0)" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">自动编号ID(Url)</td>
    <td align="center" style="cursor:hand;" onClick="FieldInsertCode('title',202,0)" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">标题</td>
    <td align="center" style="cursor:hand;" onClick="FieldInsertCode('updatetime',7,0)" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">添加/更新时间</td>
    <td align="center" style="cursor:hand;" onClick="FieldInsertCode('KeyWords',202,0)" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">关键字</td>
    </tr>
  <tr  height='20'>
    <td align="center" style="cursor:hand;" onClick="FieldInsertCode('ClassID',3,0)" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">栏目ID</td>
    <td align="center" style="cursor:hand;" onClick="FieldInsertCode('IntactTitle',202,0)" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">文章完整标题</td>
    <td align="center" style="cursor:hand;" onClick="FieldInsertCode('ArticleInput',203,0)" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">作者</td>
    <td align="center" style="cursor:hand;" onClick="FieldInsertCode('FileName',203,0)" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">生成的文件名</td>
    </tr>
  <tr  height='20'>
    <td align="center" style="cursor:hand;" onClick="FieldInsertCode('hits',3,0)" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">点击次数</td>
    <td align="center" style="cursor:hand;" onClick="FieldInsertCode('Content',203,0)" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">详细内容</td>
    <td align="center" style="cursor:hand;" onClick="FieldInsertCode('CopyFrom',202,0)" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">来源</td>
    <td align="center" style="cursor:hand;" onClick="FieldInsertCode('字段名称',202,0)" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">自定义</td>
    </tr>
</table> </td>
    </tr>
    <tr>
      <td colspan="2" align="center"><textarea name="LabelContent" cols="100%" rows="20" id="LabelContent"  wrap='on'><%=LabelContent%></textarea></td>
    </tr>
    <tr>
      <td colspan="2" align="center"><input type=button class="ACT_btn" onclick=CheckInfo()  name=Submit value=" 保 存 " />
	    <input type="reset" class="ACT_btn" name="Submit2" value="  重置  " /></td>
    </tr>
  </table>
 


 

</form>

 


<script language="javascript" >
	 CheckSel('datasourcetype','<%= datasourcetype %>');
 	 
	function CheckSel(Voption,Value)
{
	var obj = document.getElementById(Voption);
	for (i=0;i<obj.length;i++){
		if (obj.options[i].value==Value){
		obj.options[i].selected=true;
		break;
		}
	}
} 
	function ajaxcheck()
	{
  	var url=lhgajax.send("ajaxcheck.asp?A=testsource&DataType="+document.ahhfchhs.datasourcetype.value+"&str="+document.ahhfchhs.datasourcestr.value+"&m="+Math.random(),"GET");
     var DigArr=url.split('|');
		switch (DigArr[0])
		{
			 case "0":
				 alert(DigArr[1]);
				 break;
			 case "1":
				 alert('连接正常,可以使用');
				 break;
 	 
				  default:
		}
	 }
 		  function changeconnstr(datatype)
		  {
 		  
		    if (datatype==0)
			{
				document.ahhfchhs.datasourcestr.disabled=true;
				document.ahhfchhs.testbutton.disabled=true;
   			 }
			else
			{
 			document.ahhfchhs.testbutton.disabled=false;
 			document.ahhfchhs.datasourcestr.disabled=false;
 			} 
			 
		    switch (Number(datatype))
		    {
			 case 1:
			  document.ahhfchhs.datasourcestr.value='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=数据库.mdb';
 			  break;
			  
			 case 2:
  			  document.ahhfchhs.datasourcestr.value='Provider=Sqloledb; User ID=用户名; Password=密码; Initial Catalog=数据库名称; Data Source =(local);'
			  break;
			 case 3:
 		      document.ahhfchhs.datasourcestr.value='DSN=数据源名;UID=用户名;PWD=密码';
			  break;
			 case 4:
 		      document.ahhfchhs.datasourcestr.value='driver={microsoft odbc for oracle};uid=用户名;pwd=密码;server=服务器';
			  break;
			 case 5:
 		      document.ahhfchhs.datasourcestr.value='driver={microsoft excel driver (*.xls)};dbq=数据库名称';
			  break;
			 case 6:
 		      document.ahhfchhs.datasourcestr.value='driver={microsoft dbase driver (*.dbf)};dbq=数据库名称';
			  break;
 			}
		
		  }	   
	   
		  function labeltypeS(num)
		  {
		   if (num==1) 
		   {
			document.all.pagearea.style.display='';
			}
		   else
		   {
			document.all.pagearea.style.display='none';
		   }
		  }
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
        <% if datasourcetype<>0 and  action="Add" then   %>
      <script language="javascript">changeconnstr(<%= datasourcetype %>);</script>  
      <% end if  %>
      
        <% if datasourcetype<>0 and  action="EditLabel" then   %>
      <script language="javascript">
 			document.ahhfchhs.testbutton.disabled=false;
 			document.ahhfchhs.datasourcestr.disabled=false;
      
      </script>  
      <% end if  %>
      
        </body> 
</html>