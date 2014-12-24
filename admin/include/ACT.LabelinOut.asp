<!--#include file="../ACT.Function.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ACTCMS_标签目录</title>
<link href="../Images/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td class="bg_tr">您现在的位置：后台管理 &gt;&gt; 标签导入导出</td>
  </tr>
  <tr>
    <td>
<strong><a href="?A=Out">标签导出</a></strong>
	┆<strong><a href="?A=in">标签导入</a></strong>
	</td>
  </tr>
</table>
<% 
If Not ACTCMS.ACTCMS_QXYZ(0,"bqxt","") Then   Call Actcms.Alert("对不起，你没有操作权限！","")

	Dim sql, sqlCount,Sqls,intPageSize, strPageInfo,arrRecordInfo, i,pages,intPageNow,strLocalUrl,Action,Foldername,Field1
	Action=Request("A")

	  

	  Select Case Action
	  		Case "in"
				call ins()
			Case "Out"
				Call Outs()
			Case "Dout"
				Call Dout()
			Case "in2"
				Call in2()			
			Case "Doimport"
				Call Doimport()
	End  Select 

	Function LabelO(LabelType)
	  Dim AllLabel,RS
	  Set Rs = ACTCMS.ACTEXE("Select * From Label_Act " & LabelType)
	  Do While Not RS.Eof 
		AllLabel=AllLabel & "<option value='" & RS("ID") & "'>" & RS("LabelName") & "</option>"
		RS.MoveNext
	  Loop
	  RS.Close:Set RS=Nothing
	  LabelO=AllLabel
	End Function

	Sub ins()%>
		<FORM METHOD=POST ACTION="?A=in2">
			
			  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class="table" >  
			  <tr >  
			  <td height='22' align='center'><strong>标签导入（第一步）</strong></td>   
			  </tr> 
			  <tr>     
			  <td height='100' align="center">&nbsp;&nbsp;&nbsp;&nbsp;请输入要导入的标签数据库的文件名：         
	 <input name='LabelMdb' type='text' id='LabelMdb' value='<%=actcms.actsys%>Label.mdb' size='20' maxlength='50'>   
			   <input name='Submit'  class="ACT_btn" type='submit' id='Submit' value=' 下一步 '>       
			</TD>  </TR>
			</TABLE>
		</FORM>
	<%End Sub 


		Sub Doimport()
			on error resume next
			Dim n:n=0
			Dim m:m=0
			Dim k:k=0
			Dim LabelMdb:LabelMdb=ACTCMS.S("LabelMdb")
			Dim NewLabelID,cl:cl=ACTCMS.S("cl")
			Dim LabelID:LabelID=Trim(Replace(ACTCMS.S("LabelID")," ",""))
			If labelid ="" Then response.write "请选择一个标签ID":response.end
			Dim DataConn:Set DataConn = Server.CreateObject("ADODB.Connection")
			DataConn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(LabelMdb)
			If Err Then 
			Err.Clear:Set DataConn = Nothing:Response.Write "<tr><td>数据库路径不正确，连接出错</td></tr>":Response.End
			else
			 Dim rs:set rs=server.createobject("adodb.recordset")
			 rs.open "select * from Label_Act where ID in("&Trim(LabelID)&")",dataconn,1,1
			 Dim rsa:set rsa=server.createobject("adodb.recordset")
			 do while not rs.eof 
			  rsa.open "select * from Label_Act where labelname='" & rs("labelname") & "'",conn,1,3
			  if rsa.eof then
			     rsa.addnew
				 rsa("LabelName")=rs("LabelName")
				 rsa("LabelContent")=rs("LabelContent")
				 rsa("Description")=rs("Description")
				 rsa("LabelType")=rs("LabelType")
				 rsa("LabelFlag")=rs("LabelFlag")
				 rsa("AddDate")=rs("AddDate")
				 n=n+1
				rsa.update
			  else   '重名处理
			   if cl="1" then
				 rsa("LabelContent")=rs("LabelContent")
				 rsa("Description")=rs("Description")
				 rsa("LabelType")=rs("LabelType")
				 rsa("LabelFlag")=rs("LabelFlag")
				 rsa("AddDate")=rs("AddDate")
				 m=m+1
				rsa.update
			   else
			    k=K+1
			   end if
			  end if
			   rsa.close
			  rs.movenext
			 loop
			 rs.close:set rs=nothing
			 set rsa=nothing
			end if
			response.write "<br><br><br><div align=center>操作完成!成功导入了 <font color=red>" & n & "</font> 个标签,覆盖了 <font color=red>" & m & "</font> 个标签,重名跳过了 <font color=red>" & k & "</font> 个标签！  </div><br><br><br><br><br><br><br>"
           dataconn.close:set dataconn=nothing
		End Sub






	Function Label1(LabelType,DataConn)
	  Dim AllLabel,RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	  RS.Open "Select * From Label_Act " & LabelType,DataConn,1,1
	  Do While Not RS.Eof 
		AllLabel=AllLabel & "<option value='" & RS("ID") & "'>" & RS("LabelName") & "</option>"
		RS.MoveNext
	  Loop
	  RS.Close:Set RS=Nothing
	  Label1=AllLabel
	End Function


	Sub in2()
		Dim LabelMdb,LabelType:LabelMdb=ACTCMS.S("LabelMdb")
		Dim DataConn:Set DataConn = Server.CreateObject("ADODB.Connection")
	    DataConn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(LabelMdb)
		%>
		<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='table'> 
		<form name='myform' method='post' action='?A=Doimport'>  
		<tr class='title'>  
	 <td height='22' align='center'><strong>标签导入（第二步）</strong></td> 
	 </tr>
	 <tr > 
	<td height='100' align='center'>      
	 <br>       
	 <table border='0' cellspacing='0' cellpadding='0'>          
		<%
		If Err Then 
		Err.Clear:Set DataConn = Nothing:Response.Write "<tr><td>数据库路径不正确，连接出错</td></tr>":Response.End
		else
		 	%>
		  <Script language="Javascript">
			var ClassArr = new Array();
			  ClassArr[1] =new Array("<%=Label1("  Where LabelType=1",DataConn)%>");
			  ClassArr[2] =new Array("<%=Label1("  Where LabelType=2",DataConn)%>");
			  ClassArr[3] =new Array("<%=Label1("  Where LabelType=3",DataConn)%>");
			  ClassArr[9999] =new Array("<%=Label1(" ",DataConn)%>");
		  </Script>
		<tr> 
		<td colspan="2">
		<strong>选择要导入的标签的分类：</strong>
		<select id="LabelType" name="LabelType" onChange="SelectClass(this.value)">
			  <option value="9999">全部标签</option>
			  <option value="1"<%IF LabelType="1" Then Response.write " selected"%>>系统函数标签</option>
			  <option value="2"<%IF LabelType="2" Then Response.write " selected"%>>自定义静态标签</option>
			  <option value="3"<%IF LabelType="3" Then Response.write " selected"%>>自由标签</option>
		    </select></td></tr>   
  		<tr>
		<td colspan="2"><strong>重名处理方式：</strong> 
		<input type="radio" value="0" name="cl" id="cl1" checked><label for="cl1">标签重名跳过</label>
		<input type="radio" value="1" name="cl" id="cl2"><label for="cl2">标签重名覆盖</label></td>
		</tr>  

		<tr>
		<td id="ClassArea"><select name='select' size='2' multiple style='height:300px;width:350px;'>
        </select></td>
		<td width="20" align="left">
		<input type='button' class="ACT_btn" name='Submit' value=' 选定所有 ' onclick='SelectAll()'>   
	   <br><br>&nbsp;&nbsp;&nbsp;&nbsp;
	   <input type='button' class="ACT_btn" name='Submit' value=' 取消选定 ' onclick='UnSelectAll()'>
	   <br><br><br><b>&nbsp;</b>		    </td>
		</tr>  
		<%end if%>         
		<tr><td colspan='4' height='25' align='center'>
		<input type='submit' name='Submit' class="ACT_btn" value=' 导入标签 '>
		</td></tr>
		 </table>
		<input name='LabelMdb' type='hidden' id='LabelMdb' value='<%=LabelMdb%>'> 
         <br>   <br>  <br>          
	   </td>          </tr>       
		</form></table><br>  <br>  
		<script language='javascript'>
		  SelectClass(9999);
	function SelectClass(LabelType)
	{ document.all.ClassArea.innerHTML='<select name="LabelID" size="2" multiple style="height:300px;width:450px;">'+ClassArr[LabelType]+'</select>';
	}
   </script>
   <%
	End Sub 
Sub Outs() 
Dim LabelType

%>
<table width="100%" border="0" cellpadding="1" cellspacing="1" class="table">
<form name="myform" method="post" action="?A=Dout">
		  <Script language="Javascript">
			var ClassArr = new Array();
			  ClassArr[1] =new Array("<%=LabelO("  Where LabelType=1")%>");
			  ClassArr[2] =new Array("<%=LabelO("  Where LabelType=2")%>");
			  ClassArr[3] =new Array("<%=LabelO("  Where LabelType=3")%>");
			  ClassArr[9999] =new Array("<%=LabelO(" ")%>");
		  </Script>
  <tr>
    <td align="right">标签目录名称：</td>
    <td>
  <select id="LabelType" name="LabelType" onChange="SelectClass(this.value)">
			  <option value="9999">全部标签</option>
			  <option value="1"<%IF LabelType="1" Then Response.write " selected"%>>系统函数标签</option>
			  <option value="2"<%IF LabelType="2" Then Response.write " selected"%>>自定义静态标签</option>
			  <option value="3"<%IF LabelType="3" Then Response.write " selected"%>>自由标签</option>
			</select>
</td>
  </tr> <tr >      
		  <td colspan="2" align='center'>        
		    <table width="100%" border='0' cellpadding='0' cellspacing='0'>          
			   <tr>           
			     <td width="10%" align="right">标签列表：</td>
				 <td width="54%" ID="ClassArea">
				 <select name='LabelID' size='2' multiple style='height:300px;width:450px;'>
				 </select></td>                 <td width="36%" align='left'>&nbsp;&nbsp;&nbsp;&nbsp;
				   <input type='button' class="ACT_btn" name='Submit' value=' 选定所有 ' onclick='SelectAll()'>   
				   <br><br>&nbsp;&nbsp;&nbsp;&nbsp;
				   <input type='button' class="ACT_btn" name='Submit' value=' 取消选定 ' onclick='UnSelectAll()'>
				   <br><br><br><b>&nbsp;提示：按住“Ctrl”或“Shift”键可以多选</b></td>      
			 </tr>     
			 <tr height='30'>        <td colspan='2'>　目标数据库：
		 <input name='LabelMdb' type='text' id='LabelMdb' value='<%=actcms.actsys%>Label.mdb' size='20' maxlength='50'>
			 &nbsp;&nbsp;此操作将清空目标数据库</td>      
			 </tr>
  <tr>
    <td colspan="2" align="center"><input type=button onclick=CheckForm() class="ACT_btn"  name=Submit1 value="  保存  " />
        &nbsp;&nbsp;&nbsp;&nbsp;<input name="Submit2" type="reset" class="ACT_btn" value="  重置  ">
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
    </tr>
</form>

	  <script language='javascript'>
		  SelectClass(9999);
function SelectClass(LabelType)
{ document.all.ClassArea.innerHTML='<select name="LabelID" size="2" multiple style="height:300px;width:450px;">'+ClassArr[LabelType]+'</select>';
}
</script>
		  
		  
  </table>

<%end sub 



	Sub Dout()
	 Dim LabelID:LabelID=Trim(ACTCMS.S("LabelID"))
	 Dim LabelMdb:LabelMdb=ACTCMS.S("LabelMdb")
	 Dim rs:set rs=server.createobject("adodb.recordset")
	 If labelid ="" Then response.write "请选择一个标签ID":response.end
	 Dim sqlstr,n
	   n=0
	   sqlstr="select ID,LabelName,LabelContent,Description,LabelFlag,LabelType,AddDate from Label_Act where id in(" & LabelID & ")"
			 on error resume next
			 if CreateDatabase(LabelMdb)=true then
					Dim DataConn:Set DataConn = Server.CreateObject("ADODB.Connection")
					DataConn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(LabelMdb)
					If not Err Then
					   If Checktable("Label_Act",DataConn)=true Then
						 DataConn.Execute("drop table Label_Act")
					   end if
				             Dataconn.execute("CREATE TABLE [Label_Act] ([ID] int IDENTITY (1, 1) NOT NULL CONSTRAINT PrimaryKey PRIMARY KEY,[LabelName] varchar(255) Not Null,[LabelContent] text not null,[Description] text null,[LabelType] int not null,[LabelFlag] int not null,[AddDate] date not null)")
					  rs.open sqlstr,conn,1,1
					 if not rs.eof then
						Dim RST:Set RST=Server.CreateObject("ADODB.RECORDSET")
					   do while not rs.eof
						  n=n+1
						  RST.Open "Select * From Label_Act where 1=0",DataConn,1,3
						  RST.AddNew
							RST("LabelName")=rs(1)
							RST("LabelContent")=rs(2)
							RST("Description")=rs(3)
							RST("LabelFlag")=rs(4)
							RST("LabelType")=rs(5)
							RST("AddDate")=rs(6)
						  RST.Update
						  RST.Close
						  rs.movenext
					   loop
					   Set RST=Nothing
					 end if
					  rs.close:set rs=nothing
					End if
					DataConn.Close:Set DataConn=Nothing
			 end if
			response.write "<br><br><br><div align=center>操作完成!成功导出了 <font color=red>" & n & "</font> 个标签！<a href=" & LabelMdb & ">请点击这里下载</a>(右键目标另存为)  </div><br><br><br><br><br><br><br>"
	End Sub
		Function Checktable(TableName,DataConn)
			On Error Resume Next
			DataConn.Execute("select * From " & TableName)
			If Err.Number <> 0 Then
				Err.Clear()
				Checktable = False
			Else
				Checktable = True
			End If
		End Function
		Function CreateDatabase(dbname)
			  Dim fso 
			  Set Fso = Server.CreateObject("scripting.FileSystemObject")
			  If  Fso.FileExists(Server.MapPath(dbname)) Then
				  CreateDatabase = true
				  Exit Function
			  End If
				dim objcreate :set objcreate=Server.CreateObject("adox.catalog") 
				if err.number<>0 then 
					set objcreate=nothing 
					CreateDatabase=false
					exit function 
				end if 
				'建立数据库 
				objcreate.create("data source="+server.mappath(dbname)+";provider=microsoft.jet.oledb.4.0") 
				if err.number<>0 then 
					CreateDatabase=false
					set objcreate=nothing 
					exit function
				end if 
				CreateDatabase=true
		End Function

CloseConn %>

<script language="javascript">
function SelectAll(){
  for(var i=0;i<document.myform.LabelID.length;i++){
    document.myform.LabelID.options[i].selected=true;}
}
function UnSelectAll(){
  for(var i=0;i<document.myform.LabelID.length;i++){
    document.myform.LabelID.options[i].selected=false;}
}
function CheckForm()
{ var form=document.myform;
	
	 if (form.LabelMdb.value=='')
		{ alert("请输入目标数据库!");   
		  form.LabelMdb.focus();    
		   return false;
		} 
		form.Submit1.value="正在提交数据,请稍等...";
		form.Submit1.disabled=true;
		form.Submit2.disabled=true;		
	    form.submit();
        return true;
	}</script> 

</body>
</html>
