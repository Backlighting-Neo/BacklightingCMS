<!--#include file="../ACT.Function.asp"-->
<!--#include file="../../ACT_inc/cls_pageview.asp"-->
<!--#include file="../include/ACT.F.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>文章管理</title>
<link href="../Images/style.css" rel="stylesheet" type="text/css">
 <script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/lhgcore.min.js"></script>
 	<script type="text/javascript">
		var DG = frameElement.lhgDG;
   	</script>

<%
'--------禁止缓存------------  
Response.Expires   =   -1   
Response.ExpiresAbsolute   =   Now()   -   1   
Response.cachecontrol   =   "no-cache"   
%>
</head>
<body>
<% 
	if request("Action")="save" Then
   	
	 
 	  echo"<script>DG.curWin.insertHTMLToEditor('"&trim(Replace(request.form("ID")," ",""))&"');</script>" 
		 echo "<script>DG.cancel()</script>"
	  ' Response.Write "<script>D.document.getElementById('"&request("iname")&"').value="""&trim(Replace(request.form("ID")," ",""))&""";</script>"
	  ' Response.Write "<script>window.parent.cancel()</script>"
     response.end
    end if 
  	Dim ShowErr,ModeID,ModeName,Item,pages,urlact,ClassID,Action,ID,Page,url
	Dim LmID,gltj,title
	 
	 ModeID = ChkNumeric(Request("ModeID"))
	 if ModeID=0 or ModeID="" Then ModeID=1
	 ModeName= ACTCMS.ACT_C(ModeID,1) 
	 ClassID=Request("ClassID")
	 LmID=ClassID
	 
	 Page=ChkNumeric(Request("page"))
	 If Page = 0 Then Page = 1
	 ID = Request("ID")
	 title=request("title")
	 Action=Request("Action")

  	 urlact = "Action="&Action&"&title="&title&"&ClassID="&ClassID&"&ModeID="&ModeID&"&page="&Page
  	 url = "&ClassID="&LMID&"&ModeID="&ModeID&"&page="&Page

	
 	
 	Dim strLocalUrl
	strLocalUrl = request.ServerVariables("SCRIPT_NAME")
	
	Dim intPageNow
	intPageNow = request("page")
	
	Dim intPageSize, strPageInfo
	intPageSize = 20
	
	Dim arrRecordInfo, i
	Dim sql, sqlCount,Sqls
	 
	 pages = "Action="&Request("Action")&"&title="& request("title")&"&ClassID="& request("ClassID")&"&UserID="& request("UserID")&"&ModeID="&ModeID&"&page"
	'If  ClassID<>"" Then ClassID=" classid='"&Request.QueryString("ClassID")&"' and  "
	If  LmID<>"" Then gltj=" classid='"&LmID&"' and  "
	
	Select Case Request.QueryString("Action")
 		Case "ListisAccept"
			Sqls = "  where "&gltj&"  isAccept=0 and delif=0 "
		Case "ListNoAccept"
			Sqls = "  where "&gltj&"  isAccept=2 and delif=0 "
		Case "Listcg"
			Sqls = "  where "&gltj&"  isAccept=1 and delif=0 "
		Case "Listtg"
			Sqls = "  where "&gltj&"  isAccept=3 and delif=0 "
		Case "Lististop"
			Sqls = "  where "&gltj&"  istop=1 and delif=0 "
		Case "Listpic"
			Sqls = "  where "&gltj&"   picurl<>'' and delif=0 "
		Case "Slide"
			Sqls = "  where "&gltj&"  Slide=1 and delif=0 "
		Case "sh"
			Sqls = "  where "&gltj&"  isAccept>0 and delif=0 "
		Case "MyArticle"
			Sqls="  where "&gltj&"  ArticleInput='"&RSQL(Request.Cookies(AcTCMSN)("AdminName"))&"' and delif=0 "
		Case "UserID"
			Sqls="  where "&gltj&"  ArticleInput='"&RSQL(request("UserID"))&"' and delif=0 "
		Case "t"
			If Request("ClassID")="" Then 
				Sqls = " where title Like '%" & RSQL(request("title")) & "%' "
			Else 
				Sqls = " where ClassID='"&LmID&"' and title Like '%" & request("title") & "%' "
			End If 
		Case Else
			Sqls = " where delif=0 and  isAccept=0 "
			pages = "ModeID="&ModeID&"&iname="&request("iname")&"&page"
	End Select
	
	sql = "SELECT [ID], [Title], [ArticleInput], [Hits], [isAccept], [IStop],  [ClassID], [FileName],[ACTLINK],[Ismake],[InfoPurview],[ReadPoint],[picurl],[Slide],[ATT],[rev],[UserID]" & _
		" FROM ["&ACTCMS.ACT_C(ModeID,2)&"]" &Sqls& _
		"ORDER BY [UpdateTime] DESC,[ID] DESC "
   	sqlCount = "SELECT Count([ID])" & _
			" FROM ["&ACTCMS.ACT_C(ModeID,2)&"]"&Sqls
		Dim clsRecordInfo
		Set clsRecordInfo = New Cls_PageView
			clsRecordInfo.intRecordCount = 2816
			clsRecordInfo.strSqlCount = sqlCount
			clsRecordInfo.strSql = sql
			clsRecordInfo.intPageSize = intPageSize
			clsRecordInfo.intPageNow = intPageNow
			clsRecordInfo.strPageUrl = strLocalUrl
			clsRecordInfo.strPageVar = pages
		clsRecordInfo.objConn = Conn		
		arrRecordInfo = clsRecordInfo.arrRecordInfo
		strPageInfo = clsRecordInfo.strPageInfo
		Set clsRecordInfo = nothing
 	 %>
<table width="99%" border="0" align="center" cellpadding="2" cellspacing="1"  class="table">
  <tr>
    <td  class="bg_tr"><strong>您现在的位置：<%= ModeName %>系统管理 &gt;&gt; <%= ModeName %>管理</strong>
  选择模型:<select name='ModeID' style='width:110px' onChange="location=this.value;">
  <%=AF.ACT_URL_Mode(ModeID,"&ClassID="&request("ClassID"))%>
  </select>
  
  </td>
  </tr>

  <tr>
    <td ><%= ModeName %>选项：<a href="ACT.Add.asp?Action=add&ModeID=<%=ModeID%>&ClassID=<%=request("ClassID")%>"></a>┆<a href="?ModeID=<%=ModeID%>">所有文章</a> ┆ <a href="?ModeID=<%=ModeID%>&Action=ListisAccept">已审</a>┆ <a href="?ModeID=<%=ModeID%>&Action=ListNoAccept">未审</a>┆ <a href="?ModeID=<%=ModeID%>&Action=Listcg">草稿</a>   ┆ <a href="?ModeID=<%=ModeID%>&Action=Listtg">退稿</a>┆<a href="?ModeID=<%=ModeID%>&Action=Lististop">固顶<%= ModeName %></a> ┆<a href="?ModeID=<%=ModeID%>&Action=Listpic">图片<%= ModeName %></a>┆<a href="?ModeID=<%=ModeID%>&Action=Slide">幻灯</a>┆</td>
  </tr>
</table>

<table width='99%'  border='0' align="center" cellpadding='0' cellspacing='1' class="table">
   <tr> 
          <form name='form3' Action='?Action=t&ClassID=<%=Request.QueryString("ClassID")%>' method='post'>
 
          <td>
          	<table width='720' border='0' cellpadding='0' cellspacing='0'>
          	<tr>
          	<td width='90' align='center'>请选择栏目：</td>
          	<td width='160'>
<select name="ClassID" size="1" onChange="javascript:window.location=this.options[this.selectedIndex].value;">
<option value='?ModeID=<%=ModeID%>' >全部</option>
      <% 	 Response.Write Classmake(ModeID)
		%>
    </select>        </td>
        <td width='70'>
          关键字：        </td>
          <td width='160'>
          	<input type='text' name='title' value='' class="Ainput" style='width:150'>  </td>
          <td>
            <input  name="Submit"  type="submit" class="ACT_btn" value="  搜索  ">
		  </td>
          </tr>
        </table>
          </td>
        </form>
  </tr>
</table>

  <table width="99%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
<form name="Article" method="post" Action="?Action=save&iname=<%=request("iname")%>">
 
    <tr>
      <td width="28" align="center" class="bg_tr">选中 </td>
      <td width="40" align="center" class="bg_tr">ID</td>
      <td width="160" align="center" class="bg_tr">文章标题</td>
	  <td width="60" align="center" class="bg_tr">栏目</td> 
	  <td width="60" align="center" class="bg_tr">录入者</td> 
	  <td width="40" align="center" class="bg_tr">点击数</td> 
	  <td width="28" align="center" class="bg_tr">审核</td> 
	  <td width="50" align="center" class="bg_tr">生成</td> 
    </tr>
	 <%
		Dim bgColor
		If IsArray(arrRecordInfo) Then
			For i = 0 to UBound(arrRecordInfo, 2)
			bgColor="#FFFFFF"
			if i mod 2=0 then bgColor="#DFEFFF"
			Dim Rs ,ClassName,ClasseName
			  Set Rs = server.CreateObject("adodb.recordset")
					Rs.Open "select ClassName,ClasseName from Class_Act where ClassID='"& arrRecordInfo(6,i) &"'",Conn,1,1
					if  Not Rs.bof then
							ClassName =Rs("ClassName")
							ClasseName =Rs("ClasseName")
					Else
							ClassName ="<font color=red>程序出现错误</font>"
							ClasseName ="<font color=red>意外错误</font>"
					End if 
	%>
    <tr onMouseOver=overColor(this) onMouseOut=outColor(this)>
      <td align="center" >
	  <input type="checkbox" name="ID" value="<%= arrRecordInfo(0,i) %>">	  </td>
      <td align="center" ><%= arrRecordInfo(0,i) %></td>
      <td ><a target="_blank" href="<%
	  ID = arrRecordInfo(0,i)
	  
	  If arrRecordInfo(8,i) = 1 Then
		response.write arrRecordInfo(7,i)
		Else
	 response.write  ACTCMS.actsys&"List.asp?C-"&ModeID&"-"&arrRecordInfo(0,i)&".Html"
	  End if%>"><%= arrRecordInfo(1,i) %></a>&nbsp;<%
	  If arrRecordInfo(12,i)<>"" Then 
			response.write "<font color=red title=""图片文章""  style=""cursor:default"">图</font>"
	  End If 
	  If arrRecordInfo(13,i)=1 Then 
			response.write "&nbsp;<font color=green title=""幻灯片文章""  style=""cursor:default"">幻</font>"
	  End If 

	  If arrRecordInfo(8,i)=1 Then 
			response.write "&nbsp;<font color=green title=""转向链接""  style=""cursor:default"">转</font>"
	  End If 

	  %></td>
      
	  <td align="center">
	 <a href="?Action=t&ModeID=<%=ModeID%>&ClassID=<%= arrRecordInfo(6,i) %>"><%= classname %></a>	  </td>
    
	  <td align="center">
	  <a href="?Action=UserID&ModeID=<%=ModeID%>&UserID=<%= arrRecordInfo(2,i) %>"><%= arrRecordInfo(2,i) %></a></a></td>
  
      <td align="center" ><%= arrRecordInfo(3,i) %></td>
	 
  
	 
      <td align="center" ><% '0已审  1草稿  2待审 3退稿
	  
	   Select Case(arrRecordInfo(4,i))
	  			Case 0
					response.Write "<a title='取消审核' href='?A=sh&cs=2&ID="&ID&url&"'>已审</a>"	
				Case 1
					response.Write "<a title='置为待审' href='?A=sh&cs=2&ID="&ID&url&"'><font color=red>草稿</font></a>"
				Case 2
					response.Write "<a title='置为已审' href='?A=sh&cs=0&ID="&ID&url&"'><font color=red>待审</font></a>"	
				Case 3
					response.Write "<a title='置为待审' href='?A=sh&cs=2&ID="&ID&url&"'><font color=red>退稿</font></a>"
		 End Select	
	   %></td>
	   
      <td align="center" >
	<%IF arrRecordInfo(9,i) = 1 Then 
	  response.Write "<font color=red><b>√</b></font>"
	  Else
	  response.Write "<font color=red><b>×</b></font>"
	  End If %>	  </td>
    </tr>
	<% 
	Next
	End If
	%>
    <tr >
      <td height="30" colspan="9" >
	  
	  
	  <label for="chkAll"><input name="ChkAll" type="checkbox" id="ChkAll" onClick="CheckAll(this.form)" value="checkbox">
		&nbsp;选中本页显示的所有文章</label>
		
		<input name="Submit2" type="submit" class="ACT_btn" value="保存并退出">		</td>
    </tr>
    <tr >
      <td height="66" colspan="9" align="center" ><%= strPageInfo%></td>
    </tr></form>
  </table>

<script language="javascript">
 
//CSS背景控制
function overColor(Obj)
{
	var elements=Obj.childNodes;
	for(var i=0;i<elements.length;i++)
	{
		elements[i].className="tdbg1"
		Obj.bgColor="";
	}
	
}

 


function CheckAll(form)
		  {  
		 for (var i=0;i<form.elements.length;i++)  
			{  
			   var e = Article.elements[i];  
			   if (e.name != 'ChkAll'&&e.type=="checkbox")  
			   e.checked = Article.ChkAll.checked;  
		   }  
	  }

function outColor(Obj)
{
	var elements=Obj.childNodes;
	for(var i=0;i<elements.length;i++)
	{
		elements[i].className="tdbg";
		Obj.bgColor="";
	}
}
 
</script>


<% 
	Function  ACT_CL(ID)
		ACT_CL=ACTCMS.ACTEXE("Select ClassID From "&ACTCMS.ACT_C(ModeID,2)&" Where ID=" & ID)(0)
	End Function  
	Function  Classmake(ModeID)
		 Dim FolderRS,selected
		 Set FolderRS = actcms.actexe("Select * from Class_act where ParentID='0' and ACTLINK=1 Order by Orderid desc,ID desc")
		 IF FolderRS.Eof Then
		 Classmake=("<option value='?ModeID="&ModeID&"'>还没有添加任何栏目</option>")
		 Else 
		 do while Not FolderRS.Eof
			If Request("ClassID")=FolderRS("ClassID") Then selected=" selected=""selected""" Else selected=""
			 Classmake=Classmake&"<option value='?Action=t&ModeID="&FolderRS("ModeID")&"&ClassID="&FolderRS("ClassID")&"' "&selected&">"& FolderRS("ClassName") & "</option>"& vbCrLf
			 Classmake=Classmake&(GetChildClassList(FolderRS("ClassID"),""))
			 FolderRS.MoveNext
		 Loop
		 End IF

	 End  Function 
	 Function GetChildClassList(ClassID,Str)
	       Dim Sql,RsTempObj,TempImageStr,ImageStr,CheckStr,selected
	        TempImageStr = "&nbsp;└"
	        Sql = "Select * from Class_act where ParentID='" & ClassID & "'  and ACTLINK=1"
	        ImageStr = Str & "&nbsp;└"
	        Set RsTempObj = Conn.Execute(Sql)
	            do while Not RsTempObj.Eof
					If Request("ClassID")=RsTempObj("ClassID") Then selected=" selected=""selected""" Else selected=""
					   GetChildClassList = GetChildClassList  & "<option value='?Action="&Action&"&ModeID="&RsTempObj("ModeID")&"&ClassID="&RsTempObj("ClassID")&"' "&selected&">"& ImageStr & TempImageStr &" "& RsTempObj("ClassName")& "</option>"& vbCrLf
					   GetChildClassList = GetChildClassList & GetChildClassList(RsTempObj("ClassID"),ImageStr)
					   RsTempObj.MoveNext
	           loop
	       Set RsTempObj = Nothing
	 End Function 

CloseConn 
%>
</body>
</html>