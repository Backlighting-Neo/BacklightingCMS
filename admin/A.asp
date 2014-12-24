<!--#include file="ACT.Function.asp"-->
<%
If Not ACTCMS.ACTCMS_QXYZ(0,"bqxt","") Then   Call Actcms.Alert("对不起，你没有操作权限！","")
Dim ModeID

 ModeID = Request("ModeID")
%>
<html>
<head>
<base target="_self">
<title>栏目选择</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<LINK href="../ACT_inc/dtree.css" type=text/css rel=StyleSheet>
<SCRIPT src="../ACT_inc/dtree.js" type=text/javascript></SCRIPT>


<body >
<a href="javascript: d.openAll();">全部展开</a> | <a href="javascript: d.closeAll();">全部折叠</a></p>


        <div style="width:100%;height:100%;overflow:auto;background-color:#ffffff;padding:3px;">
        <%InitTree()%>
        </div>
   
</form>
</body>
</html>
<%

	Function InitTree()

    Response.Write "        <script type=""text/javascript"">" & vbCrLf
    Response.Write "        <!--" & vbCrLf
    Response.Write "        d = new dTree('d');" & vbCrLf
    Response.Write "        d.config.inputType = 0;" & vbCrLf
    Response.Write "			d.config.useIcons = true;" & vbCrLf
	Response.Write "        d.add(0, -1, '栏目列表',null,null,null,null);" & vbCrLf
	response.write  Classmake
    Response.Write "        document.write(d);" & vbCrLf
    Response.Write "        //-->" & vbCrLf
    Response.Write "        </script>" & vbCrLf
End Function 

	function Classmake
		 Dim FolderRS,ModeIDs
		 Set FolderRS = Conn.Execute("Select * from Class_act where ParentID='0'   and ChangesLinkUrl='' Order by Orderid desc,ID desc")
		 IF FolderRS.Bof And FolderRS.Eof Then
		 Response.Write("<br><li>还没有添加任何栏目!")
		 End IF
		 do while Not FolderRS.Eof
			  IsSelected = "true"
			  Response.Write "        d.add(" & FolderRS("ClassID") & ",0,'" & FolderRS("ClassName") & "','ACT_Mode/ACT.Manage.asp?ModeID="&FolderRS("Modeid")&"','"&FolderRS("ClassName")&"','main',null,null,null,0," & IsSelected & ",'" & FolderRS("ClassID") & "');" & vbCrLf
			  Classmake=Classmake&(GetChildClassList(FolderRS("ClassID")))
			  FolderRS.MoveNext
		  loop
	 End function
	 Function GetChildClassList(ClassID)
	       Dim Sql,RsTempObj,CheckStr
	        Sql = "Select * from Class_act where ParentID='" & ClassID & "'  and ChangesLinkUrl=''"
	        Set RsTempObj = Conn.Execute(Sql)
			do while Not RsTempObj.Eof
				IsSelected = "true"
				GetChildClassList = GetChildClassList & GetChildClassList(RsTempObj("ClassID"))
				Response.Write "        d.add(" & RsTempObj("ClassID") & "," & RsTempObj("ParentID") & ",'" & RsTempObj("ClassName") & "','ACT_Mode/ACT.Manage.asp?ModeID="&RsTempObj("Modeid")&"','"&RsTempObj("ClassName")&"','main',null,null,null,null,0," & IsSelected & ",'" & RsTempObj("ClassID") & "');" & vbCrLf
			 RsTempObj.MoveNext
		   loop
		   Set RsTempObj = Nothing
	 End Function 
 

%>
