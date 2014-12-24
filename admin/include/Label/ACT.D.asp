<!--#include file="../../ACT.Function.asp"-->
<%
If Not ACTCMS.ACTCMS_QXYZ(0,"bqxt","") Then   Call Actcms.Alert("对不起，你没有操作权限！","")
Dim Work,IdList,IsSelected,ModeID
    Work = Request("Work")
Select Case Work
    Case "ReturnValue"
        ReturnValue()
End Select
IdList =  Request("IdList")
IdList =Replace(IdList, "'", "")
Dim SelectorType
    SelectorType = Request("Type")
If SelectorType = "" Then
    SelectorType = 2
Else
    SelectorType = CInt(SelectorType)
End If
 ModeID = ChkNumeric(Request("ModeID"))
%>
<html>
<head>
<base target="_self">
<title>栏目选择</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<SCRIPT src="../../../ACT_inc/dtreeFunction.js"></SCRIPT>
<LINK href="../../../ACT_inc/dtree.css" type=text/css rel=StyleSheet>
<SCRIPT src="../../../ACT_inc/dtree.js" type=text/javascript></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
<!--
var SelectorType = <%=SelectorType%>;
function chkForm(obj)
{
    if(SelectorType == 1)
    {
        if(!GetRadioBox("radiod"))
        {
            return false;
        }       
    }else{
        if(!GetCheckBoxList("checkboxd"))
        {
            return false;
        }
    }
    return true;
}
//-->
</SCRIPT>
<style type="text/css">
<!--
BODY {
    SCROLLBAR-HIGHLIGHT-COLOR: buttonface;
    SCROLLBAR-SHADOW-COLOR: buttonface;
    SCROLLBAR-3DLIGHT-COLOR: buttonhighlight;
    SCROLLBAR-TRACK-COLOR: #eeeeee;
    SCROLLBAR-DARKSHADOW-COLOR: buttonshadow;
    background-color:buttonface;
    font:12px;

    margin: 3px;
    padding: 0px;
    border: none;
}
-->
</style>

<body scroll="no">

<form name="form1" method="post"  action="ACT.D.asp?Work=ReturnValue"  onsubmit="return chkForm(this)">
  <table width="100%" height="100%" border="0" cellpadding="5" cellspacing="0">
    <tr> 
      <td> 
        <div style="width:100%;height:100%;overflow:auto;background-color:#ffffff;padding:3px;">
        <%InitTree()%>
        </div>
      </td>
    </tr>
    <tr>
      <td height="22" align="right"> <input type="submit" name="Button" value="确　定"> 
	  <input type="button" name="Button2" value="取　消" onclick="window.close();">
        <input name="Work" type="hidden" id="Work" value="ReturnValue">
        <input name="Type" type="hidden" id="Type" value="<%=SelectorType%>">
        </td>
    </tr>
  </table>
</form>
</body>
</html>
<%

	Public Function FoundInArr(strArr, strToFind, strSplit)
		Dim arrTemp, i
		FoundInArr = False
		If InStr(strArr, strSplit) > 0 Then
			arrTemp = Split(strArr, strSplit)
			For i = 0 To UBound(arrTemp)
			If LCase(Trim(arrTemp(i))) = LCase(Trim(strToFind)) Then
				FoundInArr = True:Exit For
			End If
			Next
		Else
			If LCase(Trim(strArr)) = LCase(Trim(strToFind)) Then FoundInArr = True
		End If
	End Function
	Function InitTree()

    Response.Write "        <script type=""text/javascript"">" & vbCrLf
    Response.Write "        <!--" & vbCrLf
    Response.Write "        d = new dTree('d');" & vbCrLf
    If SelectorType = 1 Then
        Response.Write "        d.config.inputType = 1;" & vbCrLf
    Else
        Response.Write "        d.config.inputType = 2;" & vbCrLf   
    End If
    Response.Write "			d.config.useIcons = true;" & vbCrLf
	Response.Write "        d.add(0, -1, '栏目列表',null,null,null,null);" & vbCrLf
	response.write  Classmake
    Response.Write "        document.write(d);" & vbCrLf
    Response.Write "        //-->" & vbCrLf
    Response.Write "        </script>" & vbCrLf
End Function 

	function Classmake
		 Dim FolderRS,ModeIDs
		  if  ModeID="0" Then ModeIDs="" Else ModeIDs= " And ModeID= "&CInt(ModeID)&" "
		 Set FolderRS = Conn.Execute("Select * from Class_act where ParentID='0'  and actlink<>2 Order by Orderid desc,ID desc")
		 IF FolderRS.Bof And FolderRS.Eof Then
		 End IF
		 do while Not FolderRS.Eof
			If FoundInArr(IdList,FolderRS("ClassID"),",")=False then
				IsSelected = "false"
			Else
				IsSelected = "true"
			End If
			  Response.Write "        d.add(" & FolderRS("ClassID") & ",0,'" & FolderRS("ClassName") & "',null,null,null,null,null,null,0," & IsSelected & ",'" & FolderRS("ClassID") & "');" & vbCrLf
			  Classmake=Classmake&(GetChildClassList(FolderRS("ClassID")))
			  FolderRS.MoveNext
		  loop
	 End function
	 Function GetChildClassList(ClassID)
	       Dim Sql,RsTempObj,CheckStr
	        Sql = "Select * from Class_act where ParentID='" & ClassID & "'  and actlink<>2 "
	        Set RsTempObj = Conn.Execute(Sql)
			do while Not RsTempObj.Eof
				If FoundInArr(IdList,RsTempObj("ClassID"),",")=False then
					IsSelected = "false"
				Else
					IsSelected = "true"
				End If
				GetChildClassList = GetChildClassList & GetChildClassList(RsTempObj("ClassID"))
				Response.Write "        d.add(" & RsTempObj("ClassID") & "," & RsTempObj("ParentID") & ",'" & RsTempObj("ClassName") & "',null,null,null,null,null,null,0," & IsSelected & ",'" & RsTempObj("ClassID") & "');" & vbCrLf
			 RsTempObj.MoveNext
		   loop
		   Set RsTempObj = Nothing
	 End Function 
 
Function ReturnValue()
	Dim SelectorType,ClassIDf,i,CQ
    SelectorType = CInt(Request("Type"))
    If SelectorType = 1 Then
        ClassIDf = Split(Trim(Request("radiod")),",")
    Else
        ClassIDf = Split(Trim(Request("checkboxd")),",")
    End If
	If UBound(ClassIDf)<>0 Then CQ= "'"
	Response.Write "<html><script>" & vbCrLf
	Response.Write "var result = Array(" & vbCrLf
	 For I = LBound(ClassIDf) To UBound(ClassIDf)
        Response.Write "    {" & vbCrLf
        Response.Write "        id:"""&CQ&Trim(ClassIDf(i))&CQ&""","& vbCrLf
        Response.Write "        parent:"""","  & vbCrLf
        Response.Write "        title:"""","  & vbCrLf
        Response.Write "        creator:"""","  & vbCrLf
        Response.Write "        show:"""","  & vbCrLf
        Response.Write "        addtime:"""""  & vbCrLf
        If i=>UBound(ClassIDf) Then
            Response.Write "    }" & vbCrLf
        Else
            Response.Write "    }," & vbCrLf
        End If
	Next
	 Response.Write ");" & vbCrLf
    Response.Write "window.returnValue = result;window.close();" & vbCrLf
    Response.Write "</script></html>" & vbCrLf
	response.End
End Function
%>
