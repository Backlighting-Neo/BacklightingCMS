<!--#include file="../ACT.Function.asp"-->
<!--#include file="../include/ACT.F.asp"-->
<HTML><HEAD><TITLE>设置属性</TITLE>
<META http-equiv=Content-Type content="text/html; chaRSet=utf-8">
<LINK href="../Images/style.css" rel=stylesheet>
<SCRIPT language=javascript>
function SelectAll(){
			  for(var i=0;i<document.myform.ClassID.length;i++){
				document.myform.ClassID.options[i].selected=true;}
			}
			function UnSelectAll(){
			  for(var i=0;i<document.myform.ClassID.length;i++){
				document.myform.ClassID.options[i].selected=false;}
			}				
</SCRIPT>
<META content="MSHTML 6.00.2900.3132" name=GENERATOR>
</HEAD>
<BODY>
<% 
		Dim Action,ModeID,ID
		ModeID = ChkNumeric(Request("ModeID"))
		if ModeID=0 or ModeID="" Then ModeID=1
			If Not ACTCMS.ACTCMS_QXYZ(ModeID,"","") Then   Call Actcms.Alert("对不起，您没有"&ACTCMS.ACT_C(ModeID,1)&"系该项操作权限！","")

		ID=Trim(Request("ID"))
		Action = Request("Action")
		ConnectionDatabase
		Select Case Action
			Case "saveset"
				Call saveset
			Case Else
				Call MainArticle()
		End Select
	function Classmake(Cname)
		 Dim FolderRS
		 Set FolderRS = Conn.Execute("Select * from Class_act where ParentID='0'  and actlink=1 Order by Orderid desc,ID desc")
		 IF FolderRS.Bof And FolderRS.Eof Then
		 Response.Write("还没有添加任何栏目!")
		 End IF
		 do while Not FolderRS.Eof
			Classmake=Classmake&"<option value="&FolderRS("ClassID")&" "&Cname&"="&FolderRS("ClassID")&">"& FolderRS("ClassName") & "</option>"
			 Classmake=Classmake&(GetChildClassList(FolderRS("ClassID"),"",Cname))
		  FolderRS.MoveNext
		  loop
	 end function
	 Function GetChildClassList(ClassID,Str,Cname)
	       Dim Sql,RsTempObj,TempImageStr,ImageStr,CheckStr
	        TempImageStr = "&nbsp;└"
	        Sql = "Select * from Class_act where ParentID='" & ClassID & "'  and actlink=1"
	        ImageStr = Str & "&nbsp;└"
	        Set RsTempObj = Conn.Execute(Sql)
	            do while Not RsTempObj.Eof
					   GetChildClassList = GetChildClassList  & "<option value="&RsTempObj("ClassID")&" "&Cname&"="&RsTempObj("ClassID")&">"& ImageStr & TempImageStr &" "& RsTempObj("ClassName")& "</option>"
					  GetChildClassList = GetChildClassList & GetChildClassList(RsTempObj("ClassID"),ImageStr,Cname)
		         RsTempObj.MoveNext
	           loop
	       Set RsTempObj = Nothing
	 End Function 
	 function saveset()
	 dim id,idarr,rs,k
		id=request("id")
		     If request("choose")=0 Then
		      IDArr=Split(ID,",")
			 Else
			  IDArr=Split(Replace(Request.form("ClassID")," ",""),",")
			 End If
		      Set RS=Server.CreateObject("ADODB.RECORDSET")
			  For K=0 To Ubound(IDArr)
 				  If Request.form("choose")=0 Then
				   ModeID = ChkNumeric(Request("ModeID"))

				  RS.Open "Select * From "&ACTCMS.ACT_C(ModeID,2)&"  Where ID=" & IDArr(K), conn, 1, 3
				  Else
 				  RS.Open "Select * From "&ACTCMS.ACT_C(actcms.act_l(IDArr(K),10),2)&"  Where classid='" & IDArr(K) & "'", conn, 1, 3
				  End If
				    response.write "Select * From "&ACTCMS.ACT_C(ModeID,2)&"  Where ID=" & IDArr(K)
			  If Not RS.EOF Then
			     Do While Not RS.Eof
				  If ChkNumeric(Request.form("ACT_TemplateUrl"))=1 And Request.form("TemplateUrl")<>"" Then RS("TemplateUrl") = Request.form("TemplateUrl")
				  If ChkNumeric(Request.form("ACT_KeyWords"))=1 Then    RS("KeyWords")   = Request.form("KeyWords")
 				  If ChkNumeric(Request.form("ACT_CopyFrom"))=1 Then   RS("CopyFrom")  = Request.form("CopyFrom")
				  If ChkNumeric(Request.form("ACT_ATT"))=1 Then       RS("ATT")      = Request.form("ATT")
				  If ChkNumeric(Request.form("ACT_Hits"))=1 Then       RS("Hits")      = ChkNumeric(Request.form("Hits"))
				  If ChkNumeric(Request.form("ACT_rev"))=1 Then        RS("rev")       = ChkNumeric(Request.form("rev"))
				  If ChkNumeric(Request.form("ACT_IsTop"))=1 Then      RS("IsTop")     = ChkNumeric(Request.form("IsTop"))
				  If ChkNumeric(Request.form("ACT_Slide"))=1 Then      RS("Slide")     = ChkNumeric(Request.form("Slide"))
				  If ChkNumeric(Request.form("zy"))=1 Then			   RS("ClassID")     = Request.form("CID")
				   RS.Update
				 RS.MoveNext
				Loop
			 End If
			  RS.Close
			 Next 
		   Set RS = Nothing
		   conn.Close:Set conn = Nothing
		  Call Actcms.ActErr("批量设置成功","","")
 			response.end
		 End Function
 Function  MainArticle()%>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td  class="bg_tr"><strong>您现在的位置：文章中心管理 &gt;&gt; 文章属性设置</strong></td>
  </tr>
</table>
<TABLE class="table" style="MARGIN-TOP: 10px" cellSpacing=1 width="98%" align=center border=0>
  <FORM name=myform action=?Action=saveset&ModeID=<%=ModeID%> method=post>
  <TBODY>
  <TR class=tdbg id=choose2 <%if ID<>"" then response.write " style='display:none'"%>>
    <TD vAlign=top width=200 rowSpan=101>
	<FONT color=red>提示：</FONT>可以按住“Shift”<BR>或“Ctrl”键进行多个栏目的选择<BR>
	<SELECT style="WIDTH: 200px; HEIGHT: 380px" multiple size=2 name=ClassID> 
 <% 	 Response.Write Classmake("ClassID")
		%>
	</SELECT> 
      <DIV align=center>
	  <INPUT class="ACT_btn" onclick=SelectAll() type=button value=选定所有栏目 name=Submit><BR>
	  <INPUT class="ACT_btn" onclick=UnSelectAll() type=button value=取消选定栏目 name=Submit>
	  </DIV></TD></TR>
  <TR class=tdbg>
    <TD  align=right colSpan=2><STRONG>设置选择:</STRONG></TD>
    <TD>
	<INPUT <%if ID<>"" then response.write" checked"%> onClick="choose1.style.display='';choose2.style.display='none';" type=radio value=0 name=choose> 按文章ID&nbsp;&nbsp; 
	<INPUT <%if ID="" then response.write " checked"%> onClick="choose2.style.display='';choose1.style.display='none';" type=radio value=1 name=choose> 按文章分类</TD></TR>
  <TR class=tdbg id=choose1 <%if ID="" then response.write " style='display:none'"%>>
    <TD  align=right colSpan=2><STRONG>文章ID：</STRONG>多个ID请用“,”分开</TD>
    <TD><INPUT size=50 name=ID value="<%=ID%>"></TD></TR>
  <TR class=tdbg>
    <TD  align=middle height=25><INPUT type=checkbox value=1 name=ACT_TemplateUrl></TD>
    <TD width="150"  align=center><STRONG>文章模板:</STRONG></TD>
    <TD><input name="TemplateUrl" type="text" id="TemplateUrl">
       <input class="ACT_btn" type="button"  onClick="OpenWindowAndSetValue('../include/print/SelectPic.asp?CurrPath=<%=ACTCMS.ActSys%><%=actcms.SysThemePath&"/"&actcms.NowTheme%>',500,320,window,document.myform.TemplateUrl);" value="选择模板..."> 
	
	</TD>
  </TR>
  <TR class=tdbg>
    <TD  align=middle height=25><INPUT type=checkbox value=1 name=ACT_KeyWords></TD>
    <TD width="150"  align=center><STRONG>关 键 字:</STRONG></TD>
    <TD><INPUT  size=40 name=KeyWords>&nbsp;	</TD></TR>
 
  <TR class=tdbg>
    <TD  align=middle height=25><INPUT type=checkbox value=1 name=ACT_CopyFrom></TD>
    <TD width="150"  align=center><STRONG>文章来源:</STRONG></TD>
    <TD noWrap><input name="CopyFrom" type="text" id="CopyFrom" value="本站原创" />
        【<font color="blue"><font style="CURSOR: hand" 
            onclick="document.myform.CopyFrom.value='本站原创'" >本站原创</font></font>】 
        【<font color="blue"><font  style="CURSOR: hand" onclick="document.myform.CopyFrom.value='不详'">不详</font></font>】
        【<font color="blue"><font   style="CURSOR: hand" onclick="document.myform.CopyFrom.value='互联网'" >互联网</font></font>】
        【<font color="red"><font   style="CURSOR: hand" onclick="document.myform.CopyFrom.value=''">清空</font></font>】		</TD></TR>
  <TR class=tdbg>
    <TD  align=middle height=25><INPUT type=checkbox value=1 name=zy></TD>
    <TD  align=center><strong>转移栏目:</strong></TD>
    <TD>
	<select name="CID">
<% 	 Response.Write Classmake("CID")
		%>
		</select></TD>
  </TR>
  <TR class=tdbg>
    <TD  align=middle height=25><INPUT type=checkbox value=1  name=ACT_ATT></TD>
    <TD width="150"  align=center><STRONG>自定义属性:</STRONG></TD>
    <TD><select name="ATT" id="ATT">
					<option value="0">普通</option>
			<%=ACTCMS.ACT_ATT(0)%></select> </TD></TR>
  <TR class=tdbg>
    <TD  align=middle height=25><INPUT type=checkbox value=1  name=ACT_hits></TD>
    <TD width="150"  align=center><STRONG>点 击 数:</STRONG></TD>
    <TD><INPUT size=5 value=0 name=hits></TD></TR>
  <TR class=tdbg>
    <TD  align=middle height=25><INPUT type=checkbox value=1  name=ACT_Rec></TD>
    <TD width="150"  align=center><STRONG>是否推荐:</STRONG></TD>
    <TD><INPUT id=Rec type=radio value=1 name=Rec> 是
	<INPUT id=Rec type=radio CHECKED value=0 name=Rec> 否</TD></TR>
  <TR class=tdbg>
    <TD  align=middle height=25><INPUT type=checkbox value=1  name=ACT_IsTop></TD>
    <TD width="150"  align=center><STRONG>是否置顶:</STRONG></TD>
    <TD><INPUT type=radio value=1 name=IsTop> 是 <INPUT type=radio CHECKED value=0 name=IsTop> 否</TD></TR>
 
  <TR class=tdbg>
    <TD  align=middle height=25><INPUT type=checkbox value=1 name=ACT_rev></TD>
    <TD width="150"  align=center><STRONG>允许评论:</STRONG></TD>
    <TD><INPUT type=radio value=1 name=rev> 是 <INPUT type=radio CHECKED  value=0 name=rev> 否</TD></TR>
  <TR class=tdbg>
    <TD  align=middle height=25><INPUT type=checkbox value=1  name=ACT_Slide></TD>
    <TD width="150"  align=center><STRONG>是否幻灯:</STRONG></TD>
    <TD><INPUT type=radio value=1 name=Slide> 是 <INPUT type=radio CHECKED value=0 name=Slide> 
      否<B> 说明：</B>若要批量修改某个属性的值，请先选中其左侧的复选框，然后再设定属性值。</TD>
  </TR>
  <TR class=tdbg>
    <TD colSpan=3 height=30><INPUT class="ACT_btn" type=submit value="  保存  " name=button1>
&nbsp;
<INPUT class="ACT_btn" type=reset  value="  重置  " name=button2>    </TD></TR></FORM>
</TABLE>
<% End  Function %>
<script language="JavaScript" type="text/javascript">
		function OpenWindowAndSetValue(Url,Width,Height,WindowObj,SetObj)
		{
			var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:yes;help:no;scroll:yes;');
			if (ReturnStr!='') SetObj.value=ReturnStr;
			return ReturnStr;
		}	
		function SelectClass()
		{
			var ReturnValue='',TempArray=new Array();
			ReturnValue = OpenWindow('ACT.ClassID.asp',400,300,window);
			if (ReturnValue.indexOf('***')!=-1)
			{
				TempArray = ReturnValue.split('***');
				document.all.ClassIDs.value=TempArray[0]
				document.all.ClassNames.value=TempArray[1]
			}
		}		
		function OpenWindow(Url,Width,Height,WindowObj)
		{
			var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;status:no;help:no;scroll:no;');
			return ReturnStr;
		}	
</script>
</BODY></HTML>
