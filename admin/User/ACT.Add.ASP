<!--#include file="../ACT.Function.asp"-->
<!--#include file="../../act_inc/md5.asp"-->
<!--#include file="ACT.AddUM.ASP"-->
<!--#include file="../../Field.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>会员修改/添加 By Act</title>
<link href="../Images/editorstyle.css" rel="stylesheet" type="text/css">
<script charset="utf-8"  language="JavaScript" type="text/javascript" src="../../editor/kindeditor/kindeditor.js" ></script>
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/lhgcore.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/lhgdialog.min.js"></script>
<script language="JavaScript" type="text/javascript" src="../../ACT_inc/js/lhgcore/Main.js"></script>
<script type='text/javascript' src='../../ACT_INC/js/time/WdatePicker.js'></script>
  <script type="text/javascript" src="../../ACT_INC/js/swfobject.js"></script>
 <script type="text/javascript">
 var U="<%=ACTCMS.strToAsc(RSQL(Trim(Request.Cookies(AcTCMSN)("AdminName"))))%>";
var P="<%=ACTCMS.strToAsc(RSQL(Trim(Request.Cookies(AcTCMSN)("AdminPassword"))))%>";
</script>
 </head>
 <body>
<% 

  If Not ACTCMS.ChkAdmin() Then   Call Actcms.Alert("对不起，你没有操作权限！","")
	dim sqlUser,rsUser,UserName,Province,city,locked,realName,Sex,Email,IF_NULL,i,rs,ChargeType
	dim Birthday,Mobile,HomeTel,QQ,MSN,address,postcode,myface,GroupID,Score,A,Password,RPassWord,BeginDate,EDays,Point,Money
	dim Action,UModeID,founderr
	Action = Request("Action")
	UModeID=ChkNumeric(Request("UModeID"))
	If UModeID=0 Then UModeID=1
  	Select Case Action
			Case "Edit","Add"
 				Call AddEdit()
			Case "AddSave"
				Call AddSave
			Case "ESave"
				Call ESave
			Case "del"
				Call Del()
 			Case Else
				Call AddEdit
	End Select
 %>
 
<%    Function ActUser_MX_Arr(ModeID)'返回模型数组
	  Dim Rs
	  Set Rs=actcms.ACTEXE("Select FieldName,Title,IsNotNull,FieldType,[check],regex,regError from Table_ACT  Where ModeID=" & ModeID & " and SearchIF=1 and actcms=2  order by OrderID desc,ID Desc")
	 If Not Rs.Eof Then
	  ActUser_MX_Arr=Rs.GetRows(-1)
	 Else
	  ActUser_MX_Arr=""
	 End If
	 Rs.Close:Set Rs=Nothing
   End Function
	sub AddSave()
	
			UModeID=ChkNumeric(request("UModeID"))
 			if UModeID=0 then
				FoundErr=True
				Call ACTCMS.Alert("参数不足！","")
				exit sub
			end If
			Set rs=actcms.actexe("select UModeID from user_act where UModeID="&UModeID)
			If Not rs.eof Then UModeID=rs("UModeID")
			Password  = Trim(request("Password"))
			RPassWord= Trim(request("RPassWord"))
  			Sex       = Trim(Request("Sex"))
			Email     = Trim(request("Email"))
 			QQ        = Trim(request("QQ"))
			MSN       = Trim(request("MSN"))
			GroupID   = Trim(request("GroupID"))
			locked    = Trim(request("locked"))
 			ChargeType=ChkNumeric(Trim(request("ChargeType")))
			
			BeginDate=Trim(request("BeginDate"))
			EDays=ChkNumeric(Trim(request("EDays")))
			Point=ChkNumeric(Trim(request("Point")))
			Money=ChkNumeric(Trim(request("Money")))
			
  			province= Request("prov")
			city=     Request("city")
			UserName=RSQL(Trim(request("UserName")))
			RealName=Trim(request("RealName"))
			Address=Trim(request("Address"))
 			BirthDay=Trim(request("BirthDay"))
 			HomeTel=Trim(request("HomeTel"))
			Mobile=Trim(request("Mobile"))
			IF_NULL=ACTCMS.Act_MX_Arr(UModeID,2)
			If IsArray(IF_NULL) Then
			For I=0 To Ubound(IF_NULL,2)
			 If IF_NULL(2,I)=0 And Trim(ACTCMS.S(IF_NULL(0,I)))="" Then  Call  ACTCMS.ALERT(IF_NULL(1,I)&"不能为空","")
			Next
			End If
			Set rsUser=Server.CreateObject("Adodb.RecordSet")
			sqlUser="select * from User_ACT where  UserName='" & UserName & "'"
			rsUser.Open sqlUser,Conn,1,3
			
			If Not  rsUser.bof then
				FoundErr=True
				Call ACTCMS.Alert("数据库中已存在该用户名","")
				rsUser.close:set rsUser=Nothing
				exit sub
			end if
		    rsUser.addnew
			if Password<>RPassWord then
				founderr=true
				Call ACTCMS.Alert("密码和确认密码不一致","")
			end if
			if Sex="" then
				founderr=true
				Call ACTCMS.Alert("性别不能为空","")
			end if
			if Email<>"" and  ACTCMS.IsValidEmail(Email)=false then
					Call ACTCMS.Alert("您的Email有错误","")
					founderr=true
			end if
			
		
			if GroupID="" then
				FoundErr=True
				Call ACTCMS.Alert("请指定用户级别！","")
			else
				GroupID=CLng(GroupID)
			end if
			if locked<>0 then locked=1
 			
 				if   IsDate(BeginDate)=false then
 				BeginDate=now
				end if
 			
 				if  IsDate(BirthDay)=false then
 				BirthDay=now
				end if
 			
			
			if founderr=true then exit sub
		   rsUser("UserName") = UserName
			rsUser("RealName")=RealName
			if Password<>"" then
				rsUser("Password")=md5(Password)
			end if
  			rsUser("BeginDate")=BeginDate
 			rsUser("EDays")=EDays
 			rsUser("Point")=Point
 			rsUser("Money")=Money
 			rsUser("Email")=Email
			rsUser("Sex")=Sex
			rsUser("GroupID")=GroupID
		   rsUser("UModeID") = UModeID
			rsUser("locked")=locked
 			rsUser("Birthday")=Birthday
			rsUser("province")=province
			rsUser("ChargeType")=ChargeType
			rsUser("City")=City
 			rsUser.update
 
			
						  '自定义字段开始入库
			Dim UserID,FieldRS,FieldSql
			Set rs=ACTCMS.actexe("Select top 1 UserID from User_ACT  order by UserID desc")
			If Not rs.eof Then UserID = rs("UserID")
			
 				  Set FieldRS = Server.CreateObject("adodb.recordset")
				  FieldSql = "select * from "&ACTCMS.ACT_U(UModeID,2)&" where userid<>"&userid&""
 				  FieldRS.Open FieldSql, Conn, 1, 3
				  FieldRS.AddNew
				  FieldRS("UserID") = UserID
  				 Dim IF_NULL
				IF_NULL=ActUser_MX_Arr(UModeID)
				If IsArray(IF_NULL) Then
				For I=0 To Ubound(IF_NULL,2)
				 If IF_NULL(2,I)=0 And Trim(ACTCMS.S(IF_NULL(0,I)))="" Then  Call  ACTCMS.ALERT(IF_NULL(1,I)&"不能为空","")
				Next
				End If
				If IsArray(IF_NULL) Then
					For I=0 To Ubound(IF_NULL,2)
 						If IF_NULL(3,I)="NumberType" Then 
						   If actcms.regexField(ACTCMS.S(IF_NULL(0,I)),"^\d+$")=True Then 
							   FieldRS("" & IF_NULL(0,I) & "" )= ACTCMS.S(IF_NULL(0,I))
						   End If 
						ElseIf IF_NULL(3,I)="DateType" Then 
							If IsDate(ACTCMS.S(IF_NULL(0,I)))=False Then 
								FieldRS("" & IF_NULL(0,I) & "")= Now()
							Else 
								FieldRS("" & IF_NULL(0,I) & "")=ACTCMS.S(IF_NULL(0,I))
							End If
						ElseIf IF_NULL(4,I)="1" Then 
								 FieldRS("" & IF_NULL(0,I) & "")= actcms.AField(IF_NULL(5,I))
						ElseIf IF_NULL(4,I)="2" Then 
								If actcms.regexField(ACTCMS.S(IF_NULL(0,I)),IF_NULL(5,I))=True Then 
									FieldRS("" & IF_NULL(0,I) & "")=ACTCMS.S(IF_NULL(0,I))
								Else 
									Call Actcms.Alert(IF_NULL(6,I),"")
								End If 
						Else 
							FieldRS("" & IF_NULL(0,I) & "")=ACTCMS.S(IF_NULL(0,I))
						End If 
						actField=""
					Next
  					FieldRS.Update	
				End If
			
			
			Set  rsUser=Nothing
 			Call ACTCMS.Alert("恭喜您，添加成功！请按确定返回！","User_Admin.asp?ModeID="&UModeID&"")
 	end sub 
  	Sub AddEdit() 
	 Sex=0
 %>
 
 <form  name="actcmsfrom"   method="post" action="?Action=AddSave&UModeID=<%= UModeID %>">
 <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="table">
  <tr>
    <td colspan="4" align="left" class="bg_tr">您现在的位置：会员管理 &gt;&gt; 修改会员</td>
  </tr>
   <tr>
    <td width="15%" height="26" align="right"  class="tdclass">用户名称：</td>
    <td width="68%" height="26"align="left"  class="tdclass"> 
    
    <input name="UserName" type="text"  class="Ainput"   value="<%= UserName %>" size="30" />    </td>
<td width="17%" rowspan="8"  class="tdclass"><img src="<%If Trim(myface)<>"" Then 
		response.write myface
	Else 
		response.write "../../user/images/nophoto.gif" 
	End If 
	
	%> " alt="用户头像" height="150" width="150"/></td>
   </tr>
  <tr>
    <td height="26" align="right"  class="tdclass">用户状态：</td>
    <td  height="26"align="left" bgcolor="#FFFDEC"  class="tdclass">
    <label for="locked1"><input type="radio" name="locked" id="locked1" value="0" <%if locked=0 then Response.Write "checked"%>>
        <font color="green">正常</font></label>&nbsp;&nbsp;
       <label for="locked2"> <input type="radio" name="locked" id="locked2" value="1" <%if locked=1 then Response.Write "checked"%>>
        <font color="red">锁定</font></label>     </td>
    </tr>
  <tr>
    <td height="26" align="right"  class="tdclass">计费方式：</td>
    <td  height="26"   align="left" bgcolor="#FFFDEC"  class="tdclass">
    
 <label for="ChargeType1"><input type="radio" name="ChargeType" id="ChargeType1" value="1"  <%if ChargeType=1 then Response.Write "checked"%>>
        <font color="green">扣点数（推荐）</font></label>
       <label for="ChargeType2"> <input type="radio" name="ChargeType" id="ChargeType2" value="2" <%if ChargeType=2 then Response.Write "checked"%>>
        <font color="blue">有效期</font></label>      
      <label for="ChargeType3"> <input type="radio" name="ChargeType" id="ChargeType3" value="3" <%if ChargeType=3 then Response.Write "checked"%>>
        <font color="red">无限期</font></label>  
     </td>
  </tr>
  <tr>
    <td height="26" align="right"  class="tdclass">有效期限：</td>
    <td  height="26"   align="left" bgcolor="#FFFDEC"  class="tdclass">
    开始日期：<input name="BeginDate" type="text"  class="Ainput"   value="<%= BeginDate %>"  onClick="WdatePicker()"  size="20" />
    
若超过此期限，则用户不能阅读收费内容此功能只有当计费方式为“有效期限”时才有效 <br>
有 效 期：<input name="EDays" type="text"  class="Ainput"   value="<%= EDays %>"   size="20" />天</td>
  </tr>
  
  <tr>
    <td height="26" align="right"  class="tdclass">用户<%=actcms.ActCMS_Sys(24)%>：</td>
    <td  height="26"   align="left" bgcolor="#FFFDEC" class="tdclass"><input name="Point" type="text"  class="Ainput"   value="<%= Point %>" size="30" />
      用户资金：<input name="Money" type="text"  class="Ainput"   value="<%= Money %>" size="30" /></td>
  </tr>  
  
   <tr>
    <td height="26" align="right"  class="tdclass">真实姓名：</td>
    <td  height="26"   align="left" bgcolor="#FFFDEC"  class="tdclass">
    <input name="realName" type="text"  class="Ainput"   value="<%= realName %>" size="30" /></td>
    </tr>
   
   <tr>
    <td height="26" align="right"  class="tdclass">用户密码： </td>
    <td height="26" align="left"  class="tdclass"><input name="PassWord" type="text"  class="Ainput"  value="" size="30" />
    <font color="#FF6600">如果不想修改，请留空</font></td>
    </tr>
  <tr>
    <td height="26" align="right"  class="tdclass">重复密码：</td>
    <td height="26" align="left"  class="tdclass"><input name="RPassWord" type="text"  class="Ainput"  value="" size="30" />
    <font color="#FF6600">如果不想修改，请留空</font></td>
    </tr>
  <tr>
    <td height="26" align="right"  class="tdclass">用户性别：</td>
    <td height="26" colspan="2" align="left"  class="tdclass">
    <label for="sex1"><INPUT type=radio value="0" name=sex id="sex1"  <%if Sex="0" then Response.Write "CHECKED"%>>
        男</label> &nbsp;&nbsp;
        <label for="sex2"><INPUT type=radio value="1" name=sex id="sex2" <%if Sex="1" then Response.Write "CHECKED"%>>
        女</label></td>
  </tr>
  <tr>
    <td height="26" align="right"  class="tdclass">用户等级：</td>
    <td height="26" colspan="2" align="left"  class="tdclass">
	<select name="GroupID" id="GroupID">
 	<%=GroupOption(GroupID,UModeID)%>
     </select>    </td>
  </tr>
  <tr>
    <td height="26" align="right"  class="tdclass">用户邮箱：</td>
    <td height="26" colspan="2" align="left"  class="tdclass"><input name="Email" type="text"  class="Ainput"  value="<%= Email %>" size="30" /></td>
  </tr>
  <tr>
    <td height="26" align="right"  class="tdclass">所在省市：</td>
    <td height="26" colspan="2" align="left"  class="tdclass">
<select name="prov" onChange="selectcityarea('prov','city','actcmsfrom');" style="width:110">
        <%if Province<>"" then%>
        <option  value="<%=Province%>" selected="selected"><%=Province%></option>
        <%else%>
        <option  value="" selected="selected">==请选择省份==</option>
        <%end if%>
      </select>
          <select name="city" style="width:110">
            <%if  city<>"" then%>
            <option  value="<%=city%>" selected="selected"><%=city%></option>
            <%else%>
            <option  value="" selected="selected">==请选择城市==</option>
            <%end if%>
          </select>
          <script language="javascript" src="<%=ACTCMS.ActSys%>act_Inc/Province.js"></script>    </td>
  </tr>

   
 <tr>
    	<td align="right"  class="tdclass">出生日期：</td>
        <td colspan="2"  class="tdclass">
        <input name="Birthday" type="text" class="Ainput"  id="Birthday" size="30"  onClick="WdatePicker()"  value="<%=Birthday%>">        </td>
    </tr>
        <tr>
    	<td align="right"  class="tdclass">手机：</td>
    	<td colspan="2"  class="tdclass"><input name="Mobile"  class="Ainput" type="text" id="Mobile" size="30"  value="<%=Mobile%>">         </td>
    </tr>
        <tr>
    	<td align="right"  class="tdclass">电话：</td>
    	<td colspan="2"  class="tdclass"><input name="HomeTel"  class="Ainput"  type="text" id="HomeTel" size="30"  value="<%=HomeTel%>">   </td>
    </tr>
        <tr>
    	<td align="right"  class="tdclass">QQ：</td>
    	<td colspan="2"  class="tdclass"><input name="QQ" class="Ainput"  type="text" id="QQ" size="30"  value="<%=QQ%>">  </td>
    </tr>
        <tr>
    	<td align="right"  class="tdclass">MSN：</td>
    	<td colspan="2"  class="tdclass"><input name="MSN" class="Ainput"  type="text" id="MSN" size="30"  value="<%=MSN%>">  </td>
    </tr>
        <tr>
    	<td align="right"  class="tdclass">地址：</td>
    	<td colspan="2"  class="tdclass"><input name="address"  class="Ainput"  type="text" id="address" size="30"  value="<%=address%>"> </td>
    </tr>
    
    
    <tr>
    	<td align="right"  class="tdclass">邮编：</td>
    	<td colspan="2"  class="tdclass"><input name="postcode" class="Ainput"  type="text" id="postcode" size="30"  value="<%=Score%>">  </td>
    </tr>
   
     
  <tr>
    <td height="26" align="right"  class="tdclass">会员点数：</td>
    <td height="26" colspan="2" align="left"  class="tdclass"><input name="Score" type="text"  class="Ainput"  value="<%= Score %>" size="30"    /></td>
  </tr>
   
   
  <tr>
    <td height="23" colspan="4" align="center" class="bg_tr">扩展信息</td>
    </tr>
     <%= U_M.ACT_NoRormMXList(UModeID) %>
 

  <tr>
    <td height="23" colspan="4" align="center"   class="tdclass">
	<input type=button onclick=CheckForm() class="ACT_btn"  name=Submit1 value="  保存  " />
      &nbsp;&nbsp; <input type="reset" class="ACT_btn" name="Submit2" value="  重置  ">         </td>
    </tr>
</table>
</form> 
<% end sub

	Public Function GroupOption(Selected,UModeID)
	 Dim RSObj,GroupName:Set RSObj=Server.CreateObject("Adodb.Recordset")
	    RSObj.Open "Select GroupID,GroupSetting,GroupName From Group_ACT where modeid="&UModeID,Conn,1,1
	  	Do While Not RSObj.Eof
		   GroupName=RSObj(2)
		   IF Selected=RSObj(0) Then
			GroupOption=GroupOption & "<option value=""" & RSObj(0) & """ Selected>" &GroupName & "</option>"
		   Else
			GroupOption=GroupOption & "<option value=""" & RSObj(0) & """>" & GroupName & "</option>"
		   End If
		RSObj.MoveNext
		Loop
	  RSObj.Close:Set RSObj=Nothing
	End Function

 %>

<SCRIPT language=javascript>
<!--
function CheckForm()
{ 
var form=document.actcmsfrom;
   if (form.UserName.value=='')
    { Alert("请输入用户名!");   
	  form.UserName.focus();    
	   return false;
    }
		  
	 if (form.PassWord.value!=form.RPassWord.value)   
	   {    
	     Alert("两次输入的密码不一致!");   
		  form.PassWord.focus();   
		  return false;    
		  }
		form.Submit1.value="正在提交数据,请稍等...";
		form.Submit1.disabled=true;	
	    form.submit();
        return true;
}

//-->
 
function insertHTMLToEditor(codeStr,I)
{
  	var oEditor = CKEDITOR.instances[I];
 	if ( oEditor.mode == 'wysiwyg' )
	{
 		oEditor.insertHtml( codeStr );
	}
	else
		alert( '未定义' );
}
	
 function upload(instr,iname) 
{
  J.dialog.get({ id: 'zxsc', title: '在线上传',width: 720,height: '240', page: '../include/Upload_Admin.asp?A=add&instr='+instr+ "&ModeID="+999+ "&instrname="+iname+ "&" + Math.random() });
 }
 
 //-->

	</script>
  
</body>
</html>