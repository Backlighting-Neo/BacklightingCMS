<!--#include file="../ACT.Function.asp"-->
<% 
If Not  ACTCMS.ChkAdmin() Then 
Response.Write "您没有权限执行此操作"
Response.End()
End If
'============添加管理员日志======================
ShortDescription="【数据库总管理】"
LongDescription=""
Call ACTCMS.InsertLog(Request.Cookies(AcTCMSN)("AdminName"),2,ShortDescription,LongDescription)
'============/添加管理员日志===================
response.Redirect("http://www.access2008.cn/DetectionExterior.asp?URL=http://zdh.qdu.edu.cn/accessapi.asp&PASS=qduzdh")
 %>