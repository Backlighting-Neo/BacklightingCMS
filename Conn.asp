<%@ LANGUAGE = VBScript CodePage = 65001%>
<%
Option Explicit
response.buffer=true
session.codepage=65001
response.charset="utf-8"
Dim Conn,db,MyDbPath,actcool,actField,NowString,ConnStr,aspexe
Const isSqlDataBase = 0
Const MsxmlVersion=".3.0"  '系统采用XML版本设置
Const AcTCMSN="BKCMS1212"'系统缓存名称.在一个URL下安装多个ACTCMS请设置不同名称
Const DataBaseType="access"  '' 数据库类型: 值分别为 access   mssql
NowString = "Now()"
MyDbPath ="/"'系统安装目录,如在虚拟目录下安装.请填写 /虚拟目录名称/
db = "data_act/#BackLightingDatabase.mdb" 'ACCESS数据库的文件名
Connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(MyDbPath & db)
Sub ConnectionDatabase()
	'On Error Resume Next
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Connstr
	
End Sub

Sub CloseConn()
	On Error Resume Next
	If IsObject(Conn) Then
		Conn.Close:Set Conn = Nothing
	End If
End Sub
%>
<!--#include file="ACT_INC/ACT.Common.asp" -->
