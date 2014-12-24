<!--#include file="../../include.asp"-->
<!--#include file="Mode.asp"-->
<!--#include file="label.asp"-->
<%
	ThemeID=actcms.s("ThemeID")
	If actcms.s("Install")="yes" Then 
		Call GetLabel(ThemeID)
		actcms.actexe("Update Config_act Set ActCMS_Theme='"&ThemeID&"'  ")
		Call actcms.DelCahe("NowTheme")
 		 response.Redirect "Index.asp?title="&ThemeID&"主题安装成功"
		 respone.End
	Else 
		Call GetLabel(ThemeID)
 	End If 
	Function newm(ModeID)
		'2-产品 3-链接
 			Select Case ModeID
				Case "2"'原来老的模型ID
					newm=ACTCMS.ACTEXE("SELECT ModeID FROM Mode_Act Where ModeName='链接'")(0)
				Case "3"'原来老的模型ID
					newm=ACTCMS.ACTEXE("SELECT ModeID FROM Mode_Act Where ModeName='产品'")(0)
				Case Else 
					newm=modeid
			End Select  
 	End Function 
	' Call forModeContent
	'Call execute(actcms.LTemplate("class.asp"))
	Function forModeContent()
	Call 	drbc(actcms.LTemplate("2.act"))
	Call 	drbc(actcms.LTemplate("3.act"))
	End Function 
   Application.Contents.RemoveAll
   echo "缓存清除成功"

		   
		   %>