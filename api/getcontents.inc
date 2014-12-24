 '调用指定表的指定字段和指定条件的内容
 Function getcontents(table,Field,where) '表,字段,条件
	Dim grs,wherestr
	If where<>"" Then wherestr=" where "&where
	Set grs=actcms.actexe("select top 1  "&Field&"  from "&table&"  "&wherestr&" ")
	If Not grs.eof Then 
			getcontents=grs(""&Field&"")
	End If 
   grs.close:set grs=Nothing
   actcool=getcontents
 End Function 
 