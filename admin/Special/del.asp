<!--#include file="../ACT.Function.asp"-->
 <% 
   Dim ID,rs
   ID=ChkNumeric(Request.QueryString("ID"))
   Set rs=actcms.actexe("select * from SpecialPicUrl_ACT where id="&ID&"")
   If rs.eof Then response.write "没有找到图片":response.end

   If  actcms.DeleteFile(rs("picurl"))=false Then response.write "删除图片文件失败":response.end
   Conn.Execute ("Delete from SpecialPicUrl_ACT Where ID=" &ID )		
   Set conn=nothing
   response.write "1"
%>
 