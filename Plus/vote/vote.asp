<!--#include file="../../act_inc/ACT.User.asp"-->
<%	
 	Dim id,i,VoteType,rs1,voted
	id=ChkNumeric(RSQL(request("id")))
	voted=ChkNumeric(RSQL(request("voted")))
 	 set rs1=actcms.actexe("select * from vote_act where id="&id&" and isLock=0")
	 If Not rs1.eof Then VoteType=rs1("VoteType")
if VoteType="0" then
	if isnull(request("voted")) or request("voted")=empty then
	response.write "<script>alert('请选择投票项目。');history.go(-1);</script>"
	response.end
	end If
	response.write request("voted")
	actcms.actexe("Update vote_act set VoteNum=VoteNum+1 where id="&voted)
elseif VoteType="1" then
	if request("voted").count=0 then
	response.write "<script>alert('请选择投票项目。');window.close()</script>"
	response.end
	end If
	Dim arrvoted
	arrvoted=Split(request("voted"),",")
  	For I = 0 To UBound(arrvoted)
		 actcms.actexe("Update vote_act set VoteNum=VoteNum+1 where id="&ChkNumeric(RSQL(arrvoted(i))))
	Next 
 end If
 response.Redirect "index.asp?id="&id&""

 %>