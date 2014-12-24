function Digg(num_id,ModeID,ActSys)
{
 	var url=lhgajax.send(ActSys+"plus/digg/digg.asp?id="+num_id+"&post=digg&ModeID="+ModeID+"&m="+Math.random(),"GET");
    var DigArr=url.split('|');
 	switch (DigArr[0])
	{
		 case "err":
			 alert('系统错误,未定义操作!');
			 break;
		 case "Close":
			 alert('系统已经关闭DIGG!');
			 break;
		 case "ACT":
			 alert('您已经顶过了!');
  		 break;
			 case "Login":
			 alert('您还没有登录，请登陆后操作!');
		 break;
				
			  default:
 			  document.getElementById("dact").innerHTML = "("+DigArr[0]+")";
 			//  document.getElementById("dbar").innerHTML = "<span style=width:"+DigArr[0]+"%></span>";
 			  document.getElementById("dnum").innerHTML = "("+DigArr[0]+"%)";
 	}
 }
 
function down(num_id,ModeID,ActSys)
{
 	var url=lhgajax.send(ActSys+"plus/digg/digg.asp?id="+num_id+"&post=down&ModeID="+ModeID+"&m="+Math.random(),"GET");
    var DigArr=url.split('|');
	switch (DigArr[0])
	{
		 case "err":
			 alert('系统错误,未定义操作!');
			 break;
		 case "Close":
			 alert('系统已经关闭DIGG!');
			 break;
		 case "ACT":
			 alert('您已经顶过了!');
  		 break;
			 case "Login":
			 alert('您还没有登录，请登陆后操作!');
		 break;
				
			  default:
 			  document.getElementById("downact").innerHTML = "("+DigArr[0]+")";
 			//  document.getElementById("downbar").innerHTML = "<span style=width:"+DigArr[0]+"%></span>";
 			  document.getElementById("downnum").innerHTML = "("+DigArr[0]+"%)";
 	}
 }
 


function postBadGood(ActSys,Type,id)
{	
   	var url=lhgajax.send(ActSys+"plus/Comment/ACT.Comment.asp?Action=Support&Type="+Type+"&id="+id+"&m="+Math.random(),"GET");
    var DigArr=url.split('|');
	switch (DigArr[0])
	{
		 case "err":
			 alert('您已经顶过了');
			 break;
 		 break;
 			  default:
		if (Type==1)
		{
 			  document.getElementById("goodfb"+id).innerHTML = "<a href=#goodfb"+id+" onclick=postBadGood('"+ActSys+"','1',"+id+")>已支持</a>["+DigArr[1]+"]";
		}
		else

		{
 			  document.getElementById("badfb"+id).innerHTML = "<a href=#goodfb"+id+" onclick=postBadGood('"+ActSys+"','2',"+id+")>已反对</a>["+DigArr[1]+"]";
		}
  	}
 }
 
function CheckForm()
{ var form=document.Comment;
    if (form.Content.value=='')
    { alert("请输入评论内容!");   
	  form.Content.focus();    
	   return false;
    }
 		form.Submit.value="正在提交数据,请稍等...";
		form.Submit.disabled=true;	
	    form.submit();
	    return true;
}