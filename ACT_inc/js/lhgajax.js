/*
 *@Generator -> LiHuiGang - Email:lhg133@126.com - QQ:463214570 file:lhgajax.js
 *@Copyright (c) 2009 LiHuiGang Compostion Blog:http://www.cnblogs.com/lhgstudio/
 */

var lhgajax = (function()
{
    var gethttp = function()
	{
		try{ return new ActiveXObject('Msxml2.XMLHTTP'); }catch(e){}
	    try{ return new XMLHttpRequest(); }catch(e){} return null;
	};
	
	return {
	    send : function( url, method, sync, pdata )
		{
		    var oh = gethttp(); url = url + '?t=' + new Date().getTime();
			sync = sync ? true : false; oh.open( method, url, sync );
			
			if( method.toLocaleUpperCase() == 'GET' ) oh.send(null);
			else
			{ 
			    oh.setRequestHeader('content-type','application/x-www-form-urlencoded');
				if(pdata) oh.send(pdata); else return false;
			}
			
			if( oh.readyState == 4 && oh.status == 200 )
			{
			    var redata = oh.responseText; delete(oh); return redata;
			}else return false;
		}
	};
})();