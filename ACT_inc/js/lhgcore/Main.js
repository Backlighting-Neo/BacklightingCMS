function dohelp( h, w, y )
{
     w = w || 700; y = y || 500;
	J('#'+h).dialog({ id:h ,title:'在线帮助', loadingText:'帮助加载中...<br>逆光软件工作室制作', page:'http://help.actcms.com/act.asp?a='+h, link:true, width:700, height:500 });
}
function insertHTML(codeStr)
{
document.getElementById('DiyContent').value+=codeStr;
}