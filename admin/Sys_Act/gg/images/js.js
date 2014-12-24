function Num(){
  var berr=false;
  if (!(event.keyCode>=48 && event.keyCode<=57)) berr=true;
  return !berr;
}

function chkinput(){
if (document.form.ADID.value=="")
	{
	  alert("请填写广告ID (不能重名)");
	  document.form.ADID.focus();
	  return false
	 }
else if (document.form.ADSrc.value=="")
	{
	  alert("请填写广告链接地址 (对于窗口和网页对话框请输入打开页面地址)");
	  document.form.ADSrc.focus();
	  openhelp("ext")
	  return false
	 }
else if (isNaN(parseInt(document.form.ADWidth.value)))
	{
	  alert("请填写广告的宽 (单位：象素)");
	  document.form.ADWidth.focus();
	  return false
	 }
else if (isNaN(parseInt(document.form.ADHeight.value)))
	{
	  alert("请填写广告的高 (单位：象素)");
	  document.form.ADHeight.focus();
	  return false
	 }
else if (isNaN(parseInt(document.form.ADStopHits.value)))
	{
	  alert("请填写限制广告的点击次数 (如不限制请输入数字0)");
	  document.form.ADStopHits.value=0;
	  document.form.ADStopHits.focus();
	  return false
	 }
else if (isNaN(parseInt(document.form.ADStopViews.value)))
	{
	  alert("请填写限制广告的显示次数 (如不限制请输入数字0)");
	  document.form.ADStopViews.value=0;
	  document.form.ADStopViews.focus();
	  return false
	 }
else if (isNaN(parseInt(document.form.ADStopDate.value)))
	{
	  var d = new Date();
	  alert("请填写限制广告投放截止日期 (格式"+d.getYear()+"-"+(d.getMonth()+1)+"-"+d.getDate()+" 如不限制请输入如"+(d.getYear()+50)+"-"+(d.getMonth()+1)+"-"+d.getDate()+")");
	  document.form.ADStopDate.value=(d.getYear()+50)+"-"+(d.getMonth()+1)+"-"+d.getDate();
	  document.form.ADStopDate.focus();
	  return false
	 }
}

function ckinput(){
if (document.loadad.username.value=="")
	{
	  alert("请填管理员名称");
	  document.loadad.username.focus();
	  return false
	 }
else if (document.loadad.userpass.value=="")
	{
	  alert("请填写管理员密码");
	  document.loadad.userpass.focus();
	  return false
	 }
else if (document.loadad.secruity.value=="")
	{
	  alert("请填写随机密码 (右边的四位数字)");
	  document.loadad.secruity.focus();
	  return false
	 }
}

function ChangeType(adtype){
	//alert("你选择了广告："+adtype);
	var adsrcText="";
	switch (adtype){
	case "6":adsrcText="广告内容";
	case "5":
	case "7":
	case "8":
			if(adsrcText!=""){window.adsrc_text.innerHTML= '广告内容：';}
			else{window.adsrc_text.innerHTML= '广告地址：';}
			
			if(document.form.ADLink.value!='本项无效，无需填写'){
				ADLinkValue=document.form.ADLink.value;
				ADAltValue=document.form.ADAlt.value;

				document.form.ADLink.value='本项无效，无需填写';
				document.form.ADAlt.value='本项无效，无需填写';
			}
			document.form.ADLink.disabled=true;
			document.form.ADAlt.disabled=true;
			
			break;
	default:
			
			if(document.form.ADLink.value=='本项无效，无需填写'){document.form.ADLink.value=ADLinkValue;}
			if(document.form.ADAlt.value=='本项无效，无需填写'){document.form.ADAlt.value=ADAltValue;}
			document.form.ADLink.disabled=false;
			document.form.ADAlt.disabled=false;

			window.adsrc_text.innerHTML= '广告地址：';
			break;
	}
}

function openhelp(e){
var e
if (e=="ext"){
	alert("关于：广告地址\n\n 通常指图片或FLASH的地址 \n\n 以下类型广告除外：\n 1、【透明对话框】 指对话框中的内容 \n 2、【网页对话框】 网页的URL地址 \n 3、【窗口】 指打开或弹出窗口地址")
	}
if (e=="stop"){
	alert("关于：投放限制\n\n 显示广告的限制条件 显示数·点击数·显示日期 \n 当满足条件后将不在显示该广告(无提示) \n\n 请注意：\n 1、【不想限制】 不限制显示或点击数请填写数字0,日期请填写1月以后的日期 \n 2、【限制一个】 这三个限制可同时或单独填写,即对填写的起作用 \n 3、【日期限制格式】 2007-10-10")
	}
}