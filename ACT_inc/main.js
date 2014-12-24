function CheckSel(Voption,Value)
{
 	var obj = document.getElementById(Voption);
	for (i=0;i<obj.length;i++){
		if (obj.options[i].value==Value){
		obj.options[i].selected=true;
		break;
		}
	}
}
 
 function upload(path,ModeID,iname)  
{
  J('#'+iname).dialog({ id:'actcmscj' ,title:'ACTCMS生成', loadingText:'上传加载中...', page: path+ 'Upload_Admin.asp?A=add&ModeID='+ModeID+ "&instrname="+iname+ "&" + Math.random(),  width:720, height:240 });
 
 }

  
 function uploadform(path,iname,ids)  
{
  J('#'+ids).dialog({ id:'actcmscj' ,title:'ACTCMS生成', loadingText:'上传加载中...', page: path+ 'Upload_Admin.asp?A=add&ModeID=999'+ "&instrname="+iname+ "&" + Math.random(),  width:720, height:240 });
 
 }


 function uploadimg(iname,ModeID) 
{
   J('#'+iname+'s').dialog({ id:'actcmscj'+iname ,title:'ACTCMS上传', loadingText:'上传加载中...', page: 'user.img.asp?A=add&ModeID='+ModeID+ "&instrname="+iname+ "&" + Math.random(),  width:720, height:240 });
   }
 

function sapLoadMsg(t){
var actup=t.split('|');
 {
  	   KE.insertHtml(actup[0], actup[1]);
}
}

function insertHTMLToEditor(I,codeStr)
{

KE.insertHtml(codeStr, I);

  
}