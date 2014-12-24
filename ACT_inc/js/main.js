// JavaScript Document
var EditorType="editAreaLoader";
$(document).ready(function()	{
	
	 var editor=['DiyContent','outerfor','ForClassContent','LabelContent','content'];
	
	
		function DoEditor(){
		
		if(typeof(editAreaLoader)!="undefined"){
	     
		$(editor).each(function(index, element) {			
           if($("#"+this)[0]){
			editAreaLoader.init({
			id: this	
			,start_highlight: true	
			,allow_resize: "both"
			,allow_toggle: true
			,word_wrap: true
			,syntax: "html"	
		});   
		   }
        });
  
       }else{
		    EditorType="markItUp";
		   $(editor).each(function(index, element) {	
		   if($("#"+this)[0]){
			   $("#"+this).markItUp();
		   }});
	   }
	}
	
	
	DoEditor();
   
});

function SetDiyContent(oTextarea,strText){ 
if(EditorType!="editAreaLoader")  
   $.markItUp({target:'#'+oTextarea, placeHolder:strText});
   else
   editAreaLoader.setSelectedText(oTextarea, strText);
}   

function CheckAll(obj)  
  {  
    //$(obj).parentsUntil("form").find("input:checkbox[checked=true]").attr("checked",$(obj).attr("checked"));
   //之所以这样写（click）是为了 给后面某个操作提供服务支持,之所以加了个c属性是考虑到jquery的click中在判断checkbox属性的时候和dom的click有问题
     if($(obj).attr("checked")){$(obj).parentsUntil("form").find("input:checkbox[checked=false]").attr("c",'false').click();}else{
		 $(obj).parentsUntil("form").find("input:checkbox[checked=true]").attr("c",'true').click();
	 }
  }