function OpenMWin(url,w,h)
{
    var theDes = "status:no;center:yes;help:no;minimize:no;maximize:no;dialogWidth:"+w+"px;scroll:no;dialogHeight:"+h+"px;border:think";
    return self.showModalDialog(url,null,theDes);
}



function Selector(type, val, ModeID, all)
{
	val = val?val:'';
	ModeID = ModeID?ModeID:'';
	all = all?all:'';
	return OpenMWin("ACT.D.asp?r=" + Math.random() + "&Type=" + type + "&IdList=" + val + "&ModeID=" + ModeID + "&ShowAll=" + all, 450, 350);
}

function GetCheckBoxList(objName)
{
	var result = "";
	var coll=document.all.item(objName)
	if(!coll) return result;
	if(coll.length){
		for(var i=0;i<coll.length;i++)
		{
			if(coll.item(i).checked)
			{
				result += (result == "")?coll.item(i).value:("," + coll.item(i).value);
			}
		}
	}else{
		if(document.all.item(objName).checked)
		{
			result = document.all.item(objName).value;
		}
	}
	return result;
}

function GetRadioBox(objName)
{
    var Coll = document.all.item(objName);
	if(!Coll) return null;
	if(Coll.length)
	{
		for(var i=0;i<Coll.length;i++)
		{
			if(Coll.item(i).checked)
			{
				return Coll.item(i).value;
			}
		}
		return null;
	}else{
		return Coll.checked?Coll.value:null;
	}
}
