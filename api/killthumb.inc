Function killthumb(thumbpic)
	If InStr(thumbpic,"thumb_")>0 then 
		killthumb=rep(thumbpic,"thumb_","")	
	else 
		killthumb=thumbpic	
	end if 
 	 actcool=killthumb
end Function 