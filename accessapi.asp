<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'--------------------------------------------
'Access 数据库在线管理系统 API文件
'网址: http://www.access2008.cn
'--------------------------------------------
Response.Charset="utf-8"
Session.CodePage = "65001"
Response.Buffer = True 
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "No-Cache"
Response.ContentType = "text/xml"

const mulu=".\data_act\" '数据库所在目录
const APIPASS="qduzdh" '文件密码
const apiVersion="1.1.3" 'API版本
const apiVersionmun="113"

dim COArray:COArray = Array("Adodb.Connection","Adodb.RecordSet","Adox.CataLog","Adox.Table","Adox.Column","Adox.Index","Adox.Key","Msxml2.DOMDocument","JRO.JetEngine","Scripting.FileSystemObject")
dim cmd
dim text
dim filelj
dim mululj
dim accessPass:accessPass=""
dim accessData:accessData=""
dim comad:comad = Request("command")
dim APIFilePASS:APIFilePASS=request("APIFilePASS")
dim fs
set text = New TextData
If len(comad) > 0 Then
	if instr(mulu,":")=0 then
		mululj=server.MapPath(mulu)
	else
		mululj=mulu
	end if
	if right(mululj,1)="\" or right(mululj,1)="/" then
		mululj=left(mululj,len(mululj)-1)
	end if
	set fs = server.CreateObject(COArray(9))
	a=fs.FolderExists(mululj)
	set fs=nothing
	if a then
		if APIFilePASS=APIPASS or len(APIPASS)=0 then
			cmad(comad)
		else
			text.outerr text.gettxt(0)
		end if
	else
		text.outerr text.gettxt(1)
	end if
else
	text.outerr text.gettxt(2)
End If
function cmad(a)
	dim i,b,c
	dim a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14
	accessData=Trim(Request("access"))
	a1=Trim(Request("access"))
	a2=Trim(Request("table"))
	a3=Trim(Request("sl"))
	a4=Trim(Request("ys"))
	accessPass=Trim(Request("pass"))
	a5=Trim(Request("pass"))
	a6=Trim(Request("AbsolutePosition"))
	a7=Trim(Request("data"))
	a8=Trim(Request("bs"))
	a9=Trim(Request("field"))
	a10=Trim(Request("oldname"))
	a11=Trim(Request("newname"))
	a12=Trim(request("sql"))
	a13=Trim(Request("newpass"))
	a14=Trim(Request("type"))
	select case a
		case "index"
			call getaccess()
		case "gettable"
			b = split(a1,"|")
			if len(a5)>0 then
				c = split(a5,"|")
			else
				c= array(0)
				c(0)=""
			end if
			for i= 0 to ubound(b)
				call gettable(b(i),c(i))
			next
		case "getdatalist"
			call getdatalist(a1,a2,a3,a4,a5)
		case "deletedata"
			call deletedata(a1,a2,a3,a4,a6,a5)
		case "editdata"
			call editdata(a1,a2,a3,a4,a6,a7,a8,a5)
		case "getdata"
			call getdata(a1,a2,a6,a8,a5)
		case "getfield"
			call getfield(a1,a2,a9,a8,a5)
		case "fieldlist"
			call getfieldslist(a1,a2,a5)
		case "deletefield"
			call deletefield(a1,a2,a9,a5)
		case "editfield"
			call editfield(a1,a2,a9,a7,a8,a5)
		case "edittablename"
			call edittablename(a1,a5,a10,a11)
		case "newtable"
			call AddTable(a1,a5,a2)
		case "deletetable"
			call deletetable(a1,a5,a2)
		case "info"
			call banben()
		case "newdata"
			call newdata(a1)
		case "sqltext"
			call sqltext(a1,a5,a12,a4,a3)
		case "compressionaccess"
			call compressionaccess(a1,a5,"",1)
		case "editpass"
			call compressionaccess(a1,a5,a13,2)
		case "accessBackup"
			call accessBackup(a1,a8)
		case "accessLocale"
			call compressionaccess(a1,a5,a7,3)
		case "editPRIMARY"
			call editPRIMARY(a1,a2,a9,a5)
		case "editIndex"
			call editIndex(a1,a2,a9,a5,a14)
		case "serverinfo"
			call serverinfo()
		case "comlist"
			call comlist()
		case "Bandwidth"
			call Bandwidth()
		case else
			text.Start
			text.categoties "ok"
			text.Completed
	end select
end function
if comad="crossdomain" then
	Response.AddHeader "X-Permitted-Cross-Domain-Policies", "all"
	Response.Write("<?xml version=""1.0""?><cross-domain-policy><allow-access-from domain=""*.access2008.cn"" /></cross-domain-policy>")
else
	Response.Write(text.output)
end if
text.clase

sub banben()
	text.Start
	text.categoties "info"
	text.xmladd mululj,"mulu"
	text.xmladd apiVersion,"Version"
	text.xmladd apiVersionmun,"mun"
	text.Completed
end sub
function mdbjc(wjdz)
	On Error  resume next
	dim conn
	set conn = server.CreateObject(COArray(0))
	conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password='';Data Source="&wjdz
	mdbjc= Err.Number
	conn.close
	set conn = nothing
end function
Function KJHS(INTS)
	dim b
	if ints>=(1024*1024*1024) then
		b=ints/(1024*1024*1024)
		kjhs=formatnumber(b,2,-1)&"GB"
	elseif ints>=(1024*1024) then
		b=ints/(1024*1024)
		kjhs=formatnumber(b,2,-1)&"MB"
	elseif ints>=1000 then
		b=ints/1024
		kjhs=formatnumber(b,2,-1)&"KB"
	else
		kjhs=ints&"字节"
	end if
end Function
sub accessBackup(ByVal a,ByVal b)
	On Error resume next
	dim c
	set fso = Server.CreateObject(COArray(9))
	if b="1" then
		fso.copyfile mululj&"\"&a, mululj&"\"&Left(a, InStrRev(a, ".")) & "bak"
	else
		fso.copyfile mululj&"\"&Left(a, InStrRev(a, ".")) & "bak", mululj&"\"&a
	end if
	if err.number<>0 then
		text.outerr err.Description
	else
		if b="1" then
			text.infoshow 3,4
		else
			text.infoshow 5,4
		end if
	end if
end sub
sub connaccess(byref a)
	set a = server.CreateObject(COArray(0))
	a.open "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password='"&accessPass&"';Data Source="&mululj&"\"&accessData
end sub
sub compressionaccess(ByVal a,ByVal b,ByVal t,ByVal e)
	on error resume next
	dim c,d,conn,xwjm,ee
	ee="Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password="
	set fso = Server.CreateObject(COArray(9))
	set jro = Server.CreateObject(COArray(8))
	call connaccess(conn)
	xwjm=fso.GetTempName
	c=ee&"'"&b&"';Data Source="&mululj&"\"&a
	if e=2 then
		d=ee&"'"&t&"';Data Source="&mululj&"\"&xwjm &";Locale Identifier=" & conn.Properties("Locale Identifier").value & "; Jet OLEDB:Engine Type=" & conn.Properties("Jet OLEDB:Engine Type")
	elseif e=3 then
		d=ee&"'"&b&"';Data Source="&mululj&"\"&xwjm &";Locale Identifier=" & t & "; Jet OLEDB:Engine Type=" & conn.Properties("Jet OLEDB:Engine Type")
	else
		d=ee&"'"&b&"';Data Source="&mululj&"\"&xwjm &";Locale Identifier=" & conn.Properties("Locale Identifier").value & "; Jet OLEDB:Engine Type=" & conn.Properties("Jet OLEDB:Engine Type")
	end if
	conn.close
	jro.CompactDatabase c,d
	if err.number<>0 then
		fso.deletefile mululj&"\"&xwjm
		if e=2 then
			text.outerr 6
		elseif e=3 then
			text.outerr 7
		else
			text.outerr 8
		end if
	else
		fso.DeleteFile mululj&"\"&a
		fso.MoveFile mululj&"\"&xwjm, mululj&"\"&a
		if e=2 then
			text.infoshow 9,4
			call gettable(a,t)
		elseif e=3 then
			text.infoshow 10,4
			call gettable(a,b)
		else
			text.infoshow 11,4
			call gettable(a,b)
		end if
	end if
end sub
sub sqltext(ByVal a,ByVal b,ByVal c,ByVal d,ByVal e)
	on error resume next
	dim conn,cmdTemp,rs
	call connaccess(conn)
	Set cmdTemp = Server.CreateObject("ADODB.Command")
    set rs=server.createobject(COArray(1))
    cmdTemp.CommandText = c
    cmdTemp.CommandType = 1
    Set cmdTemp.ActiveConnection = conn   
    rs.Open cmdTemp, ,1,3
	if err.Number<>0 then
        text.outerr err.Description
    else
	rs.pagesize=e
	text.start
	text.categoties "sqltabledata"
	text.xmladd a,"dataaccess"
	text.xmladd b,"dataaccesspass"
	text.xmladd c,"sql"
	text.xmladd d,"pagenow"
	text.xmladd rs.pageCount,"pageCount"
	text.xmladd rs.recordCount,"recordCount"
	for i=0 to rs.fields.count-1
		text.xmladd rs.fields(i).name,"fields"
	next
	if not (rs.eof or err) then rs.move (cint(d)-1)*cint(e)
	do while not (rs.eof or err)
		text.add "<data1>"
			text.xmladd rs.AbsolutePosition,"datashow"
			for i=0 to rs.fields.count-1
				select case rs.fields(i).type
					case 205
						if not isnull(rs(i)) then
							text.xmladd text.gettxt(12),"datashow"
						else
							text.xmladd "","datashow"
						end if
					case 128
						if not isnull(rs(i)) then
							text.xmladd text.gettxt(13),"datashow"
						else
							text.xmladd "","datashow"
						end if
					case 204
						if not isnull(rs(i)) then
							text.xmladd text.gettxt(14),"datashow"
						else
							text.xmladd "","datashow"
						end if
					case 203
						if len(rs(i))>100 then
							text.xmladd replace(left(rs(i),90)&"...",chr(13)&chr(10),""),"datashow"
						else
							text.xmladd rs(i),"datashow"
						end if
					case else
						text.xmladd rs(i),"datashow"
				end select
		next
		text.add "</data1>"
		j=j+1
		if j>=cint(e) then exit do
		rs.movenext
	loop
	text.Completed
	end if
end sub
sub newdata(ByVal a)
	On Error resume next
	dim cat
	set cat=server.createobject(COArray(2))   
    cat.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&mululj&"/"&a&".mdb"
	call getaccess()
end sub
sub getaccess()
	dim fso,fsoml
	dim sjkbs
	dim j:j=0
 	Set fso = Server.CreateObject(COArray(9))
	Set fsom1= fso.getfolder(mululj)
	text.Start
	text.categoties "accessname"
	for each thing in fsom1.files
		if LCase(right(thing.name,len(thing.name)-InstrRev(thing.name,".")))<>"bak" then
			sjkbs=mdbjc(mululj&"\"&thing.name)
			if(sjkbs<>"-2147467259") then
				j=j+1
				if sjkbs="0" then
					text.add "<access title="""&text.xmldata(thing.name)&""" size="""&text.xmldata(thing.size)&"""  data="""&text.xmldata(thing.name)&""" pass="""" bs=""0"" icon=""iconaccess""/>"
				elseif sjkbs="-2147217843" then
					text.add "<access title="""&text.xmldata(thing.name)&" "&text.gettxt(15)&""" size="""&thing.size&"""  data="""&text.xmldata(thing.name)&""" pass="""" bs=""1"" icon=""iconaccess""/>"
				end if
			end if
		end if
	next  
	text.Completed
	if j=0 then
		text.infoshow mululj&" 目录下未发现数据库，请确认数据库地址设置","数据库目录提示"
	end if
end sub
sub gettable(ByVal a,ByVal b)
	On Error resume next
	dim conn,cat,tbl,fso,fsoml,bs,title
	set cat = server.CreateObject(COArray(2))
	set tbl= server.CreateObject(COArray(3))
	Set fso = Server.CreateObject(COArray(9))
	Set fsoml= fso.GetFile(mululj&"/"&a)
	call connaccess(conn)
	set cat.ActiveConnection = conn
	text.Start
	text.categoties "tablename"
	bs="0"
	title=a
	if len(b)>0 then
		bs="1"
		title=a&" "&text.gettxt(15)
	end if
	text.add "<access title="""&text.xmldata(title)&""" bs="""&text.xmldata(bs)&""" pass="""&text.xmldata(b)&""" size="""&fsoml.size&""" data="""&text.xmldata(a)&""" ReclaimedSpace="""&conn.Properties("Jet OLEDB:Compact Reclaimed Space Amount").Value&""" LocaleIdentifier="""&conn.Properties("Locale Identifier").Value&""" accesstype="""&conn.Properties("Jet OLEDB:Engine Type")&""">"
	for each tbl in cat.Tables
		if tbl.type = "TABLE" then
			text.add "<table access="""&text.xmldata(a)&""" pass="""&text.xmldata(b)&""" title="""&text.xmldata(tbl.name)&""" icon=""icontable""/>"
		end if 
	next
	text.add "</access>"
	text.Completed
	if err.number<>0 then
		if err.number=3709 then
			text.outerr text.gettxt(16)
		else
			text.outerr err.Description&","&err.number
		end if
	end if
end sub
sub getdatalist(ByVal a,ByVal b,ByVal c,ByVal d,ByVal pass)
	dim j,conn,rs,sql,i
	j=0
	call connaccess(conn)
	set rs=server.createobject(COArray(1))
	sql="select * from ["&b&"]"
	rs.open sql,conn,3,3
	rs.pagesize=cint(c)
	text.start
	text.categoties "tabledata"
	text.xmladd a,"dataaccess"
	text.xmladd b,"tablename"
	text.xmladd d,"pagenow"
	text.xmladd pass,"dataaccesspass"
	text.xmladd rs.pageCount,"pageCount"
	text.xmladd rs.recordCount,"recordCount"
	for i=0 to rs.fields.count-1
		text.xmladd rs.fields(i).name,"fields"
	next
	if not (rs.eof or err) then rs.move (cint(d)-1)*cint(c)
	do while not (rs.eof or err)
		text.add "<data1>"
			text.xmladd rs.AbsolutePosition,"datashow"
			for i=0 to rs.fields.count-1
				select case rs.fields(i).type
					case 205
						if not isnull(rs(i)) then
							text.xmladd text.gettxt(12),"datashow"
						else
							text.xmladd "","datashow"
						end if
					case 128
						if not isnull(rs(i)) then
							text.xmladd text.gettxt(13),"datashow"
						else
							text.xmladd "","datashow"
						end if
					case 204
						if not isnull(rs(i)) then
							text.xmladd text.gettxt(14),"datashow"
						else
							text.xmladd "","datashow"
						end if
					case 203
						if len(rs(i))>100 then
							text.xmladd "<-BIG TEXT->","datashow"
						else
							text.xmladd rs(i),"datashow"
						end if
					case else
						text.xmladd rs(i),"datashow"
				end select
		next
		text.add "</data1>"
		j=j+1
		if j>=cint(c) then exit do
		rs.movenext
	loop
	text.Completed
end sub
sub deletedata(ByVal a,ByVal b,ByVal c,ByVal d,ByVal e,ByVal pass)
	dim conn,sql,rs,data,i
	call connaccess(conn)
	set rs=server.createobject(COArray(1))
	sql="select * from ["&b&"]"
	rs.open sql,conn,3,3
	data = split(e,"|")
	for i = 0 to ubound(data)
	rs.AbsolutePosition = cint(data(i))
	rs.delete
	next
	rs.close
	conn.close
	text.start
	text.categoties "editdataend"
	text.Completed
end sub
sub getdata(ByVal a,ByVal b,ByVal c,ByVal d,ByVal pass)
	dim conn,cat,rs,sql,i
	set cat = server.CreateObject(COArray(2))
	call connaccess(conn)
	set rs=server.createobject(COArray(1))
	sql="select * from ["&b&"]"
	rs.open sql,conn,3,3
	if d=0 then
		rs.AbsolutePosition = cint(c)
	end if
	set cat.ActiveConnection = conn
	text.start
	if d=0 then
		text.categoties "editdata"
	else
		text.categoties "getfields"
	end if
	text.xmladd c,"AbsolutePosition"
	for i=0 to rs.fields.count-1
		if rs.fields(i).type<>205 and rs.fields(i).type<>128 and rs.fields(i).type<>204 and cat.Tables(b).Columns(rs.fields(i).name).Properties("Autoincrement")=false then
			text.add "<datashow>"
			text.xmladd rs.fields(i).name,"name"
			text.xmladd rs.fields(i).type,"type1"
			text.xmladd cat.Tables(b).Columns(rs.fields(i).name).Properties("Description").Value,"Description"
			if d=0 then
				text.xmladd rs(i),"data"
			end if
			text.add "</datashow>"
		end if
	next
	text.Completed
	rs.close
	conn.close
end sub
sub editdata(ByVal a,ByVal b,ByVal c,ByVal d,ByVal e,ByVal f,ByVal g,ByVal pass)
	On Error resume next
	dim conn,cat,rs,sql,i,type1,fieldtxt
	set cat = server.CreateObject(COArray(2))
	call connaccess(conn)
	set rs=server.createobject(COArray(1))
	sql="select * from ["&b&"]"
	rs.open sql,conn,3,3
	if g=1 then
		rs.addnew
	else
		rs.AbsolutePosition = cint(e)
	end if
	set cat.ActiveConnection = conn
	Set objXML = Server.CreateObject(COArray(7)) 
		objXML.async = False 
		loadResult = objXML.loadXML(f) 
		Set objNodes = objXML.getElementsByTagName("challs")
	for i=0 to rs.fields.count-1
		type1=rs.fields(i).type
		fieldtxt=objNodes(0).selectSingleNode(rs.fields(i).name).Text
		if type1<>205 and type1<>128 and type1<>204 and cat.Tables(b).Columns(rs.fields(i).name).Properties("Autoincrement")=false and len(fieldtxt)>0 then
		  if type1=11 then
			  if lcase(fieldtxt)="false" then
				  rs(i)=0
			  else
				  rs(i)=1
			  end if
		  else
			  rs(i)=fieldtxt
		  end if
		end if
	next
	rs.update
	rs.close
	conn.close
	text.start
	text.categoties "editdataend"
	text.Completed
	if err.number<>0 and err.number<>424 then
		text.outerr err.Description&","&err.number
    end if
end sub

sub getfield(ByVal a,ByVal b,ByVal c,ByVal d,ByVal pass)
On Error resume next
dim conn,cat,rs,sql,i
	set cat = server.CreateObject(COArray(2))
	call connaccess(conn)
	set rs=server.createobject(COArray(1))
	sql="select top 1 * from ["&b&"]"
	set cat.ActiveConnection = conn
	rs.open sql,conn,3,3
	text.start
	text.categoties "getfield"
	text.xmladd d,"bs"
	text.add "<datashow>"
	text.xmladd rs.fields(c).name,"name"
	text.xmladd rs.fields(c).type,"type"
	text.xmladd fieldtype(rs.fields(c).type,cat.Tables(b).Columns(c).Properties("Autoincrement")),"type1"
	text.xmladd cat.Tables(b).Columns(c).DefinedSize,"DefinedSize"
	text.xmladd cat.Tables(b).Columns(c).Properties("default").Value,"default"
	text.xmladd cat.Tables(b).Columns(c).Properties("Jet OLEDB:Allow Zero Length").Value,"AllowZeroLength"
	text.xmladd cat.Tables(b).Columns(c).Properties("Jet OLEDB:Column Validation Rule").Value,"ColumnValidationRule"
	text.xmladd cat.Tables(b).Columns(c).Properties("Jet OLEDB:Column Validation Text").Value,"ColumnValidationText"
	text.xmladd cat.Tables(b).Columns(c).Properties("Nullable").Value,"Nullable"
	text.xmladd cat.Tables(b).Columns(c).Properties("Jet OLEDB:Compressed UNICODE Strings").Value,"CompressedUNICODEStrings"
	text.xmladd cat.Tables(b).Columns(c).Properties("Autoincrement").Value,"Autoincrement"
	text.xmladd cat.Tables(b).Columns(c).Properties("Description").Value,"Description"
	text.add "</datashow>"
	text.Completed
	rs.close
	conn.close
	if err.number<>0 then
		text.outerr err.Description&","&err.number
    end if
end sub
sub deletefield(ByVal a,ByVal b,ByVal c,ByVal pass)
On Error resume next
dim conn,cat
	call connaccess(conn)
	set cat =server.createobject(COArray(2)) 
	Set cat.ActiveConnection = conn 
	cat.tables(b).columns.delete c
	text.start
	text.categoties "editfieldend"
	text.Completed
	if err.number<>0 then
		text.outerr err.Description&","&err.number
    end if
end sub

sub getfieldslist(ByVal a,ByVal b,ByVal pass)
	On Error resume next
	dim conn,cat,rs,sql,i,zjname,sybs,u,j,ui
	dim syjh(),jj,sythtype()
	dim strSQL,TempSQL,PrimaryKey,PrKey
	dim keySQL
	set cat = server.CreateObject(COArray(2))
	set ckey1 = server.CreateObject(COArray(6))
	set ckey = server.CreateObject(COArray(5))
	call connaccess(conn)
	set rs=server.createobject(COArray(1))
	sql="select top 1 * from ["&b&"]"
	strSQL="CREATE TABLE [" & b & "]"
	set cat.ActiveConnection = conn
	jj=0
	redim Preserve sythtype(1)
	for each tim in cat.Tables(b).keys
		
		if tim.type=1 then
			keySQL=keySQL&VBCrlf&"CREATE "
			set ckey1 = tim
			PrKey=tim.Name
			keySQL=keySQL&"INDEX [" & PrKey & "] ON [" & b & "]("
			for j=0 to ckey1.Columns.count-1
				zjname=ckey1.Columns(j).Name
				keySQL=keySQL&"[" & ckey1.Columns(j).Name & "],"
			next
			keySQL=Left(KeySQL,len(KeySQL)-1)
			keySQL=keySQL& ") WITH PRIMARY;"
		elseif tim.type=3 then
			keySQL=keySQL&VBCrlf&"CREATE "
			set ckey1 = tim
			keySQL=keySQL&"UNIQUE INDEX [" & tim.Name & "] ON [" & b & "]("
			redim Preserve sythtype(ckey1.Columns.count)
			for j=0 to ckey1.Columns.count-1
				sythtype(jj)=ckey1.Columns(j).Name
				keySQL=keySQL&"[" & ckey1.Columns(j).Name & "],"
				jj=jj+1
			next
			keySQL=Left(KeySQL,len(KeySQL)-1)
			keySQL=keySQL& ");"
		end if
	next
	jj=0
	dim TempKey
	if cat.Tables(b).Indexes.count<>0 then
		redim Preserve syjh(cat.Tables(b).Indexes.count)
		for u=0 to cat.Tables(b).Indexes.count-1
			TempKey="CREATE INDEX [" & cat.Tables(b).indexes(u).Name & "] ON [" & b & "]("
			set ckey = cat.Tables(b).indexes(u)
			for j=0 to ckey.Columns.count-1 
				syjh(jj)=ckey.Columns(j).Name
				TempKey=TempKey&"[" & ckey.Columns(j).Name & "],"
				jj=jj+1
			next
			TempKey=Left(TempKey,len(TempKey)-1)
			TempKey=TempKey& ");"
			if cat.Tables(b).indexes(u).Name<>PrKey then
				keySQL=keySQL&VBcrlf&TempKey
			end if
		next
	else
		redim syjh(0)
	end if
	rs.open sql,conn,3,3
	text.start
	text.categoties "getfieldslist"
	text.xmladd a,"dataaccess"
	text.xmladd b,"tablename"
	text.xmladd pass,"dataaccesspass"
	text.xmladd keySQL,"keySQL"
	for i=0 to rs.fields.count-1
			sybs=0
			PrimaryKey=False
			text.add "<datashow>"
			if rs.fields(i).name=zjname and len(zjname)>0 then
				text.xmladd "√","PrimaryKey"
				PrimaryKey=True
			else
				text.add "<PrimaryKey/>"
			end if
			sybs=""
			for u=0 to ubound(syjh)
				if rs.fields(i).name=syjh(u) then
					sybs=text.gettxt(17)
					for ui=0 to ubound(sythtype)
						if rs.fields(i).name=sythtype(ui) then
							sybs=text.gettxt(18)
							exit for
						end if
					next
					
				end if
			next
			dim Temp,DefinedSize_,IsNullable_,Default_
			DefinedSize_=cat.Tables(b).Columns(rs.fields(i).name).DefinedSize
			IsNullable_=cat.Tables(b).Columns(rs.fields(i).name).Properties("Nullable").Value
			Default_=cat.Tables(b).Columns(rs.fields(i).name).Properties("default").Value
			Temp=GetSQLTypeName(rs.fields(i).Type,PrimaryKey,DefinedSize_)
			TempSQL=TempSQL& "[" & rs.fields(i).name & "] " & Temp
			If Temp="TEXT" Then TempSQL=TempSQL & "(" & DefinedSize_ & ")"
			If Not IsNullable_ or PrimaryKey Then TempSQL=TempSQL & " NOT NULL"
			If Len(Default_)>0 Then TempSQL=TempSQL  & " DEFAULT " & Default_ 
			TempSQL=TempSQL & ","
			text.xmladd sybs,"index"
			text.xmladd rs.fields(i).name,"name"
			text.xmladd rs.fields(i).type,"type"
			text.xmladd fieldtype(rs.fields(i).type,cat.Tables(b).Columns(rs.fields(i).name).Properties("Autoincrement")),"type1"
			text.xmladd DefinedSize_,"DefinedSize"
			text.xmladd Default_,"default"
			text.xmladd cat.Tables(b).Columns(rs.fields(i).name).Properties("Jet OLEDB:Allow Zero Length").Value,"AllowZeroLength"
			text.xmladd cat.Tables(b).Columns(rs.fields(i).name).Properties("Jet OLEDB:Column Validation Rule").Value,"ColumnValidationRule"
			text.xmladd cat.Tables(b).Columns(rs.fields(i).name).Properties("Jet OLEDB:Column Validation Text").Value,"ColumnValidationText"
			text.xmladd not IsNullable_,"Nullable"
			text.xmladd cat.Tables(b).Columns(rs.fields(i).name).Properties("Jet OLEDB:Compressed UNICODE Strings").Value,"CompressedUNICODEStrings"

			text.xmladd cat.Tables(b).Columns(rs.fields(i).name).Properties("Autoincrement").Value,"Autoincrement"
			text.xmladd cat.Tables(b).Columns(rs.fields(i).name).Properties("Description").Value,"Description"
			text.add "</datashow>"
	next
	strSQL=strSQL & "("&Left(TempSQL,Len(TempSQL)-1)&");"
	text.xmladd strSQL,"strSQL"
	text.Completed
	rs.close
	conn.close
	if err.number<>0 then
		text.outerr err.Description&","&err.number
    end if
end sub
Function GetSQLTypeName(FieldType_,IsAutonumber,MaxLength_)
	Select Case FieldType_
	Case 3		if IsAutonumber then GetSQLTypeName = "COUNTER" else GetSQLTypeName = "LONG"
	Case 7		GetSQLTypeName = "DATETIME"
	Case 11		GetSQLTypeName = "BIT"
	Case 6		GetSQLTypeName = "MONEY"
	Case 128	GetSQLTypeName = "BINARY"
	Case 17		GetSQLTypeName = "TINYINT"
	Case 131	GetSQLTypeName = "DECIMAL"
	Case 5		GetSQLTypeName = "FLOAT"
	Case 2		GetSQLTypeName = "INTEGER"
	Case 4		GetSQLTypeName = "REAL"
	Case 72		GetSQLTypeName = "UNIQUEIDENTIFIER"
	Case 130	if MaxLength_ = 0 then GetSQLTypeName = "MEMO" else GetSQLTypeName = "TEXT"
	Case 202	GetSQLTypeName = "TEXT"
	Case 203	GetSQLTypeName = "MEMO"
	Case Else	GetSQLTypeName = ""
	End Select
End Function

sub editPRIMARY(ByVal a,ByVal b,ByVal c,ByVal d)
	dim conn,cat,ckey1,zjname,IndexName,tim
	set cat = server.CreateObject(COArray(2))
	set ckey1 = server.CreateObject(COArray(6))
	call connaccess(conn)
	set cat.ActiveConnection = conn
	for each tim in cat.Tables(b).keys		
		if tim.type=1 then
			IndexName = tim.Columns(0).name
			cat.Tables(b).keys.delete tim.name
		end if
	next
	if IndexName<>c and len(c)>0 then
		 conn.execute("CREATE INDEX [PrimaryKey] ON [" & b & "]([" & c & "]) WITH PRIMARY")
	end if
	call getfieldslist(a,b,d)
end sub
sub editIndex(ByVal a,ByVal b,ByVal c,ByVal d,ByVal e)
	dim conn,cat,ckey1,zjname,IndexName,tim,syjh(),u,j,bool,uu,typecf
	bool=true
	typecf=" "
	set cat = server.CreateObject(COArray(2))
	set ckey = server.CreateObject(COArray(5))
	call connaccess(conn)
	set cat.ActiveConnection = conn
	if cat.Tables(b).Indexes.count<>0 then
		redim syjh(cat.Tables(b).Indexes.count-1)
		for u=0 to cat.Tables(b).Indexes.count-1
			set ckey = cat.Tables(b).indexes(u)
			for j=0 to ckey.Columns.count-1 
				syjh(j)=ckey.Columns(j).Name
				if syjh(j)=c then
					IndexName=cat.Tables(b).indexes(u).name
					uu=u
					bool=false
					exit for
				end if
			next
		next
	end if
	if bool then
		if e="1" then
			typecf=" UNIQUE "
		end if
		conn.execute("CREATE"&typecf&"INDEX [] ON [" & b & "]([" & c & "])")
	else
		cat.Tables(b).indexes.delete IndexName
	end if
	call getfieldslist(a,b,d)
end sub
sub editfield(ByVal a,ByVal b,ByVal c,ByVal d,ByVal e,ByVal pass)
	On Error resume next
	dim ColumnName,ColumnType,ColumnLength,ColumnDefault,ColumnDescription,ColumnNullable,ColumnValidRule,ColumnValidText,ColumnZeroLength,ColumnUnicode,newname
	dim typestring
	Set objXML = Server.CreateObject(COArray(7)) 
		objXML.async = False 
		loadResult = objXML.loadXML(d) 
		Set objNodes = objXML.getElementsByTagName("challs")
		if e="0" then
			Select Case objNodes(0).selectSingleNode("type").Text
			  case 2            
			   typestring="short"
			  case 3           
			   typestring="long"
			  case 4            
			   typestring="real"
			  case 5            
			   typestring="double" 
			  case 6           
				typestring="currency"
			  case 7            
			   typestring="datetime"
			  case 11          
			   typestring="yesno" 
			  case 17          
			   typestring="byte" 
			  case 128        
			   typestring="hyperlink"  
			  case 133        
			   typestring="date"
			  case 134        
			   typestring="time" 
			  case 135        
			   typestring="datetime"
			  case 202        
			   typestring="varchar"
			  case 203        
			   typestring="memo"
			  case 204       
			   typestring="OleObject"
			  case 205     
			   typestring="OleObject"
			  case 8888
				typestring="AutoIncrement"
			  case else
			   typestring=objNodes(0).selectSingleNode("type").Text   
			 end Select
			ColumnType=typestring
			newname = objNodes(0).selectSingleNode("name").Text
			ColumnName = objNodes(0).selectSingleNode("oldname").Text
			ColumnLength = objNodes(0).selectSingleNode("DefinedSize").Text
			if ColumnLength="" then ColumnLength=0
			if int(ColumnLength)<=0 then ColumnLength=""
			ColumnDefault = objNodes(0).selectSingleNode("default").Text
			ColumnDescription = objNodes(0).selectSingleNode("Description").Text
			ColumnNullable = objNodes(0).selectSingleNode("Nullable").Text
			ColumnValidRule ="" 
			ColumnValidText =""
			ColumnZeroLength = objNodes(0).selectSingleNode("AllowZeroLength").Text
			ColumnUnicode = objNodes(0).selectSingleNode("CompressedUNICODEStrings").Text	
			call AlterTableColumn(a,b,pass,ColumnName,ColumnType,ColumnLength,ColumnDefault,ColumnDescription,ColumnNullable,ColumnValidRule,ColumnValidText,ColumnZeroLength,ColumnUnicode,newname) 
		else
			call addfield(a,b,c,d,e,pass) 
		end if
end sub
Sub addfield(ByVal a,ByVal b,ByVal c,ByVal d,ByVal e,ByVal pass)
	On Error resume next
	dim conn,cat,rs,sql,i,zjname,sybs,u,j
	set cat = server.CreateObject(COArray(2))
	set mytable=server.createobject(COArray(3))
	set myfield =server.createobject(COArray(4))
	call connaccess(conn)
	Set objXML1 = Server.CreateObject(COArray(7)) 
	objXML1.async = False 
	loadResult = objXML1.loadXML(d) 
	Set objNodes1 = objXML1.getElementsByTagName("challs")
	if objNodes1(0).selectSingleNode("type").Text = "11" then
		if (objNodes1(0).selectSingleNode("DefinedSize").Text<>"" and isNumeric(objNodes1(0).selectSingleNode("DefinedSize").Text)) or objNodes1(0).selectSingleNode("DefinedSize").Text="0" then
		  mrz = 0
		else
		  mrz = 1
	    end if
		conn.execute("ALTER TABLE ["&b&"] ADD COLUMN ["&objNodes1(0).selectSingleNode("name").Text&"] yesno default "&mrz)
		PropertiesEdit Conn,b,objNodes1(0).selectSingleNode("name").Text,"",objNodes1(0).selectSingleNode("Description").Text,objNodes1(0).selectSingleNode("AllowZeroLength").Text
	elseif objNodes1(0).selectSingleNode("type").Text = "7" then
		conn.execute("ALTER TABLE ["&b&"] ADD COLUMN ["&objNodes1(0).selectSingleNode("name").Text&"] datetime")
		PropertiesEdit Conn,b,objNodes1(0).selectSingleNode("name").Text,objNodes1(0).selectSingleNode("default").Text,objNodes1(0).selectSingleNode("Description").Text,objNodes1(0).selectSingleNode("AllowZeroLength").Text
	elseif objNodes1(0).selectSingleNode("type").Text = "133" then
		conn.execute("ALTER TABLE ["&b&"] ADD COLUMN ["&objNodes1(0).selectSingleNode("name").Text&"] date")
		PropertiesEdit Conn,b,objNodes1(0).selectSingleNode("name").Text,objNodes1(0).selectSingleNode("default").Text,objNodes1(0).selectSingleNode("Description").Text,objNodes1(0).selectSingleNode("AllowZeroLength").Text
	else
	  set cat.ActiveConnection = conn
	  set myfield.ParentCatalog=cat
	  myfield.Name = objNodes1(0).selectSingleNode("name").Text
	  if objNodes1(0).selectSingleNode("type").Text = "8888" then
	   myfield.Type = 3
	   myfield.Properties("AutoIncrement") = true
	  else
	   myfield.Type = int(objNodes1(0).selectSingleNode("type").Text)
	  end if
	  myfield.Properties("Nullable").Value=not CBool(objNodes1(0).selectSingleNode("Nullable").Text)
	  if objNodes1(0).selectSingleNode("DefinedSize").Text<>"" and isNumeric(objNodes1(0).selectSingleNode("DefinedSize").Text) then
		  myfield.DefinedSize = int(objNodes1(0).selectSingleNode("DefinedSize").Text)
	  end if
	  set mytable=cat.Tables(b)
	  mytable.Columns.Append myfield
	  set myfield=mytable.Columns(objNodes1(0).selectSingleNode("name").Text)
	  with myfield
		  .Properties("Description").Value= objNodes1(0).selectSingleNode("Description").Text
		  .Properties("Jet OLEDB:Allow Zero Length").Value= CBool(objNodes1(0).selectSingleNode("AllowZeroLength").Text)
		  .Properties("Jet OLEDB:Compressed UNICODE Strings").Value=CBool(objNodes1(0).selectSingleNode("CompressedUNICODEStrings").Text)
	  end with
	  set mytable=nothing
	  set myfield=nothing
	  set cat=nothing
	  if len(objNodes1(0).selectSingleNode("default").Text)>0 then
		  conn.Execute("ALTER TABLE [" & b & "] ALTER COLUMN [" & objNodes1(0).selectSingleNode("name").Text & "] SET DEFAULT " & objNodes1(0).selectSingleNode("default").Text)
	  end if
	end if
	if err.number<>0 and err.number<>-2147217887 then
		text.outerr err.Description&","&err.number
    end if
	text.ReInfoOrErr 20,4,"infoshow",array(b,objNodes1(0).selectSingleNode("name").Text)
	text.start
	text.categoties "editfieldend"
	text.Completed
end Sub
sub PropertiesEdit(Conn,TableName,ColumnName,ColumnDefault,ColumnDescription,ColumnZeroLength)
  dim mydb,mytable,myfield
  set mydb=server.createobject(COArray(2))
  set mytable=server.createobject(COArray(3))
  set myfield =server.createobject(COArray(4))
  MyDB.ActiveConnection =Conn
  For Each MyTable In MyDB.Tables
	if MyTable.Name=TableName then
  For Each MyField In MyTable.Columns
	if  MyField.Name=ColumnName Then
	  Res=1
	  
	  if  MyField.Properties("Default").Value<>ColumnDefault and len(ColumnDefault)>0 then
	  MyField.Properties("Default").Value=ColumnDefault
	  end if
	  
	  if  MyField.Properties("Description").Value<>ColumnDescription and len(ColumnDescription)>0 then
	  MyField.Properties("Description").Value=ColumnDescription
	  end if
	 
	  
	  if  MyField.Properties("Jet OLEDB:Allow Zero Length").Value<>ColumnZeroLength and len(ColumnZeroLength)>0 then
	  MyField.Properties("Jet OLEDB:Allow Zero Length").Value=ColumnZeroLength
	  end if
	 
	 exit for 
  end if 
  Next
  if Res=1 then exit for
  end if
  if Res=1 then exit for
Next
end sub
Sub AlterTableColumn(PathName,TableName,pass,ColumnName,ColumnType,ColumnLength,ColumnDefault,ColumnDescription,ColumnNullable,ColumnValidRule,ColumnValidText,ColumnZeroLength,ColumnUnicode,newname) 
On Error resume next
dim Conn
call connaccess(Conn)
PropertiesEdit Conn,TableName,ColumnName,ColumnDefault,ColumnDescription,ColumnZeroLength
if ColumnNullable=True then
ColumnNullable=" Null "
else
ColumnNullable=" Not Null "
end if

sql="Alter Table ["&TableName&"] Alter Column "

select case ColumnType
case  "AutoIncrement" 
sql=sql&"["&ColumnName&"] AutoIncrement "&ColumnNullable

case "varchar"
 if ColumnLength="" then
 sql=sql&"["&ColumnName&"] varchar(50) "&ColumnNullable
 else
 sql=sql&"["&ColumnName&"] varchar("&cint(ColumnLength)&") "&ColumnNullable
 end if 
 if ColumnDefault<>"" then 
 sql=sql&"  default "&ColumnDefault
 else
 sql=sql
 end if

case "memo"
  if ColumnDefault<>"" then 
  sql=sql&"["&ColumnName&"] memo "&"  default "&ColumnDefault
  else
  sql=sql&"["&ColumnName&"] memo "&ColumnNullable
  end if
  
case "integer"
  if ColumnLength="" then
  sql=sql&"["&ColumnName&"] integer "&ColumnNullable
  else
  sql=sql&"["&ColumnName&"] integer("&ColumnLength&") "&ColumnNullable
  end if 
  if ColumnDefault<>"" then 
  sql=sql&"  default "&ColumnDefault
  else
  sql=sql
  end if
  
case "number"
  if ColumnLength="" then
  sql=sql&"["&ColumnName&"] number "&ColumnNullable
  else
  sql=sql&"["&ColumnName&"] number("&ColumnLength&") "&ColumnNullable
  end if 
  if ColumnDefault<>"" then 
  sql=sql&"  default "&ColumnDefault
  else
  sql=sql
  end if
  
case "short"
  if ColumnLength="" then
  sql=sql&"["&ColumnName&"] short "&ColumnNullable
  else
  sql=sql&"["&ColumnName&"] short("&ColumnLength&") "&ColumnNullable
  end if 
  if ColumnDefault<>"" then 
  sql=sql&"  default "&ColumnDefault
  else
  sql=sql
  end if
case "long"
  if ColumnLength="" then
  sql=sql&"["&ColumnName&"] long "&ColumnNullable
  else
  sql=sql&"["&ColumnName&"] long("&ColumnLength&") "&ColumnNullable
  end if 
  if ColumnDefault<>"" then 
  sql=sql&"  default "&ColumnDefault
  else
  sql=sql
  end if 
  
case "double"
  if ColumnLength="" then
  sql=sql&"["&ColumnName&"] double "&ColumnNullable
  else
  sql=sql&"["&ColumnName&"] double("&ColumnLength&") "&ColumnNullable
  end if 
  if ColumnDefault<>"" then 
  sql=sql&"  default "&ColumnDefault
  else
  sql=sql
  end if
  
case "real"
  if ColumnLength="" then
  sql=sql&"["&ColumnName&"] real "&ColumnNullable
  else
  sql=sql&"["&ColumnName&"] real("&ColumnLength&") "&ColumnNullable
  end if 
  if ColumnDefault<>"" then 
  sql=sql&"  default "&ColumnDefault
  else
  sql=sql
  end if
  
case "numeric"
  if ColumnLength="" then
  sql=sql&"["&ColumnName&"] numeric "&ColumnNullable
  else
  sql=sql&"["&ColumnName&"] numeric("&ColumnLength&") "&ColumnNullable
  end if 
  if ColumnDefault<>"" then 
  sql=sql&"  default "&ColumnDefault
  else
  sql=sql
  end if
  
case "byte"
  if ColumnLength="" then
  sql=sql&"["&ColumnName&"] byte "&ColumnNullable
  else
  sql=sql&"["&ColumnName&"] byte("&ColumnLength&") "&ColumnNullable
  end if 
  if ColumnDefault<>"" then 
  sql=sql&"  default "&ColumnDefault
  else
  sql=sql
  end if
   
case "datetime"
 if ColumnDefault="" then
 sql=sql&"["&ColumnName&"] datetime "&ColumnNullable
 else
 sql=sql&"["&ColumnName&"] datetime "&ColumnNullable&"  default "&ColumnDefault
 end if 
 
case "date"
  if ColumnDefault="" then
  sql=sql&"["&ColumnName&"] date "&ColumnNullable
  else
  sql=sql&"["&ColumnName&"] date  "&ColumnNullable&" default "&ColumnDefault
  end if 
  
case "time"
  if ColumnDefault="" then
  sql=sql&"["&ColumnName&"] time "&ColumnNullable
  else
  sql=sql&"["&ColumnName&"] time  "&ColumnNullable&" default "&ColumnDefault
  end if 
case "yesno"
  if ColumnDefault="" then
  sql=sql&"["&ColumnName&"] yesno "&ColumnNullable
  else
  sql=sql&"["&ColumnName&"] yesno "&ColumnNullable&"  default "&ColumnDefault
  end if
  
case "currency"
  if ColumnLength="" then
  sql=sql&"["&ColumnName&"] currency "&ColumnNullable
  else
  sql=sql&"["&ColumnName&"] currency("&ColumnLength&") "&ColumnNullable
  end if 
  
  if ColumnDefault<>"" then 
  sql=sql&"  default "&ColumnDefault
  else
  sql=sql
  end if

case "hyperlink"
  if ColumnDefault="" then
  sql=sql&"["&ColumnName&"] OleObject "&ColumnNullable
  else
  sql=sql&"["&ColumnName&"] OleObject "&ColumnNullable&"  default "&ColumnDefault
  end if
case "OleObject"
  if ColumnDefault="" then
  sql=sql&"["&ColumnName&"] OleObject "&ColumnNullable
  else
  sql=sql&"["&ColumnName&"] OleObject "&ColumnNullable&"  default "&ColumnDefault
  end if  
case else
	text.ReInfoOrErr 21,"","outerr",array(ColumnType)
end select
conn.execute(sql)
if ColumnName<>newname and len(newname)>0 then
	MyField.name=newname
end if

	conn.close
	set Conn=nothing
	text.ReInfoOrErr 22,4,"infoshow",array(tablename,ColumnName)
	text.start
	text.categoties "editfieldend"
	text.Completed
   if err.number<>0 then
		text.outerr err.Description&","&err.number&","&err.line
    end if
End Sub

function fieldtype(ByVal a,ByVal b)
	 Select Case a
	  case 2            
	   fieldtype=text.gettxt(25)
	  case 3           
	   fieldtype=text.gettxt(26)
	  case 4            
	   fieldtype=text.gettxt(27)
	  case 5            
	   fieldtype=text.gettxt(28)
	  case 6           
		fieldtype=text.gettxt(29)
	  case 7            
	   fieldtype=text.gettxt(30)
	  case 11          
	   fieldtype=text.gettxt(31)
	  case 17          
	   fieldtype=text.gettxt(32)
	  case 128        
	   fieldtype=text.gettxt(33)
	  case 133        
	   fieldtype=text.gettxt(34)
	  case 134        
	   fieldtype=text.gettxt(35)
	  case 135        
	   fieldtype=text.gettxt(36)
	  case 202        
	   fieldtype=text.gettxt(37)
	  case 203        
	   fieldtype=text.gettxt(38)
	  case 204       
	   fieldtype=text.gettxt(39)
	  case 205     
	   fieldtype=text.gettxt(40)
	  case else
	   fieldtype=a
	 end Select
	 if b=true then
	   fieldtype=text.gettxt(41)
	 end if

end function 


sub edittablename(ByVal a,ByVal b,ByVal c,ByVal d)
	Dim oTbl,DBConn2
	Set DBConn2 = Server.CreateObject(COArray(2))
	DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password='"&b&"';Data Source="&mululj&"/"&a
	DBConn2.ActiveConnection = DSN
	Set oTbl = Server.CreateObject(COArray(3))
	Set oTbl = DBConn2.Tables(c) 
	oTbl.Name = d  
	Set oTbl = Nothing
	Set DBConn2 = Nothing
	text.start
	text.categoties "edittablenameend"
	text.xmladd a,"access"
	text.xmladd b,"pass"
	text.Completed
end sub

sub AddTable(ByVal a,ByVal b,ByVal c)
	Dim oTbl,DBConn,myfield,DSN
	call connaccess(DBConn)
	DBConn.execute("CREATE TABLE ["&c&"] ([ID] counter, CONSTRAINT [Index1] PRIMARY KEY ([ID]))")
	text.start
	text.categoties "addtableend"
	text.xmladd a,"access"
	text.xmladd b,"pass"
	text.Completed
End sub 

sub deletetable(ByVal a,ByVal b,ByVal c)
	Dim DBConn2,DSN
	Set DBConn2 = Server.CreateObject(COArray(2))
	DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password='"&b&"';Data Source="&mululj&"/"&a
	DBConn2.ActiveConnection = DSN
	DBConn2.Tables.Delete c
	text.start
	text.categoties "deletetableend"
	text.xmladd a,"access"
	text.xmladd b,"pass"
	text.Completed
end sub

Class TextData
	Private text()
	Private id
	Private max
	Private txt
	Private Sub Class_Initialize 
		redim text(12)
		id=0
		max=11
		txt=array("API文件密码错误,请重新输入密码！","API文件里设置的数据库目录不存在,请重新设置API文件!","服务器数据传输错误,请重新登入！","备份文件创建成功","信息提示","数据库已经还原为备份数据","密码修改失败！","数据库编码失败！","数据库修复压缩失败！","密码修改成功","数据库地区修改成功","数据库压缩成功","[OLE对象]","[超连接]","[二进制数据]","(密码)","ACCESS数据库密码不正确","(重复)","(唯一)","|TextData类文本索引错误|","数据表%1表中字段 %2 新建完成","数据类%1不可以识别或者暂时未完善此类别数据类型的建表功能","数据表%1表中字段 %2 修改常用属性完成"," (日期格式不规范)","(未知)","整形","长整形","单精度","双精度","货币","日期/时间","布尔","字节型","二进制","日期","时间","日期时间","文本","备注","二进制","OLE对象","自动编号")
	End Sub
	Public function xmldata(ByVal value)
		xmldata=replace(replace(replace(value&"","&","&amp;"),">","&gt;"),"<","&lt;")
	end function
	Public Default Sub Add(ByVal value)
		text(id) = value
		id = id+1
		If id>=max Then
			max = max + 12
			Redim Preserve text(max)
		End if
	End Sub
	Public Sub xmlAdd(ByVal value,ByVal bq)
		dim a
		if len(value)>0 then
			a=xmldata(value)
			add "<"&bq&">"
			add a
			add "</"&bq&">"
		else
			add "<"&bq&"/>"
		end if
	End Sub
	Public Sub outerr(ByVal value)
		dim a
		if VarType(value)=2 or VarType(value)=3 then
			a=GetTxt(value)
		else
			a=value
		end if
		add "<challs><categories>merr</categories><text>"&xmldata(a)&"</text></challs>"
	end Sub
	Public Sub infoshow(ByVal value,ByVal title)
		dim a,b
		if VarType(value)=2 or VarType(value)=3 then
			a=GetTxt(value)
		else
			a=value
		end if
		if VarType(title)=2 or VarType(title)=3 then
			b=GetTxt(title)
		else
			b=title
		end if
		add "<challs><categories>infoshow</categories><text>"&xmldata(a)&"</text><title>"&xmldata(b)&"</title></challs>"
	end Sub
	Public Sub ReInfoOrErr(ByVal value,ByVal title,ByVal type1,ByVal bs)
		dim i,a,b
		a=""
		b=""
		if VarType(bs)>8000 then
			if VarType(value)=2 or VarType(value)=3 then
				a=GetTxt(value)
			else
				a=value
			end if
			if type1="infoshow" then
				if VarType(title)=2 or VarType(title)=3 then
					b=GetTxt(title)
				else
					b=title
				end if	
			end if
			for i=0 to ubound(bs)
				if len(a)>0 then
					a=replace(a,"%"&(i+1),bs(i))
				end if
				if len(b)>0 then
					b=replace(b,"%"&(i+1),bs(i))
				end if
			next
			if type1="infoshow" then
				call infoshow(a,b)
			elseif type1="outerr" then
				call outerr(a)
			end if
		end if
	end sub
	Public Sub CDATA(ByVal value)
		add "<![CDATA["&value&"]]>"
	end Sub
	Public Sub categoties(ByVal value)
		add "<categories>"&xmldata(value)&"</categories>"
	end Sub
	Public Sub Start()
		add "<challs>"
	end Sub
	Public Sub Completed()
		add "</challs>"
	end Sub
	public Function GetTxt(ByVal t)
		if t>ubound(txt) then
			GetTxt=txt(19)
		else
			GetTxt=txt(t)
		end if
	end Function
	Public Function Output()
		Redim preserve text(id-1)
		Output = "<?xml version=""1.0"" encoding=""utf-8""?><date>"&join(text,"")&"</date>"
	end Function
	Public Sub Clase()'清空
		redim text(12)
		id=0
		max=11
	end Sub
end Class
sub serverinfo()
	dim tnow,oknow
	text.start
	text.categoties "serverinfo"
	text.xmladd Request.ServerVariables("SERVER_NAME"),"name" 
	text.xmladd Request.ServerVariables("LOCAL_ADDR"),"ip" 
	text.xmladd Request.ServerVariables("SERVER_PORT"),"Port"
	tnow = now():oknow = cstr(tnow)
	  if oknow <> year(tnow) & "-" & month(tnow) & "-" & day(tnow) & " " & hour(tnow) & ":" & right(FormatNumber(minute(tnow)/100,2),2) & ":" & right(FormatNumber(second(tnow)/100,2),2) then oknow = oknow & text.gettxt(23)
	text.xmladd oknow,"time"
	text.xmladd Request.ServerVariables("SERVER_SOFTWARE"),"IISVersion" 
	text.xmladd Server.ScriptTimeout,"ScriptTimeout"
	text.xmladd Request.ServerVariables("PATH_TRANSLATED"),"FilePath"
	text.xmladd ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion,"ScriptEngine"
	call getsysinfo()
	text.Completed
end sub
sub getsysinfo()
  on error resume next
  dim okCPUS, okCPU, okOS
  Set WshShell = server.CreateObject("WScript.Shell")
  Set WshSysEnv = WshShell.Environment("SYSTEM")
  okOS = cstr(WshSysEnv("OS"))
  okCPUS = cstr(WshSysEnv("NUMBER_OF_PROCESSORS"))
  okCPU = cstr(WshSysEnv("PROCESSOR_IDENTIFIER"))
  if isempty(okCPUS) then
    okCPUS = Request.ServerVariables("NUMBER_OF_PROCESSORS")
  end if
  if okCPUS & "" = "" then
    okCPUS = text.gettxt(24)
  end if
  if okOS & "" = "" then
    okOS = text.gettxt(24)
  end if
  text.xmladd okOS,"OS"
  text.xmladd okCPUS,"CPUS"
  text.xmladd okCPU,"CPU"
end sub
sub comlist()
	text.start
	text.categoties "comlist"
	for i=0 to ubound(COArray)
		text.add "<data>"
		call ObjTest(COArray(i))
		text.add "</data>"
	next
	text.Completed
end sub
sub ObjTest(strObj)
  on error resume next
  dim IsObj,VerObj
  IsObj=false
  VerObj=""
  set TestObj=server.CreateObject (strObj)
  If -2147221005 <> Err then
    IsObj = True
    VerObj = TestObj.version
    if VerObj="" or isnull(VerObj) then VerObj=TestObj.about
  end if
  text.xmladd strObj,"strObj" 
  text.xmladd IsObj,"IsObj"
  text.xmladd VerObj,"VerObj"
  set TestObj=nothing
End sub
sub Bandwidth()
	dim a(1000)
	dim b,c
	text.start
	text.categoties "Bandwidth"
	for i=1 to 1000
		a(i)="|---567890#########0#########0#########0#########0#########0#########0#########0#########012345--|"
	next
	text.xmladd join(a,""),"a"
	text.Completed
end sub
%>