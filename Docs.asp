<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
Checking "0|1"
Dim ObjectID, Conn, rs, aListValues, CountListValues, JavaMsg, i, Action, FS, FID, Week, Corr, BLType
Dim FileName1, FileName2, FileName3, FileName4, Dest1, Dest2, Dest3, Dest4, Files, FilesID, Display

	ObjectID = CheckNum(Request("OID"))
  	Action = checkNum(Request("Action"))
	Display = checkNum(Request("D"))
	BLType = checkNum(Request("AT"))
	CountListValues = -1

	OpenConn Conn
	
	Set rs = Conn.Execute("select a.Week, a.BLNumber from BLs a where a.BLID=" & ObjectID)
	If Not rs.EOF Then
		Week = rs(0)
		Corr = Left(rs(1),3) & Right(rs(1),5)
	End If
	CloseOBJ rs

	set FS=Server.CreateObject("Scripting.FileSystemObject")
	Select Case Action
	case 1
		FileName1 = Request("FileName1")
		FileName2 = Request("FileName2")
		FileName3 = Request("FileName3")
		FileName4 = Request("FileName4")
		Dest1 = Request.Form("Dest1")
		Dest2 = Request.Form("Dest2")
		Dest3 = Request.Form("Dest3")
		Dest4 = Request.Form("Dest4")
		SetFile FileName1, Dest1, FS, Conn, 1, Week, Corr, 0, ObjectID, 0
		SetFile FileName2, Dest2, FS, Conn, 1, Week, Corr, 0, ObjectID, 0
		SetFile FileName3, Dest3, FS, Conn, 1, Week, Corr, 0, ObjectID, 0
		SetFile FileName4, Dest4, FS, Conn, 1, Week, Corr, 0, ObjectID, 0
	case 3
		FID = checkNum(Request("FID"))
		FileName1 = Request("FileName1")
		SetFile FileName1, "", FS, Conn, 3, 0, 0, FID, 0, 0
	case 4
		FilesID = Request.Form("FilesID")
		if FilesID <> "" then
			Files = Split(FilesID, "|")
			Conn.Execute("update Files set InTransit=3 where FileID=" & Join(Files, " or FileID="))
			Set rs = Conn.Execute("select FileName, CountriesFinalDes, BLID from Files where FileID=" & Join(Files, " or FileID="))
			if not rs.EOF then
				for i=0 to rs.RecordCount-1
					SetFile rs(0), rs(1), FS, Conn, 4, Week, Corr, 0, ObjectID, rs(2)
					rs.MoveNext
				next
			end if
			CloseOBJ rs
		end if
	end select
	set FS=nothing

	if ObjectID <> 0  then
	
		if BLType >=0 then 'muestra los archivos de cada carta porte
			Set rs = Conn.Execute("select a.FileID, a.FileName, a.CountriesFinalDes, b.CountryDes from Files a, BLs b where a.BLID=b.BLID and b.BLID=" & ObjectID)
		else 'Cuando es consolidado muestra todos los archivos de sus cartas portes
			Set rs = Conn.Execute("select a.FileID, a.FileName, a.CountriesFinalDes, b.CountryDes from Files a, BLs b, BLGroupDetail c where a.BLID=b.BLID and b.BLID=c.BLID and c.BLGroupID=" & ObjectID)
		end if

		If Not rs.EOF Then
   			aListValues = rs.GetRows
       		CountListValues = rs.RecordCount-1
	    End If
		CloseOBJ rs
	end if
	CloseOBJ Conn

Function SetFile (FileName, CountriesFinalDes, FS, Conn, Action, Week, Corr, FID, BLID, BLIDTransit)
	Dim CreatedDate, CreatedTime, newFile, rs

	if (FileName<>"") then
		
		select case Action
		case 1 'Insert
			newFile = Week & "-" & Corr & "-" & FileName
			if FS.FileExists(Session("PhysicalPath") & "\" & newFile) then
				FS.DeleteFile(Session("PhysicalPath") & "\" & newFile)
			end if
			'response.write Session("PhysicalPath") & "\" & FileName & "-" & Session("PhysicalPath") & "\" & newFile & "<br>"
			FS.MoveFile Session("PhysicalPath") & "\" & FileName, Session("PhysicalPath") & "\" & newFile

			'Obteniendo los parametros para hacer las operaciones de Insert, Update o Delete
			FormatTime CreatedDate, CreatedTime

			Set rs = Conn.execute("select FileID from Files where FileName='" & newFile & "'" )
			if rs.EOF Then  'Si no existe el archivo, puede ingresarlo
				 Conn.Execute("insert into Files (FileName, CreatedDate, CreatedTime, CountriesFinalDes, BLID, InTransit, BLIDTransit) values ('" & newFile & "', '" & CreatedDate & "', " & CreatedTime & ", '" & CountriesFinalDes & "', " & BLID & ", 1, " & BLIDTransit & ")")
			end if
		case 3 'Delete
			Set rs = Conn.execute("select FileID, BLIDTransit from Files where FileID=" & FID)
			
			if Not rs.EOF Then 'Si existe el archivo, puede borrarlo
				'if CInt(rs(1)) = 0 then
				'	if FS.FileExists(Session("PhysicalPath") & FileName) then
				'		FS.DeleteFile(Session("PhysicalPath") & FileName)
				'	end if				
				'end if
				Conn.Execute("delete from Files where FileID=" & FID)
			end if
		case 4
			FormatTime CreatedDate, CreatedTime
			Conn.Execute("insert into Files (FileName, CreatedDate, CreatedTime, CountriesFinalDes, BLID, InTransit, BLIDTransit) values ('" & FileName & "', '" & CreatedDate & "', " & CreatedTime & ", '" & CountriesFinalDes & "', " & BLID & ", 1, " & BLIDTransit & ")")
		end select
		
		CloseOBJ rs
	end if
End Function

%>
<HTML>
<HEAD><TITLE>Aimar - Terrestre</TITLE></HEAD>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<LINK REL="stylesheet" type="text/css" HREF="img/estilos.css">
<style type="text/css">
.style4 {
	font-size:10px;
	color: #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
}
</style>
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0"  onLoad="JavaScript:self.focus()">
	<TABLE cellspacing=1 cellpadding=2 width=100% align=center>
	<%if Display=0 then%>
	<tr><td colspan="4" align="right" class="style4"><a href="UploadFiles.asp?OID=<%=ObjectID%>" class=submenu><font color="FFFFFF">Subir Archivos</font></a></tr>
	<%end if%>
	<tr><td class=titlelist><b>Pais&nbsp;Transito</b></td><td class=titlelist><b>Pais&nbsp;Destino</b></td><td class=titlelist colspan="2"><b>Archivo</b></td></tr>
	<%
	for i=0 to CountListValues%>
	<tr>
	<td class=list><a class=labellist href="<%=Session("VirtualPath") & aListValues(1,i)%>" target="_blank"><%=aListValues(3,i)%></a></td>
	<td class=list><a class=labellist href="<%=Session("VirtualPath") & aListValues(1,i)%>" target="_blank"><%=aListValues(2,i)%></a></td>
	<td class=list><a class=labellist href="<%=Session("VirtualPath") & aListValues(1,i)%>" target="_blank"><%=aListValues(1,i)%></a></td>
	<%if Display=0 then%>
	<td class=list><a class=labellist href="Docs.asp?Action=3&FileName1=<%=aListValues(1,i)%>&FID=<%=aListValues(0,i)%>&OID=<%=ObjectID%>">Borrar</a></td>
	<%end if%>
	</tr>
	<%next

	Set aListValues = Nothing
	%>
	</TABLE>
</BODY>
</HTML>