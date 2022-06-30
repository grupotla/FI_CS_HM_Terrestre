<%
Dim aBLIDs, CountBLIDs, separator, Msg
Checking "0|1"
if CountTableValues >= 0 then
	CreatedDate = ConvertDate(aTableValues(1, 0),2)
    CreatedTime = aTableValues(2, 0)
	Expired = aTableValues(3, 0)
	BLNumber = aTableValues(4, 0)
	Countries = aTableValues(5, 0)
	Week = aTableValues(6, 0)
	CountryDes = aTableValues(7, 0)
end if
Set aTableValues = Nothing
	
	OpenConn Conn
	if Action=1 or Action=2 then
		Conn.Execute("delete from BLGroupDetail where BLGroupID=" & ObjectID)
		aBLIDs = Split(Replace(Request.Form("BLIDs")," ", "",1,-1),chr(13)&chr(10))
		CountBLIDs = UBound(aBLIDs, 1)
		for i=0 to CountBLIDs
			if aBLIDs(i) <> "" then
				Set rs = Conn.Execute("select BLID from BLs where BLNumber='" & aBLIDs(i) & "'")
				if Not rs.EOF then
					Conn.Execute("insert into BLGroupDetail (BLGroupID, BLID) values (" & ObjectID & ", " & rs(0) & ")")
				else
					Msg = Msg & aBLIDs(i) & ", "
				end if
				CloseOBJ rs
			end if
		next
		Set aBLIDs = Nothing
	end if
	
	CountBLIDs = -1
	Set rs = Conn.Execute("select a.BLNumber from BLs a, BLGroupDetail b where a.BLID=b.BLID and b.BLGroupID=" & ObjectID)
	do while Not rs.EOF
		BLIDs = BLIDs & separator & rs(0)
		separator = chr(13)&chr(10)
		CountBLIDs = CountBLIDs + 1
		rs.MoveNext
	loop
	CloseOBJs rs, Conn
%>
<HTML>
<HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
	function validar(Action) {
  		if (Action != 3) {
			if (!valSelec(document.forma.Countries)){return (false)};		
			if (!valSelec(document.forma.CountryDes)){return (false)};
			if (!valTxt(document.forma.Week, 2, 4)){return (false)};
			if (!valTxt(document.forma.BLIDs, 20, 5)){return (false)};
		}
	    document.forma.Action.value = Action;
		document.forma.submit();			 
	 }
</SCRIPT>
<%if Msg <> "" then%>
	<SCRIPT>alert('NO EXISTEN LAS CARTAS PORTE <%=Msg%>');</SCRIPT>
<%end if%>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<style type="text/css">
.style10 {font-size: 10px; font-family: Verdana, Arial, Helvetica, sans-serif; }
</style>
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="Javascript:self.focus();">
	<TABLE cellspacing=0 cellpadding=2 width=400 align=center>
	<FORM name="forma" action="InsertData.asp" method="post">
	<INPUT name="Action" type=hidden value=0>
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
	<INPUT name="OID" type=hidden value="<%=ObjectID%>">
	<INPUT name="CD" type=hidden value="<%=CreatedDate%>">
	<INPUT name="CT" type=hidden value="<%=CreatedTime%>">
		<TR><TD class=label align=right><b>Código:</b></TD><TD class=label align=left><%if ObjectID <> 0 then response.write ObjectID End if%></TD></TR> 
		<TR><TD class=label align=right><b>Fecha Creación:</b></TD><TD class=label align=left><%=CreatedDate%></TD></TR> 
        <TR><TD class=label align=right><b>Activa:</b></TD><TD class=label align=left><INPUT name=Expired TYPE=checkbox class=label <%If Expired = False Then response.write " checked"  End If%>></TD></TR>
		<TR>
		  <TD class=label align=right><b>Pais Origen:</b></TD>
		  <TD class=label align=left colspan=2>
			<select name="Countries" class=label id="Pais de Origen">
			<option value='-1'>Seleccionar</option>
			<%DisplayCountries "", 2%>
		</select>
		</TD></TR>
		<TR>
		  <TD class=label align=right><b>Pais Destino:</b></TD>
		  <TD class=label align=left colspan=2>
			<select name="CountryDes" class=label id="Pais Destino">
			<option value='-1'>Seleccionar</option>
			<%DisplayCountries "", 2%>
		</select>
		</TD></TR>
		<TR><TD class=label align=right><b>Semana:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Week" id="Semana" value="<%=Week%>" maxlength="5" size="5" onKeyUp="res(this,numb);"></TD></TR>
		<TR><TD class=label align=right><b>No. de Carta Porte Grupo:</b></TD><TD class=label align=left><b><%=BLNumber%></b></TD></TR> 
		<TR><TD class=label align=right><b>Cartas Porte:</b></TD><TD class=style4 align=left colspan=2 >
			<textarea class="style10" name="BLIDs" rows="10" cols="29" id="No. Cartas Porte para agrupar"><%=BLIDs%></textarea>
		</TD></TR>
		<TR>
		<TD colspan="2" class=label align=center>
			<TABLE cellspacing=0 cellpadding=2 width=200>
			<TR>
			<%if CountTableValues = -1 then%>
			 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(4)" value="&nbsp;&nbsp;Agregar&nbsp;&nbsp;" class=label></TD>
			<%else
				if CountBLIDs > -1 then%>
			 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="Javascript:if (document.forma.BLIDs.value!='') {window.open('BLPrintConditions.asp?BLID=<%=ObjectID%>&BTP=4','BLGPrint','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750');return false;} else {alert('Debe ingresar al menos una Carta Porte');}" value="&nbsp;&nbsp;Previsualizar Carta Porte Grupo&nbsp;&nbsp;" class=label></TD>
			 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="Javascript:window.open('Reports.asp?GID=22&OID=<%=ObjectID%>','GReports','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750');return false;" value="&nbsp;&nbsp;Previsualizar Manifiesto Grupo&nbsp;&nbsp;" class=label></TD>
				<%end if%>
			 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(2)" value="&nbsp;&nbsp;Actualizar&nbsp;&nbsp;" class=label></TD>
			 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(3)" value="&nbsp;&nbsp;Eliminar&nbsp;&nbsp;" class=label></TD>
			<%end if%>
			</TR>
			</TABLE>
		<TD>
		</TR>
	</FORM>
	</TABLE>
<script>
selecciona('forma.Countries','<%=Countries%>');
selecciona('forma.CountryDes','<%=CountryDes%>');
</script>
</BODY>
</HTML>