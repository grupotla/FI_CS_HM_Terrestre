<%
Checking "0|1"
CountList1Values=-1
if CountTableValues >= 0 then
	CreatedDate = ConvertDate(aTableValues(1, 0),2)
    CreatedTime = aTableValues(2, 0)
	Expired = aTableValues(3, 0)
	WarehouseID = aTableValues(4, 0)
	StartValue = aTableValues(5, 0)
	ActualValue = aTableValues(6, 0)
	FinishValue = aTableValues(7, 0)
	BagValue = aTableValues(8, 0)	
	UserID = aTableValues(9, 0)
end if
	
	OpenConn Conn
	'listado de bodegas para asignar bolsas de marchamos
	Set rs = Conn.Execute("select WarehouseID, Countries, Name from Warehouses where Expired=0 order by Countries, Name")
	If Not rs.EOF Then
   		aList1Values = rs.GetRows
       	CountList1Values = rs.RecordCount-1
    End If
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
			if (!valSelec(document.forma.WarehouseID)){return (false)};		
			if (!valTxt(document.forma.BagValue, 1, 5)){return (false)};
			if (!valTxt(document.forma.StartValue, 3, 5)){return (false)};
			if (!valTxt(document.forma.FinishValue, 3, 5)){return (false)};
		}
	    document.forma.Action.value = Action;
		document.forma.submit();			 
	 }
</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<style type="text/css">
<!--
body {
	margin: 0px;
}
.style8 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 10px;
	font-weight: bold;
	color: #999999;
}
.style10 {font-size: 10px; font-family: Verdana, Arial, Helvetica, sans-serif; }
-->
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
		<TR><TD class=label align=right><b>Bodega:</b></TD><TD class=style4 align=left colspan=2>
		<select class="style10" name="WarehouseID" id="Bodega">
		<option value="-1">Seleccionar</option>
		<%		
			For i = 0 To CountList1Values
		%>
		<option value="<%=aList1Values(0,i)%>"><%response.write aList1Values(1,i) & " - " & aList1Values(2,i)%></option>
		<%
   			Next
		%>
		</select>
		</TD></TR>
		<TR><TD class=label align=right><b>Bolsa:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="BagValue" id="Bolsa" value="<%=BagValue%>" maxlength="60" onKeyUp="res(this,numb);"></TD></TR> 
		<TR><TD class=label align=right><b>Numero Inicial:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="StartValue" id="Numero Inicial" value="<%=StartValue%>" maxlength="60" onKeyUp="res(this,numb);"></TD></TR>
		<TR><TD class=label align=right><b>Numero Final:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="FinishValue" id="Numero Final" value="<%=FinishValue%>" maxlength="60" onKeyUp="res(this,numb);"></TD></TR>
		<TR><TD class=label align=right><b>Proximo Numero:</b></TD><TD class=label align=left><INPUT TYPE=text class=label value="<%=ActualValue%>" maxlength="45" readonly></TD></TR>
		<TD colspan="2" class=label align=center>
			<TABLE cellspacing=0 cellpadding=2 width=200>
				<TR>
				<%if CountTableValues = -1 then%>
					<TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(1)" value="&nbsp;&nbsp;Agregar&nbsp;&nbsp;" class=label></TD>
				<%else%>
					<TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(2)" value="&nbsp;&nbsp;Actualizar&nbsp;&nbsp;" class=label></TD>
					<TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(3)" value="&nbsp;&nbsp;Eliminar&nbsp;&nbsp;" class=label></TD>
				<%end if%>
				</TR>
			</TABLE>
		<TD>
		</TR>
	</TABLE>
	</FORM>
<script>
selecciona('forma.WarehouseID','<%=WarehouseID%>');
</script>
</BODY>
</HTML>