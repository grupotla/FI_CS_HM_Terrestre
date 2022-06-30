<%
Checking "0|1"
'Dim IATANo, DefaultVal
if CountTableValues >= 0 then
	CreatedDate = ConvertDate(aTableValues(1, 0),2)
    CreatedTime = aTableValues(2, 0)
	Expired = aTableValues(3, 0)
	Name = aTableValues(4, 0)
	Logo = aTableValues(5, 0)
	Footer = aTableValues(6, 0)
	Countries = aTableValues(7, 0)
	Estate = aTableValues(8, 0)
	ArrivalNotes = aTableValues(9, 0)
end if

Set aTableValues = Nothing
%>
<HTML>
<HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
var ntr = "";
var com = "";
	function validar(Action) {
  		if (Action != 3) {
			if (!valSelec(document.forma.Countries)){return (false)};
			if (!valTxt(document.forma.Name, 3, 5)){return (false)};
			if (!valTxt(document.forma.Footer, 3, 5)){return (false)};
			if (!valTxt(document.forma.Estate, 3, 5)){return (false)};
	    }
		document.forma.Action.value = Action;
		document.forma.submit();			 
	 }
</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0"  onLoad="Javascript:self.focus();">
	<TABLE cellspacing=0 cellpadding=2 width=400 align=center>
	<FORM name="forma" action="InsertData.asp" method="post">
	<INPUT name="Action" type=hidden value=0>
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
	<INPUT name="OID" type=hidden value="<%=ObjectID%>">
	<INPUT name="CD" type=hidden value="<%=CreatedDate%>">
	<INPUT name="CT" type=hidden value="<%=CreatedTime%>">
		<TR><TD class=label align=right><b>Código:</b></TD><TD class=label align=left><%if ObjectID <> 0 then response.write ObjectID End if%></TD></TR> 
		<TR><TD class=label align=right><b>Fecha Creación:</b></TD><TD class=label align=left><%=CreatedDate%></TD></TR> 
        <TR><TD class=label align=right><b>Activa:</b></TD><TD class=label align=left><INPUT name=Expired TYPE=checkbox class=label <%If Expired = 0 Then response.write " checked"  End If%>></TD></TR>
		<TR><TD class=label align=right><b>Pais:</b></TD><TD class=label align=left colspan=2>
			<select name="Countries" id="Pais" class="label">
				<option value="-1">Seleccionar</option>
				<%DisplayCountries Countries, 2%>
			</select>	
		</TD></TR>
		<TR><TD class=label align=right><b>Logotipo&nbsp;(Encabezado):</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Logo" id="Encabezado" value="<%=Logo%>" maxlength="60" size="50"></TD></TR> 
		<TR><TD class=label align=right><b>Nombre:</TD><TD class=label align=left><INPUT TYPE=text class=label name="Name" id="Nombre" value="<%=Name%>" maxlength="60" size="50"></TD></TR>
		<TR><TD class=label align=right><b>Pie&nbsp;de&nbsp;Página:</TD><TD class=label align=left><TEXTAREA class=label wrap="off" cols="100" rows="6" name="Footer" id="Pie de Página"><%=Footer%></TEXTAREA></TD></TR>
		<TR><TD class=label align=right><b>Hacienda&nbsp;(Impuestos):</TD><TD class=label align=left><TEXTAREA class=label wrap="off" cols="100" rows="6" name="Estate" id="Hacienda (Impuestos)"><%=Estate%></TEXTAREA></TD></TR>
		<TR><TD class=label align=right><b>Observaciones para Nota Arribo:</TD><TD class=label align=left><TEXTAREA class=label wrap="off" cols="100" rows="6" name="ArrivalNotes" id="Observaciones Nota Arribo"><%=ArrivalNotes%></TEXTAREA></TD></TR>
		<TR>
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
	</FORM>
	</TABLE>
</BODY>
</HTML>