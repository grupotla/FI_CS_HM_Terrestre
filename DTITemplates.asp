<%
Checking "0|1|2"
'Dim TaxNo, Address, Phone1, Phone2, AccountNo, Attn, Expired
if CountTableValues >= 0 then
	CreatedDate = ConvertDate(aTableValues(1, 0),2)
    CreatedTime = aTableValues(2, 0)
	Expired = aTableValues(3, 0)
	Name = aTableValues(4, 0)
	Countries = aTableValues(5, 0)
end if

%>
<HTML>
<HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
	function validar(Action) {
	   	if (!valTxt(document.forma.Name, 3, 5)){return (false)};
		if (!valSelec(document.forma.Countries)){return (false)};
		document.forma.Action.value = Action;
		document.forma.submit();			 
	 }
</script>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
	<TABLE cellspacing=0 cellpadding=2 width=500 align=center>
	<FORM name="forma" action="InsertData.asp" method="post">
	<INPUT name="Action" type=hidden value=0>
	<INPUT name="Expired" type=hidden value="on">
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
	<INPUT name="OID" type=hidden value="<%=ObjectID%>">
	<INPUT name="CD" type=hidden value="<%=CreatedDate%>">
	<INPUT name="CT" type=hidden value="<%=CreatedTime%>">
		<TR><TD class=label align=right><b>Fecha Creación:</b></TD><TD class=label align=left colspan="2"><%=CreatedDate%></TD></TR> 
		<TR><TD class=label align=right><b>Código:</b></TD><TD class=label align=left colspan="2"><%if ObjectID <> 0 then response.write ObjectID End if%></TD></TR> 
		<TR><TD class=label align=right><b>Nombre:</b></TD><TD class=label align=left colspan="2"><INPUT name="Name" id="Nombre de Plantilla" type=text value="<%=Name%>" size=23 maxLength=45 class=label></TD></TR>
		<TR><TD class=label align=right><b>Pais:</b></TD><TD class=label align=left colspan="2">
			<select name="Countries" id="Pais" class="label">
				<option value="-1">Seleccionar</option>
				<%DisplayCountries Countries, 2%>
			</select>
		</TD></TR>
		<%if CountTableValues>0 then%>
	 	<TR><TD class=label align=right><b>Formato de Plantilla:</b></TD><TD class=label align=center>Horizontal</TD><TD class=label align=center>Vertical</TD></TR>
		<TR><TD class=label align=right>1. Exportador / Embarcador / Remitente</TD><TD class=label align=center><input type=text class=label name='hor_SenderData' id='1. Exportador / Embarcador / Remitente Horizontal' value='<%=aTableValues(6,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_SenderData' id='1. Exportador / Embarcador / Remitente Vertical' value='<%=aTableValues(7,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>2. Aduana de Partida / Pais</TD><TD class=label align=center><input type=text class=label name='hor_BrokerName' id='2. Aduana de Partida / Pais Horizontal' value='<%=aTableValues(8,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_BrokerName' id='2. Aduana de Partida / Pais Vertical' value='<%=aTableValues(9,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>5. Fecha de Aceptacion</TD><TD class=label align=center><input type=text class=label name='hor_BLExitDate' id='5. Fecha de Aceptacion Horizontal' value='<%=aTableValues(10,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_BLExitDate' id='5. Fecha de Aceptacion Vertical' value='<%=aTableValues(11,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>6. Consignatario</TD><TD class=label align=center><input type=text class=label name='hor_ConsignerData' id='6. Consignatario Horizontal' value='<%=aTableValues(12,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_ConsignerData' id='6. Consignatario Vertical' value='<%=aTableValues(13,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>7. Transportista</TD><TD class=label align=center><input type=text class=label name='hor_ProviderName' id='7. Transportista Horizontal' value='<%=aTableValues(14,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_ProviderName' id='7. Transportista Vertical' value='<%=aTableValues(15,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>8. Codigo Transportista</TD><TD class=label align=center><input type=text class=label name='hor_CodProv' id='8. Codigo Transportista Horizontal' value='<%=aTableValues(16,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_CodProv' id='8. Codigo Transportista Vertical' value='<%=aTableValues(17,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>9. Nombre del Conductor</TD><TD class=label align=center><input type=text class=label name='hor_PilotName' id='9. Nombre del Conductor Horizontal' value='<%=aTableValues(18,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_PilotName' id='9. Nombre del Conductor Vertical' value='<%=aTableValues(19,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>10. Pasaporte del Conductor</TD><TD class=label align=center><input type=text class=label name='hor_PilotPassport' id='10. Pasaporte del Conductor Horizontal' value='<%=aTableValues(20,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_PilotPassport' id='10. Pasaporte del Conductor Vertical' value='<%=aTableValues(21,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>11. Pais del Conductor</TD><TD class=label align=center><input type=text class=label name='hor_PilotCountries' id='11. Pais del Conductor Horizontal' value='<%=aTableValues(22,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_PilotCountries' id='11. Pais del Conductor Vertical' value='<%=aTableValues(23,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>12. Licencia del Conductor</TD><TD class=label align=center><input type=text class=label name='hor_PilotLicense' id='12. Licencia del Conductor Horizontal' value='<%=aTableValues(24,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_PilotLicense' id='12. Licencia del Conductor Vertical' value='<%=aTableValues(25,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>14. Pais de Procedencia</TD><TD class=label align=center><input type=text class=label name='hor_CountryDep' id='14. Pais de Procedencia Horizontal' value='<%=aTableValues(26,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_CountryDep' id='14. Pais de Procedencia Vertical' value='<%=aTableValues(27,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>15.  Pais de Destino</TD><TD class=label align=center><input type=text class=label name='hor_CountryDes' id='15.  Pais de Destino Horizontal' value='<%=aTableValues(28,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_CountryDes' id='15.  Pais de Destino Vertical' value='<%=aTableValues(29,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>17. Matricula del Transporte</TD><TD class=label align=center><input type=text class=label name='hor_TruckNo' id='17. Matricula del Transporte Horizontal' value='<%=aTableValues(30,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_TruckNo' id='17. Matricula del Transporte Vertical' value='<%=aTableValues(31,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>18. Pais de Registro del Transporte</TD><TD class=label align=center><input type=text class=label name='hor_TruckCountries' id='18. Pais de Registro del Transporte Horizontal' value='<%=aTableValues(32,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_TruckCountries' id='18. Pais de Registro del Transporte Vertical' value='<%=aTableValues(33,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>19. No. Ejes del Transporte</TD><TD class=label align=center><input type=text class=label name='hor_TruckAxes' id='19. No. Ejes del Transporte Horizontal' value='<%=aTableValues(34,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_TruckAxes' id='19. No. Ejes del Transporte Vertical' value='<%=aTableValues(35,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>20. Tara del Transporte</TD><TD class=label align=center><input type=text class=label name='hor_TruckTara' id='20. Tara del Transporte Horizontal' value='<%=aTableValues(36,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_TruckTara' id='20. Tara del Transporte Vertical' value='<%=aTableValues(37,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>21. Marca del Transporte</TD><TD class=label align=center><input type=text class=label name='hor_TruckMark' id='21. Marca del Transporte Horizontal' value='<%=aTableValues(38,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_TruckMark' id='21. Marca del Transporte Vertical' value='<%=aTableValues(39,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>22. Motor del Transporte</TD><TD class=label align=center><input type=text class=label name='hor_TruckMotor' id='22. Motor del Transporte Horizontal' value='<%=aTableValues(40,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_TruckMotor' id='22. Motor del Transporte Vertical' value='<%=aTableValues(41,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>23. Chasis del Transporte</TD><TD class=label align=center><input type=text class=label name='hor_TruckChassis' id='23. Chasis del Transporte Horizontal' value='<%=aTableValues(42,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_TruckChassis' id='23. Chasis del Transporte Vertical' value='<%=aTableValues(43,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>24. Matricula del Remolque</TD><TD class=label align=center><input type=text class=label name='hor_ContainerTruckNo' id='24. Matricula del Remolque Horizontal' value='<%=aTableValues(44,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_ContainerTruckNo' id='24. Matricula del Remolque Vertical' value='<%=aTableValues(45,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>25. Pais de Registro del Remolque</TD><TD class=label align=center><input type=text class=label name='hor_ContainerCountries' id='25. Pais de Registro del Remolque Horizontal' value='<%=aTableValues(46,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_ContainerCountries' id='25. Pais de Registro del Remolque Vertical' value='<%=aTableValues(47,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>26. No. Ejes del Remolque</TD><TD class=label align=center><input type=text class=label name='hor_ContainerAxes' id='26. No. Ejes del Remolque Horizontal' value='<%=aTableValues(48,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_ContainerAxes' id='26. No. Ejes del Remolque Vertical' value='<%=aTableValues(49,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>27. Tara del Remolque</TD><TD class=label align=center><input type=text class=label name='hor_ContainerTara' id='27. Tara del Remolque Horizontal' value='<%=aTableValues(50,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_ContainerTara' id='27. Tara del Remolque Vertical' value='<%=aTableValues(51,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>28, 29, 30, 31 y 32 Detalle de las Mercancias</TD><TD class=label align=center><input type=text class=label name='hor_DiceContener' id='29. Numero y Clase de Bultos, Descripcion de las Mercancias Horizontal' value='<%=aTableValues(52,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_DiceContener' id='29. Numero y Clase de Bultos, Descripcion de las Mercancias Vertical' value='<%=aTableValues(53,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>Peso Bruto Total (Kg)</TD><TD class=label align=center><input type=text class=label name='hor_TotNoOfPieces' id='Peso Bruto Total (Kg) Horizontal' value='<%=aTableValues(54,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_TotNoOfPieces' id='Peso Bruto Total (Kg) Vertical' value='<%=aTableValues(55,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>33. Nombre y Firma del Transportista</TD><TD class=label align=center><input type=text class=label name='hor_ContactSignature' id='33. Nombre y Firma del Transportista Horizontal' value='<%=aTableValues(56,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD>
		<TD class=label align=center><input type=text class=label name='ver_ContactSignature' id='33. Nombre y Firma del Transportista Vertical' value='<%=aTableValues(57,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
	 	<TR><TD class=label align=right><b>Anchos para Detalle DTI:</b></TD><TD class=label align=center>Ancho de Columna</TD><TD class=label align=center>&nbsp;</TD></TR>
		<TR><TD class=label align=right>28. Marca de expedicion, No.Contenedor, dimensiones</TD><TD class=label align=left colspan="2"><input type=text class=label name='ObservationsWidth' id='28. Marca de expedicion, No.Contenedor, dimensiones' value='<%=aTableValues(58,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>29. Numero y Clase de Bultos, Descripcion de las Mercancias</TD><TD class=label align=left colspan="2"><input type=text class=label name='DiceContenerWidth' id='29. Numero y Clase de Bultos, Descripcion de las Mercancias' value='<%=aTableValues(59,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>30. Incio Arancelario de las Mercancias</TD><TD class=label align=left colspan="2"><input type=text class=label name='ArancelWidth' id='30. Inciso Arancelario de las Mercancias' value='<%=aTableValues(60,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>31. Peso Bruto de la Mercancia</TD><TD class=label align=left colspan="2"><input type=text class=label name='WeightsWidth' id='31. Peso Bruto de la Mercancia' value='<%=aTableValues(61,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<TR><TD class=label align=right>32. Valor $C.A.</TD><TD class=label align=left colspan="2"><input type=text class=label name='DiceContenerValueWidth' id='32. Valor $C.A.' value='<%=aTableValues(62,0)%>' maxlength='5' onKeyUp='res(this,numb);'></TD></TR>
		<%end if%>
		<TD colspan="3" class=label align=center>
			<TABLE cellspacing=0 cellpadding=2 width=200>
			<TR>
				 <%if CountTableValues<0 then%>	
				 <TD class=label align=left colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(1)" value="&nbsp;&nbsp;Agregar&nbsp;&nbsp;" class=label></TD>
				 <%else%>
				 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="Javascript:window.open('DTIPrint.asp?GID=<%=GroupID%>&OID=<%=ObjectID%>&CTR=<%=Countries%>&AT=1','DTI','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750,height=600');return false;" value="&nbsp;&nbsp;Previsualizar&nbsp;DTI&nbsp;&nbsp;" class=label></TD>
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
</script>	
</BODY>
</HTML>
<%Set aTableValues = Nothing%>