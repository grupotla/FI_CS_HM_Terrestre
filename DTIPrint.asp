<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"
Dim Conn, rs, BLID, CID, CAID, AID, SEP, BLType, CountTableValues1, CountTableValues2, CountPositionValues
Dim QueryPositions, QuerySelect1, QuerySelect2, QuerySelect3, aTableValues1, aTableValues2, Countries, aPositionValues, i

	BLID = CheckNum(Request("OID"))
	CID = CheckNum(Request("CID"))
	CAID = CheckNum(Request("CAID"))
	AID = CheckNum(Request("AID"))
	SEP = CheckNum(Request("SEP"))
	BLType = CheckNum(Request("AT"))
	Countries = Request("CTR")
	CountPositionValues = -1
	CountTableValues1 = -1
	CountTableValues2 = -1
	
	QueryPositions = "select hor_SenderData, a.ver_SenderData, a.hor_BrokerName, a.ver_BrokerName, a.hor_BLExitDate, a.ver_BLExitDate, " & _
			"a.hor_ConsignerData, a.ver_ConsignerData, a.hor_ProviderName, a.ver_ProviderName, a.hor_CodProv, a.ver_CodProv, " & _
			"a.hor_PilotName, a.ver_PilotName, a.hor_PilotPassport, a.ver_PilotPassport, a.hor_PilotCountries, a.ver_PilotCountries, a.hor_PilotLicense, " & _
			"a.ver_PilotLicense, a.hor_CountryDep, a.ver_CountryDep, a.hor_CountryDes, a.ver_CountryDes, a.hor_TruckNo, a.ver_TruckNo, " & _
			"a.hor_TruckCountries, a.ver_TruckCountries, a.hor_TruckAxes, a.ver_TruckAxes, a.hor_TruckTara, a.ver_TruckTara, a.hor_TruckMark, a.ver_TruckMark, " & _
			"a.hor_TruckMotor, a.ver_TruckMotor, a.hor_TruckChassis, a.ver_TruckChassis, a.hor_ContainerTruckNo, " & _
			"a.ver_ContainerTruckNo, a.hor_ContainerCountries, a.ver_ContainerCountries, a.hor_ContainerAxes, a.ver_ContainerAxes, " & _
			"a.hor_ContainerTara, a.ver_ContainerTara, a.hor_DiceContener, a.ver_DiceContener, a.hor_TotNoOfPieces, a.ver_TotNoOfPieces, " & _
			"a.hor_ContactSignature, a.ver_ContactSignature, a.ObservationsWidth, a.DiceContenerWidth, a.ArancelWidth, a.WeightsWidth, a.DiceContenerValueWidth "
		if BLType=1 then
			QueryPositions = QueryPositions & "from DTITemplates a where a.Countries='" & Countries & "'"'Cuando es Demo
		else
			QueryPositions = QueryPositions & "from DTITemplates a, BLs b where a.Countries=b.Countries and b.BLID=" & BLID 'Cuando es un DTI Real
		end if			

	if CID=0 then
		QuerySelect1 = "select a.ShipperData, b.Name, b.Countries, a.BLExitDate, a.ConsignerData, d.Name, d.CodProv, " & _
				"c.Name, c.Passport, c.Countries, c.License, a.CountryDep, a.CountryDes, e.TruckNo, e.Countries, e.Axes, e.Tara, " & _
				"e.Mark, e.Motor, e.Chassis, f.TruckNo, f.Countries, f.Axes, f.Tara, a.TotNoOfPieces, a.TotWeight, " & _
				"a.TotDiceContenerValue, a.ContactSignature, a.Countries, a.DTIObservations, a.Comment4 " & _
				"from (((((BLs a left outer join Trucks f on a.Container = f.TruckID) " & _
				"inner join Brokers b on a.BrokerID=b.BrokerID) inner join Pilots c on a.PilotID=c.PilotID) " & _
				"inner join Providers d on c.ProviderID=d.ProviderID) inner join Trucks e on a.TruckID=e.TruckID) " & _
				"where BLID=" & BLID
				
		QuerySelect2 = "select a.CommoditiesID, a.NoOfPieces, a.ClassNoOfPieces, a.DiceContener, a.Weights, a.DiceContenerValue, '', '', '' " & _
				"from BLDetail a where a.BLID=" & BLID
	else
		QuerySelect1 = "select g.Agents, b.Name, b.Countries, a.BLExitDate, g.Clients, d.Name, d.CodProv, " & _
				"c.Name, c.Passport, c.Countries, c.License, a.CountryDep, a.CountryDes, e.TruckNo, e.Countries, e.Axes, e.Tara, " & _
				"e.Mark, e.Motor, e.Chassis, f.TruckNo, f.Countries, f.Axes, f.Tara, sum(g.NoOfPieces), sum(g.Weights), " & _
				"sum(DiceContenerValue), a.ContactSignature, a.Countries, g.DTIObservations, g.Comment4 " & _
				"from ((((((BLs a left outer join Trucks f on a.Container = f.TruckID) " & _
				"inner join Brokers b on a.BrokerID=b.BrokerID) inner join Pilots c on a.PilotID=c.PilotID) " & _
				"inner join Providers d on c.ProviderID=d.ProviderID) inner join Trucks e on a.TruckID=e.TruckID) " & _
				"inner join BLDetail g on g.BLID=a.BLID) " & _
				"where a.BLID=" & BLID & " and g.ClientsID=" & CID & " and g.AgentsID=" & AID & " and g.Seps=" & SEP & " group by a.BLID, g.ClientsID"
				
		QuerySelect2 = "select a.CommoditiesID, a.NoOfPieces, a.ClassNoOfPieces, a.DiceContener, a.Weights, a.DiceContenerValue, a.AgentsID, a.AgentsAddrID, '' " & _
				"from BLDetail a where a.BLID=" & BLID & " and ClientsID=" & CID & " and AgentsID=" & AID & " and Seps=" & SEP
	end if

	'Posiciones 		
	'"0 a.SenderData, 1 b.Name, 2 b.Countries, 3 a.BLExitDate, 4 a.ConsignerData, 5 d.Name, 6 d.CodProv, "
	'"7 c.Name, 8 c.Passport, 9 c.Countries, 10 c.License, 11 a.CountryDep, 12 a.CountryDes, 13 e.TruckNo, 14 e.Countries, 15 e.Axes, 16 e.Tara, "
	'"17 e.Mark, 18 e.Motor, 19 e.Chassis, 20 f.TruckNo, 21 f.Countries, 22 f.Axes, 23 f.Tara, 24 a.TotNoOfPieces, 25 a.TotWeight: 
	'"26 a.TotDiceContenerValue, 27 a.ContactSignature, 28 a.Countries, 29 a.DTIObservations, 30 a.Comment4 "
	
	OpenConn Conn
	'Obteniendo las posiciones de impresion del DTI
	Set rs = Conn.Execute(QueryPositions)
	If Not rs.EOF Then
   		aPositionValues = rs.GetRows
   		CountPositionValues = rs.RecordCount-1
    End If
   	closeOBJ rs	
	
	if CountPositionValues >= 0 then

		select case BLType
		case 0
			'Obteniendo los datos Generales del DTI
			Set rs = Conn.Execute(QuerySelect1)
			If Not rs.EOF Then
				aTableValues1 = rs.GetRows
				CountTableValues1 = rs.RecordCount-1
			End If
			closeOBJ rs
			'Obteniendo el detalle del DTI
			Set rs = Conn.Execute(QuerySelect2)
			If Not rs.EOF Then
				aTableValues2 = rs.GetRows
				CountTableValues2 = rs.RecordCount-1
			End If	
			closeOBJs rs, Conn
			
			'Obteniendo los Incisos Arancelarios en la base Postgress correspondientes a cada Producto
			if CountTableValues1>=0 then
				if CountTableValues2>=0 then
					OpenConn2 Conn
						for i=0 to CountTableValues2
							Set rs = Conn.Execute("select arancel_" & SetArancel(aTableValues1(28,0)) & " from commodities where commodityid=" & aTableValues2(0,i))
							If Not rs.EOF Then
								aTableValues2(8,i) = rs(0)
							End If
							CloseOBJ rs
						next						
						
						if CID <> 0 then 'Cuando es DTI Individual se busca la informacion del Consignatario y Exportador Individual
							'Obteniendo los datos del Consignatario Individual
							QuerySelect3 = "select a.nombre_cliente, d.direccion_completa, d.phone_number " & _
													"from clientes a, direcciones d, niveles_geograficos n, paises p " & _
													" where a.id_cliente = d.id_cliente" & _
													" and d.id_nivel_geografico = n.id_nivel" & _
													" and n.id_pais = p.codigo" & _
													" and a.id_cliente = " & CID
							if CAID <> 0 then
								QuerySelect3 = QuerySelect3 & " and d.id_direccion = " & CAID
							end if
							
							'response.write QuerySelect3 & "<br>"
							Set rs = Conn.Execute(QuerySelect3)
							if Not rs.EOF then
								aTableValues1(4,0) = aTableValues1(4,0) & "<br>" & rs(1) & "<br>"
								if rs(2) <> "" then
									aTableValues1(4,0) = aTableValues1(4,0) & rs(2)
								end if
							end if
							CloseOBJ rs
						
							set rs = Conn.Execute("select numero_telefono from cli_telefonos where id_cliente=" & CID)
							if Not rs.EOF then
								aTableValues1(4,0) = aTableValues1(4,0) & "    " & rs(0)
							end if
							CloseOBJ rs
							'set rs = Conn.Execute("select nombres from contactos where id_cliente=" & CID)
                            set rs = Conn.Execute("select nombres from contactos where id_cliente=" & CID & " and (tipo_persona = 'Principal' or tipo_persona IS NULL) and activo = true")
							if Not rs.EOF then
								aTableValues1(4,0) = aTableValues1(4,0) & "<br>ATTN:" & rs(0)
							end if

							'Obteniendo los datos del Exportador Individual
							QuerySelect3 = "select a.nombre_cliente, d.direccion_completa, d.phone_number " & _
													"from clientes a, direcciones d, niveles_geograficos n, paises p " & _
													" where a.id_cliente = d.id_cliente" & _
													" and d.id_nivel_geografico = n.id_nivel" & _
													" and n.id_pais = p.codigo" & _
													" and a.id_cliente = " & aTableValues2(6,0)
							if CAID <> 0 then
								QuerySelect3 = QuerySelect3 & " and d.id_direccion = " & aTableValues2(7,0)
							end if
							
							'response.write QuerySelect3 & "<br>"
							Set rs = Conn.Execute(QuerySelect3)
							if Not rs.EOF then
								aTableValues1(0,0) = aTableValues1(0,0) & "<br>" & rs(1) & "<br>"
								if rs(2) <> "" then
									aTableValues1(0,0) = aTableValues1(0,0) & rs(2)
								end if
							end if
							CloseOBJ rs
						
							set rs = Conn.Execute("select numero_telefono from cli_telefonos where id_cliente=" & aTableValues2(6,0))
							if Not rs.EOF then
								aTableValues1(0,0) = aTableValues1(0,0) & "    " & rs(0)
							end if
							CloseOBJ rs
                            
							'set rs = Conn.Execute("select nombres from contactos where id_cliente=" & aTableValues2(6,0))
                            set rs = Conn.Execute("select nombres from contactos where id_cliente=" & aTableValues2(6,0) & " and (tipo_persona = 'Principal' or tipo_persona IS NULL) and activo = true")
							if Not rs.EOF then
								aTableValues1(0,0) = aTableValues1(0,0) & "<br>ATTN:" & rs(0)
							end if



						end if
					CloseOBJs rs, Conn
				end if
			end if
		case 1
			CountTableValues1 = 0
			CountTableValues2 = 0
			Redim aTableValues1(30,0)
			Redim aTableValues2(5,0)
			Redim aArancelValues(0,0)
			aTableValues1(0,0)="1.EXPORTADOR / EMBARCADOR/ REMITENTE"
			aTableValues1(1,0)="2.ADUANA DE PARTIDA"
			aTableValues1(2,0)="XX"
			aTableValues1(3,0)="5.FECHA DE ACEPTACION"
			aTableValues1(4,0)="6.CONSIGNATARIO"
			aTableValues1(5,0)="7.TRANSPORTISTA"
			aTableValues1(6,0)="8.CODIGO TRANSPORTISTA"
			aTableValues1(7,0)="9.NOMBRE DEL CONDUCTOR"
			aTableValues1(8,0)="10.PASAPORTE"
			aTableValues1(9,0)="11.PAIS"
			aTableValues1(10,0)="12.LICENCIA"
			aTableValues1(11,0)="14.PAIS PROCEDENCIA"
			aTableValues1(12,0)="15.PAIS DESTINO"
			aTableValues1(13,0)="17.MATRICULA"
			aTableValues1(14,0)="18.PAIS"
			aTableValues1(15,0)="19.EJES"
			aTableValues1(16,0)="20.TARA"
			aTableValues1(17,0)="21.MARCA"
			aTableValues1(18,0)="22.MOTOR"
			aTableValues1(19,0)="23.CHASIS"
			aTableValues1(20,0)="24.MATRICULA"
			aTableValues1(21,0)="25.PAIS"
			aTableValues1(22,0)="26.EJES"
			aTableValues1(23,0)="27.TARA"
			aTableValues1(25,0)="PESO BRUTO TOTAL"
			aTableValues1(27,0)="NOMBRE Y FIRMA DEL TRANSPORTISTA"
			aTableValues1(29,0)="28.MARCA DE EXPEDICION, # CONTENEDOR, DIMENSIONES"

			aTableValues2(1,0)=""
			aTableValues2(3,0)="29.NUMERO Y CLASE DE BULTOS, DESCRIPCION DE LAS MERCANCIAS"
			aTableValues2(4,0)="31.PESO BRUTO DE LAS MERCANCIAS"
			aTableValues2(5,0)="32.VALOR $C.A."
			
			aArancelValues(0,0)="30.INCISO ARANCELARIO DE LAS MERCANCIAS"
		end select
			
		if CountTableValues1>=0 then
%>
<html>
<style type="text/css">
<!--
body {
	margin:0px;
}
.style11 {
	font-size:11px;
	color: #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: normal;
}
.style1 {
	font-size:11px;
	color: #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: normal;
}
-->
</style>
<body onLoad="JavaScript:self.focus();">

<!--SenderData-->
<DIV style='LEFT: <%=aPositionValues(0,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(1,0)%>px; WIDTH:<%=aPositionValues(2,0)-aPositionValues(0,0)-25%>;' class='style11'><%=FRegExp(chr(13) & chr(10), aTableValues1(0,0), "<br>", 4)%></DIV>
<!--BrokerName-->
<DIV style='LEFT: <%=aPositionValues(2,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(3,0)%>px;' class='style11'><%=aTableValues1(1,0)%><br>/&nbsp;<%=TranslateCountry(aTableValues1(2,0))%></DIV>
<!--BLExitDate-->
<DIV style='LEFT: <%=aPositionValues(4,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(5,0)%>px;' class='style11'><%=aTableValues1(3,0)%></DIV>
<!--ConsignerData-->
<DIV style='LEFT: <%=aPositionValues(6,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(7,0)%>px; WIDTH:<%=aPositionValues(8,0)-aPositionValues(6,0)-25%>;' class='style11'><%=FRegExp(chr(13) & chr(10), aTableValues1(4,0), "<br>", 4)%></DIV>
<!--ProviderName-->
<DIV style='LEFT: <%=aPositionValues(8,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(9,0)%>px;' class='style11'><%=aTableValues1(5,0)%></DIV>
<!--CodProv-->
<DIV style='LEFT: <%=aPositionValues(10,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(11,0)%>px;' class='style11'><%=aTableValues1(6,0)%></DIV>
<!--PilotName-->
<DIV style='LEFT: <%=aPositionValues(12,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(13,0)%>px;' class='style11'><%=aTableValues1(7,0)%></DIV>
<!--PilotPassport-->
<DIV style='LEFT: <%=aPositionValues(14,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(15,0)%>px;' class='style11'><%=aTableValues1(8,0)%></DIV>
<!--PilotCountries-->
<DIV style='LEFT: <%=aPositionValues(16,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(17,0)%>px;' class='style11'><%=aTableValues1(9,0)%></DIV>
<!--PilotLicense-->
<DIV style='LEFT: <%=aPositionValues(18,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(19,0)%>px;' class='style11'><%=aTableValues1(10,0)%>&nbsp;&nbsp;<%=aTableValues1(9,0)%></DIV>
<!--CountryDep-->
<DIV style='LEFT: <%=aPositionValues(20,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(21,0)%>px;' class='style11'><%=TranslateCountry(aTableValues1(11,0))%></DIV>
<!--CountryDes-->
<DIV style='LEFT: <%=aPositionValues(22,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(23,0)%>px;' class='style11'><%=TranslateCountry(aTableValues1(12,0))%></DIV>
<!--TruckNo-->
<DIV style='LEFT: <%=aPositionValues(24,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(25,0)%>px;' class='style11'><%=aTableValues1(13,0)%></DIV>
<!--TruckCountries-->
<DIV style='LEFT: <%=aPositionValues(26,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(27,0)%>px;' class='style11'><%=aTableValues1(14,0)%></DIV>
<!--TruckAxes-->
<DIV style='LEFT: <%=aPositionValues(28,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(29,0)%>px;' class='style11'><%=aTableValues1(15,0)%></DIV>
<!--TruckTara-->
<DIV style='LEFT: <%=aPositionValues(30,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(31,0)%>px;' class='style11'><%=aTableValues1(16,0)%></DIV>
<!--TruckMark-->
<DIV style='LEFT: <%=aPositionValues(32,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(33,0)%>px;' class='style11'><%=aTableValues1(17,0)%></DIV>
<!--TruckMotor-->
<DIV style='LEFT: <%=aPositionValues(34,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(35,0)%>px;' class='style11'><%=aTableValues1(18,0)%></DIV>
<!--TruckChassis-->
<DIV style='LEFT: <%=aPositionValues(36,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(37,0)%>px;' class='style11'><%=aTableValues1(19,0)%></DIV>
<!--ContainerTruckNo-->
<DIV style='LEFT: <%=aPositionValues(38,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(39,0)%>px;' class='style11'><%=aTableValues1(20,0)%></DIV>
<!--ContainerCountries-->
<DIV style='LEFT: <%=aPositionValues(40,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(41,0)%>px;' class='style11'><%=aTableValues1(21,0)%></DIV>
<!--ContainerAxes-->
<DIV style='LEFT: <%=aPositionValues(42,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(43,0)%>px;' class='style11'><%=aTableValues1(22,0)%></DIV>
<!--ContainerTara-->
<DIV style='LEFT: <%=aPositionValues(44,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(45,0)%>px;' class='style11'><%=aTableValues1(23,0)%></DIV>
<!--DiceContener-->
<DIV style='LEFT: <%=aPositionValues(46,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(47,0)%>px;' class='style11'>
	<table cellpadding="3" cellspacing="0" border="<%=BLType%>">
	<tr>
		<td class='style1' rowspan="<%=CountTableValues2+2%>" width="<%=aPositionValues(52,0)%>" valign="top"><%=aTableValues1(29,0)%></td>
	</tr>
	<%for i=0 to CountTableValues2%>
	<tr>
		<td class='style1' valign="top"><%=aTableValues2(1,i)%>&nbsp;<%if Trim(aTableValues2(2,i)) <> "" then response.write aTableValues2(2,i)%></td>
		<td class='style1' valign="top" width="<%=aPositionValues(53,0)%>"><%=aTableValues2(3,i)%></td>
		<td class='style1' valign="top" width="<%=aPositionValues(54,0)%>">
		<%if aTableValues2(8,i)<>"" then 
			response.write aTableValues2(8,i) 
		else
			response.write "<script>alert('El producto """ & aTableValues2(3,i) & """ no tiene Aranceles Asignados');</script>"
		end if%></td>
		<td class='style1' valign="top" width="<%=aPositionValues(55,0)%>" align="right"><%=aTableValues2(4,i)%></td>
		<td class='style1' valign="top" width="<%=aPositionValues(56,0)%>" align="right">$&nbsp;<%=aTableValues2(5,i)%></td>
	</tr>
	<%next%>
	<tr>
		<td class='style11' colspan="6">&nbsp;</td>
	</tr>
	<tr>
		<td class='style1'><b>TOTAL</b></td>
		<td class='style1'><b><%=aTableValues1(24,0)%>&nbsp;BLTS.</b></td>
		<td class='style1' colspan="2">&nbsp;</td>
		<td class='style1' align="right"><b><%=aTableValues1(25,0)%>&nbsp;KGS.</b></td>
		<td class='style1' align="right"><b>$&nbsp;<%=aTableValues1(26,0)%></b></td>
	</tr>
	<tr>
		<td class='style11' colspan="6"><br><br><%=aTableValues1(30,0)%></td>
	</tr>
	</table>
</DIV>
<!--TotNoOfPieces-->
<DIV style='LEFT: <%=aPositionValues(48,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(49,0)%>px;' class='style11'><%=aTableValues1(25,0)%> Kgs.</DIV>
<!--ContactSignature-->
<DIV style='LEFT: <%=aPositionValues(50,0)%>px; POSITION: absolute; TOP: <%=aPositionValues(51,0)%>px;' class='style11'><%=aTableValues1(7,0)%></DIV>
</body>
</html>
<%
		Else
%>
			<script>alert("Algun dato esta incompleto, puede ser que el Piloto o Transporte no tenga Asignado un Proveedor");</script>
<%
		End if 
		Set aTableValues1 = Nothing
		Set aTableValues2 = Nothing
		Set aPositionValues = Nothing		
	Else
		CloseOBJ Conn
%>
	<script>alert("No Existe plantilla");</script>
<%
	End if
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>