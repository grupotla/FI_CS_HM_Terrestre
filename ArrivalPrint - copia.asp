<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"
Dim Conn, rs, i, j, Sep, Action, BLID, ObjectID, ClientID, AgentID, SBLIDS, ConsignerData, QuerySelect, com, IsColoader, ColoaderID, ShipperID, query
Dim CountTableValues, aTableValues, CountTableValues2, aTableValues2, CountTableValues3, aTableValues3, aTableValues4, CountTableValues4
Dim TotUSDCurrencyPrepaid, TotOtherCurrenciesPrepaid, TotUSDCurrencyCollect, TotOtherCurrenciesCollect, ResultEXDBCountry, ClientConsult, AddrConsult

	ObjectID = CheckNum(Request("OID"))
	ClientID = CheckNum(Request("CID"))
	AgentID = CheckNum(Request("AID"))
	Sep = CheckNum(Request("SEP"))
    com  = ""
    TotUSDCurrencyPrepaid = 0
    TotOtherCurrenciesPrepaid = 0
	TotUSDCurrencyCollect = 0
    TotOtherCurrenciesCollect = 0
	CountTableValues = -1
	CountTableValues2 = -1
	CountTableValues3 = -1
	CountTableValues4 = -1
    IsColoader = 0
    ClientConsult = 0
	
	OpenConn Conn
	'Obteniendo los datos del Encabezado
	'response.write("select a.ClientsID, a.AddressesID, b.CountryDes, a.HBLNumber, a.LtArrivalDate, b.BLArrivalDate, c.Name, a.LtArrivalDeliveryDocs, b.FinalDes, a.CPDocType, a.ManifestDocType, a.EndorseDocType, a.DTIDocType, a.BLsType, a.BillType, a.EndorseObservations, a.Countries, a.BLs, a.EXDBCountry, a.ColoadersID, a.ShippersID from ((BLDetail a left outer join BLs b on a.BLID=b.BLID) left outer join Warehouses c on b.DestinyType=c.WareHouseID) where a.BLID=" & ObjectID & " and a.ClientsID=" & ClientID & " and a.AgentsID=" & AgentID & " and a.Seps=" & Sep & "<br>")
	Set rs = Conn.Execute("select a.ClientsID, a.AddressesID, b.CountryDes, a.HBLNumber, a.LtArrivalDate, b.BLArrivalDate, c.Name, a.LtArrivalDeliveryDocs, b.FinalDes, a.CPDocType, a.ManifestDocType, a.EndorseDocType, a.DTIDocType, a.BLsType, a.BillType, a.EndorseObservations, a.Countries, a.BLs, a.EXDBCountry, a.ColoadersID, a.ShippersID, a.ColoadersAddrID from ((BLDetail a left outer join BLs b on a.BLID=b.BLID) left outer join Warehouses c on b.DestinyType=c.WareHouseID) where a.BLID=" & ObjectID & " and a.ClientsID=" & ClientID & " and a.AgentsID=" & AgentID & " and a.Seps=" & Sep)
	If not rs.EOF then
    	aTableValues = rs.GetRows
    	CountTableValues = rs.RecordCount-1
	end if
	CloseOBJ rs
	'Obteniendo el detalle de la carga
	'response.write "select BLDetailID, CountryOrigen, Container, BLs, Shippers, NoOfPieces, ClassNoOfPieces, Weights, Volumes from BLDetail where ClientsID=" & ClientID & " and AgentsID=" & AgentID & " and Seps=" & Sep & " order by BLDetailID<br>"
	Set rs = Conn.Execute("select BLDetailID, CountryOrigen, Container, BLs, Shippers, NoOfPieces, ClassNoOfPieces, Weights, Volumes, DiceContener, Countries, CountriesFinalDes from BLDetail where BLID=" & ObjectID & " and ClientsID=" & ClientID & " and AgentsID=" & AgentID & " and Seps=" & Sep & " order by BLDetailID")
	If not rs.EOF then
    	aTableValues2 = rs.GetRows
    	CountTableValues2 = rs.RecordCount-1
	end if
	CloseOBJ rs
    ColoaderID = CheckNum(aTableValues(19, 0)) '2019-08-30
	'Obteniendo los datos generales de la nota de Arribo
    'If ColoaderID = 93049 Or ColoaderID = 77222 Or ColoaderID = 75002 Or ColoaderID = 67728 Or ColoaderID = 65768 Or ColoaderID = 63709 Or ColoaderID = 61920 Or ColoaderID = 43421 Or ColoaderID = 29457 Then 'GTTLA
    '    query = "SELECT a.ArrivalNotes, b.CountryDes, a.Footer, 1 as orden FROM Letters a, BLs b WHERE a.Countries=CONCAT(b.CountryDes,'TLA') AND a.Expired=0 AND b.BLID=" & ObjectID	   
    'else
        query = "SELECT a.ArrivalNotes, b.CountryDes, a.Footer, 1 as orden FROM Letters a, BLs b WHERE a.Countries=b.CountryDes AND a.Expired=0 AND b.BLID=" & ObjectID	
    'end if
    'query = query & " UNION SELECT 'N/A', 'N/A', 'N/A', 2 as orden ORDER BY orden LIMIT 1"
   'response.write query & "<br><br>"
    Set rs = Conn.Execute(query)
	If not rs.EOF then
    	aTableValues3 = rs.GetRows
    	CountTableValues3 = rs.RecordCount-1
	end if
	CloseOBJ rs
	'Obteniendo los rubros de la carga
	for i=0 to CountTableValues2
		SBLIDS = SBLIDS & com & aTableValues2(0,i)
		com = ","
	next
	'response.write("select SBLID, ItemName, Currency, Value, OverSold, PrepaidCollect from ChargeItems where Expired=0 and InterProviderType<>5 and InterChargeType<>2 and SBLID in (" & SBLIDS & ") order by SBLID, PrepaidCollect, Local, Currency, ItemName <br>")
	if SBLIDS <> "" then
		Set rs = Conn.Execute("select SBLID, ItemName, Currency, Value, OverSold, PrepaidCollect from ChargeItems where Expired=0 and InterProviderType<>5 and InterChargeType<>2 and SBLID in (" & SBLIDS & ") order by SBLID, PrepaidCollect, Local, Currency, ItemName")
		If not rs.EOF then
			aTableValues4 = rs.GetRows
			CountTableValues4 = rs.RecordCount-1
		end if
		CloseOBJ rs
	end if
	CloseOBJ Conn
	
	'Obteniendo datos del cliente desde la tabla master
	if aTableValues(0,0) <> 0 then
		if ColoaderID = 0 then
            ClientConsult = CheckNum(aTableValues(0,0))
            AddrConsult = CheckNum(aTableValues(1,0))
        else
            ClientConsult = ColoaderID
            AddrConsult = CheckNum(aTableValues(21,0))
        end if
        OpenConn2 Conn
		QuerySelect = "select a.nombre_cliente, d.direccion_completa, d.phone_number " & _
								"from clientes a, direcciones d, niveles_geograficos n, paises p " & _
								" where a.id_cliente = d.id_cliente" & _
								" and d.id_nivel_geografico = n.id_nivel" & _
								" and n.id_pais = p.codigo" & _
								" and a.id_cliente = " & ClientConsult
		
		if AddrConsult <> 0 then
			QuerySelect = QuerySelect & " and d.id_direccion = " & AddrConsult
		end if
	
		'response.write QuerySelect & "<br>"
		Set rs = Conn.Execute(QuerySelect)
		if Not rs.EOF then
			ConsignerData = ConsignerData & rs(0) & "<br>" & rs(1) & "<br>"
			if rs(2) <> "" then
				ConsignerData = ConsignerData & rs(2)
			end if
		end if
		CloseOBJ rs	
		set rs = Conn.Execute("select numero_telefono from cli_telefonos where id_cliente=" & ClientConsult)
		if Not rs.EOF then
			if rs(0) <> "" then
				ConsignerData = ConsignerData & "    " & rs(0)
			end if
		end if
		CloseOBJ rs
		set rs = Conn.Execute("select nombres from contactos where id_cliente=" & ClientConsult)
		if Not rs.EOF then
			if rs(0) <> "" then
				ConsignerData = ConsignerData & "<br>ATTN:" & rs(0)
			end if
		end if
		CloseOBJs rs, Conn
	end if

    'Se toma primero el pais de la base de datos para desplegar el logo, si viene vacio se toma el pais donde se crea el registro
    if aTableValues(18,0) = "" then
        aTableValues(18,0) = aTableValues(16,0)
    end if
    ShipperID = aTableValues(20, 0)
    'ColoaderID = aTableValues(19, 0)	2019-08-30
    if ColoaderID <> 0  then
        IsColoader = 1
    end if
%>
<html>
<style type="text/css">
<!--
body {
	margin: 0px;
}
.style3 {
	font-size:10px;
	color: #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
}
.style4 {
	font-size:10px;
	color: #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
	border-bottom-style:solid;
	border-left-style:solid;
	border-right-style:solid;
	border-top-style:solid;
	border-collapse:collapse;
	border-width: 1px;
}
.style10 {
	font-size:10px; 
	font-family: Verdana, Arial, Helvetica, sans-serif; 
	font-weight:normal;
}
.style11 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 16px;
}	
.styleborder {
	border-bottom-style:solid;
	border-left-style:solid;
	border-right-style:solid;
	border-top-style:solid;
	border-width: 1px;
	border-collapse:collapse;
}
-->
</style>
<body onLoad="JavaScript:self.focus();">
<table width="641" cellpadding="2" cellspacing="0" align="center">
  <tr>
	<td class="style11" align="left" width="60%">
	<%
        'select case ColoaderID
        'case 7052 'GTLTF
        '    response.write "<img src=http://www.latinfreightneutral.com/img/logo_latin.png border=0>"
        'case 93049, 77222, 75002, 67728, 65768, 63709, 61920, 43421, 29457 'GTTLA
        '    response.write "<img src=http://www.aimargroup.com/img/tla.jpg border=0>"		
        'case else
            response.write DisplayLogo(aTableValues(18,0), 0, 0, 0)
        'end select		
	%>
	<br><br>
	</td>
	<td class="style3" align="right">FO-TR-06</td>
  </tr>
</table>
<table width="641" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
	<td class="style11" align="left"><font color="#0000FF">NOTA&nbsp;DE&nbsp;ARRIBO 
    <%=aTableValues(3,0)%>
    </font></td>
  </tr>
</table>
<table width="641" cellpadding="2" cellspacing="0" align="center" border="0">
  <tr>
    <td class="style10" align="left" colspan="2" valign="top"><br><b>Atención:</b><br><%=ConsignerData%><br></td>
  </tr>
  <tr>
    <td class="style10" align="left" colspan="2" valign="top"><br><b>Fecha de Notificación:</b>&nbsp;<%=aTableValues(4,0)%><br></td>
  </tr>
  <tr>
    <td class="style10" align="left" colspan="2" valign="top">
	<br>Por medio de la presente nos es grato informarle que recibiremos mercaderia consignada a Uds., de acuerdo a los siguientes datos:<br><br></td>
  </tr>	
  <tr>
    <td class="style10" valign="top">
	<table width="100%" cellpadding="2" cellspacing="0" class="styleborder" >
		<tr>
			<td class="style4" valign="middle" width="36%"><B>Fecha Arribo:</B></td>
			<td class="style4" valign="top"><span class="style10"><%=aTableValues(5,0)%></span></td>
		</tr>
		<tr>
			<td class="style4" valign="middle"><B>Bodega Descarga:</B></td>
			<td class="style4" valign="top">
			<span class="style10">
			<% if aTableValues(6,0) <> "" then%>
			<%=aTableValues(6,0)%>
			<%else%>
			<%=aTableValues(8,0)%>
			<%end if%>
			</span>
			</td>
		</tr>
		<tr>
			<td class="style4" valign="middle"><B>Fecha Libre para Entrega Documentos:</B></td>
			<td class="style4" valign="top"><span class="style10"><%=aTableValues(7,0)%></span></td>
		</tr>
		<%if aTableValues(9,0)>=0 then%>	
		<tr>
			<td class="style4" valign="middle"><B>Carta de Porte:</B></td>
			<td class="style4" valign="top"><span class="style10"><%=SetType(aTableValues(9,0),2)%></span></td>
		</tr>
		<%end if%>
		<%if aTableValues(10,0)>=0 then%>	
		<tr>
			<td class="style4" valign="middle"><B>Manifiesto:</B></td>
			<td class="style4" valign="top"><span class="style10"><%=SetType(aTableValues(10,0),2)%></span></td>
		</tr>
		<%end if%>
		<%if aTableValues(11,0)>=0 then%>	
		<tr>
			<td class="style4" valign="middle"><B>Carta de Endoso:</B></td>
			<td class="style4" valign="top"><span class="style10"><%=SetType(aTableValues(11,0),2)%></span></td>
		</tr>
		<%end if%>
		<%if aTableValues(12,0)>=0 then%>	
		<tr>
			<td class="style4" valign="middle"><B>DTI:</B></td>
			<td class="style4" valign="top"><span class="style10"><%=SetType(aTableValues(12,0),2)%></span></td>
		</tr>
		<%end if%>
		<%if aTableValues(13,0)>=0 then%>	
		<tr>
			<td class="style4" valign="middle"><B>BLs:</B></td>
			<td class="style4" valign="top"><span class="style10"><%=SetType(aTableValues(13,0),2)%></span></td>
		</tr>
		<%end if%>
		<%if aTableValues(14,0)>=0 then%>
		<tr>
			<td class="style4" valign="middle"><B>Factura:</B></td>
			<td class="style4" valign="top"><span class="style10"><%=SetType(aTableValues(14,0),2)%></span></td>
		</tr>
		<%end if%>
		<%if aTableValues(15,0)<>"" then%>
		<tr>
			<td class="style4" valign="middle"><B>Observaciones:</B></td>
			<td class="style4" valign="top"><span class="style10"><%=FRegExp(chr(13) & chr(10), aTableValues(15,0), "<br>", 4)%></span></td>
		</tr>
		<%end if%>
	</table>
	<br>
	</td>
  </tr>
  <%for i=0 to CountTableValues2%>
  <tr>
    <td class="style10" align="left" colspan="2" valign="top">		
		<table width="640" class="styleborder" cellpadding="2" cellspacing="0" align="left">
		<tr>
			<td class="style4" align="center" valign="middle">Procedencia de Carga</td>
            <%if aTableValues2(10,i)<>"SV" and aTableValues2(10,i)<>"SVLTF" and aTableValues2(11,i)<>"SV" and aTableValues2(11,i)<>"SVLTF" then %>
			<td class="style4" align="center" valign="middle">Contenedor</td>
            <%end if %>
			<td class="style4" align="center" valign="middle">BL/RO</td>
			<td class="style4" align="center" valign="middle">Embarcador</td>
            <td class="style4" align="center" valign="middle">Producto</td>
            <td class="style4" align="center" valign="middle">Bultos</td>
			<td class="style4" align="center" valign="middle">Peso Bruto</td>
			<td class="style4" align="center" valign="middle">Volumen</td>
		</tr>
		<tr>
			<td class="style4" align="center" valign="top"><span class="style10"><%=TranslateCountry(aTableValues2(1,i))%></span></td>
			<%if aTableValues2(10,i)<>"SV" and aTableValues2(10,i)<>"SVLTF" and aTableValues2(11,i)<>"SV" and aTableValues2(11,i)<>"SVLTF" then %>
			<td class="style4" align="right" valign="top"><span class="style10"><%=aTableValues2(2,i)%></span></td>
			<%end if %>
            <td class="style4" align="right" valign="top"><span class="style10"><%=aTableValues2(3,i)%></span></td>
			<td class="style4" align="center" valign="top"><span class="style10"><%=aTableValues2(4,i)%></span></td>
            <td class="style4" align="center" valign="top"><span class="style10"><%=aTableValues2(9,i)%></span></td>
			<td class="style4" align="center" valign="top"><span class="style10"><%=aTableValues2(5,i) & " " & aTableValues2(6,i)%></span></td>
			<td class="style4" align="center" valign="top"><span class="style10"><%=aTableValues2(7,i)%></span></td>
			<td class="style4" align="center" valign="top"><span class="style10"><%=aTableValues2(8,i)%></span></td>
		</tr>
		</table>
	</td>
  </tr>
  <tr>
  	<td class="style10" align="left" colspan="2" valign="top">		
		<table class="style10" width="20%" cellpadding="2" cellspacing="0" align="left" border="0">
			<br><b>Costos Estimados a Cancelar</b><br>
			<%for j=0 to CountTableValues4
				if aTableValues2(0,i)=aTableValues4(0,j) then
			%>
			<tr>
				<td class="style10" align="left" valign="top"><%=FRegExp(" ", aTableValues4(1,j), "&nbsp;", 4)%></td>
				<td class="style10" align="right" valign="top"><%=aTableValues4(2,j)%></td>
				<td class="style10" align="right" valign="top"><%=FormatNumber(aTableValues4(3,j)+aTableValues4(4,j),2)%></td>
				<td class="style10" align="right" valign="top"><%=SetType(aTableValues4(5,j),5)%></td>
			</tr>
			<%	
                    if aTableValues4(2,j) = "USD" then
                        if aTableValues4(5,j) = 0 then
                            TotUSDCurrencyPrepaid = TotUSDCurrencyPrepaid + CheckNum(aTableValues4(3,j)+aTableValues4(4,j))
                        else
                            TotUSDCurrencyCollect = TotUSDCurrencyCollect + CheckNum(aTableValues4(3,j)+aTableValues4(4,j))
                        end if
                    else
                        if aTableValues4(5,j) = 0 then
                            TotOtherCurrenciesPrepaid = TotOtherCurrenciesPrepaid + CheckNum(aTableValues4(3,j)+aTableValues4(4,j))
                        else
                            TotOtherCurrenciesCollect = TotOtherCurrenciesCollect + CheckNum(aTableValues4(3,j)+aTableValues4(4,j))
                        end if
                    end if
                end if
			next%>	
		</table>
	</td>
  </tr>
   <%next%>
   <tr>
  	<td class="style10" align="left" colspan="2" valign="top">		
		<table class="style10" width="50%" cellpadding="2" cellspacing="0" align="left" border="0">
            <%if TotUSDCurrencyPrepaid <> 0 or TotUSDCurrencyCollect <> 0 then
                Sep = ""%>
            <tr>
	            <td class="style10" align="right" valign="top"><b>TOTAL A PAGAR EN DOLARES (USD)</b></td>
	            <td class="style10" align="right" valign="top">
                <%if TotUSDCurrencyPrepaid <> 0 then
                    Sep = "y&nbsp;"%>
                    <%=FormatNumber(TotUSDCurrencyPrepaid,2)%>&nbsp;Prepaid
                <%end if%>
                <%if TotUSDCurrencyCollect <> 0 then%>
                    <%=Sep & FormatNumber(TotUSDCurrencyCollect,2)%>&nbsp;Collect
                <%end if%>
                </td>
	            <td class="style10" align="right" valign="top">&nbsp;</td>
            </tr>
            <%end if
            if TotOtherCurrenciesPrepaid <> 0 or TotOtherCurrenciesCollect<>0 then            
                Sep = ""%>
            <tr>
	            <td class="style10" align="right" valign="top"><b>TOTAL A PAGAR Otras Monedas</b></td>
	            <td class="style10" align="right" valign="top">
                <%if TotOtherCurrenciesPrepaid <> 0 then
                    Sep = "y&nbsp;"%>
                    <%=FormatNumber(TotOtherCurrenciesPrepaid,2)%>&nbsp;Prepaid
                <%end if%>
                <%if TotOtherCurrenciesCollect <> 0 then%>
                    <%=Sep & FormatNumber(TotOtherCurrenciesCollect,2)%>&nbsp;Collect
                <%end if%>
                </td>
	            <td class="style10" align="right" valign="top">&nbsp;</td>
            </tr>
            <%end if %>
            </table>
        </td>
    </tr>
  <tr>
  	<td class="style10" align="justify" valign="top" colspan="2">
	<br>
	<div align="justify">
    <%=CheckCreditClient(ClientID,SetCountryBAW(aTableValues(16,0)))%><br>
	</div>
	</td>
  </tr>
  <tr>
  	<td class="style10" align="justify" valign="top" colspan="2">
	<br>
	<div align="justify">
	<%=FRegExp(chr(13) & chr(10), aTableValues3(0,0), "<br>", 4)%><br>
	</div>
	</td>
  </tr>  
</table><br>
<table width="641" cellpadding="2" cellspacing="0" align="center">
<%
    ResultEXDBCountry = aTableValues(18,0)

    'response.write ColoaderID & " " & ResultEXDBCountry & "<br>"
%>
<tr>
    <td class="style4" align="left" width="100%">
<%
    select case (ColoaderID-ColoaderID) 'de esta forma para que no entre a TLA 2019-10-08
    'case 40396 'GTAIM
    'case 7052 'GTLTF

    case 93049, 77222, 75002, 67728, 65768, 63709, 61920, 43421, 29457 'GTTLA

        select case aTableValues3(1,0)
            case "GT"
%>
        <u>Observaciones Impotantes:</u>
        <br>
        <u>Para pagos locales:</u>
        <br>
        Para realizar sus pagos en Quetzales, puede realizarlos a través de depósitos monetarios y a continuación detallamos los números de cuenta:<br>
        <br>
	        <table cellpadding="2" cellspacing="0" class="styleborder" width=90%>
		        <tr>
		        <td class="style4" align="left">
		        <span class=style10>BAC (Quetzales)</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>No. 90051701-2</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>Grupo TLA Guatemala, S.A.</span>
		        </td>
		        </tr>
	        </table>
        <br>
        - No se recibirá: efectivo.<br>
        - Enviar copia de la boleta escaneada al correo: cgamboa@grupotla.com, operaciones@ltmcarrier.net con las instrucciones de referencia de pago (detalle de facturas).<br>
        <br>
        Para realizar sus pagos en dólares puede realizar a través de depósitos monetarios a las cuentas detalladas a continuación:<br>
        <br>
	        <table cellpadding="2" cellspacing="0" class="styleborder" width=90%>
		        <tr>
		        <td class="style4" align="left">
		        <span class=style10>BAC (dólares)</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>No. 90075739-4</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>Grupo TLA Guatemala, S.A.</span>
		        </td>
		        </tr>
	        </table>
        <br>
        - Enviar copia de la boleta escaneada al correo: cgamboa@grupotla.com, operaciones@ltmcarrier.net con las instrucciones de referencia de pago (detalle de facturas).<br>
        <br>
        <u>Para pagos internacionales:</u>
        <br>
        CITIBANK<br>
        56: INTERMEDIARY BANK<br>
        NAME: CITIBANK<br>
        ADDRESS: NEW YORK USA<br>
        SWIFT CODE: CITIUS33<br>
        ABA: 021000089<br>
        <br>
        57:BENEFICIARY BANK<br>
        Nombre: Banco de América Central S.A.<br>
        ADDRESS: 7ª ave 6-26 zona 9, Guatemala<br>
        Account number: 36243565<br>
        SWIFT: AMCNGTGT<br>
        <br>
        59:Final Beneficiary<br>
        Name: GRUPO TLA GUATEMALA, S.A.<br>
        Address: 42 CALLE 22-17 INTERIOR 7&8 ZONA 12 GUATEMALA, GUATEMALA<br>
        Account Number: 900757394<br>
        <br>				
 
<%
            case "SV"
%>
<u>Observaciones Impotantes:</u>
<br />
<u>Para pagos locales:</u>
<br />
Para realizar sus pagos en Dolares, puede realizarlos a través de depósitos monetarios y a continuación detallamos los números de cuenta:
<br />
	        <table cellpadding="2" cellspacing="0" class="styleborder" width=90%>
		        <tr>
		        <td class="style4" align="left">
		        <span class=style10>BAC (Dólar)</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>No. 200937837</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>Grupo TLA El Salvador, S.A. de C.V.</span>
		        </td>
		        </tr>
	        </table>
<br />
- No se recibirá: efectivo.<br />
- Enviar copia de la boleta escaneada al correo: rquintanillav@grupotla.com, mecheverria@grupotla.com con las instrucciones de referencia de pago (detalle de facturas).
<br />
<u>Para pagos internacionales:</u><br />
<b>Razón Social : </b> Grupo TLA El Salvador, S.A. de C.V.<br />
<b>Dirección : </b> Km 12 1/2 carretera al puerto La Libertad, Zona Franca Santa Tecla Bodega 6<br />
<b>Teléfono : </b> 2536-6903<br />
<b>Correo Electrónico : </b> mecheverria@grupotla.com<br />
<b>Contacto : </b> Mauridcio Echeverria<br />
<br />
<b><u>Datos bancarios para recibir transferencias locales</u></b><br />
<b><u>Beneficiario:</u></b> Grupo TLA El Salvador, S.A. de C.V.<br />
<b><u>Cuenta:</u></b> 200937837<br />
<br />
<b><u>Datos bancarios para recibir transferencias internacionales</u></b><br />
<b><u>Paying Bank:</u></b> Banco de America Central, S.A.<br />
55 Av. Entre calle Roosevelt y Avenida Olimpica, San Salvador, El Salvador.<br />
<b><u>Tel:</u></b> (503) 2206 4685<br />
<b><u>Contact:</u></b> Rhina de Romero<br />
International Department <br />
<b><u>SWIFT :</u></b> BAMCSVSS<br />
<b><u>CTA. No.:</u></b> 36148605<br />

<%
            case "HN"
%>
<u>Observaciones Impotantes:</u>
<br />
<u>Para pagos locales:</u>
<br />
Para realizar sus pagos en Lempiras, puede realizarlos a través de depósitos monetarios y a continuación detallamos los números de cuenta:
<br />

	        <table cellpadding="2" cellspacing="0" class="styleborder" width=90%>
		        <tr>
		        <td class="style4" align="left">
		        <span class=style10>BAC (Lempiras)</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>No. 730264761</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>Grupo TLA Honduras S.A.</span>
		        </td>
		        </tr>
	        </table>

- Se reciben cheques de empresa local o personales, cheques de caja certificados emitidos a nombre de GRUPO TLA HONDURAS S.A. <br />
- No se recibirá: efectivo.<br />
- Enviar copia de la boleta escaneada al correo: yayala@grupotla.com y pavila@grupotla.com  con las instrucciones de referencia de pago (detalle de facturas o notas de débito).
<br />
<br />
Para realizar sus pagos en dólares puede realizar a través de depósitos monetarios a las cuentas detalladas a continuación:<br />
<br />

	        <table cellpadding="2" cellspacing="0" class="styleborder" width=90%>
		        <tr>
		        <td class="style4" align="left">
		        <span class=style10>BAC (dólares)</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>No. 730264771</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>Grupo TLA Honduras S.A.</span>
		        </td>
		        </tr>
	        </table>
<br />
Enviar copia de la boleta escaneada al correo: yayala@grupotla.com y pavila@grupotla.com  con las instrucciones de referencia de pago (detalle de facturas o notas de débito).
<br /><br />
<u>Para pagos internacionales:</u><br />
<br />
Transferencias Internacionales a BAC Honduras:<br />
<br />
Banco Intermediario: Citibank N.A. 111  Wall Street New York, New York 10043<br />
ABA: 021000089<br />
SWIFT: CITIUS33<br />
CUENTA: 36022113 (Cuenta de BAC Honduras en Citibank)<br />
SWIFT: BMILHNTE<br />
Para finalmente acreditar a:<br />
Nombre del Beneficiario: Grupo TLA Honduras S.A<br />
Cuenta: 730264761<br />

<%
            case "NI"
%>
<u>Observaciones Impotantes:</u>
<br />					
A continuación le detallamos las diferentes formas de pago:<br />
<br />
<u>Para pagos locales:</u><br />
<br />				
Cancelación puede ser:						
<br />				
Transferencia o deposito a nuestras cuentas a nombre de Grupo TLA Nicaragua, S.A. :	<br />

	        <table cellpadding="2" cellspacing="0" class="styleborder" width=90%>
		        <tr>
		        <td class="style4" align="left">
		        <span class=style10>BAC</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>$ 357-49487-1</span>
		        </td>
		        </tr>
		        <tr>
		        <td class="style4" align="left">
		        <span class=style10>BANCENTRO</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>C$ 357-49471-5 T/C oficial.</span>
		        </td>
		        </tr>
		        <tr>
		        <td class="style4" align="left">
		        <span class=style10>BANCENTRO</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>$ 101200644</span>
		        </td>
		        </tr>
		        <tr>
		        <td class="style4" align="left">
		        <span class=style10>BANCENTRO</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>C$ 100200853 T/C oficial.</span>
		        </td>
		        </tr>
	        </table>
<br />                               						
En caso de realizar transferencias, favor enviar comprobante de pago a: Karina Lopez Paz al correo klopez@grupotla.com y hacer referencia del número de conocimiento de embarque que se esté cancelando (bill of lading, carta porte o guía aérea). 						
<br /><br />				
<u>Para pagos internacionales:</u>
<br />
Transferencias Internacionales a BAC Nicaragua:	<br />
<br />
Banco Intermediario: BANK OF AMERICA, N.A.<br />
Dirección: 1st SE Third St Miami, FL 33131 United States<br />
SWIFT: BOFAUS3M<br />
ABA/ROUTING: 026009593<br />
Intermediary Account: 1901621325<br />
Banco Beneficiario: Banco de America Central Nicaragua<br />
Dirección: Km. 4.5 Carretera a Masaya, Complejo Pellas Edificio Norte, 1er Piso, Managua, Nicaragua<br />
SWIFT: BAMCNIMA<br />
Beneficiario Final: GRUPO TLA NICARAGUA, S.A<br />
Número de la Cuenta: 357-49487-1<br />
<%
            case "CR"
%>
<u>Observaciones Impotantes:</u>
<br />
<u>Para pagos locales:</u><br />
<br />				
Para realizar sus pagos puede llevarlos a cabo a través de depósitos monetarios y a continuación detallamos las cuentas bancarias:
<br /><br />

                <table cellpadding="2" cellspacing="0" class="styleborder" width=90%>
		        <tr>
		        <td class="style4" align="left">
		        <span class=style10>EMPRESA</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>BANCO</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>MONEDA</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>CUENTA CORRIENTE</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>CUENTA CLIENTE</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>IBAN</span>
		        </td>
		        </tr>

                <tr>
		        <td class="style4" align="left">
		        <span class=style10>EMPRESA</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>BAC CREDOMATIC</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>COLONES</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>902671148</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>#10200009026711485</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>CR49010200009026711485</span>
		        </td>
		        </tr>
                		        

                <tr>
		        <td class="style4" align="left">
		        <span class=style10>GRUPO TLA</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>BAC CREDOMATIC</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>DOLARES</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>902671932</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>#10200009026719328</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>CR39010200009026719328</span>
		        </td>
		        </tr>


                <tr>
		        <td class="style4" align="left">
		        <span class=style10>GRUPO TLA</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>BANCO NACIONAL DE COSTA RICA</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>COLONES</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>100-01-000-220644-0</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>#15100010012206444</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>CR78015100010012206444</span>
		        </td>
		        </tr>


                <tr>
		        <td class="style4" align="left">
		        <span class=style10>GRUPO TLA</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>BANCO NACIONAL DE COSTA RICA</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>DOLARES</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>100-02-000-063114-3</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>#15100010020631144</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>CR21015100010020631144</span>
		        </td>
		        </tr>


                <tr>
		        <td class="style4" align="left">
		        <span class=style10>GRUPO TLA</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>BANCO DE COSTA RICA</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>COLONES</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>001-0117691-9</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>#15201001011769195</span>
		        </td>
		        <td class="style4" align="left">
		        <span class=style10>CR77015201001011769195</span>
		        </td>
		        </tr>

	        </table>
<br />
Para pagos internacionales:<br />
<br />
INFORMACION PARA RECIBIR TRANSFERENCIAS DEL EXTERIOR, USANDO UNO DE LOS SIGUIENTES CORRESPONSALES:<br />
<br />
JUST ONE OF THEM<br />
<br />

                <table cellpadding="2" cellspacing="0" class="styleborder" width=90%>
		        <tr><td colspan=2 class="style4" align="left">1) WACHOVIA NATIONAL BANK, NEW YORK</td></tr>                                
                <tr><td class="style4" align="left">Cover through:</td><td class="style4" align="left">WACHOVIA NATIONAL BANK, NEW YORK</td></tr>
                <tr><td class="style4" align="left">Account:</td><td class="style4" align="left">2000192000042 (between BAC San José and Wachovia Nat. Bank)</td></tr>
                <tr><td class="style4" align="left">No. ABA</td><td class="style4" align="left">026005092 Swift PNBPUS3NNYC</td></tr>
                <tr><td class="style4" align="left">Transfer to:</td><td class="style4" align="left">BAC San José (formerly Banco San José, S.A.)</td></tr>
                <tr><td class="style4" align="left">Swift</td><td class="style4" align="left">BSNJCRSJ</td></tr>
                <tr><td class="style4" align="left">Beneficiary Name:</td><td class="style4" align="left">Grupo TLA</td></tr>
                <tr><td class="style4" align="left">Beneficiary Account:</td><td class="style4" align="left">902671932</td></tr>
                </table>
	<br />
                <table cellpadding="2" cellspacing="0" class="styleborder" width=90%>
		        <tr><td colspan=2 class="style4" align="left">2) BANK OF AMERICA, MIAMI</td></tr>                                
                <tr><td class="style4" align="left">Cover through:</td><td class="style4" align="left">BANK OF AMERICA, MIAMI</td></tr>
                <tr><td class="style4" align="left">Account:</td><td class="style4" align="left">19019-05932 (between BAC San José and Bank of América)</td></tr>
                <tr><td class="style4" align="left">No. ABA</td><td class="style4" align="left">026009593 Swift BOFAUS3M</td></tr>
                <tr><td class="style4" align="left">Transfer to:</td><td class="style4" align="left">BAC San José (formerly Banco San José, S.A.)</td></tr>
                <tr><td class="style4" align="left">Swift</td><td class="style4" align="left">BSNJCRSJ</td></tr>
                <tr><td class="style4" align="left">Beneficiary Name:</td><td class="style4" align="left">Grupo TLA</td></tr>
                <tr><td class="style4" align="left">Beneficiary Account:</td><td class="style4" align="left">902671932</td></tr>
                </table>
	<br />
                <table cellpadding="2" cellspacing="0" class="styleborder" width=90%>
		        <tr><td colspan=2 class="style4" align="left">3) CITIBANK N.A., NEW YORK</td></tr>                                
                <tr><td class="style4" align="left">Cover through:</td><td class="style4" align="left">CITIBANK N.A., NEW YORK</td></tr>
                <tr><td class="style4" align="left">Account:</td><td class="style4" align="left">36026966 (between BAC San José and Citibank)</td></tr>
                <tr><td class="style4" align="left">No. ABA</td><td class="style4" align="left">021000089 Swift CITIUS33</td></tr>
                <tr><td class="style4" align="left">Transfer to:</td><td class="style4" align="left">BAC San José (formerly Banco San José, S.A.)</td></tr>
                <tr><td class="style4" align="left">Swift</td><td class="style4" align="left">BSNJCRSJ</td></tr>
                <tr><td class="style4" align="left">Beneficiary Name:</td><td class="style4" align="left">Grupo TLA</td></tr>
                <tr><td class="style4" align="left">Beneficiary Account:</td><td class="style4" align="left">902671932</td></tr>
                </table>
	<br />
                <table cellpadding="2" cellspacing="0" class="styleborder" width=90%>
		        <tr><td colspan=2 class="style4" align="left">4) AMERICAN EXPRESS BNAK, NEW YORK</td></tr>                                
                <tr><td class="style4" align="left">Cover through:</td><td class="style4" align="left">AMERICAN EXPRESS BNAK, NEW YORK</td></tr>
                <tr><td class="style4" align="left">Account:</td><td class="style4" align="left">745273 (between BAC San José and American Express Bank)</td></tr>
                <tr><td class="style4" align="left">No. ABA</td><td class="style4" align="left">026001591 Swift CITIUS33</td></tr>
                <tr><td class="style4" align="left">Transfer to:</td><td class="style4" align="left">BAC San José (formerly Banco San José, S.A.)</td></tr>
                <tr><td class="style4" align="left">Swift</td><td class="style4" align="left">BSNJCRSJ</td></tr>
                <tr><td class="style4" align="left">Beneficiary Name:</td><td class="style4" align="left">Grupo TLA</td></tr>
                <tr><td class="style4" align="left">Beneficiary Account:</td><td class="style4" align="left">902671932</td></tr>
                </table>
<br />

<%
            case "PA"
%>
<u>Observaciones Impotantes:</u>
<br />
A continuación le detallamos las diferentes formas de pago:<br />
<br />
<u>Para pagos locales:</u><br />
<br />
Elaborar el cheque a nombre de: Grupo TLA PANAMA S.A<br />
<u>Cheque certificado</u> monto mayor a $100.00<br />
Banco : BAC DE PANAMA<br />
Cta corriente:  104117866<br />
<br />
<u>Para pagos internacionales:</u><br />
	<br />
                <table cellpadding="2" cellspacing="0" class="styleborder" width=90%>		        
                <tr><td class="style4" align="left">Intermediary Bank 56A:</td><td class="style4" align="left">JP MORGAN CHASE BANK NA Swift: CHASUS33 Or ABA:021000021</td></tr>
                <tr><td class="style4" align="left">Address:</td><td class="style4" align="left">New York, N.Y. USA</td></tr>
                <tr><td class="style4" align="left">Beneficiary Bank 57A:</td><td class="style4" align="left">BAC International Bank, Inc / Swift  BCINPAPA / our acct. with JP Morgan Chase Bank NA</td></tr>
                <tr><td class="style4" align="left"></td><td class="style4" align="left">No. 777142555 / Address : Aquilino de la Guardia Street, Urb. Marbella, Panama, Rep. Panama</td></tr>
                <tr><td class="style4" align="left">Beneficiary 59:</td><td class="style4" align="left">GRUPO TLA PANAMA, S.A.</td></tr>
                <tr><td class="style4" align="left">Beneficiary's Acct.:</td><td class="style4" align="left">104117866</td></tr>
                <tr><td class="style4" align="left">By order of 50K:</td><td class="style4" align="left">(your customer)</td></tr>
                <tr><td class="style4" align="left">Important note:</td><td class="style4" align="left">Plase complete all information requested in this form.</td></tr>
                <tr><td class="style4" align="left" colspan=2>All imcomplete incomming money transfer will be RETURNED	</td></tr>
               </table>
<br />	

En caso de realizar transferencias, favor enviar comprobante de pago a: Diego Portugal Dolores al correo ddolores@grupotla.com y hacer referencia del número de conocimiento de embarque que se esté cancelando (bill of lading, carta porte o guía aérea). 					
<br />
<%
        end select
%>        
   </td>
</tr>
<%
    case else    


    If Len(ResultEXDBCountry) > 2 then
     
        select case ResultEXDBCountry
            Case "SVLTF"
%>
 <tr>
	    <td class="style4" align="left" width="100%">
        <b>(EMITIR CHEQUE A NOMBRE DE LATIN FREIGHT DE GUATEMALA, S.A.)</b>
        <br><br>
        CUENTAS BANCARIAS
        <table cellpadding="2" cellspacing="0" class="styleborder">
            <tr>
            <td class="style4" align="left">
            <span class=style10>CITYBANK / BANCO UNO</span>
            </td>
            <td class="style4" align="left">
            <span class=style10>000-170131519011</span>
            </td>
            <td class="style4" align="left">
            <span class=style10>LATIN FREIGHT DE GUATEMALA, S.A.</span>
            </td>
            </tr>
        </table>
        <BR>
        OBSERVACIONES IMPORTANTES
        <table cellpadding="2" cellspacing="0" class="styleborder">
            <tr>
            <td class="style4" align=justify>
            <span class=style10>NO SE ACEPTA EFECTIVO, SOLO CHEQUE DE CAJA O DEPOSITO EN LAS CUENTAS INDICADAS. *** PARA PAGOS EN DOLARES
            EMITIR GIRO BANCARIO ***</br></br>
            POR INSTRUCCIONES DE NUESTRO AGENTE SE SOLICITA UN BL ORIGINAL PARA REGOGER DOCUMENTOS.Favor de hacer este
            pago para poder entregar Copia de poliza de traslado y su respectivo endoso.En caso de RECLAMO debe hacerse por escrito dentro de
            los primeros 03 DIAS DEL CALENDARIO (Contando Sabado y Domingo) mismos que seran contados apartir de la FECHA DE
            DESCARGA ARRIBA DESCRITA de lo contrario EL RECLAMO NO SERA TOMADO EN CUENTA NI SE LE DARA TRAMITE ALGUNO.</br></br>
            TOMAR EN CUENTA QUE LA FACTURA SE REALIZO SEGÚN LOS INFORMACION COLOCADA EN EL REQUERIMIENTO DE PARTIDAS
            O INFORMACION ANTICIPADA, POR CAMBIO DE LA MISMA TIENE UN RECARGO DE Q250.00 ó $35.00 SI SE SOLICITA CAMBIO DE
            FACTURA DE MESES ANTERIORES DEBERA CANCELAR EL VALOR DE IVA E ISR.</span>
            </td>
            </tr>
            </table>
        </td>
      </tr>
<%
   
        end select

        select case aTableValues3(1,0)

        case "GT"
        %>
         <tr>
	        <td class="style4" align="left" width="100%">
            <b>EMITIR CHEQUE A NOMBRE DE Latin Freight de Guatemala , S.A.</b>
            <br><br>
            CUENTAS BANCARIAS
            <table cellpadding="2" cellspacing="0" class="styleborder">
                <tr>
                <td class="style4" align="left">
                <span class=style10>Banco Industrial</span>
                </td>
                <td class="style4" align="left">
                <span class=style10>No. 453-0059049</span>
                </td>
                <td class="style4" align="left">
                <span class=style10>Latin Freight de Guatemala , S.A.</span>
                </td>
                </tr>
                <tr>
                <td class="style4" align="left">
                <span class=style10>G&T Continental</span>
                </td>
                <td class="style4" align="left">
                <span class=style10>No. 77-005164-9</span>
                </td>
                <td class="style4" align="left">
                <span class=style10>Latin Freight de Guatemala , S.A.</span>
                </td>
                </tr>
                <tr>      
                <td class="style4" align="left">
                <span class=style10>Citi Bank</span>
                </td>
                <td class="style4" align="left">
                <span class=style10>No. 1701315190-11</span>
                </td>
                <td class="style4" align="left">
                <span class=style10>Latin Freight de Guatemala , S.A.</span>
                </td>
                </tr>
                <tr>      
                <td class="style4" align="left">
                <span class=style10>G&T Continental (DOLARES)</span>
                </td>
                <td class="style4" align="left">
                <span class=style10>No. 775806180-3</span>
                </td>
                <td class="style4" align="left">
                <span class=style10>Latin Freight de Guatemala , S.A.</span>
                </td>
                </tr>
            </table>
            <BR>
            OBSERVACIONES IMPORTANTES
            <table cellpadding="2" cellspacing="0" class="styleborder">
                <tr>
                <td class="style4" align=justify>
                <span class=style10>Se reciben cheques de empresa local o personales, cheques de caja emitidos a nombre de Latin Freight de Guatemala, S.A.
                Nota: La reincidencia de cheques rechazados conlleva a que se acepten pagos únicamente por medio de cheques de caja y tiene un costo de $35.00 y en quetzales Q168.00
        <BR><BR>
        No se recibirá: efectivo o giros bancarios
        <BR><BR><b>
        Si realiza depósito monetario enviar copia de la boleta escaneada al correo: 
        <br>asiste-creditos3@latinfreightneutral.com con las instrucciones de aplicación de pago (detalle de facturas o notas de débito)
        </b><BR><BR>
        POR INSTRUCCIONES DE NUESTRO AGENTE SE SOLICITA UN BL ORIGINAL PARA REGOGER DOCUMENTOS.Favor de hacer este
        pago para poder entregar Copia de poliza de traslado y su respectivo endoso.En caso de RECLAMO debe hacerse por escrito dentro de
        los primeros 10 DIAS DEL CALENDARIO (Contando Sabado y Domingo) mismos que seran contados apartir de la FECHA DE
        DESCARGA ARRIBA DESCRITA de lo contrario EL RECLAMO NO SERA TOMADO EN CUENTA NI SE LE DARA TRAMITE ALGUNO.
        <BR><BR>
        TOMAR EN CUENTA QUE LA FACTURA SE REALIZO SEGÚN LA INFORMACION COLOCADA EN EL REQUERIMIENTO DE PARTIDAS
        O INFORMACION ANTICIPADA, POR CAMBIO DE LA MISMA TIENE UN RECARGO DE Q250.00 ó $30.00 SI SE SOLICITA CAMBIO DE
        FACTURA DE MESES ANTERIORES DEBERA CANCELAR EL VALOR DE IVA E ISR.</span>
        <BR><BR>
                </td>
                </tr>
                </table>
            </td>
    </tr>
    <%

    end select

    else
        select Case aTableValues3(1,0)
            case "GT"
%>
 <tr>
	    <td class="style4" align="left" width="100%">
        OBSERVACIONES IMPORTANTES
        <table cellpadding="2" cellspacing="0" class="styleborder">
            <tr>
            <td class="style4" align=justify>
            <span class=style10>
            Para realizar sus pagos en quetzales, puede realizarlos a través de depósitos monetarios y  a continuación detallamos los números de convenio:
            <BR><BR>
            <table cellpadding="2" cellspacing="0" class="styleborder">
                <tr>
                <td class="style4" align="left">
                <span class=style10>Banco G&T Continental (Quetzales)</span>
                </td>
                <td class="style4" align="left">
                <span class=style10>No. 012-0001068-6 </span>
                </td>
                <td class="style4" align="left">
                <span class=style10>Aimar, S.A.</span>
                </td>
                <td class="style4" align="left">
                <span class=style10><b>No de convenio: 8292</b></span>
                </td>
                </tr>
                <tr>      
                <td class="style4" align="left">
                <span class=style10>Banco Industrial (Quetzales)</span>
                </td>
                <td class="style4" align="left">
                <span class=style10>No. 027-018962-1</span>
                </td>
                <td class="style4" align="left">
                <span class=style10>Aimar, S.A.</span>
                </td>
                <td class="style4" align="left">
                <span class=style10><b>No de convenio: 2253</b></span>
                </td>
                </tr>
                <tr>      
                <td class="style4" align="left">
                <span class=style10>Banrural (Quetzales)</span>
                </td>
                <td class="style4" align="left">
                <span class=style10>No. 30-3343978-4 </span>
                </td>
                <td class="style4" align="left">
                <span class=style10>Aimar, S.A.</span>
                </td>
                <td class="style4" align="left">
                <span class=style10><b>No de convenio: ATX-249-426-1</b></span>
                </td>
                </tr>
            </table>
            <BR>Se reciben cheques de empresa local o personales, cheques de caja <b>emitidos a nombre de Aimar, S.A. o Agencia Internacional Marítima, S.A.</b>
Nota: La reincidencia de cheques rechazados conlleva a que se acepten pagos únicamente por medio de cheques de caja y tiene un costo de $35.00 y en quetzales Q168.00
        <BR><BR>
        No se recibirá: efectivo o giros bancarios
        <BR><BR>
        Favor utilizar la boleta de depósito que se adjunta en la factura electrónica en la parte inferior del documento.
        <BR><BR>Enviar copia de la boleta escaneada al correo: <b>creditosycobros-gt@aimargroup.com</b> con las instrucciones de aplicación de pago (detalle de facturas o notas de débito)
 
        <BR><BR>Para realizar sus pagos en dólares puede realizar a través de depósitos monetarios a las cuentas en dólares detalladas a continuación:

        <BR><BR>

        <table cellpadding="2" cellspacing="0" class="styleborder">
            <tr>
            <td class="style4" align="left">
            <span class=style10>Banco G&T Continental (Dólares)</span>
            </td>
            <td class="style4" align="left">
            <span class=style10>No. 7858059517</span>
            </td>
            <td class="style4" align="left">
            <span class=style10>Aimar, S.A.</span>
            </td>
            </tr>
            <tr>      
            <td class="style4" align="left">
            <span class=style10>Banco Industrial, S.A. (Dólares)</span>
            </td>
            <td class="style4" align="left">
            <span class=style10>No. 027-003599-1 </span>
            </td>
            <td class="style4" align="left">
            <span class=style10>Aimar, S.A.</span>
            </td>
            </tr>
            <tr>      
            <td class="style4" align="left">
            <span class=style10>Banrural (Dólares)</span>
            </td>
            <td class="style4" align="left">
            <span class=style10>No. 6445015801 </span>
            </td>
            <td class="style4" align="left">
            <span class=style10>Aimar, S.A.</span>
            </td>
            </tr>
        </table>

        <BR><BR>
        Enviar copia de la boleta escaneada al correo: <b>creditosycobros-gt@aimargroup.com</b> con las instrucciones de aplicación de pago (detalle de facturas o notas de débito)

        <BR><BR>O bien puede realizar transferencia bancaria a los datos detallados a continuación:
 
        <BR><BR>Transferencias Internacionales a Banco G&T Continental, S.A.:
        <BR>Banco Intermediario: BANK OF AMERICA, N.A., NEW YORK USA
        <BR><b>ABA:</b> 026009593
        <BR><b>SWIFT:</b> BOFAUS3N
        <BR><b>CUENTA:</b> 1901734945 de Banco G&T Continental, S.A., Guatemala
        <BR><b>SWIFT:</b> GTCOGTGC
        <BR>Para finalmente acreditar a:
        <BR><b>Nombre del Beneficiario:</b> Aimar, S.A.
        <BR>Cuenta: 7858059517
        <BR><BR>
    POR INSTRUCCIONES DE NUESTRO AGENTE SE SOLICITA UN BL ORIGINAL PARA REGOGER DOCUMENTOS.Favor de hacer este
    pago para poder entregar Copia de poliza de traslado y su respectivo endoso.En caso de RECLAMO debe hacerse por escrito dentro de
    los primeros 10 DIAS DEL CALENDARIO (Contando Sabado y Domingo) mismos que seran contados apartir de la FECHA DE
    DESCARGA ARRIBA DESCRITA de lo contrario EL RECLAMO NO SERA TOMADO EN CUENTA NI SE LE DARA TRAMITE ALGUNO.
    <BR><BR>
    TOMAR EN CUENTA QUE LA FACTURA SE REALIZO SEGÚN LA INFORMACION COLOCADA EN EL REQUERIMIENTO DE PARTIDAS
    O INFORMACION ANTICIPADA, POR CAMBIO DE LA MISMA TIENE UN RECARGO DE Q250.00 Ó $30.00 SI SE SOLICITA CAMBIO DE
    FACTURA DE MESES ANTERIORES DEBERA CANCELAR EL VALOR DE IVA E ISR.</span>
            </td>
            </tr>
            </table>
        </td>
</tr>
<%
        case "HN"
%>
  <tr>
	    <td class="style10" align="justify" width="100%">
        Se aceptan las siguientes formas de pago:
        <br><br>
        <b>-Pago por medio de deposito en efectivo.</b>
        <br><br>
        <b>-Pago por medio de deposito de cheque:</b>
        Se espera hasta que los fondos sean reflejados en nuestra cuenta para poder libera documentos o carga.  Si el cheque es devuelto por el banco por falta de fondos, se le cobrara el recargo del banco mas un recargo nuestro de $50.00 si el cheque es en dolares y Lps500.00 si el queche es en Lempiras.
        <br><br>
        <b>-Pago por medio de Transferencia bancaria:</b>
        Se espera hasta que los fondos sean reflejados en nuestra cuenta para poder libera documentos y carga.
        <br><br>
        <b>NO se aceptan Pago en efectivo Ni en Dolares , ni en Lempiras</b>
        <br><br>
        <b>- Pago de facturas en Dolares al cambio en Lempiras</b>
        <br>
        Se devolvera lo depositado en Lempiras restando Lps 600.00 por gastos administrativos y se exigira el pago correspondiente en Dolares para poder liberar documentos o la carga.
        </td>
      </tr>
<%
        end select
        

   end if 

    end select  'fin del case de los ColoaderID de TLA

  %>
 
  <tr>
	<td class="style10" align="left" width="50%"><br>Atentamente,<br><br><%=Session("Sign")%></td>
  </tr>
  <tr>
	<td class="style10" align="center" width="50%"><%=FRegExp(chr(13) & chr(10), aTableValues3(2,0), "<br>", 4)%></td>
  </tr>
</table>
</body>
</html>
<%
	Set aTableValues = Nothing
	Set aTableValues2 = Nothing
	Set aTableValues3 = Nothing
	Set aTableValues4 = Nothing
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>