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


'/////////////////////////////////////////////////////2019-12-04////////////////////////////////////////////////////////////////       

        ResultEXDBCountry = aTableValues(18,0)
        
        response.write ColoaderID & " " & ResultEXDBCountry & " " & aTableValues3(1,0) & "<br>"

        dim CountTableValues5, aTableValues5
        dim iEdicion, iTitulo, iEmpresa, iDireccion, iObservaciones

        'aTableValues5 = EmpresaParametros(aTableValues3(1,0), ColoaderID, ResultEXDBCountry)

        'iEdicion = rs(3)
        'iTitulo = rs(4)
        'iEmpresa = rs(5)
        'iDireccion = rs(10)
        iObservaciones = aTableValues5(19,0)



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
<tr>
    <td class="style4" align="left" width="100%">

<%'/////////////////////////////////////////////////////2019-12-04////////////////////////////////////////////////////////////////
    response.write iObservaciones   
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