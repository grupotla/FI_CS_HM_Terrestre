<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"

Dim Conn, rs, QuerySelect, BLID, aTableValues, BLNumber
Dim ProviderData, sep, PilotData, PilotName, TransportData, PilotInstructions
Dim CreatedDate, CountryOrig, CountryDep, BrokerName, CountryDes, BLType
Dim ItineraryData, Comment2, TotVolume, TotWeight, TotNoOfPieces, ShipperData, ConsignerData

	BLID = CheckNum(Request("BLID"))
	BLType = CheckNum(Request("BTP"))
	
    QuerySelect = "select a.BLID, a.BLNumber, f.Name, f.Address, f.Address2, f.Phone1, f.Phone2, f.Attn, f.Email, " & _
        "c.Name, c.Phone1, c.Passport, c.License, c.Countries, " & _
        "d.TruckNo, e.TruckNo, a.ContainerDep, a.Chassis, " & _
        "a.CreatedDate, a.Countries, a.CountryDep, b.Name, a.CountryDes, a.BLType, a.BLDispatchDate, a.BLEstArrivalDate, " & _
        "a.Comment2, a.TotVolume, a.TotWeight, a.TotNoOfPieces, a.ShipperData, a.ConsignerData, PilotInstructions " & _
        "from (((((BLs a left outer join Trucks e on a.Container = e.TruckID) " & _
		"inner join Brokers b on a.BrokerID = b.BrokerID) inner join Pilots c on a.PilotID = c.PilotID) " & _
		"inner join Trucks d on a.TruckID = d.TruckID) inner join Providers f on d.ProviderID = f.ProviderID) " & _
		"where a.BLID = " & BLID
	
	OpenConn Conn
	'response.Write QuerySelect & "<br><br>"
	Set rs = Conn.Execute(QuerySelect)
	If Not rs.EOF Then
		aTableValues = rs.GetRows
	End If
	closeOBJs rs, Conn

    BLNumber = aTableValues(1,0)
    
    'Nombre Proveedor
    ProviderData = aTableValues(2,0)
    
    'Direccion Proveedor
    if aTableValues(3,0) <> "" then
        ProviderData = ProviderData & "<br>" & aTableValues(3,0)
    end if
    if aTableValues(4,0) <> "" then
        ProviderData = ProviderData & "<br>" & aTableValues(4,0)
    end if
    
    'Telefonos Proveedor
    if aTableValues(5,0) <> "" then
        ProviderData = ProviderData & "<br>" & aTableValues(5,0)
        sep = " -- "
    end if
    if aTableValues(6,0) <> "" then
        ProviderData = ProviderData & sep & aTableValues(6,0)
    end if

    'Contacto Proveedor
    if aTableValues(7,0) <> "" then
        ProviderData = ProviderData & "<br>CONTACTO: " & aTableValues(7,0)
    end if
    if aTableValues(8,0) <> "" then
        ProviderData = ProviderData & "<br>EMAIL: " & aTableValues(8,0)
    end if
    
    'Datos Piloto
    PilotData = aTableValues(9,0)
    PilotName = aTableValues(9,0)
    if aTableValues(10,0) <> "" then
        PilotData = PilotData & "<br>" & aTableValues(10,0)
    end if
    if aTableValues(11,0) <> "" then
        PilotData = PilotData & "<br>IDENTIFICACION:" & aTableValues(11,0)
    end if
    if aTableValues(12,0) <> "" then
        PilotData = PilotData & "<br>LICENCIA:" & aTableValues(12,0)
    end if
    if aTableValues(13,0) <> "" then
        PilotData = PilotData & "<br>NACIONALIDAD:" & TranslateCountry(aTableValues(13,0))
    end if

    'Datos Transporte
    if aTableValues(14,0) <> "" then
        TransportData = "CABEZAL: " & aTableValues(14,0)
    end if
    if aTableValues(15,0) <> "" then
        TransportData = TransportData & "<br>FURGON: " & aTableValues(15,0)
    end if
    if aTableValues(16,0) <> "" then
        TransportData = TransportData & "<br>CONTENEDOR: " & aTableValues(16,0)
    end if
    if aTableValues(17,0) <> "" then
        TransportData = TransportData & "<br>CHASSIS: " & aTableValues(17,0)
    end if
    
    CreatedDate = aTableValues(18,0)
    CountryOrig = aTableValues(19,0)
    CountryDep = TranslateCountry(aTableValues(20,0))
    BrokerName = aTableValues(21,0)
    CountryDes = TranslateCountry(aTableValues(22,0))
    if aTableValues(23,0) = 0 then
        BLType = "CONSOLIDADO (LTL)"
    else
        BLType = "EXPRESS (FTL)"
    end if

    if aTableValues(24,0) <> "" then
        ItineraryData = "FECHA DE SALIDA: " & aTableValues(24,0)
    end if
    if aTableValues(25,0) <> "" then
        ItineraryData = ItineraryData & "<br>FECHA DE LLEGADA APROX: " & aTableValues(25,0)
    end if
    if aTableValues(26,0) <> "" then
        Comment2 = aTableValues(26,0)
    else
        Comment2 = "<br><br>"
    end if
    TotVolume = aTableValues(27,0)
    TotWeight = aTableValues(28,0)
    TotNoOfPieces = aTableValues(29,0)
    ShipperData = aTableValues(30,0)
    ConsignerData = aTableValues(31,0)
    PilotInstructions = Replace(aTableValues(32,0),chr(13) & chr(10),"<br>",1,-1)

    Set aTableValues = Nothing

        Dim iResult, iLogo, iEdicion, iTitulo, iEmpresa, iDireccion, iObservaciones, iPlantilla    
        iResult = WsGetLogo(CountryOrig, "TERRESTRE",  "8",  "",  "")
        iLogo = iResult(20)
        iEdicion = iResult(2)
        iTitulo = iResult(3)
        iEmpresa = iResult(4)
        iDireccion = iResult(6)
        iObservaciones = iResult(1)
        iPlantilla = iResult(22)


    'CountryOrig = "GTTLA"
    'dim iLogo, iTitulo, iEdicion, iEmpresa, iDireccion, iObservaciones, aTableValues5
    'aTableValues5 = EmpresaParametros(CountryOrig, "8", "TERRESTRE")
    'if aTableValues5(1,0) <> "" then    
    '    iLogo = aTableValues5(20,0)
    '    iTitulo = aTableValues5(4,0)    
    '    iObservaciones = aTableValues5(11,0)
    '    iEmpresa = aTableValues5(5,0)
    '    iDireccion = aTableValues5(7,0)
    '    iEdicion = aTableValues5(3,0)
    'end if
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
.style10 {font-size: 10px; font-family: Verdana, Arial, Helvetica, sans-serif; font-weight:normal;}
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
    .style12
    {
        border-style: solid;
        font-size: 10px;
        color: #000000;
        font-family: Verdana, Arial, Helvetica, sans-serif;
        font-weight: bold;
        border-collapse: collapse;
        border-width: 1px;
        width: 374px;
    }
    .style13
    {
        border-style: solid;
        font-size: 10px;
        color: #000000;
        font-family: Verdana, Arial, Helvetica, sans-serif;
        font-weight: bold;
        border-collapse: collapse;
        border-width: 1px;
        width: 85px;
    }
-->
</style>

<body onLoad="JavaScript:self.focus();">
<table width="641" cellpadding="2" cellspacing="0" align="center">
  <tr>
	<td class="style11" align="left" width="50%"><%=DisplayLogo(CountryOrig, 0, 0, 0, iLogo)%></td>
	<td class="style3" align="right"><%=IIf(iEdicion = "", "EDICION 1", iEdicion)%> <% 'IIf(iTitulo = "","CARTA DE INSTRUCCIONES TRANSPORTISTA TERRESTRE INTERNACIONAL",iTitulo)%></td>
  </tr>
</table>
<table width="641" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
	<td class="style11" align="left"><font color="#0000FF"><%=IIf(iTitulo = "","INSTRUCCIONES&nbsp;DE&nbsp;CARTA&nbsp;PORTE",iTitulo)%>&nbsp;No.:&nbsp;<%=BLNumber%></font></td>
  </tr>
</table>
<table width="641" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
    <td class="style4" align="left" valign="top" rowspan=3>Transportista SubContratado:<br>
      <span class="style10"><%=ProviderData%></span></td>    
    <td class="style4" align="left" valign="top">Fecha:<br><span class="style10"><%=CreatedDate%></span></td>
  </tr>
  <tr>
    <td class="style4" align="left" valign="top">Lugar y Pais de Origen:<br><span class="style10"><%=TranslateCountry(CountryOrig)%></span></td>
  </tr>
  <tr>
    <td class="style4" align="left" valign="top">Lugar de Carga:<br><span class="style10"><%=CountryDep%></span></td>
  </tr>
  <tr>
    <td class="style4" align="left" valign="top" rowspan=3>Piloto Asignado:<br>
    <span class="style10"><%=PilotData%></span></td>
    <td class="style4" align="left" valign="top">Aduana de Salida (Exportacion):<br><span class="style10"><%=BrokerName%></span></td>
  </tr>
  <tr>
    <td class="style4" align="left" valign="top">Destino Final:<br><span class="style10"><%=CountryDes%></span></td>
  </tr>
  <tr>
    <td class="style4" align="left" valign="top">Tipo de Servicio:<br><span class="style10"><%=BLType%></span></td>
  </tr>
  <tr>
    <td class="style4" align="left" valign="top" rowspan=3>Equipo Asignado para Servicio 
        Terrestre:<br><span class="style10"><%=TransportData%></span></td>
    <td class="style4" align="left" valign="top">Itinerario de Ruta:<br><span class="style10"><%=ItineraryData%></span></td>
  </tr>
  </table>
  <table width="641" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
    <td class="style4" align="left" valign="top">Mercancias con Notas Tecnicas (Permiso):&nbsp;&nbsp;&nbsp;&nbsp;SI&nbsp;(&nbsp;&nbsp;)&nbsp;&nbsp;&nbsp;&nbsp;NO (&nbsp;&nbsp;)   <br>
      </td>
  </tr>
  <tr>
    <td class="style4" align="left" valign="top">Cartas de Porte de Mercancias Sujetas a Tramites de Notas Tecnicas en Frontera:<br>
      <span class="style10"><%=Comment2%></span></td>
  </tr>
  </table>
  <table width="641" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
    <td class="style4" align="left" valign="top">Volumen Total:<br><span class="style10"><%=TotVolume%></span></td>
    <td class="style4" align="left" valign="top">Peso Neto:<br><span class="style10"><%=TotWeight%></span></td>
    <td class="style4" align="left" valign="top">Total de No. Bultos:<br><span class="style10"><%=TotNoOfPieces%></span></td>
  </tr>  
  </table>
  <table width="641" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
    <td class="style4" align="left" valign="top">Contacto Envio:<br><span class="style10"><%=ShipperData%></span></td>
    <td class="style4" align="left" valign="top">Contacto Destino:<br><span class="style10"><%=ConsignerData%></span></td>
  </tr>  
  </table>
  <table width="641" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
    <td class="style4" align="left" valign="top">Observaciones de la Carga:<br>
    <%=iObservaciones%>
    </td>
  </tr>
  </table>
  <table width="641" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
    <td class="style4" align="left" valign="top">Instrucciones Precisas que debe de Efectuar el Transportista segun contratacion y servicio ofrecido:<br>
    <span class="style10">
    <%=PilotInstructions%>
    </span>
    </td>
  </tr>
  </table>
  <table width="641" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
    <td align="center" class="style12">Documentacion Proporcionada</td>
    <td align="center" class="style13">Original</td>
	<td align="center" class="style13">Copia</td>
	<td width="488" align="center" class="style4">Compromiso</td>
  </tr>
 <tr>
    <td align="left" class="style12"><span class=style10>No. DE FACTURAS</span></td>
    <td align="left" class="style13">&nbsp;</td>
	<td align="left" class="style13">&nbsp;</td>
    <td width="488" align="justify" valign=top class="style4" rowspan=8>
    <span class=style10>
    EL TRANSPORTISTA SE COMPROMETE A CUMPLIR CADA UNA DE LAS INSTRUCCIONES CONSIGANADAS EN ESTA CARTA.<br>
    EL INCUMPLIMIENTO DE ALGUNA DE LAS INSTRUCCIONES DARA NACIMIENTO DE OBLIGACIONES POR PARTE DEL TRANSPORTISTA EN LO QUE RESPECTA A PAGO DE COSTOS EXTRAS, MULTAS Y CARGOS ASOCIADOS QUE SE IMPONGAN POR INCUMPLIMIENTO DE LA INSTRUCCION.
    </span>
    </td>
 </tr>
 <tr>
    <td align="left" class="style12"><span class=style10>CARTA DE INSTRUCCIONES</span></td>
    <td align="left" class="style13">&nbsp;</td>
	<td align="left" class="style13">&nbsp;</td>
 </tr>
 <tr>
    <td align="left" class="style12"><span class=style10>NO. DE CARTAS DE PORTE</span></td>
    <td align="left" class="style13">&nbsp;</td>
	<td align="left" class="style13">&nbsp;</td>
 </tr>
 <tr>
    <td align="left" class="style12"><span class=style10>MANIFIESTO DE CARGA MASTER</span></td>
    <td align="left" class="style13">&nbsp;</td>
	<td align="left" class="style13">&nbsp;</td>
 </tr>
 <tr>
    <td align="left" class="style12"><span class=style10>MANIFIESTO DE CARGA INDIVIDUAL</span></td>
    <td align="left" class="style13">&nbsp;</td>
	<td align="left" class="style13">&nbsp;</td>
 </tr>
 <tr>
    <td align="left" class="style12"><span class=style10>PERMISOS DE MERCANCIA</span></td>
    <td align="left" class="style13">&nbsp;</td>
	<td align="left" class="style13">&nbsp;</td>
 </tr>
 <tr>
    <td align="left" class="style12"><span class=style10>DUT</span></td>
    <td align="left" class="style13">&nbsp;</td>
	<td align="left" class="style13">&nbsp;</td>
 </tr>
 <tr>
    <td align="left" class="style12"><span class=style10>DUAS</span></td>
    <td align="left" class="style13">&nbsp;</td>
	<td align="left" class="style13">&nbsp;</td>
 </tr>
</table>
<table width="641" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
    <td class="style4" align="left" valign="top">Recibido:<br>
    <span class=style10>PILOTO ASIGNADO: <%=PilotName%> <br><br><br>
    FIRMA:_____________________________________ IDENTIFICACION: _____________________________________<br><br>
    ELABORADO POR: <%=Session("Sign")%>
    </span>
    <br><br></td>
  </tr>

  <tr>
	<td class="style10" align="center" width="50%"><%=iEmpresa&"<br>"&iDireccion%></td>
  </tr>

</table>
</body>
</html>
<%
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>