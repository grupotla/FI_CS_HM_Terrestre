<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"
Dim Action, ObjectID, GroupID, QuerySelect, Conn, Conn2, rs, ntr, CountryTitle, QuerySelect2, QuerySelect3, BLType
Dim CountTableValues, aTableValues, CountTableValues2, aTableValues2, CountTableValues3, aTableValues3
Dim BLNumber, BLExitDate, PilotName, ShipperName, License, Countries, TruckNo, Mark, Model, CodProv, Attn, Chassis
Dim SenderData, ConsignerData, CountryDes, Bail, Container, ContainerDep, FinalDes, BLArrivalDate, Consolidated
Dim SubTotNoOfPieces, SubTotWeight, SubTotVolume, TotNoOfPieces, TotWeight, TotVolume, Week, i, j, Nacionality, DTI
Dim LtAcceptNumber, LtAcceptDate, BrokerRecepName, BrokerName, Logo, Footer, BusinessName, Estate, LtEndorseDate, YR
Dim DiceContener(), Weights(), Volumes(), Clients(), NoOfPieces(), CountriesFinalDes(), BLs(), DischargeDate()

	GroupID = CheckNum(Request("GID"))
	Week = CheckNum(Request("W"))
	BLType = CheckNum(Request("BTP"))
	YR = CheckNum(Request("YR"))
	CountTableValues = -1
	CountTableValues2 = -1
	CountTableValues3 = -1
	SubTotNoOfPieces = 0
	ntr = chr(13) & chr(10)
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
-->
</style>
<body onLoad="JavaScript:self.focus();">
<table width="641" cellpadding="2" cellspacing="0" align="center">
  <tr>
	<td class="style11" align="left" width="50%"><img src="img/aimar.gif" border="0"></td>
	<td class="style3" align="right">CA-P-T-04</td>
  </tr>
</table>
<%	
	OpenConn Conn
	Set rs = Conn.Execute("select BLID from BLs where Week=" & Week & " and Year(CreatedDate)=" & YR & " and BLType=" & BLType & " order by CountryDes, BLID Desc")
	If Not rs.EOF Then
    	aTableValues = rs.GetRows
    	CountTableValues = rs.RecordCount-1
	End If
	CloseOBJ rs
	
	for i=0 to CountTableValues

		QuerySelect = "select a.BLNumber, c.Name, a.Week, a.BLExitDate, a.BLArrivalDate, a.ShipperID, a.ContainerDep, a.TotNoOfPieces, a.TotWeight, a.TotVolume, a.Consolidated " & _
					  "from BLs a, Pilots c where a.BLID=" & aTableValues(0,i) & " and a.PilotID=c.PilotID"
		'QuerySelect = "select a.BLNumber, c.Name, a.Week, a.BLExitDate, a.BLArrivalDate, b.Name, a.ContainerDep, a.TotNoOfPieces, a.TotWeight, a.TotVolume, a.Consolidated " & _
		'			  "from BLs a, Shippers b, Pilots c where a.ShipperID = b.ShipperID and a.PilotID = c.PilotID and a.BLID=" & aTableValues(0,i)
		QuerySelect2 = "select a.DiceContener, a.Weights, a.Volumes, a.NoOfPieces, a.CountriesFinalDes, a.Clients, a.BLs, a.DischargeDate from BLDetail a where a.BLID=" & aTableValues(0,i)

		Set rs = Conn.Execute(QuerySelect)
		If Not rs.EOF Then
			aTableValues2 = rs.GetRows
			CountTableValues2 = rs.RecordCount-1
		End If
		CloseOBJ rs
	
		Set rs = Conn.Execute(QuerySelect2)
		If Not rs.EOF Then
			aTableValues3 = rs.GetRows
			CountTableValues3 = rs.RecordCount-1
		End If
		CloseOBJ rs

		if CountTableValues2 >= 0 then
			BlNumber = aTableValues2(0, 0)
			PilotName = aTableValues2(1, 0)
			Week = aTableValues2(2, 0)
			BLExitDate = aTableValues2(3, 0)
			BLArrivalDate = aTableValues2(4, 0)
			OpenConn2 Conn2
			set rs = Conn2.Execute("select nombre_cliente from clientes where es_shipper=true and id_cliente=" & aTableValues2(5, 0))
			if Not rs.EOF then
				ShipperName = rs(0)
			Else
				ShipperName = ""
			End if
			CloseOBJs rs, Conn2
			ContainerDep = aTableValues2(6, 0)
			TotNoOfPieces = aTableValues2(7, 0)
			TotWeight = aTableValues2(8, 0)
			TotVolume = aTableValues2(9, 0)
			Consolidated = aTableValues2(10, 0)
		end if

		if CountTableValues3 >= 0 then
			Redim DiceContener(CountTableValues3)
			Redim Weights(CountTableValues3)
			Redim Volumes(CountTableValues3)
			Redim Clients(CountTableValues3)
			Redim NoOfPieces(CountTableValues3)
			Redim CountriesFinalDes(CountTableValues3)
			Redim BLs(CountTableValues3)
			Redim DischargeDate(CountTableValues3)
	
			for j=0 to CountTableValues3
				DiceContener(j) = aTableValues3(0,j)
				Weights(j) = aTableValues3(1,j)
				Volumes(j) = aTableValues3(2,j)
				NoOfPieces(j) = aTableValues3(3,j)
				CountriesFinalDes(j) = aTableValues3(4,j)
					if Consolidated = 1 then
					Clients(j) = aTableValues3(5,j)
					BLs(j) = aTableValues3(6,j)
					DischargeDate(j) = aTableValues3(7,j)
				else
					Clients(0) = aTableValues3(5,j)
					BLs(0) = aTableValues3(6,j)
					DischargeDate(0) = aTableValues3(7,j)
				end if
			next
		end if
		Set aTableValues2 = Nothing
		Set aTableValues3 = Nothing
%>
<table border="1" align="center">
<tr>
<td>
<table width="641" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
	<td class="style11" align="left">ITINERARIO&nbsp;DE&nbsp;CARGA:&nbsp;<%=BLNumber%></td>
  </tr>
</table>
<table width="641" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
    <td class="style4" align="left" colspan="5" valign="top">Fecha de Salida:<br>
      <span class="style10"><%=BLExitDate%></span></td>
    <td class="style4" align="left" colspan="5" valign="top">Fecha de Llegada:<br>
      <span class="style10"><%=BLArrivalDate%></span></td>
    <td class="style4" align="left" colspan="5" valign="top">Semana:<br>
      <span class="style10"><%=Week%></span></td>
  </tr>
</table>
<table width="641" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
    <td class="style4" align="left" colspan="5" valign="top">Piloto:<br><span class="style10"><%=PilotName%></span></td>
    <td class="style4" align="left" colspan="5" valign="top">Contenedor:<br><span class="style10"><%=ContainerDep%></span></td>
    <td class="style4" align="left" colspan="5" valign="top">Agente:<br><span class="style10"><%=ShipperName%></span></td>
</table>
<br>
<table width="641" height="250" class="styleborder" cellpadding="2" cellspacing="0" align="center">
	<tr height="8">
		<td class="style4" align="center" valign="middle">No. Bultos</td>
		<td class="style4" align="center" valign="middle">Descripci&oacute;n de Carga</td>
		<td class="style4" align="center" valign="middle">Volumen<br>(CBM)</td>
		<td class="style4" align="center" valign="middle">Peso&nbsp;Bruto<br>(Kg)</td>
		<td class="style4" align="center" valign="middle">Consignee<br>(Consolidado)</td>
		<td class="style4" align="center" valign="middle">Fecha<br>Descarga</td>
	</tr>
	<%for j=0 to CountTableValues3%>
		<%if CountryTitle <> CountriesFinalDes(j) then
			CountryTitle = CountriesFinalDes(j)
			if SubTotNoOfPieces > 0 then
		%>
				<tr height="8">
					<td class="style4" align="right" valign="top"><b><%=SubTotNoOfPieces%></b></td>
					<td class="style4" align="center" valign="top"><b>SUB-TOTALES</b></td>
					<td class="style4" align="right" valign="top"><b><%=SubTotWeight%></b></td>
					<td class="style4" align="right" valign="top"><b><%=SubTotVolume%></b></td>
					<td class="style4" align="right" valign="top" colspan="2">&nbsp;</td>
				</tr>
		<%		SubTotNoOfPieces = 0
				SubTotWeight = 0
				SubTotVolume = 0
			end if%>
			<tr>
				<td class="style4" align="left" valign="top" colspan="6"><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<span class="style10"><b><%=TranslateCountry(CountryTitle)%></b></span></td>
			</tr>
		<%end if%>
	<tr>
		<td class="style4" align="right" valign="top"><span class="style10"><%=NoOfPieces(j)%></span></td>
		<td class="style4" align="left" valign="top"><span class="style10"><%=DiceContener(j)%></span></td>
		<td class="style4" align="right" valign="top"><span class="style10"><%=Weights(j)%></span></td>
		<td class="style4" align="right" valign="top"><span class="style10"><%=Volumes(j)%></span></td>
		<td class="style4" align="left" valign="top"><span class="style10"><%=Clients(j)%></span></td>
		<td class="style4" align="right" valign="top"><span class="style10"><%=DischargeDate(j)%></span></td>
	</tr>
	<%	SubTotNoOfPieces = SubTotNoOfPieces + NoOfPieces(j)*1
		SubTotWeight = SubTotWeight + Weights(j)*1
		SubTotVolume = SubTotVolume + Volumes(j)*1
	next%>
	<tr height="8">
		<td class="style4" align="right" valign="top"><b><%=SubTotNoOfPieces%></b></td>
		<td class="style4" align="center" valign="top"><b>SUB-TOTALES</b></td>
		<td class="style4" align="right" valign="top"><b><%=SubTotWeight%></b></td>
		<td class="style4" align="right" valign="top"><b><%=SubTotVolume%></b></td>
		<td class="style4" align="right" valign="top" colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="style4" align="right" valign="top" colspan="6" height="100%">&nbsp;</td>
	</tr>
	<tr height="8">
		<td class="style4" align="right" valign="top"><b><%=TotNoOfPieces%></b></td>
		<td class="style4" align="center" valign="top"><b>TOTALES</b></td>
		<td class="style4" align="right" valign="top"><b><%=TotWeight%></b></td>
		<td class="style4" align="right" valign="top"><b><%=TotVolume%></b></td>
		<td class="style4" align="right" valign="top" colspan="2">&nbsp;</td>
	</tr>
</table>
</td>
</tr>
</table><br>
<%
		SubTotNoOfPieces = 0
		SubTotWeight = 0
		SubTotVolume = 0
CountTableValues2 = -1
		CountTableValues3 = -1
	next
	closeOBJ Conn		
	Set aTableValues = Nothing
%>
</body>
</html>
<%
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>
