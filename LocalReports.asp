<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
Checking "0|1|2"
Dim Conn, rs, i, aTableValues, CountTableValues, HTMLCode, Countries, Week, Yr, Mth, Dy, Yrs, BLType, Action, Profit, ProfitPercent
Dim TotWeight, TotVol, TotPieces, TotProfit, SQLFilter

	CountTableValues = -1
	Week = CheckNum(Request("Week"))
	BLType = CheckNum(Request("BLType"))
	Countries  = Request("Countries")
	Yr = CheckNum(Request("Yr"))
	Mth = CheckNum(Request("Mth"))
	Dy = CheckNum(Request("Dy"))
	SQLFilter = ""
	
	Action= 2

	'0(Consolidado) es el default de Itinerario en Transito, 2(Recoleccion) es el default en Itinerario Local
	if BLType = 0 then 
		BLType = 2
	end if
	
	'1 es el default del reporte = Seguimiento TMP
	if Action=0 then
		Action=1
	end if

	'Al no indicarse el pais, se selecciona el primer pais asignado al usuario
	if Countries = "" then
		Countries = Countries = SetDefaultCountry
	end if
	
	'Si no se indica el anio, se toma el actual
	if Yr = 0 then
		Yr = Year(Now)
	end if
	
	OpenConn Conn
	'Creando el listado automatico de anios que tiene el sistema
	Set rs = Conn.Execute("select distinct Year(CreatedDate) as Yr from BLDetail order by Yr Desc")
	do while Not rs.EOF
		Yrs = Yrs & "<option value=" & rs(0) & ">" & rs(0) & "</option>"
		rs.MoveNext
	loop
	CloseOBJ rs
	
	'Si no se indica la semana, se toma la ultima ingresada en el anio indicado
	if Mth<>0 and Dy<>0 then
		SQLFilter = " and Month(b.CreatedDate)=" & Mth & " and Day(b.CreatedDate)=" & Dy		
		Set rs = Conn.Execute("select distinct(b.Week) from BLDetail b where b.BLType=" & BLType & SQLFilter)
		if Not rs.EOF then
			Week = CheckNum(rs(0))
		end if
	else
		'Si no se indica la semana, se toma la ultima ingresada en el anio indicado
		if Week=0 then
			Set rs = Conn.Execute("(select Max(Week) from BLDetail where Year(CreatedDate)=" & Yr & " and Countries='" & Countries & "' and BLType=" & BLType & ")")
			if Not rs.EOF then
				Week = CheckNum(rs(0))
			end if
		end if
		SQLFilter = " and b.Week = " & Week
		Mth = 0
		Dy = 0		
	end if	
	
	Select Case Action
	Case 1	
		set rs = Conn.Execute("select a.BLType, a.BLNumber, a.PolicyNo, a.Marchamo, b.BLs, b.DischargeDate, b.DischargeTime, a.DeliveryPolicyDate, a.DeliveryPolicyHour, a.DeliveryPolicyMin, " & _
			"b.DeliveryDate, a.BLArrivalDate, a.BLArrivalHour, a.BLArrivalMin, a.BLFinishHour, a.BLFinishMin, b.Clients, b.PickUpData, b.DeliveryData, " & _
		 	"b.ClientContact, b.PhoneContact, b.Weights, b.Volumes, b.NoOfPieces, b.DiceContener, c.Name, d.TruckNo, b.Freight+b.Insurance+b.AnotherChargesCollect, " & _
			"b.Freight2+b.Insurance2+b.AnotherChargesPrepaid, sum(e.Cost) " & _
			"from ((((BLs a inner join BLDetail b on a.BLID=b.BLID) inner join Pilots c on a.PilotID=c.PilotID) inner join Trucks d on a.TruckID=d.TruckID) left outer join Costs e on a.BLID=e.BLID) " & _
			"where b.Expired=0 and b.Countries = '" & Countries & "' and Year(b.CreatedDate)=" & Yr & SQLFilter & " and a.BLType = " & BLType & " Group by a.BLID Order by a.CreatedDate, a.BLNumber, b.Priority")
	Case 2
		set rs = Conn.Execute("select a.BLType, e.Name, a.Marchamo, b.PickUpData, b.DeliveryData, a.CreatedDate, a.BLNumber, b.Clients, d.TruckNo, c.Name " & _
			"from BLs a, BLDetail b, Pilots c, Trucks d, Warehouses e " & _
			"where a.BLID=b.BLID and a.PilotID=c.PilotID and a.TruckID=d.TruckID and a.ChargeType=e.WarehouseID " & _
			"and b.Expired=0 and b.Countries = '" & Countries & "' and Year(b.CreatedDate)=" & Yr & SQLFilter & " and a.BLType = " & BLType & _
			" UNION " & _
			"select 2, a.Name, b.Marchamo, '', '', b.CreatedDate, '', 'ANULADO', '', '' from Warehouses a, DelMarchamos b " & _
			"where a.WarehouseID=b.WarehouseID and a.Countries = '" & Countries & "' and Year(b.CreatedDate)=" & Yr & SQLFilter & _
			" order by Marchamo")
	End Select
	
	If Not rs.EOF Then
		aTableValues = rs.GetRows
		CountTableValues = rs.RecordCount-1
	End If
	closeOBJs rs, Conn

	Select Case Action
	Case 1
		for i=0 to CountTableValues
			Profit = aTableValues(27,i)+aTableValues(28,i)-aTableValues(29,i)
			if Profit <> 0 then
				ProfitPercent = Round(Profit/aTableValues(27,i)+aTableValues(28,i),2)*100
			else
				Profit = 0
				ProfitPercent = 0
			end if

			HTMLCode = HTMLCode & "<tr><td class=label>" & i+1 & "</td>" & _
				"<td class=label align=left>" & SetTypeItinerary(0,aTableValues(0,i)) & "</td>" & _
				"<td class=label>" & aTableValues(1,i) & "</td>" & _
				"<td class=label>" & aTableValues(2,i) & "</td>" & _
				"<td class=label>" & aTableValues(3,i) & "</td>" & _
				"<td class=label>" & aTableValues(4,i) & "</td>" & _
				"<td class=label align=center>"& ConvertDate(aTableValues(5,i),1) & "<br>" & TwoDigits(Hour(aTableValues(6,i))) & ":" & TwoDigits(Minute(aTableValues(6,i))) & "</td>" & _
				"<td class=label align=center>"& ConvertDate(aTableValues(7,i),1) & "<br>" & TwoDigits(aTableValues(8,i)) & ":" & TwoDigits(aTableValues(9,i)) & "</td>" & _
				"<td class=label align=center>"& ConvertDate(aTableValues(10,i),1) & "<br>" & TwoDigits(Hour(aTableValues(10,i))) & ":" & TwoDigits(Minute(aTableValues(10,i))) & "</td>" & _
				"<td class=label align=center>"& ConvertDate(aTableValues(11,i),1) & "<br>" & TwoDigits(aTableValues(12,i)) & ":" & TwoDigits(aTableValues(13,i)) & "</td>" & _
				"<td class=label align=center>"& TwoDigits(aTableValues(14,i)) & ":" & TwoDigits(aTableValues(15,i)) & "</td>" & _
				"<td class=label>" & aTableValues(16,i) & "</td>" & _
				"<td class=label>" & aTableValues(17,i) & "</td>" & _
				"<td class=label>" & aTableValues(18,i) & "</td>" & _
				"<td class=label>" & aTableValues(19,i) & "</td>" & _
				"<td class=label>" & aTableValues(20,i) & "</td>" & _
				"<td class=label>" & aTableValues(21,i) & "</td>" & _
				"<td class=label>" & aTableValues(22,i) & "</td>" & _
				"<td class=label>" & aTableValues(23,i) & "</td>" & _
				"<td class=label>" & aTableValues(24,i) & "</td>" & _
				"<td class=label>" & aTableValues(25,i) & "</td>" & _
				"<td class=label>" & aTableValues(26,i) & "</td>" & _
				"<td class=label>" & aTableValues(27,i) & "</td>" & _
				"<td class=label>" & aTableValues(28,i) & "</td>" & _
				"<td class=label>" & aTableValues(29,i) & "</td>" & _
				"<td class=label>" & Profit & "</td>" & _
				"<td class=label>" & ProfitPercent & "%</td>" & _
				"</tr>" & _
				"<tr><td class=submenu colspan=27></td></tr>"
				TotWeight = TotWeight + CheckNum(aTableValues(21,i))
				TotVol = TotVol + CheckNum(aTableValues(22,i))
				TotPieces = TotPieces + CheckNum(aTableValues(23,i))
				TotProfit = TotProfit + Profit
		next
		HTMLCode = HTMLCode & "<tr><td class=submenu colspan=27></td></tr><tr><td class=label colspan=15>&nbsp;</td>" & _
				"<td class=label><b>TOTALES</b></td>" & _
				"<td class=label><b>" & TotWeight & "</b></td>" & _
				"<td class=label><b>" & TotVol & "</b></td>" & _
				"<td class=label><b>" & TotPieces & "</b></td>" & _
				"<td class=label colspan=6>&nbsp;</td>" & _
				"<td class=label><b>" & TotProfit & "</b></td>" & _
				"<td class=label>&nbsp;</td>" & _
				"</tr>" & _
				"<tr><td class=submenu colspan=27></td></tr>"
	Case 2
		for i=0 to CountTableValues
			HTMLCode = HTMLCode & "<tr><td class=label>" & i+1 & "</td>" & _
				"<td class=label align=left>" & SetTypeItinerary(0,aTableValues(0,i)) & "</td>" & _
				"<td class=label>" & aTableValues(1,i) & "</td>" & _
				"<td class=label>" & aTableValues(2,i) & "</td>" & _
				"<td class=label>" & aTableValues(3,i) & "</td>" & _
				"<td class=label>" & aTableValues(4,i) & "</td>" & _
				"<td class=label>" & ConvertDate(aTableValues(5,i),1) & "</td>" & _
				"<td class=label>" & aTableValues(6,i) & "</td>" & _
				"<td class=label>" & aTableValues(7,i) & "</td>" & _
				"<td class=label>" & aTableValues(8,i) & "</td>" & _
				"<td class=label>" & aTableValues(9,i) & "</td>" & _
				"</tr>" & _
				"<tr><td class=submenu colspan=24></td></tr>"
		next
	End Select
%>
<HTML>
<HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/config.js"></SCRIPT>
<SCRIPT src="img/mainLib.js" language="JavaScript"></SCRIPT>
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/validaciones.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<script>
function Validar(Action) {
	if (!valTxt(document.forma.Week, 1, 5)){return (false)};
	if (!valSelec(document.forma.BLType)){return (false)};
	
	document.forma.Action.value = Action;
	document.forma.submit();
 }

</script>
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
.style11 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 10px;
}
.style4 {
	font-size:10px;
	color: #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
}
.style5 {font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 10px;
	color: #000000;
}
.style12 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; color: #FFFFFF; }
.style13 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-weight: bold; color: #FFFFFF; }
.style14 {font-size: 10px; font-family: Verdana, Arial, Helvetica, sans-serif; }
.style14 {font-size: 10px; font-family: Verdana, Arial, Helvetica, sans-serif; }
.style15 {	font-size:10px;
	color: #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
}
.style15 {	font-size:10px;
	color: #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
}
-->
</style>
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
		<TABLE cellspacing=5 cellpadding=2 width=100% align=center>
		<FORM name="forma" action="LocalReports.asp" method="get">
		<input type="hidden" name="Action" value="0">
		<TR>
			<TD class=label align=left colspan="20">
				<table cellspacing="1" cellpadding="0" width="100%">
				<tr>
				<td align="left" width="100">&nbsp;</td>
				<td align="center">
					<%select Case Action
					Case 1%>
					Seguimiento
					<%Case 2%>
					Marchamo
					<%End Select%>
					<%=Countries & "-" & SetType(BLType,3) & "-" & Week  & "-" & Yr%>
				</td>
				<td align="right">
					<table cellspacing="1" cellpadding="0" width="400">
					<tr>
					<TD class=label align=right><b>Ver&nbsp;Reporte</b></TD>
					<TD class=label align=right>&nbsp;</TD>
					<TD class=label align=left>
						<select name="BLType" class=label id="Tipo de Transporte">
							<option value='-1'>Seleccionar</option>
							<option value='2'>LOCAL</option>
							<!--<option value='3'>ENTREGA</option>-->
						</select>
					</TD>
					<TD class=label align=left>
						<select class="style10" name="Yr">
						<%=Yrs%>
						</select>
					</TD>
					<%if Len(Session("Countries"))>6 then%>
					<TD class=label align=left>
						<select name="Countries" class=label>
							<option value=''>Seleccionar</option>
							<%DisplayCountries "", 1%>
						</select>
					</TD>
					<%else%>
						<INPUT name="Countries" type=hidden value="<%=Countries%>">
					<%end if%>
					<TD class=label align=right><b>&nbsp;de&nbsp;</b></TD>
					<TD class=label align=left>
					<select name="Week" class=label id="Semana" onChange="javascript:document.forma.Mth.value=0;document.forma.Dy.value=0;">
							<option value='0'>SEMANA</option>
							<%=SetPriority(0,1)%>
						</select>
					</TD>					
					<TD class=label align=right><b>&nbsp;o&nbsp;</b></TD>
					<TD class=label align=left>
					<select name="Mth" class=label id="Mes" onChange="javascript:document.forma.Week.value=0;">
						<option value='0'>MES</option>
						<option value='1'>01</option>
						<option value='2'>02</option>
						<option value='3'>03</option>
						<option value='4'>04</option>
						<option value='5'>05</option>
						<option value='6'>06</option>
						<option value='7'>07</option>
						<option value='8'>08</option>
						<option value='9'>09</option>
						<option value='10'>10</option>
						<option value='11'>11</option>
						<option value='12'>12</option>
					</select>
					</TD>
					<TD class=label align=left>
					<select name="Dy" class=label id="Dia" onChange="javascript:document.forma.Week.value=0;">>
						<option value='0'>DIA</option>
						<option value='1'>01</option>
						<option value='2'>02</option>
						<option value='3'>03</option>
						<option value='4'>04</option>
						<option value='5'>05</option>
						<option value='6'>06</option>
						<option value='7'>07</option>
						<option value='8'>08</option>
						<option value='9'>09</option>
						<option value='10'>10</option>
						<option value='11'>11</option>
						<option value='12'>12</option>
						<option value='13'>13</option>
						<option value='14'>14</option>
						<option value='15'>15</option>
						<option value='16'>16</option>
						<option value='17'>17</option>
						<option value='18'>18</option>
						<option value='19'>19</option>
						<option value='20'>20</option>
						<option value='21'>21</option>
						<option value='22'>22</option>
						<option value='23'>23</option>
						<option value='24'>24</option>
						<option value='25'>25</option>
						<option value='26'>26</option>
						<option value='27'>27</option>
						<option value='28'>28</option>
						<option value='29'>29</option>
						<option value='30'>30</option>
						<option value='31'>31</option>
					</select>
					</TD>					
					<!--<TD class=label align=left colspan="10"><INPUT name=enviar type=button onClick="JavaScript:Validar(1);return(false);" value="&nbsp;Ver&nbsp;Seguimiento&nbsp;" class=label></TD>-->
					<TD class=label align=left colspan="10"><INPUT name=enviar type=button onClick="JavaScript:Validar(2);return(false);" value="&nbsp;Ver&nbsp;Marchamos&nbsp;" class=label></TD>
					</tr>
					</table>				
				</td>
				</tr>
				</table>
			</TD>
		</TR>
		</TABLE>
		<TABLE cellspacing=5 cellpadding=2 width=100% align=center>
		<TR>
			<%select Case Action
			Case 1%>
			<TD class=titlelist align=left>&nbsp;</TD>
			<TD class=titlelist align=left><b>Tipo de Movimiento:</b></TD>
			<TD class=titlelist align=left><b>CP:</b></TD>
			<TD class=titlelist align=left><b>Poliza:</b></TD>
			<TD class=titlelist align=left><b>Marchamo:</b></TD>
			<TD class=titlelist align=left><b>BL/RO:</b></TD>
			<TD class=titlelist align=left><b>Fecha Solicitud:</b></TD>
			<TD class=titlelist align=left><b>Fecha Liberacion Poliza:</b></TD>
			<TD class=titlelist align=left><b>Fecha Programada de Entrega:</b></TD>
			<TD class=titlelist align=left><b>Fecha LLegada:</b></TD>
			<TD class=titlelist align=left><b>Fecha Finalizacion:</b></TD>
			<TD class=titlelist align=left><b>Cliente:</b></TD>
			<TD class=titlelist align=left><b>Origen:</b></TD>
			<TD class=titlelist align=left><b>Destino:</b></TD>
			<TD class=titlelist align=left><b>Contacto:</b></TD>
			<TD class=titlelist align=left><b>Telefono:</b></TD>
			<TD class=titlelist align=left><b>Peso:</b></TD>
			<TD class=titlelist align=left><b>CBM:</b></TD>
			<TD class=titlelist align=left><b>Bultos:</b></TD>
			<TD class=titlelist align=left><b>Descripci&oacute;n&nbsp;de&nbsp;Carga:</b></TD>
			<TD class=titlelist align=left><b>Piloto:</b></TD>
			<TD class=titlelist align=left><b>Unidad:</b></TD>
			<TD class=titlelist align=left><b>Revenue PP:</b></TD>
			<TD class=titlelist align=left><b>Revenue CC:</b></TD>
			<TD class=titlelist align=left><b>Payout:</b></TD>
			<TD class=titlelist align=left><b>Profit($):</b></TD>
			<TD class=titlelist align=left><b>Profit(%):</b></TD>
			<%Case 2%>
			<TD class=titlelist align=left>&nbsp;</TD>
			<TD class=titlelist align=left><b>Tipo de Movimiento:</b></TD>
			<TD class=titlelist align=left><b>Bodega:</b></TD>
			<TD class=titlelist align=left><b>Marchamo:</b></TD>
			<TD class=titlelist align=left><b>Origen:</b></TD>
			<TD class=titlelist align=left><b>Destino:</b></TD>
			<TD class=titlelist align=left><b>Fecha:</b></TD>
			<TD class=titlelist align=left><b>CP:</b></TD>
			<TD class=titlelist align=left><b>Cliente:</b></TD>
			<TD class=titlelist align=left><b>Unidad:</b></TD>
			<TD class=titlelist align=left><b>Piloto:</b></TD>
			<%End Select%>
		</TR> 
		<%=HTMLCode%>
	</FORM>
	</TABLE>
<script>

selecciona('forma.Yr','<%=Yr%>');
selecciona('forma.BLType','<%=BLType%>');
selecciona('forma.Week','<%=Week%>');
selecciona('forma.Mth','<%=Mth%>');
selecciona('forma.Dy','<%=Dy%>');
<%if Len(Session("Countries"))>6 then%>
	selecciona('forma.Countries','<%=Countries%>');
<%end if%>
</script>	
</BODY>
</HTML>
<%
	Set aTableValues = Nothing
%>