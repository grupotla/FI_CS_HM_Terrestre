<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
Checking "0|1|2"
Dim Conn, rs, i, aTableValues, CountTableValues, SBLIDS, PIDS, STATUSIDs, CantPIDS, HTMLCode
Dim SQLFilter, Countries, Week, Yr, Mth, Dy, Yrs, BLType, ItineraryType, Query, ModifyItinerary
Dim Profit, ProfitPercent, Profit2, ProfitPercent2, TotWeight, TotVol, TotPieces, TotProfit, TotProfit2
Dim Empresa, CtrsTemp

	CountTableValues = -1
	SBLIDS = Request("SBLIDS")
	Week = CheckNum(Request("Week"))
	BLType = CheckNum(Request("BLType"))
	Countries  = Request("Countries")
	Yr = CheckNum(Request("Yr"))
	ItineraryType = CheckNum(Request("IT"))
	Mth = CheckNum(Request("Mth"))
	Dy = CheckNum(Request("Dy"))
	SQLFilter = ""
    ModifyItinerary = 0
    if Instr(1, Session("Countries"), Countries)>0 then
        ModifyItinerary = 1        
    end if


	'0(Consolidado) es el default de Itinerario en Transito, 2(Recoleccion) es el default en Itinerario Local
	if ItineraryType=2 and BLType = 0 then 
		BLType = 2
	end if

	'Al no indicarse el pais, se selecciona el primer pais asignado al usuario
	if Countries = "" then
        Countries = SetDefaultCountry
	end if

    'Select Case Countries
    'Case "GT","GTLTF"
    '    CtrsTemp = "GT','GTLTF"
    'Case "SV","SVLTF"
    '    CtrsTemp = "SV','SVLTF"
    'Case "HN","HN1","HN2","HNLTF"
    '    CtrsTemp = "HN','HN1','HN2','HNLTF"
    'Case "NI","NILTF"
    '    CtrsTemp = "NI','NILTF"
    'Case "CR","CRLTF"
    '    CtrsTemp = "CR','CRLTF"
    'Case "PA","PALTF"
    '    CtrsTemp = "PA','PALTF"
    'Case "MX","MXLTF"
    '    CtrsTemp = "MX','MXLTF"
    'End Select

    CtrsTemp = Left(Countries,2)
	
	'Si no se indica el anio, se toma el actual
	if Yr = 0 then
		Yr = Year(Now)
	end if
	
	OpenConn Conn
	'Eliminando o Desasignado Registros a un Itinerario
	if SBLIDS <> "" then
		select Case CheckNum(Request("Action"))
		Case 2
			'Desasignando Registros Nuevos, borra el HBLNumber, para obtener el que corresponda cuando vuelva a ser asignado
			Conn.Execute("update BLDetail set InTransit=0, Pos=0, BLID=-1, Week=0, BLType=-" & ItineraryType & ", HBLNumber='--' where BLDetailID in (" & SBLIDS & ") and BLIDTransit=0 and BLID = -1")
			'Desasignando Registros en Transito, no borra el HBLNumber porque se debe mantener por la poliza de seguro
			Conn.Execute("update BLDetail set InTransit=0, Pos=0, BLID=-1, Week=0, BLType=-" & ItineraryType & " where BLDetailID in (" & SBLIDS & ") and BLIDTransit<>0")
		Case 3 'Actualizando Prioridades en un Itinerario Local
			PIDS = Split(Request("PIDS"),",") 'Prioridades
			SBLIDS = Split(Request("SBLIDS"),",") 'BLDetail IDs
			STATUSIDs = Split(Request("STATUSIDs"),",") 'Sstatus de IDs
			CantPIDS = UBound(PIDS)
			for i=0 to CantPIDS
				Conn.Execute("update BLDetail set Priority=" & PIDS(i) & " where BLDetailID = " & SBLIDS(i))
			next
		End Select
		SBLIDS = ""
	end if

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
		'Set rs = Conn.Execute("select distinct(b.Week) from BLDetail b where b.BLType=" & BLType & SQLFilter)
		'if Not rs.EOF then
	'		Week = CheckNum(rs(0))
	'	end if
        Week = 0
    else 

        if Mth<>0 then

            SQLFilter = " and Month(b.CreatedDate)=" & Mth 
            Week = 0

	    else
		    'Si no se indica la semana, se toma la ultima ingresada en el anio indicado
		    if Week=0 then
			    'Set rs = Conn.Execute("(select Max(Week) from BLDetail where Year(CreatedDate)=" & Yr & " and Countries in ('" & CtrsTemp & "') and BLType=" & BLType & ")")
                Set rs = Conn.Execute("(select Max(Week) from BLDetail where Year(CreatedDate)=" & Yr & " and substr(Countries,1,2) = '" & CtrsTemp & "' and BLType=" & BLType & ")")
			    if Not rs.EOF then
				    Week = CheckNum(rs(0))
			    end if
		    end if
		    SQLFilter = " and b.Week = " & Week
		    Mth = 0
		    Dy = 0		
	    end if

    end if
	
    dim qry 

   
	'Buscando Registros pendientes para Eliminar o Asignar a un Itinerario
	Select Case ItineraryType
	Case 1 'Asignados en Transito
		'								0			1		2			  3			   4			       5			6			7		   8			9			        10	           11       12          13          14		   15		 16	       17		    18          19
		'response.write("select b.BLDetailID, b.Clients, b.BLIDTransit, b.BLType, b.DischargeDate, b.DiceContener, b.Weights, b.Volumes, b.NoOfPieces, b.CountryOrigen, b.CountriesFinalDes, Notify, b.Agents, b.Shippers, b.Coloaders, b.Contact, b.Container, b.BLs, CreatedDate, b.CreatedTime, b.EXType, b.EXDBCountry from BLDetail b where b.InTransit in (1,2) and b.Expired=0 and substr(b.Countries, 1, 2) like '" & Mid(CtrsTemp, 1, 2) & "%' and Year(b.CreatedDate)=" & Yr & SQLFilter & " and b.BLType = " & BLType & " Order by b.CountriesFinalDes, b.Pos")
		'set rs = Conn.Execute("select b.BLDetailID, b.Clients, b.BLIDTransit, b.BLType, b.DischargeDate, b.DiceContener, b.Weights, b.Volumes, b.NoOfPieces, b.CountryOrigen, b.CountriesFinalDes, Notify, b.Agents, b.Shippers, b.Coloaders, b.Contact, b.Container, b.BLs, CreatedDate, b.CreatedTime, b.EXType, b.EXDBCountry from BLDetail b where b.InTransit in (1,2) and b.Expired=0 and substr(b.Countries, 1, 2) like '" & Mid(CtrsTemp, 1, 2) & "%' and Year(b.CreatedDate)=" & Yr & SQLFilter & " and b.BLType = " & BLType & " Order by b.CountriesFinalDes, b.Pos")     

        qry = "select b.BLDetailID, b.Clients, b.BLIDTransit, b.BLType, b.DischargeDate, b.DiceContener, b.Weights, b.Volumes, b.NoOfPieces, b.CountryOrigen, b.CountriesFinalDes, Notify, b.Agents, b.Shippers, b.Coloaders, b.Contact, b.Container, b.BLs, CreatedDate, b.CreatedTime, b.EXType, b.EXDBCountry from BLDetail b where b.InTransit in (1,2) and b.Expired=0 and substr(b.Countries,1,2) = '" & CtrsTemp & "' and Year(b.CreatedDate)=" & Yr & SQLFilter & " and b.BLType = " & BLType & " " & _ 
        " Order by b.CountriesFinalDes, b.Pos"
        '"AND COALESCE( CURRENT_DATE - cast(b.CreatedDate as date),0)  < 90 " & _ 
        
	Case 2 'Asignados Local
		qry = "select b.BLStatusName, a.BLNumber, a.PolicyNo, a.Marchamo, b.BLs, b.DischargeDate, b.DischargeTime, a.DeliveryPolicyDate, a.DeliveryPolicyHour, a.DeliveryPolicyMin, " & _
			"b.DeliveryDate, a.BLArrivalDate, a.BLArrivalHour, a.BLArrivalMin, a.BLFinishHour, a.BLFinishMin, b.Clients, b.PickUpData, b.DeliveryData, " & _
		 	"b.ClientContact, b.PhoneContact, b.Weights, b.Volumes, b.NoOfPieces, b.DiceContener, c.Name, f.Name, d.TruckNo, " & _
			"(select sum(g.Value+g.OverSold) from ChargeItems g where g.SBLID=b.BLDetailID and g.Currency='USD' and g.Expired=0 and InvoiceID<>0), " & _
			"(select sum(h.Value+h.OverSold) from ChargeItems h where h.SBLID=b.BLDetailID and h.Currency<>'USD' and h.Expired=0 and InvoiceID<>0), " & _
			"(select sum(e.Cost) from Costs e where a.BLID=e.BLID and Currency='USD' and e.Expired=0), " & _
			"b.BLDetailID, b.Priority, b.CreatedDate, b.CreatedTime, " & _
			"(select sum(i.Cost) from Costs i where a.BLID=i.BLID and Currency<>'USD' and i.Expired=0) " & _
			"from ((((BLDetail b left join BLs a on a.BLID=b.BLID) left join Pilots c on a.PilotID=c.PilotID) left join Trucks d on a.TruckID=d.TruckID) left join Providers f on c.ProviderID=f.ProviderID) " & _
			"where b.InTransit in (1,2) and b.Expired=0 and substr(b.Countries,1,2) = '" & CtrsTemp & "' and Year(b.CreatedDate)=" & Yr & SQLFilter & " and b.BLType = " & BLType & " " & _             
            " Order by b.DeliveryDate, b.Priority"
			'"where b.InTransit in (1,2) and b.Expired=0 and substr(b.Countries, 1, 2) in ('" & CtrsTemp & "') and Year(b.CreatedDate)=" & Yr & SQLFilter & " and b.BLType = " & BLType & " Order by b.DeliveryDate, b.Priority")
        	'"AND COALESCE( CURRENT_DATE - cast(b.CreatedDate as date),0)  < 90 " & _ 
            
    End Select

    'response.write qry & "<br>"

    set rs = Conn.Execute(qry)

	If Not rs.EOF Then
		aTableValues = rs.GetRows
		CountTableValues = rs.RecordCount-1
	End If
	closeOBJs rs, Conn

	Select Case ItineraryType
	Case 1
        'OpenConn2 Conn
        'for i=0 to CountTableValues
		'    set rs = Conn.Execute("select id_cliente, nombres, telefono from contactos where id_cliente=" & aTableValues(0,i) & " order by telefono desc")
	    '    if Not rs.EOF then
        '        aTableValues(11,i) = rs(0)
		'        aTableValues(12,i) = rs(1)
	    '    end if
        '    CloseOBJ rs
        'next
        'CloseOBJ Conn
		
        for i=0 to CountTableValues
			Select Case aTableValues(20,i)
                Case 0,1,2,11,12,13,8,15
                    'Select Case aTableValues(21,i)
                    '    Case "GTLTF","SVLTF","HNLTF","NILTF","CRLTF","PALTF","BZLTF"
                    '        Empresa = " / LATIN FREIGHT"
                    '    Case "MXLTF","MX"
                    '        Empresa = ""
                    '    Case Else
                    '        Empresa = " / AIMAR GROUP"
                    'End Select
                    Empresa = TranslateCompany(aTableValues(21,i))
                Case Else
                    Empresa = ""
            End Select
            SBLIDS = SBLIDS & "SBLIDS[" & i & "]=" & aTableValues(0,i) & ";" & vbCrLf
			HTMLCode = HTMLCode & "<tr><td class=label><input type=checkbox name='Pos" & i & "'></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(1,i) & "</a></td>" & _
				"<td class=label align=center><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & SetTypeItinerary(aTableValues(2,i),aTableValues(3,i)) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(4,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(5,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(6,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(7,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(8,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(9,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(10,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(11,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(12,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(13,i) & "" & Empresa & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(14,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(15,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(16,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(17,i) & "</a></td>" & _
				"<td class=style4 align=center bgcolor=#999999>"
                if ModifyItinerary = 1 then
                    HTMLCode = HTMLCode & "<a href=# onClick=JavaScript:Editar(" & aTableValues(0,i) & ",'" & aTableValues(18,i) & "'," & aTableValues(19,i) & ");return (false); class=submenu><font color=FFFFFF><b>&nbsp;Editar&nbsp;</b></font></a>"
                end if
				HTMLCode = HTMLCode & "</td></tr><tr><td class=submenu colspan=18></td></tr>"
		next
	Case 2
		for i=0 to CountTableValues
			SBLIDS = SBLIDS & "SBLIDS[" & i & "]=" & aTableValues(31,i) & ";" & vbCrLf
			
			Profit = aTableValues(28,i)-aTableValues(30,i)
			if Profit <> 0 then
				ProfitPercent = Round(Profit/aTableValues(28,i),2)*100
			else
				Profit = 0
				ProfitPercent = 0
			end if

			Profit2 = aTableValues(29,i)-aTableValues(35,i)
			if Profit2 <> 0 then
				ProfitPercent2 = Round(Profit2/aTableValues(29,i),2)*100
			else
				Profit2 = 0
				ProfitPercent2 = 0
			end if

			HTMLCode = HTMLCode & "<tr><td class=label><input type=checkbox name='Pos" & i & "'></td>" & _
				"<td class=label><select name='S" & i & "' class=label>" & SetPriority(aTableValues(32,i),2) & "</select></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(0,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(1,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(2,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(3,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(4,i) & "</a></td>" & _
				"<td class=label align=center><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & ConvertDate(aTableValues(5,i),1) & "<br>" & TwoDigits(Hour(aTableValues(6,i))) & ":" & TwoDigits(Minute(aTableValues(6,i))) & "</a></td>" & _
				"<td class=label align=center><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & ConvertDate(aTableValues(7,i),1) & "<br>" & TwoDigits(aTableValues(8,i)) & ":" & TwoDigits(aTableValues(9,i)) & "</a></td>" & _
				"<td class=label align=center><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & ConvertDate(aTableValues(10,i),1) & "<br>" & TwoDigits(Hour(aTableValues(10,i))) & ":" & TwoDigits(Minute(aTableValues(10,i))) & "</a></td>" & _
				"<td class=label align=center><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & ConvertDate(aTableValues(11,i),1) & "<br>" & TwoDigits(aTableValues(12,i)) & ":" & TwoDigits(aTableValues(13,i)) & "</a></td>" & _
				"<td class=label align=center><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & TwoDigits(aTableValues(14,i)) & ":" & TwoDigits(aTableValues(15,i)) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(16,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(17,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(18,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(19,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(20,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(21,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(22,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(23,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(24,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(25,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(26,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(27,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(28,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(29,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(30,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(35,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & Profit & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & Profit2 & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & ProfitPercent & "%</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & ProfitPercent2 & "%</a></td>" & _
				"<td class=style4 align=center bgcolor=#999999><a href=# onClick=JavaScript:Editar(" & aTableValues(31,i) & ",'" & aTableValues(33,i) & "'," & CheckNum(aTableValues(34,i)) & ");return (false); class=submenu><font color=FFFFFF><b>&nbsp;Editar&nbsp;</b></font></a></td></tr>" & _
                "<tr><td class=submenu colspan=33></td></tr>"
				
				TotWeight = TotWeight + CheckNum(aTableValues(21,i))
				TotVol = TotVol + CheckNum(aTableValues(22,i))
				TotPieces = TotPieces + CheckNum(aTableValues(23,i))
				TotProfit = TotProfit + Profit
				TotProfit2 = TotProfit2 + Profit2
		next
		HTMLCode = HTMLCode & "<tr><td class=submenu colspan=33></td></tr><tr><td class=label colspan=16>&nbsp;</td>" & _
				"<td class=label><b>TOTALES</b></td>" & _
				"<td class=label><b>" & TotWeight & "</b></td>" & _
				"<td class=label><b>" & TotVol & "</b></td>" & _
				"<td class=label><b>" & TotPieces & "</b></td>" & _
				"<td class=label colspan=8>&nbsp;</td>" & _
				"<td class=label><b>" & TotProfit & "</b></td>" & _
				"<td class=label><b>" & TotProfit2 & "</b></td>" & _
				"<td class=label>&nbsp;</td>" & _
				"</tr>" & _
				"<tr><td class=submenu colspan=33></td></tr>"
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
<SCRIPT LANGUAGE="JavaScript">
var SBLIDS = new Array();
<%=SBLIDS%>
function SetList(Pos) {
	if (document.forma.elements[Pos].checked){
		document.forma.elements[Pos].checked = false;
	} else {
		document.forma.elements[Pos].checked = true;
	}
}

function SetAll() {
	if (document.forma.Set.checked) {
		for (var i=0; i<<%=i%>; i++) {
			document.forma.elements["Pos" + i].checked = true;
		}
	} else {
		for (var i=0; i<<%=i%>; i++) {
			document.forma.elements["Pos" + i].checked = false;
		}

	}
}

function SetIDs() {
	var sep = "";
	document.forma.SBLIDS.value = "";
	for (var i=0; i<<%=i%>; i++) {
		if (document.forma.elements["Pos" + i].checked) {
			document.forma.SBLIDS.value = document.forma.SBLIDS.value + sep + SBLIDS[i];
			sep = ","
		}
	}
}

function SetPIDs() {
	var sep = "";
	document.forma.SBLIDS.value = "";
	document.forma.PIDS.value = "";
	document.forma.STATUSIDs.value = "";
	for (var i=0; i<<%=i%>; i++) {
		document.forma.SBLIDS.value = document.forma.SBLIDS.value + sep + SBLIDS[i];
		document.forma.PIDS.value = document.forma.PIDS.value + sep + document.forma.elements["S" + i].value;
		document.forma.STATUSIDs.value = document.forma.STATUSIDs.value + sep + document.forma.elements["STATUS" + i].value;
		sep = ","
	}
}

function Editar(OID, CD, CT){
		window.open('InsertData.asp?GID=27&OID='+OID+'&CD='+CD+'&CT='+CT,'EData','height=650,width=600,menubar=0,resizable=1,scrollbars=1,toolbar=0');
}


function Validar(Action) {
	SetIDs();
	if (Action == 1) {
		if (!valTxt(document.forma.Week, 1, 5)){return (false)};
		if (!valSelec(document.forma.BLType)){return (false)};
	}
	if (Action == 2) {
        if (document.forma.SBLIDS.value == "") {
		    alert("Por favor seleccione al menos un registro para Eliminar");
			return (false);
		};		
		document.forma.Week.value = '<%=Week%>';
		document.forma.BLType.value = '<%=BLType%>';
		document.forma.Countries.value = '<%=Countries%>';
        //document.forma.CtrsTemp.value = '<%=CtrsTemp%>';
		document.forma.Yr.value = '<%=Yr%>';
	}
	if (Action == 3) {
		SetPIDs();
	}
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
	<FORM name="forma" action="ItineraryAsigs.asp" method="get">
	<input name="IT" type=hidden value=<%=ItineraryType%>>
	<input name="Action" type=hidden value=0>
	<input type="hidden" name="SBLIDS" value="">
	<input type="hidden" name="PIDS" value="">
	<input type="hidden" name="STATUSIDs" value="">
		<TR>
			<TD class=label align=left colspan="16">
				<table cellspacing="1" cellpadding="0" width="100%">
				<tr>
				<td align="center">
				Itinerario <%=CtrsTemp & "-" & SetType(BLType,3) & "-" & Week  & "-" & Yr%>
				</td>
				<td>&nbsp;</td>
				</tr>
				<tr>
				<td align="center">
				    <%if ModifyItinerary=1 then %>
                	<table cellspacing="1" cellpadding="0" width="200">
					<tr>
					<TD class=label align=center colspan="10"><INPUT name=enviar type=button onClick="JavaScript:Validar(2);" value="&nbsp;Desasignar&nbsp;Seleccionados&nbsp;" class=label></TD>
					<%if ItineraryType=2 then%>
					<TD class=label align=center colspan="10"><INPUT name=enviar type=button onClick="JavaScript:Validar(3);return(false);" value="&nbsp;Actualizar&nbsp;Prioridades&nbsp;" class=label></TD>
					<%end if%>
					</tr>
					</table>				
                    <%end if%>
				</td>
				<td align="right">
					<table cellspacing="1" cellpadding="0" width="200">
					<tr>
					<TD class=label align=right><b>Ver&nbsp;Itinerario:</b></TD>
					<TD class=label align=right>&nbsp;</TD>
					<TD class=label align=left>
						<select name="BLType" class=label id="Tipo de Transporte">
							<option value='-1'>Seleccionar</option>
							<%if ItineraryType=1 then%>
							<option value='0'>CONSOLIDADO</option>
							<option value='1'>EXPRESS</option>
							<%else%>
							<option value='2'>LOCAL</option>
							<!--<option value='3'>ENTREGA</option>-->
							<%end if%>
						</select>
					</TD>
					<TD class=label align=left>
						<select class="style10" name="Yr">
						<%=Yrs%>
						</select>
					</TD>
					<TD class=label align=left>
						<select name="Countries" class=label>
							<option value=''>Seleccionar</option>
							<%DisplayCountries "", 1%>
						</select>
					</TD>
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
					<TD class=label align=left colspan="10"><INPUT name=enviar type=button onClick="JavaScript:Validar(1);return(false);" value="&nbsp;Buscar&nbsp;" class=label></TD>
					<TD class=label align=left colspan="10"><INPUT name=enviar1 type=button 
                            onClick="Javascript:window.open('ItineraryPrint.asp?T=0&BT=<%=BLType%>&RT=0&YR=<%=Yr%>&Mth=<%=Mth%>&W=<%=Week%>&CT=<%=Countries%>&filter=<%=SQLFilter%>','ItineraryPrint','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750');return false;" 
                            value="&nbsp;Imprimir&nbsp;" class=label></TD>
					<%if ItineraryType=1 then%>
					<TD class=label align=left colspan="10"><INPUT name=enviar type=button onClick="Javascript:window.open('ItineraryPrint.asp?YR=<%=Yr%>&W=<%=Week%>&CT=<%=Countries%>&T=1&BT=<%=BLType%>&RT=0&filter=<%=SQLFilter%>','ItineraryPrint2','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750');return false;" value="&nbsp;Adicional&nbsp;" class=label></TD>
					<%end if%>
					<TD class=label align=left colspan="10"><INPUT name=enviar type=button onClick="Javascript:window.open('ItineraryPrint.asp?YR=<%=Yr%>&W=<%=Week%>&CT=<%=Countries%>&T=0&BT=<%=BLType%>&RT=1&filter=<%=SQLFilter%>','ItineraryPrint3','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750');return false;" value="&nbsp;Regional&nbsp;" class=label></TD>
					</tr>
					</table>				
				</td>
				</tr>
				</table>
			</TD>
		</TR> 		
		<%select case ItineraryType
		case 1%>
		<TR>
			<TD class=titlelist align=left><input class=titlelist type=checkbox name='Set' onClick='Javascript:SetAll();'></TD>
			<TD class=titlelist align=left><b>Cliente:</b></TD>
			<TD class=titlelist align=left><b>Tipo:</b></TD>
			<TD class=titlelist align=left><b>Fecha&nbsp;Descarga:</b></TD>
			<TD class=titlelist align=left><b>Descripci&oacute;n&nbsp;de&nbsp;Carga:</b></TD>
			<TD class=titlelist align=left><b>Peso:</b></TD>
			<TD class=titlelist align=left><b>CBM:</b></TD>
			<TD class=titlelist align=left><b>Bultos:</b></TD>
			<TD class=titlelist align=left><b>Origen:</b></TD>
			<TD class=titlelist align=left><b>Destino:</b></TD>
            <TD class=titlelist align=left><b>Contacto:</b></TD>
			<TD class=titlelist align=left><b>Exportador:</b></TD>
			<TD class=titlelist align=left><b>Agente:</b></TD>
			<TD class=titlelist align=left><b>Coloader:</b></TD>
			<TD class=titlelist align=left><b>Operador:</b></TD>
			<TD class=titlelist align=left><b>Contenedor:</b></TD>
			<TD class=titlelist align=left><b>BL:</b></TD>
			<TD class=titlelist align=left><b>Editar:</b></TD>
		</TR> 
		<%case 2%>
		<TR>
			<TD class=titlelist align=left><input class=titlelist type=checkbox name='Set' onClick='Javascript:SetAll();'>
			<TD class=titlelist align=left><b>Prioridad:</b></TD>
			<TD class=titlelist align=left><b>Estado:</b></TD>
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
			<TD class=titlelist align=left><b>Proveedor:</b></TD>
			<TD class=titlelist align=left><b>Unidad:</b></TD>
			<TD class=titlelist align=left><b>Revenue en $:</b></TD>
			<TD class=titlelist align=left><b>Revenue dif $:</b></TD>
			<TD class=titlelist align=left><b>Payout en $:</b></TD>
			<TD class=titlelist align=left><b>Payout dif $:</b></TD>
			<TD class=titlelist align=left><b>Profit en $:</b></TD>
			<TD class=titlelist align=left><b>Profit dif $:</b></TD>
			<TD class=titlelist align=left><b>Profit(%) en $:</b></TD>
			<TD class=titlelist align=left><b>Profit(%) dif $:</b></TD>
			<TD class=titlelist align=left><b>Editar:</b></TD>
		</TR> 
		<%end select%>		
		<%=HTMLCode%>
	</FORM>
	</TABLE>
<script>

selecciona('forma.Yr','<%=Yr%>');
selecciona('forma.BLType','<%=BLType%>');
selecciona('forma.Week','<%=Week%>');
selecciona('forma.Mth','<%=Mth%>');
selecciona('forma.Dy','<%=Dy%>');
selecciona('forma.Countries', '<%=Countries%>');
selecciona('forma.CtrsTemp', '<%=CtrsTemp%>');
</script>	
</BODY>
</HTML>
<%
	Set aTableValues = Nothing
%>