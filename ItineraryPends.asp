<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
Checking "0|1|2"

Dim Conn, Conn2, rs, rs2, i, aTableValues, CountTableValues, SBLIDS, CIDS, SIDS, TIDS, PIDS, INVIDS, HTMLCode, BLType, Countries, ItineraryType, Query, CantPIDS, SetList, BLS, Consolidado, Express, Tipo, EXID
     
	CountTableValues = -1
	SBLIDS = Request("SBLIDS") 'BLDetailIDs
	CIDS = Request("CIDS") 'ClientsIDs
	SIDS = Request("SIDS") 'AgentsIDs (Shippers)
	BLType = CheckNum(Request("BLType"))
	Countries  = Request("Countries")
	ItineraryType = CheckNum(Request("IT"))
    Tipo = Request("Tipo") 'Tipo de cada RO
    TIDS = Request("TIDS") 'BLIDTransit
    EXID = Request("EXID") 'EXID de cada RO
    
    'Al no indicarse el pais, se selecciona el primer pais asignado al usuario
    if Countries = "" then
        Countries = SetDefaultCountry
	end if

    Select Case Countries
        Case "'GT'","'GTLTF'"
            Countries = "GT','GTLTF"
        Case "'SV'","'SVLTF'"
            Countries = "SV','SVLTF"
        Case "'HN'","'HN1'","'HN2'","'HNLTF'"
            Countries = "HN','HN1','HN2','HNLTF"
        Case "'NI'","'NILTF'"
            Countries = "NI','NILTF"
        Case "'CR'","'CRLTF'"
            Countries = "CR','CRLTF"
        Case "'PA'","'PALTF'"
            Countries = "PA','PALTF"
        Case "'MX'","'MXLTF'"
            Countries = "MX','MXLTF"
    End Select
    
	'Buscando Registros pendientes para Eliminar o Asignar a un Itinerario
	Select Case ItineraryType
	Case 1 'Pendientes en Transito
		'					0			1		2			3			4		  5			    6			7		8		    9			  10			11	    	  12		13  	 14		  15		16		   17			18		 19			20			21	        22	
		Query = "select BLDetailID, Priority, Clients, BLIDTransit, BLType, DischargeDate, DiceContener, Weights, Volumes, NoOfPieces, CountryOrigen, CountriesFinalDes, Agents, Shippers, Contact, Container, BLs, CreatedDate, CreatedTime, ClientsID, AgentsID, Coloaders, (select count(SBLID) from ChargeItems where Expired=0 and InvoiceID<>0 and SBLID=BLDetailID), EXType, EXID from BLDetail where InTransit=0 and Expired=0 and BLType in (-1,0,1) and Countries in ('" & Countries & "') " & _ 
        " AND TIMESTAMPDIFF(DAY, CreatedDate , CURRENT_DATE) < 90 " & Iif(Session("PerfilColgate") = 0," AND EXType <> 99 ","") & " " & _ 
        " Order by Priority, CountriesFinalDes, Pos"
	Case 2 'Pendientes Local
		'					0			1		2			3			4		  5				6			   7			8		9		  10			11		  12		13	    14		 15		  16		   17		  18		19      20 21    22
		Query = "select BLDetailID, Priority, Clients, BLIDTransit, BLType, DischargeDate, DischargeTime, DiceContener, Weights, Volumes, NoOfPieces, CountryOrigen, Agents, Shippers, Contact, BLs, CreatedDate, CreatedTime, ClientsID, AgentsID, 0, '',  (select count(SBLID) from ChargeItems where Expired=0 and InvoiceID<>0 and SBLID=BLDetailID), EXType, EXID from BLDetail where InTransit=0 and Expired=0 and BLType in (-2,2,3) and Countries = '" & Countries & "' " & _ 
        " AND TIMESTAMPDIFF(DAY, CreatedDate , CURRENT_DATE) < 90 " & Iif(Session("PerfilColgate") = 0," AND EXType <> 99 ","") & " " & _  
        " Order by Priority, DischargeDate, DischargeTime"
	End Select
	'response.write Query & "<br>"
	
	OpenConn Conn
	if SBLIDS <> "" then
		select case CheckNum(Request("Action"))
		case 1 'Asignado Registros a un Itinerario
			'response.write("update BLDetail set Week=" & CheckNum(Request("Week")) & ", BLType=" & BLType & ", InTransit=1 where BLDetailID in (" & SBLIDS & ")")
            Conn.Execute("update BLDetail set Week=" & CheckNum(Request("Week")) & ", BLType=" & BLType & ", InTransit=1 where BLDetailID in (" & SBLIDS & ")")
			'Actualizando la BD Master indicando la fecha y tipo de servicio que realizo el cliente y el shipper
			if CIDS <> "" then
				OpenConn2 Conn2
				    'response.write "update clientes set ultima_fecha_descarga='" & ConvertDate(Now,2) & "', ultimo_tipo_movimiento=" & SetType(BLType,4) & " where id_cliente in (" & CIDS & "," & SIDS & ")<br>"
				    Conn2.Execute("update clientes set ultima_fecha_descarga='" & ConvertDate(Now,2) & "', ultimo_tipo_movimiento=" & SetType(BLType,4) & " where id_cliente in (" & CIDS & "," & SIDS & ")")
				CloseOBJ Conn2
			end if
		case 2 'Eliminando Registros a un Itinerario
			Conn.Execute("update BLDetail set Expired=" & CheckNum(Request("Expired")) & " where BLDetailID in (" & SBLIDS & ") and BLID=-1")

            'Desbloqueando el RO para que lo puedan borrar, porque ya no esta asociado al terrestre
            SBLIDS = Split(Request("SBLIDS"),",") 'BLDetail IDs
			CantPIDS = UBound(SBLIDS)

			for i=0 to CantPIDS
                Set rs = Conn.Execute("select BLs, EXID, EXType, EXDBCountry from BLDetail where BLDetailID=" & SBLIDS(i))
                if Not rs.EOF then
                    Select Case rs(2)
                    Case 4,5,6,7
                        OpenConn2 Conn2
                            Set rs2 = Conn2.Execute("select routing_seg, routing_adu from routings where routing = '" & rs(0) & "' and borrado = false ")
                            if (rs2(0) <> 0 and rs2(1) = 0) then 
                                Conn2.Execute("update routings set activo=true, bl_id=0 where routing='" & rs(0) & "'")
                                Conn2.Execute("update routings set bl_id=0 where id_routing='" & rs2(0) & "'")
                            elseif (rs2(0) = 0 and rs2(1) <> 0) then 
                                Conn2.Execute("update routings set activo=true where routing='" & rs(0) & "' and borrado=false and seguro = false")
                                Conn2.Execute("update routings set bl_id=0 where routing='" & rs(0) & "' and borrado=false")
                                Conn2.Execute("update routings set activo= true, bl_id=0 where id_routing='" & rs2(1) & "'")
                            elseif (rs2(0) <> 0 and rs2(1) <> 0) then 
                                Conn2.Execute("update routings set activo=true, bl_id=0 where routing='" & rs(0) & "'")
                                Conn2.Execute("update routings set bl_id=0 where id_routing='" & rs2(0) & "' or id_routing='" & rs2(1) & "'")
                                Conn2.Execute("update routings set activo=true where id_routing='" & rs2(1) & "'")
                            else
                                Conn2.Execute("update routings set activo=true where routing='" & rs(0) & "' and borrado=false and seguro = false")
                                Conn2.Execute("update routings set bl_id=0 where routing='" & rs(0) & "' and borrado=false")
                            end if
                        CloseOBJs Conn2, rs2
                    Case 1,2,12,13
                        openConnOcean Conn2, "ventas_" & LCase(rs(3))
                            Set rs2 = Conn2.Execute("update bill_of_lading set ref_id = 0 where bl_id = " & rs(1) & " ")
                        CloseOBJs Conn2, rs2
                    Case 0,11
                        openConnOcean Conn2, "ventas_" & LCase(rs(3))
                            Set rs2 = Conn2.Execute("update bl_completo set ref_id = 0 where bl_id = " & rs(1) & " ")
                        CloseOBJs Conn2, rs2
                    End Select
                end if
                CloseOBJ rs
            next            

		case 3 'Actualizando Prioridades en un Itinerario Local
			PIDS = Split(Request("PIDS"),",") 'Prioridades
			SBLIDS = Split(Request("SBLIDS"),",") 'BLDetail IDs
			CantPIDS = UBound(PIDS)
			for i=0 to CantPIDS
				Conn.Execute("update BLDetail set Priority=" & PIDS(i) & " where BLDetailID = " & SBLIDS(i))
			next
		end select
		SBLIDS = ""
	end if
	
    'response.write Query & "<br>"
	Set rs = Conn.Execute(Query)
	If Not rs.EOF Then
		aTableValues = rs.GetRows
		CountTableValues = rs.RecordCount-1
	End If
	closeOBJs rs, Conn

	CIDS = ""
	SIDS = ""
	Select Case ItineraryType
	Case 1 
		for i=0 to CountTableValues
			HTMLCode = HTMLCode & "<tr>" & _
            "<td class=style4 align=center bgcolor=#999999>" & _
                "<a href=# onClick=JavaScript:Editar(" & aTableValues(0,i) & ",'" & aTableValues(17,i) & "'," & CheckNum(aTableValues(18,i)) & ");return (false); class=submenu><font color=FFFFFF><b>&nbsp;Editar&nbsp;</b></font></a>" & _ 
                "<br>" &  aTableValues(0,i) & _
                "<br><a href=# onClick=JavaScript:Desglose(" & aTableValues(0,i) & ",'" & aTableValues(17,i) & "'," & CheckNum(aTableValues(18,i)) & ");return (false); class=submenu1><font color=FFFFFF><b>&nbsp;Desglose&nbsp;</b></font></a>" & _ 
            "</td>" & _
            "<td class=label>"
            
            'Validando si el Cliente tiene dato asignado, Si lo tiene se puede asignar a Itinerario
            'si no lo tiene muestra aviso Pend. y no permite asignar a Itinerario
            if Trim(aTableValues(2,i))="" then
                SBLIDS = SBLIDS & "SBLIDS[" & i & "]=0;" & vbCrLf
                Tipo = Tipo & "Tipo[" & i & "]=0;" & vbCrLf
                INVIDS = INVIDS & "INVIDS[" & i & "]=0;" & vbCrLf
                HTMLCode = HTMLCode & "<input type=checkbox name='Pos" & i & "' disabled><font color=red><b>PEND.</b><font>"
            else
                SBLIDS = SBLIDS & "SBLIDS[" & i & "]=" & CheckNum(aTableValues(0,i)) & ";" & vbCrLf
                INVIDS = INVIDS & "INVIDS[" & i & "]=" & CheckNum(aTableValues(22,i)) & ";" & vbCrLf
                Tipo = Tipo & "Tipo[" & i & "]=" & CheckNum(aTableValues(23,i)) & ";" & vbCrLf
                EXID = EXID & "EXID[" & i & "]=" & CheckNum(aTableValues(24,i)) & ";" & vbCrLf
                if CheckNum(aTableValues(3,i)) > 0 then
                    HTMLCode = HTMLCode & "<input type=checkbox name='Pos" & i & "'><font color=brown><b>TRANSIT.</b><font>"
                else
                    HTMLCode = HTMLCode & "<input type=checkbox name='Pos" & i & "'>"
                end if
                SetList = "onclick=Javascript:SetList('Pos" & i & "');"
            end if
            
            HTMLCode = HTMLCode & "</td>" & _
				"<td class=label><a href=# class=label><select name='S" & i & "' class=label>" & SetPriority(aTableValues(1,i),1) & "</select></a></td>" & _
				"<td class=label><a href=# class=label >" & aTableValues(2,i) & "</a></td>" & _
				"<td class=label align=center><a href=# class=label " & SetList & ">" & SetTypeItinerary(aTableValues(3,i),aTableValues(4,i)) & "</a></td>" & _
				"<td class=label><a href=# class=label " & SetList & ">" & aTableValues(5,i) & "</a></td>" & _
				"<td class=label><a href=# class=label " & SetList & ">" & aTableValues(6,i) & "</a></td>" & _
				"<td class=label><a href=# class=label " & SetList & ">" & aTableValues(7,i) & "</a></td>" & _
				"<td class=label><a href=# class=label " & SetList & ">" & aTableValues(8,i) & "</a></td>" & _
				"<td class=label><a href=# class=label " & SetList & ">" & aTableValues(9,i) & "</a></td>" & _
				"<td class=label><a href=# class=label " & SetList & ">" & aTableValues(10,i) & "</a></td>" & _
				"<td class=label><a href=# class=label " & SetList & ">" & aTableValues(11,i) & "</a></td>" & _
				"<td class=label><a href=# class=label " & SetList & ">" & aTableValues(12,i) & "</a></td>" & _
				"<td class=label><a href=# class=label " & SetList & ">" & aTableValues(13,i) & "</a></td>" & _
				"<td class=label><a href=# class=label " & SetList & ">" & aTableValues(21,i) & "</a></td>" & _
				"<td class=label><a href=# class=label " & SetList & ">" & aTableValues(14,i) & "</a></td>" & _
				"<td class=label><a href=# class=label " & SetList & ">" & aTableValues(15,i) & "</a></td>" & _
				"<td class=label><a href=# class=label " & SetList & ">" & aTableValues(16,i) & "</a></td>" & _
				"</tr>" & _
				"<tr><td class=submenu colspan=18></td></tr>"
			CIDS = CIDS & "CIDS[" & i & "]=" & aTableValues(19,i) & ";" & vbCrLf 'Ids de Clientes para indicar la fecha y tipo de servicio contratado
			SIDS = SIDS & "SIDS[" & i & "]=" & aTableValues(20,i) & ";" & vbCrLf 'Ids de Shippers para indicar la fecha y tipo de servicio contratado
			TIDS = TIDS & "TIDS[" & i & "]=" & aTableValues(3,i) & ";" & vbCrLf 'Tipo de Carga (Transito<>0 o Nueva=0)
		next
	Case 2
		for i=0 to CountTableValues
			SBLIDS = SBLIDS & "SBLIDS[" & i & "]=" & CheckNum(aTableValues(0,i)) & ";" & vbCrLf
            Tipo = Tipo & "Tipo[" & i & "]=" & CheckNum(aTableValues(23,i)) & ";" & vbCrLf
            INVIDS = INVIDS & "INVIDS[" & i & "]=" & CheckNum(aTableValues(21,i)) & ";" & vbCrLf
            EXID = EXID & "EXID[" & i & "]=" & CheckNum(aTableValues(24,i)) & ";" & vbCrLf
			HTMLCode = HTMLCode & "<tr>" & _
                "<td class=style4 align=center bgcolor=#999999>" & _
                    "<a href=# onClick=JavaScript:Editar(" & aTableValues(0,i) & ",'" & aTableValues(16,i) & "'," & CheckNum(aTableValues(17,i)) & ");return (false); class=submenu><font color=FFFFFF><b>&nbsp;Editar&nbsp;</b></font></a>" & _
                    "<br><br><a href=# onClick=JavaScript:Desglose(" & aTableValues(0,i) & ",'" & aTableValues(16,i) & "'," & CheckNum(aTableValues(17,i)) & ");return (false); class=submenu1><font color=FFFFFF><b>&nbsp;Desglose&nbsp;</b></font></a>" & _
                "</td>" & _
                "<td class=label><input type=checkbox name='Pos" & i & "'></td>" & _
				"<td class=label><a href=# class=label><select name='S" & i & "' class=label>" & SetPriority(aTableValues(1,i),2) & "</select></a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(2,i) & "</a></td>" & _
				"<td class=label align=center><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & SetTypeItinerary(aTableValues(3,i),aTableValues(4,i)) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(5,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(6,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(7,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(8,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(9,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(10,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(11,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(12,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(13,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(14,i) & "</a></td>" & _
				"<td class=label><a href=# class=label onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(15,i) & "</a></td>" & _
				"</tr>" & _
				"<tr><td class=submenu colspan=16></td></tr>"
			CIDS = CIDS & "CIDS[" & i & "]=" & aTableValues(18,i) & ";" & vbCrLf 'Ids de Clientes para indicar la fecha y tipo de servicio contratado
			SIDS = SIDS & "SIDS[" & i & "]=" & aTableValues(19,i) & ";" & vbCrLf 'Ids de Shippers para indicar la fecha y tipo de servicio contratado
			TIDS = TIDS & "TIDS[" & i & "]=" & aTableValues(3,i) & ";" & vbCrLf 'Tipo de Carga (Transito<>0 o Nueva=0)
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
<SCRIPT LANGUAGE="JavaScript">
var SBLIDS = new Array();
var CIDS = new Array();
var SIDS = new Array();
var TIDS = new Array();
var PIDS = new Array();
var INVIDS = new Array();
var Tipo = new Array();
var EXID = new Array();
<%=SBLIDS%>
<%=CIDS%>
<%=SIDS%>
<%=TIDS%>
<%=INVIDS%>
<%=Tipo%>
<%=EXID%>

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
			if (SBLIDS[i] != 0) {
                document.forma.elements["Pos" + i].checked = true;
            }
		}
	} else {
		for (var i=0; i<<%=i%>; i++) {
			document.forma.elements["Pos" + i].checked = false;
		}

	}
}

function SetIDs() {
	var sep = "";
	var sep2 = "";
	document.forma.SBLIDS.value = "";
    document.forma.TIDS.value = "";
	document.forma.CIDS.value = "";
	document.forma.SIDS.value = "";
    document.forma.Tipo.value = "";
    document.forma.EXID.value = "";
	for (var i=0; i<<%=i%>; i++) {
		if (document.forma.elements["Pos" + i].checked) {
			document.forma.SBLIDS.value = document.forma.SBLIDS.value + sep + SBLIDS[i];
            document.forma.Tipo.value = document.forma.Tipo.value + sep + Tipo[i];
            document.forma.TIDS.value = document.forma.TIDS.value + sep + TIDS[i];
            if (document.forma.BLType.value == 0){
                if (Tipo[i] == 5){
                    alert("No es posible asignar la línea " + (i + 1) + " en carga Consolidada, ya que es Express.");
                    document.forma.elements["Pos" + i].checked = false;
                    returnToPreviousPage();
                }
            }
            else if (document.forma.BLType.value == 1){
                if (Tipo[i] == 4){
                    alert("No es posible asignar la línea " + (i + 1) + " en carga Express, ya que es Consolidada.");
                    document.forma.elements["Pos" + i].checked = false;
                    returnToPreviousPage();
                }
            }
            document.forma.Tipo.value = document.forma.Tipo.value + sep + Tipo[i];
            document.forma.EXID.value = document.forma.EXID.value + sep + EXID[i];
			if (TIDS[i]==0) {
				document.forma.CIDS.value = document.forma.CIDS.value + sep2 + CIDS[i];
				document.forma.SIDS.value = document.forma.SIDS.value + sep2 + SIDS[i];
				sep2 = ","
			}			
			sep = ","
		}
	}
}

function SetPIDs() {
	var sep = "";
	document.forma.SBLIDS.value = "";
	document.forma.PIDS.value = "";
    document.forma.Tipo.value = "";
    document.forma.EXID.value = "";
	for (var i=0; i<<%=i%>; i++) {
		document.forma.SBLIDS.value = document.forma.SBLIDS.value + sep + SBLIDS[i];
        document.forma.Tipo.value = document.forma.Tipo.value + sep + Tipo[i];
        document.forma.EXID.value = document.forma.EXID.value + sep + EXID[i];
		document.forma.PIDS.value = document.forma.PIDS.value + sep + document.forma.elements["S" + i].value;
		sep = ","
	}
}

function IData(GID){
	if (GID==37) {
		window.open('Search_ResultsAdmin.asp?GID='+GID+'&CTR=<%=Countries%>&IT=<%=ItineraryType%>','SIData37','height=330,width=750,menubar=0,resizable=1,scrollbars=1,toolbar=0');
	} 
	if (GID==27) {
		window.open('Search_Admin.asp?GID='+GID+'&CTR=<%=Countries%>&IT=<%=ItineraryType%>','SIData27','height=330,width=750,menubar=0,resizable=1,scrollbars=1,toolbar=0');
	} 
    if (GID==28) {
    	window.open('Search_BLData.asp?GID='+GID+'&CTR=<%=Countries%>&IT=<%=ItineraryType%>','IData28','height=330,width=660,menubar=0,resizable=1,scrollbars=1,toolbar=0');
	}
    if (GID==33) {
        <%if ItineraryType=1 then 'se coloca ET=8 %>
    	window.open('InsertData.asp?GID='+GID+'&ET=8&OID=0&CTR=<%=Countries%>&CTR2=<%=Countries%>','IData33','height=330,width=660,menubar=0,resizable=1,scrollbars=1,toolbar=0');
        <%else %>
        window.open('InsertData.asp?GID='+GID+'&ET=15&OID=0&CTR=<%=Countries%>&CTR2=<%=Countries%>','IData33','height=330,width=660,menubar=0,resizable=1,scrollbars=1,toolbar=0');
        <%end if %>
	}   
}

function Editar(OID, CD, CT){
	window.open('InsertData.asp?GID=27&OID='+OID+'&CD='+CD+'&CT='+CT,'EData','height=650,width=600,menubar=0,resizable=1,scrollbars=1,toolbar=0');
}

function Desglose(OID, CD, CT){
	window.open('InsertData.asp?GID=36&OID='+OID+'&CD='+CD+'&CT='+CT,'EDataD','height=650,width=600,menubar=0,resizable=1,scrollbars=1,toolbar=0');
}


function MaritimTransit(){
	window.open('MaritimTransit.asp','MaritimTransit','height=650,width=980,menubar=0,resizable=1,scrollbars=1,toolbar=0');
}

function Validar(Action) {
    SetIDs();
	if (Action == 1) {
        if (document.forma.SBLIDS.value == "") {
		    alert("Por favor seleccione al menos un registro para Asignar al Itinerario");
			return (false);
		};
		if (!valTxt(document.forma.Week, 1, 5)){return (false)};
		if (!valSelec(document.forma.BLType)){return (false)};
	}
	if (Action == 2) {
        if (document.forma.SBLIDS.value == "") {
		    alert("Por favor seleccione al menos un registro para Eliminar");
			return (false);
		};
        //Se revisa que los IDs seleccionados para eliminar no tenga factura relacionada
        var DelSBLID = document.forma.SBLIDS.value.split(",");
        var DelTipo = document.forma.Tipo.value.split(",");
        var DelTIDS = document.forma.TIDS.value.split(",");
        for(i=0; i<DelSBLID.length; i++) {
            for(j=0; j<<%=i%>; j++) {
                if (DelSBLID[i]==SBLIDS[j]) {
                    switch (Tipo[j]) {
                        case 4:
                        case 5:
                        case 6:
                        case 7:
                            window.open('RoutingError.asp?id_trafico=4&id_usuario=<%=Session("OperatorID")%>&id_routing=' + EXID[j], 'prueba' + [j], 'height=420,width=530,menubar=0,resizable=1,scrollbars=1,toolbar=0');
                            
                    }
                    if (INVIDS[j]>0) {
                        alert("No puede borrar uno de los registros porque tiene facturas relacionadas, si desea hacerlo primero debe anular las facturas respectivas");
                        alert(INVIDS[j])
			            return (false);
                    };
                };
            };
            if (DelTIDS[i] > 0)
            {
                alert("No se puede eliminar uno de los registros ya que es una carga en tránsito.");
                return (false);
            }
        }
        document.forma.Expired.value = 1;
	}
	if (Action == 3) {
		SetPIDs();
	}
	
    document.forma.Action.value = Action;
    document.forma.submit();
    if (Action == 2 || Action == 3)
    {
        location.reload();
    }
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
		<FORM name="forma" action="ItineraryPends.asp" method="post">
		<INPUT name="IT" type=hidden value=<%=ItineraryType%>>
		<INPUT name="Expired" type=hidden value=0>
		<INPUT name="Action" type=hidden value=0>
		<input type="hidden" name="SBLIDS" value="">
        <input type="hidden" name="TIDS" value="">
		<input type="hidden" name="CIDS" value="">
		<input type="hidden" name="SIDS" value="">
		<input type="hidden" name="PIDS" value="">
        <input type="hidden" name="Tipo" value="">
        <input type="hidden" name="EXID" value="">
		<TR>
			<TD class=label align=left colspan="17">
				<table cellspacing="0" cellpadding="0" width="100%">
				<tr>
				<td align="left" valign=top>
					<table cellspacing="1" cellpadding="0" width="200">
					<tr>
					<%if Len(Session("Countries"))>6 then%>
					<TD class=label align=left>
						<select name="Countries" id="Country" class=label>
							<option value=''>Seleccionar</option>
							<%DisplayCountries "", 4%>
						</select>
					</TD>
					<TD class=label align=left olspan="10"><INPUT name=enviar type=button onClick="JavaScript:Validar(0);return(false);" value="&nbsp;Ver&nbsp;" class=label>&nbsp;&nbsp;&nbsp;</TD>
					<%else%>
					<input type="hidden" name="Countries" value="<%=Countries%>">
					<%end if%>
					<TD class=label align=left colspan="10">&nbsp;<INPUT name=enviar type=button onClick="JavaScript:Validar(3);return(false);" value="Actualizar&nbsp;Seleccionados" class=label>&nbsp;</TD>
					<TD class=label align=left colspan="10">&nbsp;<INPUT name=enviar type=button onClick="JavaScript:Validar(2);return(false);" value="Eliminar&nbsp;Seleccionados" class=label>&nbsp;</TD>
					<TD class=label align=left>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>Semana:</b></TD>
					<TD class=label align=left><input class="style10" name="Week" value="" id="Semana" type="text" size="5" maxlength="2" onKeyUp="res(this,numb);"></TD>
					<TD class=label align=left>
						<select name="BLType" class=label id="Tipo de Transporte">
						<option value='-1'>Seleccionar</option>
						<%Select Case ItineraryType
						Case 1%>
						<option value='0'>CONSOLIDADO</option>
						<option value='1'>EXPRESS</option>
						<%Case 2%>
						<option value='2'>LOCAL</option>
						<!--<option value='3'>ENTREGA</option>-->
						<%End Select%>
						</select>
					</TD>
					<TD class=label align=left colspan="10"><INPUT name=enviar type=button onClick="JavaScript:Validar(1);return(false);" value="Asignar&nbsp;Seleccionados" class=label></TD>
					</tr>
					</table>				
				</td>
				<td align="right" valign=top>
					<table cellspacing="1" cellpadding="4" border=0>
					<tr>
						<td class=label align=left><b>Agregar:</b></td>

                        <% if Session("PerfilColgate") = 1 then %>
                        <td class="style4" align="right" bgcolor="#999999">
						<a href="#" onClick="Javascript:IData(37);return (false);" class="submenu"><font color="FFFFFF">&nbsp;Carga&nbsp;Colgate&nbsp;</font></a>
						</td>
                        <% end if %>

						<td class="style4" align="right" bgcolor="#999999">
						<a href="#" onClick="Javascript:IData(27);return (false);" class="submenu"><font color="FFFFFF">&nbsp;Carga&nbsp;en&nbsp;Tr&aacute;nsito&nbsp;</font></a>
						</td>
						<td class="style4" align="right" bgcolor="#999999">
						<a href="#" onClick="Javascript:IData(28);return (false);" class="submenu"><font color="FFFFFF">&nbsp;Carga&nbsp;General&nbsp;</font></a>
						</td>
                        <%if isValidCIF then%>
                        <td class="style4" align="right" bgcolor="#999999">
						<a href="#" onClick="Javascript:IData(33);return (false);" class="submenu"><font color="FFFFFF">&nbsp;Carga&nbsp;CIF&nbsp;</font></a>
						</td>
                        <%end if %>
					</tr>
                    <tr>
                    <td colspan=2></td>
                    <td colspan=3 align=center>
                    <INPUT name=enviar type=button onClick="JavaScript:MaritimTransit();return(false);" value="Carga&nbsp;en&nbsp;Transito" class=label>	
                    </td>
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
        	<TD class=titlelist align=left><b>Accion:</b></TD>
			<TD class=titlelist align=left style="width:5px"><input class=titlelist type=checkbox name='Set' onClick='Javascript:SetAll();'>
			<TD class=titlelist align=left><b>Semana:</b></TD>
			<TD class=titlelist align=left><b>Cliente:</b></TD>
			<TD class=titlelist align=left><b>Tipo:</b></TD>
			<TD class=titlelist align=left style="width:5px"><b>Fecha Descarga:</b></TD>
			<TD class=titlelist align=left><b>Descripci&oacute;n&nbsp;de&nbsp;Carga:</b></TD>
			<TD class=titlelist align=left><b>Peso:</b></TD>
			<TD class=titlelist align=left><b>CBM:</b></TD>
			<TD class=titlelist align=left><b>Bultos:</b></TD>
			<TD class=titlelist align=left><b>Origen:</b></TD>
			<TD class=titlelist align=left><b>Destino:</b></TD>
			<TD class=titlelist align=left><b>Exportador:</b></TD>
			<TD class=titlelist align=left><b>Agente:</b></TD>
			<TD class=titlelist align=left><b>Coloader:</b></TD>
			<TD class=titlelist align=left><b>Contacto:</b></TD>
			<TD class=titlelist align=left><b>Contenedor:</b></TD>
			<TD class=titlelist align=left><b>BL:</b></TD>
		</TR> 
		<%case 2%>
		<TR>
        	<TD class=titlelist align=left><b>Accion:</b></TD>
			<TD class=titlelist align=left><input class=titlelist type=checkbox name='Set' onClick='Javascript:SetAll();'>
			<TD class=titlelist align=left><b>Prioridad:</b></TD>
			<TD class=titlelist align=left><b>Cliente:</b></TD>
			<TD class=titlelist align=left><b>Tipo:</b></TD>
			<TD class=titlelist align=left><b>Fecha&nbsp;Programada:</b></TD>
			<TD class=titlelist align=left><b>Hora&nbsp;Programada:</b></TD>
			<TD class=titlelist align=left><b>Descripci&oacute;n&nbsp;de&nbsp;Carga:</b></TD>
			<TD class=titlelist align=left><b>Peso:</b></TD>
			<TD class=titlelist align=left><b>CBM:</b></TD>
			<TD class=titlelist align=left><b>Bultos:</b></TD>
			<TD class=titlelist align=left><b>Origen:</b></TD>
			<TD class=titlelist align=left><b>Exportador:</b></TD>
			<TD class=titlelist align=left><b>Agente:</b></TD>
			<TD class=titlelist align=left><b>Contacto:</b></TD>
			<TD class=titlelist align=left><b>RO:</b></TD>
		</TR> 
		<%end select%>		
		<%=HTMLCode%>
	</FORM>
	</TABLE>
<%if Len(Session("Countries"))>6 then%>
<script>
    selecciona('forma.Countries', '<%=Countries%>');
</script>	
<%end if%>
</BODY>
</HTML>
<%
	Set aTableValues = Nothing
%>