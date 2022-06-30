<%
Checking "0|1|2"

Dim WeekPrev2, WeekPrev, WeekAct, WeekPost, WeekPost2, YearPrev2, YearPrev, YearAct, YearPost, YearPost2


if Action <> 3 then
	if CountTableValues >= 0 then
        BLID = aTableValues(0,0)
		CreatedDate = aTableValues(1,0)
		CreatedTime = aTableValues(2,0)
		Expired = aTableValues(3,0)
		BLNumber = aTableValues(4,0)
		SenderData = aTableValues(5,0)
		ShipperData = aTableValues(6,0)
		ConsignerData = aTableValues(7,0)
		CountryDep = aTableValues(8,0)
		Attn = aTableValues(9,0)
		HandlingInformation = aTableValues(10,0)
		BrokerID = aTableValues(11,0)
		CountryDes = aTableValues(12,0)
		ChargePlace = aTableValues(13,0)
		FinalDes = aTableValues(14,0)
		PilotID = aTableValues(15,0)
		TruckID = aTableValues(16,0)
		Container = aTableValues(17,0)
		TotNoOfPieces = aTableValues(18,0)
		TotWeight = aTableValues(19,0)
		TotVolume = aTableValues(20,0)
		TotPrepaid = aTableValues(21,0)
		TotCollect = aTableValues(22,0)
		BLsFreight = aTableValues(23,0)
		BLsInsurance = aTableValues(24,0)
		BLsAnotherChargesCollect = aTableValues(25,0)
		Observations = aTableValues(26,0)
		ContactSignature = aTableValues(27,0)
		BLExitDate = aTableValues(28,0)
		Countries = aTableValues(29,0)
		Week = aTableValues(30,0)
		ContainerDep = aTableValues(31,0)
		BLType = aTableValues(32,0)
		Consolidated = aTableValues(33,0)
		LtAcceptDate = aTableValues(34,0)
		SenderID = aTableValues(35,0)
		ShipperID = aTableValues(36,0)
		ConsignerID = aTableValues(37,0)
		SenderAddrID = aTableValues(38,0)
		ShipperAddrID = aTableValues(39,0)
		ConsignerAddrID = aTableValues(40,0)
		Chassis = aTableValues(41,0)
		HasSecure = aTableValues(42,0)
		UnitNo = aTableValues(43,0)
		BLSize = aTableValues(44,0)
		Marchamo = aTableValues(45,0)
		BLsFreight2 = aTableValues(46,0)
		BLsInsurance2 = aTableValues(47,0)
		BLsAnotherChargesPrepaid = aTableValues(48,0)
		ChargeType = aTableValues(49,0)
		DestinyType = aTableValues(50,0)
		Closed = aTableValues(51,0)
		BLEstArrivalDate = aTableValues(52,0)
		DTIObservations = aTableValues(53,0)
		TotDiceContenerValue = aTableValues(54,0)
		Bail = aTableValues(55,0)
		Comment2 = aTableValues(56,0)
		DTI = aTableValues(57,0)
		Comment4 = aTableValues(58,0)
		PolicyNo = aTableValues(59,0)
		DeliveryPolicyDate = aTableValues(60,0)
		DeliveryPolicyHour = aTableValues(61,0)
		DeliveryPolicyMin = aTableValues(62,0)
		BLDispatchDate = aTableValues(63,0)
        PilotInstructions = aTableValues(64,0)
        ClientColoader = aTableValues(65,0)
        ShipperColoader = aTableValues(66,0)
        AgentNeutral = aTableValues(67,0)
        GuiaRemision = aTableValues(68,0)
        BLArrivalDate = aTableValues(69,0)
	else
		CreatedDate = Request.Form("CreatedDate")
		CreatedTime = Request.Form("CreatedTime")
		Expired = checkNum(Request.Form("Expired"))
		BLNumber = Request.Form("BLNumber")
		SenderData = Request.Form("SenderData")
		ShipperData = Request.Form("ShipperData")
		ConsignerData = Request.Form("ConsignerData")
		CountryDep = Request.Form("CountryDep")
		Attn = Request.Form("Attn")
		HandlingInformation = Request.Form("HandlingInformation")
		BrokerID = checkNum(Request.Form("BrokerID"))
		CountryDes = Request.Form("CountryDes")
		ChargePlace = Request.Form("ChargePlace")
		FinalDes = Request.Form("FinalDes")
		PilotID = checkNum(Request.Form("PilotID"))
		TruckID = checkNum(Request.Form("TruckID"))
		Container = Request.Form("Container")	
		TotNoOfPieces = Request.Form("TotNoOfPieces")
		TotWeight = Request.Form("TotWeight")
		TotVolume = Request.Form("TotVolume")
		TotPrepaid = Request.Form("TotPrepaid")
		if TotPrepaid = "" then TotPrepaid = "0.00" end if
		TotCollect = Request.Form("TotCollect")
		if TotCollect = "" then TotCollect = "0.00" end if
		BLsFreight = Request.Form("BLsFreight")
		if BLsFreight = "" then BLsFreight = "0.00" end if
		BLsInsurance = Request.Form("BLsInsurance")
		if BLsInsurance = "" then BLsInsurance = "0.00" end if
		BLsAnotherChargesCollect = Request.Form("BLsAnotherChargesCollect")
		if BLsAnotherChargesCollect = "" then BLsAnotherChargesCollect = "0.00" end if
		Observations = Request.Form("Observations")
		ContactSignature = Session("OperatorName")
		BLExitDate = Request.Form("BLExitDate")
		Countries = Request.Form("Countries")
		if Countries = "" then
            Countries = SetDefaultCountry
		end if
		Week = Request.Form("Week")
		ContainerDep = Request.Form("ContainerDep")
		BLType = Request.Form("BLType")
		
		Consolidated = 1
		
		LtAcceptDate = null ' Request.Form("LtAcceptDate")
		SenderID = Request.Form("SenderID")
		ShipperID = Request.Form("ShipperID")
		ConsignerID = Request.Form("ConsignerID")
		SenderAddrID = Request.Form("SenderAddrID")
		ShipperAddrID = Request.Form("ShipperAddrID")
		ConsignerAddrID = Request.Form("ConsignerAddrID")
		Chassis = Request.Form("Chassis")
		HasSecure = Request.Form("HasSecure")
		UnitNo = Request.Form("UnitNo")
		BLSize = Request.Form("BLSize")
		Marchamo = Request.Form("Marchamo")
		BLsFreight2 = "0.00"
		BLsInsurance2 = "0.00"
		BLsAnotherChargesPrepaid = "0.00"
		Closed = 0
		BLEstArrivalDate = Request.Form("BLEstArrivalDate")
		DTIObservations = Request.Form("DTIObservations")
		TotDiceContenerValue = Request.Form("TotDiceContenerValue")
		Bail = Request.Form("Bail")
		Comment2 = Request.Form("Comment2")
		DTI = Request.Form("DTI")
		Comment4 = Request.Form("Comment4")
		PolicyNo = Request.Form("PolicyNo")
		DeliveryPolicyDate = Request.Form("DeliveryPolicyDate")
		DeliveryPolicyHour = Request.Form("DeliveryPolicyHour")
		DeliveryPolicyMin = Request.Form("DeliveryPolicyMin")
		BLDispatchDate = Request.Form("BLDispatchDate")
        PilotInstructions = Request.Form("PilotInstructions")
        if Request("Action")="" then
            PilotInstructions="COMUNICACION DE ESTATUS CADA 3 HORAS A INTERMODAL TERRESTRE DE AIMAR EN PAIS EN EL QUE SE ENCUENTRE" & Chr(13) & Chr(10) & _
            "ALERTA DE LLEGADA DE UNIDAD A FRONTERA" & Chr(13) & Chr(10) & _ 
            "ALERTA DE SALIDA DE ADUANA DE PAIS DE ORIGEN Y ENTRADA A ADUANA DE PAIS DE DESTINO" & Chr(13) & Chr(10) & _ 
            "CONTACTAR A LOS SENORES DE LA AGENCIA ADUANAL SANTOS EN FRONTERA LADO DE PANAMA PARA RETIRAR PERMISOS PARA PRODUCTOS VETERINARIOS Y MADERA"
        end if
	end if
end if



if Action = 10 then 

        'response.write "(" & Action & ")(" & Request.Form("ConsignerID") & ")(" & Request.Form("ConsignerData") & ")"

		ConsignerData = Request.Form("ConsignerData")

		Attn = Request.Form("Attn")

		ConsignerID = Request.Form("ConsignerID")

		ConsignerAddrID = Request.Form("ConsignerAddrID")

		'ConsignerColoader = Request.Form("ConsignerColoader")

end if


Set aTableValues = Nothing

if Request("IT") <> "" then
	ItineraryType = Request("IT")
else
	select case BLType
	case 0,1 'Transito
		ItineraryType = 1
	case 2,3 'Local
		ItineraryType = 2
	end select
end if

CountList1Values = -1
CountList2Values = -1
CountList3Values = -1
CountList4Values = -1
CountList5Values = -1
CountList6Values = -1


    fleet = GroupData(ConsignerID)

    'response.write "(" & fleet & ")<br>"

OpenConn Conn
    JavaMsg = ""
	'Guardando el Detalle del BL
	If Action = 1 or Action = 2 Then

        '///////////////////////////////////////////////////////////////////////////////////
        'response.write "(" & Action & ")"
        '///////////////////////////////////////////////////////////////////////////////////

		ntr = chr(13) & chr(10)
        'response.write request.form("Clients") & "<br>"
		DetailToErase Split(request.form("DetailToErase"), "|", -1, 1)
		JavaMsg = DetailToUpdate(Split(request.form("BLDetailID"), "|", -1, 1), Split(request.form("Pos"), "|",  -1, 1), Split(request.form("BLIDTransit"), "|",  -1, 1), Split(request.form("NoOfPieces"), ntr,  -1, 1), Split(request.form("ClassNoOfPieces"), ntr,  -1, 1), Split(request.form("CommoditiesID"), "|",  -1, 1), Split(request.form("DiceContener"), ntr,  -1, 1), Split(request.form("DiceContenerValue"), ntr,  -1, 1), Split(request.form("Volumes"), ntr,  -1, 1), Split(request.form("Weights"), ntr,  -1, 1), Split(request.form("ClientsID"), "|",  -1, 1), Split(request.form("AddressesID"), "|",  -1, 1), Split(request.form("Clients"), ntr,  -1, 1), Split(request.form("BLs"), ntr,  -1, 1), Split(request.form("DischargeDate"), ntr,  -1, 1), Split(request.form("CountriesFinalDes"), ntr,  -1, 1), Split(request.form("InTransit"), "|",  -1, 1), Split(request.form("AgentsID"), "|",  -1, 1), Split(request.form("Agents"), ntr,  -1, 1), Split(request.form("HBLNumber"), "|",  -1, 1), Split(request.form("AgentsAddrID"), "|",  -1, 1), BLNumber, Split(request.form("Seps"), ntr, -1, 1), Split(request.form("CountriesOrigen"), ntr,  -1, 1), Countries)
	
    End If	



    if Action = 9 then 'validar placa

        QuerySelect = "select case when TruckGps = 1 then '(GPS)' else '' end, case when TruckAvailable = 1 then '(ON)' else '(OFF)' end, a.TruckNo, IFNULL(b.BLNumber,'') from Trucks a left join BLs b ON a.TruckID = b.TruckID where a.TruckID = " & TruckID & " order by BLID desc LIMIT 1"
        'response.write QuerySelect & "<br>"
	    Set rs = Conn.Execute(QuerySelect)
	    If Not rs.EOF Then
   	       
            On Error Resume Next      
        
                if rs(1) = "(OFF)" and rs(0) = "(GPS)" then
                     response.write  "<script>alert('Placa " & Replace(Replace(rs(2),"(GPS)",""),"(ON)","") & " ya esta en ruta"
                    if rs(3) <> "" then
                         response.write  ", fue cargada en " & rs(3) 
                    end if
                    response.write "');</script>"   'document.getElementById('Cabezal').value='';
                    TruckID = ""
                end if
               
            If Err.Number<>0 then
                response.write "Validar Placa :" & Err.Number & " - " & Err.Description & "<br>"  
            end if

        End If
	    CloseOBJ rs
    end if

	'Obteniendo listado de Aduanas
	'Set rs = Conn.Execute("select BrokerID, Name, Countries from Brokers where Expired = 0 and Countries in " & Session("Countries") & " order by Countries, Name")
	Set rs = Conn.Execute("select BrokerID, Name, Countries from Brokers where Expired = 0 order by Name, Countries")
	If Not rs.EOF Then
   		aList1Values = rs.GetRows
       	CountList1Values = rs.RecordCount-1
    End If
	CloseOBJ rs

	'Obteniendo listado de Pilotos
	'Set rs = Conn.Execute("select PilotID, Name, License, Countries from Pilots where Expired = 0 and Countries in " & Session("Countries") & " order by Countries, Name")
	Set rs = Conn.Execute("select PilotID, Name, License, Countries from Pilots where Expired = 0 order by Name, Countries")
	If Not rs.EOF Then
   		aList2Values = rs.GetRows
       	CountList2Values = rs.RecordCount-1
    End If
	CloseOBJ rs

	'Obteniendo listado de Cabezales
	'Set rs = Conn.Execute("select TruckID, TruckNo, Countries from Trucks where Expired = 0 and Countries in " & Session("Countries") & " order by Countries, TruckNo")

    


    QuerySelect = "select TruckID, TruckNo, Countries, TruckType, Mark, Model, UPPER(Replace(Replace(TruckNo, '-', ''), ' ', '')), case when TruckGps = 1 then '(GPS)' else '' end, case when TruckAvailable = 1 then '(ON)' else '(OFF)' end FROM Trucks WHERE Expired = 0 "
    
    'if fleet <> "NA" AND fleet <> "" then

    'QuerySelect = QuerySelect & " AND UPPER(Replace(Replace(TruckNo, '-', ''),' ','')) NOT IN ( " & _
    '"SELECT DISTINCT UPPER(Replace(Replace(nombre, '-', ''),' ','')) FROM disatel_json " & _
    '"WHERE activo = 1 AND IFNULL(geocerca,'') = '' AND IFNULL(geocerca_mov,'') = '') " 

    'end if

    QuerySelect = QuerySelect & " ORDER BY UPPER(Replace(Replace(TruckNo, '-', ''),' ','')), TruckType, Countries"


	Set rs = Conn.Execute(QuerySelect)
	If Not rs.EOF Then
   		aList3Values = rs.GetRows
       	CountList3Values = rs.RecordCount-1
    End If
	CloseOBJ rs

	'Obteniendo listado de Bodegas
	Set rs = Conn.Execute("select WarehouseID, Countries, Name, Address, Address2, Phone1, Phone2, Attn from Warehouses where Expired=0 order by Name, Countries")
	If Not rs.EOF Then
   		aList4Values = rs.GetRows
       	CountList4Values = rs.RecordCount-1
    End If
	CloseOBJ rs
	
	'Obteniendo Detalle del BL
	'                                 0            1            2            3                4             5                6             7        8         9           10         11    12         13               14             15        16       17       18           19        20         21                                                     22                                                   23       24         25
    'response.write("select BLDetailID, BLIDTransit, NoOfPieces, ClassNoOfPieces, CommoditiesID, DiceContener, DiceContenerValue, Volumes, Weights, ClientsID, AddressesID, Clients, BLs, DischargeDate, CountriesFinalDes, InTransit, AgentsID, Agents, HBLNumber, AgentsAddrID, Seps, CountryOrigen, (select count(SBLID) from ChargeItems where Expired=0 and InvoiceID<>0 and SBLID=BLDetailID), GuiaRemision, Pos from BLDetail where BLID = " & ObjectID & " and Expired = 0 Order by Pos")
    'Set rs = Conn.Execute("select BLDetailID, BLIDTransit, NoOfPieces, ClassNoOfPieces, CommoditiesID, DiceContener, DiceContenerValue, Volumes, Weights, ClientsID, AddressesID, Clients, BLs, DischargeDate, CountriesFinalDes, InTransit, AgentsID, Agents, HBLNumber, AgentsAddrID, Seps, CountryOrigen, (select count(SBLID) from ChargeItems where Expired=0 and InvoiceID<>0 and SBLID=BLDetailID), GuiaRemision, Pos, CodeReference from BLDetail  where BLID = " & ObjectID & " and Expired = 0 Order by Pos")


	'                                   0               1             2               3                  4              5                    6             7              8         9           10         11       12         13                14                      15        16        17         18           19            20         21                                                                                                          22        23         24                   25                                                                   26                             27       28       29           30                    31                              32                33
    QuerySelect = "select DISTINCT a.BLDetailID, a.BLIDTransit, a.NoOfPieces, a.ClassNoOfPieces, a.CommoditiesID, a.DiceContener, a.DiceContenerValue, a.Volumes, a.Weights, a.ClientsID, a.AddressesID, a.Clients, a.BLs, a.DischargeDate, a.CountriesFinalDes, a.InTransit, a.AgentsID, a.Agents, a.HBLNumber, a.AgentsAddrID, a.Seps, a.CountryOrigen, (select count(SBLID) from ChargeItems where Expired=0 and InvoiceID<>0 and SBLID=a.BLDetailID) as c, a.GuiaRemision, a.Pos, IFNULL(a.CodeReference,0), REPLACE(REPLACE(IFNULL(TruckNo,''),'-',''),' ','') as TruckNo, IFNULL(d.nombre,'') as placa, d.BLID, d.HBLNumber, IFNULL(a.BLID,0),  a.HBLNumber, IFNULL(d.id_json,0) as id_json, TruckNo as plate " & _
    "FROM BLDetail a " & _
    "LEFT JOIN BLs b ON a.BLID = b.BLID " & _
    "LEFT JOIN Trucks c ON b.TruckID = c.TruckID " & _
    "LEFT JOIN disatel_json d ON (a.BLDetailID = d.BLDetailID AND geocerca = '' AND geocerca_mov = '') " & _
    "WHERE a.BLID = " & ObjectID & " and a.Expired = 0 and a.BLDetailID IS NOT NULL Order by Pos"

    'response.write QuerySelect & "<br>"
    
    Set rs = Conn.Execute(QuerySelect)
    If Not rs.EOF Then
   		aList5Values = rs.GetRows
       	CountList5Values = rs.RecordCount-1
    End If
    CloseOBJs rs, Conn
	
        'response.write QuerySelect & "<br>" & aList5Values(30,0) & "<br>"
        i = 0
        
        On Error Resume Next      
                    
            i = CheckNum(aList5Values(30,0))

        If Err.Number<>0 then
	        'response.write "disatelgps :" & Err.Number & " - " & Err.Description & "<br>"  
        end if
                   
        'response.write "(" & CountList5Values & ")(" & i & ")<br>"

	if CountList5Values >= 0 and i > 0 then
		ntr = chr(13) & chr(10)
		Redim Preserve BLDetailID(CountList5Values)
		Redim Preserve BLIDTransit(CountList5Values)
		Redim Preserve NoOfPieces(CountList5Values)
		Redim Preserve ClassNoOfPieces(CountList5Values)
		Redim Preserve CommoditiesID(CountList5Values)
		Redim Preserve DiceContener(CountList5Values)
		Redim Preserve DiceContenerValue(CountList5Values)
		Redim Preserve Volumes(CountList5Values)
		Redim Preserve Weights(CountList5Values)
		Redim Preserve ClientsID(CountList5Values)
		Redim Preserve AddressesID(CountList5Values)
		Redim Preserve Clients(CountList5Values)
		Redim Preserve BLs(CountList5Values)
		Redim Preserve DischargeDate(CountList5Values)
		Redim Preserve CountriesFinalDes(CountList5Values)
		Redim Preserve InTransit(CountList5Values)
		Redim Preserve AgentsID(CountList5Values)
		Redim Preserve AgentsAddrID(CountList5Values)
		Redim Preserve Agents(CountList5Values)
		Redim Preserve HBLNumber(CountList5Values)
        Redim Preserve HaveInvoices(CountList5Values)
		Redim Preserve Seps(CountList5Values)
		Redim Preserve CountriesOrigen(CountList5Values)
        Redim Preserve GuiaRemisionDet(CountList5Values)
        Redim Preserve Posi(CountList5Values)
		
		for i=0 to CountList5Values
			BLDetailID(i) = aList5Values(0,i)
			BLIDTransit(i) = aList5Values(1,i)
			NoOfPieces(i) = aList5Values(2,i)
			if aList5Values(3,i) = "" or isnull(aList5Values(3,i)) then
				ClassNoOfPieces(i) = ""
			else
				ClassNoOfPieces(i) = aList5Values(3,i)
			end if
			CommoditiesID(i) = aList5Values(4,i)
			DiceContener(i) = aList5Values(5,i)
			DiceContenerValue(i) = aList5Values(6,i)
			Volumes(i) = aList5Values(7,i)
			Weights(i) = aList5Values(8,i)

			'if Consolidated = 1 then
				'ClientsID(i) = aList5Values(9,i)
				'AddressesID(i) = aList5Values(10,i)
				'Clients(i) = aList5Values(11,i)
				'BLs(i) = aList5Values(12,i)
				'DischargeDate(i) = aList5Values(13,i)
				'CountriesFinalDes(i) = aList5Values(14,i)
				'AgentsID(i) = aList5Values(16,i)
				'Agents(i) = aList5Values(17,i)				
				'AgentsAddrID(i) = aList5Values(19,i)
				'CountriesOrigen(i) = aList5Values(21,i)
			'else
				'ClientsID(0) = aList5Values(9,i)
				'AddressesID(0) = aList5Values(10,i)
				'Clients(0) = aList5Values(11,i)
				'BLs(0) = aList5Values(12,i)
				'DischargeDate(0) = aList5Values(13,i)
				'CountriesFinalDes(0) = aList5Values(14,i)
				'AgentsID(0) = aList5Values(16,i)
				'Agents(0) = aList5Values(17,i)	
				'AgentsAddrID(0) = aList5Values(19,i)
				'CountriesOrigen(0) = aList5Values(21,i)
			'end if
			ClientsID(i) = aList5Values(9,i)
			AddressesID(i) = aList5Values(10,i)
			Clients(i) = aList5Values(11,i)
			BLs(i) = aList5Values(12,i)
			DischargeDate(i) = aList5Values(13,i)
			CountriesFinalDes(i) = aList5Values(14,i)
			InTransit(i) = aList5Values(15,i)
			AgentsID(i) = aList5Values(16,i)
			Agents(i) = aList5Values(17,i)				
			HBLNumber(i) = aList5Values(18,i)
			AgentsAddrID(i) = aList5Values(19,i)
			Seps(i) = aList5Values(20,i)
            CountriesOrigen(i) = aList5Values(21,i)
            HaveInvoices(i) = aList5Values(22,i)
            GuiaRemisionDet(i) = aList5Values(23,i)
            Posi(i) = aList5Values(24,i)



            If Action = 1 or Action = 2 Then

                if fleet <> "NA" AND fleet <> "" AND BLType = 1 then 'solo grupos tipo colgate y express

                    'response.write "(" & aList5Values(26,i) & ")(" & aList5Values(27,i) & ")<br>"

                    '          TruckNo                 placa
                    if aList5Values(26,i) <> aList5Values(27,i) then

                        On Error Resume Next      
                            'response.write "(" & aList5Values(0,i) & ")<br>"   
                            'response.write "(" & aList5Values(25,i) & ")<br>"  
                            'response.write "(" & aList5Values(26,i) & ")<br>"  
                            'response.write "(" & aList5Values(30,i) & ")<br>"  
                            'response.write "(" & aList5Values(31,i) & ")<br>"  
                            'response.write "(" & aList5Values(32,i) & ")<br>"  
                            'response.write "(" & aList5Values(9,i) & ")<br>"  
                            'response.write "(" & aList5Values(33,i) & ")<br>" 

                            '               BLDetailID          codreference        TruckNo                 BLID                HBLNumber       id_json          ConsignerID        plate
                            disatelgps aList5Values(0,i), aList5Values(25,i), aList5Values(26,i), aList5Values(30,i), aList5Values(31,i), aList5Values(32,i), aList5Values(9,i), aList5Values(33,i), fleet
                    
                        If Err.Number<>0 then
	                        response.write "disatelgps :" & Err.Number & " - " & Err.Description & "<br>"  
                        end if
                   
                    end if


                end if

            end if

		next


            If Action = 1 or Action = 2 Then

                if fleet <> "NA" AND fleet <> "" AND BLType = 1 then

	                'response.write "disatelgps :(" & BLID & ") - (" & fleet & ")<br>"  

                    'bloquea placa          
                    ActivePlate CheckNum(BLID), 0
 
                end if

            end if

	end if

    
                        On Error Resume Next      

    if (ObjectID > 0 and Join(BLDetailID, ",") <> "") then
        openConnBAW Conn
            'response.write("select concat(a.tfa_routing , '|', a.tfa_serie, '-', a.tfa_correlativo, '(FA)', '|' , case when b.pai_nombre = 'GUATEMALA' then 'GT' when b.pai_nombre = 'EL SALVADOR' then 'SV' when b.pai_nombre = 'HONDURAS' then 'HN' when b.pai_nombre = 'NICARAGUA' then 'NI' when b.pai_nombre = 'COSTA RICA' then 'CR' when b.pai_nombre = 'PANAMA' then 'PA' when b.pai_nombre = 'BELICE' then 'BZ' when b.pai_nombre = 'SALVADOR 2' then 'SVLOG' when b.pai_nombre = 'GRH' then 'GRH' when b.pai_nombre = 'ISI SURVEYOR' then 'ISI-NI' when b.pai_nombre = 'MAYAN LOGISTIC' then 'MAYAN-NI' when b.pai_nombre = 'APROA' then 'APROA' when b.pai_nombre = 'LATIN FREIGHT - GUATEMALA' then 'GTLTF' when b.pai_nombre = 'MAYAN LOGISTICS GT' then 'MAYAN-GT' when b.pai_nombre = 'REIMAR' then 'REIMAR' when b.pai_nombre = 'ISI SURVEYOR GT' then 'ISI-GT' when b.pai_nombre = 'MEXICO' then 'MX' when b.pai_nombre = 'AGENTE ADUANERO COSTA RICA' then 'AGAD-CR' when b.pai_nombre = 'LATIN FREIGHT - COSTA RICA' then 'CRLTF' when b.pai_nombre = 'LATIN FREIGHT - BELICE' then 'BZLTF' when b.pai_nombre = 'LATIN FREIGHT - HONDURAS' then 'HNLTF' when b.pai_nombre = 'LATIN FREIGHT - NICARAGUA' then 'NILTF' when b.pai_nombre = 'LATIN FREIGHT - PANAMA' then 'PALTF' when b.pai_nombre = 'LATIN FREIGHT - EL SALVADOR' then 'SVLTF' when b.pai_nombre = 'EQUITRANS GUATEMALA' then 'EQ-GT' when b.pai_nombre = 'EQUITRANS COSTA RICA' then 'EQ-CR' end) hbl_fact from tbl_facturacion a left join tbl_pais b on b.pai_id=a.tfa_pai_id left join tbl_sucursal c on c.suc_id=a.tfa_suc_id where a.tfa_tto_id in (5,6,7,15,16) and a.tfa_ted_id not in (3) and a.tfa_routing <> '' and a.tfa_ted_id not in (3) and a.tfa_routing in ('" & Join(BLs, "','") & "') UNION select concat(a.tnd_routing , '|', a.tnd_serie, '-', a.tnd_correlativo, '(ND)', '|' , case when b.pai_nombre = 'GUATEMALA' then 'GT' when b.pai_nombre = 'EL SALVADOR' then 'SV' when b.pai_nombre = 'HONDURAS' then 'HN' when b.pai_nombre = 'NICARAGUA' then 'NI' when b.pai_nombre = 'COSTA RICA' then 'CR' when b.pai_nombre = 'PANAMA' then 'PA' when b.pai_nombre = 'BELICE' then 'BZ' when b.pai_nombre = 'SALVADOR 2' then 'SVLOG' when b.pai_nombre = 'GRH' then 'GRH' when b.pai_nombre = 'ISI SURVEYOR' then 'ISI-NI' when b.pai_nombre = 'MAYAN LOGISTIC' then 'MAYAN-NI' when b.pai_nombre = 'APROA' then 'APROA' when b.pai_nombre = 'LATIN FREIGHT - GUATEMALA' then 'GTLTF' when b.pai_nombre = 'MAYAN LOGISTICS GT' then 'MAYAN-GT' when b.pai_nombre = 'REIMAR' then 'REIMAR' when b.pai_nombre = 'ISI SURVEYOR GT' then 'ISI-GT' when b.pai_nombre = 'MEXICO' then 'MX' when b.pai_nombre = 'AGENTE ADUANERO COSTA RICA' then 'AGAD-CR' when b.pai_nombre = 'LATIN FREIGHT - COSTA RICA' then 'CRLTF' when b.pai_nombre = 'LATIN FREIGHT - BELICE' then 'BZLTF' when b.pai_nombre = 'LATIN FREIGHT - HONDURAS' then 'HNLTF' when b.pai_nombre = 'LATIN FREIGHT - NICARAGUA' then 'NILTF' when b.pai_nombre = 'LATIN FREIGHT - PANAMA' then 'PALTF' when b.pai_nombre = 'LATIN FREIGHT - EL SALVADOR' then 'SVLTF' when b.pai_nombre = 'EQUITRANS GUATEMALA' then 'EQ-GT' when b.pai_nombre = 'EQUITRANS COSTA RICA' then 'EQ-CR' end) from tbl_nota_debito a left join tbl_pais b on b.pai_id=a.tnd_pai_id left join tbl_sucursal c on c.suc_id=a.tnd_suc_id where a.tnd_tto_id in (5,6,7,15,16) and a.tnd_ted_id not in (3) and a.tnd_routing <> '' and a.tnd_ted_id not in (3) and a.tnd_blid in (" & Join(BLDetailID, ",") & ") and a.tnd_routing in ('" & Join(BLs, "','") & "') order by hbl_fact ")
            Set rs = Conn.Execute("select concat(a.tfa_routing , '|', a.tfa_serie, '-', a.tfa_correlativo, '(FA)', '|' , case when b.pai_nombre = 'GUATEMALA' then 'GT' when b.pai_nombre = 'EL SALVADOR' then 'SV' when b.pai_nombre = 'HONDURAS' then 'HN' when b.pai_nombre = 'NICARAGUA' then 'NI' when b.pai_nombre = 'COSTA RICA' then 'CR' when b.pai_nombre = 'PANAMA' then 'PA' when b.pai_nombre = 'BELICE' then 'BZ' when b.pai_nombre = 'SALVADOR 2' then 'SVLOG' when b.pai_nombre = 'GRH' then 'GRH' when b.pai_nombre = 'ISI SURVEYOR' then 'ISI-NI' when b.pai_nombre = 'MAYAN LOGISTIC' then 'MAYAN-NI' when b.pai_nombre = 'APROA' then 'APROA' when b.pai_nombre = 'LATIN FREIGHT - GUATEMALA' then 'GTLTF' when b.pai_nombre = 'MAYAN LOGISTICS GT' then 'MAYAN-GT' when b.pai_nombre = 'REIMAR' then 'REIMAR' when b.pai_nombre = 'ISI SURVEYOR GT' then 'ISI-GT' when b.pai_nombre = 'MEXICO' then 'MX' when b.pai_nombre = 'AGENTE ADUANERO COSTA RICA' then 'AGAD-CR' when b.pai_nombre = 'LATIN FREIGHT - COSTA RICA' then 'CRLTF' when b.pai_nombre = 'LATIN FREIGHT - BELICE' then 'BZLTF' when b.pai_nombre = 'LATIN FREIGHT - HONDURAS' then 'HNLTF' when b.pai_nombre = 'LATIN FREIGHT - NICARAGUA' then 'NILTF' when b.pai_nombre = 'LATIN FREIGHT - PANAMA' then 'PALTF' when b.pai_nombre = 'LATIN FREIGHT - EL SALVADOR' then 'SVLTF' when b.pai_nombre = 'EQUITRANS GUATEMALA' then 'EQ-GT' when b.pai_nombre = 'EQUITRANS COSTA RICA' then 'EQ-CR' end) hbl_fact from tbl_facturacion a left join tbl_pais b on b.pai_id=a.tfa_pai_id left join tbl_sucursal c on c.suc_id=a.tfa_suc_id where a.tfa_tto_id in (5,6,7,15,16) and a.tfa_ted_id not in (3) and a.tfa_routing <> '' and a.tfa_ted_id not in (3) and a.tfa_routing in ('" & Join(BLs, "','") & "') UNION select concat(a.tnd_routing , '|', a.tnd_serie, '-', a.tnd_correlativo, '(ND)', '|' , case when b.pai_nombre = 'GUATEMALA' then 'GT' when b.pai_nombre = 'EL SALVADOR' then 'SV' when b.pai_nombre = 'HONDURAS' then 'HN' when b.pai_nombre = 'NICARAGUA' then 'NI' when b.pai_nombre = 'COSTA RICA' then 'CR' when b.pai_nombre = 'PANAMA' then 'PA' when b.pai_nombre = 'BELICE' then 'BZ' when b.pai_nombre = 'SALVADOR 2' then 'SVLOG' when b.pai_nombre = 'GRH' then 'GRH' when b.pai_nombre = 'ISI SURVEYOR' then 'ISI-NI' when b.pai_nombre = 'MAYAN LOGISTIC' then 'MAYAN-NI' when b.pai_nombre = 'APROA' then 'APROA' when b.pai_nombre = 'LATIN FREIGHT - GUATEMALA' then 'GTLTF' when b.pai_nombre = 'MAYAN LOGISTICS GT' then 'MAYAN-GT' when b.pai_nombre = 'REIMAR' then 'REIMAR' when b.pai_nombre = 'ISI SURVEYOR GT' then 'ISI-GT' when b.pai_nombre = 'MEXICO' then 'MX' when b.pai_nombre = 'AGENTE ADUANERO COSTA RICA' then 'AGAD-CR' when b.pai_nombre = 'LATIN FREIGHT - COSTA RICA' then 'CRLTF' when b.pai_nombre = 'LATIN FREIGHT - BELICE' then 'BZLTF' when b.pai_nombre = 'LATIN FREIGHT - HONDURAS' then 'HNLTF' when b.pai_nombre = 'LATIN FREIGHT - NICARAGUA' then 'NILTF' when b.pai_nombre = 'LATIN FREIGHT - PANAMA' then 'PALTF' when b.pai_nombre = 'LATIN FREIGHT - EL SALVADOR' then 'SVLTF' when b.pai_nombre = 'EQUITRANS GUATEMALA' then 'EQ-GT' when b.pai_nombre = 'EQUITRANS COSTA RICA' then 'EQ-CR' end) from tbl_nota_debito a left join tbl_pais b on b.pai_id=a.tnd_pai_id left join tbl_sucursal c on c.suc_id=a.tnd_suc_id where a.tnd_tto_id in (5,6,7,15,16) and a.tnd_ted_id not in (3) and a.tnd_routing <> '' and a.tnd_ted_id not in (3) and a.tnd_blid in (" & Join(BLDetailID, ",") & ") and a.tnd_routing in ('" & Join(BLs, "','") & "') order by hbl_fact ")
	        If Not rs.EOF Then
   		        aList6Values = rs.GetRows
       	        CountList6Values = rs.RecordCount-1
            End If
        CloseOBJs rs, Conn
    End if

                        If Err.Number<>0 then
	                        response.write "BLDetailID :" & Err.Number & " - " & Err.Description & "<br>"  
                        end if
                   
    if CountList6Values >= 0 then
        
        Redim Preserve DetailInvoices(CountList6Values)
        
        for i=0 to CountList6Values
            For j=0 to CountList5Values
            If InStr(Right(aList6Values(0,i),Len(aList6Values(0,i))-instrrev(aList6Values(0,i),"|")), Countries) or (BLIDTransit(j) = 0 and Split(aList6Values(0,i), "|")(0) = BLs(j)) or (InStr(CountriesFinalDes(j), CountryDes) and InStr(Right(aList6Values(0,i),Len(aList6Values(0,i))-instrrev(aList6Values(0,i),"|")), CountryDes)) Then
                DetailInvoices(i) = aList6Values(0,i)
            End If
            Next
        next

    end if
    
    'Obteniendo numero de semana y año
	OpenConn2 Conn
    Set rs = Conn.Execute("select	a.id, d.semana, case when d.semana = 1 then date_part('y',d.fecha_fin) else date_part('y',d.fecha_ini) end, b.semana, case when b.semana = 1 then date_part('y',b.fecha_fin) else date_part('y',b.fecha_ini) end, a.semana, case when a.semana = 1 then date_part('y',a.fecha_fin) else date_part('y',a.fecha_ini) end, c.semana, case when c.semana = 1 then date_part('y',c.fecha_fin) else date_part('y',c.fecha_ini) end, e.semana, case when e.semana = 1 then date_part('y',e.fecha_fin) else date_part('y',e.fecha_ini) end from numero_semana a inner join numero_semana b on a.id-1 = b.id inner join numero_semana c on a.id+1 = c.id inner join numero_semana d on a.id-2 = d.id inner join numero_semana e on a.id+2 = e.id where current_date between a.fecha_ini and a.fecha_fin")
	If Not rs.EOF Then
   		WeekPrev2 = rs(1)
        WeekPrev = rs(3)
        YearPrev2 = rs(2)
        YearPrev = rs(4)
        WeekAct = rs(5)
        YearAct = rs(6)
        WeekPost = rs(7)
        WeekPost2 = rs(9)
        YearPost = rs(8)
        YearPost2 = rs(10)
    End If
	CloseOBJs rs, Conn
%>
<html>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<LINK REL="stylesheet" type="text/css" HREF="img/estilos.css">
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
#myBar {
    width: 10%;
    height: 15px;
    background-color: #4CAF50;
    text-align: center; /* To center it horizontally (if you want) */
    line-height: 15px; /* To center it vertically */
    color: white;
    font-weight: bold;
    display: none;
}
-->
</style>
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
<%if JavaMsg <> "" then %>
    alert("<%=JavaMsg%>");
<%end if %>
var chkDiceContener=0;
var chkClients=0;
var ntr = "";
var com = "";
var WareHouseID = <%=CheckNum(ChargeType)%>;
var Marchamo="<%=Marchamo%>";

<%if Consolidated = 0 and CountList5Values >= 0 then%>
	var Consignee="<%=Clients(0)%>";
	var CountryShipper="<%=CountriesOrigen(0)%>";
	var CountryConsignee="<%=CountriesFinalDes(0)%>";
	var HBLN="<%=HBLNumber(0)%>";
<%else%>
	var Consignee="";
	var CountryConsignee="";
	var CountryShipper="";
	var HBLN="--";
<%end if%>

	function SetLabelID(Label) {
		var LabelID = "";
		if (Label == "BLExitDate") {
			LabelID = "Fecha Estimada de Salida";
		} else if (Label == "DeliveryPolicyDate") {
			LabelID = "Fecha en que se libero la Poliza";
		} else if (Label == "DischargeDate") {
			LabelID = "Fecha de Descarga";
		} else if (Label == "BLEstArrivalDate") {
			LabelID = "Fecha Estimada de Llegada";
		} else if (Label == "ChargePlace") {
			LabelID = "Lugar de Carga";
		} else if (Label == "FinalDes") {
			LabelID = "Destino Final / Entrega";
		} else if (Label == "BLDispatchDate") {
			LabelID = "Fecha de Despacho";
		}		
		return LabelID;
	}

	function abrir(Label){
	var DateSend, Subject;
		if (parseInt(navigator.appVersion) < 5) {
			DateSend = document.forma(Label).value;
		} else {
			var LabelID = SetLabelID(Label);
			DateSend = document.getElementById(LabelID).value;
		}
		Subject = '';
		window.open('Agenda.asp?Action=1&Label=' + Label + '&DateSend=' + DateSend + '&Subj=' + Subject,'Seleccionar','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,width=170,height=160,top=250,left=250');
	}
	
	function GetData(GID){
		if (GID==14) {
			if (!valSelec(document.forma.CountryDep)){return (false)};
		}
		//if ((document.forma.Consolidated.value!=1) && (GID==14) && (document.forma.ConsignerID.value=="")){
		//	alert (" Primero debe indicar los datos del Consignatario "); 
		//	document.forma.ConsignerData.focus();
		//	return (false);
			//break;
		//};
		var CSD = 0;
		if (document.forma.Consolidated.value==1) {
			CSD = 1;
		}		
		if (GID!=14) {
			window.open('Search_BLData.asp?GID='+GID+'&BTP='+document.forma.BLType.value,'BLData','height=200,width=450,menubar=0,resizable=1,scrollbars=1,toolbar=0,status=1');
		} else {
			window.open('Search_ResultsBLData.asp?GID='+GID+'&CTD='+document.forma.CountryDep.value+'&CID='+document.forma.ConsignerID.value+'&CSD='+CSD+'&BTP='+document.forma.BLType.value,'BLData','height=300,width=550,menubar=0,resizable=1,scrollbars=1,toolbar=0,status=1');
		}
	}
	function IData(GID){
		window.open('InsertData.asp?GID='+GID+'&SO=1','IData','height=330,width=460,menubar=0,resizable=1,scrollbars=1,toolbar=0');
	}
	function GetDoc(ObjID){
		window.open('Docs.asp?OID='+ObjID,'BLDocs','height=200,width=500,menubar=0,resizable=1,scrollbars=1,toolbar=0,status=1');
	}
	function LTrim(s){
		// Devuelve una cadena sin los espacios del principio
		var i=0;
		var j=0;
		// Busca el primer caracter <> de un espacio
		for(i=0; i<=s.length-1; i++)
			if(s.substring(i,i+1) != ' '){
				j=i;
				break;
			}
		return s.substring(j, s.length);
	}
	function RTrim(s){
		// Quita los espacios en blanco del final de la cadena
		var j=0;
		// Busca el último caracter <> de un espacio
		for(var i=s.length-1; i>-1; i--)
			if(s.substring(i,i+1) != ' '){
				j=i;
				break;
			}
		return s.substring(0, j+1);
	}
	function Trim(s){
		// Quita los espacios del principio y del final
		return LTrim(RTrim(s));
	}
	function SetFlag(Obj){
		if (Obj.value.length == 0) {
			return 0;
		} else {
			return 1;
		}		
	}

    function ValidarPlaca(obj) {

        var placa, id;
      
        id = obj.value;

        placa = obj.options[obj.selectedIndex].text
 
         <% if fleet <> "NA" and fleet <> "" then %>

            //move();
		    //document.forma.Action.value = 9; esta accion esta deshibilitada por el momento
		    //document.forma.submit();	
/*2020-12-04 se deshabilita temporalmente
            if (placa.indexOf('(GPS)') !== -1) {   //validar que placa sea colgate
            
                if (placa.indexOf('(OFF)') !== -1) {   //validar que placa esta disponible       
                    obj.value = '';
                    alert("Placa " + placa.replace("(OFF)", "").replace("(GPS)", "") + " ya se encuentra en ruta ");
                }

            } else {
                obj.value = '';
                alert("Placa " + placa.replace("(OFF)", "").replace("(GPS)", "") + " No esta registrada con GPS");
            }
*/                          
        <% end if %>
                   


    }
	

	function validar(Action) {
  		if (Action != 3) {
			SetClassNoOfPieces(document.forma.NoOfPieces);
			if ((document.forma.BLType.value==0)||(document.forma.BLType.value==1)) {
				if (!valTxt(document.forma.Week, 1, 3)){return (false)};
				if (!valTxt(document.forma.SenderData, 3, 6)){return (false)};
				if (!valTxt(document.forma.ShipperData, 3, 6)){return (false)};
				if (!valTxt(document.forma.ConsignerData, 3, 6)){return (false)};

                //Validacion de Latin Freight y Aimar, el resto de empresas no tiene esta validacion, por ejemplo N1 (GRH)
                <%if FilterAimarLatin = 1 then%>
                if ((document.forma.Countries.value!="N1") && (document.forma.Countries.value!="A2")) {
                    if (document.forma.Countries.value.substr(2,3)=="LTF") {
                        if (document.forma.AgentNeutral.value == 0) {
                            alert("Para operaciones de Latin Freight solo puede utilizar agentes Neutrales");
				            document.forma.SenderData.focus();
                            return (false);
			            }
                    } else {
                        var EconoCodes = /<%=PtrnEconoCodes%>/;
                        var Result = EconoCodes.exec(document.forma.AgentsID.value)
                        if (Result == null) {
                            if (document.forma.ConsignerColoader.value == 1) {
                                alert("Para operaciones de Aimar, solo cuando el Agente es Econocaribe puede asignar Clientes o Shippers Coloaders, favor de consultar con su supervisor y revisar el administrador de catalogos");
				                document.forma.ConsignerData.focus();
                                return (false);
			                }
                            if (document.forma.ShipperColoader.value == 1) {
                                alert("Para operaciones de Aimar, solo cuando el Agente es Econocaribe puede asignar Clientes o Shippers Coloaders, favor de consultar con su supervisor y revisar el administrador de catalogos");
				                document.forma.ShipperData.focus();
                                return (false);
			                }
			            }
                    }
                }
                <%end if %>

				if (!valSelec(document.forma.CountryDep)){return (false)};
				if (!valSelec(document.forma.BrokerID)){return (false)};
				if (!valSelec(document.forma.CountryDes)){return (false)};
				if (document.forma.CountryDep.value == document.forma.CountryDes.value) {
						alert ("EL pais de Origen (Transito) debe ser diferente al Pais de Destino (Transito)"); 
						document.forma.CountryDep.focus();
						return (false);	
				}
				if (!valTxt(document.forma.ChargePlace, 3, 5)){return (false)};
				if (!valTxt(document.forma.FinalDes, 3, 5)){return (false)};
				if (!valSelec(document.forma.PilotID, 3, 5)){return (false)};
				if (!valSelec(document.forma.TruckID, 3, 5)){return (false)};
				if (!valTxt(document.forma.NoOfPieces, 1, 5)){return (false)};
				if (!valTxt(document.forma.Contener, 3, 5)){return (false)};
				if (!valTxt(document.forma.Volumes, 1, 5)){return (false)};
				if (!valTxt(document.forma.Weights, 1, 5)){return (false)};
				if (document.forma.Consolidated.value==1){
					if (!valTxt(document.forma.Clients, 3, 5)){return (false)};
				}
				if (!valTxt(document.forma.BLs, 3, 8)){return (false)};
				if (!valTxt(document.forma.DischargeDate, 10, 5)){return (false)};
				if (!valTxt(document.forma.BLDispatchDate)){return (false)};
				//if (!valTxt(document.forma.BLExitDate, 3, 5)){return (false)};
				//if (!valTxt(document.forma.BLEstArrivalDate, 3, 5)){return (false)};
				if (!valTxt(document.forma.Agents, 3, 5)){return (false)};
				<%if Countries<>"GT" and Countries<>"SV" then%>
				if (!valTxt(document.forma.DiceContenerValue, 1, 5)){return (false)};
				<%end if%>
				if (!valTxt(document.forma.ContactSignature, 3, 5)){return (false)};
				if (!valTxt(document.forma.TotPrepaid, 3, 5)){return (false)};
				if (!valTxt(document.forma.TotCollect, 3, 5)){return (false)};
				if (!valSelec(document.forma.Countries)){return (false)};                
			} else {
				if (!valTxt(document.forma.Week, 1, 3)){return (false)};
				if (!valTxt(document.forma.ConsignerData, 3, 6)){return (false)};
				if (!valTxt(document.forma.Attn, 3, 5)){return (false)};
				if (!valTxt(document.forma.FinalDes, 3, 5)){return (false)};
				if (!valSelec(document.forma.PilotID, 3, 5)){return (false)};
				if (!valSelec(document.forma.TruckID, 3, 5)){return (false)};
				if (!valTxt(document.forma.NoOfPieces, 1, 5)){return (false)};
				if (!valTxt(document.forma.Contener, 3, 5)){return (false)};
				if (!valTxt(document.forma.Volumes, 1, 5)){return (false)};
				if (!valTxt(document.forma.Weights, 1, 5)){return (false)};
				<%if ItineraryType=2 then %>
                if (!valTxt(document.forma.BLExitDate, 3, 5)){return (false)};
				if (!valTxt(document.forma.BLEstArrivalDate, 3, 5)){return (false)};
				<%end if %>
                if (!valTxt(document.forma.ContactSignature, 3, 5)){return (false)};
				if (!valTxt(document.forma.TotPrepaid, 3, 5)){return (false)};
				if (!valTxt(document.forma.TotCollect, 3, 5)){return (false)};
				if (!valSelec(document.forma.Countries)){return (false)};
				if (document.forma.Closed.value==1) {
					if (!valTxt(document.forma.PolicyNo, 3, 5)){return (false)};
				}
				if (document.forma.PolicyNo.value!="") {
					if (!valTxt(document.forma.DeliveryPolicyDate, 3, 5)){return (false)};
					if (!valSelec(document.forma.DeliveryPolicyHour)){return (false)};
					if (!valSelec(document.forma.DeliveryPolicyMin)){return (false)};
				}
			}
			var i = 0;
			var j = 0;
			var k = 0;
			var Lengs = new Array();
			var Temp = new Array();
			/*
			Lengs[0] = document.forma.NoOfPieces.value.split("\r\n"); //13
			Lengs[1] = document.forma.ClassNoOfPieces.value.split("\r\n"); //14
			Lengs[2] = document.forma.DiceContener.value.split("|"); //15
			Lengs[3] = document.forma.DiceContenerValue.value.split("\r\n"); //16
			Lengs[4] = document.forma.Volumes.value.split("\r\n"); //17
			Lengs[5] = document.forma.Weights.value.split("\r\n"); //18
			Lengs[6] = document.forma.Clients.value.split("\r\n"); //19
			Lengs[7] = document.forma.BLs.value.split("\r\n"); //20
			Lengs[8] = document.forma.DischargeDate.value.split("\r\n"); //21
			Lengs[9] = document.forma.CountriesOrigen.value.split("\r\n"); //22
			Lengs[10] = document.forma.CountriesFinalDes.value.split("\r\n"); //23
			Lengs[11] = document.forma.Agents.value.split("\r\n"); //24
			Lengs[12] = document.forma.Seps.value.split("\r\n"); //25
            */

            Lengs[0] = FixCart(document.forma.NoOfPieces); //13
			Lengs[1] = FixCart(document.forma.ClassNoOfPieces); //14
			Lengs[2] = document.forma.DiceContener.value.split("|"); //15
			Lengs[3] = FixCart(document.forma.DiceContenerValue); //16
			Lengs[4] = FixCart(document.forma.Volumes); //17
			Lengs[5] = FixCart(document.forma.Weights); //18
			Lengs[6] = FixCart(document.forma.Clients); //19
			Lengs[7] = FixCart(document.forma.BLs); //20
			Lengs[8] = FixCart(document.forma.DischargeDate); //21
			Lengs[9] = FixCart(document.forma.CountriesOrigen); //22
			Lengs[10] = FixCart(document.forma.CountriesFinalDes); //23
			Lengs[11] = FixCart(document.forma.Agents); //24
			Lengs[12] = FixCart(document.forma.Seps); //25


			//Eliminando espacios innecesarios al final de cada columna
			for(i=1; i<13; i++) {
				while ((Lengs[0].length!=Lengs[i].length) && (Lengs[i].length!=0) && (Trim(Lengs[i][Lengs[i].length-1])==" ")) {
					Lengs[i].pop();
				};
			}			
			//Obteniendo las longitudes de cada textarea para compararlas entre si
			for(i=0; i<13; i++) {
				Lengs[i+13] = Lengs[i].length-1;
				//alert((i+11)+"-"+Lengs[i+11]);
			}
			Lengs[26] = "NoOfPieces";
			Lengs[27] = "ClassNoOfPieces";
			Lengs[28] = "Contener";
			Lengs[29] = "DiceContenerValue";
			Lengs[30] = "Volumes"; 
			Lengs[31] = "Weights";
			Lengs[32] = "Clients";
			Lengs[33] = "BLs";
			Lengs[34] = "DischargeDate";
			Lengs[35] = "CountriesOrigen";
			Lengs[36] = "CountriesFinalDes";
			Lengs[37] = "Agents";
			Lengs[38] = "Seps";
			Lengs[39] = document.forma.BLDetailID.value.split("|"); 
			Lengs[40] = document.forma.BLIDTransit.value.split("|"); 
			Lengs[41] = document.forma.CommoditiesID.value.split("|"); 
			Lengs[42] = document.forma.ClientsID.value.split("|"); 
			Lengs[43] = document.forma.AddressesID.value.split("|"); 
			Lengs[44] = document.forma.InTransit.value.split("|"); 
			Lengs[45] = document.forma.AgentsID.value.split("|"); 
			Lengs[46] = document.forma.AgentsAddrID.value.split("|"); 
			Lengs[47] = document.forma.HBLNumber.value.split("|"); 

			//Checking Tabulacion de Datos para Consolidado y Express
			for(i=14; i<19; i++) {
				if (Lengs[13] != Lengs[i]) { //Todos se comparan con la cantidad de Bultos
					//Mensajes en el array Lengs de la posicion ClassNoOfPieces-Weights
					
                    var temp; //2020-08-20
                    try {
                        temp = document.forma.elements(Lengs[i+13]);
                        alert ("La cantidad de líneas ingresadas en la casilla '" + temp.id + "' no coincide con las líneas ingresadas en 'Cantidad de Bultos'"); 
                        temp.focus();
                    }
                    catch(err) {
                        alert ("La cantidad de líneas ingresadas no coincide " + (i + 13)); 								
                    }

					break;
				}
			}
			//Checking Tabulacion de mas Datos solo para Consolidado
			if (document.forma.BLType.value==0) {
				for(i=19; i<=25; i++) {
					if (Lengs[13] != Lengs[i]) { //Todos se comparan con la cantidad de Bultos
						//Mensajes en el array Lengs de la posicion Clients-Seps
						alert ("La cantidad de líneas ingresadas en la casilla '" + document.forma.elements(Lengs[i+13]).id + "' no coincide con las líneas ingresadas en 'Cantidad de Bultos'"); 
						document.forma.elements(Lengs[i+13]).focus();
						return (false);
						break;					
					};
				};	
			} else {
				if (document.forma.BLType.value==1) {//Express
					//for(i=19; i<=19; i++) {
						//if (Lengs[i]>0) {
					for(n=0; n<Lengs[42].length; n++) {//Comparando cada ClientID del detalle, contra el ConsignerID del encabezado, en Express no pueden ser distintos clientes
						if (Lengs[42][n]!=document.forma.ConsignerID.value) { //Lengs[42] es el arreglo de ClientsID
							
                            
                            //alert ("Como el documento es Express no puede agregar distintos datos en '" + document.forma.elements(Lengs[i+13]).id + "'"); 
							//document.forma.elements(Lengs[i+13]).focus();

                            var temp; //2020-08-20
                            try {
                                temp = document.forma.elements(Lengs[i+13]);
                                alert ("Como el documento es Express no puede agregar distintos datos en '" + temp.id + "'"); 							
                                temp.focus();
                            }
                            catch(err) {
                                alert ("Como el documento es Express no puede agregar distintos datos"); 								
                            }

							return (false);
							break;
						};
					};
				};
			};

			Lengs[48] = new Array(); //BLDetailID
			Lengs[49] = new Array(); //Position
			Lengs[50] = new Array(); //BLIDTransit
			Lengs[51] = new Array(); //NoOfPieces
			Lengs[52] = new Array(); //ClassNoOfPieces
			Lengs[53] = new Array(); //CommoditiesID
			Lengs[54] = new Array(); //DiceContener
			Lengs[55] = new Array(); //DiceContenerValue
			Lengs[56] = new Array(); //Volumes
			Lengs[57] = new Array(); //Weights
			Lengs[58] = new Array(); //ClientsID
			Lengs[59] = new Array(); //AddressesID
			Lengs[60] = new Array(); //Clients
			Lengs[61] = new Array(); //BLs
			Lengs[62] = new Array(); //DischargeDate
			Lengs[63] = new Array(); //CountriesOrigen
			Lengs[64] = new Array(); //CountriesFinalDes
			Lengs[65] = new Array(); //InTransit
			Lengs[66] = new Array(); //AgentsID
			Lengs[67] = new Array(); //AgentsAddrID
			Lengs[68] = new Array(); //Agents
			Lengs[69] = new Array(); //HBLNumber
			Lengs[70] = new Array(); //Seps
            
			//Ordenando Geograficamente los datos consolidados
			if (document.forma.BLType.value==0) {
				var Countries = new Array(); //Orden de Paises Geograficamente
				Countries[0] = "MX";
				Countries[1] = "BZ";
				Countries[2] = "GT";			
				Countries[3] = "SV";			
				Countries[4] = "HN";			
				Countries[5] = "NI";			
				Countries[6] = "CR";
				Countries[7] = "PA";
				Countries[8] = "PR";
				Countries[9] = "US";
				Countries[10] = "DO";
				Countries[11] = "PL"; //Polonia
				Countries[12] = "FR"; //Francia
				Countries[13] = "AF"; //AFGANISTAN
				Countries[14] = "AL"; //ALBANIA
				Countries[15] = "DE"; //ALEMANIA
				Countries[16] = "AD"; //ANDORRA
				Countries[17] = "AO"; //ANGOLA
				Countries[18] = "AQ"; //ANTARTIDA
				Countries[19] = "AG"; //ANTIGUA Y BARBUDA
				Countries[20] = "AN"; //ANTILLAS DE LOS PAISES BAJOS
				Countries[21] = "SA"; //ARABIA SAUDITA
				Countries[22] = "DZ"; //ARGELIA
				Countries[23] = "AR"; //ARGENTINA
				Countries[24] = "AM"; //ARMENIA
				Countries[25] = "AW"; //ARUBA
				Countries[26] = "AU"; //AUSTRALIA
				Countries[27] = "AT"; //AUSTRIA
				Countries[28] = "AZ"; //AZERBAIJN
				Countries[29] = "BS"; //BAHAMAS
				Countries[30] = "BH"; //BAHRAYN
				Countries[31] = "BD"; //BANGLADESH
				Countries[32] = "BB"; //BARBADOS -ISLAS DE BARLOVENTO-
				Countries[33] = "BY"; //BELARUS
				Countries[34] = "BE"; //BELGICA
				Countries[35] = "BJ"; //BENN
				Countries[36] = "BM"; //BERMUDAS
				Countries[37] = "MM"; //BIRMANIA
				Countries[38] = "BO"; //BOLIVIA
				Countries[39] = "BA"; //BOSNIA Y HERZEGOVINA
				Countries[40] = "BW"; //BOTSWANA ESTADO DE  AFRICA AUSTRAL
				Countries[41] = "BV"; //BOUVET ISLA NORUETA DEL ATLANTICO SUR
				Countries[42] = "BR"; //BRASIL
				Countries[43] = "BN"; //BRUNEI DARUSSALAM
				Countries[44] = "BG"; //BULGARIA
				Countries[45] = "BF"; //BURKINA FASO
				Countries[46] = "BI"; //BURUNDI
				Countries[47] = "BT"; //BUTN
				Countries[48] = "KH"; //CAMBOYA
				Countries[49] = "CM"; //CAMERN
				Countries[50] = "CA"; //CANADA
				Countries[51] = "CV"; //CAPE VERDE
				Countries[52] = "TD"; //CHAD
				Countries[53] = "CL"; //CHILE
				Countries[54] = "CN"; //CHINA
				Countries[55] = "CY"; //CHIPRE
				Countries[56] = "CO"; //COLOMBIA
				Countries[57] = "KM"; //COMORES
				Countries[58] = "CG"; //CONGO
				Countries[59] = "HR"; //CROACIA
				Countries[60] = "CU"; //CUBA
				Countries[61] = "DK"; //DINAMARCA
				Countries[62] = "DJ"; //DJIBOUTI
				Countries[63] = "DM"; //DOMINICA
				Countries[64] = "EC"; //ECUADOR
				Countries[65] = "EG"; //EGYPTO
				Countries[66] = "VA"; //EL VATICANO
				Countries[67] = "AE"; //EMIRATOS RABES UNIDOS
				Countries[68] = "ER"; //ERITREA
				Countries[69] = "SK"; //ESLOVAQUIA
				Countries[70] = "SI"; //ESLOVENIA
				Countries[71] = "ES"; //ESPAÑA
				Countries[72] = "LV"; //ESTADO RUSO DE LATVIA
				Countries[73] = "EE"; //ESTONIA
				Countries[74] = "ET"; //ETIOPA
				Countries[75] = "RU"; //FEDERACIN RUSIA
				Countries[76] = "FJ"; //FIDJI
				Countries[77] = "PH"; //FILIPINAS
				Countries[78] = "FI"; //FINLANDIA
				Countries[79] = "GA"; //GABN
				Countries[80] = "GM"; //GAMBIA
				Countries[81] = "GE"; //GEORGIA
				Countries[82] = "GS"; //GEORGIA DEL SUR Y LAS ISLAS DEL SUR DE SANDWICH
				Countries[83] = "GH"; //GHNA
				Countries[84] = "GI"; //GIBRALTAR
				Countries[85] = "GD"; //GRANADA
				Countries[86] = "GR"; //GRECIA
				Countries[87] = "GL"; //GROENLANDIA
				Countries[88] = "GP"; //GUADALUPE
				Countries[89] = "GU"; //GUAM
				Countries[90] = "GN"; //GUINEA
				Countries[91] = "GQ"; //GUINEA ECUATORIAL
				Countries[92] = "GF"; //GUINEA FRANCESA
				Countries[93] = "GW"; //GUINEA PORTUGUESA
				Countries[94] = "GY"; //GUYANA
				Countries[95] = "HT"; //HAIT
				Countries[96] = "HK"; //HONG KONG
				Countries[97] = "HU"; //HUNGRIA
				Countries[98] = "IN"; //INDIA
				Countries[99] = "IO"; //INDIAS BRITANICAS TERRITORIO DEL OCENO INDICO
				Countries[100] = "ID"; //INDONESIA
				Countries[101] = "IR"; //IRAN
				Countries[102] = "IQ"; //IRAQ
				Countries[103] = "IE"; //IRLANDIA
				Countries[104] = "MQ"; //ISLA DE MARTINICA
				Countries[105] = "CX"; //ISLA DE NAVIDAD
				Countries[106] = "MU"; //ISLA MAURICIO
				Countries[107] = "NF"; //ISLA NORFOLK
				Countries[108] = "PN"; //ISLA PITCAIRN
				Countries[109] = "IS"; //ISLANDIA
				Countries[110] = "VG"; //ISLAS BRITNICAS VIRGINIA
				Countries[111] = "KY"; //ISLAS CAYMAN
				Countries[112] = "CK"; //ISLAS DE COOK
				Countries[113] = "VI"; //ISLAS DE ESTADOS UNIDOS VIRGINIA
				Countries[114] = "FK"; //ISLAS DE FALKAND MALVINAS
				Countries[115] = "FO"; //ISLAS DE FAROE
				Countries[116] = "CC"; //ISLAS DE LOS COCOS
				Countries[117] = "MP"; //ISLAS MARIANAS DEL NORTE
				Countries[118] = "MH"; //ISLAS MARSHALL
				Countries[119] = "UM"; //ISLAS MENORES Y PERFERICAS DE ESTADOS UNIDOS
				Countries[120] = "SB"; //ISLAS SALOMON
				Countries[121] = "TC"; //ISLAS TURCAS Y TURCOS
				Countries[122] = "IL"; //ISRAEL
				Countries[123] = "IT"; //ITALIA
				Countries[124] = "JM"; //JAMAICA
				Countries[125] = "JP"; //JAPóN
				Countries[126] = "JO"; //JORDANIA
				Countries[127] = "KZ"; //KASAJISTN
				Countries[128] = "KE"; //KENYA
				Countries[129] = "KG"; //KIRGUIZISTN
				Countries[130] = "KI"; //KIRIBATI
				Countries[131] = "KW"; //KUWAIT
				Countries[132] = "LA"; //LAOS
				Countries[133] = "LS"; //LESOTHO O BASUTOLANDIA
				Countries[134] = "LB"; //LBANO
				Countries[135] = "LR"; //LIBERIA
				Countries[136] = "LY"; //LIBIA ARABE JAMAHIRYA
				Countries[137] = "LI"; //LIECHTENSTEIN
				Countries[138] = "LT"; //LITUNIA
				Countries[139] = "LU"; //LUXEMBURGO
				Countries[140] = "MO"; //MACAO
				Countries[141] = "MK"; //MACEDONIA, ANTIGUA REPBLICA DE YUGOESLAVIA
				Countries[142] = "MG"; //MADAGASCAR
				Countries[143] = "MY"; //MALASYA
				Countries[144] = "MW"; //MALAWI
				Countries[145] = "MV"; //MALDIVAS
				Countries[146] = "ML"; //MAL
				Countries[147] = "MT"; //MALTA
				Countries[148] = "MA"; //MARRUECOS
				Countries[149] = "MR"; //MAURITANIA
				Countries[150] = "YT"; //MAYOTT
				Countries[151] = "MC"; //MNACO
				Countries[152] = "MN"; //MONGOLIA
				Countries[153] = "MS"; //MONTSERRAT
				Countries[154] = "MZ"; //MOZAMBIQUE
				Countries[155] = "NA"; //NAMIBIA
				Countries[156] = "NR"; //NAURU
				Countries[157] = "NP"; //NEPAL
				Countries[158] = "NE"; //NGER
				Countries[159] = "NG"; //NIGERIA
				Countries[160] = "NO"; //NORUEGA
				Countries[161] = "NC"; //NUEVA CALEDONIA
				Countries[162] = "PG"; //NUEVA GUINEA - PAPUASIA
				Countries[163] = "NZ"; //NUEVA ZELANDA
				Countries[164] = "OM"; //OMN
				Countries[165] = "NL"; //PAISES BAJOS
				Countries[166] = "PK"; //PAKISTN
				Countries[167] = "PW"; //PALAOS
				Countries[168] = "PS"; //PALESTINA
				Countries[169] = "PY"; //PARAGUAY
				Countries[170] = "PE"; //PERU
				Countries[171] = "PF"; //POLINESIA FRANCESA
				Countries[172] = "PT"; //PORTUGAL
				Countries[173] = "QA"; //QATAR
				Countries[174] = "GB"; //REINO UNIDO O INGLATERRA
				Countries[175] = "SY"; //REPUBLICA ARABE DE SIRIA
				Countries[176] = "CZ"; //REPUBLICA CHECA
				Countries[177] = "CF"; //REPUBLICA DE AFRICA CENTRAL
				Countries[178] = "KR"; //KOREA
				Countries[179] = "MD"; //REPUBLICA DE MOLDOVIA
				Countries[180] = "CD"; //REPUBLICA DEMOCRTICA DEL CONGO
				Countries[181] = "TZ"; //REPUBLICA UNIDA DE TANZANIA
				Countries[182] = "KP"; //REPUBLICAS DEMOCRTICAS DE COREA
				Countries[183] = "RE"; //RUNION
				Countries[184] = "RW"; //RUANDA
				Countries[185] = "RO"; //RUMANIA
				Countries[186] = "EH"; //SAHARA OCCIDENTAL
				Countries[187] = "WS"; //SAMOA
				Countries[188] = "SM"; //SAN MARINO
				Countries[189] = "VC"; //SAN VICENTE Y LAS GRANADINAS
				Countries[190] = "SH"; //SANTA HELENA
				Countries[191] = "KN"; //SANTA KITTS Y NEVIS 
				Countries[192] = "LC"; //SANTA LUCIA
				Countries[193] = "PM"; //SANTO PIER Y MIKELON
				Countries[194] = "NU"; //SAVAGE ISLA DEL PACFICO
				Countries[195] = "SN"; //SENEGAL
				Countries[196] = "CS"; //SERBIA Y MONTENEGRO
				Countries[197] = "SC"; //SEYCHELLES
				Countries[198] = "SL"; //SIERRA LEONA
				Countries[199] = "SG"; //SINGAPUR
				Countries[200] = "SO"; //SOMALIA
				Countries[201] = "LK"; //SRI LANKA
				Countries[202] = "SD"; //SUDAN
				Countries[203] = "SE"; //SUECIA
				Countries[204] = "CH"; //SUIZA
				Countries[205] = "ZA"; //SUR AFRICA
				Countries[206] = "SR"; //SURINAM
				Countries[207] = "SJ"; //SVALBARD Y ENERO MAYEN
				Countries[208] = "SZ"; //SWAZILANDIA
				Countries[209] = "TH"; //TAILANDIA
				Countries[210] = "TW"; //TAIWAN PROVINCIA DE CHINA
				Countries[211] = "TJ"; //TAJIKISTAN
				Countries[212] = "TL"; //TIMOR-LESTE
				Countries[213] = "TG"; //TOGO
				Countries[214] = "TK"; //TOKELAU
				Countries[215] = "ST"; //TOMO DE SAO Y PRINCIPE
				Countries[216] = "TO"; //TONGA
				Countries[217] = "TT"; //TRINIDAD Y TOBAGO
				Countries[218] = "TN"; //TUNISIA
				Countries[219] = "TR"; //TURQUÍA
				Countries[220] = "TM"; //TURKMENISTN
				Countries[221] = "TV"; //TUVALU
				Countries[222] = "UA"; //UCRANIA
				Countries[223] = "UG"; //UGANDA
				Countries[224] = "UY"; //URUGUAY
				Countries[225] = "UZ"; //UZBEKISTN
				Countries[226] = "VU"; //VANUATU
				Countries[227] = "VE"; //VENEZUELA
				Countries[228] = "VN"; //VIETNAM
				Countries[229] = "WF"; //WALLIS Y FUTUNA
				Countries[230] = "YE"; //YEMEN
				Countries[231] = "ZM"; //ZAMBIA
				Countries[232] = "ZW"; //ZIMBABWE
				Countries[233] = "N1"; //NICARAGUA-GRH
                Countries[234] = "A2"; //NICARAGUA-NAZEL
                Countries[235] = "GTLTF"; //LATIN FREIGHT GT
                Countries[236] = "SVLTF"; //LATIN FREIGHT SV
                Countries[237] = "HNLTF"; //LATIN FREIGHT HN
                Countries[238] = "NILTF"; //LATIN FREIGHT NI
                Countries[239] = "CRLTF"; //LATIN FREIGHT CR
                Countries[240] = "PALTF"; //LATIN FREIGHT PA
                Countries[241] = "BZLTF"; //LATIN FREIGHT BZ
                Countries[242] = "HN1"; //HONDURAS TGU
                Countries[243] = "MXLTF"; //LATIN FREIGHT MX
                Countries[244] = "HN2"; //HONDURAS SPS / TGU

                //2021-10-13 hoy se agregaron estos paises El Incidente # 2876
                
                Countries[245] = "PALGX"; //PA LOGISTICS
                Countries[246] = "CRLGX"; //CR LOGISTICS
                Countries[247] = "GTLGX"; //GT LOGISTICS

                Countries[248] = "GTTLA"; //  GT
                Countries[249] = "SVTLA"; //  SV
                Countries[250] = "HNTLA"; //  HN
                Countries[251] = "NITLA"; //  NI
                Countries[252] = "CRTLA"; //  CR
                Countries[253] = "PATLA"; //  PA
                Countries[254] = "BZTLA"; //  BZ

				
				for (i=0; i<Countries.length;i++) //Ordenando por paises
				{
					for (j=0; j<=Lengs[13];j++)
					{
						if (Lengs[10][j] == Countries[i])
						{	//alert(k);
							Lengs[48][k] = Lengs[39][j] //BLDetailID
							Lengs[49][k] = k //Position
							Lengs[50][k] = Lengs[40][j] //BLIDTransit
							Lengs[51][k] = Lengs[0][j] //NoOfPieces
							Lengs[52][k] = Lengs[1][j] //ClassNoOfPieces
							Lengs[53][k] = Lengs[41][j] //CommoditiesID
							Lengs[54][k] = Lengs[2][j] //DiceContener
							Lengs[55][k] = Lengs[3][j] //DiceContenerValue
							Lengs[56][k] = Lengs[4][j] //Volumes
							Lengs[57][k] = Lengs[5][j] //Weights
							Lengs[58][k] = Lengs[42][j] //ClientsID
							Lengs[59][k] = Lengs[43][j] //AddressesID
							Lengs[60][k] = Lengs[6][j] //Clients
							Lengs[61][k] = Lengs[7][j] //BLs
							Lengs[62][k] = Lengs[8][j] //DischargeDate
							Lengs[63][k] = Lengs[9][j] //CountriesOrigen
							Lengs[64][k] = Lengs[10][j] //CountriesFinalDes
							Lengs[65][k] = Lengs[44][j] //InTransit
							Lengs[66][k] = Lengs[45][j] //AgentsID
							Lengs[67][k] = Lengs[46][j] //AgentsAddrID
							Lengs[68][k] = Lengs[11][j] //Agents
							Lengs[69][k] = Lengs[47][j] //HBLNumber
							Lengs[70][k] = Lengs[12][j] //Seps
							k++;
						};
					};
				}
			} else {
				//if (document.forma.BLType.value==1) { //Express
					for (j=0; j<=Lengs[13];j++) {
						Lengs[48][k] = Lengs[39][j] //BLDetailID
						Lengs[49][k] = k //Position
						Lengs[50][k] = Lengs[40][j] //BLIDTransit
						Lengs[51][k] = Lengs[0][j] //NoOfPieces
						Lengs[52][k] = Lengs[1][j] //ClassNoOfPieces
						Lengs[53][k] = Lengs[41][j] //CommoditiesID
						Lengs[54][k] = Lengs[2][j] //DiceContener
						Lengs[55][k] = Lengs[3][j] //DiceContenerValue
						Lengs[56][k] = Lengs[4][j] //Volumes
						Lengs[57][k] = Lengs[5][j]	//Weights
						Lengs[58][k] = document.forma.ConsignerID.value	//Lengs[42][j] //ClientsID
						Lengs[59][k] = document.forma.ConsignerAddrID.value //Lengs[43][j] //AddressesID
						Lengs[60][k] = document.forma.Clients.value //2020-08-20 Consignee   //Lengs[6][j] //Clients
						//Lengs[61][k] = Lengs[7][0] //BLs
						Lengs[61][k] = Lengs[7][j] //BLs
						//Lengs[62][k] = Lengs[8][0] //DischargeDate CON, EXP
						Lengs[62][k] = Lengs[8][j] //DischargeDate CON, EXP
						//Lengs[63][k] = CountryShipper	//Lengs[9][0] //CountriesOrigen
						Lengs[63][k] = Lengs[9][k] //CountriesOrigen
						//Lengs[64][k] = CountryConsignee	//Lengs[10][0] //CountriesFinalDes
						Lengs[64][k] = Lengs[10][j] //CountriesFinalDes
						Lengs[65][k] = Lengs[44][j] //InTransit
						//Lengs[66][k] = Lengs[45][0] //AgentsID
						//Lengs[67][k] = Lengs[46][0] //AgentsAddrID
						//Lengs[68][k] = Lengs[11][0] //Agents
						Lengs[66][k] = Lengs[45][j] //AgentsID
						Lengs[67][k] = Lengs[46][j] //AgentsAddrID
						Lengs[68][k] = Lengs[11][j] //Agents
						//Lengs[69][k] = HBLN //Lengs[47][j] //HBLNumber
						Lengs[69][k] = Lengs[47][j] //HBLNumber
						//Lengs[70][k] = Lengs[12][0] //Seps
						Lengs[70][k] = Lengs[12][j] //Seps
						k++;
					}
				/*} else { //Recoleccion-Entrega
					for (j=0; j<=Lengs[13];j++) {
						Lengs[48][k] = Lengs[39][j] //BLDetailID
						Lengs[49][k] = k //Position
						Lengs[50][k] = Lengs[40][j] //BLIDTransit
						Lengs[51][k] = Lengs[0][j] //NoOfPieces
						Lengs[52][k] = Lengs[1][j] //ClassNoOfPieces
						Lengs[53][k] = Lengs[41][j] //CommoditiesID
						Lengs[54][k] = Lengs[2][j] //DiceContener
						Lengs[55][k] = "0.00" //Lengs[3][j] //DiceContenerValue
						Lengs[56][k] = Lengs[4][j] //Volumes
						Lengs[57][k] = Lengs[5][j]	//Weights
						Lengs[58][k] = document.forma.ConsignerID.value	//Lengs[42][j] //ClientsID
						Lengs[59][k] = document.forma.ConsignerAddrID.value //Lengs[43][j] //AddressesID
						Lengs[60][k] = Consignee   //Lengs[6][j] //Clients
						Lengs[61][k] = " " //Lengs[7][j] //BLs
						Lengs[62][k] = ' ' //Lengs[8][j] //DischargeDate REC, ENT
						Lengs[63][k] = CountryShipper	//Lengs[9][j] //CountriesOrigen
						Lengs[64][k] = CountryConsignee	//Lengs[10][j] //CountriesFinalDes
						Lengs[65][k] = Lengs[44][j] //InTransit
						Lengs[66][k] = "414" //Lengs[45][j] //AgentsID
						Lengs[67][k] = "19729" //Lengs[46][j] //AgentsAddrID
						Lengs[68][k] = 'AIMAR' //Lengs[11][j] //Agents
						Lengs[69][k] = HBLN //Lengs[47][j] //HBLNumber
						Lengs[70][k] = Lengs[12][j] //Seps
						k++;
					};
				};*/	
			};
			document.forma.BLDetailID.value = Lengs[48].join("|");
			document.forma.Pos.value = Lengs[49].join("|");
			document.forma.BLIDTransit.value = Lengs[50].join("|");
			document.forma.NoOfPieces.value = Lengs[51].join("\n");
			document.forma.ClassNoOfPieces.value = Lengs[52].join("\n");
			document.forma.CommoditiesID.value = Lengs[53].join("|");
			document.forma.DiceContener.value = Lengs[54].join("\n");
			document.forma.DiceContenerValue.value = Lengs[55].join("\n");
			document.forma.Volumes.value = Lengs[56].join("\n");
			document.forma.Weights.value = Lengs[57].join("\n");
			document.forma.ClientsID.value = Lengs[58].join("|");
			document.forma.AddressesID.value = Lengs[59].join("|");
			document.forma.Clients.value = Lengs[60].join("\n");
			document.forma.BLs.value = Lengs[61].join("\n");
			document.forma.DischargeDate.value = Lengs[62].join("\n");
			document.forma.CountriesOrigen.value = Lengs[63].join("\n");
			document.forma.CountriesFinalDes.value = Lengs[64].join("\n");
			document.forma.InTransit.value = Lengs[65].join("|");
			document.forma.AgentsID.value = Lengs[66].join("|");
			document.forma.AgentsAddrID.value = Lengs[67].join("|");
			document.forma.Agents.value = Lengs[68].join("\n");
			document.forma.HBLNumber.value = Lengs[69].join("|");
			document.forma.Seps.value = Lengs[70].join("\n");
			//alert (document.forma.BLDetailID.value);
			//alert (document.forma.Pos.value);
			//alert (document.forma.BLIDTransit.value);
			//alert (document.forma.NoOfPieces.value);
			//alert (document.forma.ClassNoOfPieces.value);
			//alert (document.forma.CommoditiesID.value);
			//alert (document.forma.DiceContener.value);
			//alert (document.forma.Volumes.value);
			//alert (document.forma.Weights.value);
			//alert (document.forma.ClientsID.value)
			//alert (document.forma.AddressesID.value);
			//alert (document.forma.Clients.value);			
			//alert (document.forma.BLs.value);
			//alert (document.forma.DischargeDate.value);
			//alert (document.forma.CountriesOrigen.value);
			//alert (document.forma.CountriesFinalDes.value);
			//alert (document.forma.InTransit.value);
			//alert (document.forma.AgentsID.value);
			//alert (document.forma.AgentsAddrID.value);
			//alert (document.forma.Agents.value);
			//alert (document.forma.HBLNumber.value);
			//alert (document.forma.Seps.value);

            SumVals(document.forma.NoOfPieces, document.forma.TotNoOfPieces);
            SumVals(document.forma.Volumes, document.forma.TotVolume);
            SumVals(document.forma.Weights, document.forma.TotWeight);

            move();
		    document.forma.Action.value = Action;
		    document.forma.submit();			
		}
        else if (Action == 3)
        {
            var Hinvs = document.forma.DetailInvoices.value.split("|").join("\t\t").split("*").join("\n");
            if (Hinvs.length > 0)
            {
                //alert("No puede eliminar la carta porte porque tiene " + <%=CountList6Values + 1%> + " documentos contables asociados: \n\n RO/BL \t\t\t Tipo \t\t SERIE \t\t CORR \t\t EMPRESA \n\n" + Hinvs);
                alert("No puede eliminar la carta porte porque tiene " + <%=CountList6Values + 1%> + " documentos contables asociados: \r\nRO/BL\t\t\tSE-CO\t\t\tPAIS\r\n" + Hinvs);
                alert("Si desea eliminar esta carta de porte, antes debe anular los documentos contables de cada carta porte hija.");
                return (false);
            }
            else
            {        
                SumVals(document.forma.NoOfPieces, document.forma.TotNoOfPieces);
                SumVals(document.forma.Volumes, document.forma.TotVolume);
                SumVals(document.forma.Weights, document.forma.TotWeight);

                document.forma.Action.value = Action;
		        document.forma.submit();
            }
        }
	}

	function CheckDate(Val)
	{
		var Day;
		var Month;
		var Year;

		Day = Val.substr(0,2)*1;
		Month = Val.substr(3,2)*1;
		Year = Val.substr(6,4)*1;

		if ((isNaN(Day))||(isNaN(Month))||(isNaN(Year))) {
			return(false);
		} else {
			if ((Day > 31) || (Day < 1)) {
				return(false);
			} else {
				if ((Month > 12) || (Month < 1)) {
					return(false);
				} else {
					if ((Year > 3000) || (Year < 2007)) {
						return(false);
					} else {
						return(true);
					}			
				}
			}
		}
	}
	
	function SetClassNoOfPieces(obj) {
	var Vals = FixCart(obj);//.value.split("\r\n");
	var ValsLen = Vals.length-1;
	var ClassPieces = FixCart(document.forma.ClassNoOfPieces);//.value.split("\r\n");
	var ClassPiecesLen = ClassPieces.length-1;	
	//alert (ValsLen + "-" + ClassPiecesLen);
		if (ValsLen > ClassPiecesLen) {
			for (i=0;i<ValsLen-ClassPiecesLen;i++) {
				document.forma.ClassNoOfPieces.value = document.forma.ClassNoOfPieces.value + "\r\n"; 	
			}
		} else {
			if (ClassPiecesLen==0 && document.forma.ClassNoOfPieces.value=="") {
				document.forma.ClassNoOfPieces.value = " "; 	
			}
		}
	}

	function SetDiceContenerValue(obj) {
	var Vals = FixCart(obj);//.value.split("\r\n");
	var ValsLen = Vals.length-1;
	var ContenerVals = FixCart(document.forma.DiceContenerValue);//.value.split("\r\n");
	var ContenerValsLen = ContenerVals.length-1;	
	//alert (ValsLen + "-" + ClassPiecesLen);
		if (ValsLen > ContenerValsLen) {
			for (i=0;i<ValsLen-ContenerValsLen;i++) {
				document.forma.DiceContenerValue.value = document.forma.DiceContenerValue.value + "\r\n0.00"; 	
			}
		} else {
			if (ContenerValsLen==0) {
				document.forma.DiceContenerValue.value = "0.00"; 	
			}
		}
	}

    function FixCart(obj) {
    	var Vals = obj.value.split("\r\n");
		if (Vals.length == 1) {
			var Vals2 = Vals[0].split("\n");
			if (Vals2.length > 1) {
				Vals = obj.value.split("\n");
                console.log(Vals);
			}
		}		
		return Vals;	
	}
	 
	function SumVals(obj, destination) {

    var Vals = FixCart(obj);
	var TotVals = 0;
	var Values = "";
	var Val;
	var ExistsNaN = false;
	var ValsLen = Vals.length-1;
	var spacer = "";
	
		for (i=0;i<=ValsLen;i++) { 
			if (Vals[i] != ""){
				Val = Vals[i]*1;
				if (isNaN(Val) == false) {
					TotVals = TotVals + (Vals[i]*1);
				} else {
					ExistsNaN = true;			
				}
				Values = Values + spacer + Vals[i];
				spacer = "\r\n";
				//if (i!=ValsLen){
				//	Values = Values + "\r\n";
				//}
			}
		}

		obj.value = Values;
		if (!ExistsNaN) {
			destination.value = Round(TotVals);
		} else {
			alert ("Solo debe ingresar Numeros en la casilla '" + obj.id + "'");
			destination.value = "";
			obj.focus();
		}
	}
	
	function Round(value){
		var number = (Math.round(value * 100)) / 100;
		return (number == Math.floor(number)) ? number + '.00' : ((number * 10 == Math.floor(number * 10)) ? number + '0' : number);
	}
	
	var numb = "0123456789./\r/\n";
	function res(t,v){
		var w = "";
		for (i=0; i < t.value.length; i++) {
		x = t.value.charAt(i);
		if (v.indexOf(x,0) != -1)
			w += x;
		}
		t.value = w;
	}

	var Lics = new Array();
	var Pils = new Array();
	var Warehouses = new Array();
	
	<%'Desplegando Datos de Pilotos
	For i = 0 To CountList2Values
		j = i+1
		response.write "Lics[" & j & "]='" & aList2Values(2,i) & "';	Pils[" & j & "]=" & aList2Values(0,i) & ";" & vbCrLf
	Next
	
	'Desplegando Datos de Bodegas
	'WarehouseID, Countries, Name, Address, Address2, Phone1, Phone2, Attn
	For i = 0 To CountList4Values
		for j = 4 to 7
			if aList4Values(j,i) <> "" then
				aList4Values(3,i) = aList4Values(3,i) & "\n" & aList4Values(j,i)
			end if
		Next
		response.write "Warehouses[" & aList4Values(0,i) & "]='" & aList4Values(2,i) & "\n" & aList4Values(3,i) & "';" & vbCrLf
	Next
	%>
	
	function SetLics(pos){
		if (pos > 0) {
			document.forma.License.value = Lics[pos];
		} else {
			document.forma.License.value = "";
		}
	}
	
	function getWeekNr()
	{
		NumberOfWeek = "";
		BLExitDate = document.forma.BLExitDate.value;
		if (BLExitDate != "") {
			Day = BLExitDate.substr(0,2)*1;
			Month = (BLExitDate.substr(3,2)*1)-1;
			Year = BLExitDate.substr(6,4)*1;
			
			now = Date.UTC(Year,Month,Day+1,0,0,0);
			var Firstday = new Date();
			Firstday.setYear(Year);
			Firstday.setMonth(0);
			Firstday.setDate(1);
			then = Date.UTC(Year,0,1,0,0,0);
			var Compensation = Firstday.getDay();
			if (Compensation > 3) Compensation -= 4;
			else Compensation += 3;
			NumberOfWeek =  Math.round((((now-then)/86400000)+Compensation)/7);
		}
		return NumberOfWeek;		
	}	

	function takeYear(theDate)
	{
		x = theDate.getYear();
		var y = x % 100;
		y += (y < 38) ? 2000 : 1900;
		return y;
	}
	
	function EnumDiceContener(Val, Id)
	{
		var i=0;
		var Vals = document.forma.DiceContener.value.split("|");
		var Vals2 = document.forma.CommoditiesID.value.split("|");
		var ValsLen = Vals.length-1;
		ntr = "";
		com = "";		

		document.forma.Contener.value = "";
			
		if (document.forma.DiceContener.value != "")
		{
			document.forma.DiceContener.value = "";
			document.forma.CommoditiesID.value = "";

			for (i=0;i<=ValsLen;i++) { 
				document.forma.DiceContener.value = document.forma.DiceContener.value + com + Vals[i];
				document.forma.Contener.value = document.forma.Contener.value + ntr + (i+1) + "-" + Vals[i];
				document.forma.CommoditiesID.value = document.forma.CommoditiesID.value + com + Vals2[i];
				ntr = "\n";
				com = "|";
			}
		}
		
		if (Val != "") {
			document.forma.DiceContener.value = document.forma.DiceContener.value + com + Val;
			document.forma.Contener.value = document.forma.Contener.value + ntr + (i+1) + "-" + Val;
			document.forma.CommoditiesID.value = document.forma.CommoditiesID.value + com + Id;
		}
	}
	
	function EraseLines(Line)
	{
		//Si la carga aun no tiene facturas, es posible eliminarlo de la carta porte, caso contrario no debido a que ya tiene documentos relacionados
        //con el numero de carta porte a la que esta asociado

        //var Hinvs = document.forma.HaveInvoices.value.split("|");

        var x = FixCart(document.forma.BLs);//.value.split("\r\n");
        var y = document.forma.DetailInvoices.value.split("*");
        var z = "";

        for (i=0;i<y.length;i++)
        {
            if(y[i].indexOf(x[Line]) >= 0)
            {
                z = z + y[i].split("|").join("\t\t") + "\n";
            }; 
        };

		if (z.length == 0) {
            SetDetailToErase (document.forma.BLDetailID, document.forma.BLIDTransit, Line, "|", "|");

		    EraseLine (document.forma.NoOfPieces, Line, "\r\n", "\n");
		    EraseLine (document.forma.ClassNoOfPieces, Line, "\r\n", "\n");
		    EraseLine (document.forma.CommoditiesID, Line, "|", "|");
		    EraseLine (document.forma.DiceContener, Line, "|", "|");
		    EraseLine (document.forma.DiceContenerValue, Line, "\r\n", "\n");
		    EraseLine (document.forma.Volumes, Line, "\r\n", "\n");
		    EraseLine (document.forma.Weights, Line, "\r\n", "\n");
		    EraseLine (document.forma.ClientsID, Line, "|", "|");
		    EraseLine (document.forma.AddressesID, Line, "|", "|");
		    EraseLine (document.forma.Clients, Line, "\r\n", "\n");
		    EraseLine (document.forma.BLs, Line, "\r\n", "\n");
		    EraseLine (document.forma.DischargeDate, Line, "\r\n", "\n");
		    EraseLine (document.forma.CountriesOrigen, Line, "\r\n", "\n");
		    EraseLine (document.forma.CountriesFinalDes, Line, "\r\n", "\n");
		    EraseLine (document.forma.InTransit, Line, "|", "|");
		    EraseLine (document.forma.AgentsID, Line, "|", "|");
		    EraseLine (document.forma.AgentsAddrID, Line, "|", "|");
		    EraseLine (document.forma.Agents, Line, "\r\n", "\n");
		    EraseLine (document.forma.Seps, Line, "\r\n", "\n");
		
		    EnumDiceContener("", "");
		
		    SumVals(document.forma.NoOfPieces, document.forma.TotNoOfPieces);
		    SetClassNoOfPieces(document.forma.NoOfPieces);
		    <%if Countries="GT" or Countries="SV" then%>SetDiceContenerValue(document.forma.NoOfPieces);<%end if%>
		    SumVals(document.forma.Volumes, document.forma.TotVolume);
		    SumVals(document.forma.Weights, document.forma.TotWeight);
		    SumVals(document.forma.DiceContenerValue,document.forma.TotDiceContenerValue);
        } else {
            
            alert("No puede eliminar la carga porque tiene los siguientes documentos contables asociados: \r\nRO/BL\t\t\tSE-CO\t\t\tPAIS\r\n" + z);
            
            alert("Si desea eliminar dicha carga, antes debe anular los documentos contables correspondientes.");
            
            // SetDetailToErase (document.forma.BLDetailID, document.forma.BLIDTransit, Line, "|", "|");

            return (false);
        }
	}
	
	function SetDetailToErase(obj1, obj2, Line, sepSplit, sepJoin)
	{
		var Vals1 = obj1.value.split(sepSplit);
		var Vals2 = obj2.value.split(sepSplit);
		var ValsLen = Vals1.length-1;
		ntr = "";
		obj1.value = "";
		obj2.value = "";

		for (i=0; i<=ValsLen; i++) {
			if (i != Line) {
				obj1.value = obj1.value + ntr + Vals1[i];
				obj2.value = obj2.value + ntr + Vals2[i];
				ntr = sepJoin;
			} else {
				if (Vals1[i] != 0) {
					if (document.forma.DetailToErase.value == "") {
						document.forma.DetailToErase.value = Vals1[i];
					} else {
						document.forma.DetailToErase.value = document.forma.DetailToErase.value + "|" + Vals1[i];
					}				
				}
				if (Vals2[i] != 0) {
					if (document.forma.DetailToRestore.value == "") {
						document.forma.DetailToRestore.value = Vals2[i];
					} else {
						document.forma.DetailToRestore.value = document.forma.DetailToRestore.value + "|" + Vals2[i];
					}
				}
			}
		}	
	}	

    function EraseLine(obj, Line, sepSplit, sepJoin)
	{
        var Vals; 
            
        if (sepSplit == "|")
		    Vals = obj.value.split(sepSplit);
        else
		    Vals = FixCart(obj)

		var ValsLen = Vals.length-1;
		ntr = "";
		obj.value = "";
		
		for(i=0; i<=ValsLen; i++) {
			if (i != Line) {
				obj.value = obj.value + ntr + Vals[i];
				ntr = sepJoin;
			}
		}	
	}

	function CleanSpaces(obj, sepSplit, sepJoin)
	{
        var Vals;

        if (sepSplit == "|")
		    Vals = obj.value.split(sepSplit);
        else
		    Vals = FixCart(obj)

		var ValsLen = Vals.length-1;
		ntr = "";
		obj.value = "";
		
		for(i=0; i<=ValsLen; i++) {
			if (Vals[i] != "") {
				obj.value = obj.value + ntr + Vals[i];
				ntr = sepJoin;
			}
		}	
	}	

	function SetConsolidated()
	{
		if (document.forma.BLType.value==0) {
			document.forma.Consolidated.value=1;
			document.getElementById("Consig").style.visibility = "visible";
            <%'Se da acceso solo a usuarios Cesar Sanchez, Gabriel Morales, Edgard Campos y Erlin Carcamo para editar fechas estimadas en BLs Consolidados
            if Session("OperatorID")=318 or Session("OperatorID")=782 or Session("OperatorID")=1055 or Session("OperatorID")=248 then%>
            document.getElementById("EstExitDate").style.visibility = "visible";
            document.getElementById("EstArrivalDate").style.visibility = "visible";
            <%else %>
            document.getElementById("EstExitDate").style.visibility = "hidden";
            document.getElementById("EstArrivalDate").style.visibility = "hidden";
            <%end if %>

            document.forma.BLExitDate.value = "<%=BLExitDate%>";
            document.forma.BLEstArrivalDate.value = "<%=BLEstArrivalDate%>";
		} else {
			document.forma.Consolidated.value=0;
			document.getElementById("Consig").style.visibility = "hidden";
            document.getElementById("EstExitDate").style.visibility = "visible";
            document.getElementById("EstArrivalDate").style.visibility = "visible";
		}
		
		//if (document.forma.Consolidated.value==0) {
			//document.forma.ClientsID.value = "";
			//document.forma.AddressesID.value = "";
			//document.forma.Clients.value = "";
			//document.forma.CountriesOrigen.value = "";			
			//document.forma.CountriesFinalDes.value = "";			
		//}
		//DisplayDivs();		
	}
	
	function DisplayDivs() {
		if ((document.forma.BLType.value==0)||(document.forma.BLType.value==1)) {
			for(i=1; i<=22; i++) {
			document.getElementById("D" + i).style.visibility = "visible";
			}
			for(i=1; i<=8; i++) {
			document.getElementById("I" + i).style.visibility = "hidden";
			}
		} else {
			for(i=1; i<=22; i++) {
			document.getElementById("D" + i).style.visibility = "hidden";
			}
			for(i=1; i<=8; i++) {
			document.getElementById("I" + i).style.visibility = "visible";
			}
		}
		DisplayAddr(document.getElementById("ChargeType"), 'ChargePlace', 0);
		DisplayAddr(document.getElementById("DestinyType"), 'FinalDes', 0);
	}
	
	function DisplayAddr(obj, element, action) {
		if (obj.value>=0) {
			var LabelID = SetLabelID(element);
			if (obj.value==0) {
				document.getElementById(LabelID).readOnly = "";
				if (action==1) {
					document.getElementById(LabelID).value = "";
				}
			} else {
				document.getElementById(LabelID).readOnly = "readOnly";
				if (action==1) {
				document.getElementById(LabelID).value = Warehouses[obj.value];
				}
				
				<%'if BLType=2 then%>
				//if (obj.value==WareHouseID) {
				//	document.getElementById("Marchamo").value = Marchamo;
				//} else {
				//	document.getElementById("Marchamo").value = "";
				//}
				<%'end if%>
			}
		}
	}

    function move() {
        document.forma.style.display = "none";
        document.getElementById('myBar').style.display = "block";
        var elem = document.getElementById("myBar");
        var width = 10;
        var id = setInterval(frame, 45);
        function frame() {
            if (width >= 100) {
                clearInterval(id);
            } else {
                width++;
                elem.style.width = width + '%';
                elem.innerHTML = width * 1 + '%';
            }
        }
    }

</SCRIPT>
<body>
<%if JavaMsg <> "" then %>
    <div class=label><font color=<%if InStr(BAWResult,"Exitosamente") then %>blue<%else %>red<%end if %>><%=Replace(JavaMsg,"\n","<br>")%></font></div>
<%end if %>
<div id="myProgress">
  <div id="myBar">10%</div>
</div>
<form name="forma" action="InsertData.asp" method="post">
	<INPUT name="Action" type=hidden value=0>
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
	<INPUT name="OID" type=hidden value="<%=ObjectID%>">
	<INPUT name="CD" type=hidden value="<%=CreatedDate%>">
	<INPUT name="CT" type=hidden value="<%=CreatedTime%>">
	<INPUT name="BT" type=hidden value="<%=BLType%>">
	<INPUT name="IT" type=hidden value="<%=ItineraryType%>">
	<INPUT name="SenderID" type=hidden value="<%=SenderID%>">
    <INPUT name="ShipperID" type=hidden value="<%=ShipperID%>">
	<INPUT name="ConsignerID" type=hidden value="<%=ConsignerID%>">
    <INPUT name="AgentNeutral" type=hidden value="<%=AgentNeutral%>">
    <INPUT name="GuiaRemision" type=hidden value="<%=GuiaRemision%>">
    <INPUT name="BLArrivalDate" type=hidden value="<%=BLArrivalDate%>">
    <INPUT name="ShipperColoader" type=hidden value="<%=ShipperColoader%>">    
	<INPUT name="ConsignerColoader" type=hidden value="<%=ClientColoader%>">    
	<INPUT name="SenderAddrID" type=hidden value="<%=SenderAddrID%>">
	<INPUT name="ShipperAddrID" type=hidden value="<%=ShipperAddrID%>">
	<INPUT name="ConsignerAddrID" type=hidden value="<%=ConsignerAddrID%>">
	<INPUT name="BLDetailID" type=hidden value="<%=Join(BLDetailID, "|")%>">
	<INPUT name="Pos" type=hidden value="">
	<INPUT name="BLIDTransit" type=hidden value="<%=Join(BLIDTransit, "|")%>">
	<INPUT name="CommoditiesID" type=hidden value="<%=Join(CommoditiesID, "|")%>">
	<INPUT name="DiceContener" type=hidden value="<%=Join(DiceContener, "|")%>">
	<INPUT name="ClientsID" type=hidden value="<%=Join(ClientsID, "|")%>">
	<INPUT name="AddressesID" type=hidden value="<%=Join(AddressesID, "|")%>">
	<INPUT name="AgentsID" type=hidden value="<%=Join(AgentsID, "|")%>">
	<INPUT name="AgentsAddrID" type=hidden value="<%=Join(AgentsAddrID, "|")%>">
	<INPUT name="InTransit" type=hidden value="<%=Join(InTransit, "|")%>">
	<INPUT name="HBLNumber" type=hidden value="<%=Join(HBLNumber, "|")%>">
    <INPUT name="HaveInvoices" type=hidden value="<%=Join(HaveInvoices, "|")%>">
    <INPUT name="DetailInvoices" type=hidden value="<%=Join(DetailInvoices, "*")%>">
    <INPUT name="DetailToErase" type=hidden value="">
    <INPUT name="BLsToErase" type=hidden value="">
	<INPUT name="DetailToRestore" type=hidden value="">	
	<INPUT name="Consolidated" type=hidden value="<%=Consolidated%>">
	<INPUT name="Closed" type=hidden value="<%=Closed%>">
	<%if ItineraryType=2 then%>
	<INPUT name="CountryDes" type=hidden value="<%=Countries%>">
	<INPUT name="CountryDep" type=hidden value="<%=Countries%>">
	<%end if%>
	
<table width="1076" border="1" cellpadding="2" cellspacing="0" align="center">
  <tr>  	
    <td class="style4" align="center" width="35%"><%if BLNumber<>"" then%>CARTA&nbsp;PORTE&nbsp;No.&nbsp;<%=BLNumber%><% end if%>&nbsp;</td>
	<td class="style4" align="center" width="15%"><%if ObjectID <> 0 then%><a href="#" onClick="Javascript:GetDoc(<%=ObjectID%>);return (false);" class="menu"><font color="FFFFFF">&nbsp;Archivos&nbsp;</font></a><%end if%>&nbsp;</td>
	<td class="style4" align="right" width="20%">Tipo&nbsp;de&nbsp;Transporte:</td>
    <td class="style4" align="left" bgcolor="#999999" width="15%">
		<select class="style10" name="BLType" id="Tipo de Transporte" onChange="Javascript:SetConsolidated();">
			<%if ItineraryType=1 then%>
			<option value="0">CONSOLIDADO</option>
			<option value="1">EXPRESS</option>
			<%else%>
			<option value="2">LOCAL</option>
			<!--<option value="3">ENTREGA</option>-->
			<%end if%>
		</select>
     </td>
      <% CountGuiaRemision = int(len(cstr(GuiaRemision))) 
         select case CountGuiaRemision
       case 1
			NumeroCerosGuiaRemision = "0000000"
       case 2
            NumeroCerosGuiaRemision = "000000"
       case 3
            NumeroCerosGuiaRemision = "00000"
       case 4
            NumeroCerosGuiaRemision = "00000"
       case 5
            NumeroCerosGuiaRemision = "0000"
       case 6
            NumeroCerosGuiaRemision = "000"
       case 7
            NumeroCerosGuiaRemision = "000"
       case 8
            NumeroCerosGuiaRemision = "00"
	   case else
			NumeroCerosGuiaRemision = "0"
	   end select
    %>
    <td class="style4" align="right" width="10%"><%if CountryDep="HN" or CountryDes="HN" or CountryDep="HN1" or CountryDes="HN1" or CountryDep="HN2" or CountryDes="HN2" then %>GUIA&nbsp;DE&nbsp;REMISION&nbsp;000-001-08-<%=NumeroCerosGuiaRemision %><%=GuiaRemision%><% end if%></td>
	<td class="style4" align="right" width="10%">Semana:</td>
    <td class="style4" align="left" bgcolor="#999999" width="5%">
        <select class="style10" name="Week" id="Semana">
            <%if Week = 0 then %>
                <option value=<%=WeekPrev2%>><%=WeekPrev2&"-"&YearPrev2%></option>
                <option value=<%=WeekPrev%>><%=WeekPrev&"-"&YearPrev%></option>
                <option value=<%=WeekAct%> selected="selected"><%=WeekAct&"-"&YearAct%></option>
                <option value=<%=WeekPost%>><%=WeekPost&"-"&YearPost%></option>
                <option value=<%=WeekPost2%>><%=WeekPost2&"-"&YearPost2%></option>
            <%else%>
                <option value=<%=Week%>><%=Week&"-"&Mid(BLNumber,5,4)%></option>
            <%end if %>
        </select>
        <!--<input class="style10" name="Week" value="<%=Week%>" id="Semana" type="text" size="5" maxlength="2" onKeyUp="res(this,numb);">-->
    </td>
  </tr>
</table>
<table width="1076" border="1" cellpadding="2" cellspacing="0" align="center">
  <%if ItineraryType=1 then%>
  <tr>
    <td class="style4" align="left">Exportador:</td>
    <td class="style4" align="left">Shipper / Embarcador:</td>
  </tr>
  <tr>
    <td class="style4" align="left" bgcolor="#999999">
	<textarea class="style10" name="SenderData" rows="5" cols="70" id="Exporter / Exportador"  readonly><%=SenderData%></textarea>
	<%if isnull(LtAcceptDate) then%>
	<table width="80%" cellpadding="0" cellspacing="0">
	<tr>
		<td class="style4" width="80%">&nbsp;</td>
		<td class="style4" align="right">
			<a class="submenu" href="http://10.10.1.20/catalogo_admin/login.php" target="_blank"><font color="FFFFFF">Nuevo</font></a>
		</td>
		<td class="style4">&nbsp;&nbsp;</td>
		<td class="style4" align="right">
			<a href="#" onClick="Javascript:GetData(2);return (false);" class="menu"><font color="FFFFFF">Buscar</font></a>
		</td>
	</tr>
	</table>
	<%end if%>
	<td class="style4" align="left" bgcolor="#999999">
	<textarea class="style10" name="ShipperData" rows="5" cols="70" id="Shipper / Embarcador"  readonly="readonly"><%=ShipperData%></textarea>
	<%if isnull(LtAcceptDate) then%>
	<table width="80%" cellpadding="0" cellspacing="0">
	<tr>
		<td class="style4" width="80%">&nbsp;</td>
		<td class="style4" align="right">
			<a class="submenu" href="http://10.10.1.20/catalogo_admin/login.php" target="_blank"><font color="FFFFFF">Nuevo</font></a>
		</td>
		<td class="style4">&nbsp;&nbsp;</td>
		<td class="style4" align="right">
			<a href="#" onClick="Javascript:GetData(3);return (false);" class="menu"><font color="FFFFFF">Buscar</font></a>
		</td>
	</tr>
	</table>
	<%end if%>
  </tr>
  <%end if%>
  <tr>
    <td class="style4" align="left">Consignee / Consignatario: </td>
    <td class="style4" align="left">
	<%if ItineraryType=1 then%>
	Pais de Origen (Transito):
	<%else%>
	P&oacute;liza:
	<%end if%></td>
  </tr>
  <tr>
    <td class="style4" align="left" bgcolor="#999999"><textarea class="style10" name="ConsignerData" rows="5" cols="70" id="Consignee / Consignatario"  readonly="readonly"><%=ConsignerData%></textarea>
	<%if ItineraryType=1 then%>
		<%if isnull(LtAcceptDate) then%>
		<table width="80%" cellpadding="0" cellspacing="0">
		<tr>
			<td class="style4" width="80%">&nbsp;</td>
			<td class="style4" align="right">
			<a class="submenu" href="http://10.10.1.20/catalogo_admin/login.php" target="_blank"><font color="FFFFFF">Nuevo</font></a>
			</td>
			<td class="style4">&nbsp;&nbsp;</td>
			<td class="style4" align="right">
				<a href="#" onClick="Javascript:GetData(4);return (false);" class="menu"><font color="FFFFFF">Buscar</font></a>
			</td>
		</tr>
		</table>
		<%end if%>
	<%else%>
		<table width="80%" cellpadding="0" cellspacing="0">
		<tr>
			<td class="style4" width="80%">&nbsp;</td>
			<td class="style4" align="right">
				<a href="#" onClick="Javascript:GetData(14);return (false);" class="menu"><font color="FFFFFF">Agregar</font></a>
			</td>			
		</tr>
		</table>	
	<%end if%>
	</td>
  	<td class="style4" align="left" valign="top" bgcolor="#999999">
	<%if ItineraryType=1 then%>
	<select class="style10" name="CountryDep" id="Pais de Origen">
		<option value='-1'>Seleccionar</option>
        <%DisplayCountries Countries, 2%>		
	</select>
	<%else%>
	<input class="style10" type="text" name="PolicyNo" value="<%=PolicyNo%>" id="No. de Poliza">
	<%end if%>&nbsp;
	</td>
  </tr>
  <tr>
    <td class="style4" align="left">Notificaci&oacute;n a:  </td>
    <td class="style4" align="left">
	<%if ItineraryType=1 then%>
	Instrucciones de Exportaci&oacute;n:
	<%else%>
	Fecha Liberaci&oacute;n de P&oacute;liza:
	<%end if%></td>
  </tr>
  <tr>
    <td class="style4" align="left" bgcolor="#999999"><textarea class="style10" name="Attn" id="Notificación a" rows="5" cols="70"><%=Attn%></textarea></td>
    <td class="style4" align="left" bgcolor="#999999" valign="top">
	<%if ItineraryType=1 then%>
	<textarea class="style10" name="HandlingInformation" rows="5" cols="70"><%=HandlingInformation%></textarea>
	<%else%>
	<INPUT readonly name="DeliveryPolicyDate" type=text value="<%=DeliveryPolicyDate%>" size=23 maxLength=19 class=label id="Fecha en que se libero la Poliza">&nbsp;<a href="#" onClick="JavaScript:abrir('DeliveryPolicyDate');return (false);" class="menu"><font color="FFFFFF">Seleccionar</font></a>
	<br>
	<select name="DeliveryPolicyHour" id="Hora en que se libero la Poliza" class="label">
			<option value="-1">Hora</option>
			<option value="0">00</option>
			<option value="1">01</option>
			<option value="2">02</option>
			<option value="3">03</option>
			<option value="4">04</option>
			<option value="5">05</option>
			<option value="6">06</option>
			<option value="7">07</option>
			<option value="8">08</option>
			<option value="9">09</option>
			<option value="10">10</option>
			<option value="11">11</option>
			<option value="12">12</option>
			<option value="13">13</option>
			<option value="14">14</option>
			<option value="15">15</option>
			<option value="16">16</option>
			<option value="17">17</option>
			<option value="18">18</option>
			<option value="19">19</option>
			<option value="20">20</option>
			<option value="21">21</option>
			<option value="22">22</option>
			<option value="23">23</option>
		</select>	
		<select name="DeliveryPolicyMin" id="Minutos en que se libero la Poliza" class="label">
			<option value="-1">Minuto</option>
			<option value="0">00</option>
			<option value="5">05</option>
			<option value="10">10</option>
			<option value="15">15</option>
			<option value="20">20</option>
			<option value="25">25</option>
			<option value="30">30</option>
			<option value="35">35</option>
			<option value="40">40</option>
			<option value="45">45</option>
			<option value="50">50</option>
			<option value="55">55</option>
		</select>	
	<%end if%>
	</td>
  </tr>
  <%if ItineraryType=1 then%>
  <tr>
    <td class="style4" align="left">Aduana de Tr&aacute;nsito:</td>
    <td class="style4" align="left">Pais de Destino (Transito):</td>
  </tr>
  <tr>
    <td class="style4" align="left" bgcolor="#999999">
	<select class="style10" name="BrokerID" id="Aduana">
	<option value="-1">Seleccionar</option>
	<%
		For i = 0 To CountList1Values
	%>
	<option value="<%=aList1Values(0,i)%>"><%response.write aList1Values(1,i) & " - " & aList1Values(2,i)%></option>
	<%
   		Next
	%>
	</select>
	</td>
  	<td class="style4" align="left" valign="top" bgcolor="#999999">
	<select class="style10" name="CountryDes" id="Pais de Destino">
		<option value='-1'>Seleccionar</option>
		<%DisplayCountries Countries, 2%>
	</select>
	</td>
  </tr>
  <%end if%>
  <tr>
    <td class="style4" align="left">Lugar de Carga:</td>
    <td class="style4" align="left">Destino Final / Entrega:</td>
  </tr>
  <tr>
    <td class="style4" align="left" bgcolor="#999999">
	<select class="style10" name="ChargeType" id="Lugar de  Carga" onChange="javascript:DisplayAddr(this,'ChargePlace',1);">
		<option value="-1">Seleccionar</option>
		<%		
			'WarehouseID, Countries, Name, Address, Address2, Phone1, Phone2, Attn, Email
			For i = 0 To CountList4Values
		%>
		<option value="<%=aList4Values(0,i)%>"><%response.write aList4Values(2,i) & " - " & aList4Values(1,i)%></option>
		<%
   			Next
		%>
		<option value="0">OTRO</option>
	</select><br>
	<textarea class="style10" name="ChargePlace" rows="5" cols="70" id="Lugar de Carga" readonly="readonly"><%=ChargePlace%></textarea>	
	</td>
    <td class="style4" align="left" bgcolor="#999999">
	<select class="style10" name="DestinyType" id="Destino Final /  Entrega" onChange="javascript:DisplayAddr(this,'FinalDes',1);">
		<option value="-1">Seleccionar</option>
		<%		
			'WarehouseID, Name, Countries, Address, Address2, Phone1, Phone2, Attn, Email
			For i = 0 To CountList4Values
		%>
		<option value="<%=aList4Values(0,i)%>"><%response.write aList4Values(2,i) & " - " & aList4Values(1,i)%></option>
		<%
   			Next
		%>
		<option value="0">OTRO</option>
	</select><br>
	<textarea class="style10" name="FinalDes" rows="5" cols="70" id="Destino Final / Entrega" readonly="readonly"><%=FinalDes%></textarea>
	</td>
  </tr>
  </table>
  <table width="1076" border="1" cellpadding="2" cellspacing="0" align="center">
  <tr>
    <td class="style4" align="left">Conductor: </td>
    <td class="style4" align="left">Transporte: </td>
    <td class="style4" align="left">Contenedor: </td>
    <td class="style4" align="left">Chasis: </td>
  </tr>
  <tr>
    <td class="style4" align="left" bgcolor="#999999">
	<select class="style10" name="PilotID" id="Piloto" onChange="Javascript:SetLics(this.selectedIndex);">
	<option value="-1">Seleccionar</option>
	<%
		For i = 0 To CountList2Values
	%>
	<option value="<%=aList2Values(0,i)%>"><%response.write aList2Values(1,i) & " - " & aList2Values(3,i)%></option>
	<%
   		Next
	%>
	</select>
	</td>
    <td class="style4" align="left" bgcolor="#999999">
	<select class="style10" name="TruckID" id="Cabezal" onchange="Javascript:ValidarPlaca(this);">
	<option value="-1">Seleccionar</option>
	<%
		For i = 0 To CountList3Values
			if aList3Values(3,i) = 0 or aList3Values(3,i) = 2 then
	%>
	<option value="<%=aList3Values(0,i)%>"><%response.write aList3Values(6,i) & " " & aList3Values(5,i) & " " & aList3Values(4,i) & " - " & aList3Values(2,i) & " " & aList3Values(7,i) & " " & aList3Values(8,i)%></option>
	<%
			end if
   		Next
	%>
	</select>
	</td>
    <td class="style4" align="left" bgcolor="#999999">
	<input class="style10" type="text" name="ContainerDep" value="<%=ContainerDep%>" id="Contenedor">
	</td>
    <td class="style4" align="left" bgcolor="#999999">
	<input class="style10" type="text" name="Chassis" value="<%=Chassis%>" id="Contenedor">	
	</td>
  </tr>
  <tr>
    <td class="style4" align="left" colspan="2">Licencia Conductor: </td>
    <td class="style4" align="left" colspan="2">Furgon TC: </td>
  </tr>
  <tr>
    <td class="style4" align="left" bgcolor="#999999" colspan="2">
		<input class="style10" type="text" name="License" value="<%=License%>" readonly>
	</td>	
    <td class="style4" align="left" bgcolor="#999999" colspan="2">
		<select class="style10" name="Container" id="Furgon">
	<option value="-1">Seleccionar</option>
	<%
		For i = 0 To CountList3Values
			if aList3Values(3,i) = 1 then '1=Furgon
	%>
	<option value="<%=aList3Values(0,i)%>"><%response.write aList3Values(6,i) & " " & aList3Values(5,i) & " " & aList3Values(4,i) & " - " & aList3Values(2,i)%></option>
	<%
			end if
   		Next
	%>
	</select>
	<!--<input class="style10" type="text" value="< %=Container%>" name="Container" id="Furgon TC">-->
	</td>	
  </tr>
  <%if ItineraryType=2 then%>
  <tr>
    <td class="style4" align="left">Unidad No:</td>
    <td class="style4" align="left">Tamaño:</td>
    <td class="style4" align="left">Marchamo:</td>
    <td class="style4" align="left">Seguro:</td>
  </tr>
  <tr>
    <td class="style4" align="left" bgcolor="#999999">
		<input class="style10" type="text" name="UnitNo" value="<%=UnitNo%>" >
	</td>	
    <td class="style4" align="left" bgcolor="#999999">
		<input class="style10" type="text" name="BLSize" value="<%=BLSize%>">
	</td>	
    <td class="style4" align="left" bgcolor="#999999">
		<input class="style10" type="text" name="Marchamo" id="Marchamo" value="<%=Marchamo%>">
		<input type="hidden" name="OldMarchamo" value="<%=Marchamo%>">
		<!--<a href="#" onClick="Javascript:document.forma.Marchamo.value='';return(false);" class="menu"><font color="FFFFFF">X</font></a>-->
	</td>	
    <td class="style4" align="left" bgcolor="#999999">
	<select class="style10" name="HasSecure" id="Seguro">
	<option value="-1">Seleccionar</option>
	<option value="1">SI</option>
	<option value="0">NO</option>
	</select>
	</td>
  </tr>
  <%end if%>
</table>
<table width="988" height="309" border="1" cellpadding="0" cellspacing="0" align="center">
	<tr>
		<td class="style4" align="center">No.<br>Bultos</td>
		<td class="style4" align="center">Clase Bultos</td>
		<td class="style4" align="center"><span class="style15">Descripci&oacute;n de Carga</span></td>
		<td class="style4" align="center">Volumen<br>(CBM) </td>
		<td class="style4" align="center"><span class="style15">Peso Bruto <br>(Kg)</span></td>
		<td class="style4" align="center">Consignee / <br>Consignatario</td>
		<td class="style4" align="center">BL/AWB/RO</td>
		<td class="style4" align="center">Fecha<br>Descarga</td>
		<td class="style4" align="center">Origen</td>		
		<td class="style4" align="center">Destino<br>Final</td>		
		<td class="style4" align="center">Shipper / Exportador</td>		
		<td class="style4" align="center">Valor $ C.A.</td>
		<td class="style4" align="center">Sep</td>		
	</tr>
	<tr bgcolor="#999999">
	  <td class="style4" valign="top">
		<textarea class="style10" cols="7" rows="20" wrap="off"  name="NoOfPieces" id="Numero de Bultos" onBlur="javascript:SumVals(this, document.forms[0].TotNoOfPieces);SetClassNoOfPieces(this);<%if Countries="GT" or Countries="SV" then%>SetDiceContenerValue(this);<%end if%>"><%=Join(NoOfPieces, ntr)%></textarea><br>
		<input class="style10" name="TotNoOfPieces" value="<%=TotNoOfPieces%>"  type="text" size="8" readonly>
	  </td>
	  <td class="style4" valign="top">
		<textarea class="style10" cols="9" rows="20" wrap="off"  name="ClassNoOfPieces" id="Clase de Bultos"><%=Join(ClassNoOfPieces, ntr)%></textarea><br>
	  </td>
	  <td class="style4" valign="top">
		<textarea class="style10" cols="41" rows="20" wrap="off" name="Contener" id="Descripción de Carga" readonly></textarea>
		<%if isnull(LtAcceptDate) and ItineraryType=1 then%>
		<table width="100%" cellpadding="0" cellspacing="0">
		<tr>
			<td class="style4" width="95%" valign="top">
			<INPUT type=button value="Borrar Linea #" onClick="JavaScript:EraseLines(document.forma.LineToErase.value-1);" class=label>
			<INPUT name="LineToErase" type=text value="" size=1 maxLength=2 class=label>			
			</td>
			<!--<td class="style4" align="right" valign="top">
				<a href="#" onClick="Javascript:IData(9);return (false);" class="submenu"><font color="FFFFFF">Nuevo</font></a>
			</td>-->
			<td class="style4">&nbsp;&nbsp;</td>
			<td class="style4" align="right" valign="top">
				<a href="#" onClick="Javascript:GetData(14);return (false);" class="menu"><font color="FFFFFF">Agregar</font></a>
			</td>
		</tr>
		</table>
		<%end if%>
	  </td>
		<td class="style4" valign="top">
	    <textarea class="style10"  cols="9" rows="20" wrap="off" name="Volumes" id="Volumen"  onBlur="javascript:SumVals(this, document.forms[0].TotVolume);"><%=Join(Volumes, ntr)%></textarea>
		<input class="style10"  type="text" size="10" name="TotVolume" value="<%=TotVolume%>" readonly>		
		</td>
		<td class="style4" valign="top">
		<textarea class="style10"  cols="9" rows="20" wrap="off" name="Weights" id="Peso Bruto" onBlur="javascript:SumVals(this, document.forms[0].TotWeight);"><%=Join(Weights, ntr)%></textarea><BR>
		<input class="style10"  type="text" size="10" name="TotWeight" value="<%=TotWeight%>" readonly>		
		</td>
		<td class="style4" valign="top">        
		  	<textarea class="style10"  cols="18" rows="20" wrap="off" name="Clients" id="Consignatarios (Consolidados)" readonly><%=Join(Clients, ntr)%></textarea>
			<div id="Consig" style="visibility:visible;">
			<%'if isnull(LtAcceptDate) then%>			
			<!--<table width="100%" cellpadding="0" cellspacing="0">
			<tr>
				<td class="style4" width="80%">&nbsp;</td>
				<td class="style4" align="right">
					<a href="#" onClick="Javascript:IData(11);return (false);" class="submenu"><font color="FFFFFF">Nuevo</font></a>
				</td>
				<td class="style4">&nbsp;&nbsp;</td>
				<td class="style4" align="right">
					<a href="#" onClick="Javascript:GetData(11);return (false);" class="menu"><font color="FFFFFF">Buscar</font></a>
				</td>
			</tr>
			</table>-->
			<%'end if%>
			</div>
		</td>
	    <td class="style4" valign="top">
		<textarea class="style10"  cols="13" rows="20" wrap="off" name="BLs" id="BLs o AWBs de Origen" readonly><%=Join(BLs, ntr)%></textarea>
		</td>
	    <td class="style4" valign="top">
		<textarea class="style10" cols="13" rows="20" wrap="off" name="DischargeDate" id="Fecha de Descarga" readonly><%=Join(DischargeDate, ntr)%></textarea><br>
			<%if isnull(LtAcceptDate) and ItineraryType=1 then%>			
			<table width="100%" cellpadding="0" cellspacing="0">
			<tr>
				<td class="style4" align="right">
					<a href="#" onClick="JavaScript:abrir('DischargeDate');return (false);" class="menu"><font color="FFFFFF">Seleccionar</font></a>
				</td>
			</tr>
			</table>
			<%end if%>
		</td>
		<td class="style4" valign="top">
		<textarea class="style10"  cols="4" rows="20" wrap="off" name="CountriesOrigen" id="Pais de Origen" readonly><%=Join(CountriesOrigen, ntr)%></textarea>
		</td>
		<td class="style4" valign="top">
		<textarea class="style10"  cols="4" rows="20" wrap="off" name="CountriesFinalDes" id="Pais de Destino" readonly><%=Join(CountriesFinalDes, ntr)%></textarea>
		</td>
		<td class="style4" valign="top">
		  <textarea class="style10"  cols="15" rows="20" wrap="off" name="Agents" id="Agentes" readonly><%=Join(Agents, ntr)%></textarea><br>
			<%'if isnull(LtAcceptDate) then%>			
			<!--<table width="100%" cellpadding="0" cellspacing="0">
			<tr>
				<td class="style4" width="80%">&nbsp;</td>
				<td class="style4" align="right">
					<a href="#" onClick="Javascript:IData(20);return (false);" class="submenu"><font color="FFFFFF">Nuevo</font></a>
				</td>
				<td class="style4">&nbsp;&nbsp;</td>
				<td class="style4" align="right">
					<a href="#" onClick="Javascript:GetData(20);return (false);" class="menu"><font color="FFFFFF">Buscar</font></a>
				</td>
			</tr>
			</table>-->
			<%'end if%>
		</td>
		<td class="style4" valign="top">
			<textarea class="style10" cols="7" rows="20" wrap="off"  name="DiceContenerValue" id="Valor $ C.A." onBlur="javascript:SumVals(this,document.forms[0].TotDiceContenerValue);"><%=Join(DiceContenerValue, ntr)%></textarea>
			<input class="style10"  type="text" size="8" name="TotDiceContenerValue" value="<%=TotDiceContenerValue%>" readonly>
	    </td>
		<td class="style4" valign="top">
			<textarea class="style10" cols="3" rows="20" wrap="off"  name="Seps" id="Separacion de Documentos" onBlur="javascript:SumVals(this,document.forms[0].SumSeps);"><%=Join(Seps, ntr)%></textarea>
			<input type="hidden" name="SumSeps">
	    </td>
	</tr>
</table>
<table width="1076" border="1" cellpadding="2" cellspacing="0" align="center">
  <tr>
    <td width="206" align="center" class="style4">PREPAGADO</td>
    <td width="206" align="center" class="style4">AL COBRO</td>
	<td class="style4" align="center"&nbsp;</td>
	<td align="center" class="style4">Observaciones para Carta Porte </td>
  </tr>
 <tr>
    <td class="style4" align="center" colspan="2">Flete</td>
	<td class="style4" align="center">Fecha Despacho</td>
	<td rowspan="2" align="center" class="style4" bgcolor="#999999">
	<textarea class="style10" cols="45" rows="2" name="Observations"><%=Observations%></textarea>
	</td>
 </tr>
  <tr>
    <td class="style4" align="center" bgcolor="#999999">
	<input type="text" class="style10" name="BLsFreight2" size="16" value="<%=BLsFreight2%>" onBlur="javascript:this.value=Round(this.value);" onKeyUp="res(this,numb);">
	<a href="#" onClick="Javascript:document.forma.BLsFreight2.value='AS AGREED';return(false);" class="menu"><font color="FFFFFF">AA</font></a>
	</td>
    <td class="style4" align="center" bgcolor="#999999">
	<input type="text" class="style10" name="BLsFreight" size="16" value="<%=BLsFreight%>" onBlur="javascript:this.value=Round(this.value);" onKeyUp="res(this,numb);">
	<a href="#" onClick="Javascript:document.forma.BLsFreight.value='AS AGREED';return(false);" class="menu"><font color="FFFFFF">AA</font></a>
	</td>
    <td class="style4" align="center" bgcolor="#999999">
	<INPUT type="text"  readonly name="BLDispatchDate" value="<%=BLDispatchDate%>" size=23 maxLength=19 class=label id="Fecha de Despacho">&nbsp;<a href="#" onClick="JavaScript:abrir('BLDispatchDate');return (false);" class="menu"><font color="FFFFFF">Seleccionar</font></a>
	</td>
  </tr>
 <tr>
    <td class="style4" align="center" colspan="2">Seguros</td>
	<td class="style4" align="center">ETD / Fecha Estimada de Salida</td>
	<td class="style4" align="center"><span class="style15">Encargado</span></td>
 </tr>
  <tr>
    <td class="style4" align="center" bgcolor="#999999">
	<input type="text" class="style10" name="BLsInsurance2" value="<%=BLsInsurance2%>" onBlur="javascript:this.value=Round(this.value);" onKeyUp="res(this,numb);">
	</td>
    <td class="style4" align="center" bgcolor="#999999">
	<input type="text" class="style10" name="BLsInsurance" value="<%=BLsInsurance%>" onBlur="javascript:this.value=Round(this.value);" onKeyUp="res(this,numb);">
	</td>
    <td class="style4" align="center" bgcolor="#999999">
    <table border=0 cellpadding=0 cellspacing=0>
        <tr>
        <td class="style4" align="center" bgcolor="#999999">
        <INPUT readonly name="BLExitDate" type=text value="<%=BLExitDate%>" size=23 maxLength=19 class=label id="Fecha Estimada de Salida">
        </td>
        <td class="style4" align="center" bgcolor="#999999">
        <div id="EstExitDate">
        &nbsp;<a href="#" onClick="JavaScript:abrir('BLExitDate');return (false);" class="menu"><font color="FFFFFF">Seleccionar</font></a>
        </div>
        </td>
        </tr>
    </table>
    </td>
    <td class="style4" align="center" bgcolor="#999999"><input type="text" class="style10" name="ContactSignature" value="<%=ContactSignature%>" maxlength="45" readonly></td>
	</tr>
   <tr>
    <td class="style4" align="center" colspan="2">Otros</td>
	<td class="style4" align="center"><span class="style15">ETA / Fecha Estimada de Llegada</span></td>
	<td class="style4" align="center">Lugar de Emisi&oacute;n</td>
 </tr>
  <tr>
    <td class="style4" align="center" bgcolor="#999999">
	<input type="text" class="style10" name="BLsAnotherChargesPrepaid" value="<%=BLsAnotherChargesPrepaid%>" onBlur="javascript:this.value=Round(this.value);" onKeyUp="res(this,numb);">
	</td>
    <td class="style4" align="center" bgcolor="#999999">
	<input type="text" class="style10" name="BLsAnotherChargesCollect" value="<%=BLsAnotherChargesCollect%>" onBlur="javascript:this.value=Round(this.value);" onKeyUp="res(this,numb);">
	</td>
    <td class="style4" align="center" bgcolor="#999999">
    <table border=0 cellpadding=0 cellspacing=0>
        <tr>
        <td class="style4" align="center" bgcolor="#999999">
        <input readonly name="BLEstArrivalDate" type=text value="<%=BLEstArrivalDate%>" size=23 maxLength=19 class=label id="Fecha Estimada de Llegada">
        </td>
        <td class="style4" align="center" bgcolor="#999999">
        <div id="EstArrivalDate">
        &nbsp;<a href="#" onClick="JavaScript:abrir('BLEstArrivalDate');return (false);" class="menu"><font color="FFFFFF">Seleccionar</font></a>
        </div>
        </td>
        </tr>
    </table>
    </td>
    <td class="style4" align="center" bgcolor="#999999">
	<select name="Countries" id="Pais" class="label">
		<option value="-1">Seleccionar</option>
		<%DisplayCountries Countries, 2%>
	</select></td>	
  <tr>
    <td class="style4" align="center">TOTAL PREPAGADO</td>
    <td class="style4" align="center">TOTAL AL COBRO</td>
	<td class="style4" align="center"><%if ItineraryType=1 then%>Fianza o Marchamo para Manifiesto:<%end if%>&nbsp;</td>
	<td class="style4" align="center"><%if ItineraryType=1 then%>Observaciones  para Manifiesto<%end if%>&nbsp;</td>
  </tr>
  <tr>
    <td class="style4" align="center" bgcolor="#999999">
	<input type="text" class="style10" name="TotPrepaid" value="<%=TotPrepaid%>" onBlur="javascript:this.value=Round(this.value);" onKeyUp="res(this,numb);" id="Total Prepagado">
	</td>
    <td class="style4" align="center" bgcolor="#999999">
	<input type="text" class="style10" name="TotCollect" value="<%=TotCollect%>" onBlur="javascript:this.value=Round(this.value);" onKeyUp="res(this,numb);" id="Total Al Cobro">
	</td>
    <td align="center" class="style4" bgcolor="#333333" valign="middle">
	<%if ItineraryType=1 then%>
	<INPUT name="Bail" type=text value="<%=Bail%>" size=23 maxLength=45 class=label>
	<%end if%>&nbsp;
	</td>
	<td align="center" class="style4" bgcolor="#333333">
	<%if ItineraryType=1 then%>
	<span class="style15">
	  <textarea class="style14" cols="45" rows="2" name="Comment2"><%=Comment2%></textarea>
	</span>
	<%end if%>&nbsp;
	</td>
  </tr>
  <%if ItineraryType=1 then%>
  <tr>
    <td class="style4" align="center">No.DTI</td>
	<td class="style4" colspan="2" align="center">Marcas de Expedicion, Nos. Contenedor, dimensiones para inciso 28 de DTI</td>
	<td class="style4" align="center">Observaciones para DTI</td>
  </tr>
  <tr bgcolor="#CCCCCC">
    <td class="style4" align="center"><INPUT name="DTI" id="DTI" type=text value="<%=DTI%>" size=23 maxLength=45 class=label></td>
	<td class="style4" colspan="2" align="center"><textarea class="style10" cols="45" rows="2" name="DTIObservations"><%=DTIObservations%></textarea></td>
	<td class="style4" align="center"><Textarea class="style10" name="Comment4" id="Comentario" cols="45" rows="2"><%=Comment4%></Textarea></td>
  </tr>
  <tr>
    <td class="style4" align="center" colspan="4">Instrucciones Precisas que debe de Efectuar el Transportista segun contratacion y servicio ofrecido</td>
    </tr>
    <tr bgcolor="#999999">
    <td class="style4" align="center" colspan="4"><textarea class="style10" cols="140" rows="5" name="PilotInstructions"><%=PilotInstructions%></textarea></td>
  </tr>
  <%end if%>

</table>
<table width="1076" border="1" cellpadding="2" cellspacing="0" align="center">
    <%if Expired = 0 then %>	    
        <INPUT name="Expired" type=hidden value="on">
    <%else %>
        <INPUT name="Expired" type=hidden value="">
    <%end if %>
    <%if CountTableValues = -1 then%>
	 <TD class=label align=center>
	 <!--<INPUT name=enviar type=button onClick="javascript:if(confirm('Si Actualiza y Cierra ya no podra hacer modificaciones y la informacion continuara su proceso')){document.forma.Closed.value=1;validar(1);};" value="Cerrar&nbsp;&nbsp;" class=label>
	 &nbsp;&nbsp;&nbsp;&nbsp;-->
	<INPUT name=enviar type=button onClick="JavaScript:validar(1)" value="&nbsp;&nbsp;Guardar&nbsp;&nbsp;" class=label></TD>
	<%else
	%>	 
	 <TD class=label align=center >
	 <INPUT type=button onClick="Javascript:window.open('BLPrintConditions.asp?BLID=<%=ObjectID%>&BTP=<%=BLType%>','BLPrint','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=800');return false;" value="&nbsp;CP&nbsp;Master&nbsp;" class=label>
	 <INPUT type=button onClick="Javascript:window.open('Reports.asp?GID=13&OID=<%=ObjectID%>','Manifest','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=800,height=600');return false;" value="&nbsp;Manifiesto&nbsp;Master&nbsp;" class=label>
	 <INPUT type=button onClick="Javascript:window.open('DTIPrint.asp?GID=25&OID=<%=ObjectID%>','DTI','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750,height=600');return false;" value="&nbsp;DTI&nbsp;Master&nbsp;" class=label><br>
	 <!--<INPUT type=button onClick="Javascript:window.open('MultipleDocs.asp?GID=<%=15-BLType%>&BLID=<%=ObjectID%>&BTP=5&Typ=0','Insurances','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750,height=800');return false;" value="Cartas&nbsp;Seguro" class=label>-->
	 <%select Case BLType
     case 0,1
     %>
     <INPUT type=button onClick="Javascript:window.open('MultipleDocs.asp?GID=<%=15-BLType%>&BLID=<%=ObjectID%>&BTP=<%=BLType%>&Typ=0','CPIndividual','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=850,height=800');return false;" value="CPs&nbsp;Individuales" class=label>
	 <INPUT type=button onClick="Javascript:window.open('MultipleDocs.asp?GID=<%=15-BLType%>&BLID=<%=ObjectID%>&BTP=<%=BLType%>&Typ=2','CPPrealerts','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750,height=800');return false;" value="PreAlertas&nbsp;Individuales" class=label><br>
	 <%end select %>

     <INPUT type=button onClick="Javascript:window.open('MultipleDocs.asp?GID=<%=15-BLType%>&BLID=<%=ObjectID%>&BTP=<%=BLType%>&Typ=1','CPManifests','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=800,height=800');return false;" value="Manifiestos&nbsp;Individuales" class=label>
	 
     <%select Case BLType
     case 0,1
     %>
     <INPUT type=button onClick="Javascript:window.open('InstructionsPrint.asp?BLID=<%=ObjectID%>&BTP=<%=BLType%>','InstrucPrint','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750');return false;" value="Carta&nbsp;Instrucciones" class=label>
	 <%end select %>	
       <%if Closed=0 then%>
		 	<TD class=label align=center><INPUT name=enviar type=button onClick="javascript:if(confirm('Si Actualiza y Cierra ya no podra hacer modificaciones y la informacion continuara su proceso')){document.forma.Closed.value=1;validar(2);};" value="&nbsp;Cerrar&nbsp;" class=label></TD>
			<TD class=label align=center><INPUT name=enviar type=button onClick="JavaScript:validar(2);" value="&nbsp;Actualizar&nbsp;" class=label></TD>
			<%if ObjectID<>19384 and BLArrivalDate = "" then %>
            <TD class=label align=center><INPUT name=enviar type=button onClick="JavaScript:if(confirm('¿Está seguro de querer eliminar la CP máster?')){validar(3);};" value="&nbsp;Eliminar&nbsp;" class=label></TD>
            <%end if %>
	   <%else%>
			<%if Session("OperatorLevel") = 0 then%>
			<TD class=label align=center><INPUT name=enviar type=button onClick="Javascript:document.forma.Closed.value=0;validar(2);" value="&nbsp;&nbsp;Abrir&nbsp;&nbsp;" class=label></TD>
			<%end if%>
	   <%end if%>
	<%end if%>
</table>
</form>
<script>
selecciona('forma.BLType','<%=BLType%>');
selecciona('forma.CountryDes','<%=CountryDes%>');
selecciona('forma.CountryDep','<%=CountryDep%>');
selecciona('forma.BrokerID','<%=BrokerID%>');
selecciona('forma.Container','<%=Container%>');
selecciona('forma.PilotID','<%=PilotID%>');
selecciona('forma.TruckID','<%=TruckID%>');
selecciona('forma.HasSecure','<%=HasSecure%>');
<%select case BLType
case 2,3%>
selecciona('forma.DeliveryPolicyHour','<%=DeliveryPolicyHour%>');
selecciona('forma.DeliveryPolicyMin','<%=DeliveryPolicyMin%>');
<%end select%>
CleanSpaces(document.forma.Clients,"\r\n", "\n");
CleanSpaces(document.forma.BLs,"\r\n", "\n");
CleanSpaces(document.forma.DischargeDate,"\r\n", "\n");
CleanSpaces(document.forma.CountriesOrigen,"\r\n", "\n");
CleanSpaces(document.forma.CountriesFinalDes,"\r\n", "\n");
SetLics(document.forma.PilotID.selectedIndex);
EnumDiceContener("", "");
selecciona('forma.ChargeType','<%=ChargeType%>');
selecciona('forma.DestinyType','<%=DestinyType%>');
<%if ChargeType=0 then%>
	document.getElementById('Lugar de Carga').readOnly = "";
<%end if%>
<%if DestinyType=0 then%>
	document.getElementById('Destino Final / Entrega').readOnly = "";
<%end if%>
SetConsolidated();
<%'if BLType=2 then%>
//if ((document.forma.ChargeType.value>0)||(document.forma.DestinyType.value>0)) {
//	if (document.forma.Marchamo.value==0) {
//		alert("Aviso: no se asigno Marchamo, ya no hay disponibles para la bodega seleccionada, favor de agregar nuevos marchamos");
//	}
//}
<%'end if%>
</script>
<%
Set aList1Values = Nothing
Set aList2Values = Nothing
Set aList3Values = Nothing
Set aList4Values = Nothing
Set aList5Values = Nothing
Set aList6Values = Nothing
%>
</body>
</html>