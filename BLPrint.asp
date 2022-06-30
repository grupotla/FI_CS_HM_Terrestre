<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
'Checking "0|1|2"

On Error Resume Next 

Dim Conn, Conn2, rs, BLID, aTableValues, CountTableValues, aTableValues2, CountTableValues2, QuerySelect, QuerySelect2, QuerySelect3, IsColoader
Dim SenderData, ShipperData, ConsignerData, CountryDep, Attn, HandlingInformation, CountryDes, GroupID, ShowConsigners, ntr, AsAgreed, TransitConsignerData
Dim ChargePlace, FinalDes, Container, TotNoOfPieces, BrokerName, PilotName, License, TruckNo, BLNumber, i, Countries, Chassis, PO, Nit1, Nit2
Dim TotWeight, TotVolume, TotPrepaid, TotCollect, Freight, Insurance, AnotherChargesCollect, Observations, CountriesFinalDes, Bill, AgentsID
Dim ContactSignature, BLDispatchDate, ContainerDep, SBLID, Consolidated, BLType, HasSecure, UnitNo, BLSize, Marchamo, Phone1, PolicyNo, ShowProcedencia
Dim Freight2, Insurance2, AnotherChargesPrepaid, GroupValues, Pos, ClientID, AgentID, BLTransit, Sep, ShipperID, MBLs, PickupData, DeliveryData 
Dim NoOfPieces(), DiceContener(), Volumes(), Weights(), Clients(), ClassNoOfPieces(), CountryFinalDes(), CountryOrigen(), BLs()
Dim ColoaderData, ColoaderID, CompanyName, EXDBCountry, GuiaRemision, GuiaRemisionDet, NumeroCerosGuiaRemision, CountGuiaRemision
Dim FreightColoader, FreightColoader2, InsuranceColoader, InsuranceColoader2, AnotherChargesColoader, AnotherChargesColoader2
Dim FinalDesMaster, NotifyPartyID, NotifyParty, Footer, CodProv

	GroupID = CheckNum(Request("GID"))
	BLID = CheckNum(Request("BLID"))
	SBLID = CheckNum(Request("SBLID"))
	BLType = CheckNum(Request("BTP"))
	ClientID = CheckNum(Request("CID"))
	AgentID = CheckNum(Request("AID"))
	Sep = CheckNum(Request("SEP"))
	CountTableValues = -1
	Pos = ""
	CodProv = ""


if Session("OperatorID") = 1237 then
    response.write "Entro BLPrint<br>"
end if

	'response.write GroupID & "<br>"
	select case GroupID
	case 0 '0=Master
		select Case BLType
		case 0,1 '0=Consolidado, 1=Express
			QuerySelect = "SELECT a.BLID, a.BLNumber, a.SenderData, a.ShipperData, a.CountryDep, a.HandlingInformation, " & _
					  "a.CountryDes, a.ChargePlace, a.FinalDes, e.TruckNo, a.Observations, a.Consolidated, " & _
					  "a.ContactSignature, a.BLDispatchDate, a.Countries, b.Name, c.Name, c.License, d.TruckNo, a.BLType, a.Chassis, a.ContainerDep, " & _
					  "a.TotNoOfPieces, a.TotWeight, a.TotVolume, a.TotPrepaid, a.TotCollect, a.Freight, a.Insurance, a.AnotherChargesCollect, " & _
					  "a.ConsignerData, a.Attn, a.HasSecure, a.UnitNo, a.BLSize, a.Marchamo, c.Phone1, a.Freight2, a.Insurance2, " & _
					  "a.AnotherChargesPrepaid, a.ConsignerID, f.AgentsID, f.Seps, f.HBLNumber, f.MBLs, f.PickupData, f.DeliveryData, f.PolicyNo, a.GuiaRemision, " & _
                      "f.NotifyPartyID, f.NotifyParty, CONCAT(CASE WHEN IFNULL(g.CodProv,'') = '' THEN '' ELSE CONCAT(g.CodProv,' - ') END, g.Name) as CodProv " &_
					  "from BLs a left outer join Trucks e on a.Container = e.TruckID " & _
					  "inner join Brokers b on a.BrokerID = b.BrokerID inner join Pilots c on a.PilotID = c.PilotID " & _
					  "inner join Trucks d on a.TruckID = d.TruckID inner join BLDetail f on a.BLID=f.BLID " & _

                      "INNER JOIN Providers g ON g.ProviderID=d.ProviderID " & _		

					  "WHERE a.BLID = "
			if GroupID<>14 then
				QuerySelect2 = "select a.NoOfPieces, a.DiceContener, a.Volumes, a.Weights, a.Clients, a.Freight, a.Freight2, a.ClassNoOfPieces, a.Insurance, a.Insurance2, a.AnotherChargesCollect, a.AnotherChargesPrepaid, a.CountriesFinalDes, a.CountryOrigen, a.HBLNumber from BLDetail a where a.BLID = " & BLID & " order by a.Pos"
			else
				QuerySelect2 = "select a.NoOfPieces, a.DiceContener, a.Volumes, a.Weights, a.Clients, a.Freight, a.Freight2, a.ClassNoOfPieces, a.Insurance, a.Insurance2, a.AnotherChargesCollect, a.AnotherChargesPrepaid, a.CountriesFinalDes, a.CountryOrigen, a.HBLNumber from BLDetail a where a.BLID = " & BLID & " and a.ClientsID=" & ClientID & " and a.AgentsID=" & AgentID & " and a.Seps=" & Sep & " order by a.Pos"
			end if
		case 2,3 'Master Recoleccion o Entrega
			QuerySelect = "select a.BLID, a.BLNumber, a.SenderData, a.ShipperData, a.CountryDep, a.HandlingInformation, " & _
					  "a.CountryDes, a.ChargePlace, a.FinalDes, e.TruckNo, a.Observations, a.Consolidated, " & _
					  "a.ContactSignature, a.BLDispatchDate, a.Countries, c.Name, c.Name, c.License, d.TruckNo, a.BLType, a.Chassis, a.ContainerDep, " & _
					  "a.TotNoOfPieces, a.TotWeight, a.TotVolume, a.TotPrepaid, a.TotCollect, a.Freight, a.Insurance, a.AnotherChargesCollect, " & _
					  "a.ConsignerData, a.Attn, a.HasSecure, a.UnitNo, a.BLSize, a.Marchamo, c.Phone1, a.Freight2, a.Insurance2, " & _
					  "a.AnotherChargesPrepaid, a.ConsignerID, f.AgentsID, f.Seps, f.HBLNumber, f.MBLs, a.ChargePlace, a.FinalDes, f.PolicyNo, a.GuiaRemision, f.GuiaRemision, " & _
                      "f.NotifyPartyID, f.NotifyParty, f.EXDBCountry, CONCAT(CASE WHEN IFNULL(g.CodProv,'') = '' THEN '' ELSE CONCAT(g.CodProv,' - ') END, g.Name) as CodProv " &_
					  "from BLs a left outer join Trucks e on a.Container = e.TruckID " & _
					  "inner join Pilots c on a.PilotID = c.PilotID " & _
					  "inner join Trucks d on a.TruckID = d.TruckID inner join BLDetail f on a.BLID=f.BLID " & _

                      "INNER JOIN Providers g ON g.ProviderID=d.ProviderID " & _		

					  "where a.BLID = "
			QuerySelect2 = "select a.NoOfPieces, a.DiceContener, a.Volumes, a.Weights, a.Clients, a.Freight, a.Freight2, a.ClassNoOfPieces, a.Insurance, a.Insurance2, a.AnotherChargesCollect, a.AnotherChargesPrepaid, a.CountriesFinalDes, a.CountryOrigen, a.HBLNumber from BLDetail a where a.BLID = " & BLID & " order by a.Pos"
		case 4 'Master Grupo
			QuerySelect = "select a.BLID, e.BLNumber, a.SenderData, a.ShipperData, e.Countries, a.HandlingInformation, " & _
					  "e.CountryDes, a.ChargePlace, a.FinalDes, g.TruckNo, a.Observations, a.Consolidated, " & _
					  "a.ContactSignature, a.BLDispatchDate, a.Countries, b.Name, c.Name, c.License, d.TruckNo, a.BLType, a.Chassis, a.ContainerDep, " & _
					  "a.TotNoOfPieces, a.TotWeight, a.TotVolume, 0, 0, 0, 0, 0, " & _
					  "a.ConsignerData, a.Attn, a.HasSecure, a.UnitNo, a.BLSize, a.Marchamo, c.Phone1, 0, 0, 0, a.ConsignerID, a.GuiaRemision, CONCAT(CASE WHEN IFNULL(g.CodProv,'') = '' THEN '' ELSE CONCAT(g.CodProv,' - ') END, g.Name) as CodProv " & _
					  "from BLs a left outer join Trucks g on a.Container = g.TruckID " & _
					  "inner join Brokers b on a.BrokerID = b.BrokerID inner join Pilots c on a.PilotID = c.PilotID " & _
					  "inner join Trucks d on a.TruckID = d.TruckID inner join BLGroupDetail f on a.BLID=f.BLID " & _
					  "inner join BLGroups e on e.BLGroupID = f.BLGroupID " & _

                      "INNER JOIN Providers g ON g.ProviderID=d.ProviderID " & _		

					  "where f.BLGroupID = "
			QuerySelect2 = "select a.NoOfPieces, a.DiceContener, a.Volumes, a.Weights, a.Clients, 0, 0, a.ClassNoOfPieces, 0, 0, 0, 0, a.CountriesFinalDes, a.CountryOrigen, a.HBLNumber from BLDetail a, BLs b, BLGroups e, BLGroupDetail f where a.BLID=b.BLID and b.BLID=f.BLID and e.BLGroupID=f.BLGroupID and e.BLGroupID=" & BLID & " order by b.BLNumber, a.Pos"
		'case 5 'Carta Seguro, Similar al query cuando es GroupID=15, solo que busca por BLID en lugar de BLDetailID
		'QuerySelect = "select a.BLID, a.BLNumber, a.SenderData, a.ShipperData, a.CountryDep, a.HandlingInformation, " & _
		'		  "a.CountryDes, a.ChargePlace, a.FinalDes, e.TruckNo, a.Observations, a.Consolidated, " & _
		'		  "a.ContactSignature, a.BLDispatchDate, a.Countries, b.Name, c.Name, c.License, d.TruckNo, a.BLType, a.Chassis, a.ContainerDep, " & _
		'		  "f.ClientsID, f.AddressesID, f.Pos, f.CountriesFinalDes, f.AgentsID, f.AgentsAddrID, f.Seps, f.HBLNumber " & _
		'		  "from (((((BLs a left outer join Trucks e on a.Container = e.TruckID) " & _
		'		  "inner join Brokers b on a.BrokerID = b.BrokerID) inner join Pilots c on a.PilotID = c.PilotID) " & _
		'		  "inner join Trucks d on a.TruckID = d.TruckID) inner join BLDetail f on a.BLID=f.BLID) " & _
		'		  "where f.BLID = "
		'QuerySelect2 = "select a.NoOfPieces, a.DiceContener, a.Volumes, a.Weights, a.Clients, a.Freight, a.Freight2, a.ClassNoOfPieces, a.Insurance, a.Insurance2, a.AnotherChargesCollect, a.AnotherChargesPrepaid from BLDetail a where a.BLID = " & BLID & " and a.ClientsID=" & ClientID & " and a.AgentsID=" & AgentID & " and a.Seps=" & Sep & " order by a.Pos"
		end select
	case 14, 15 '14=Cliente Express 15=Cliente Consolidado
		QuerySelect = "select a.BLID, a.BLNumber, a.SenderData, f.Shippers, f.CountryOrigen, a.HandlingInformation, " & _
				  "f.CountriesFinalDes, a.ChargePlace, a.FinalDes, e.TruckNo, f.Observations, a.Consolidated, " & _
				  "a.ContactSignature, a.BLDispatchDate, f.Countries, b.Name, c.Name, c.License, d.TruckNo, a.BLType, a.Chassis, a.ContainerDep, " & _
				  "f.ClientsID, f.AddressesID, f.Pos, f.CountriesFinalDes, f.AgentsID, f.AgentsAddrID, f.Seps, f.HBLNumber, f.Notify, f.AsAgreed, " & _
                  "f.BLs, f.PO, f.Bill, f.ShippersID, a.CountryDes, a.ConsignerID, a.ShipperData, f.EXType, f.MBLs, " & _
				  "f.ShipperColoader, f.ColoadersID, f.ColoadersAddrID, f.Coloaders, c.Phone1, EXDBCountry, a.GuiaRemision, f.GuiaRemision, " & _
                  "f.NotifyPartyID, f.NotifyParty, CONCAT(CASE WHEN IFNULL(g.CodProv,'') = '' THEN '' ELSE CONCAT(g.CodProv,' - ') END, g.Name) as CodProv " &_
                    "FROM BLDetail f " & _
                    "left join BLs a on a.BLID=f.BLID " & _
                    "left join Trucks e on a.Container = e.TruckID " & _
                    "left join Brokers b on a.BrokerID = b.BrokerID " & _
                    "left join Pilots c on a.PilotID = c.PilotID " & _
                    "left join Trucks d on a.TruckID = d.TruckID " & _

                "left join Providers g ON g.ProviderID=d.ProviderID " & _		

				  "where f.BLDetailID = "
                  '2021-10-13 segun reunion con Emanuel y Cesar, se decidio cambiar el Countries a.Countries (master) por f.Countries (hija)
                  '"from (((((BLs a left outer join Trucks e on a.Container = e.TruckID) " & _
				  '"inner join Brokers b on a.BrokerID = b.BrokerID) inner join Pilots c on a.PilotID = c.PilotID) " & _
				  '"inner join Trucks d on a.TruckID = d.TruckID) inner join BLDetail f on a.BLID=f.BLID) " & _

		'if GroupID<>15 then
		'	QuerySelect2 = "select a.NoOfPieces, a.DiceContener, a.Volumes, a.Weights, a.Clients, a.Freight, a.Freight2, a.ClassNoOfPieces, a.Insurance, a.Insurance2, a.AnotherChargesCollect, a.AnotherChargesPrepaid from BLDetail a where a.BLDetailID = " & SBLID & " order by a.Pos"
		'else
		QuerySelect2 = "select a.NoOfPieces, a.DiceContener, a.Volumes, a.Weights, a.Clients, a.Freight, a.Freight2, a.ClassNoOfPieces, a.Insurance, a.Insurance2, a.AnotherChargesCollect, a.AnotherChargesPrepaid, a.CountriesFinalDes, a.CountryOrigen, a.HBLNumber, a.FreightColoader, a.FreightColoader2, a.InsuranceColoader, a.InsuranceColoader2, a.AnotherChargesColoader, a.AnotherChargesColoader2 from BLDetail a where a.BLID = " & BLID & " and a.ClientsID=" & ClientID & " and a.AgentsID=" & AgentID & " and a.Seps=" & Sep & " order by a.Pos"
	end select
	
	OpenConn Conn
	
	'if BLType=5 then 'Carta Seguro
	'	BLTransit = 0
	'	'Revisando si el BLID es transito u origen
	'	Set rs = Conn.Execute("select BLIDTransit from BLDetail where BLDetailID=" & SBLID)
	'		if Not rs.EOF then
	'			BLTransit = rs(0)
	'		end if
	'	CloseOBJ rs
	'	
	'	'Buscando el Origen si el BLID es Transito
	'	Do while BLTransit<>0
	'		Set rs = Conn.Execute("select BLID, BLDetailID, BLIDTransit from BLDetail where BLDetailID=" & BLTransit)
	'			BLID = rs(0)
	'			SBLID = rs(1)
	'			BLTransit = rs(2)
	'		CloseOBJ rs
	'	loop
	'end if

	'Seteando el Query con el BLID Origen
	if GroupID=14 or GroupID=15 then
		QuerySelect = QuerySelect & SBLID
	else
		QuerySelect = QuerySelect & BLID
	end if
	

    'response.write GroupID & " Group ID<br>"

    QuerySelect2 = Replace(QuerySelect2,"order by a.Pos","") 

    '2021-10-05 hoy se comento todo este codigo que da error ************************************
    'QuerySelect2 = QuerySelect2 & " AND (SELECT count(*) FROM BLDetail2 WHERE BLID = " & SBLID & " AND Expired = 0) = 0 " & _
'"UNION " & _
'"select a.NoOfPieces, a.DiceContener, a.Volumes, a.Weights, a.Clients, a.Freight, a.Freight2, a.ClassNoOfPieces, a.Insurance, a.Insurance2, a.AnotherChargesCollect, a.AnotherChargesPrepaid, a.CountriesFinalDes, a.CountryOrigen, a.HBLNumber"

    'if GroupID=14 or GroupID=15 then 
    '    QuerySelect2 = QuerySelect2 & ", a.FreightColoader, a.FreightColoader2, a.InsuranceColoader, a.InsuranceColoader2, a.AnotherChargesColoader, a.AnotherChargesColoader2" 
    'end if

    'QuerySelect2 = QuerySelect2 & " FROM BLDetail2 a where a.BLID = " & SBLID & " AND a.Expired = 0 " 'ORDER BY a.Pos, BLDetailID" 
    'hasta aqui se comento hoy *********************************************
    
	'response.Write QuerySelect2 & "<br>"



if Session("OperatorID") = 1237 then
	response.Write GroupID & "<br>" & QuerySelect & "<br><br>"
end if

	Set rs = Conn.Execute(QuerySelect)
	If Not rs.EOF Then
    	
        On Error Resume Next 
            CodProv = IFNULL(rs("CodProv"))

        If Err.Number<>0 then
            Err.Number = 0
            CodProv = " -- "
        end if


		aTableValues = rs.GetRows
		CountTableValues = rs.RecordCount-1
	End If
	CloseOBJ rs

	Set rs = Conn.Execute(QuerySelect2)
	If Not rs.EOF Then
		aTableValues2 = rs.GetRows
		CountTableValues2 = rs.RecordCount-1
	End If
	closeOBJs rs, Conn

	if CountTableValues >= 0 then
		ntr = chr(13) & chr(10)
		BLID = aTableValues(0,0)
		
		'if BLNumber = "" then
		'	BLNumber = aTableValues(1,0)
		'end if
		'if GroupID = 14 and BLType<>5 then
			'BLNumber = BLNumber & "-" & FiveDigits(aTableValues(40,0)+aTableValues(41,0)+aTableValues(42,0))
		'end if

		select case GroupID
		case 0
			BLNumber = aTableValues(1,0)

		'case 14
			'select case BLType
			'case 0,1,2,3
			'	BLNumber = aTableValues(43,0)
			'case 5
			'	BLNumber = aTableValues(29,0)
			'end select
		case else

            'response.write "(" & aTableValues(40,0) & ")(" & aTableValues(32,0) & ")(" & aTableValues(29,0) & ")<br>(" & PtrnViewBLAgents & ")(" & aTableValues(35,0) & ")"

           'if InStr(1,ValidPatrnCIF,aTableValues(14,0))>0 then
            'if InStr(1,ValidPatrnCIF,Mid(aTableValues(29,0),2,2))>0 then
            if FRegExp(PtrnViewBLAgents, aTableValues(35,0),  "", 2) then
                Select case CheckNum(aTableValues(39,0))
                Case 4,5 '0=FCL,1=LCL-SinDivision,2=LCL-ConDivision,4=RO-Consolidado,5=RO-Express,6=RO-Recoleccion,7=RO-Entrega,8=CIF-Externo(Mexico)
                    BLNumber = aTableValues(40,0)
                case else
                    BLNumber = aTableValues(32,0)
                end Select
            else
                BLNumber = aTableValues(29,0)
            end if
            PO = aTableValues(33,0)
            Bill = aTableValues(34,0)
            AgentsID = aTableValues(35,0)
		end select

		ShipperData =FRegExp(ntr, aTableValues(3,0), "<br>", 4)
		CountryDep = aTableValues(4,0)
		HandlingInformation = FRegExp(ntr, aTableValues(5,0), "<br>", 4)
		CountryDes = aTableValues(6,0)
		ChargePlace = FRegExp(ntr, aTableValues(7,0), "<br>", 4)
		FinalDes = FRegExp(ntr, aTableValues(8,0), "<br>", 4)
		Container = aTableValues(9,0)
		Observations = FRegExp(ntr, aTableValues(10,0), "<br>", 4)
		Observations = aTableValues(10,0)
		Consolidated = aTableValues(11,0)
		ContactSignature = aTableValues(12,0)
		BLDispatchDate = aTableValues(13,0)
		Countries = aTableValues(14,0)
		BrokerName = aTableValues(15,0)
		PilotName = aTableValues(16,0)
		License = aTableValues(17,0)
		TruckNo = aTableValues(18,0)
        'GuiaRemision = aTableValues(48,0)
		if BLType<>4 and BLType<>5 then
			BLType = CheckNum(aTableValues(19,0))
		end if
		Chassis = aTableValues(20,0)
		ContainerDep = aTableValues(21,0)




		'response.write(GroupID & " - GroupID")
        'response.write(BLType & " - BLType")
        TotNoOfPieces = 0
		if GroupID=0 then 'Master o Cliente Express
			if BLType <> 4 then
				TotNoOfPieces = aTableValues(22,0)
				TotWeight = aTableValues(23,0)
				TotVolume = aTableValues(24,0)
				TotPrepaid = aTableValues(25,0)
				TotCollect = aTableValues(26,0)
				Freight = aTableValues(27,0)
				Insurance = aTableValues(28,0)
				AnotherChargesCollect = aTableValues(29,0)
				Freight2 = aTableValues(37,0)
				Insurance2 = aTableValues(38,0)
				AnotherChargesPrepaid = aTableValues(39,0)
			else
				for i=0 to CountTableValues
					TotNoOfPieces = TotNoOfPieces + Cdbl(aTableValues(22,i))
					TotWeight = TotWeight + Cdbl(aTableValues(23,i))
					TotVolume = TotVolume + Cdbl(aTableValues(24,i))
					TotPrepaid = TotPrepaid + Cdbl(aTableValues(25,i))
					TotCollect = TotCollect + Cdbl(aTableValues(26,i))
					Freight = Freight + Cdbl(aTableValues(27,i))
					Insurance = Insurance + Cdbl(aTableValues(28,i))
					AnotherChargesCollect = AnotherChargesCollect + Cdbl(aTableValues(29,i))
					Freight2 = Freight2 + Cdbl(aTableValues(37,i))
					Insurance2 = Insurance2 + Cdbl(aTableValues(38,i))
					AnotherChargesPrepaid = AnotherChargesPrepaid + Cdbl(aTableValues(39,i))
				next
			end if
			'ShipperData =FRegExp(ntr, aTableValues(3,0), "<br>", 4)
			SenderData = FRegExp(ntr, aTableValues(2,0), "<br>", 4)
			ConsignerData = FRegExp(ntr, aTableValues(30,0), "<br>", 4)
			Attn = FRegExp(ntr, aTableValues(31,0), "<br>", 4)
			HasSecure = aTableValues(32,0)
			UnitNo = aTableValues(33,0)
			BLSize = aTableValues(34,0)
			Marchamo = aTableValues(35,0)
			Phone1 = aTableValues(36,0)
		else
			TotPrepaid = ""
			TotCollect = ""
			Freight = ""
			Insurance = ""
			AnotherChargesCollect = ""
			Freight2 = ""
			Insurance2 = ""
			AnotherChargesPrepaid = ""
			Pos = aTableValues(24,0) + 1
			CountriesFinalDes = aTableValues(25,0)
			Attn = FRegExp(ntr, aTableValues(30,0), "<br>", 4)
            AsAgreed = aTableValues(31,0)
            Phone1 = aTableValues(45,0)
            EXDBCountry = aTableValues(46,0)

			OpenConn2 Conn
			'Obteniendo datos del cliente desde la tabla master
			if aTableValues(22,0) <> 0 then
				QuerySelect3 = "select a.nombre_cliente, d.direccion_completa, d.phone_number, codigo_tributario, codigo_tributario2 " & _
										"from clientes a, direcciones d, niveles_geograficos n, paises p " & _
										" where a.id_cliente = d.id_cliente" & _
										" and d.id_nivel_geografico = n.id_nivel" & _
										" and n.id_pais = p.codigo" & _
										" and a.id_cliente = " & aTableValues(22,0)
				aTableValues(23,0) = CheckNum(aTableValues(23,0))
                'response.write aTableValues(22,0)
				if aTableValues(23,0) <> 0 then
					QuerySelect3 = QuerySelect3 & " and d.id_direccion = " & aTableValues(23,0)
				end if
				
				'response.write QuerySelect3 & "<br>"
				Set rs = Conn.Execute(QuerySelect3)
				if Not rs.EOF then
					'Phone1 = rs(2)
                    'response.write ConsignerData
					ConsignerData = ConsignerData & "<br>" & rs(1) & "<br>"
                    'response.write ConsignerData
                    if rs(2) <> "" then
						ConsignerData = ConsignerData & rs(2)
					end if
                    Nit1 = rs(3)
                    Nit2 = rs(4)
				end if
				CloseOBJ rs
			
				set rs = Conn.Execute("select numero_telefono from cli_telefonos where id_cliente=" & aTableValues(22,0))
				if Not rs.EOF then
					ConsignerData = ConsignerData & "    " & rs(0)
				end if
				CloseOBJ rs
				set rs = Conn.Execute("select nombres from contactos where id_cliente=" & aTableValues(22,0) & " and activo = true")
				if Not rs.EOF then
					ConsignerData = ConsignerData & "<br>ATTN:" & rs(0)
				end if
				CloseOBJ rs		
			end if

            QuerySelect3 = ""
            'Cuando la carga pasa por CR en Transito, en lugar de mostrar los nit del cliente, se deben mostrar los de Aimar (cliente en Master)
            'if (GroupID=14 or GroupID=15) and aTableValues(25,0)<>"CR" and aTableValues(36,0)="CR" then
                'QuerySelect3 = "select a.nombre_cliente, codigo_tributario, codigo_tributario2 " & _
										'"from clientes a, direcciones d, niveles_geograficos n, paises p " & _
										'" where a.id_cliente = d.id_cliente" & _
										'" and d.id_nivel_geografico = n.id_nivel" & _
										'" and n.id_pais = p.codigo" & _
										'" and a.id_cliente = " & aTableValues(37,0)
                'response.write aTableValues(37,0)
                'Set rs = Conn.Execute(QuerySelect3)
				'if Not rs.EOF then
					'TransitConsignerData = rs(0) & "/"
					'Nit1 = rs(1)
                    'Nit2 = rs(2)
				'end if
				'CloseOBJ rs
            'end if

            if (GroupID=14 or GroupID=15) and aTableValues(36,0)="PA" then
                ShipperData = aTableValues(38,0)
            end if

            'Se colocan los Nit del Cliente o de Aimar si pasa en Transito por CR
            if Nit1 <> "" then
                ConsignerData = ConsignerData & "<br>" & Nit1
            end if

            if Nit2 <> "" then
                ConsignerData = ConsignerData & "<br>" & Nit2
            end if

			QuerySelect3 = ""
			'Obteniendo datos del Shipper desde la tabla master
			ShipperID = CheckNum(aTableValues(26,0))
			if ShipperID <> 0 then
				QuerySelect3 = "select a.nombre_cliente, d.direccion_completa, d.phone_number " & _
										"from clientes a, direcciones d, niveles_geograficos n, paises p " & _
										" where a.id_cliente = d.id_cliente" & _
										" and d.id_nivel_geografico = n.id_nivel" & _
										" and n.id_pais = p.codigo" & _
										" and a.id_cliente = " & ShipperID
				aTableValues(27,0) = CheckNum(aTableValues(27,0))
				if aTableValues(27,0) <> 0 then
					QuerySelect3 = QuerySelect3 & " and d.id_direccion = " & aTableValues(27,0)
				end if										
				
				'response.Write QuerySelect3 & "<br>"
				Set rs = Conn.Execute(QuerySelect3)
				if Not rs.EOF then
					SenderData = rs(0) & "<br>" & rs(1) & "<br>"
					if rs(2) <> "" then
					SenderData = SenderData & rs(2)
					end if
				end if
				CloseOBJ rs

				set rs = Conn.Execute("select numero_telefono from cli_telefonos where id_cliente=" & ShipperID)
				if Not rs.EOF then
					SenderData = SenderData & "    " & rs(0)
				end if
				CloseOBJ rs
				set rs = Conn.Execute("select nombres from contactos where id_cliente=" & ShipperID & " and activo = true")
				if Not rs.EOF then
					SenderData = SenderData & "<br>ATTN:" & rs(0)
				end if
				CloseOBJ rs
			end if

            'Obteniendo datos del Coloader desde la tabla master
            QuerySelect3 = ""
			'Obteniendo datos del Shipper del Coloader desde la tabla master
			ColoaderID = CheckNum(aTableValues(42,0))
            if ColoaderID <> 0 then
				IsColoader = 1 'Se Setea en 1 para mostrar datos del Coloader en lugar del Agente en la casilla "Embarcador"
                QuerySelect3 = "select a.nombre_cliente, d.direccion_completa, d.phone_number " & _
										"from clientes a, direcciones d, niveles_geograficos n, paises p " & _
										" where a.id_cliente = d.id_cliente" & _
										" and d.id_nivel_geografico = n.id_nivel" & _
										" and n.id_pais = p.codigo" & _
										" and a.id_cliente = " & ColoaderID
				aTableValues(43,0) = CheckNum(aTableValues(43,0))
				if aTableValues(43,0) <> 0 then
					QuerySelect3 = QuerySelect3 & " and d.id_direccion = " & aTableValues(43,0)
				end if										
				
				'response.Write QuerySelect3 & "<br>"
				Set rs = Conn.Execute(QuerySelect3)
				if Not rs.EOF then
					ColoaderData = rs(0) & "<br>" & rs(1) & "<br>"
					if rs(2) <> "" then
					ColoaderData = ColoaderData & rs(2)
					end if
				end if
				CloseOBJ rs

				set rs = Conn.Execute("select numero_telefono from cli_telefonos where id_cliente=" & ColoaderID)
				if Not rs.EOF then
					ColoaderData = ColoaderData & "    " & rs(0)
				end if
				CloseOBJ rs
				set rs = Conn.Execute("select nombres from contactos where id_cliente=" & ColoaderID & " and activo = true")
				if Not rs.EOF then
					ColoaderData = ColoaderData & "<br>ATTN:" & rs(0)
				end if
				CloseOBJ rs
			end if
            CloseOBJ Conn
		end if
		
        Select Case BLType
		case 2,3
			MBLs = aTableValues(44,0)
			PickupData  = aTableValues(45,0)
			DeliveryData = aTableValues(46,0)
			PolicyNo = aTableValues(47,0)
            GuiaRemisionDet = aTableValues(48,0)
            NotifyPartyID = aTableValues(49,0)
            NotifyParty = aTableValues(50,0)
            EXDBCountry = aTableValues(52,0)
            GuiaRemision = aTableValues(48,0)
        case 0,1
            GuiaRemision = aTableValues(48,0)
            NotifyPartyID = aTableValues(49,0)
            NotifyParty = aTableValues(50,0)
        case 4
            GuiaRemision = aTableValues(41,0)
		End Select
	end if
	
	if CountTableValues2 >= 0 then	
		Redim Preserve NoOfPieces(CountTableValues2)
		Redim Preserve DiceContener(CountTableValues2)
		Redim Preserve Volumes(CountTableValues2)
		Redim Preserve Weights(CountTableValues2)
		Redim Preserve Clients(CountTableValues2)
		Redim Preserve ClassNoOfPieces(CountTableValues2)
		Redim Preserve CountryFinalDes(CountTableValues2)
		Redim Preserve CountryOrigen(CountTableValues2)
		Redim Preserve BLs(CountTableValues2)

		for i=0 to CountTableValues2
			NoOfPieces(i) = aTableValues2(0,i)
			DiceContener(i) = aTableValues2(1,i)
			Volumes(i) = aTableValues2(2,i)
			Weights(i) = aTableValues2(3,i)
			Clients(i) = aTableValues2(4,i)
			ClassNoOfPieces(i) = aTableValues2(7,i)

            'Cuando la carga ingresa en transito por CR (DestinoFinal<>CR) el destino final que se imprime es CR por temas de aduana
            'El dato real que esta en base de datos es diferente a CR (PA,NI, etc.)
            'cuando GroupID es 14,15 CountryDes -> CountriesFinalDes House y aTableValues(36,0) -> CountryDes Master
            ShowProcedencia = 1
            if (GroupID=0 and aTableValues2(12,i)<>"CR" and CountryDes="CR") or ((GroupID=14 or GroupID=15) and CountryDes<>"CR" and aTableValues(36,0)="CR") then
                CountryFinalDes(i) = "CR"
                CountryDes = "CR"
            else
                CountryFinalDes(i) = aTableValues2(12,i)
            end if
            'Cuando la carga sale en transito de CR con destino PA, no debe mostrarse la procedencia
			if (GroupID=0 and CountryDes="PA") or ((GroupID=14 or GroupID=15) and aTableValues(36,0)="PA") then
                ShowProcedencia = 0
            end if

			CountryOrigen(i) = aTableValues2(13,i)
			Bls(i) = aTableValues2(14,i)
		next

		if GroupID=14 or GroupID=15 then
            if AsAgreed = 0 then
			    response.write TransitConsignerData
                ConsignerData = TransitConsignerData & aTableValues2(4,0) & ConsignerData
			    Freight = aTableValues2(5,0)
			    Freight2 = aTableValues2(6,0)
			    Insurance = aTableValues2(8,0)
			    Insurance2 = aTableValues2(9,0)
			    AnotherChargesCollect = aTableValues2(10,0)
			    AnotherChargesPrepaid = aTableValues2(11,0)
                FreightColoader = aTableValues2(15,0)
			    FreightColoader2 = aTableValues2(16,0)
			    InsuranceColoader = aTableValues2(17,0)
			    InsuranceColoader2 = aTableValues2(18,0)
			    AnotherChargesColoader = aTableValues2(19,0)
			    AnotherChargesColoader2 = aTableValues2(20,0)
                FinalDesMaster = aTableValues(36,0)
			    if (FreightColoader = 0 and InsuranceColoader = 0 and AnotherChargesColoader = 0) then
                    TotCollect = Freight + Insurance + AnotherChargesCollect
                elseif (FreightColoader = 0 and InsuranceColoader = 0) then
                    TotCollect = Freight + Insurance + AnotherChargesColoader
                elseif (FreightColoader = 0 and AnotherChargesColoader = 0) then
                    TotCollect = Freight + InsuranceColoader + AnotherChargesCollect
                elseif (InsuranceColoader = 0 and AnotherChargesColoader = 0) then
                    TotCollect = FreightColoader + Insurance + AnotherChargesCollect
                elseif (FreightColoader = 0) then
                    TotCollect = Freight + InsuranceColoader + AnotherChargesColoader
                elseif (InsuranceColoader = 0) then
                    TotCollect = FreightColoader + Insurance + AnotherChargesColoader
                elseif (AnotherChargesColoader = 0) then
                    TotCollect = FreightColoader + InsuranceColoader + AnotherChargesCollect
                else
                    TotCollect = FreightColoader + InsuranceColoader + AnotherChargesColoader
                end if
			    if (FreightColoader2 = 0 and InsuranceColoader2 = 0 and AnotherChargesColoader2 = 0) then
                    TotPrepaid = Freight2 + Insurance2 + AnotherChargesPrepaid
                elseif (FreightColoader2 = 0 and InsuranceColoader2 = 0) then
                    TotPrepaid = Freight2 + Insurance2 + AnotherChargesColoader2
                elseif (FreightColoader2 = 0 and AnotherChargesColoader2 = 0) then
                    TotPrepaid = Freight2 + InsuranceColoader2 + AnotherChargesPrepaid
                elseif (InsuranceColoader2 = 0 and AnotherChargesColoader2 = 0) then
                    TotPrepaid = FreightColoader2 + Insurance2 + AnotherChargesPrepaid
                elseif (FreightColoader2 = 0) then
                    TotPrepaid = Freight2 + InsuranceColoader2 + AnotherChargesColoader2
                elseif (InsuranceColoader2 = 0) then
                    TotPrepaid = FreightColoader2 + Insurance2 + AnotherChargesColoader2
                elseif (AnotherChargesColoader2 = 0) then
                    TotPrepaid = FreightColoader2 + InsuranceColoader2 + AnotherChargesPrepaid
                else
                    TotPrepaid = FreightColoader2 + InsuranceColoader2 + AnotherChargesColoader2
                end if
			    for i=0 to CountTableValues2
				    TotNoOfPieces = TotNoOfPieces + NoOfPieces(i)
				    TotVolume = TotVolume + Volumes(i)
				    TotWeight = TotWeight + Weights(i)
			    next
            else
                ConsignerData = aTableValues2(4,0) & ConsignerData
			    TotCollect = "AS AGREED"
			    TotPrepaid = "AS AGREED"
            end if
		end if
	end if	

    
    'Countries = "GT"
    dim xy
    xy = 1
    On Error Resume Next 'cuando se crea un bl en pendientes countries no trae valor poque aun no hay cp master
        xy = CheckNum(Countries + 1)

    If Err.Number<>0 then
        Err.Number = 0
	    'response.write "Trin Info 1 :" & Err.Number & " - " & Err.Description & "<br>"  
    end if

    if xy = 0 then        
        Countries = EXDBCountry
    end if
	
    'response.write EXDBCountry & " (" & Countries & ") - " & len(Countries) & "<br>"

    if CheckNum(Request("id_routing")) then  '2019-06-21 cuando es LTF      
        
        'ColoaderData = "LATIN FREIGHT NEUTRAL LOGISTICS"        
        'ShipperData = "LATIN FREIGHT NEUTRAL LOGISTICS."

        'if ColoaderID = NotifyPartyID then
            OpenConn2 Conn        
            Set rs = Conn.Execute("SELECT id_cliente, id_shipper, id_notify, id_coloader, id_routing FROM routings WHERE id_routing = '" & CheckNum(Request("id_routing")) & "'")
            if Not rs.EOF then                	                                    
                'SenderData = "(" & rs(3) & ")(" & ColoaderID & ") " & ColoaderData            
                SenderData = ColoaderID & ". " & ColoaderData            
                'ConsignerData = "(" &  rs(2) & ")(" & NotifyPartyID & " " & NotifyParty & "<br><br>" & Attn              
                ConsignerData = NotifyPartyID & ". " & NotifyParty & "<br><br>" & Attn              
	        end if
            'ColoaderData = "(" &  rs(1) & ")(" & ShipperID & ") " & ShipperData
            ColoaderData = ShipperID & ". " & ShipperData
            CloseOBJs rs, Conn    
        'end if

    end if


Set GroupValues = Nothing
Set aTableValues = Nothing
Set aTableValues2 = Nothing

ShowConsigners = 0
if GroupID=0 and Consolidated=1 then
	ShowConsigners = 1
end if	

'Se toma primero el pais de la base de datos para desplegar el logo, si viene vacio se toma el pais donde se crea el registro
if EXDBCountry = "" then
    EXDBCountry = Countries
end if

select case Countries
Case "N1"
    CompanyName = "GRH"
Case "GTLTF","SVLTF","HNLTF","NILTF","CRLTF","PALTF"
    CompanyName = "Latin Freight"
Case Else
    CompanyName = "Aimar"
End Select

    '9 CP individual
    '10 CP master
    'dim iEdicion, iTitulo, iEmpresa, iDireccion, iLogo, iObservaciones, iDocID, aTableValues5
    
        Dim iResult, iLogo, iEdicion, iTitulo, iEmpresa, iDireccion, iObservaciones, iPlantilla, iDocID, iMedida, iVol, iBulto
            
        iDocID = IIf(GroupID = 15 or GroupID = 14, "9", IIf(GroupID = 0, "10", "")) '9 cp individual 10 cp master
        
        'response.write "(" & GroupID & ")(" & Countries & ")(" & iDocID & ")"

        iResult = WsGetLogo(Countries, "TERRESTRE",  iDocID,  "",  "")
        iLogo = iResult(20)
        iEdicion = iResult(2)
        iTitulo = iResult(3)
        iEmpresa = iResult(4)
        iDireccion = iResult(6)
        iObservaciones = iResult(1)
        iPlantilla = iResult(22)
        CompanyName = iResult(4)
        Footer = iResult(6) 

    'EXDBCountry = "GTTLA"
    'if iDocID <> "" then
    '    'aTableValues5 = EmpresaParametros(EXDBCountry, iDocID, "TERRESTRE")    
    '    aTableValues5 = EmpresaParametros(Countries, iDocID, "TERRESTRE")    
    '    if aTableValues5(1,0) <> "" then
    '        'iLogo = "<img src='data:image/jpeg;base64," & aTableValues5(20,0) & "'>"
    '        iLogo = aTableValues5(20,0)
    '        iEdicion = aTableValues5(3,0)
    '        iTitulo = aTableValues5(4,0)
    '        iEmpresa = aTableValues5(5,0)
    '        iDireccion = aTableValues5(7,0)
    '        iObservaciones = aTableValues5(11,0)
    '        CompanyName = iEmpresa
    '        Footer = iDireccion 
    '    end if    
    'end if

    'response.write "(BLType=" & BLType & ")(GroupID=" & GroupID & ")(" & Countries & ")(" & iEdicion & ")(" & CompanyName & ")(" & CountryDep & ")(" & EXDBCountry & ")"

    iMedida = "&nbsp;Kgs"
    iVol = "&nbsp;CBM"
    iBulto = "&nbsp;Bultos"


%>
<html>
<HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
    <!--
body {
	margin: 0px;
        font-weight: 700;
    }
.style3 {
	font-size:8px;
	color: #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
}
.style4 {
	font-size:9px;
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
.style10 {font-size: 9px; font-family: Verdana, Arial, Helvetica, sans-serif; font-weight:normal;}
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

    .style13
    {
        font-family: Arial, Helvetica, sans-serif;
        font-size: 15px;
        height: 22px;
    }

-->
</style>
</head>
<body onLoad="JavaScript:self.focus();">
<% 


select case BLType
case 12%>
    <table width="750" cellpadding="2" cellspacing="0" align="center">
        <tr>
            <td class="style11" align="left" width="50%">
                <%
                if FRegExp(PtrnLogo, AgentsID,  "", 2) or (ColoaderID > 0) then
		            response.write DisplayLogo(AgentsID, ColoaderID, IsColoader, EXDBCountry, iLogo)
	            else
		            response.write DisplayLogo(EXDBCountry, 0, 0, 0, iLogo)
	            end if
	            %>
            </td>
            <td class="style3" align="right" width="50%"><%=IIf(iEdicion = "", "FO-TR-03<br>ORIGINAL", iEdicion)%></td>
        </tr>    
    </table>
    <br /><br />
<%case 0, 1, 4%>
<table width="750" cellpadding="2" cellspacing="0" align="center">
  <tr>
	<td class="style11" align="left" width="45%">
	<%
	if CheckNum(Request("id_routing")) then     
    	response.write DisplayLogo(EXDBCountry, 0, 0, 0, iLogo)
	else
            'response.write "(" & PtrnLogo & ")(" & AgentsID & ")(" & ColoaderID & ")<br>"
        if FRegExp(PtrnLogo, AgentsID,  "", 2) or (ColoaderID > 0) then
		    response.write DisplayLogo(AgentsID, ColoaderID, IsColoader, EXDBCountry, iLogo)
            'response.write AgentsID & "-" & ColoaderID & "-" & IsColoader
	    else
            'response.write "(" & EXDBCountry & ")<br>"
		    response.write DisplayLogo(EXDBCountry, 0, 0, 0, iLogo)
	    end if
	end if
	%></td>
          <%if (CountryDep="HN" or CountryDep="HN1" or CountryDep="HN2") and (inStrRev(EXDBCountry,"LTF")=0) then %>
    <td class="style3" width="45%">Agencia Internacional Maritima S.A</br>
        Col. Brisas de la Mesa contiguo a la Base Aerea,</br>
        Aereopuerto Ramon Villeda Morales</br>
        PBX (504) 2564-0099/ 2668-0121 FAX (504) 2668-0353</br>
        La Lima, San Pedro Sula, Hondruas, C.A.</br>
        RTN.: 05019000044051</br>
        <%if (GuiaRemision < 4000) then%>
            CAI-5890CI-026427-254380-AB3502-39D69C-87</br>
		<%elseif (GuiaRemision > 4000 and GuiaRemision < 7501) then%>
            CAI-6040EC-C37B99-0D43B1-B4A85F-E93234-39</br>
        <%elseif (GuiaRemision > 7500 and GuiaRemision < 12001) then %>
            CAI-681E5F-D789A5-D547B3-DFFFD9-BD9737-57</br>
        <%elseif (GuiaRemision > 12000 and GuiaRemision < 16001) then %>
            CAI-D48F40-024639-9E4EA9-0B4494-A06D6C-6C</br>
        <%end if%>
    </td>
  <%end if %>
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
    </td>
	<td class="style3" align="right"><%=IIf(iEdicion = "", "FO-TR-03<br>ORIGINAL", iEdicion)%></td>
  </tr>
</table>
<table width="750" class="styleborder" cellpadding="2" cellspacing="0" align="center" border=0>
  <tr>
    <!--BLN es un tag para encontrar el BLNumber en el HTML cuando se parsea para el envio de prealerta -->
	<td class="style13" align="left"><font color="#0000FF"><%=IIf(iTitulo = "", "CARTA&nbsp;PORTE&nbsp;No.", iTitulo)%>:&nbsp;<!--BLN--><%=BLNumber%><!--/BLN--></font>&nbsp;&nbsp;&nbsp;</td>
    <td class="style4" align="right" style="border-left:0px;"><%if (CountryDep="HN" or CountryDes="HN" or CountryDep="HN1" or CountryDes="HN1" or CountryDep="HN2" or CountryDes="HN2") then %>&nbsp;&nbsp;&nbsp;&nbsp;Guia de Remision: &nbsp;&nbsp;000-001-08-<%=NumeroCerosGuiaRemision %><%=GuiaRemision%><% end if%></td>
  </tr>
</table>
<table width="750" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
    <td width="60%" class="style4" align="left" valign="top">Shipper / Exportador:<br>
      <span class="style10">
      <%=SenderData%>
      </span></td>
    <td class="style4" align="left" valign="top">Embarcador:<br>
    <span class="style10">
      <%if IsColoader=1 then %>
      <%=ColoaderData%>
      <%else %>
      <%=ShipperData%>
      <%end if %>
    </span></td>
  </tr>
  <tr>
    <td class="style4" align="left" valign="top">Consignee / Consignatario:<br>
    <span class="style10"><%=ConsignerData%></span></td>
    <%if GroupID = 15 or GroupID = 14 then %>
        <%for i=0 to CountTableValues2%>
            <%if ((CountryFinalDes(i) = "NI") and (CountryFinalDes(i) = (FinalDesMaster))) then %>
                <td class="style4" align="left" valign="top">Pais de Origen:<br><span class="style10"><%=TranslateCountry(Mid(BLs(i),2,2))%></span>
            <%else%>
                <td class="style4" align="left" valign="top">Pais de Origen:<br><span class="style10"><%=TranslateCountry(Left(CountryDep,2))%></span>
            <%end if%>
        <%next %>
    <%else %>
        <td class="style4" align="left" valign="top">Pais de Origen:<br><span class="style10"><%=TranslateCountry(Left(CountryDep,2))%></span>
    <%end if %>
	<%if GroupID=14 or GroupID=15 then%>
        <%if ShowProcedencia=1 then%>		
        <br><br>Pais de Procedencia:<br><span class="style10"><%=TranslateCountry(Left(Countries,2))%></span>
        <!--<br><br>Pais de Procedencia:<br><span class="style10"><%=TranslateCountry(Mid(BLNumber,2,2))%></span>-->
        <%end if %>
	<%end if%>
	</td>
  </tr>
  <tr>
    <%if GroupID=14 or GroupID=15 then%>
        <td class="style4" align="left" valign="top">Notificaci&oacute;n a:<br><span class="style10"><%=NotifyPartyID & " " & NotifyParty & "<br><br>" & Attn%></span></td>
    <%else %>
        <td class="style4" align="left" valign="top">Notificaci&oacute;n a:<br><span class="style10"><%=Attn%></span></td>
    <%end if%>
    <td class="style4" align="left" valign="top">Instrucciones de Exportaci&oacute;n:<br><span class="style10"><%=HandlingInformation%></span></td>
  </tr>
  <tr>
    <td class="style4" align="left" valign="top">Aduana de Transito:<br>
      <span class="style10"><%=BrokerName%></span></td>
    <td class="style4" align="left" valign="top">Pais de Destino:<br><span class="style10"><%=TranslateCountry(Left(CountryDes,2))%></span></td>
  </tr>
  <tr>
    <td class="style4" align="left" valign="top">Lugar de Carga:<br><span class="style10"><%=ChargePlace%></span></td>
    <td class="style4" align="left" valign="top">Destino Final:<br><span class="style10"><%=FinalDes%></span></td>
  </tr>
  </table>
  <table width="750" class="styleborder" cellpadding="2" cellspacing="0" align="center" border="0">
  <tr>
    <td class="style4" align="left" valign="top">Conductor:<br><span class="style10"><%=PilotName%></span></td>
    <td class="style4" align="left" valign="top">Licencia:<br><span class="style10"><%=License%></span></td>
    <td class="style4" align="left" valign="top">Tel&eacute;fono:<br><span class="style10"><%=Phone1%></span></td>
  </tr>
  <tr>
  
    <td align="left" valign="top" class="style4">
    Cabezal/Cami&oacute;n:<br><span class="style10"><%=TruckNo%>
    </td>

    <!--
    <td align="left" valign="top">
    
        <table width="100%" border="0" cellpadding=0 cellspacing=0>
        <tr>
        <td width="75%" class="style4" style="border:0px;border-right:1px solid black">C&oacute;digo:<br /><span style="font-weight:normal"><%=CodProv%></span></td>
        <td width="25%" class="style4" style="border:0px;padding-left:3px;">Cabezal/Cami&oacute;n:<br><span class="style10"><%=TruckNo%></td>
        </tr>
        </table>
     -->
    
    <%if Container <> "" then%>
	<td class="style4" align="left" valign="top" colspan="2">Furgon:<br><span class="style10"><%=Container%></span></td>
	<%else%>
	<td class="style4" align="left" valign="top">Contenedor:<br><span class="style10"><%=ContainerDep%></span></td>
	<td class="style4" align="left" valign="top">Chassis:<br><span class="style10"><%=Chassis%></span></td>
	<%end if%>
  </tr>
  <%if GroupID=14 or GroupID=15 then%>
  <tr>
    <td class="style4" align="left" valign="top">Documento de Transporte Interno / PO:<br><span class="style10"><%=PO%></span></td>
    <td class="style4" align="left" valign="top" colspan=2>Factura:<br><span class="style10"><%=Bill%></span></td>
  </tr>
  <%end if%>
</table>

<table width="750" height="280" class="styleborder" cellpadding="2" cellspacing="0" align="center">
	<tr height="8">
		<td class="style4" align="center" valign="middle">No. Bultos</td>
		<td class="style4" align="center" valign="middle">Clase Bultos</td>
		<td class="style4" align="center" valign="middle">Descripci&oacute;n de Carga</td>
		<td class="style4" align="center" valign="middle">Volumen (CBM)</td>
		<td class="style4" align="center" valign="middle">Peso Bruto(Kg)</td>
		<td class="style4" align="center" valign="middle">Origen</td>
		<%if ShowProcedencia=1 then%>
        <td class="style4" align="center" valign="middle">Procedencia</td>
		<%end if%>
        <td class="style4" align="center" valign="middle">Destino</td>
		<%if ShowConsigners=1 then%>
		<td class="style4" align="center" valign="middle">Consignatario Final</td>
		<%end if%>
	</tr>
	<%for i=0 to CountTableValues2%>
	<tr>
		<td class="style4" align="right" valign="top"><span class="style10"><%=NoOfPieces(i)%></span></td>
		<td class="style4" align="left" valign="top"><span class="style10"><%=ClassNoOfPieces(i)%></span></td>
		<td class="style4" align="left" valign="top"><span class="style10"><%=DiceContener(i)%></span></td>
		<td class="style4" align="right" valign="top"><span class="style10"><%=Volumes(i)%><%=iVol%></span></td>
		<td class="style4" align="right" valign="top"><span class="style10"><%=Weights(i)%><%=iMedida%></span></td>
		<%if (CountryFinalDes(i) = "NI" and (CountryFinalDes(i) = FinalDesMaster)) then %>
            <td class="style4" align="right" valign="top"><span class="style10"><%=Left(Mid(BLs(i),2,2),2)%></span></td>
        <%else %>
            <td class="style4" align="right" valign="top"><span class="style10"><%=Left(CountryOrigen(i),2)%></span></td>
        <%end if %>
        <%if ShowProcedencia=1 then%>
        <td class="style4" align="right" valign="top"><span class="style10"><%=Left(Countries,2)%></span></td>
        <%end if%>
		<!--<td class="style4" align="right" valign="top"><span class="style10"><%=Mid(BLs(i),2,2)%></span></td>-->
        <td class="style4" align="right" valign="top"><span class="style10"><%=Left(CountryFinalDes(i),2)%></span></td>
        <%if ShowConsigners=1 then%>
		<td class="style4" align="right" valign="top"><span class="style10"><%=Clients(i)%></span></td>
		<%end if%>
	</tr>
	<%next%>
	<tr>
		<td class="style4" align="right" valign="top" colspan="8" height="100%">&nbsp;</td>
	</tr>
	<tr height="8">
		<td class="style4" align="right" valign="top"><%=TotNoOfPieces%><%=iBulto%></td>
		<td class="style4" align="center" valign="top" colspan="2">TOTALES</td>
		<td class="style4" align="right" valign="top"><%=TotVolume%><%=iVol%></td>
		<td class="style4" align="right" valign="top"><%=TotWeight%><%=iMedida%></td>
		<%if ShowConsigners=1 then%>
		<td class="style4" align="right" valign="top" colspan="4">&nbsp;</td>
		<%end if%>
<!--		<td class="style4" align="right" valign="top">9</td>
		<td class="style4" align="center" valign="top">TOTALES</td>
		<td class="style4" align="right" valign="top">0.28</td>
		<td class="style4" align="right" valign="top">104.50</td>
		<%if ShowConsigners=1 then%>
		<td class="style4" align="right" valign="top">&nbsp;</td>
		<%end if%>
-->
	</tr>
</table>
<br>
<table width="750" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
    <td width="168" align="center" class="style4">PREPAGADO (USD)</td>
    <td width="165" align="center" class="style4">AL COBRO (USD)</td>
	<td width="488" align="center" class="style4" colspan="2">Observaciones / Comentarios</td>
  </tr>
 <tr>
    <td class="style4" align="center" colspan="2">Flete</td>
	<td width="488" rowspan="4" align="center" class="style10" colspan="2" valign="top"><%=Observations%>&nbsp;</td>
 </tr>
  <tr>
    <%if FreightColoader2 <> 0 then %>
        <td class="style4" align="right"><span class="style10"><%=FreightColoader2%></span></td>
    <%else %>
        <td class="style4" align="right"><span class="style10"><%=Freight2%></span></td>
    <%end if %>
    <%if FreightColoader <> 0 then %>
        <td class="style4" align="right"><span class="style10"><%=FreightColoader%></span></td>
    <%else %>
        <td class="style4" align="right"><span class="style10"><%=Freight%></span></td>
    <%end if %>
  </tr>
 <tr>
    <td class="style4" align="center" colspan="2">Seguros</td>
 </tr>
  <tr>
    <%if InsuranceColoader2 <> 0 then %>
        <td class="style4" align="right"><span class="style10"><%=InsuranceColoader2%></span></td>
    <%else %>
        <td class="style4" align="right"><span class="style10"><%=Insurance2%></span></td>
    <%end if %>
    <%if InsuranceColoader <> 0 then %>
        <td class="style4" align="right"><span class="style10"><%=InsuranceColoader%></span></td>
    <%else %>
        <td class="style4" align="right"><span class="style10"><%=Insurance%></span></td>
    <%end if %>
  </tr>
   <tr>
    <td class="style4" align="center" colspan="2">Otros</td>
	<td class="style4" align="center">Fecha</td>
	<td class="style4" align="center">Lugar de Emisi&oacute;n </td>
 </tr>
  <tr>
    <%if AnotherChargesColoader2 <> 0 then %>
        <td class="style4" align="right"><span class="style10"><%=AnotherChargesColoader2%></span></td>
    <%else %>
        <td class="style4" align="right"><span class="style10"><%=AnotherChargesPrepaid%></span></td>
    <%end if %>
    <%if AnotherChargesColoader <> 0 then %>
        <td class="style4" align="right"><span class="style10"><%=AnotherChargesColoader%></span></td>
    <%else %>
        <td class="style4" align="right"><span class="style10"><%=AnotherChargesCollect%></span></td>
    <%end if %>
    <td class="style4" align="center"><span class="style10"><%=BLDispatchDate%></span></td>
    <td class="style4" align="center"><span class="style10"><%=TranslateCountry(Left(Countries,2))%></span></td>
  </tr>
  <tr>
    <td class="style4" align="center">TOTAL PREPAGADO (USD)</td>
    <td class="style4" align="center">TOTAL AL COBRO (USD)</td>
	<td class="style4" align="center">Medio de Transporte: TERRESTRE </td>
	<td class="style4" align="center"><%=BLNumber%></td>
  </tr>
  <tr>
    <td class="style4" align="right"><%=TotPrepaid%></td>
    <td class="style4" align="right"><%=TotCollect%></td>
    <td class="style4" align="center">&nbsp;</td>
    <td class="style4" align="center">&nbsp;</td>
  </tr>
</table>

<%case 2,3%>
<table width="750" cellpadding="2" cellspacing="0" align="center">
  <tr>
	<td class="style11" align="left" width="45%"><%=DisplayLogo(EXDBCountry, 0, 0, 0, iLogo)%></td>
        <%if (CountryDep="HN" or CountryDep="HN1" or CountryDep="HN2") then %>
            <td class="style3" width="45%"><%=IIf(iEmpresa = "", "Agencia Internacional Maritima S.A.", iEmpresa)%></br>
            Col. Brisas de la Mesa contiguo a la Base Aerea,</br>
            Aereopuerto Ramon Villeda Morales</br>            
            La Lima, San Pedro Sula, Hondruas, C.A.</br>
            PBX (504) 2564-0099/ 2668-0121 FAX (504) 2668-0353</br>
            RTN.: 05019000044051</br>
            <%if (GuiaRemision < 4000) then%>
                CAI-5890CI-026427-254380-AB3502-39D69C-87</br>
		    <%elseif (GuiaRemision > 4000 and GuiaRemision < 7501) then%>
                CAI-6040EC-C37B99-0D43B1-B4A85F-E93234-39</br>
            <%elseif (GuiaRemision > 7500 and GuiaRemision < 12001) then %>
                CAI-681E5F-D789A5-D547B3-DFFFD9-BD9737-57</br>
            <%elseif (GuiaRemision > 12000 and GuiaRemision < 16001) then %>
                CAI-D48F40-024639-9E4EA9-0B4494-A06D6C-6C</br>
            <%end if%>
        <%end if %>
    </td>
	<td class="style3" align="right"><%=IIf(iEdicion = "", "FO-TR-03", iEdicion)%></td>
  </tr>
</table>
<table width="750" class="styleborder" cellpadding="2" cellspacing="0" align="center" border="1">
  <tr>
    <% CountGuiaRemision = int(len(cstr(GuiaRemisionDet))) 
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
	<td class="style11" align="left" width="80%"><font color="#0000FF">CARTA&nbsp;PORTE&nbsp;No.:&nbsp;<%=BLNumber%></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%if (CountryDep="HN" or CountryDes="HN" or CountryDep="HN1" or CountryDes="HN1" or CountryDep="HN2" or CountryDes="HN2") then %>Guia Remision: <% end if%></td>
	<td class="style3" align="right" width="20%"><%if (CountryDep="HN" or CountryDes="HN" or CountryDep="HN1" or CountryDes="HN1" or CountryDep="HN2" or CountryDes="HN2") then %> 000-001-08-<%=NumeroCerosGuiaRemision %><%=GuiaRemisionDet%><% else %><%if BLType<>4 then%>Entrega<%else%>Recolecci&oacute;n<%end if%> de Carga <% end if%></td>
  </tr>
</table>
<table width="750" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
    <td class="style4" align="left" valign="top" colspan="2">Consignee /Consignatario:<br>
    <span class="style10"><%=ConsignerData%></span></td>
  </tr>
  <tr>
    <td class="style4" align="left" valign="top" colspan="2">Notificaci&oacute;n a:<br><span class="style10"><%=Attn%></span></td>
  </tr>
  </table>
  <table width="750" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
    <td class="style4" align="left" valign="top" colspan="2">Recolecci&oacute;n:<br><span class="style10"><%=PickupData%></span></td>
    <td class="style4" align="left" valign="top" colspan="2">Seguridad <%=CompanyName%>:<br><span class="style10"><%if HasSecure=1 then%>SI<%else%>NO<%end if%></span></td>
  </tr>
  <tr>
    <td class="style4" align="left" valign="top" colspan="2">Entrega:<br><span class="style10"><%=Deliverydata%></span></td>
    <td class="style4" align="left" valign="top" colspan="2">P&oacute;liza:<br><span class="style10"><%=PolicyNo%></span></td>
  </tr>
  <tr>
    <td class="style4" align="left" valign="top">Conductor:<br><span class="style10"><%=PilotName%></span></td>
    <td class="style4" align="left" valign="top">Licencia:<br><span class="style10"><%=License%></span></td>
    <td class="style4" align="left" valign="top" colspan="2">Tel&eacute;fono:<br><span class="style10"><%=Phone1%></span></td>
  </tr>
  <tr>
    <td class="style4" align="left" valign="top">Cabezal/Cami&oacute;n:<br><span class="style10"><%=TruckNo%></span></td>
    <%if Container = "" then%>
	<td class="style4" align="left" valign="top" colspan="3">Furgon:<br><span class="style10"><%=Container%></span></td>
	<%else%>
	<td class="style4" align="left" valign="top">Contenedor:<br><span class="style10"><%=ContainerDep%></span></td>
	<td class="style4" align="left" valign="top" colspan="2">Chassis:<br><span class="style10"><%=Chassis%></span></td>
	<%end if%>
  </tr>
  <tr>
    <td class="style4" align="left" valign="top">BL/AWB/RO:<br><span class="style10"><%=MBLs%></span></td>
    <td class="style4" align="left" valign="top">Unidad:<br><span class="style10"><%=UnitNo%></span></td>
	<td class="style4" align="left" valign="top">Tama&ntilde;o:<br><span class="style10"><%=BLSize%></span></td>
	<td class="style4" align="left" valign="top">Marchamo:<br><span class="style10"><%=Marchamo%></span></td>
  </tr>
</table>
<table width="750" height="280" class="styleborder" cellpadding="2" cellspacing="0" align="center">
	<tr height="8">
		<td class="style4" align="center" valign="middle">No. Bultos</td>
		<td class="style4" align="center" valign="middle">Clase Bultos</td>
		<td class="style4" align="center" valign="middle">Descripci&oacute;n de Carga</td>
		<td class="style4" align="center" valign="middle">Volumen (CBM)</td>
		<td class="style4" align="center" valign="middle">Peso Bruto(Kg)</td>
		<td class="style4" align="center" valign="middle">Origen</td>
		<td class="style4" align="center" valign="middle">Destino</td>
		<%if ShowConsigners=1 then%>
		<td class="style4" align="center" valign="middle">Consignatario Final</td>
		<%end if%>
	</tr>
	<%for i=0 to CountTableValues2%>
	<tr>
		<td class="style4" align="right" valign="top"><span class="style10"><%=NoOfPieces(i)%></span></td>
		<td class="style4" align="left" valign="top"><span class="style10"><%=ClassNoOfPieces(i)%></span></td>
		<td class="style4" align="left" valign="top"><span class="style10"><%=DiceContener(i)%></span></td>
		<td class="style4" align="right" valign="top"><span class="style10"><%=Volumes(i)%></span></td>
		<td class="style4" align="right" valign="top"><span class="style10"><%=Weights(i)%></span></td>
		<td class="style4" align="right" valign="top"><span class="style10"><%=Left(CountryOrigen(i),2)%></span></td>
		<td class="style4" align="right" valign="top"><span class="style10"><%=Left(Mid(BLs(i),2,2),2)%></span></td>
		<td class="style4" align="right" valign="top"><span class="style10"><%=Left(CountryFinalDes(i),2)%></span></td>
		<%if ShowConsigners=1 then%>
		<td class="style4" align="right" valign="top"><span class="style10"><%=Clients(i)%></span></td>
		<%end if%>
	</tr>
	<%next%>
	<tr>
		<td class="style4" align="right" valign="top" colspan="8" height="100%">&nbsp;</td>
	</tr>
	<tr height="8">
		<td class="style4" align="right" valign="top"><%=TotNoOfPieces%></td>
		<td class="style4" align="center" valign="top" colspan="2">TOTALES</td>
		<td class="style4" align="right" valign="top"><%=TotVolume%></td>
		<td class="style4" align="right" valign="top"><%=TotWeight%></td>
		<%if ShowConsigners=1 then%>
		<td class="style4" align="right" valign="top" colspan="4">&nbsp;</td>
		<%end if%>
	</tr>
</table>
<br>
<table width="750" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
	<td align="left" class="style4" colspan="2">Observaciones / Comentarios<br><span class="style10"><%=Observations%></span><br><br><br><br><br><br><br></td>
  </tr>
 <tr>
	<td class="style4" align="center"><br><br></td>
	<td class="style4" align="center"><br><br></td>
 </tr>
 <tr>
	<td class="style4" align="center"><span class="style10">Firma, Nombre Quien Recibo/Entrega por <%=CompanyName%></span></td>
	<td class="style4" align="center"><span class="style10">Firma, Nombre, y/o Sello Quien Recibe/Entrega la Carga</span></td>
 </tr>
 <tr>
	<td class="style4" align="center">Fecha</td>
	<td class="style4" align="center">Lugar de Emisi&oacute;n </td>
 </tr>
  <tr>
    <td class="style4" align="center"><span class="style10"><%=BLDispatchDate%></span></td>
    <td class="style4" align="center"><span class="style10"><%=TranslateCountry(Left(Countries,2))%></span></td>
  </tr>
  <tr>
	<td align="left" class="style4" colspan="2">
	<span class="style10">
    <% if iObservaciones = "" then  %>
	    Si la mercader&iacute;a no es recibida a su satisfacci&oacute;n y <%=CompanyName%> no aseguro el mismo favor comunicarse con su Agente de Seguros.<br>
	    Si la mercader&iacute;a es asegurada por <%=CompanyName%> comunicarse a nuestro departamento de Calidad y Reclamos<%if Countries="GT" then%>, al 2329-8200 Ext.2801.<%end if%><br>
	    Cualquier atraco y/o asalto <%=CompanyName%> no es responsable del mismo, se proceder&aacute; con los tramites legales seg&uacute;n nuestra constituci&oacute;n.<br>
	    Cualquier comentario o sugerencia adicional la puede expresar comunicandose con nosotros<%if Countries="GT" then%> al 2329-8200, Ext.2801.<%end if%><br>
	    <%if Countries="GT" then%>
	    Entregas y Recolecci&oacute;n en unidades mayores de 3.5 toneladas de 09:00 a 16:30 (esto se debe a restricci&oacute;n de horario seg&uacute;n acuerdo COM-005-07
	    <%end if%>
    <%else    
        'response.write "///////////////////////CODIGO NUEVO//////////////////////////////<BR>"
        response.write iObservaciones
    end if%>

    <%=Session("Sign") %>
	</span><br><br></td>
  </tr>
</table>
<%case 5%>
<table width="750" cellpadding="2" cellspacing="0" align="center">
  <tr>
	<td class="style11" align="left" width="50%" colspan="2"><%=DisplayLogo(EXDBCountry, 0, 0, 0, iLogo)%></td>
  </tr>
</table>
<table width="750" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
	<td class="style11" align="left"><font color="#0000FF">CARTA&nbsp;SEGURO&nbsp;No.:&nbsp;<%=BLNumber%></font></td>
  </tr>
</table>
<table width="750" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
    <td class="style4" align="left" valign="top" width="50%">Shipper / Exportador:<br>
    <span class="style10"><%=SenderData%></span></td>
    <td class="style4" align="left" valign="top">Embarcador:<br>
    <span class="style10"><%=ShipperData%></span></td>
  </tr>
  <tr>
    <td class="style4" align="left" valign="top">Consignee / Consignatario:<br>
      <span class="style10"><%=ConsignerData%></span></td>
    <td class="style4" align="left" valign="top">
	Pais de Origen:<br><span class="style10"><%=TranslateCountry(Left(CountryDep,2))%></span><br><br>
    Pais de Destino:<br><span class="style10"><%=TranslateCountry(Left(CountriesFinalDes,2))%></span></td>
  </tr>
</table>

<table width="750" height="370" class="styleborder" cellpadding="2" cellspacing="0" align="center">
	<tr height="8">
		<td class="style4" align="center" valign="middle">No. Bultos</td>
		<td class="style4" align="center" valign="middle">Clase Bultos</td>
		<td class="style4" align="center" valign="middle">Descripci&oacute;n de Carga</td>
		<td class="style4" align="center" valign="middle">Volumen (CBM)</td>
		<td class="style4" align="center" valign="middle">Peso Bruto(Kg)</td>
	</tr>
	<%for i=0 to CountTableValues2%>
	<tr>
		<td class="style4" align="right" valign="top"><span class="style10"><%=NoOfPieces(i)%></span></td>
		<td class="style4" align="left" valign="top"><span class="style10"><%=ClassNoOfPieces(i)%></span></td>
		<td class="style4" align="left" valign="top"><span class="style10"><%=DiceContener(i)%></span></td>
		<td class="style4" align="right" valign="top"><span class="style10"><%=Volumes(i)%></span></td>
		<td class="style4" align="right" valign="top"><span class="style10"><%=Weights(i)%></span></td>
	</tr>
	<%next%>
	<tr>
		<td class="style4" align="right" valign="top" colspan="6" height="100%">&nbsp;</td>
	</tr>
	<tr height="8">
		<td class="style4" align="right" valign="top"><%=TotNoOfPieces%></td>
		<td class="style4" align="center" valign="top" colspan="2">TOTALES</td>
		<td class="style4" align="right" valign="top"><%=TotVolume%></td>
		<td class="style4" align="right" valign="top"><%=TotWeight%></td>
	</tr>
</table>
<%end select%>
</body>
</html>



<%

If Err.Number<>0 then
	response.write "BLPrint :" & Err.Number & " - " & Err.Description & "<br>"  
    Err.Number = 0
end if

%>