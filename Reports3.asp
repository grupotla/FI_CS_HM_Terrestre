<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"
Dim Action, ObjectID, GroupID, QuerySelect, Conn, rs, ntr, CountryTitle, QuerySelect2, QuerySelect3, BLID
Dim CountTableValues, aTableValues, CountTableValues2, aTableValues2, CountTableValues3, aTableValues3
Dim BLNumber, BLDispatchDate, PilotName, ShipperName, License, Countries, TruckNo, Mark, Model, CodProv, Attn, Chassis, SObjectID
Dim SenderData, ConsignerData, CountryDes, Bail, Container, ContainerDep, FinalDes, BLArrivalDate, Consolidated, GroupValues
Dim SubTotNoOfPieces, SubTotWeight, SubTotVolume, TotNoOfPieces, TotWeight, TotVolume, Week, i, Nacionality, DTI, BLType, ShipperID
Dim LtAcceptNumber, LtAcceptDate, BrokerRecepName, BrokerName, Logo, Footer, BusinessName, Estate, LtEndorseDate, ClientID, AgentID
Dim DiceContener(), Weights(), Volumes(), Clients(), ClassNoOfPieces(), NoOfPieces(), CountriesFinalDes(), BLs(), ExBLs(), MBLs()
Dim DischargeDate(), Agents(), Pos(), Consigners(), CountriesDes(), CountryOrigen(), Consig, ManifComment, Sep
Dim CIFBrokerIn, BL, MBL, CIFLandFreight, Contener, ColoaderID, IsColoader, EXDBCountry, ExType

	BLID = CheckNum(Request("BLID"))
	GroupID = CheckNum(Request("GID"))
	ObjectID = CheckNum(Request("OID"))
	SObjectID = CheckNum(Request("SOID"))
	ClientID = CheckNum(Request("CID"))
	AgentID = CheckNum(Request("AID"))
	Sep = CheckNum(Request("SEP"))
	BLType = CheckNum(Request("AT"))
	CountTableValues = -1
	CountTableValues2 = -1
	CountTableValues3 = -1
	SubTotNoOfPieces = 0
	ntr = chr(13) & chr(10)
	
	Select case GroupID
	case 12 'Itinerario Individual
		QuerySelect = "select a.BLNumber, c.Name, a.Week, a.BLDispatchDate, a.BLArrivalDate, a.ShipperID, a.ContainerDep, a.TotNoOfPieces, a.TotWeight, a.TotVolume, a.Consolidated " & _
					  "from BLs a, Pilots c where a.PilotID = c.PilotID and a.BLID=" & ObjectID
		QuerySelect2 = "select a.DiceContener, a.Weights, a.Volumes, a.NoOfPieces, a.CountriesFinalDes, a.Clients, a.BLs, a.DischargeDate, a.ClassNoOfPieces, a.CountryOrigen from BLDetail a where a.BLID=" & ObjectID
	case 13 'Manifiesto
    	select case BLType
		case 0 'Manifiesto Master
			QuerySelect = "select a.BLNumber, a.BLDispatchDate, b.Name, b.License, b.Countries, c.TruckNo, c.Mark, c.Model, d.CodProv, a.Bail, e.TruckNo, " & _
					  "a.SenderData, a.ConsignerData, a.CountryDes, a.TotNoOfPieces, a.TotWeight, a.Countries, a.FinalDes, a.Comment2 " & _
					  "from ((((BLs a left outer join Trucks e on a.Container = e.TruckID) " & _
					  "inner join  Pilots b on a.PilotID = b.PilotID) inner join Trucks c on a.TruckID = c.TruckID) " & _
					  "inner join Providers d on c.ProviderID=d.ProviderID) " & _
					  "where a.BLID = " & ObjectID					  
			QuerySelect2 = "select a.NoOfPieces, a.Weights, a.DiceContener, b.BLNumber, a.Agents, a.Clients, a.CountriesFinalDes, a.ClientsID, a.ClassNoOfPieces, a.AgentsID, a.Seps, a.HBLNumber, a.BLs, a.CountryOrigen, a.MBLs, a.ExType from BLDetail a, BLs b where b.BLID=a.BLID and a.BLID=" & ObjectID
		case 1,2 'Manifiesto Cliente 1=Consolidado, 2=Express
			QuerySelect = "select f.HBLNumber, a.BLDispatchDate, b.Name, b.License, b.Countries, c.TruckNo, c.Mark, c.Model, d.CodProv, a.Bail, e.TruckNo, " & _
					  "a.SenderData, a.ConsignerData, a.CountryDes, f.NoOfPieces, f.Weights, a.Countries, a.FinalDes, '', f.BLs, f.EXType, f.MBLs, f.ShippersID, f.ColoadersID, f.EXDBCountry " & _
					  "from (((((BLs a left outer join Trucks e on a.Container = e.TruckID) " & _
					  "inner join  Pilots b on a.PilotID = b.PilotID) inner join Trucks c on a.TruckID = c.TruckID) " & _
					  "inner join Providers d on c.ProviderID=d.ProviderID) " & _
					  "inner join BLDetail f on a.BLID=f.BLID) " & _
					  "where f.BLDetailID = " & ObjectID		  
					  '"from BLs a, Pilots b, Trucks c, Providers d, Trucks e " & _
					  '"where a.PilotID = b.PilotID and a.TruckID = c.TruckID and c.ProviderID=d.ProviderID and a.Container=e.TruckID and a.BLID = " & ObjectID
			'QuerySelect2 = "select a.NoOfPieces, a.Weights, a.DiceContener, b.BLNumber, a.Agents, a.Clients, a.CountriesFinalDes, a.Pos from BLDetail a, BLs b where b.BLID=a.BLID and a.BLDetailID=" & ObjectID
            QuerySelect2 = "select a.NoOfPieces, a.Weights, a.DiceContener, b.BLNumber, a.Agents, a.Clients, a.CountriesFinalDes, a.ClientsID, a.ClassNoOfPieces, a.AgentsID, a.Seps, a.HBLNumber, a.BLs, a.CountryOrigen, a.MBLs, a.ExType from BLDetail a, BLs b where b.BLID=a.BLID and a.BLID=" & BLID & " and a.ClientsID=" & ClientID & " and a.AgentsID=" & AgentID & " and a.Seps=" & Sep

		case 3 'Manifiesto Cliente 3=ItineraryAdds CIF
			QuerySelect = "select f.HBLNumber, '', '', '', f.Countries, '', '', '', '', '', '', " & _
					  "f.Agents, f.Clients, f.CountriesFinalDes, f.NoOfPieces, f.Weights, f.Countries, f.CountriesFinalDes, '' " & _
					  "from BLDetail f " & _
					  "where f.BLDetailID = " & ObjectID		  
					  '"from BLs a, Pilots b, Trucks c, Providers d, Trucks e " & _
					  '"where a.PilotID = b.PilotID and a.TruckID = c.TruckID and c.ProviderID=d.ProviderID and a.Container=e.TruckID and a.BLID = " & ObjectID
			'QuerySelect2 = "select a.NoOfPieces, a.Weights, a.DiceContener, b.BLNumber, a.Agents, a.Clients, a.CountriesFinalDes, a.Pos from BLDetail a, BLs b where b.BLID=a.BLID and a.BLDetailID=" & ObjectID
			
            QuerySelect2 = "select a.NoOfPieces, a.Weights, a.DiceContener, a.BLs, a.Agents, a.Clients, a.CountriesFinalDes, a.ClientsID, a.ClassNoOfPieces, a.AgentsID, a.Seps, a.HBLNumber, a.BLs, a.CountryOrigen, a.MBLs, a.ExType from BLDetail a where a.BLDetailID=" & ObjectID & " and a.ClientsID=" & ClientID & " and a.AgentsID=" & AgentID & " and a.Seps=" & Sep
        'case 2 'Manifiesto Cliente Express
		'	QuerySelect = "select f.HBLNumber, a.BLDispatchDate, b.Name, b.License, b.Countries, c.TruckNo, c.Mark, c.Model, d.CodProv, a.Bail, e.TruckNo, " & _
		'			  "a.SenderData, a.ConsignerData, a.CountryDes, f.NoOfPieces, f.Weights, a.Countries, a.FinalDes, '' " & _
		'			  "from (((((BLs a left outer join Trucks e on a.Container = e.TruckID) " & _
		'			  "inner join  Pilots b on a.PilotID = b.PilotID) inner join Trucks c on a.TruckID = c.TruckID) " & _
		'			  "inner join Providers d on c.ProviderID=d.ProviderID) " & _
		'			  "inner join BLDetail f on a.BLID=f.BLID) " & _
		'			  "where f.BLID = " & BLID		  
		'			  '"from BLs a, Pilots b, Trucks c, Providers d, Trucks e " & _
		'			  '"where a.PilotID = b.PilotID and a.TruckID = c.TruckID and c.ProviderID=d.ProviderID and a.Container=e.TruckID and a.BLID = " & ObjectID
		'	'QuerySelect2 = "select a.NoOfPieces, a.Weights, a.DiceContener, b.BLNumber, a.Agents, a.Clients, a.CountriesFinalDes, a.Pos from BLDetail a, BLs b where b.BLID=a.BLID and a.BLDetailID=" & ObjectID
		'	QuerySelect2 = "select a.NoOfPieces, a.Weights, a.DiceContener, b.BLNumber, a.Agents, a.Clients, a.CountriesFinalDes, a.ClientsID, a.ClassNoOfPieces, a.AgentsID, a.Seps, a.HBLNumber from BLDetail a, BLs b where b.BLID=a.BLID and a.BLID=" & BLID & " and a.ClientsID=" & ClientID & " and a.AgentsID=" & AgentID & " and a.Seps=" & Sep
		end select
	case 14, 15, 27 'Carta Endoso
		if BLType<>3 then
            QuerySelect = "select a.Logo, a.Footer, a.Name, a.Estate, b.CountryDes, c.HBLNumber, " & _
					  "c.Clients, c.LtEndorseDate, c.DTI, b.TotNoOfPieces, b.TotWeight, d.TruckNo, e.TruckNo, b.ContainerDep, " & _
                      "b.Chassis, c.Countries, c.CIFBrokerIn, c.BLs, c.Container, c.CIFLandFreight, c.ShippersID, c.EXDBCountry, c.MBLs, c.ExType " & _
					  "from ((((BLs b left outer join Trucks e on b.Container = e.TruckID) " & _
					  "inner join  Letters a on a.Countries=b.CountryDes) inner join BLDetail c on b.BLID=c.BLID) " & _
					  "inner join Trucks d on b.TruckID=d.TruckID) " & _
					  "where c.BLDetailID=" & ObjectID
					  '"from Letters a, BLs b, BLDetail c, Trucks d, Trucks e " & _
					  '"where a.Countries=b.CountryDes and b.BLID=c.BLID and b.TruckID=d.TruckID and b.Container=e.TruckID and c.BLDetailID=" & ObjectID
		    if GroupID <> 27 then
                QuerySelect2 = "select a.NoOfPieces, a.Weights, a.DiceContener, a.ClassNoOfPieces from BLDetail a where a.BLID=" & BLID & " and a.ClientsID=" & ClientID & " and a.AgentsID=" & AgentID & " and a.Seps=" & Sep
            else
                QuerySelect2 = "select a.NoOfPieces, a.Weights, a.DiceContener, a.ClassNoOfPieces from BLDetail a where a.BLDetailID=" & ObjectID & " and a.ClientsID=" & ClientID & " and a.AgentsID=" & AgentID & " and a.Seps=" & Sep
            end if
        else 
            QuerySelect = "select a.Logo, a.Footer, a.Name, a.Estate, c.CountriesFinalDes, c.HBLNumber, " & _
					  "c.Clients, c.LtEndorseDate, c.DTI, c.NoOfPieces, c.Weights, '', '', '', " & _
                      "'', c.Countries, c.CIFBrokerIn, c.BLs, c.Container, c.CIFLandFreight " & _
					  "from Letters a inner join BLDetail c on a.Countries=c.CountriesFinalDes " & _
					  "where c.BLDetailID=" & ObjectID
					  '"from Letters a, BLs b, BLDetail c, Trucks d, Trucks e " & _
					  '"where a.Countries=b.CountryDes and b.BLID=c.BLID and b.TruckID=d.TruckID and b.Container=e.TruckID and c.BLDetailID=" & ObjectID
            QuerySelect2 = "select a.NoOfPieces, a.Weights, a.DiceContener, a.ClassNoOfPieces from BLDetail a where a.BLDetailID=" & ObjectID & " and a.ClientsID=" & ClientID & " and a.AgentsID=" & AgentID & " and a.Seps=" & Sep
        end if
	case 16 'Carta Aceptacion
		if BLType >= 0 then
			QuerySelect = "select b.Name, c.Name, d.TruckNo, a.ShipperID, f.Name, a.Consolidated, " & _
					  "a.BLNumber, a.LtAcceptDate, a.CountryDes, a.TotNoOfPieces, a.TotWeight " & _
					  "from BLs a, Brokers b, Brokers c, Trucks d, Pilots f " & _
					  "where b.BrokerID = a.BrokerRecepID and " & _
					  "c.BrokerID = a.BrokerID and " & _
					  "d.TruckID = a.TruckID and " & _
					  "f.PilotID = a.PilotID and a.BLID=" & ObjectID
			QuerySelect2 = "select a.NoOfPieces, a.Weights, a.BLs, a.Clients, a.ClassNoOfPieces from BLDetail a where a.BLID=" & ObjectID
			QuerySelect3 = "select a.Logo, a.Footer, a.Name from Letters a, BLs b where a.Countries=b.CountryDes and b.BLID=" & ObjectID
		else
			QuerySelect = "select b.Name, c.Name, d.TruckNo, a.ShipperID, f.Name, a.Consolidated, " & _
					  "g.BLNumber, g.LtAcceptDate, g.CountryDes " & _
					  "from BLs a, Brokers b, Brokers c, Trucks d, Pilots f, BLGroups g, BLGroupDetail h " & _
					  "where b.BrokerID = a.BrokerRecepID and " & _
					  "c.BrokerID = a.BrokerID and " & _
					  "d.TruckID = a.TruckID and " & _
					  "f.PilotID = a.PilotID and a.BLID=h.BLID and g.BLGroupID=h.BLGroupID and h.BLGroupID=" & ObjectID
			QuerySelect2 = "select a.NoOfPieces, a.Weights, a.BLs, a.Clients, a.ClassNoOfPieces from BLDetail a, BLs b, BLGroups c, BLGroupDetail d where a.BLID=b.BLID and b.BLID=d.BLID and c.BLGroupID=d.BLGroupID and d.BLGroupID=" & ObjectID
			QuerySelect3 = "select a.Logo, a.Footer, a.Name from Letters a, BLGroups b where a.Countries=b.CountryDes and b.BLGroupID=" & ObjectID
		end if
	case 22 'Manifiesto Grupo
		QuerySelect = "select f.BLNumber, a.BLDispatchDate, b.Name, b.License, b.Countries, c.TruckNo, c.Mark, c.Model, d.CodProv, a.Bail, g.TruckNo, " & _
					  "a.SenderData, a.ConsignerData, f.CountryDes, a.TotNoOfPieces, a.TotWeight, f.Countries, a.FinalDes " & _
					  "from ((((((BLs a left outer join Trucks g on a.Container = g.TruckID) " & _
					  "inner join Pilots b on a.PilotID = b.PilotID) inner join Trucks c on a.TruckID = c.TruckID) " & _
					  "inner join Providers d on c.ProviderID=d.ProviderID) inner join BLGroupDetail e on a.BLID=e.BLID) " & _
					  "inner join BLGroups f on e.BLGroupID=f.BLGroupID) " & _
					  "where e.BLGroupID=" & ObjectID
					  '"from BLs a, Pilots b, Trucks c, Providers d, BLGroupDetail e, BLGroups f, Trucks g " & _
					  '"where a.PilotID = b.PilotID and a.TruckID = c.TruckID and c.ProviderID=d.ProviderID and e.BLID=a.BLID and f.BLGroupID=e.BLGroupID and a.Container=g.TruckID and e.BLGroupID=" & ObjectID
		QuerySelect2 = "select a.NoOfPieces, a.Weights, a.DiceContener, a.HBLNumber, a.Agents, a.Clients, a.CountriesFinalDes, b.BLNumber, b.ConsignerData, b.CountryDes, a.ClassNoOfPieces, a.CountryOrigen, BLs, MBLs from BLDetail a, BLs b, BLGroupDetail c where b.BLID=a.BLID and c.BLID=a.BLID and c.BLGroupID=" & ObjectID & " order by BLNumber, Pos"
	    QuerySelect3 = "select distinct b.Comment2 from BLs b, BLGroupDetail c where b.BLID=c.BLID and c.BLGroupID=" & ObjectID & " order by b.BLNumber"
    case 24 'Itinerario Grupo
		QuerySelect = "select b.BLNumber, c.Name, a.Week, a.BLDispatchDate, a.BLArrivalDate, a.ShipperID, a.ContainerDep, sum(a.TotNoOfPieces), sum(a.TotWeight), sum(a.TotVolume), a.Consolidated " & _
					  "from BLs a, BLGroups b, Pilots c,  BLGroupDetail d where a.PilotID=c.PilotID and a.BLID=d.BLID and b.BLGroupID=d.BLGroupID and d.BLGroupID=" & ObjectID & " group by d.BLGroupID"
		QuerySelect2 = "select a.DiceContener, a.Weights, a.Volumes, a.NoOfPieces, a.CountriesFinalDes, a.Clients, a.BLs, a.DischargeDate, a.ClassNoOfPieces, a.CountryOrigen from BLDetail a, BLGroupDetail b, BLs c where a.BLID=b.BLID and a.BLID=c.BLID and b.BLGroupID=" & ObjectID & " order by a.CountriesFinalDes, c.BLNumber, a.Pos"
	end select

	'response.write QuerySelect & "<br>"
	'response.write QuerySelect2 & "<br>"
	OpenConn Conn
	Set rs = Conn.Execute(QuerySelect)
	If Not rs.EOF Then
    	aTableValues = rs.GetRows
    	CountTableValues = rs.RecordCount-1
	End If
	CloseOBJ rs
	
	if QuerySelect2 <> "" then
		Set rs = Conn.Execute(QuerySelect2)
		If Not rs.EOF Then
			aTableValues2 = rs.GetRows
			CountTableValues2 = rs.RecordCount-1
		End If
		CloseOBJ rs
	end if	

	if QuerySelect3 <> "" then
		Set rs = Conn.Execute(QuerySelect3)
		If Not rs.EOF Then
			aTableValues3 = rs.GetRows
			CountTableValues3 = rs.RecordCount-1
		End If
		CloseOBJ rs
	end if
	closeOBJ Conn	
		
	if CountTableValues >= 0 then
		Select case GroupID
		case 12, 24
			BlNumber = aTableValues(0, 0)
			PilotName = aTableValues(1, 0)
			Week = aTableValues(2, 0)
			BLDispatchDate = aTableValues(3, 0)
			BLArrivalDate = aTableValues(4, 0)
			OpenConn2 Conn
			set rs = Conn.Execute("select nombre_cliente from clientes where es_shipper=true and id_cliente=" & aTableValues(5, 0))
			if Not rs.EOF then
				ShipperName = rs(0)
			Else
				ShipperName = ""
			End if
			CloseOBJs rs, Conn
			ContainerDep = aTableValues(6, 0)
			TotNoOfPieces = aTableValues(7, 0)
			TotWeight = aTableValues(8, 0)
			TotVolume = aTableValues(9, 0)
			Consolidated = aTableValues(10, 0)
		case 13
			if BLNumber = "" then 'No es vacio cuando proviene del grupo 22
				BLNumber = aTableValues(0, 0)
				'if BLType > 0 then 'Cliente
				'	BLNumber = BLNumber & "-" & FiveDigits(ClientID+AgentID+Sep)
				'end if 
			end if
			BLDispatchDate = aTableValues(1, 0)
			PilotName = aTableValues(2, 0)
			License = aTableValues(3, 0)
			Nacionality = aTableValues(4, 0)
			TruckNo = aTableValues(5, 0)
			Mark = aTableValues(6, 0)
			Model = aTableValues(7, 0)
			CodProv = aTableValues(8, 0)
			Bail = aTableValues(9, 0)
			Container = aTableValues(10, 0)
			SenderData = FRegExp(ntr, aTableValues(11,0), "<br>", 4)
			ConsignerData = FRegExp(ntr, aTableValues(12,0), "<br>", 4)
			CountryDes = aTableValues(13, 0)
			'TotNoOfPieces = aTableValues(14, 0)
			'TotWeight = aTableValues(15, 0)
			Countries = aTableValues(16, 0)
			FinalDes = aTableValues(17, 0)
			ManifComment = aTableValues(18, 0)	
            'Caso especial de MX, no debe mostrar el numero de CP nuestro, sino el de MX, y si tiene RO muestra el dato que se guarda en MBLs
            select Case BLType
            Case 1,2
                ShipperID = aTableValues(22, 0)	
                if FRegExp(PtrnViewBLAgents, ShipperID,  "", 2) then
                'if InStr(1,ValidPatrnCIF,Mid(BLNumber,2,2))>0 then
                'if InStr(1,ValidPatrnCIF,Countries)>0 then
                    Select Case CheckNum(aTableValues(20, 0))
                    Case 4,5
                        BLNumber = aTableValues(21, 0)
                    Case Else
                        BLNumber = aTableValues(19, 0)	
                    End Select                    
                end if
                ColoaderID = aTableValues(23, 0)	
                if ColoaderID <> 0  then
                    IsColoader = 1
                end if
                EXDBCountry = aTableValues(24, 0)
            End Select            

		case 14, 15, 27
			Logo = aTableValues(0, 0)
			Footer = FRegExp(chr(13) & chr(10), aTableValues(1, 0), "<br>", 4)
			BusinessName = aTableValues(2, 0)
			Estate = aTableValues(3, 0)
			CountryDes = aTableValues(4, 0)
			BLNumber = aTableValues(5, 0)
			ConsignerData = aTableValues(6, 0)
			LtEndorseDate = aTableValues(7, 0)
			DTI = aTableValues(8, 0)
			TotNoOfPieces = aTableValues(9, 0)
			TotWeight = aTableValues(10, 0)
			TruckNo = aTableValues(11, 0)
			Container = aTableValues(12, 0)
			ContainerDep = aTableValues(13, 0)
			Chassis = aTableValues(14, 0)
			Countries = aTableValues(15, 0)
            CIFBrokerIn = aTableValues(16, 0)
            BL = aTableValues(17, 0)
            Contener = aTableValues(18, 0)
            CIFLandFreight = aTableValues(19, 0)
            ShipperID = aTableValues(20, 0)
            EXDBCountry = aTableValues(21, 0)
            MBL = aTableValues(22, 0)
            ExType = aTableValues(23, 0)
		case 16
			BrokerRecepName = aTableValues(0, 0)
			BrokerName = aTableValues(1, 0)
			TruckNo = aTableValues(2, 0)
			OpenConn2 Conn
				set rs = Conn.Execute("select nombre_cliente from clientes where es_shipper=true and id_cliente=" & aTableValues(3, 0))
				if Not rs.EOF then
					ShipperName = rs(0)
				Else
					ShipperName = ""
				End if
				CloseOBJ rs
			
				set rs = Conn.Execute("select nombres from contactos where id_cliente=" & aTableValues(3, 0))
				if Not rs.EOF then
					Attn = rs(0)
				else
					Attn = ""
				end if
			CloseOBJs rs, Conn
			PilotName = aTableValues(4, 0)
			Consolidated = aTableValues(5, 0)
			BLNumber = aTableValues(6, 0)
			LtAcceptDate = aTableValues(7, 0)
			CountryDes = aTableValues(8, 0)
			if BLType >= 0 then
				TotNoOfPieces = aTableValues(9, 0)
				TotWeight = aTableValues(10, 0)
			end if
		case 22
			BLNumber = aTableValues(0, 0)
			BLDispatchDate = aTableValues(1, 0)
			PilotName = aTableValues(2, 0)
			License = aTableValues(3, 0)
			Nacionality = aTableValues(4, 0)
			TruckNo = aTableValues(5, 0)
			Mark = aTableValues(6, 0)
			Model = aTableValues(7, 0)
			CodProv = aTableValues(8, 0)
			Bail = aTableValues(9, 0)
			Container = aTableValues(10, 0)
			SenderData = FRegExp(ntr, aTableValues(11,0), "<br>", 4)
			ConsignerData = FRegExp(ntr, aTableValues(12,0), "<br>", 4)
			CountryDes = aTableValues(13, 0)
			'TotNoOfPieces = 0
			'TotWeight = 0
			'for i=0 to CountTableValues
			'	TotNoOfPieces = TotNoOfPieces + Cdbl(aTableValues(14, i))
			'	TotWeight = TotWeight + Cdbl(aTableValues(15, i))
			'next 
			Countries = aTableValues(16, 0)
			FinalDes = aTableValues(17, 0)			
		end select
	end if
	
	if CountTableValues2 >= 0 then
		Select case GroupID
		case 12, 24		
			Redim Preserve DiceContener(CountTableValues2)
			Redim Preserve Weights(CountTableValues2)
			Redim Preserve Volumes(CountTableValues2)
			Redim Preserve Clients(CountTableValues2)
			Redim Preserve NoOfPieces(CountTableValues2)
			Redim Preserve CountriesFinalDes(CountTableValues2)
			Redim Preserve BLs(CountTableValues2)
			Redim Preserve DischargeDate(CountTableValues2)
			Redim Preserve ClassNoOfPieces(CountTableValues2)
			Redim Preserve CountryOrigen(CountTableValues2)

			for i=0 to CountTableValues2
				DiceContener(i) = aTableValues2(0,i)
				Weights(i) = aTableValues2(1,i)
				Volumes(i) = aTableValues2(2,i)
				NoOfPieces(i) = aTableValues2(3,i)
				CountriesFinalDes(i) = aTableValues2(4,i)
				if Consolidated = 1 then
					Clients(i) = aTableValues2(5,i)
					BLs(i) = aTableValues2(6,i)
					DischargeDate(i) = aTableValues2(7,i)
				else
					Clients(0) = aTableValues2(5,i)
					BLs(0) = aTableValues2(6,i)
					DischargeDate(0) = aTableValues2(7,i)
				end if
				ClassNoOfPieces(i) = aTableValues2(8,i)
				CountryOrigen(i) = aTableValues2(9,i)
			next
		case 13, 22
			Redim Preserve BLs(CountTableValues2)
			Redim Preserve ExBLs(CountTableValues2)
			Redim Preserve MBLs(CountTableValues2)
			Redim Preserve Agents(CountTableValues2)
			Redim Preserve Clients(CountTableValues2)
			Redim Preserve CountriesFinalDes(CountTableValues2)
			Redim Preserve Pos(CountTableValues2)
			Redim Preserve NoOfPieces(CountTableValues2)
			Redim Preserve Weights(CountTableValues2)
			Redim Preserve DiceContener(CountTableValues2)
			Redim Preserve Consigners(CountTableValues2)
			Redim Preserve CountriesDes(CountTableValues2)
			Redim Preserve ClassNoOfPieces(CountTableValues2)
			Redim Preserve CountryOrigen(CountTableValues2)

			for i=0 to CountTableValues2
				BLs(i) = aTableValues2(3,i)
				if GroupID=13 then
					'BLs(i) = BLNumber 'En el manifiesto Hijo mostraba el dato BLs de la tabla BLDetail, se elimino porque ahora debe mostrar la CP Hija
					'if BLType=0 then
						'BLs(i) = BLs(i) & "-" & FiveDigits(aTableValues2(7,i)+aTableValues2(9,i)+aTableValues2(10,i))
						BLs(i) = aTableValues2(11,i)
					'end if

                    if Mid(aTableValues2(11,i),2,2) = "MX" then
                    'if aTableValues2(13,i) = "MX" then
                        Select Case CheckNum(aTableValues2(15,i))
                        Case 4,5
                            ExBLs(i) = aTableValues2(14,i) 'MBLs
                            MBLs(i) = aTableValues2(12,i) 'BLs
                        Case Else
                            ExBLs(i) = aTableValues2(12,i) 'BLs
                            MBLs(i) = aTableValues2(14,i) 'MBLs
                        end Select
                    else
                        ExBLs(i) = aTableValues2(12,i) 'BLs
                        MBLs(i) = aTableValues2(14,i) 'MBLs
                    end if
					
                    'ExBLs(i) = aTableValues2(12,i)
                    ClassNoOfPieces(i) = aTableValues2(8,i)
					CountryOrigen(i) = aTableValues2(13,i)
                    'MBLs(i) = aTableValues2(14,i)
				end if
				Agents(i) = aTableValues2(4,i)
				Clients(i) = aTableValues2(5,i)
				CountriesFinalDes(i) = aTableValues2(6,i)
				if GroupID = 22 then
					Consig = Split(aTableValues2(8,i),chr(13)&chr(10))
					Consigners(i) = Consig(0)
					CountriesDes(i) = aTableValues2(9,i)
					ClassNoOfPieces(i) = aTableValues2(10,i)
					CountryOrigen(i) = aTableValues2(11,i)
                    ExBLs(i) = aTableValues2(12,i)
                    MBLs(i) = aTableValues2(13,i)
				end if
				NoOfPieces(i) = aTableValues2(0,i)
				Weights(i) = aTableValues2(1,i)
				DiceContener(i) = aTableValues2(2,i)
				TotNoOfPieces = TotNoOfPieces + aTableValues2(0,i)
				TotWeight = TotWeight + aTableValues2(1,i)
			next
		case 14, 15, 27
			Redim Preserve NoOfPieces(CountTableValues2)
			Redim Preserve Weights(CountTableValues2)
			Redim Preserve DiceContener(CountTableValues2)
			Redim Preserve ClassNoOfPieces(CountTableValues2)
	
			for i=0 to CountTableValues2
				NoOfPieces(i) = aTableValues2(0,i)
				Weights(i) = aTableValues2(1,i)
				DiceContener(i) = aTableValues2(2,i)
				ClassNoOfPieces(i) = aTableValues2(3,i)
			next
		case 16
			Redim Preserve NoOfPieces(CountTableValues2)
			Redim Preserve Weights(CountTableValues2)
			Redim Preserve BLs(CountTableValues2)
			Redim Preserve Clients(CountTableValues2)
			Redim Preserve ClassNoOfPieces(CountTableValues2)
	
			for i=0 to CountTableValues2
				NoOfPieces(i) = aTableValues2(0,i)
				Weights(i) = aTableValues2(1,i)
				if Consolidated = 1 then
					BLs(i) = aTableValues2(2,i)
					Clients(i) = aTableValues2(3,i)
				else
					BLs(0) = aTableValues2(2,i)
					Clients(0) = aTableValues2(3,i)
				end if
				if BLType < 0 then 'Grupo = -1
					TotNoOfPieces = TotNoOfPieces + NoOfPieces(i)
					TotWeight = TotWeight + Weights(i)
				end if				
				ClassNoOfPieces(i) = aTableValues2(4,i)
			next
		end select
	end if
	
	if CountTableValues3>=0 then
		if GroupID<>22 then
            Logo = aTableValues3(0, 0)
		    Footer = FRegExp(chr(13) & chr(10), aTableValues3(1, 0), "<br>", 4)
		    BusinessName = aTableValues3(2, 0)
        else
            for i=0 to CountTableValues3
                ManifComment = ManifComment & aTableValues3(0,i) & "<br>"
            next
        end if
	end if

    'Se toma primero el pais de la base de datos para desplegar el logo, si viene vacio se toma el pais donde se crea el registro
    if EXDBCountry = "" then
        EXDBCountry = Countries
    end if
	
	Set aTableValues = Nothing
	Set aTableValues2 = Nothing
	Set aTableValues3 = Nothing
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
	font-size:<%if GroupID=13 or GroupID=22 then%>10<%else%>10<%end if%>px;
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
	font-size: <%if GroupID=13 or GroupID=22 then%>10<%else%>10<%end if%>px; 
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
<%select Case GroupID
  Case 12, 24%>
<table width="641" cellpadding="2" cellspacing="0" align="center">
  <tr>
	<td class="style11" align="left" width="50%"><%=DisplayLogo(EXDBCountry, 0, 0, 0)%></td>
	<td class="style3" align="right">EDICION 1 </td>
  </tr>
</table>
<table width="641" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
	<td class="style11" align="left"><font color="#0000FF">ITINERARIO&nbsp;DE&nbsp;CARGA:&nbsp;<%=BLNumber%></font></td>
  </tr>
</table>
<table width="641" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
    <td class="style4" align="left" colspan="5" valign="top">Fecha de Salida:<br>
      <span class="style10"><%=BLDispatchDate%></span></td>
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
		<td class="style4" align="center" valign="middle">Clase Bultos</td>
		<td class="style4" align="center" valign="middle">Descripci&oacute;n de Carga</td>
		<td class="style4" align="center" valign="middle">Volumen<br>(CBM)</td>
		<td class="style4" align="center" valign="middle">Peso&nbsp;Bruto<br>(Kg)</td>
		<td class="style4" align="center" valign="middle">Consignee<br>(Consolidado)</td>
		<td class="style4" align="center" valign="middle">Fecha<br>Descarga</td>
	</tr>
	<%for i=0 to CountTableValues2%>
		<%if CountryTitle <> CountriesFinalDes(i) then
			CountryTitle = CountriesFinalDes(i)
			if SubTotNoOfPieces > 0 then
		%>
				<tr height="8">
					<td class="style4" align="right" valign="top"><b><%=SubTotNoOfPieces%></b></td>
					<td class="style4" align="center" valign="top" colspan="2"><b>SUBTOTALES</b></td>
					<td class="style4" align="right" valign="top"><b><%=SubTotVolume%></b></td>
					<td class="style4" align="right" valign="top"><b><%=SubTotWeight%></b></td>
					<td class="style4" align="right" valign="top" colspan="2">&nbsp;</td>
				</tr>
		<%		SubTotNoOfPieces = 0
				SubTotWeight = 0
				SubTotVolume = 0
			end if%>
			<tr>
				<td class="style4" align="left" valign="top" colspan="7"><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<span class="style10"><b><%=TranslateCountry(CountryTitle)%></b></span></td>
			</tr>
		<%end if%>
	<tr>
		<td class="style4" align="right" valign="top"><span class="style10"><%=NoOfPieces(i)%></span></td>
		<td class="style4" align="left" valign="top"><span class="style10"><%=ClassNoOfPieces(i)%></span></td>
		<td class="style4" align="left" valign="top"><span class="style10"><%=DiceContener(i)%></span></td>
		<td class="style4" align="right" valign="top"><span class="style10"><%=Volumes(i)%></span></td>
		<td class="style4" align="right" valign="top"><span class="style10"><%=Weights(i)%></span></td>
		<td class="style4" align="left" valign="top"><span class="style10"><%=Clients(i)%></span></td>
		<td class="style4" align="right" valign="top"><span class="style10"><%=DischargeDate(i)%></span></td>		
	</tr>
	<%	SubTotNoOfPieces = SubTotNoOfPieces + NoOfPieces(i)*1
		SubTotWeight = SubTotWeight + Weights(i)*1
		SubTotVolume = SubTotVolume + Volumes(i)*1
	next%>
	<tr height="8">
		<td class="style4" align="right" valign="top"><b><%=SubTotNoOfPieces%></b></td>
		<td class="style4" align="center" valign="top" colspan="2"><b>SUBTOTALES</b></td>
		<td class="style4" align="right" valign="top"><b><%=SubTotWeight%></b></td>
		<td class="style4" align="right" valign="top"><b><%=SubTotVolume%></b></td>
		<td class="style4" align="right" valign="top" colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="style4" align="right" valign="top" colspan="7" height="100%">&nbsp;</td>
	</tr>
	<tr height="8">
		<td class="style4" align="right" valign="top"><b><%=TotNoOfPieces%></b></td>
		<td class="style4" align="center" valign="top" colspan="2"><b>TOTALES</b></td>
		<td class="style4" align="right" valign="top"><b><%=TotWeight%></b></td>
		<td class="style4" align="right" valign="top"><b><%=TotVolume%></b></td>
		<td class="style4" align="right" valign="top" colspan="2">&nbsp;</td>
	</tr>
</table>

<%Case 13, 22%>
<table width="641" cellpadding="2" cellspacing="0" align="center">
  <tr>
	<td class="style11" align="left" width="50%">
    <%
    if FRegExp(PtrnLogo, ShipperID,  "", 2) or (ColoaderID > 0) then
		response.write DisplayLogo(ShipperID, ColoaderID, IsColoader, EXDBCountry)
	else
		response.write DisplayLogo(EXDBCountry, 0, 0, 0)
	end if	
	%>
    </td>
	<td class="style3" align="right">EDICION 1</td>
  </tr>
</table>
<% if BLType=3 then%>
<table width="641" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
	<td class="style11" align="left"><font color="#0000FF">MANIFIESTO&nbsp;DE&nbsp;CARGA:</font></td>
  </tr>
</table>
<% else%>
<table width="641" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
	<td class="style11" align="left"><font color="#0000FF">MANIFIESTO&nbsp;DE&nbsp;CARGA:&nbsp;MAN<%=BLNumber%></font></td>
  </tr>
</table>
<table width="641" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
	<td class="style4" align="left" colspan="5" valign="top">Fecha Salida:<br><span class="style10"><%=BLDispatchDate%></span></td>
	<td class="style4" align="left" colspan="5" valign="top">Carta Porte:<br><span class="style10"><%=BLNumber%></span></td>
	<td class="style4" align="left" colspan="5" valign="top">Destino:<br><span class="style10"><%=TranslateCountry(CountryDes)%></span></td>
  </tr>
</table>
<table width="641" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
	<td class="style4" align="left" colspan="5" valign="top">Conductor:<br><span class="style10"><%=PilotName%></span></td>
	<td class="style4" align="left" colspan="5" valign="top">Licencia:<br><span class="style10"><%=License%></span></td>
	<td class="style4" align="left" colspan="4" valign="top">Nacionalidad:<br>
	  <span class="style10"><%=TranslateCountry(Nacionality)%></span></td>
  </tr>
</table>
<table width="641" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
	<td class="style4" align="left" colspan="5" valign="top">Placas Cami&oacute;n:<br>
	  <span class="style10"><%=TruckNo%></span></td>
	<td class="style4" align="left" colspan="5" valign="top">Marca:<br>
	<span class="style10"><%=Mark%></span></td>
	<td class="style4" align="left" colspan="5" valign="top">Modelo:<br>
	  <span class="style10"><%=Model%></span></td>
  </tr>
  <tr>
	<td class="style4" align="left" colspan="5" valign="top">C&oacute;digo:<br>
	<span class="style10"><%=CodProv%></span></td>
	<td class="style4" align="left" colspan="5" valign="top"><%if Countries="GT" then%>Fianza<%else%>Marchamo<%end if%>:<br>
	  <span class="style10"><%=Bail%></span></td>
	<td class="style4" align="left" colspan="5" valign="top">Furg&oacute;n:<br>
	  <span class="style10"><%=Container%></span></td>
  </tr>
</table>
<%end if %>
<%if Countries="MX" or Countries="GT" or Countries="SV" or Countries="HN" or Countries="NI" or Countries="N1" or Countries="CR" or Countries="PA" or Countries="A2" or Countries="BZ" then%>
	<table width="641" height="250" class="styleborder" cellpadding="2" cellspacing="0" align="center">
		<tr height="8">
			<%if BLType<>3 then%>
            <td class="style4" align="center" valign="middle" width="25%">CP</td>
            <td class="style4" align="center" valign="middle" width="20%">BL/RO</td>
            <td class="style4" align="center" valign="middle" width="20%"><%if Countries="PA" then %>DMC<%else%>MBL<%end if%></td>
			<%else %>
            <td class="style4" align="center" valign="middle" width="20%" colspan=3>BL/RO</td>
            <%end if %>
			<%if Countries="CR" or Countries="NI" or Countries="N1" then%>
			<td class="style4" align="center" valign="middle">Exportador</td>
			<%end if%>
			<td class="style4" align="center" valign="middle">Consignatario</td>
			<td class="style4" align="center" valign="middle">Origen</td>
			<%if BLType<>3 then%>
                <%if CountryDes<>"PA" then%>
                <td class="style4" align="center" valign="middle">Procedencia</td>
                <%end if %>
            <%end if %>
			<td class="style4" align="center" valign="middle">Destino</td>
			<td class="style4" align="center" valign="middle">No. Bultos</td>
			<td class="style4" align="center" valign="middle">Clase Bultos</td>
			<td class="style4" align="center" valign="middle">Descripci&oacute;n de Carga</td>
			<td class="style4" align="center" valign="middle">Peso Bruto</td>
		</tr>
		<%
		if CountTableValues2>=0 then
			BrokerName = CountriesFinalDes(0)
		end if		
		for i=0 to CountTableValues2
			if BrokerName <> CountriesFinalDes(i) then%>
			<tr>
				<td class="style4" align="right" valign="top" colspan="<%if Countries="CR" then%>8<%else%>7<%end if%>"><span class="style10">SUBTOTAL</span></td>
				<td class="style4" align="right" valign="top"><span class="style10"><%=SubTotNoOfPieces%></span></td>
				<td class="style4" align="right" valign="top" colspan="2"><span class="style10">SUBTOTAL</span></td>
				<td class="style4" align="right" valign="top"><span class="style10"><%=SubTotWeight%></span></td>
			</tr>
			<tr>
				<td class="style4" align="left" valign="top" colspan="9"><span class="style10">&nbsp;</span></td>
			</tr>
			<%  BrokerName=CountriesFinalDes(i)
				SubTotNoOfPieces = 0
				SubTotWeight = 0
			 end if
			 SubTotNoOfPieces = SubTotNoOfPieces + NoOfPieces(i)
			 SubTotWeight = SubTotWeight + Weights(i)
		%>
		<tr>
            <%if BLType<>3 then%>
                <%if Mid(BLs(i),2,2)="MX" then 'if CountryOrigen(i)="MX" then %>
			    <td class="style4" align="left" valign="top"><span class="style10"><%=ExBLs(i)%></span></td>
                <td class="style4" align="left" valign="top"><span class="style10"><%=BLs(i)%></span></td>
                <%else %>
                <td class="style4" align="left" valign="top"><span class="style10"><%=BLs(i)%></span></td>
                <td class="style4" align="left" valign="top"><span class="style10"><%=ExBLs(i)%></span></td>
                <%end if %>
                <td class="style4" align="left" valign="top"><span class="style10"><%=MBLs(i)%></span></td>
			<%else %>
            <td class="style4" align="left" valign="top" colspan=3><span class="style10"><%=ExBLs(i)%></span></td>
			<%end if %>
			<%if Countries="CR" or Countries="NI" or Countries="N1" then%>
			<td class="style4" align="left" valign="top"><span class="style10"><%=Agents(i)%></span></td>
			<%end if%>
			<td class="style4" align="left" valign="top"><span class="style10"><%=Clients(i)%></span></td>
			<td class="style4" align="left" valign="top"><span class="style10"><%=CountryOrigen(i)%></span></td>
			<%if BLType<>3 then%>
                <%if CountryDes<>"PA" then%>
                <td class="style4" align="left" valign="top"><span class="style10"><%=Countries%></span></td>
			    <!--<td class="style4" align="left" valign="top"><span class="style10"><%=Mid(BLs(i),2,2)%></span></td>-->
			    <%end if%>
            <%end if%>
            <td class="style4" align="left" valign="top"><span class="style10"><%=CountriesFinalDes(i)%></span></td>
			<td class="style4" align="right" valign="top"><span class="style10"><%=NoOfPieces(i)%></span></td>
			<td class="style4" align="left" valign="top"><span class="style10"><%=ClassNoOfPieces(i)%></span></td>
			<td class="style4" align="left" valign="top"><span class="style10"><%=DiceContener(i)%></span></td>
			<td class="style4" align="right" valign="top"><span class="style10"><%=Weights(i)%></span></td>
		</tr>
		<%next%>
		<tr>
			<td class="style4" align="right" valign="top" colspan="<%if Countries="CR" then%>8<%else%>7<%end if%>"><span class="style10">SUBTOTAL</span></td>
			<td class="style4" align="right" valign="top"><span class="style10"><%=SubTotNoOfPieces%></span></td>
			<td class="style4" align="right" valign="top" colspan="2"><span class="style10">SUBTOTAL</span></td>
			<td class="style4" align="right" valign="top"><span class="style10"><%=SubTotWeight%></span></td>
		</tr>
		<tr>
			<td class="style4" align="right" valign="top" colspan="<%if Countries="CR" then%>11<%else%>10<%end if%>" height="100%">&nbsp;</td>
		</tr>
		<tr height="8">
			<td class="style4" align="right" valign="top" colspan="<%if Countries="CR" then%>8<%else%>7<%end if%>">&nbsp;</td>
			<td class="style4" align="right" valign="top"><b><%=TotNoOfPieces%></b></td>
			<td class="style4" align="center" valign="top" colspan="2"><b>TOTALES</b></td>
			<td class="style4" align="right" valign="top"><b><%=TotWeight%></b></td>
		</tr>
	</table>
<%else%>
	<table width="641" class="styleborder" cellpadding="2" cellspacing="0" align="center">
	  <tr>
		<td width="50%" class="style4" align="left" valign="top">Exporter / Exportador:<br>
		  <span class="style10"><%=SenderData%></span></td>
		<td width="50%" class="style4" align="left" valign="top">Consignee / Consignatario:<br>
		<span class="style10"><%=ConsignerData%></span></td>
	  </tr>
	</table>
	<table width="641" height="250" class="styleborder" cellpadding="2" cellspacing="0" align="center">
		<tr height="8">
			<td class="style4" align="center" valign="middle">No. Bultos</td>
			<td class="style4" align="center" valign="middle">Clase Bultos</td>
			<td class="style4" align="center" valign="middle">Descripci&oacute;n de Carga</td>
			<td class="style4" align="center" valign="middle">Peso Bruto</td>
		</tr>
		<%
		if CountTableValues2>=0 then
			BrokerName = CountriesFinalDes(0)
		end if		
		for i=0 to CountTableValues2
			if BrokerName <> CountriesFinalDes(i) then%>
			<tr>
				<td class="style4" align="right" valign="top"><span class="style10"><%=SubTotNoOfPieces%></span></td>
				<td class="style4" align="right" valign="top" colspan="2"><span class="style10">SUBTOTAL</span></td>
				<td class="style4" align="right" valign="top"><span class="style10"><%=SubTotWeight%></span></td>
			</tr>
			<tr>
				<td class="style4" align="left" valign="top" colspan="8"><span class="style10">&nbsp;</span></td>
			</tr>
			<%  BrokerName=CountriesFinalDes(i)		
				SubTotNoOfPieces = 0
				SubTotWeight = 0
			 end if
			 SubTotNoOfPieces = SubTotNoOfPieces + NoOfPieces(i)
			 SubTotWeight = SubTotWeight + Weights(i)
		%>
		<tr>
			<td class="style4" align="right" valign="top"><span class="style10"><%=NoOfPieces(i)%></span></td>
			<td class="style4" align="left" valign="top"><span class="style10"><%=ClassNoOfPieces(i)%></span></td>
			<td class="style4" align="left" valign="top"><span class="style10"><%=DiceContener(i)%></span></td>
			<td class="style4" align="right" valign="top"><span class="style10"><%=Weights(i)%></span></td>
		</tr>
		<%next%>
		<tr>
			<td class="style4" align="right" valign="top"><span class="style10"><%=SubTotNoOfPieces%></span></td>
			<td class="style4" align="right" valign="top" colspan="2"><span class="style10">SUBTOTAL</span></td>
			<td class="style4" align="right" valign="top"><span class="style10"><%=SubTotWeight%></span></td>
		</tr>
		<tr>
			<td class="style4" align="right" valign="top" colspan="4" height="100%">&nbsp;</td>
		</tr>
		<tr height="8">
			<td class="style4" align="right" valign="top"><b><%=TotNoOfPieces%></b></td>
			<td class="style4" align="center" valign="top" colspan="2"><b>TOTALES</b></td>
			<td class="style4" align="right" valign="top"><b><%=TotWeight%></b></td>
		</tr>
	</table>
<%end if%>
<br>
<table width="641" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
	<td class="style4" align="left" valign="top">Lugar de Emisi&oacute;n :<br>
	  <span class="style10"><%=TranslateCountry(Countries)%></span></td>
	<td class="style4" align="left" valign="top">Destino Final:<br>
	<span class="style10"><%=FinalDes%></span></td>
  </tr>
  <tr>
	<td class="style4" align="left" valign="top" colspan="2">Observaciones :<br><span class="style10"><%=ManifComment%></span></td>
  </tr>
</table>
<%Case 14, 15, 27%>

<table width="641" cellpadding="2" cellspacing="0" align="center">
  <tr>
	<td class="style11" align="left" width="50%">
	<%=DisplayLogo(EXDBCountry, 0, 0, 0)%>
	<br><br></td>
	<td class="style3" align="right">EDICION 1 </td>
  </tr>
</table>
<table width="641" class="styleborder" cellpadding="2" cellspacing="0" align="center">
  <tr>
	<td class="style11" align="left">CARTA&nbsp;DE&nbsp;ENDOSO&nbsp;ADUANAL&nbsp;DE&nbsp;MERCADERIA</td>
  </tr>
</table>
<table width="641" cellpadding="4" cellspacing="0" align="center">
  <tr>
	<td class="style10" align="left"><%if LtEndorseDate <> "" then response.write TranslateCountry(CountryDes) & " " & ConvertDate(LtEndorseDate,4) end if%><br></td>
  </tr>
  <tr>
	<td class="style10" align="left"><font color="#0000FF">
    <%if GroupID<>27 then%>
        Carta de Endoso No.:&nbsp;
        
        <%if Mid(BLNumber,1,3)="CMX" then %>
            <%if ExType=8 then %>
            <b><%=BL%></b>
            <%else %>
            <b><%=MBL%></b>
            <%end if %>
        <% else%>
            <%if InStr(1,ValidPatrnCIF,Countries)>0 or FRegExp(PtrnViewBLAgents, ShipperID,  "", 2) then %>
            <b><%=BL%></b>
            <%else %>
            <b><%=BLNumber%></b>
            <%end if %>
        <%end if%>
    <%end if %>
    </font><br></td>
  </tr>
</table>
<table width="641" cellpadding="2" cellspacing="0" align="center">
  <tr>
    <td class="style10" align="left" colspan="2" valign="top">Señor(es):<br><%=Estate%><br>Presente<br><br></td>
  </tr>
  <tr>
    <td class="style10" align="left" colspan="2" valign="top">Respetable(s) Señor(es):<br> Por medio del presente endoso autorizamos a ustedes para que la mercader&iacute;a
	consignada a <%=BusinessName%>, pueda ser entregada a:
	<br><br><b><%=ConsignerData%></b><br><br>
	<%if GroupID<>27 then%>
    Dicha mercader&iacute;a viene amparada por Carta de Porte 
        <%if Mid(BLNumber,1,3)="CMX" then %>
            <%if ExType=8 then %>
            Hija <b><%=BL%></b> con No de Carta Porte Master <%=MBL%>
            <%else %>
            <b><%=MBL%></b>
            <%end if %>
        <% else%>
            <%if InStr(1,ValidPatrnCIF,Countries)>0 or FRegExp(PtrnViewBLAgents, ShipperID,  "", 2) then %>
            <b><%=BL%></b>
            <%else %>
            <b><%=BLNumber%></b>
            <%end if %>
        <%end if%>
    , con el siguiente detalle:<br><br>
    <%end if %>
	<table width="585" class="styleborder" cellpadding="2" cellspacing="0" align="left">
		<tr>
			<td class="style4" align="center" valign="middle">No. Bultos</td>
			<td class="style4" align="center" valign="middle">Clase Bultos</td>
			<td class="style4" align="center" valign="middle">Peso Bruto</td>
			<td class="style4" align="center" valign="middle">Descripci&oacute;n de Carga</td>
		</tr>
		<%TotNoOfPieces = 0
		  TotWeight = 0
		for i=0 to CountTableValues2%>
		<tr>
			<td class="style4" align="right" valign="top"><span class="style10"><%=NoOfPieces(i)%></span></td>
			<td class="style4" align="left" valign="top"><span class="style10"><%=ClassNoOfPieces(i)%></span></td>
			<td class="style4" align="right" valign="top"><span class="style10"><%=Weights(i)%></span></td>
			<td class="style4" align="right" valign="top"><span class="style10"><%=DiceContener(i)%></span></td>
		</tr>
		<%TotNoOfPieces = TotNoOfPieces + NoOfPieces(i)
		  TotWeight = TotWeight + Weights(i)
		%>
		<%next%>
		<tr>
			<td class="style4" align="right" valign="top"><span class="style10"><b><%=TotNoOfPieces%></b></span></td>
			<td class="style4" align="right" valign="top"><span class="style10">&nbsp;</span></td>
			<td class="style4" align="right" valign="top"><span class="style10"><b><%=TotWeight%></b></span></td>
			<td class="style4" align="right" valign="top"><span class="style10">&nbsp;</span></td>
		</tr>
	</table>
	</td>
  </tr>
  <%if CIFBrokerIn <> "" and GroupID=27 then%>
  <tr>
    <td class="style10" align="left" valign="top" width="15%"><B>Declaraci&oacute;n&nbsp;Aduanera:</B></td>
    <td class="style10" align="left" valign="top"><%=CIFBrokerIn%></td>
  </tr>
  <%end if
    if BL <> "" and GroupID=27 then%>
  <tr>
    <td class="style10" align="left" valign="top" width="15%"><B>No. de Bill of Lading:</B></td>
    <td class="style10" align="left" valign="top"><%=BL%></td>
  </tr>
  <%end if
    if Contener <> "" and GroupID=27 then%>
  <tr>
    <td class="style10" align="left" valign="top" width="15%"><B>Contenedor:</B></td>
    <td class="style10" align="left" valign="top"><%=Contener%></td>
  </tr>
  <%end if
    if DTI <> "" and GroupID<>27 then%>
  <tr>
    <td class="style10" align="left" valign="top" width="15%"><B>Declaraci&oacute;n&nbsp;de&nbsp;Tr&aacute;nsito:</B></td>
    <td class="style10" align="left" valign="top"><%=DTI%></td>
  </tr>
  <%end if
  	if TruckNo <> "" and GroupID<>27 then%>
  <tr>
    <td class="style10" align="left" valign="top" width="15%"><B>Placa:</B></td>
    <td class="style10" align="left" valign="top"><%=TruckNo%></td>
  </tr>
  <%end if
  	if Container <> "" and GroupID<>27 then%>
  <tr>
    <td class="style10" align="left" valign="top" width="15%"><B>Furg&oacute;n TC:</B></td>
    <td class="style10" align="left" valign="top"><%=Container%></td>
  </tr>
  <%end if
  	if ContainerDep <> "" and GroupID<>27 then%>
  <tr>
    <td class="style10" align="left" valign="top" width="15%"><B>Contenedor:</B></td>
    <td class="style10" align="left" valign="top"><%=ContainerDep%></td>
  </tr>
  <%end if
  	if Chassis <> "" and GroupID<>27 then%>  
  <tr>
    <td class="style10" align="left" valign="top" width="15%"><B>Chasis:</B></td>
    <td class="style10" align="left" valign="top"><%=Chassis%></td>
  </tr>
  <%end if
  	if CIFLandFreight <> "" and GroupID=27 then%>  
  <tr>
    <td class="style10" align="left" valign="top" width="15%"><B>Valor Flete Terrestre:</B></td>
    <td class="style10" align="left" valign="top"><%=CIFLandFreight%></td>
  </tr>
  <%end if%>
  <tr>
  	<td class="style10" align="left" valign="top" colspan="2"><BR>
	Siendo <%=ConsignerData%> propietarios de dicha mercader&iacute;a y los responsables por los gastos e impuestos que ocasione esta importaci&oacute;n.<br><br>
	Sin otro particular y agradeciendo de antemano su atenci&oacute;n prestada a la presente.<br><br>
	Atentamente,

	</td>
  </tr>  
</table><br>
<table width="641" cellpadding="2" cellspacing="0" align="center">
  <tr>
	<td class="style10" align="left" width="50%"><%=Session("Sign")%></td>
  </tr>
  <tr>
	<td class="style10" align="center" width="50%"><%=Footer%></td>
  </tr>
</table>

<%Case 16%>

<table width="641" cellpadding="2" cellspacing="0" align="center">
  <tr>
	<td class="style11" align="left" width="50%">
	<%if Logo <> "" then%>
	<img src="img/<%=Logo%>" border="0">
	<%else%>
	<%=DisplayLogo(Countries, 0, 0, 0)%>
	<%end if%>
	<br><br></td>
	<td class="style3" align="right">EDICION 1 </td>
  </tr>
</table>
<table width="641" cellpadding="4" cellspacing="0" align="center">
  <tr>
	<td class="style10" align="left"><%response.write TranslateCountry(CountryDes) & " " & ConvertDate(LtAcceptDate,4)%><br></td>
  </tr>
  <tr>
	<td class="style10" align="left"><font color="#0000FF">Carta de Aceptaci&oacute;n No.:<b><%=BLNumber%></b></font><br></td>
  </tr>
</table>
<table width="641" cellpadding="2" cellspacing="0" align="center">
  <tr>
    <td class="style10" align="left" colspan="2" valign="top">Señor(es):<br><%=BrokerRecepName%><br><%if Attn <> "" then response.write Attn & "<br>"%>Presente<br><br></td>
  </tr>
  <tr>
    <td class="style10" align="left" colspan="2" valign="top">Por medio de la presente hago de su conocimiento que las mercader&iacute;as amparadas en B/L o Manifiesto
	de carga que se detalla abajo, <b><%=BusinessName%></b>, como bodega de deposito, despues de haber revisado la documentacion correspondiente estamos 
	en la disposicion de recibir el embarque descrito a continuaci&oacute;n:<br /><br>
    Las declaraciones no han sido presentadas en frontera por lo que solicitamos transito hacia ZFST.<br />
        <br><br></td>
  </tr>
  <tr>
    <td class="style10" width="15%" align="left" valign="top">Frontera&nbsp;de&nbsp;Ingreso:</td>
    <td class="style10" align="left" valign="top"><%=BrokerName%></td>
  </tr>
  <tr>
    <td class="style10" align="left" valign="top">Placa:</td>
    <td class="style10" align="left" valign="top"><%=TruckNo%></td>
  </tr>
  <tr>
    <td class="style10" align="left" valign="top">Consolidador:</td>
    <td class="style10" align="left" valign="top"><%=ShipperName%></td>
  </tr>
  <tr>
    <td class="style10" align="left" valign="top">Piloto:</td>
    <td class="style10" align="left" valign="top"><%=PilotName%></td>
  </tr>
</table><br>
<table width="641" height="250" class="styleborder" cellpadding="2" cellspacing="0" align="center">
	<tr height="8">
		<td class="style4" align="center" valign="middle">#</td>
		<td class="style4" align="center" valign="middle">B/L</td>
		<td class="style4" align="center" valign="middle">No. Bultos</td>
		<td class="style4" align="center" valign="middle">Clase Bultos</td>
		<td class="style4" align="center" valign="middle">Peso Bruto</td>
		<td class="style4" align="center" valign="middle">Consignatario</td>
	</tr>
	<%for i=0 to CountTableValues2%>
	<tr>
		<td class="style4" align="center" valign="middle"><span class="style10"><%=(i+1)%></span></td>
		<td class="style4" align="right" valign="top"><span class="style10"><%=BLs(i)%></span></td>
		<td class="style4" align="right" valign="top"><span class="style10"><%=NoOfPieces(i)%></span></td>
		<td class="style4" align="left" valign="top"><span class="style10"><%=ClassNoOfPieces(i)%></span></td>
		<td class="style4" align="right" valign="top"><span class="style10"><%=Weights(i)%></span></td>
		<td class="style4" align="right" valign="top"><span class="style10"><%=Clients(i)%></span></td>
	</tr>
	<%next%>
	<tr>
		<td class="style4" align="right" valign="top" colspan="6" height="100%">&nbsp;</td>
	</tr>
	<tr height="8">
		<td class="style4" colspan="2">&nbsp;</td>
		<td class="style4" align="right" valign="top"><b><%=TotNoOfPieces%></b></td>
		<td class="style4">&nbsp;</td>
		<td class="style4" align="right" valign="top"><b><%=TotWeight%></b></td>
		<td class="style4">&nbsp;</td>
	</tr>
</table><br>
<table width="641" cellpadding="2" cellspacing="0" align="center">
  <tr>
	<td class="style10" align="left" width="50%"><%=Session("Sign")%></td>
  </tr>
  <tr>
	<td class="style10" align="center" width="50%"><%=Footer%></td>
  </tr>
</table>

<%end select%>

</body>
</html>
<%
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>