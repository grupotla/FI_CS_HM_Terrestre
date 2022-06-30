<%
Checking "0|1|2"
Dim CreatedDate2, CreatedTime2, CreatedTime3, Flag
if CountTableValues >= 0 then
	CreatedDate = ConvertDate(aTableValues(1, 0),2)
    CreatedTime = aTableValues(2, 0)
	Expired = aTableValues(3, 0)
	BLNumber = aTableValues(4, 0)
	BLArrivalDate = aTableValues(5, 0)
	Comment = aTableValues(6, 0)
	CountryDes = aTableValues(7, 0)
	BLArrivalHour = aTableValues(8, 0)
	BLArrivalMin = aTableValues(9, 0)
	Week = aTableValues(10, 0)
	BLType = aTableValues(11, 0)
	Closed = aTableValues(12, 0)
	BLFinishHour = aTableValues(13, 0)
	BLFinishMin = aTableValues(14, 0)
    BLRealExitDate = aTableValues(15, 0)
end if
Set aTableValues = Nothing

    Comment2 = " "
    'Closed = 1
    OpenConn2 Conn
		if BLType = 2 then
			'Obteniendo el listado de Status Terrestre Local
			set rs = Conn.Execute("select id, estatus, notificar_agente, notificar_cliente, notificar_shipper, publico, estatus_es from aimartrackings where local_land=1 and activo=1 and id = 4 order by estatus")
		else
			'Obteniendo el listado de Status Terrestre Intermodal
			set rs = Conn.Execute("select id, estatus, notificar_agente, notificar_cliente, notificar_shipper, publico, estatus_es from aimartrackings where land=1 and activo=1 and id = 4 order by estatus")
		end if
		'LLenando el listado html de los Status seleccionados
		Do While not rs.EOF
            RangeID = RangeID & "RangeID[" & rs(0) & "]='" & rs(1) & "';" & vbCrLf
            RangeESID = RangeESID & "RangeESID[" & rs(0) & "]='" & rs(6) & "';" & vbCrLf
            NotifyAgentID = NotifyAgentID & "NotifyAgentID[" & rs(0) & "]='" & rs(2) & "';" & vbCrLf
            NotifyClientID = NotifyClientID & "NotifyClientID[" & rs(0) & "]='" & rs(3) & "';" & vbCrLf
            NotifyShipperID = NotifyShipperID & "NotifyShipperID[" & rs(0) & "]='" & rs(4) & "';" & vbCrLf
			rs.MoveNext
		Loop
		CloseOBJ rs

	CloseOBJ Conn
    	
   CountList1Values = -1
   OpenConn Conn

        'Plantillas de Estatus
        Set rs = Conn.Execute("select TrackingID, Template from TrackingTemplates where Expired=0 and TrackingID = 4")
		Do While not rs.EOF
			'Comment2 = Comment2 & rs(1)
            Comment2 = Comment2 & mid(rs(1),1,32)
            rs.MoveNext
		Loop
		CloseOBJ rs
        
        'Obteniendo datos para envio de Notificaciones por Mail para Agente, Cliente, Shipper cuando se ingresa o actualiza informacion
            Set rs = Conn.Execute("select BLs, HBLNumber, ClientsID, AgentsID, ShippersID, Countries, ColoadersID, ExType, Clients, Agents, CountriesFinalDes, Shippers, Coloaders, BLID, EXID, EXDBCountry, MBls, BLType from BLDetail where BLID=" & ObjectID)

            if Not rs.EOF then
                aList1Values = rs.GetRows
        		CountList1Values = rs.RecordCount - 1
            end if
            CloseOBJ rs

	CloseOBJ Conn	

    'response.write "(" & request("BLOldArrivalDate") & ")(" & request("BLArrivalDate") & ")(" & Action & ")"

'If Action=2 then 'Se utiliza el BLOLDArrivalDate, para no repetir duplicidad de registros
If request("BLOldArrivalDate") = "" and request("BLArrivalDate") <> "" and Action=2 then 'Se utiliza el BLOLDArrivalDate, para no repetir duplicidad de registros
'If Action=2 then
Flag = 1
'Transaccion
FormatTime CreatedDate2, CreatedTime2
	CreatedTime3 = CreatedTime2
	CountTableValues = -1
	OpenConn Conn
		
        'Arribo de Carga
        SQLQuery = "update BLDetail set InTransit=2 where BLID=" & ObjectID
        'response.write SQLQuery & "<br>"
		Conn.Execute(SQLQuery)

        'Arribo de Carga Desglose 2021-02-27 segun lo platicado ayer viernes con Cesar
        SQLQuery = "UPDATE BLDetail2 INNER JOIN BLDetail ON BLDetail2.BLID = BLDetail.BLDetailID AND SUBSTRING(BLDetail2.CountriesFinalDes,1,2) = '" & SetDestinationCountry(Left(CountryDes,2)) & "' SET BLDetail2.InTransit=2, BLDetail2.LtArrivalDate=NOW(), BLDetail2.LtArrivalDeliveryDocs='" & Session("OperatorID") & "' WHERE BLDetail.BLID=" & ObjectID '& " AND BLDetail.BLDetailID = 94465" 'sirvio para test con blid=-1  
        'response.write SQLQuery & "<br><br>"
		Conn.Execute(SQLQuery)
		
        'Arribo de Documentos
		SQLQuery = "update Files set InTransit=2 where BLID=" & ObjectID 
        'response.write SQLQuery & "<br>"
        Conn.Execute(SQLQuery)
		
        
        'response.write "select BLDetailID, EXID, EXType, ClientsID, AddressesID, Clients, CommoditiesID, DiceContener, Weights, Volumes, NoOfPieces, CountriesFinalDes, AgentsID, AgentsAddrID, Agents, Contact, Container, BLs, MBLs, ChargeType, Endorse, EndorseType, Declaration, DeclarationType, RequestNo, RequestType, BLsType, BillType, Bill, PackingListType, PackingList, Observations, HBLNumber, ShippersID, ShippersAddrID, Shippers, CountryOrigen, BLType, ClassNoOfPieces from BLDetail where BLID=" & ObjectID & " and (CountriesFinalDes<>'" & SetDestinationCountry(CountryDes) & "')<br>"
		'response.write "select BLDetailID, EXID, EXType, ClientsID, AddressesID, Clients, CommoditiesID, DiceContener, Weights, Volumes, NoOfPieces, CountriesFinalDes, AgentsID, AgentsAddrID, Agents, Contact, Container, BLs, MBLs, ChargeType, Endorse, EndorseType, Declaration, DeclarationType, RequestNo, RequestType, BLsType, BillType, Bill, PackingListType, PackingList, Observations, HBLNumber, ShippersID, ShippersAddrID, Shippers, CountryOrigen, BLType, ClassNoOfPieces, Notify, EXDBCountry, Seps, DTIDocType, ManifestDocType, EndorseDocType, EndorseObservations, PolicyNo, BLStatus, AsAgreed, CIFLandFreight, CIFBrokerIn, PO, ClientsTemp, AgentsTemp, DiceContenerTemp, ActionTemp, ClientColoader, ShipperColoader, AgentNeutral, ColoadersID, ColoadersAddrID, Coloaders, IncotermsID, Incoterms, ClientCollectID, ClientsCollect from BLDetail where BLID=" & ObjectID & " and (CountriesFinalDes<>'" & SetDestinationCountry(CountryDes) & "')"
        'Seleccionando Carga en Transito para cargarla automaticamente al Itinerario del Pais
        '					  0			1	  2     	3			4			5			6			7			8			9		10			11				12				13		14		  15		16		17	 18		  19		20			21			22			23				24			25			26		27		  28	  	29				30				31		  32		33				34			 35			  36      	 37			38			  39	 	40		  41        42            43               44               45              46      47          48          49             50      51       52          53              54            55            56              57              58           59              60            61          62          63            64              65             66           67                68                 69                70                      71                     72                  73                74              75		
        SQLQuery = "select BLDetailID, EXID, EXType, ClientsID, AddressesID, Clients, CommoditiesID, DiceContener, Weights, Volumes, NoOfPieces, CountriesFinalDes, AgentsID, AgentsAddrID, Agents, Contact, Container, BLs, MBLs, ChargeType, Endorse, EndorseType, Declaration, DeclarationType, RequestNo, RequestType, BLsType, BillType, Bill, PackingListType, PackingList, Observations, HBLNumber, ShippersID, ShippersAddrID, Shippers, CountryOrigen, BLType, ClassNoOfPieces, Notify, EXDBCountry, Seps, DTIDocType, ManifestDocType, EndorseDocType, EndorseObservations, PolicyNo, BLStatus, AsAgreed, CIFLandFreight, CIFBrokerIn, PO, ClientsTemp, AgentsTemp, DiceContenerTemp, ActionTemp, ClientColoader, ShipperColoader, AgentNeutral, ColoadersID, ColoadersAddrID, Coloaders, IncotermsID, Incoterms, ClientCollectID, ClientsCollect, Countries, FreightColoader, FreightColoader2, InsuranceColoader, InsuranceColoader2, AnotherChargesColoader, AnotherChargesColoader2, NotifyPartyID, NotifyPartyAddrID, NotifyParty, RefHBLNumber, CPLock, RoClientID " & _
        "from BLDetail where BLID=" & ObjectID & " and (SUBSTRING(CountriesFinalDes,1,2) <> '" & SetDestinationCountry(Left(CountryDes,2)) & "')"
        'response.write SQLQuery & "<br><br>"

        Set rs = Conn.Execute(SQLQuery)
        '"from BLDetail where BLID=" & ObjectID & " and (CountriesFinalDes<>'" & SetDestinationCountry(CountryDes) & "')") 2020-09-17
				
        if Not rs.EOF then
   			aTableValues = rs.GetRows
   			CountTableValues = rs.RecordCount-1
		end if
		CloseOBJ rs

        for i=0 to CountTableValues
			aTableValues(0,i) = CheckNum(aTableValues(0,i))
			'Almacenando las cargas en Transito
            CouDesValue = ""
            if (inStr(aTableValues(66,i),"LTF" ) and inStrRev(CountryDes,"LTF") = 0) then
                CouDesValue = Left(CountryDes,2) & "LTF"
            elseif (inStrRev(aTableValues(66,i),"LTF" ) = 0 and inStr(CountryDes,"LTF")) then
                CouDesValue = Replace(CountryDes,"LTF","")
            else
                CouDesValue = CountryDes
            end if

            if (inStr(aTableValues(66,i),"TLA" ) and inStrRev(CountryDes,"TLA") = 0) then
                CouDesValue = Left(CountryDes,2) & "TLA"
            elseif (inStrRev(aTableValues(66,i),"TLA" ) = 0 and inStr(CountryDes,"TLA")) then
                CouDesValue = Replace(CountryDes,"TLA","")
            else
                CouDesValue = CountryDes
            end if

            Conn.Execute("insert into BLDetail (CreatedDate, CreatedTime, BLIDTransit, EXID, EXType, DischargeDate, ClientsID, AddressesID, Clients, CommoditiesID, " & _
			"DiceContener, Weights, Volumes, NoOfPieces, CountriesFinalDes, AgentsID, AgentsAddrID, Agents, Contact, Container, BLs, MBLs, " & _
			"ChargeType, Endorse, EndorseType, Declaration, DeclarationType, RequestNo, RequestType, BLsType, BillType, Bill, PackingListType, " & _
			"PackingList, Observations, Countries, InTransit, HBLNumber, ShippersID, ShippersAddrID, Shippers, CountryOrigen, BLType, ClassNoOfPieces, Notify, EXDBCountry, Seps, " & _
            "DTIDocType, ManifestDocType, EndorseDocType, EndorseObservations, PolicyNo, BLStatus, AsAgreed, CIFLandFreight, CIFBrokerIn, PO, ClientsTemp, AgentsTemp, DiceContenerTemp, " & _
            "ActionTemp, ClientColoader, ShipperColoader, AgentNeutral, ColoadersID, ColoadersAddrID, Coloaders, IncotermsID, Incoterms, ClientCollectID, ClientsCollect, " & _
            "FreightColoader, FreightColoader2, InsuranceColoader, InsuranceColoader2, AnotherChargesColoader, AnotherChargesColoader2, NotifyPartyID, NotifyPartyAddrId, NotifyParty, UserID, RefHBLNumber, CPLock, RoClientID) values (" & _
				"'" & CreatedDate2 & "', " & _
				CreatedTime2 & ", " & _
				aTableValues(0,i) & ", " & _
				CheckNum(aTableValues(1,i)) & ", " & _
				CheckNum(aTableValues(2,i)) & ", " & _
				"'" & BLArrivalDate & "', " & _
				CheckNum(aTableValues(3,i)) & ", " & _
				CheckNum(aTableValues(4,i)) & ", " & _
				"'" &aTableValues(5,i) & "', " & _
				CheckNum(aTableValues(6,i)) & ", " & _
				"'" & aTableValues(7,i) & "', " & _
				CheckNum(aTableValues(8,i)) & ", " & _
				CheckNum(aTableValues(9,i)) & ", " & _
				CheckNum(aTableValues(10,i)) & ", " & _
				"'" & aTableValues(11,i) & "', " & _
				CheckNum(aTableValues(12,i)) & ", " & _
				CheckNum(aTableValues(13,i)) & ", " & _
				"'" & aTableValues(14,i) & "', " & _
				"'" & aTableValues(15,i) & "', " & _
				"'" & aTableValues(16,i) & "', " & _
				"'" & aTableValues(17,i) & "', " & _
				"'" & aTableValues(18,i) & "', " & _
				CheckNum(aTableValues(19,i)) & ", " & _
				CheckNum(aTableValues(20,i)) & ", " & _
				CheckNum(aTableValues(21,i)) & ", " & _
				CheckNum(aTableValues(22,i)) & ", " & _
				CheckNum(aTableValues(23,i)) & ", " & _
				CheckNum(aTableValues(24,i)) & ", " & _
				CheckNum(aTableValues(25,i)) & ", " & _
				CheckNum(aTableValues(26,i)) & ", " & _
				CheckNum(aTableValues(27,i)) & ", " & _
				"'" & aTableValues(28,i) & "', " & _
				CheckNum(aTableValues(29,i)) & ", " & _
				"'" & aTableValues(30,i) & "', " & _
				"'" & aTableValues(31,i) & "', " & _
				"'" & CouDesValue & "', " & _
				"0, " & _
				"'" & aTableValues(32,i) & "', " & _
				CheckNum(aTableValues(33,i)) & ", " & _
				CheckNum(aTableValues(34,i)) & ", " & _
				"'" & aTableValues(35,i) & "', " & _
				"'" & aTableValues(36,i) & "', " & _
				CheckNum(aTableValues(37,i)) & ", " & _
				"'" & aTableValues(38,i) & "', " & _
				"'" & aTableValues(39,i) & "', " & _
				"'" & aTableValues(40,i) & "', " & _
				CheckNum(aTableValues(41,i)) & ", " & _
				CheckNum(aTableValues(42,i)) & ", " & _
				CheckNum(aTableValues(43,i)) & ", " & _
				CheckNum(aTableValues(44,i)) & ", " & _
				"'" & aTableValues(45,i) & "', " & _
				"'" & aTableValues(46,i) & "', " & _
				"'" & aTableValues(47,i) & "', " & _
				CheckNum(aTableValues(48,i)) & ", " & _
				"'" & aTableValues(49,i) & "', " & _
				"'" & aTableValues(50,i) & "', " & _
				"'" & aTableValues(51,i) & "', " & _
				"'" & aTableValues(52,i) & "', " & _
				"'" & aTableValues(53,i) & "', " & _
				"'" & aTableValues(54,i) & "', " & _
				CheckNum(aTableValues(55,i)) & ", " & _
				CheckNum(aTableValues(56,i)) & ", " & _
				CheckNum(aTableValues(57,i)) & ", " & _
				CheckNum(aTableValues(58,i)) & ", " & _
                CheckNum(aTableValues(59,i)) & ", " & _
                CheckNum(aTableValues(60,i)) & ", " & _
                "'" & aTableValues(61,i) & "', " & _
				CheckNum(aTableValues(62,i)) & ", " & _
                "'" & aTableValues(63,i) & "', " & _
                CheckNum(aTableValues(64,i)) & ", " & _
                "'" & aTableValues(65,i) & "', " & _
                CheckNum(aTableValues(67,i)) & ", " & _
                CheckNum(aTableValues(68,i)) & ", " & _
                CheckNum(aTableValues(69,i)) & ", " & _
                CheckNum(aTableValues(70,i)) & ", " & _
                CheckNum(aTableValues(71,i)) & ", " & _
                CheckNum(aTableValues(72,i)) & "," & _
                CheckNum(aTableValues(73,i)) & ", " & _
                CheckNum(aTableValues(74,i)) & ", " & _
                "'" & aTableValues(75,i) & "', " & _
                Session("OperatorID") & "," & _
                "'" & aTableValues(76,i) & "'," & _
                CheckNum(aTableValues(77,i)) & ", " & _
                CheckNum(aTableValues(78,i)) & "" & _
				")")

			'Almacenando los cargos de las cargas en Transito
			Freight = 0
			Freight2 = 0
			Insurance = 0
			Insurance2 = 0
			AnotherChargesCollect = 0
			AnotherChargesPrepaid = 0
			'response.write "select BLDetailID from BLDetail where CreatedDate='" & CreatedDate2 & "' and CreatedTime=" & CreatedTime2 & "<br>"
			Set rs = Conn.Execute("select BLDetailID from BLDetail where CreatedDate='" & CreatedDate2 & "' and CreatedTime=" & CreatedTime2)	
			if Not rs.EOF then
				BL = rs(0)
				CloseOBJ rs				
				Set rs = Conn.Execute("select Currency, ItemID, Value, Local, AgentTyp, ItemName, PrepaidCollect, OverSold, ServiceID, ServiceName, AccountType, DocType, CalcInBL, InterChargeType, InterCompanyID, InterGroupID, InterProviderType, InRO from ChargeItems where Expired=0 and InterProviderType<> 5 and InterCompanyID <> 3 and SBLID=" & aTableValues(0,i))
				Do While Not rs.EOF
					'response.write("insert into ChargeItems (SBLID, CreatedDate, CreatedTime, UserID, Currency, ItemID, Value, Local, AgentTyp, ItemName, PrepaidCollect, OverSold) values (" & _
					if rs(0)="USD" then
						Conn.Execute("insert into ChargeItems (SBLID, CreatedDate, CreatedTime, UserID, Currency, ItemID, Value, Local, AgentTyp, ItemName, PrepaidCollect, OverSold, ServiceID, ServiceName, AccountType, DocType, CalcInBL, InterChargeType, InterCompanyID, InterGroupID, InterProviderType, InRO) values (" & _
						BL & ", '" & CreatedDate2 & "', " & CreatedTime3 & ", " & Session("OperatorID") & ", '" & _
						rs(0) & "', " & rs(1) & ", " & rs(2) & ", " & rs(3) & ", " & rs(4) & ", '" & rs(5) & "', " & rs(6) & ", " & rs(7) & ", " & rs(8) & ", '" & rs(9) & "', " & rs(10) & ", " & rs(11) & ", " & rs(12) & ", " & rs(13) & ", " & rs(14) & ", " & rs(15) & ", " & rs(16) & ", " & rs(17) & ")")
						CreatedTime3 = CreatedTime3+1
                        if CheckNum(rs(12))=1 then 'Se suman si deben calcularse en el BL Hijo
						    SetCharges rs(6), rs(0), rs(1), CheckNum(rs(2))+CheckNum(rs(7)) 
                        end if
					end if
					rs.MoveNext
				loop
                'Actualiza los datos calculados en SetCharges
				Conn.Execute("update BLDetail set Freight=" & Freight & ", Freight2=" & Freight2 & ", Insurance=" & Insurance & ", Insurance2=" & Insurance2 & ", AnotherChargesCollect=" & AnotherChargesCollect & ", AnotherChargesPrepaid=" & AnotherChargesPrepaid & " where BLDetailID=" & BL)
                'response.write("update BLDetail set Freight=" & Freight & ", Freight2=" & Freight2 & ", Insurance=" & Insurance & ", Insurance2=" & Insurance2 & ", AnotherChargesCollect=" & AnotherChargesCollect & ", AnotherChargesPrepaid=" & AnotherChargesPrepaid & " where BLDetailID=" & BL)
			
            
            
                'traslada los productos desglose a nuevo bldetail
                SQLQuery = "INSERT INTO BLDetail2 (BLID, NoOfPieces, ClassNoOfPieces, CommoditiesID, DiceContener, Volumes, Weights, Countries, CountryOrigen, CountriesFinalDes, InTransit) " & _
               "SELECT " & BL & ", NoOfPieces, ClassNoOfPieces, CommoditiesID, DiceContener, Volumes, Weights, Countries, CountryOrigen, CountriesFinalDes, 1 FROM BLDetail2 WHERE BLID = " & aTableValues(0,i)
               'response.write SQLQuery & "<br><br>"
               Conn.Execute(SQLQuery)
 
                         
            end if
			CloseOBJ rs
			CreatedTime2 = CreatedTime2+1
		next

		'Carga directa del pais
		'Conn.Execute("update BLDetail set InTransit=3 where BLID=" & ObjectID & " and CountriesFinalDes='" & CountryDes & "'")
		'Documentos directos del pais
		'Conn.Execute("update Files set InTransit=3 where BLID=" & ObjectID & " and CountriesFinalDes='" & CountryDes & "'")
		'Carga en transito
		'Conn.Execute("update BLDetail set InTransit=2 where BLID=" & ObjectID & " and CountriesFinalDes<>'" & CountryDes & "'")
		'Documentos en transito
		'Conn.Execute("update Files set InTransit=2 where BLID=" & ObjectID & " and CountriesFinalDes<>'" & CountryDes & "'")

		'Rastreo / Tracking
		'set rs = Conn.Execute("select BLID from Tracking where BLID=" & ObjectID & " and ClientID=0 and Comment='La Carga Arribo en " & CountryDes & "'")
		'if rs.EOF then
		'	FormatTime CreatedDate2, CreatedTime2	
		'	Conn.Execute("insert into Tracking (BLID, ClientID, Comment, OperatorID, CreatedDate, CreatedTime) values (" & ObjectID & ", 0, 'La Carga Arribo en " & CountryDes & "' , " & Session("OperatorID") & ", '" & CreatedDate2 & "', " & CreatedTime2 & ")")
		'end if	
	CloseOBJ Conn
	
        'Notificaciones
        AgentID = CheckNum(Request.Form("NAgentID"))
        ConsignerID = CheckNum(Request.Form("NClientID"))
        ShipperID = CheckNum(Request.Form("NShipperID"))


            Header = "<html xmlns='http://www.w3.org/1999/xhtml'>" & _
		    "<head><meta http-equiv='Content-Type' Content='text/html; charset=iso-8859-1' />" & _
		    "<title>Agencia Internacional Maritima S.A. a Logistic Company Representing</title>" & _
		    "<style type='text/css'>" & _
		    "<!--" & _
		    "body {" & _
		    "	margin-left: 0px;" & _
		    "	margin-top: 0px;" & _
		    "	margin-right: 0px;" & _
		    "	margin-bottom: 0px;" & _
		    "}" & _
		    ".contenido1 {" & _
		    "   font-family: Verdana, Arial, Helvetica, sans-serif;" & _
		    "   font-size: 10px;" & _
		    "   font-weight: normal;" & _
		    "   color: #666666;" & _
		    "   text-decoration: none;" & _
		    "   padding: 3px;" & _
		    "}" & _
		    "-->" & _
		    "</style></head><body>" & _
		    "<table align=left cellpadding=0 cellspacing=2>"

             for i=0 to CountList1Values

               ' Si country destino de master es igual a country destino hija manda notificacion
               if (CountryDes = aList1Values(10,i)) then
                
                  'Si no es para salvador o costa rica envia la noticacion
                  if (aList1Values(10,i) = "SV") then
                    Flag = 0
                  end if

                  if (aList1Values(10,i) = "CR") then
                    Flag = 0
                  end if

                  if (Flag = 1) then

                    DBVentas = aList1Values(15,i)

                    if (DBVentas <> "") then

                        OpenConnOcean Conn, "ventas_" & SetDBCountry(DBVentas)
        
                            EXID = aList1Values(14,i)
                            EXType = aList1Values(7,i)
                            Select Case EXType
                            Case 0 'FCL
                                    set rs = Conn.Execute("select bl_id, en_transito, id_pais_final from bl_completo where bl_id=" & EXID)
			                        If not rs.EOF then
				                        CorreoTranshipper = rs(2)
			                        end If
                            Case 1,2 'LCL
                                    set rs = Conn.Execute("select bl_id, en_intermodal, id_pais_final2 from bill_of_lading where bl_id=" & EXID)
			                        If not rs.EOF then
                                        CorreoTranshipper = rs(2)
			                        end If
                            Case 11 'FCL
                                    set rs = Conn.Execute("select bl_id, en_transito, id_pais_final from bl_completo where bl_id=" & EXID)
			                        If not rs.EOF then
                                        CorreoTranshipper = rs(2)
			                        end If
                            Case 12,13 'LCL
                                    set rs = Conn.Execute("select bl_id, en_intermodal, id_pais_final2 from bill_of_lading where bl_id=" & EXID)
			                        If not rs.EOF then
                                        CorreoTranshipper = rs(2)
			                        end If
                            Case else
                                    CorreoTranshipper = ""
                            End Select

                        CloseOBJ Conn
                    end if

                        OpenConn Conn
		                    'Arribo de Carga		    
                                Conn.Execute("insert into Tracking (BLID, ClientID, CreatedDate, CreatedTime, Comment, OperatorID, BLStatus, BLStatusName) values ('" & aList1Values(13,i) & "','" & aList1Values(2,i)  & "','" & CreatedDate2 & "','" & CreatedTime2 & "','" & Request.Form("BLStatusName") & " " & " " & aList1Values(10,i) & "','" & Session("OperatorID") & "',4,'" & Request.Form("BLStatusName") & "')")
                        CloseOBJ Conn

                        'Ticket#2019082131000175 — RV: ESTATUS: CGT-2019-25-0031-83326/COTI: 15951/LL-NI-I-24-19-0870/S: REPARACIONES INTERMODALES, S. A./C: CLEAN MASTER DE NICARAGUA S.A. 
                        '2019-08-23 version 1.9.2
                        'Si no hay Coloader, se notifica al Cliente, Shipper, Agente segun configuracion en BBDD
                        if CheckNum(aList1Values(6,i))=0 then
                            'Notificacion al Cliente 
                            if ConsignerID = 1 then
                                SendNotification Header, "Cliente", aList1Values(0,i), aList1Values(1,i), Request.Form("BLStatusName") & " (<i>" & Request.Form("BLStatusNameES") & "</i>)", Comment2 & " (<i>" & Request.Form("BLArrivalDate") & "</i>)", "select a.email as email, b.id_grupo as id_grupo from contactos a inner join clientes b on a.id_cliente = b.id_cliente where a.id_cliente=" & CheckNum(aList1Values(2,i)) & " and a.activo=true and character_length(a.email)>5", aList1Values(5,i), aList1Values(8,i), aList1Values(9,i), aList1Values(10,i), aList1Values(7,i), CheckNum(aList1Values(6,i)), aList1Values(12,i), 1, CorreoTranshipper, aList1Values(16,i), aList1Values(17,i), aList1Values(15,i)
                                'SendNotification Header, "Cliente", aList1Values(0,i), aList1Values(1,i), Request.Form("BLStatusName") & " (<i>" & Request.Form("BLStatusNameES") & "</i>)", Comment2 & " (<i>" & Request.Form("BLArrivalDate") & "</i>)", "select a.email as email, b.id_grupo as id_grupo from contactos a inner join clientes b on a.id_cliente = b.id_cliente where a.id_cliente=76153  and a.activo=true and character_length(a.email)>5 UNION select email as email, id_grupo as id_grupo from clientes where character_length(email)>5 and id_cliente=76153", aList1Values(5,i), aList1Values(8,i), aList1Values(9,i), aList1Values(10,i), aList1Values(7,i), CheckNum(aList1Values(6,i)), aList1Values(12,i), 1, CorreoTranshipper, aList1Values(16,i)
                            end if
                            'Notificacion al Shipper
                            if ShipperID = 1 then                    
                                Select Case CheckNum(aList1Values(7,i))
                                Case 4,5,6,7 'Solo cuando es RO se envia al Shipper: 4=RO-Consolidado,5=RO-Express,6=RO-Recoleccion,7=RO-Entrega
                                    SendNotification Header, "Shipper", aList1Values(0,i), aList1Values(1,i), Request.Form("BLStatusName") & " (<i>" & Request.Form("BLStatusNameES") & "</i>)", Comment2 & " (<i>" & Request.Form("BLArrivalDate") & "</i>)", "select a.email as email, b.id_grupo as id_grupo from contactos a inner join clientes b on a.id_cliente = b.id_cliente where a.id_cliente=" & CheckNum(aList1Values(3,i)) & " and a.activo=true and character_length(a.email)>5", aList1Values(5,i), aList1Values(8,i), aList1Values(9,i), aList1Values(10,i), aList1Values(7,i), CheckNum(aList1Values(6,i)), aList1Values(12,i), 2, CorreoTranshipper, aList1Values(16,i), aList1Values(17,i), aList1Values(15,i)
                                End Select
                            end if
                            'Notificacion al Agente, si no tiene RO
                            '4=RO-Consolidado,5=RO-Express,6=RO-Recoleccion,7=RO-Entrega
                            if AgentID = 1 and (aList1Values(7,i)<4 or aList1Values(7,i)>7) then
                               SendNotification Header, "Agente", aList1Values(0,i), aList1Values(1,i), Request.Form("BLStatusName") & " (<i>" & Request.Form("BLStatusNameES") & "</i>)", Comment2 & " (<i>" & Request.Form("BLArrivalDate") & "</i>)", "select correo from agentes where agente_id=" & CheckNum(aList1Values(4,i)) & " and character_length(correo)>5 UNION ALL select email from agentes_contactos where agente_id =" & CheckNum(aList1Values(4,i)) & " and character_length(email)>5", aList1Values(5,i), aList1Values(8,i), aList1Values(9,i), aList1Values(10,i), aList1Values(7,i), CheckNum(aList1Values(4,i)), aList1Values(11,i), 0, ""
                            end if
                        else
                            if ConsignerID = 1 or ShipperID = 1 then
                                SendNotification Header, "Cliente", aList1Values(0,i), aList1Values(1,i), Request.Form("BLStatusName") & " (<i>" & Request.Form("BLStatusNameES") & "</i>)", Comment2 & " (<i>" & Request.Form("BLArrivalDate") & "</i>)", "select a.email as email, b.id_grupo as id_grupo from contactos a inner join clientes b on a.id_cliente = b.id_cliente where a.id_cliente=" & CheckNum(aList1Values(6,i)) & " and a.activo=true and character_length(a.email)>5", aList1Values(5,i), aList1Values(8,i), aList1Values(9,i), aList1Values(10,i), aList1Values(7,i), CheckNum(aList1Values(6,i)), aList1Values(12,i), 3, CorreoTranshipper, aList1Values(16,i), aList1Values(17,i), aList1Values(15,i)         
                            end if
                        end if
                  end if
               end if
            next
            Set aList1Values = Nothing
          
end if
%>
<HTML>
<HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<style>
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
</style>
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
var RangeID = new Array();
var RangeESID = new Array();
var NotifyAgentID = new Array();
var NotifyClientID = new Array();
var NotifyShipperID = new Array();
<%=RangeID%>
<%=RangeESID%>
<%=NotifyAgentID%>
<%=NotifyClientID%>
<%=NotifyShipperID%>


	function validar(Action) {
		if ((document.forma.BLArrivalDate.value=="") || (document.forma.BLRealExitDate.value=="")){
            alert("Debe Ingresar Fecha de Llegada o Salida");
            return (false);
        };
        document.forma.BLStatusName.value = RangeID[4];
        document.forma.BLStatusNameES.value = RangeESID[4];
        document.forma.NAgentID.value = NotifyAgentID[4];
        document.forma.NClientID.value = NotifyClientID[4];
        document.forma.NShipperID.value = NotifyShipperID[4];

        move();
	    document.forma.Action.value = Action;
		document.forma.submit();			 
	 }	

    function move() {
        document.forma.style.display = "none";
        document.getElementById('myBar').style.display = "block";
        var elem = document.getElementById("myBar");
        var width = 10;
        var id = setInterval(frame, 65);
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
	 
	function SetLabelID(Label) {
		var LabelID = "";
		if (Label == "BLArrivalDate") {
			LabelID = "Fecha de Llegada, Entrega o Recolección";
		} 
        if (Label == "BLRealExitDate") {
			LabelID = "Fecha Real de Salida";
		}
        
		return LabelID;
	} 
	 
	function abrir(Label){
		var Closed = <%=Closed%>;
		var DateSend, Subject;
		if (Closed==1) {

           console.log(navigator.appVersion); 	

           try {                   

			    if (parseInt(navigator.appVersion) < 5) {
				    DateSend = document.forma(Label).value;
			    } else {
				    var LabelID = SetLabelID(Label);
				    DateSend = document.getElementById(LabelID).value;
			    }

            }
            catch(err) {
                console.log(err); 								
            }

			Subject = '';	
			window.open('Agenda.asp?Action=1&Label=' + Label + '&DateSend=' + DateSend + '&Subj=' + Subject,'Seleccionar','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,width=170,height=160,top=250,left=250');
		} else {
			alert('Para indicar una fecha, la Carta Porte generada en origen debe estar "Cerrada"');
		}	
	}
	
_editor_url = "Javascripts/";                     // URL to htmlarea files
var win_ie_ver = parseFloat(navigator.appVersion.split("MSIE")[1]);
if (navigator.userAgent.indexOf('Mac')        >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Windows CE') >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Opera')      >= 0) { win_ie_ver = 0; }
if (win_ie_ver >= 5.5) {
  document.write('<scr' + 'ipt src="' +_editor_url + 'editor.js"');
  document.write(' language="Javascript1.2"></scr' + 'ipt>');  
} else { document.write('<scr'+'ipt>function editor_generate() { return false; }</scr'+'ipt>'); }
</script>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
	<div id="myProgress">
      <div id="myBar">10%</div>
    </div>
	<TABLE cellspacing=0 cellpadding=2 width=400 align=center>
	<FORM name="forma" action="InsertData.asp" method="post">
	<INPUT name="Action" type=hidden value=0>
	<INPUT name="Expired" type=hidden value="on">
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
	<INPUT name="OID" type=hidden value="<%=ObjectID%>">
	<INPUT name="CD" type=hidden value="<%=CreatedDate%>">
	<INPUT name="CT" type=hidden value="<%=CreatedTime%>">
    <input name="BLStatusName" type=hidden value="">
    <input name="BLStatusNameES" type=hidden value="">
    <input name="NAgentID" type=hidden value="0">
    <input name="NClientID" type=hidden value="0">
    <input name="NShipperID" type=hidden value="0">
	<INPUT name="BLOldArrivalDate" type=hidden value="<%=BLArrivalDate%>">
		<TR><TD class=label align=right><b>No. Carta Porte:</b></TD><TD class=label align=left><%=BLNumber%></TD></TR> 
		<TR><TD class=label align=right><b>Fecha Creaci&oacute;n:</b></TD><TD class=label align=left><%=CreatedDate%></TD></TR> 
		<TR><TD class=label align=right><b>C&oacute;digo:</b></TD><TD class=label align=left><%if ObjectID <> 0 then response.write ObjectID End if%></TD></TR> 
		<TR><TD class=label align=right><b>
		<%select case BLType
		case 0, 1%>
		Fecha&nbsp;Real&nbsp;de&nbsp;Salida:
        <%case 2, 3%>
		Fecha&nbsp;Real&nbsp;de&nbsp;Recolecci&oacute;n:
		<%end select%>
		</b></TD><TD class=label align=left>
		<INPUT readonly="readonly" name="BLRealExitDate" id="Fecha Real de Salida" type=text value="<%=BLRealExitDate%>" size=23 maxLength=19 class=label>&nbsp;
		<INPUT type=button value="Seleccionar" onClick="JavaScript:abrir('BLRealExitDate');" class=label></TD></TR>
		<TR><TD class=label align=right><b>
        <%select case BLType
		case 0, 1%>
        Fecha&nbsp;Real&nbsp;de&nbsp;Llegada:
		<%case 2, 3%>
		Fecha&nbsp;Real&nbsp;de&nbsp;Entrega:
		<%end select%>
		</b></TD><TD class=label align=left>
		<INPUT readonly="readonly" name="BLArrivalDate" id="Fecha Real de Llegada, Entrega o Recolección" type=text value="<%=BLArrivalDate%>" size=23 maxLength=19 class=label>&nbsp;
		<INPUT type=button value="Seleccionar" onClick="JavaScript:abrir('BLArrivalDate');" class=label></TD></TR>
		<TR><TD class=label align=right><b>
		<%select case BLType
		case 0, 1%>
		Tiempo de Llegada:
		<%case 2, 3%>
		Tiempo&nbsp;de&nbsp;Entrega&nbsp;o&nbsp;Recolecci&oacute;n:
		<%end select%>
		</b></TD><TD class=label align=left>
		<select name="BLArrivalHour" id="Hora de Llegada, Entrega o Recolección" class="label">
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
		<select name="BLArrivalMin" id="Minutos de Llegada, Entrega o Recolección" class="label">
			<option value="-1">Minuto</option>
			<option value="00">00</option>
			<option value="05">05</option>
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
			</TD></TR>
		<%if BLType=2 or BLType=3 then 'Transportel Local%>
		<TR><TD class=label align=right><b>Tiempo de Finalizaci&oacute;n:</b></TD><TD class=label align=left>
		<select name="BLFinishHour" id="Hora de Finalizacion de Servicio" class="label">
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
		<select name="BLFinishMin" id="Minutos de Finalizacion de Servicio" class="label">
			<option value="-1">Minuto</option>
			<option value="00">00</option>
			<option value="05">05</option>
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
			</TD></TR>
	
		<%end if%>
		<TR><TD class=label align=right><b>Comentario:</b></TD><TD class=label align=left><Textarea name="Comment" id="Comentario" cols="40" rows="10" class="style10"><%=Comment%></Textarea></TD></TR> 
		<TD colspan="2" class=label align=center>
			<TABLE cellspacing=0 cellpadding=2 width=200>
		    <TR>
				 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="Javascript:window.open('BLPrint.asp?BLID=<%=ObjectID%>&BTP=<%=BLType%>','BLPrint','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750,height=600');return false;" value="Ver&nbsp;Carta&nbsp;Porte" class=label></TD>
				 <!--<TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="Javascript:window.open('ConsolReports.asp?GID=<%=GroupID%>&W=<%=Week%>&BTP=<%=BLType%>&YR=<%=Year(CreatedDate)%>','ConsolidItinerary','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750,height=600');return false;" value="Ver&nbsp;Itinerario&nbsp;Consolidado&nbsp;por&nbsp;semana" class=label></TD>
				 -->
				 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(2)" value="&nbsp;&nbsp;Actualizar&nbsp;&nbsp;" class=label></TD>
			</TR>
			</TABLE>
		<TD>
		</TR>
	</FORM>
	</TABLE>
</BODY>
<script language="javascript1.2">
selecciona('forma.BLArrivalHour','<%=BLArrivalHour%>');
selecciona('forma.BLArrivalMin','<%=BLArrivalMin%>');
<%select case BLType
case 2,3%>
selecciona('forma.BLFinishHour','<%=BLFinishHour%>');
selecciona('forma.BLFinishMin','<%=BLFinishMin%>');
<%end select%>
editor_generate('Comment');
</script>

</HTML>