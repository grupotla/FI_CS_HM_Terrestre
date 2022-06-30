<%
Dim BAWResult, EXID_Ant
Dim file, TypeReference, CodeReference, TripType, ClientWarehouse, TripPriority, TripRequestDate, OtherDocs, CodeReferenceValid : CodeReferenceValid = CheckNum(Request("CodeReferenceValid"))

Checking "0|1|2"
EXID_Ant = 0
EXID = CheckNum(Request("EID"))
ExType = CheckNum(Request("ET"))
Countries = Request("CTR")
CountryDes = Request("CTR2")
CountList1Values = -1
CountTableValues2 = -1
CountList3Values = -1
CommodityCode = 0
Pos = -1
ClientID=0
ShipperID=0
AgentID=0
Freight = 0
Freight2 = 0
Insurance = 0
Insurance2 = 0
AnotherChargesCollect = 0
AnotherChargesPrepaid = 0
BLType = -3
BLID = -1
ReservationDate = ""
NotifyPartyID = 0
NotifyPartyAddrID = 0
NotifyParty = null
RefHBLNumber = ""
RefBLID = 0
RosClientID = 0

RefBLID = 0

TypeReference = Request.Form("TypeReference")
CodeReference = Request.Form("CodeReference")
TripType = Request.Form("TripType")
ClientWarehouse = Request.Form("ClientWarehouse")
TripPriority = Request.Form("TripPriority")
TripRequestDate = Request.Form("TripRequestDate")
OtherDocs = Request.Form("OtherDocs")


WareHouseDischargeDate = Request.Form("WareHouseDischargeDate")
ClientID = Request.Form("ClientsID")
AddressID = Request.Form("AddressesID")
Name = Request.Form("Clients")
ClientColoader = Request.Form("ClientColoader")
file = Request.Form("file")

'response.write "(" & CountTableValues & ")<br>"


    Dim QryTest 
    QryTest = QuerySelect & TableName & " where EXID=" & EXID & " and ExType=" & ExType & " and Countries='" & Countries & "' and EXDBCountry='" & CountryDes & "' and Expired = 0"
    'response.write CountTableValues & ") entro " & CountTableValues & "<br>" & QryTest & "<br>"

    'response.write QryTest & "<br><br>"

'Al no encontrarse los datos ya grabados en la base terrestre por medio del ObjectID, se busca por el ID externo EXID, media vez no sea ExType=8 (CIF)
if CountTableValues < 0 and ExType<>8 and ExType<>15 and ExType<>99 then

'response.write "ENTRO 1<br>"


	Set aTableValues = nothing
	OpenConn Conn
    Set rs = Conn.Execute(QryTest)
    if Not rs.EOF then
		aTableValues = rs.GetRows
		CountTableValues = rs.RecordCount-1

'response.write "ENTRO 2<br>"

	else
		CountTableValues = -1
	end if
	CloseOBJs rs, Conn
end if

'response.write "(" & EXID & ")(" & ExType & ")(" & EXID_Ant & ")(" & CheckNum(Request("GID")) & ")<br>"

'response.write QuerySelect & "<br>"
'response.write(CountTableValues)
'Si se encuentran datos por medio del ObjectID o EXID se despliegan sino se traen de la base de ventas externa correspondiente por medio del EXID
if CountTableValues >= 0 then

	    ObjectID = aTableValues(0, 0)
	    CreatedDate = ConvertDate(aTableValues(1, 0),2)
        CreatedTime = aTableValues(2, 0)

        if CheckNum(Request("GID")) = 37 and ExType = 99 and EXID > 0 then	    
            EXID = 0        
        else
            EXID = CDbl(aTableValues(3, 0))
        end if

        EXID_Ant = CDbl(aTableValues(3, 0)) '2020-05-18

	    ExType = CDbl(aTableValues(4, 0))
	    WareHouseDischargeDate = aTableValues(5, 0)

        if CheckNum(Request("cambio")) = 0 then
	        ClientID = aTableValues(6, 0)     
	        AddressID = aTableValues(7, 0)
            ClientColoader = aTableValues(54, 0)
	        Name = aTableValues(8, 0)
	        NameES = aTableValues(8, 0)
        end if

	    CommodityCode = aTableValues(9, 0)
	    Commodity = aTableValues(10, 0)
	    Weight = aTableValues(11, 0)
	    Volume = aTableValues(12, 0)
	    TotNoOfPieces = aTableValues(13, 0)
	    FinalDes = aTableValues(14, 0)
	    ShipperID = aTableValues(15, 0)
	    ShipperAddrID = aTableValues(16, 0)
	    ShipperData = aTableValues(17, 0)
	    ContactSignature = aTableValues(18, 0)
	    Container = aTableValues(19, 0)
	    BL = aTableValues(20, 0)
	    MBL = aTableValues(21, 0)
	    ChargeType = aTableValues(22, 0)
	    Endorse = aTableValues(23, 0)
	    EndorseType = aTableValues(24, 0)
	    Declaration = aTableValues(25, 0)
	    DeclarationType = aTableValues(26, 0)
	    RequestNo = aTableValues(27, 0)
	    RequestType = aTableValues(28, 0)
	    BLsType = aTableValues(29, 0)
	    BillType = aTableValues(30, 0)
	    Bill = aTableValues(31, 0)
	    PackingListType = aTableValues(32, 0)
	    PackingList = aTableValues(33, 0)
	    Observations = aTableValues(34, 0)
	    Countries = aTableValues(35, 0)
        AgentID = aTableValues(36, 0)
	    AgentAddrID = aTableValues(37, 0)
	    AgentData = aTableValues(38, 0)
        Contener = aTableValues(39, 0)
        if aTableValues(40, 0) <> "" then
	        CountryDes = aTableValues(40, 0)
        end if 
	    BLType = aTableValues(41, 0)
	    Notify = aTableValues(42, 0)
	    CountryOrigen = aTableValues(43, 0)
	    DeliveryDate = aTableValues(44, 0)
        Sep = aTableValues(45, 0)
        CIFLandFreight = aTableValues(46, 0)
        CIFBrokerIn = aTableValues(47, 0)
        PO = aTableValues(48, 0)
        ClientsTemp = aTableValues(49, 0)
        AgentsTemp = aTableValues(50, 0)
        DiceContenerTemp = aTableValues(51, 0)
        BLID = aTableValues(52, 0)
        BLNumber = replace(Trim(aTableValues(53, 0)),"--","",1,-1)
    
        ShipperColoader = aTableValues(55, 0)
        AgentNeutral = aTableValues(56, 0)
        ColoaderID = aTableValues(57, 0)
        ColoaderAddrID = aTableValues(58, 0)
        ColoaderData = aTableValues(59, 0)
        IncotermsID = aTableValues(60, 0)
        Incoterms = aTableValues(61, 0)
        SenderData = aTableValues(62, 0)
        ConsignerData = aTableValues(63, 0)
        AgentSignature = aTableValues(64, 0)
        Phone1 = aTableValues(65, 0)
        BLArrivalDate = aTableValues(66, 0)
        ClientCollectID = aTableValues(67, 0)
        ClientsCollect = aTableValues(68, 0)
        Week = aTableValues(69, 0)
        file = aTableValues(80, 0)

        if Action <> 9 then 'valida codigo referencia

            if TypeReference = "" and  aTableValues(70, 0) <> "" then
                TypeReference = aTableValues(70, 0)
            end if  

            if CodeReference = "" and  aTableValues(71, 0) <> "" then
                CodeReference = aTableValues(71, 0)
            end if  
        

            if TripType = "" and  aTableValues(72, 0) <> "" then
                TripType = aTableValues(72, 0)
            end if

            if ClientWarehouse = "" and  aTableValues(73, 0) <> "" then
                ClientWarehouse = aTableValues(73, 0)
            end if

            if TripPriority = "" and  aTableValues(74, 0) <> "" then
                TripPriority = aTableValues(74, 0)
            end if

            'response.write "(" & TripRequestDate &   ")(" & aTableValues(75, 0) & ")"

            'if TripRequestDate = "" and  aTableValues(75, 0) <> "" then
                TripRequestDate = aTableValues(75, 0)
            'end if

            if OtherDocs = "" and  aTableValues(76, 0) <> "" then
                OtherDocs = aTableValues(76, 0)
            end if
   
        end if

end if

'response.write "(" & CreatedDate & ")<br>"
'response.write "(" & CreatedTime & ")<br>"


'Dim isLocal     
'isLocal = IIf(Request.ServerVariables("remote_addr") = "::1" or Request.ServerVariables("remote_addr") = "127.0.0.1", true, false)
'if isLocal = true then    
'    TripRequestDate1 = TwoDigits(Day(TripRequestDate)) & "/" & TwoDigits(Month(TripRequestDate)) & "/" & Year(TripRequestDate)
'else
'    TripRequestDate1 = TwoDigits(Month(TripRequestDate)) & "/" & TwoDigits(Day(TripRequestDate)) & "/" & Year(TripRequestDate)
'end if

TripRequestDate = ConvertDate(TripRequestDate,2)

'response.write "(" & TripRequestDate1 & ")<br>"





'cuando es intercom debe asignar fecha 2021-02-23
if ExType = 8 then

    'response.write "(" & CountTableValues & ")<br>"

    'if CheckNum(Trim(WareHouseDischargeDate)) = 0 then
    if Trim(WareHouseDischargeDate) = "" then
        'WareHouseDischargeDate = ConvertDate(Now,1)

        WareHouseDischargeDate = TwoDigits(Day(Now)) & "/" & TwoDigits(Month(Now)) & "/" & Year(Now)		 		

    end if



    if CountTableValues >= 0 then


            'BUSCA COINCIDENCIAS PARA DATOS QUE VIENEN DEL RO
            
            OpenConn2 Conn

            'ClientsTemp = aTableValues(49, 0)
            'AgentsTemp = aTableValues(50, 0)
            'DiceContenerTemp = aTableValues(51, 0)

            if ClientsTemp <> "" then               
                'QryTest = "SELECT id_cliente, nombre_cliente FROM clientes WHERE nombre_cliente ILIKE '%" & ClientsTemp & "%' AND UPPER(SUBSTRING(id_pais,1,2)) = UPPER(SUBSTRING('" & FinalDes & "',1,2))"
                
                
QryTest = "SELECT id_cliente, nombre_cliente, id_pais, " & _
"TRIM(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(UPPER(nombre_cliente),'S',''),'A',''),'.',''),',',''),' DE ',' ')),  " & _
"TRIM(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(UPPER('" & ClientsTemp & "'),'S',''),'A',''),'.',''),',',''),' DE ',' '))  " & _
"FROM clientes " & _
"WHERE id_estatus = 1   AND  " & _
"TRIM(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(UPPER(nombre_cliente),'S',''),'A',''),'.',''),',',''),' DE ',' ')) ILIKE '%' || " & _
"TRIM(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(UPPER('" & ClientsTemp & "'),'S',''),'A',''),'.',''),',',''),' DE ',' '))  || '%' " & _
" AND UPPER(SUBSTRING(id_pais,1,2)) = UPPER(SUBSTRING('" & FinalDes & "',1,2)) ORDER BY id_cliente DESC" & _
" -- LIMIT 10   "                 
                'response.write "" & QryTest & "<br>"
                Set rs = Conn.Execute(QryTest)
                if Not rs.EOF then
                    if rs.RecordCount = 1 then
                        ClientID = CheckNum(rs(0))
                        Name = rs(1)
                        'response.write "" & ClientID & " " & Name & "<br>"
                    end if 
                end if
            end if

             if AgentsTemp <> "" then               
                'QryTest = "SELECT agente_id, agente FROM agentes WHERE agente ILIKE '%" & AgentsTemp & "%' AND UPPER(SUBSTRING(countries,1,2)) = UPPER(SUBSTRING('" & FinalDes & "',1,2))"
                'QryTest = "SELECT id_cliente, nombre_cliente FROM clientes WHERE nombre_cliente ILIKE '%" & AgentsTemp & "%' AND UPPER(SUBSTRING(id_pais,1,2)) = UPPER(SUBSTRING('" & FinalDes & "',1,2))"
                               
QryTest = "SELECT clientes.id_cliente, nombre_cliente, id_direccion, id_pais, " & _
"TRIM(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(UPPER(nombre_cliente),'S',''),'A',''),'.',''),',',''),' DE ',' ')),  " & _
"TRIM(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(UPPER('" & AgentsTemp & "'),'S',''),'A',''),'.',''),',',''),' DE ',' '))  " & _
"FROM clientes, direcciones " & _
"WHERE clientes.id_cliente = direcciones.id_cliente AND id_estatus = 1 AND  " & _
"TRIM(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(UPPER(nombre_cliente),'S',''),'A',''),'.',''),',',''),' DE ',' ')) ILIKE '%' || " & _
"TRIM(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(UPPER('" & AgentsTemp & "'),'S',''),'A',''),'.',''),',',''),' DE ',' '))  || '%' " & _
" -- AND UPPER(SUBSTRING(id_pais,1,2)) = UPPER(SUBSTRING('" & FinalDes & "',1,2)) " & _ 
" ORDER BY clientes.id_cliente DESC" & _
" -- LIMIT 10   "        
                'response.write "" & QryTest & "<br>"
                Set rs = Conn.Execute(QryTest)
                if Not rs.EOF then
                    if rs.RecordCount = 1 then
                        ShipperID = CheckNum(rs(0))
                        ShipperData = rs(1)
                        'response.write "" & ShipperID & " " & Shippers & "<br>"

                        ShipperAddrID = CheckNum(rs(2))

                    end if 
                end if
            end if



             if DiceContenerTemp <> "" then               
           
QryTest = "SELECT commodityid, namees, " & _ 
"TRIM(REPLACE(REPLACE(REPLACE(REPLACE(UPPER(namees),               ' PARA ',' '),'.',''),',',''),' DE ',' ')),  " & _
"TRIM(REPLACE(REPLACE(REPLACE(REPLACE(UPPER('" & DiceContenerTemp & "'),' PARA ',' '),'.',''),',',''),' DE ',' '))  " & _
"FROM commodities " & _ 
"WHERE expired = 0 AND  " & _
"TRIM(REPLACE(REPLACE(REPLACE(REPLACE(UPPER(namees),               ' PARA ',' '),'.',''),',',''),' DE ',' ')) ILIKE '%' || " & _
"TRIM(REPLACE(REPLACE(REPLACE(REPLACE(UPPER('" & DiceContenerTemp & "'),' PARA ',' '),'.',''),',',''),' DE ',' '))  || '%' ORDER BY commodityid DESC"
       
                'response.write "" & QryTest & "<br>"
                Set rs = Conn.Execute(QryTest)
                if Not rs.EOF then
                    if rs.RecordCount = 1 then
                        CommodityCode = CheckNum(rs(0))
                        Commodity = rs(1)
                        'response.write "" & ShipperID & " " & DiceContener & "<br>"
                    end if 
                end if
            end if






            CloseOBJs rs, Conn

    end if

end if



if Action = 8 then 'valida codigo referencia
    CodeReference = ""
    CodeReferenceValid = 0
end if

        if CheckNum(CodeReference) > 0 then
            CodeReferenceValid = 1
        end if      
        
        
        'response.write "(" & CheckNum(CodeReferenceValid) & ")<br>"
              

        'response.write "(" & Action & ")<br>"
      
if Action = 9 then 'valida codigo referencia
        
        if CheckNum(CodeReference) > 0 then

            CodeReference = ValidCodeReference(CodeReference)

            if CodeReference = "" then
                response.write "<script>alert('Codigo Referencia ya fue utilizado anteriormente');document.getElementById('Codigo Referencia').focus();</script>"
                CodeReferenceValid = 0
            else
                CodeReferenceValid = 1
            end if
        end if

        'OpenConn Conn        
		'Set rs = Conn.Execute("select BLDetailID, CodeReference, HBLNumber FROM BLDetail WHERE CodeReference='" & CodeReference & "'")
		'if Not rs.EOF then

        '    if CodeReference = rs(1) or CheckNum(rs(1)) > 0 then
        '        CodeReference = ""
        '        CodeReferenceValid = 0
        '        Action = 0
        '        response.write "<script>alert('Codigo Referencia ya fue utilizado anteriormente');document.getElementById('Codigo Referencia').focus();</script>"
        '    else
        '        CodeReferenceValid = 1
        '    end if
                      	
		'end if
    	'CloseOBJs rs, Conn

end if

    'response.write "(" & ObjectID & ")<br>"


'Guardando los rubros a cobrar del Routing o BL
if Action=1 and ObjectID<>0 then

   On Error Resume Next 

	OpenConn Conn
        'Almacenando Rubros de cobro directo
		Select Case Extype
        Case 0, 1, 2, 4, 5, 6, 7
            SaveChargeItems Conn, ObjectID, Action, 1
            'Almacenando Rubros de Intercompany, solo guarda no ejecuta Webservice de BAW
            BAWResult = SaveInterChargeItems (Conn, ObjectID, BLType, Countries, 0, 1)
        End Select
	CloseOBJ Conn

    If Err.Number<>0 then
        response.write "SaveInterChargeItems : "  & Err.Number & " - " & Err.Description & "<br>"  
    end if

end if


'response.write "(" & EXID & ")(" & ExType & ")(" & EXID_Ant & ")(" & CheckNum(Request("GID")) & ")<br>"

if CheckNum(Request("GID")) = 27 and ExType = 99 and EXID_Ant > 0 then
    EXID = EXID_Ant
end if


'Si EXID <> 0 debe recargar los datos que se ingresaron en el BL o RO, por si estos fueron modificados en Trafico Maritimo o Ventas
'response.write(ExType)
if EXID <> 0 then
     'Para las cargas Maritmas, Aereas o Intermodal que se mueven en Terrestre Local se le asigna la fecha que registran el documento al sistema
     Select Case ExType
     Case 9,10,11,12,13,14
        ReservationDate = Time
		ReservationDate = Hour(ReservationDate) & ":" & TwoDigits(Minute(ReservationDate)) & ":" & TwoDigits(Second(ReservationDate))
     End Select

	 'Al no indicarse el pais, se selecciona el primer pais asignado al usuario
	 if Countries = "" then
		Countries = SetDefaultCountry
	 end if
	 select Case ExType 'Carga Maritima En Transito
	 case 0,1,2,11,12,13
		 select Case ExType
		 case 0,11 'FCL
			 QuerySelect = "select B.fecha_descarga, B.id_cliente, C.comodity_id, C.peso, C.volumen, C.no_piezas, B.id_destino_final, B.id_shipper, B.user_id, C.no_contenedor, B.no_bl, B.mbl, B.agente_id, B.id_puerto_origen, C.id_tipo_paquete, B.id_routing, B.id_incoterms, B.id_colectar, B.id_coloader ab, B.id_cliente_order from bl_completo B, contenedor_completo C where B.bl_id = C.bl_id and B.activo and C.activo and C.contenedor_id=" & EXID
			 SQLQuery = "select a.id_moneda, a.id_rubro, a.valor_prepaid, a.valor_collect, a.local, '', a.valor_sobreventa, a.id_servicio, '', a.inter_company from cargos_bl a, contenedor_completo b where a.bl_id=b.bl_id and b.contenedor_id=" & EXID & " and a.tipo_bl='F' and a.factura_id=0 and a.activo=true"
		 case 1,12 'LCL Sin Division
			 QuerySelect = "select VC.fecha_descarga, B.id_cliente, B.comodity_id, B.peso, B.volumen, B.no_piezas, B.id_destino_final, B.id_shipper, V.user_id, VC.no_contenedor, B.no_bl, VC.mbl, V.agente_id, V.id_puerto_origen, B.id_tipo_paquete, B.id_routing, B.id_incoterms, B.id_colectar, B.id_coloader cd, B.id_cliente_order from viaje_contenedor VC, viajes V, bill_of_lading B where B.viaje_contenedor_id = VC.viaje_contenedor_id and VC.viaje_id = V.viaje_id and B.activo and VC.activo and V.activo and B.bl_id=" & EXID
			 SQLQuery = "select a.id_moneda, a.id_rubro, a.valor_prepaid, a.valor_collect, a.local, '', a.valor_sobreventa, a.id_servicio, '', a.inter_company from cargos_bl a where bl_id=" & EXID & " and tipo_bl='L' and a.factura_id=0 and a.activo=true"
		 case 2,13 'LCL Con Division
			 QuerySelect = "select VC.fecha_descarga, DB.id_cliente, DB.comodity_id, DB.peso, DB.volumen, DB.no_bultos, B.id_destino_final, B.id_shipper, V.user_id, VC.no_contenedor, B.no_bl || ' (' || DB.no_bl || ')', VC.mbl, V.agente_id, V.id_puerto_origen, B.id_tipo_paquete, B.id_routing, B.id_incoterms, B.id_colectar, B.id_coloader ef, B.id_cliente_order from viaje_contenedor VC, viajes V, bill_of_lading B, divisiones_bl DB where DB.bl_asoc=B.bl_id and B.viaje_contenedor_id = VC.viaje_contenedor_id and VC.viaje_id = V.viaje_id and B.activo and VC.activo and V.activo and DB.division_id=" & EXID
			 SQLQuery = "select a.id_moneda, a.id_rubro, a.valor_prepaid, a.valor_collect, a.local, '', a.valor_sobreventa, a.id_servicio, '', a.inter_company from cargos_bl a where bl_id=0 and tipo_bl='L' and a.factura_id=0 and a.activo=true"
		 end select
		'Se conecta a la base de ventas correspondiente para obtener la informacion del BL en Transito
		OpenConnOcean Conn, "ventas_" & setDBCountry(CountryDes)
		'OpenConnOcean Conn, "ventas_sv" 
		'response.write setDBCountry(CountryDes)
        
        'response.write QuerySelect & "<br>"
        'response.write("ventas_" & setDBCountry(CountryDes))
		Set rs = Conn.Execute(QuerySelect)
		if Not rs.EOF then
			WareHouseDischargeDate = ConvertDate(rs(0),1)
			if ClientID=0 then
				ClientID = CheckNum(rs(1))
			end if
			'if CommodityCode=0 and ExType=2 then
				CommodityCode = CheckNum(rs(2))
			'end if
            Weight = CheckNum(rs(3))
			Volume = CheckNum(rs(4))
			TotNoOfPieces = CheckNum(rs(5))
			if FinalDes="" then
				FinalDes = CheckNum(rs(6))
			end if
            if CheckNum(rs(7))<>0 and ShipperID=0 then
				ShipperID = CheckNum(rs(7))
			end if
				
            ContactSignature = CheckNum(rs(8))
			Container = rs(9)
			BL = rs(10)
			MBL = rs(11)
			if CheckNum(rs(12))<>0 and AgentID=0 then
				AgentID = CheckNum(rs(12))
			end if				
			if CountryOrigen="" then
				CountryOrigen = CheckNum(rs(13))
			end if
			Contener = CheckNum(rs(14))
			RoutingID = CheckNum(rs(15))
            IncotermsID = CheckNum(rs(16))
            ClientCollectID = CheckNum(rs(17))
            if CheckNum(rs(18))<>0 and ColoaderID=0 then
				    ColoaderID = CheckNum(rs(18))
			end if
            NotifyPartyID = CheckNum(rs(19))
		end if
		CloseOBJ rs
		'Obteniendo los rubros del BL
		'response.write SQLQuery & "<br>"
		Set rs = Conn.Execute(SQLQuery)
		if Not rs.EOF then
			aList1Values = rs.GetRows
			CountList1Values = rs.RecordCount-1
		end if
		CloseOBJs rs, Conn
    
    case 9,10 'Carga Aerea En Transito
		 select Case ExType
		 case 9 'Import
			 QuerySelect = "select a.AWBDate, a.ConsignerID, a.Commodities, a.TotWeightChargeable, 0, a.NoOfPieces, a.Countries, a.ShipperID, a.UserID, a.Voyage, a.HAWBNumber, a.AWBNumber, a.AgentID, b.Country, 4, a.RoutingID, 0 from Awbi a, Airports b where a.AirportDepID=b.AirportID and AWBID=" & EXID
             'Aereo no tiene tipo paquete se dejo seteado en el numero 4 (bultos)
             'El primer 0 es Volumen pero Aereo no tiene
             'el ultimo 0 es Incoterms pero Aereo no tiene
		 end select
		
        OpenConnAir Conn
		'response.write QuerySelect & " *** " & ExType & "<br>"
		Set rs = Conn.Execute(QuerySelect)
		if Not rs.EOF then
			WareHouseDischargeDate = ConvertDate(rs(0),1)
			if ClientID=0 then
				ClientID = CheckNum(rs(1))
			end if
			'if CommodityCode=0 and ExType=2 then
				CommodityCode = CheckNum(rs(2))
			'end if
            Weight = CheckNum(rs(3))
			Volume = CheckNum(rs(4))
			TotNoOfPieces = CheckNum(rs(5))
			FinalDes = rs(6)
			if CheckNum(rs(7))<>0 and ShipperID=0 then
				ShipperID = CheckNum(rs(7))
			end if
				
            ContactSignature = CheckNum(rs(8))
			Container = rs(9)
			BL = rs(10)
			MBL = rs(11)
			if CheckNum(rs(12))<>0 and AgentID=0 then
				AgentID = CheckNum(rs(12))
			end if				
			CountryOrigen = rs(13)
			Contener = CheckNum(rs(14))
            RoutingID = CheckNum(rs(15))
		end if
		CloseOBJs rs, Conn
    
    case 14 'Carga Intermodal
		QuerySelect = "select a.DischargeDate, a.ClientsID, a.CommoditiesID, a.Weights, a.Volumes, a.NoOfPieces, a.Countries, a.AgentsID, b.UserID, '', a.HBLNumber, b.BLNumber, a.ShippersID, a.CountriesFinalDes, a.ClassNoOfPieces, 0, a.IncotermsID, a.ClientCollectID, a.ColoadersID, a.ColoadersAddrID, a.Coloaders from BLDetail a, BLs b where a.BLID=b.BLID and BLDetailID=" & EXID
        
        OpenConn Conn
		'response.write QuerySelect & "<br>"
		Set rs = Conn.Execute(QuerySelect)
		if Not rs.EOF then
			WareHouseDischargeDate = ConvertDate(rs(0),1)
			if ClientID=0 then
				ClientID = CheckNum(rs(1))
			end if
			'if CommodityCode=0 and ExType=2 then
				CommodityCode = CheckNum(rs(2))
			'end if
            Weight = CheckNum(rs(3))
			Volume = CheckNum(rs(4))
			TotNoOfPieces = CheckNum(rs(5))
			FinalDes = rs(6)
			if CheckNum(rs(7))<>0 and ShipperID=0 then
				ShipperID = CheckNum(rs(7))
			end if
				
            ContactSignature = CheckNum(rs(8))
			Container = rs(9)
			BL = rs(10)
			MBL = rs(11)
			if CheckNum(rs(12))<>0 and AgentID=0 then
				AgentID = CheckNum(rs(12))
			end if				
			CountryOrigen = rs(6)
			Contener = rs(14)
			RoutingID = CheckNum(rs(15))
            IncotermsID = CheckNum(rs(16))
            ClientCollectID = CheckNum(rs(17))
            ColoaderID = CheckNum(rs(18))
            ColoaderAddrID = CheckNum(rs(19))
            ColoaderData = CheckNum(rs(20))
		end if
    	CloseOBJs rs, Conn
	end select

    
    Dim ColgateData, ClientsAddTemp, AgentsAddTemp, gS_MessageID, TransactionSetPurposeCode, ShipmentIDNumber, tch_BLDetailID

    ColgateData = "" 
    ClientsAddTemp = "" 
    AgentsAddTemp = ""
    gS_MessageID = ""
    TransactionSetPurposeCode = "" 
    ShipmentIDNumber = ""
    tch_BLDetailID = ""

    
	OpenConn2 Conn
	Select Case ExType 'RO Terrestre Consolidado/Express/Local (Carga General)

    Case 99

    QuerySelect = "SELECT  " & vbCrLf & _
" --	0					1						2					3					4								5								6								7								8						9							10									11										12						13							14							15								16				" & vbCrLf & _
"a.""tch_pk"", a.""tch_BLDetailID"", a.""tch_gS_MessageID"", a.""tch_fecha"", a.""tch_gS_TranmissionDate"", a.""tch_gS_TransmissionTime"", a.""tch_b2_ShipmentIDNumber"", a.""tch_b2_WeightUnitCode"", a.""tch_b2_ShipmentQualifier"", a.""tch_SetPurposeCode"", a.""tch_mS3_RoutingSequenceCode"", a.""tch_mS3_TransportationMethodCode"", a.""tch_ntE_Instructions"", a.""tch_l3_Weight"", a.""tch_l3_WeightQualifier"", a.""tch_l3_LandingQuantity"", a.""tch_l3_WeightUnitQualifier"",  " & vbCrLf & _

" -- 			17						18					19						20								21										22							23							24					25											26								27								28									29								30							31					32								33						34						35						36							37 " & vbCrLf & _
"s.""tcd_n1_EntityIDCode"", s.""tcd_n1_Name"", s.""tcd_n7_Weight"", s.""tcd_n7_WeightQualifier"", s.""tcd_n7_EquipmentDescriptionCode"", s.""tcd_n7_EquipmentLength"", s.""tcd_s5_StopOffReasonCode"", s.""tcd_s5_Weight"", s.""tcd_s5_WeightUnitQualifier"", s.""tcd_s5_NumberofUnitsShipped"", s.""tcd_s5_UnitofMeasureCode"", s.""tcd_l11_ReferenceNumber"", s.""tcd_l11_ReferenceNumberQualifier"", s.""tcd_0g62_DateQualifier"", s.""tcd_0g62_Date"", s.""tcd_0g62_TimeQualifier"", s.""tcd_0g62_Time"", s.""tcd_1g62_DateQualifier"", s.""tcd_1g62_Date"", s.""tcd_1g62_TimeQualifier"", s.""tcd_1g62_Time"",  " & vbCrLf & _

" -- 			38						39					40							41							42							43						44						45						46							47							48				49					50									51  " & vbCrLf & _
"f.""tcd_n1_EntityIDCode"", f.""tcd_n1_Name"", f.""tcd_0g62_DateQualifier"", f.""tcd_0g62_Date"", f.""tcd_0g62_TimeQualifier"", f.""tcd_0g62_Time"", f.""tcd_1g62_DateQualifier"", f.""tcd_1g62_Date"", f.""tcd_1g62_TimeQualifier"", f.""tcd_1g62_Time"", f.""tcd_n3_Address"", f.""tcd_n4_CityName"", f.""tcd_n4_StateorProvinceCode"", f.""tcd_n4_PostalCode"",  " & vbCrLf & _

" -- 				52					53				54							55						56									57					" & vbCrLf & _
"t.""tcd_n1_EntityIDCode"", t.""tcd_n1_Name"", t.""tcd_n3_Address"", t.""tcd_n4_CityName"", t.""tcd_n4_StateorProvinceCode"", t.""tcd_n4_PostalCode"", " & vbCrLf & _

"        CASE WHEN a.""tch_SetPurposeCode"" <> '00' THEN " & vbCrLf & _

"           (SELECT b.""tch_BLDetailID"" FROM ti_colgate_header b WHERE b.""tch_b2_ShipmentIDNumber"" = a.""tch_b2_ShipmentIDNumber"" AND b.""tch_SetPurposeCode"" <> '02' AND b.""tch_SetPurposeCode"" <> a.""tch_SetPurposeCode"" ORDER BY b.""tch_SetPurposeCode"" DESC LIMIT 1)" & vbCrLf & _

"        ELSE -1 END as t, ""tch_b2A_TransactionSetPurposeCode"", " & vbCrLf & _

"t.""tcd_n3_1Address"", t.""tcd_n4_1CityName"", t.""tcd_n4_1StateorProvinceCode"", t.""tcd_n4_1PostalCode"" " & vbCrLf & _

"FROM ti_colgate_header a " & vbCrLf & _

"LEFT JOIN ti_colgate_details s ON s.""tcd_tch_fk"" = a.tch_pk AND s.""tcd_n1_EntityIDCode"" = 'SH - Shipper' " & vbCrLf & _

"LEFT JOIN ti_colgate_details f ON f.""tcd_tch_fk"" = a.tch_pk AND f.""tcd_n1_EntityIDCode"" = 'SF - Ship From' " & vbCrLf & _

"LEFT JOIN ti_colgate_details t ON t.""tcd_tch_fk"" = a.tch_pk AND t.""tcd_n1_EntityIDCode"" = 'CN - Ship To' " & vbCrLf & _

"WHERE a.tch_pk = " & EXID & vbCrLf 'CheckNum(Request("EID"))


        'response.write QuerySelect & "<br>"

        Set rs = Conn.Execute(QuerySelect)
        if Not rs.EOF then   

            'if CheckNum(Request("OID")) > 0 and CheckNum(Request("EID")) > 0 then
            if CheckNum(Request("SetPurposeCode")) > 0 then
    			TransactionSetPurposeCode = Request("SetPurposeCode")
            else
                TransactionSetPurposeCode = IFNULL(rs(9))
            end if

           'response.write "(" & TransactionSetPurposeCode & ")(" & rs(9) & ")<br>"

           if TransactionSetPurposeCode = "02" then
    			tch_BLDetailID = CheckNum(Request("OID"))
           else
    			tch_BLDetailID = IFNULL(rs(58))
           end if



            if CheckNum(ObjectID) = 0 then
                DiceContenerTemp = "ARTICULOS DE HIGIENE"
				ClientsTemp = Replace(IFNULL(rs(53)),", S.A.","")				
				AgentsTemp = Replace(Replace(Replace(IFNULL(rs(39)),"(",""),")",""),"-"," ")				
				Weight = IFNULL(rs(24))
				TotNoOfPieces = IFNULL(rs(26))
				PO = IFNULL(rs(28))
				Container = IFNULL(rs(6))
				BL = IFNULL(rs(6))
				MBL = IFNULL(rs(6))    
				WareHouseDischargeDate = right(IFNULL(rs(41)),2) & "/" & mid(IFNULL(rs(41)),5,2) & "/" & left(IFNULL(rs(41)),4)
            end if


				gS_MessageID = IFNULL(rs(2))

				AgentsAddTemp = IFNULL(rs(48)) & " " & IFNULL(rs(49)) & " " & IFNULL(rs(50)) & " " & IFNULL(rs(51)) 								
				ClientsAddTemp = IFNULL(rs(54)) & " " & IFNULL(rs(55)) & " " & IFNULL(rs(56)) & " " & IFNULL(rs(57)) 				
				ShipmentIDNumber = IFNULL(rs(6))

ColgateData = "<TR><TD class=label colspan=2 align=center><h4><button onclick=if(document.getElementById('ShipC').style.display=='none'){document.getElementById('ShipC').style.display='inline'}else{document.getElementById('ShipC').style.display='none'};return(false);>Shipment Colgate</button></h4></TD></TR>" & _ 

        "<TR><TD class=label colspan=2 align=center>" & _ 

        "<TABLE style='margin:3px;border:0px solid blue;width:80%;display:none' id='ShipC'>" & _ 

        "<TR><TH class=label colspan=2><hr></TD></TR>" & _ 

        "<TR><TH class=label colspan=2>" & _ 

            "<TABLE style='margin:3px;width:100%;background-color:rgb(247,252,253)'>" & _ 

                "<TR><TD class=label align=right><b> #</b></TD><TD class=label align=left>" & IFNULL(rs(0)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> MessageID:</b></TD><TD class=label align=left>" & gS_MessageID & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> TransactionSetPurposeCode:</b></TD><TD class=label align=left>" & IFNULL(rs(59)) & "</TD></TR>" & _ 

                "<TR><TH class=label colspan=2><b><br>GroupHeader</b></TD></TR>" & _ 

                "<TR><TD class=label align=right><b> TranmissionDate:</b></TD><TD class=label align=left>" & IFNULL(rs(4)) & "</TD></TR>" & _ 
	  
                "<TR><TD class=label align=right><b> TransmissionTime:</b></TD><TD class=label align=left>" & IFNULL(rs(5)) & "</TD></TR>" & _ 
	  
                "<TR><TH class=label colspan=2><br>ShipmentInformation<b></TD></TR>" & _ 
	  
                "<TR><TD class=label align=right><b> ShipmentIDNumber:</b></TD><TD class=label align=left>" & IFNULL(rs(6)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> WeightUnitCode:</b></TD><TD class=label align=left>" & IFNULL(rs(7)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> ShipmentQualifier:</b></TD><TD class=label align=left>" & IFNULL(rs(8)) & "</TD></TR>" & _ 
	  
                "<TR><TH class=label colspan=2><br>InterlineInformation<b></TD></TR>" & _ 
	  
                "<TR><TD class=label align=right><b> RoutingSequenceCode:</b></TD><TD class=label align=left>" & IFNULL(rs(10)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> TransportationMethodCode:</b></TD><TD class=label align=left>" & IFNULL(rs(11)) & "</TD></TR>" & _ 
	  
                "<TR><TH class=label colspan=2><br>Notes<b></TD></TR>" & _ 
	  
                "<TR><TD class=label align=right><b> Instructions:</b></TD><TD class=label align=left>" & IFNULL(rs(12)) & "</TD></TR>" & _ 

                "<TR><TH class=label colspan=2><br>TotalWeightandCharges<b></TD></TR>" & _ 
	  
                "<TR><TD class=label align=right><b> Weight:</b></TD><TD class=label align=left>" & IFNULL(rs(13)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> WeightQualifier:</b></TD><TD class=label align=left>" & IFNULL(rs(14)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> LandingQuantity:</b></TD><TD class=label align=left>" & IFNULL(rs(15)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> WeightUnitQualifier:</b></TD><TD class=label align=left>" & IFNULL(rs(16)) & "</TD></TR>" & _ 

            "</TABLE>" & _ 

        "</TD></TR>" & _ 

        "<TR><TH class=label colspan=2><hr></TD></TR>" & _ 

        "<TR><TH class=label colspan=2>" & _ 

            "<TABLE style='margin:3px;width:100%;background-color:rgb(247,252,253)'>" & _ 
         
                " <!-- //////////////////////////////////////////////////////////////////////SHIPPER/////////////////////////////////////-->" & _         
                "<TR><TH class=label colspan=2><br>Shipper<b></TD></TR>" & _ 
                "<TR><TD class=label align=right><b> EntityIDCode:</b></TD><TD class=label align=left>" & IFNULL(rs(17)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> Name:</b></TD><TD class=label align=left>" & IFNULL(rs(18)) & "</TD></TR>" & _ 
	  
                "<TR><TH class=label colspan=2><br>EquipmentDetails<b></TD></TR>" & _ 
	  
                "<TR><TD class=label align=right><b> Weight:</b></TD><TD class=label align=left>" & IFNULL(rs(19)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> WeightQualifier:</b></TD><TD class=label align=left>" & IFNULL(rs(20)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> EquipmentDescriptionCode:</b></TD><TD class=label align=left>" & IFNULL(rs(21)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> EquipmentLength:</b></TD><TD class=label align=left>" & IFNULL(rs(22)) & "</TD></TR>" & _ 
	  
                "<TR><TH class=label colspan=2><br>StopOffDetails<b></TD></TR>" & _ 
	  
                "<TR><TD class=label align=right><b> StopOffReasonCode:</b></TD><TD class=label align=left>" & IFNULL(rs(23)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> Weight:</b></TD><TD class=label align=left>" & IFNULL(rs(24)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> WeightUnitQualifier:</b></TD><TD class=label align=left>" & IFNULL(rs(25)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> NumberofUnitsShipped:</b></TD><TD class=label align=left>" & IFNULL(rs(26)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> UnitofMeasureCode:</b></TD><TD class=label align=left>" & IFNULL(rs(27)) & "</TD></TR>" & _ 

                "<TR><TH class=label colspan=2><br>BusinessInstructionsAndReferenceNumber<b></TD></TR>" & _ 
	  
                "<TR><TD class=label align=right><b> ReferenceNumber:</b></TD><TD class=label align=left>" & IFNULL(rs(28)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> ReferenceNumberQualifier:</b></TD><TD class=label align=left>" & IFNULL(rs(29)) & "</TD></TR>" & _ 
	  
                "<TR><TH class=label colspan=2><br>DateTime <b></TD></TR>" & _ 
	  
                "<TR><TD class=label align=right><b> DateQualifier:</b></TD><TD class=label align=left>" & IFNULL(rs(30)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> Date:</b></TD><TD class=label align=left>" & IFNULL(rs(31)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> TimeQualifier:</b></TD><TD class=label align=left>" & IFNULL(rs(32)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> Time:</b></TD><TD class=label align=left>" & IFNULL(rs(33)) & "</TD></TR>" & _ 
	  
                "<TR><TH class=label colspan=2><br>DateTime <b></TD></TR>" & _ 

                "<TR><TD class=label align=right><b> DateQualifier:</b></TD><TD class=label align=left>" & IFNULL(rs(34)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> Date:</b></TD><TD class=label align=left>" & IFNULL(rs(35)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> TimeQualifier:</b></TD><TD class=label align=left>" & IFNULL(rs(36)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> Time:</b></TD><TD class=label align=left>" & IFNULL(rs(37)) & "</TD></TR>" & _ 
            "</TABLE>" & _ 

        "</TD></TR>" & _ 

        "<TR><TH class=label colspan=2><hr></TD></TR>" & _ 

        "<TR><TH class=label colspan=2>" & _ 

            "<TABLE style='margin:3px;width:100%;background-color:rgb(247,252,253)'>" & _ 

                " <!-- //////////////////////////////////////////////////////////////////////SHIP FROM/////////////////////////////////////-->" & _ 
                "<TR><TH class=label colspan=2><br>Ship From<b></TD></TR>" & _ 
                "<TR><TD class=label align=right><b> EntityIDCode:</b></TD><TD class=label align=left>" & IFNULL(rs(38)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> Name:</b></TD><TD class=label align=left>" & IFNULL(rs(39)) & "</TD></TR>" & _ 
	  
                "<TR><TH class=label colspan=2><br>DateTime <b></TD></TR>" & _ 
	  
                "<TR><TD class=label align=right><b> DateQualifier:</b></TD><TD class=label align=left>" & IFNULL(rs(40)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> Date:</b></TD><TD class=label align=left>" & IFNULL(rs(41)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> TimeQualifier:</b></TD><TD class=label align=left>" & IFNULL(rs(42)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> Time:</b></TD><TD class=label align=left>" & IFNULL(rs(43)) & "</TD></TR>" & _ 
	  
                "<TR><TH class=label colspan=2><br>DateTime <b></TD></TR>" & _ 

                "<TR><TD class=label align=right><b> DateQualifier:</b></TD><TD class=label align=left>" & IFNULL(rs(44)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> Date:</b></TD><TD class=label align=left>" & IFNULL(rs(45)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> TimeQualifier:</b></TD><TD class=label align=left>" & IFNULL(rs(46)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> Time:</b></TD><TD class=label align=left>" & IFNULL(rs(47)) & "</TD></TR>" & _ 

                "<TR><TH class=label colspan=2><br>AddressInformation<b></TD></TR>" & _ 

                "<TR><TD class=label align=right><b> Address:</b></TD><TD class=label align=left>" & IFNULL(rs(48)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> CityName:</b></TD><TD class=label align=left>" & IFNULL(rs(49)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> StateorProvinceCode:</b></TD><TD class=label align=left>" & IFNULL(rs(50)) & "</TD></TR>" & _ 
                "<TR><TD class=label align=right><b> PostalCode:</b></TD><TD class=label align=left>" & IFNULL(rs(51)) & "</TD></TR>" & _ 
            
            "</TABLE>" & _ 

        "</TD></TR>" & _ 

        "<TR><TH class=label colspan=2><hr></TD></TR>" & _ 

        "<TR><TH class=label colspan=2>" & _ 

            "<TABLE style='margin:3px;width:100%;background-color:rgb(255,241,193)'>" & _ 

            " <!-- //////////////////////////////////////////////////////////////////////SHIP TO/////////////////////////////////////-->" & _
            "<TR><TH class=label colspan=2><br>Ship To<b></TD></TR>" & _ 
            "<TR><TD class=label align=right><b> EntityIDCode:</b></TD><TD class=label align=left>" & IFNULL(rs(52)) & "</TD></TR>" & _ 
            "<TR><TD class=label align=right><b> Name:</b></TD><TD class=label align=left>" & IFNULL(rs(53)) & "</TD></TR>" & _ 
	  
            "<TR><TH class=label colspan=2><br>AddressInformation<b></TD></TR>" & _ 

            "<TR><TD class=label align=right><b> Address:</b></TD><TD class=label align=left>" & IFNULL(rs(54)) & "</TD></TR>" & _ 
            "<TR><TD class=label align=right><b> CityName:</b></TD><TD class=label align=left>" & IFNULL(rs(55)) & "</TD></TR>" & _ 
            "<TR><TD class=label align=right><b> StateorProvinceCode:</b></TD><TD class=label align=left>" & IFNULL(rs(56)) & "</TD></TR>" & _ 
            "<TR><TD class=label align=right><b> PostalCode:</b></TD><TD class=label align=left>" & IFNULL(rs(57)) & "</TD></TR>" & _


            "<TR><TH class=label colspan=2><br>AddressInformation1<b></TD></TR>" & _ 

            "<TR><TD class=label align=right><b> Address:</b></TD><TD class=label align=left>" & IFNULL(rs("tcd_n3_1Address")) & "</TD></TR>" & _ 
            "<TR><TD class=label align=right><b> CityName:</b></TD><TD class=label align=left>" & IFNULL(rs("tcd_n4_1CityName")) & "</TD></TR>" & _ 
            "<TR><TD class=label align=right><b> StateorProvinceCode:</b></TD><TD class=label align=left>" & IFNULL(rs("tcd_n4_1StateorProvinceCode")) & "</TD></TR>" & _ 
            "<TR><TD class=label align=right><b> PostalCode:</b></TD><TD class=label align=left>" & IFNULL(rs("tcd_n4_1PostalCode")) & "</TD></TR>" & _

            "</TABLE>" & _ 

        "</TD></TR>" & _ 

        "<TR><TH class=label colspan=2><hr></TD></TR>" & _ 

        "</TABLE>" & _ 

        "</TD></TR>"
	  
      ColgateData = replace(ColgateData,"CityName:","Ciudad:")
ColgateData = replace(ColgateData,"Name:","Nombre:")

ColgateData = replace(ColgateData,"MessageID:","MesssageID:")
ColgateData = replace(ColgateData,"TransactionSetPurposeCode:","Transaccion:")
ColgateData = replace(ColgateData,"GroupHeader","Encabezado")
ColgateData = replace(ColgateData,"TranmissionDate:","Fecha Transmision:")
ColgateData = replace(ColgateData,"TransmissionTime:","Hora Transmision:")
ColgateData = replace(ColgateData,"ShipmentInformation","Informacion de Shipment")
ColgateData = replace(ColgateData,"ShipmentIDNumber:","No. Shipment:")
ColgateData = replace(ColgateData,"WeightUnitCode:","Unidad de Peso:")
ColgateData = replace(ColgateData,"ShipmentQualifier:","Calificador de Embarque:")
ColgateData = replace(ColgateData,"InterlineInformation","Informacion de Carrier")
ColgateData = replace(ColgateData,"RoutingSequenceCode:","Enrutado por:")
ColgateData = replace(ColgateData,"TransportationMethodCode:","Tipo Transporte:")
ColgateData = replace(ColgateData,"Notes","Nota:")
ColgateData = replace(ColgateData,"Instructions:","Instrucciones:")
ColgateData = replace(ColgateData,"TotalWeightandCharges","Datos de Carga")
ColgateData = replace(ColgateData,"Weight:","Peso:")
ColgateData = replace(ColgateData,"WeightQualifier:","Unidad de Peso:")
ColgateData = replace(ColgateData,"LandingQuantity:","Cantidad:")
ColgateData = replace(ColgateData,"WeightUnitQualifier:","Unidad de Medida:")


'///////////////////////////////////////////////////////////////////////////SHIPPER
ColgateData = replace(ColgateData,"Shipper","Embarcador:")
ColgateData = replace(ColgateData,"EntityIDCode:","Tipo de Embarcador:")

ColgateData = replace(ColgateData,"EquipmentDetails","Detalle de Equipo")
ColgateData = replace(ColgateData,"Weight:","Peso:")
ColgateData = replace(ColgateData,"WeightQualifier:","Unidad de Peso:")
ColgateData = replace(ColgateData,"EquipmentDescriptionCode:","Tipo de Equipo:")
ColgateData = replace(ColgateData,"EquipmentLength:","Longitud de Equipo:")
ColgateData = replace(ColgateData,"StopOffDetails","Detalle de Entrega")
ColgateData = replace(ColgateData,"StopOffReasonCode:","Actividad:")
ColgateData = replace(ColgateData,"Weight:","Peso:")
ColgateData = replace(ColgateData,"WeightUnitQualifier:","Unidad de Peso:")
ColgateData = replace(ColgateData,"NumberofUnitsShipped:","Unidades a Enviar:")
ColgateData = replace(ColgateData,"UnitofMeasureCode:","Unidad de Medida:")
ColgateData = replace(ColgateData,"BusinessInstructionsAndReferenceNumber","Instrucciones de Negocio")
ColgateData = replace(ColgateData,"ReferenceNumber:","No. Referencia:")
ColgateData = replace(ColgateData,"ReferenceNumberQualifier:","Tipo Referencia:")
ColgateData = replace(ColgateData,"DateTime","Tiempos de Servicio")
ColgateData = replace(ColgateData,"DateQualifier:","Condicion de Fecha:")
ColgateData = replace(ColgateData,"Date:","Fecha:")
ColgateData = replace(ColgateData,"TimeQualifier:","Condicion de Hora:")
ColgateData = replace(ColgateData,"Time:","Hora:")
ColgateData = replace(ColgateData,"DateQualifier:","Condicion de Fecha:")
ColgateData = replace(ColgateData,"Date:","Fecha:")
ColgateData = replace(ColgateData,"TimeQualifier:","Condicion de Hora:")
ColgateData = replace(ColgateData,"Time:","Hora:")


'///////////////////////////////////////////////////////////////////////////SHIP FROM
ColgateData = replace(ColgateData,"Ship From","Salida")
ColgateData = replace(ColgateData,"EntityIDCode:","Tipo:")
ColgateData = replace(ColgateData,"Name:","Origen:")
ColgateData = replace(ColgateData,"DateTime","Tiempos de Servicio")
ColgateData = replace(ColgateData,"DateQualifier:","Condicion de Fecha:")
ColgateData = replace(ColgateData,"Date:","Fecha:")
ColgateData = replace(ColgateData,"TimeQualifier:","Condicion de Hora:")
ColgateData = replace(ColgateData,"Time:","Hora:")
ColgateData = replace(ColgateData,"DateQualifier:","Condicion de Fecha:")
ColgateData = replace(ColgateData,"Date:","Fecha:")
ColgateData = replace(ColgateData,"TimeQualifier:","Condicion de Hora:")
ColgateData = replace(ColgateData,"Time:","Hora:")

ColgateData = replace(ColgateData,"AddressInformation1","Destino 2")

ColgateData = replace(ColgateData,"AddressInformation","Destino 1")
ColgateData = replace(ColgateData,"Address:","Direccion:")
ColgateData = replace(ColgateData,"StateorProvinceCode:","Provincia:")
ColgateData = replace(ColgateData,"PostalCode:","Codigo Postal:")



'///////////////////////////////////////////////////////////////////////////SHIP TO
ColgateData = replace(ColgateData,"Ship To","Destino")
ColgateData = replace(ColgateData,"EntityIDCode:","Tipo:")
ColgateData = replace(ColgateData,"Name:","Destino:")

'ColgateData = replace(ColgateData,"Address:","Direccion:")
'ColgateData = replace(ColgateData,"CityName:","Ciudad:")
'ColgateData = replace(ColgateData,"StateorProvinceCode:","Provincia:")
'ColgateData = replace(ColgateData,"PostalCode:","Codigo Postal:")


		end if	

    


	Case 4,5,6,7
		Select Case ExType
		case 4,5 '4=Consolidado, 5=Express

			'QuerySelect = "select a.fecha, a.id_cliente, a.comodity_id, a.peso, a.volumen, a.no_piezas, a.id_pais_destino, a.id_shipper, a.vendedor_id, a.routing, not(e.prepaid), a.agente_id, a.id_pais_origen, a.id_tipo_paquete, a.notificar_a, a.id_coloader, a.id_incoterms, a.id_colectar, a.routing_seg, a.seguro, coalesce(a.poliza_seguro, ''), b.borrado, coalesce(b.poliza_seguro, ''), b.bl_id, a.routing_adu, c.activo, c.borrado, d.numero_dua, a.id_cliente_order, a.routing_cli         from routings a left join routings b on a.routing_seg=b.id_routing left join routings c on a.routing_adu=c.id_routing left join routings_dua d on a.routing_adu=d.id_routing inner join cargos_routing e on a.id_routing = e.id_routing where a.id_routing=" & EXID)
		    QuerySelect = "select a.fecha, a.id_cliente, a.comodity_id, a.peso, a.volumen, a.no_piezas, a.id_pais_destino, a.id_shipper, a.vendedor_id, a.routing, not(e.prepaid), a.agente_id, a.id_pais_origen, a.id_tipo_paquete, a.notificar_a, a.id_coloader, a.id_incoterms, a.id_colectar, a.routing_seg, a.seguro, coalesce(a.poliza_seguro, ''), b.borrado, coalesce(b.poliza_seguro, ''), b.bl_id, a.routing_adu, c.activo, c.borrado, d.numero_dua, a.id_cliente_order, a.routing_cli, a.file from routings a left join routings b on a.routing_seg=b.id_routing left join routings c on a.routing_adu=c.id_routing left join routings_dua d on a.routing_adu=d.id_routing inner join cargos_routing e on a.id_routing = e.id_routing where a.id_routing=" & EXID
        
            'response.write QuerySelect & "<br>"

			'Se conecta a la base master donde estan los routings terrestres de carga consolidado o express
            Set rs = Conn.Execute(QuerySelect)
		case 6 'Recoleccion
			'Se conecta a la base master donde estan los routings terrestres de recoleccion
            'response.write("select a.fecha, a.id_cliente, a.comodity_id, a.peso, a.volumen, a.no_piezas, a.id_pais_destino, a.id_shipper, a.vendedor_id, a.routing, not(a.prepaid), a.agente_id, a.id_pais_origen, a.id_tipo_paquete, date_part('Hour',a.hora_ingreso)||':'||date_part('Minute',a.hora_ingreso)||':'||date_part('Second',a.hora_ingreso), b.bl_ro, b.fecha_recoleccion||' '||b.hora_recoleccion, b.contacto_recoleccion, b.telefono_recoleccion, b.direccion_recoleccion, b.direccion_entrega, b.notificar_entrega, b.no_poliza, b.id_almacen_recoleccion, b.id_almacen_entrega, b.contacto_entrega, a.id_incoterms, a.routing_seg, a.seguro, coalesce(a.poliza_seguro, ''), c.borrado, coalesce(c.poliza_seguro, ''), c.bl_id, a.routing_adu, d.activo, d.borrado, e.numero_dua, a.id_cliente_order, a.id_coloader, a.routing_cli from routings a left join routing_terrestre b on a.id_routing=b.id_routing left join routings c on a.routing_seg=c.id_routing left join routings d on a.routing_adu=d.id_routing left join routings_dua e on a.routing_adu=e.id_routing where a.id_routing=" & EXID)
            Set rs = Conn.Execute("select a.fecha, a.id_cliente, a.comodity_id, a.peso, a.volumen, a.no_piezas, a.id_pais_destino, a.id_shipper, a.vendedor_id, a.routing, not(a.prepaid), a.agente_id, a.id_pais_origen, a.id_tipo_paquete, date_part('Hour',a.hora_ingreso)||':'||date_part('Minute',a.hora_ingreso)||':'||date_part('Second',a.hora_ingreso), b.bl_ro, b.fecha_recoleccion||' '||b.hora_recoleccion, b.contacto_recoleccion, b.telefono_recoleccion, b.direccion_recoleccion, b.direccion_entrega, b.notificar_entrega, b.no_poliza, b.id_almacen_recoleccion, b.id_almacen_entrega, b.contacto_entrega, a.id_incoterms, a.routing_seg, a.seguro, coalesce(a.poliza_seguro, ''), c.borrado, coalesce(c.poliza_seguro, ''), c.bl_id, a.routing_adu, d.activo, d.borrado, e.numero_dua, a.id_cliente_order, a.id_coloader, a.routing_cli, a.file from routings a left join routing_terrestre b on a.id_routing=b.id_routing left join routings c on a.routing_seg=c.id_routing left join routings d on a.routing_adu=d.id_routing left join routings_dua e on a.routing_adu=e.id_routing where a.id_routing=" & EXID)
		'case 7 'Entrega
		'	'Se conecta a la base master donde estan los routings terrestres de entrega
		'	Set rs = Conn.Execute("select a.fecha, a.id_cliente, a.comodity_id, a.peso, a.volumen, a.no_piezas, a.id_pais_destino, a.id_shipper, a.vendedor_id, a.routing, not(a.prepaid), a.agente_id, a.id_pais_origen, a.id_tipo_paquete, date_part('Hour',a.hora_ingreso)||':'||date_part('Minute',a.hora_ingreso)||':'||date_part('Second',a.hora_ingreso), b.bl_ro, b.fecha_entrega||' '||b.hora_entrega, b.contacto_entrega, b.telefono_entrega, b.direccion_recoleccion, b.direccion_entrega, a.notificar_a, b.no_poliza from routings a, routing_terrestre b where a.id_routing=b.id_routing and a.id_routing=" & EXID)
		end Select
		if Not rs.EOF then
			Select Case ExType
			Case 4,5
				WareHouseDischargeDate = ConvertDate(rs(0),1)
                if Notify="" then
					Notify = rs(14)
				end if
                Routing_Seg = rs(18)
                Seguro = rs(19)
                Poliza_Seguro = rs(20)
                RSeg_Borrado = rs(21)
                RSeg_Poliza = rs(22)
                RSeg_BLID = rs(23)
                Routing_Adu = rs(24)
                RAdu_Activo = rs(25)
                RAdu_Borrado = rs(26)
                RAdu_NDUA = rs(27)
                NotifyPartyID = CheckNum(rs(28))
                RosClientID = CheckNum(rs(29))
                file = rs(30)
			Case 6,7
				WareHouseDischargeDate = rs(0)
				if Notify="" then
					Notify = rs(21)
				end if
                PolicyNo = rs(22)
                Routing_Seg = rs(27)
                Seguro = rs(28)
                Poliza_Seguro = rs(29)
                RSeg_Borrado = rs(30)
                RSeg_Poliza = rs(31)
                RSeg_BLID = rs(32)
                NotifyPartyID = CheckNum(rs(37))
                ColoaderID = rs(38)
                RosClientID = CheckNum(rs(39))
                file = rs(40)
			End Select

		
            if CheckNum(rs(1))<>0 and ClientID=0 then
				ClientID = CheckNum(rs(1))
			end if
			
            if CommodityCode=0 then
				CommodityCode = CheckNum(rs(2))
			end if
			
            Weight = Round(CheckNum(rs(3)),2)
			Volume = Round(CheckNum(rs(4)),2)
			TotNoOfPieces = CheckNum(rs(5))
			
            if FinalDes = "" then
				FinalDes = rs(6)
			end if
			
            if CheckNum(rs(7))<>0 and ShipperID=0 then
				ShipperID = CheckNum(rs(7))
			end if
			
            ContactSignature = CheckNum(rs(8))
			'Container = ""
			BL = rs(9)
			'MBL = ""
			TotPrepaid = CheckNum(rs(10)) 'en la tabla routing Prepaid=true=1, pero en trafico y facturacion Prepaid=false=0
			
            if CheckNum(rs(11))<>0 and AgentID=0 then
				AgentID = CheckNum(rs(11))
			end if
			if CountryOrigen = "" then
				CountryOrigen = rs(12)
			end if
			Contener = CheckNum(rs(13))
            
            'Cuando el Shipper es Coloader, en esta casilla se consulta el "SubShipper" que viene con el Coloader
            Select Case ExType
            Case 4,5
                if CheckNum(rs(15))<>0 and ColoaderID=0 then
				    ColoaderID = CheckNum(rs(15))
			    end if
            Case 6,7
                if CheckNum(rs(33))<>0 and CheckNum(ColoaderID)=0 then
				    ColoaderID = CheckNum(rs(33))
			    end if
            End Select

			Select Case ExType
			Case 6,7
				ReservationDate = rs(14) 'hora del servicio
				MBL = rs(15) 'BL /AWB
				if DeliveryDate="" then
					BLArrivalDate = rs(16) 'fecha y hora del servicio
				else
					BLArrivalDate = DeliveryDate
				end if
				AgentSignature = rs(17) 'Contacto del Cliente
				Phone1 = rs(18) 'Telefono de Contacto
				SenderData = rs(19) 'Direccion Carga
				ConsignerData = rs(20) 'Direccion Descarga
				if rs(17) <> "" then
					SenderData = SenderData & "<BR><b>CONTACTO:</b> " & rs(17) 'Agregando Contacto de Carga
				end if
				if rs(25) <> "" then
					ConsignerData = ConsignerData & "<BR><b>CONTACTO:</b> " & rs(25) 'Agregando Contacto de Descarga
				end if                
                IncotermsID = CheckNum(rs(26))
            Case Else
                IncotermsID = CheckNum(rs(16))
                ClientCollectID = CheckNum(rs(17))
			End Select

            if RosClientID > 0 then '2019-08-22 Ticket#2019081331000136 — CARTAS DE PORTE TLA CON INICIALES CNI // ROS CRLTF 
            Select Case ExType
            Case 4,5,6,7                                                
                OpenConn Connx
                'response.write("select RefHBLNumber, RefBLID from BLDetail where RoClientID = " & RosClientID & " and RefHBLNumber <> '--' order by BLDetailID LIMIT 1")
                Set rs2 = Connx.Execute("select RefHBLNumber, RefBLID from BLDetail where RoClientID = " & RosClientID & " and RefHBLNumber <> '--' order by BLDetailID LIMIT 1")
                If Not rs2.EOF then
                    aList14Values = rs2.GetRows
                    CountList3Values = rs2.RecordCount-1
                    RefHBLNumber = aList14Values(0,0)
                    RefBLID = aList14Values(1,0)
                End if
                CloseOBJs rs2, Connx
            End Select
            end if

		end if
		CloseOBJ rs

		'Obteniendo los rubros del Routing

		QuerySelect = "select c.simbolo, a.id_rubro, a.valor, a.local, concat(b.desc_rubro_es,' (',a.id_rubro,')'), a.id_servicio, '', a.inter_company, not(a.prepaid) from cargos_routing a, rubros b, monedas c, routings d where d.id_routing=" & EXID & " and a.id_rubro=b.id_rubro and a.id_moneda=c.moneda_id and (case when d.routing_seg = 0 then a.id_routing=d.id_routing else (case when (select e.routing_fac from routings e where e.id_routing=d.routing_seg)<> 0 then a.id_routing=d.id_routing else a.id_routing in (d.id_routing,d.routing_seg) end) end) and d.borrado=false"
        
        'response.write QuerySelect & "<br>"

		Set rs = Conn.Execute(QuerySelect)
		if Not rs.EOF then
			aList1Values = rs.GetRows
			CountList1Values = rs.RecordCount-1
		end if
		CloseOBJ rs
			
		for i=0 to CountList1Values
			Set rs = Conn.Execute("select nombre_servicio from servicios where id_servicio=" & CheckNum(aList1Values(5,i)))
			if Not rs.EOF then
				aList1Values(6,i)=rs(0)
			end if
			CloseOBJ rs
		next
						
		'Obteniendo las condiciones del Routing Terrestre
		Set rs = Conn.Execute("select tipo_carga, factura, lista_empaque from routing_terrestre where id_routing=" & EXID)
		if Not rs.EOF then
			HandlingInformation = rs.GetRows
			Pos = rs.RecordCount-1
		end if
		CloseOBJ rs
	Case Else
		if isNumeric(FinalDes) then
			if FinalDes>0 then
				Set rsFilter = Conn.Execute("select pais from unlocode where unlocode_id=" & FinalDes)
				if Not rsFilter.EOF then
					FinalDes = rsFilter(0)
				end if
				CloseOBJ rsFilter			
			end if
		end if

		if IsNumeric(CountryOrigen) then
			if CountryOrigen>0 then
				Set rsFilter = Conn.Execute("select pais from unlocode where unlocode_id=" & CountryOrigen)
				if Not rsFilter.EOF then
					CountryOrigen = rsFilter(0)
				end if
				CloseOBJ rsFilter			
			end if
		end if
	End Select

	if ClientID>0 then
		Set rsFilter = Conn.Execute("select a.id_direccion, b.nombre_cliente, b.es_coloader, id_grupo from direcciones a, clientes b where a.id_cliente=b.id_cliente and b.id_cliente=" & ClientID)
		if Not rsFilter.EOF then
			AddressID = rsFilter(0)
			Name = rsFilter(1)
            ClientColoader = CheckNum(rsFilter(2))
           
		end if
		CloseOBJ rsFilter
	end if


    if CheckNum(NotifyPartyID)>0 then
		Set rsFilter = Conn.Execute("select a.id_direccion, b.nombre_cliente, b.es_coloader from direcciones a, clientes b where a.id_cliente=b.id_cliente and b.id_cliente=" & NotifyPartyID)
		if Not rsFilter.EOF then
			NotifyPartyAddrID = rsFilter(0)
			NotifyParty = rsFilter(1)
            NotifyPartyColoader = CheckNum(rsFilter(2))
		end if
		CloseOBJ rsFilter
	end if

    if CommodityCode>0 then
        Set rsFilter = Conn.Execute("select namees from commodities where commodityid=" & CommodityCode)
		if Not rsFilter.EOF then
			Commodity = rsFilter(0)
		end if
		CloseOBJ rsFilter		
	end if
	    
    if ShipperID<>0 then
        Set rsFilter = Conn.Execute("select b.id_direccion, a.nombre_cliente, a.es_coloader from clientes a left join direcciones b on b.id_cliente=a.id_cliente where a.id_cliente = " & ShipperID)
		if Not rsFilter.EOF then
			ShipperAddrID = rsFilter(0)
			ShipperData = rsFilter(1)
            ShipperColoader = CheckNum(rsFilter(2))
		end if
		CloseOBJ rsFilter
	end if

    if CheckNum(ColoaderID)>0 then
		Set rsFilter = Conn.Execute("select a.id_direccion, b.nombre_cliente, b.es_coloader from direcciones a, clientes b where a.id_cliente=b.id_cliente and b.id_cliente=" & ColoaderID)
		if Not rsFilter.EOF then
			ColoaderAddrID = rsFilter(0)
			ColoaderData = rsFilter(1)
		end if
		CloseOBJ rsFilter
	end if

    if ClientCollectID>0 then
        Set rsFilter = Conn.Execute("select b.nombre_cliente from direcciones a, clientes b where a.id_cliente=b.id_cliente and b.id_cliente=" & ClientCollectID)
		if Not rsFilter.EOF then
			ClientsCollect = rsFilter(0)
		end if
		CloseOBJ rsFilter
    end if

	if AgentID>0 then
		Set rsFilter = Conn.Execute("select agente, es_neutral from agentes where agente_id=" & AgentID)
		if Not rsFilter.EOF then
			AgentAddrID = 0
			AgentData = rsFilter(0)
            AgentNeutral = CheckNum(rsFilter(1))
		end if
		CloseOBJ rsFilter
	end if
		
	if isnumeric(ContactSignature) then
		if ContactSignature>0 then
			Set rsFilter = Conn.Execute("select pw_gecos from usuarios_empresas where id_usuario=" & ContactSignature)
			if Not rsFilter.EOF then
				ContactSignature = UCase(rsFilter(0))
			end if
			CloseOBJ rsFilter
		end if
	end if
	'Tipo de Paquete (Cajas, Cartones, etc.)
	if IsNumeric(Contener) then
        if Contener>0 then
			Set rsFilter = Conn.Execute("select tipo from tipo_paquete where tipo_id=" & Contener)
            if Not rsFilter.EOF then
				Contener = rsFilter(0)
			end if
			CloseOBJ rsFilter
        end if
	end if
    'Datos de a quien se debe notificar
	if RoutingID>0 and Notify="" then
		Set rsFilter = Conn.Execute("select notificar_a from routings where id_routing=" & RoutingID)
		if Not rsFilter.EOF then
			Notify = rsFilter(0)
		end if
		CloseOBJ rsFilter
	end if
    'Incoterms
	if IncotermsID>0 then
		Set rsFilter = Conn.Execute("select descripcion from incoterms where id_incoterms=" & IncotermsID)
		if Not rsFilter.EOF then
			Incoterms = rsFilter(0)
		end if
		CloseOBJ rsFilter
	end if
	CloseOBJ Conn
end if
Set aTableValues = Nothing

'response.write(CountList1Values & " - CountList1Values <br>")

'Asignando los rubros para guardarlos
if CountList1Values >= 0 then
	Select Case ExType
	Case 0,1,2
		Val = "" 'separador
		OpenConn2 Conn
		for i=0 to CountList1Values
			'Obteniendo los rubros del Routing

			'id_moneda, a.id_rubro, a.valor_prepaid, a.valor_collect, a.local, '', valor_sobreventa
			'response.write "select simbolo from monedas where moneda_id=" & CheckNum(aList1Values(0,i)) & "<br>"
			Set rs = Conn.Execute("select simbolo from monedas where moneda_id=" & CheckNum(aList1Values(0,i)))
			if Not rs.EOF then
				aList1Values(0,i)=rs(0)
			end if
			CloseOBJ rs
			
			'response.write "select desc_rubro_es from rubros where id_rubro=" & CheckNum(aList1Values(1,i)) & "<br>"
			Set rs = Conn.Execute("select concat(desc_rubro_es,' (',id_rubro,')') from rubros where id_rubro=" & CheckNum(aList1Values(1,i)))
			if Not rs.EOF then
				aList1Values(5,i)=rs(0)
			end if
			CloseOBJ rs
			
			'response.write "select nombre_servicio from servicios where id_servicio=" & CheckNum(aList1Values(7,i)) & "<br>"
			Set rs = Conn.Execute("select nombre_servicio from servicios where id_servicio=" & CheckNum(aList1Values(7,i)))
			if Not rs.EOF then
				aList1Values(8,i)=rs(0)
			end if
			CloseOBJ rs

			aList2Values = aList2Values & Val & aList1Values(0,i) 'simbolo por id_moneda
			aList3Values = aList3Values & Val & CInt(aList1Values(1,i)) 'id_rubro
			aList5Values = aList5Values & Val & aList1Values(4,i) 'local
			aList6Values = aList6Values & Val & aList1Values(5,i) 'desc_rubro_es por id_rubro
			aList7Values = aList7Values & Val & aList1Values(6,i) 'valor_sobreventa
			aList9Values = aList9Values & Val & aList1Values(7,i) 'id_servicio
			aList10Values = aList10Values & Val & aList1Values(8,i) 'nombre_servicio
			aList11Values = aList11Values & Val & "0" 'factura ID
            aList12Values = aList12Values & Val & "0" 'Si se debe calcular en el BL, el usuario puede cambiarlo luego en "Cobros y Documentos"
            aList13Values = aList13Values & Val & aList1Values(9,i) 'ID del Intercompany

			if CheckNum(aList1Values(2,i)) <> 0 then 'Guarda el valor Prepaid
				aList4Values = aList4Values & Val & aList1Values(2,i) 'valor_prepaid
				aList8Values = aList8Values & Val & "0"
				
				SetCharges 0, aList1Values(0,i), CInt(aList1Values(1,i)), CheckNum(aList1Values(2,i))+CheckNum(aList1Values(6,i))
			else 'Guarda el valor Collect
				aList4Values = aList4Values & Val & aList1Values(3,i) 'valor_collect
				aList8Values = aList8Values & Val & "1"

				SetCharges 1, aList1Values(0,i), CInt(aList1Values(1,i)), CheckNum(aList1Values(3,i))+CheckNum(aList1Values(6,i))
			end if
			
			Val = "|"
		next	

        'response.write(aList4Values & " - aList4Values")
		CloseOBJ Conn
		Set aList1Values = Nothing
	
	Case 4,5,6,7
		'c.simbolo, a.id_rubro, a.valor, a.local, b.desc_rubro_es, a.id_servicio, '', a.intercompany
		Val = "" 'separador
		for i=0 to CountList1Values
			aList2Values = aList2Values & Val & aList1Values(0,i) 'simbolo
			aList3Values = aList3Values & Val & CInt(aList1Values(1,i)) 'id_rubro
			aList4Values = aList4Values & Val & aList1Values(2,i) 'valor
			aList5Values = aList5Values & Val & aList1Values(3,i) 'local
			aList6Values = aList6Values & Val & aList1Values(4,i) 'desc_rubro_es
			aList7Values = aList7Values & Val & "0" 'sobreventa: en routing no hay
			aList8Values = aList8Values & Val & aList1Values(8,i)
			aList9Values = aList9Values & Val & aList1Values(5,i) 'id_servicio
			aList10Values = aList10Values & Val & aList1Values(6,i) 'nombre_servicio
			aList11Values = aList11Values & Val & "0" 'factura ID
			if aList1Values(7,i) <> 0  then
                aList12Values = aList12Values & Val & "0" 'Intercompany no se debe calcular en el BL, el usuario puede cambiarlo luego en "Cobros y Documentos"
            else
                aList12Values = aList12Values & Val & "1" 'Si se debe calcular en el BL, el usuario puede cambiarlo luego en "Cobros y Documentos"
            end if
            aList13Values = aList13Values & Val & aList1Values(7,i) 'ID del Intercompany
            
			Val = "|"

            'response.write "(" & aList1Values(8,i) & ")(" & aList1Values(0,i) & ")(" & aList1Values(1,i) & ")(" & CheckNum(aList1Values(2,i)) & ")<br>" 

			SetCharges aList1Values(8,i), aList1Values(0,i), CInt(aList1Values(1,i)), CheckNum(aList1Values(2,i))'No suma Sobreventa porque en RO no hay
		next

            'response.write "(" & Freight2 & ")(" & Insurance2 & ")(" & AnotherChargesPrepaid & ")<br>" 

		Set aList1Values = Nothing
	End Select
	
	'Al ser asignado por primera vez, no se sabe si el tipo de trafico que tendra sera 
	'Transito -> 0=consolidado, 1=express, ó
	'Local -> 3=recoleccion, 4=entrega
	'Solo se asigna a su tipo de Itinerario -1=Transito,-2=Local
end if

if BLType < 0 then
	Select Case ExType
	case 0,1,2,4,5,8,99
			BLType = -1 'Transito
	case 6,7,9,10,11,12,13,14,15
			BLType = -2 'Local
	End Select
End if

'En el Caso de CIF, se guarda el usuario del sistema terrestre que esta ingresando los datos
Select Case ExType
Case 8,9,10,11,12,13,14,15,99
    ContactSignature = UCase(Session("OperatorName"))
    OpenConn2 Conn
    	'Obteniendo listado de tipos de paquete
        Set rsFilter = Conn.Execute("select tipo from tipo_paquete order by tipo")
		if Not rsFilter.EOF then
            aTableValues = rsFilter.GetRows
        	CountList2Values = rsFilter.RecordCount -1
		end if
        CloseOBJ rsFilter

        'Obteniendo el listado de incoterms
        Set rsFilter = Conn.Execute("select id_incoterms, descripcion from incoterms order by descripcion")
		if Not rsFilter.EOF then
            aTableValues2 = rsFilter.GetRows
        	CountTableValues2 = rsFilter.RecordCount -1
		end if
    CloseOBJs rsFilter, Conn

    if CountryOrigen="" then
        CountryOrigen=Session("OperatorCountry")
    end if
End Select

'Verificando si la carta es de cliente especial para que le cobren Custodio y GPS
if Action<>0 then
    OpenConn2 Conn
        Set rs = Conn.Execute("select id_grupo from clientes where id_cliente in (" & ClientID & "," & ShipperID & "," & NotifyPartyID & ")")
	    do while not rs.EOF
            if FRegExp(PtrnSpecialClient, CheckNum(rs(0)),  "", 2) then
                AlertSpecialClient = "AVISO: para esta carga consulte los cargos a facturar para colocar GPS y Custodio"
            end if
            rs.MoveNext
	    loop
    CloseOBJs rs, Conn
end if



'response.write("fleet : " & ExType & "<br>")

if ExType <> 99 then

    fleet = GroupData(ClientID) '2020-09-11

end if

'response.write("fleet : " & fleet & "<br>")

'response.write("SEGURO - " & Seguro & "<br>")
'response.write("Poliza_Seguro - " & Poliza_Seguro & "<br>")
'response.write("ObjectID - " & ObjectID & "<br>")
'response.write("Routing_Seg - " & Routing_Seg & "<br>")
'response.write("RSeg_Borrado - " & RSeg_Borrado & "<br>")
'response.write("RSeg_Poliza - " & RSeg_Poliza & "<br>")
'response.write("RSeg_BLID - " & RSeg_BLID & "<br>")
'response.write("Routing_Adu - " & Routing_Adu & "<br>")
'response.write("RAdu_Activo - " & RAdu_Activo & "<br>")
'response.write("RAdu_Borrado - " & RAdu_Borrado & "<br>")
'response.write(RefHBLNumber & " - " & RefBLID)



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

.button {
    display: block;
    width: 76%;
    height: 25px;
    background: #4E9CAF;
    padding: 10px;
    text-align: center;
    border-radius: 5px;
    color: white;
    font-weight: bold;
    line-height: 25px;
    text-decoration:none;
    font-family:Arial;
}

</style>
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
    <%if BAWResult <> "" then %>
        alert(BAWResult);
    <%end if %>
    var Incoterms = new Array();
    <%
    for i=0 to CountTableValues2
        response.write "Incoterms[" & aTableValues2(0,i) & "] = '" & aTableValues2(1,i) & "';" & vbCrLf
    next

    dim iMensaje 

    iMensaje = ""

    if (((Seguro = 1) and (Poliza_Seguro = "")) and (ObjectID = 0)) then 
        iMensaje = "El Routing no tiene número de póliza. No se puede cargar el RO al Sistema Terrestre."
    elseif (((Routing_Seg <> 0) and (RSeg_Borrado = 1)) and (ObjectID = 0)) then 
        iMensaje = "El Routing de seguro asociado a éste routing fue borrado. No se puede cargar el RO al Sistema Terrestre."
    elseif (((Routing_Seg <> 0) and (RSeg_Borrado = 0) and (RSeg_Poliza = "")) and (ObjectID = 0)) then 
        iMensaje = "El Routing de seguro asociado no tiene número de póliza. No se puede cargar el RO al Sistema Terrestre."
    elseif (((Routing_Seg <> 0) and (RSeg_Borrado = 0) and (RSeg_Poliza <> "") and (RSeg_BLID <> 0)) and (ObjectID = 0)) then 
        iMensaje = "El Routing de seguro asociado ya está siendo utilizado. No se puede cargar el RO al Sistema Terrestre."
    elseif (((Routing_Adu <> 0) and (RAdu_Activo = 0)) and (ObjectID = 0)) then 
        iMensaje = "El Routing de aduana asociado ya está siendo utilizado. No se puede cargar el RO al Sistema Terrestre."
    elseif (((Routing_Adu <> 0) and (RAdu_Borrado = 1)) and (ObjectID = 0)) then 
        iMensaje = "El Routing de aduana asociado a éste routing fue borrado. No se puede cargar el RO al Sistema Terrestre."
    end if



    %>

    function submit_(Action){

        move();
        document.forma.Action.value = Action;
        document.forma.submit();	
        return false;
    }

    <%if iMensaje <> "" then %>
        alert('<%=iMensaje%>');
    <%else %>
	    function validar(Action) {

            if (Action == 9) { //valida codigo referencia
                <% if fleet <> "NA" AND fleet <> "" then%>
                ///////////////////////////////////////////////////////////2020-09-08
                if (!valTxt(document.forma.CodeReference, 1, 10)){return (false)};
                //if (document.forma.TripRequestDate.value == ""){ alert("Seleccione Fecha cd Requerimiento"); return (false)}; 
                ////////////////////////////////////////////////////////////
                <% end if%>


                return submit_(Action);
                //move();
                //document.forma.Action.value = Action;
                //document.forma.submit();	
                //return false;
            }

            if (Action == 8) { //limpiar codigo referencia
                return submit_(Action);
                //document.forma.Action.value = Action;
                //document.forma.submit();	
                //return false;
            }

            if (Action != 3) {
                if (!valTxt(document.forma.Clients, 3, 8)){return (false)};     //3, 5      Caso No. 591 / IM-1464-1-591
			    if (document.forma.AddressesID.value==0){
				    alert("El Cliente no tienen direccion asignada en el Catalogo de Clientes");
				    return (false)
			    };
			    if (!valTxt(document.forma.Shippers, 3, 8)){return (false)};	//3, 5      Caso No. 591 / IM-1464-1-591		
			
			    if (document.forma.ShippersAddrID.value==0){
				    alert("El exportador/coloader no tienen direccion asignada en el Catalogo de Clientes");
				    return (false)
			    };


                <%if fleet <> "NA" AND fleet <> "" then%>
                ///////////////////////////////////////////////////////////2020-09-08
                if (!valTxt(document.forma.CodeReference, 1, 10)){return (false)};
                if (document.forma.TripRequestDate.value == ""){ alert("Seleccione Fecha de Requerimiento"); return (false)}; 
                ////////////////////////////////////////////////////////////
                <%end if%>

                

                if (!valSelec(document.forma.CountrySession)){return (false)}; //2020-10-13
			    
            extype = <%=ExType%>;
  		    <%Select Case ExType 'Carga en Transito
		    Case 0,1,2,4,5,8,99%>
                <%if ExType <> 8 then%>
			    if (document.forma.WareHouseDischargeDate.value == "")  {
				    alert("No se puede Agregar al Itinerario si no tiene Fecha de Descarga indicada desde Trafico");
				    return (false);
			    }
			    if (document.forma.CountryOrigen.value == "")  {
				    alert("No se puede Agregar al Itinerario si no indica el Pais de Origen");
				    return (false);
			    }
                <%else %>
                if (!valTxt(document.forma.WareHouseDischargeDate, 3, 5)){return (false)};
			    if (!valTxt(document.forma.Agents, 3, 5)){return (false)};
			
                    //Validacion de Latin Freight y Aimar, el resto de empresas no tiene esta validacion, por ejemplo N1 (GRH)
                    <%if FilterAimarLatin = 1 then%>
                    if ((document.forma.Countries.value!="N1") && (document.forma.Countries.value!="A2")) {
                        if (document.forma.Countries.value.substr(2,3)=="LTF") {
                            if (document.forma.AgentNeutral.value == 0) {
                                alert("Para operaciones de Latin Freight solo puede utilizar agentes Neutrales");
				                document.forma.Agents.focus();
                                return (false);
			                }
                        } else {
                            var EconoCodes = /<%=PtrnEconoCodes%>/;
                            var Result = EconoCodes.exec(document.forma.AgentsID.value)
                            if (Result == null) {
                                if (document.forma.ClientColoader.value == 1) {
                                    alert("Para operaciones de Aimar, solo cuando el Agente es Econocaribe puede asignar Clientes o Shippers Coloaders, favor de consultar con su supervisor y revisar el administrador de catalogos");
				                    document.forma.Clients.focus();
                                    return (false);
			                    }
                                if (document.forma.ShipperColoader.value == 1) {
                                    alert("Para operaciones de Aimar, solo cuando el Agente es Econocaribe puede asignar Clientes o Shippers Coloaders, favor de consultar con su supervisor y revisar el administrador de catalogos");
				                    document.forma.Shippers.focus();
                                    return (false);
			                    }
                                if (document.forma.ColoadersID.value != 0) {
                                    alert("Para operaciones de Aimar, solo cuando el Agente es Econocaribe puede utilizar Coloaders, favor de consultar con su supervisor y revisar el administrador de catalogos");
				                    document.forma.Coloaders.focus();
                                    return (false);
			                    }
			                }
                        }
                    }
                    <%end if %>
            
                if (!valSelec(document.forma.Contener, 3, 5)){return (false)};
			    if (!valTxt(document.forma.Weights, 1, 10)){return (false)};
			    if (!valTxt(document.forma.Volumes, 1, 10)){return (false)};
			    if (!valTxt(document.forma.NoOfPieces, 1, 10)){return (false)};
                if (!valSelec(document.forma.IncotermsID, 3, 5)){return (false)};
			    if (!valTxt(document.forma.Container, 3, 5)){return (false)};
			    if (!valTxt(document.forma.MBLs, 3, 5)){return (false)};
			    if (!valTxt(document.forma.BLs, 3, 5)){return (false)};
                document.forma.Incoterms.value = Incoterms[document.forma.IncotermsID.value];

			    <%end if %>

                if (!valTxt(document.forma.DiceContener, 3, 5)){return (false)};


                

			    if (!valSelec(document.forma.CountryOrigen)){return (false)};
			    if (!valSelec(document.forma.CountriesFinalDes)){return (false)};

                <%'Validacion para Cliente BEARCOM
                select Case ClientID
                Case 54287, 54058, 46815, 28208, 1380%>
			        if (!valTxt(document.forma.PO, 3, 5)){return (false)};
                <%end Select %>

                <%'Validacion para Shipper BEARCOM
                select Case ShipperID
                Case 54287, 54058, 46815, 28208, 1380%>
			        if (!valTxt(document.forma.PO, 3, 5)){return (false)};
                <%end Select %>

			    if (!valSelec(document.forma.ChargeType)){return (false)};
                /*
			    if (!valSelec(document.forma.Endorse)){return (false)};
			    if (!valSelec(document.forma.EndorseType)){return (false)};
			    if (!valSelec(document.forma.Declaration)){return (false)};
			    if (!valSelec(document.forma.DeclarationType)){return (false)};
			    if (!valSelec(document.forma.RequestNo)){return (false)};
			    if (!valSelec(document.forma.RequestType)){return (false)};
			    if (!valSelec(document.forma.BLsType)){return (false)};
			    if (!valSelec(document.forma.BillType)){return (false)};
                */
			    if (document.forma.BillType.value!=2){
				    if (!valTxt(document.forma.Bill, 3, 5)){return (false)};
			    };
			    if (!valSelec(document.forma.PackingListType)){return (false)};
			    if (document.forma.PackingListType.value!=2){
				    if (!valTxt(document.forma.PackingList, 3, 5)){return (false)};
			    };
			    if ((document.forma.ShippersID.value==8621) || (document.forma.ShippersID.value==6841) || (document.forma.ShippersID.value==4796)){	
				    if (!valTxt(document.forma.Notify, 3, 5)){return (false)};
			    };
		    <%Case 6,7,9,10,11,12,13,14,15
                Select Case ExType
                Case 11,12,13%>
                if (document.forma.WareHouseDischargeDate.value == "")  {
				    alert("No se puede Agregar al Itinerario si no tiene Fecha de Descarga indicada desde Trafico");
				    return (false);
			    }
                <%End Select%>
                <%Select Case ExType
                Case 9,10,11,12,13,14,15%>
                if (!valSelec(document.forma.ChargeType)){return (false)};
                    if (!valSelec(document.forma.IncotermsID, 3, 5)){return (false)};
                    document.forma.Incoterms.value = Incoterms[document.forma.IncotermsID.value];
			    <%End Select%>
                <%if ExType = 15 then%>
			    if (!valTxt(document.forma.WareHouseDischargeDate, 1, 10)){return (false)};
                if (!valSelec(document.forma.Contener, 3, 5)){return (false)};
			    if (!valTxt(document.forma.Weights, 1, 10)){return (false)};
			    if (!valTxt(document.forma.Volumes, 1, 10)){return (false)};
			    if (!valTxt(document.forma.NoOfPieces, 1, 10)){return (false)};
                if (!valTxt(document.forma.Container, 3, 5)){return (false)};
			    if (!valTxt(document.forma.MBLs, 3, 5)){return (false)};
                if (!valTxt(document.forma.BLs, 3, 5)){return (false)};
                <%end if%>

			
                if (!valSelec(document.forma.Yr)){return (false)};
			    if (!valSelec(document.forma.Mh)){return (false)};
			    if (!valSelec(document.forma.Dy)){return (false)};
			    if (!valSelec(document.forma.Hr)){return (false)};
			    if (!valSelec(document.forma.Mn)){return (false)};
			    if (!valTxt(document.forma.SenderData, 3, 5)){return (false)};
                if (!valTxt(document.forma.ConsignerData, 3, 5)){return (false)};
                if (!valTxt(document.forma.AgentSignature, 3, 5)){return (false)};
                if (!valTxt(document.forma.Phone1, 3, 5)){return (false)};

                document.forma.BLArrivalDate.value = document.forma.Yr.value + "/" + document.forma.Mh.value + "/" + document.forma.Dy.value + " " + document.forma.Hr.value + ":" + document.forma.Mn.value + ":00";
			    <%if InStr(1,Session("Countries"),"GT")>0 then%>
			    if (!valTxt(document.forma.DiceContener, 3, 5)){return (false)};
			    //if (!valSelec(document.forma.CountriesFinalDes)){return (false)};
			    <%end if%>
		    <%End Select%>
		    }
        
            if (document.forma.Notify.value.length <= 3) {
                document.forma.Notify.value = "---";
            }
            //move();
            //document.forma.Action.value = Action;
            //document.forma.submit();	
            submit_(Action);		 
	     }
    <%end if %>
	 
	function GetData(GID,DiceContenerTemp){
		window.open('Search_BLData.asp?GID='+GID+'&DiceContenerTemp='+DiceContenerTemp,'BLData','height=400,width=460,menubar=0,resizable=1,scrollbars=1,toolbar=0,status=1');
	}
	function IData(GID){
		window.open('InsertData.asp?GID='+GID+'&SO=1','IData','height=400,width=460,menubar=0,resizable=1,scrollbars=1,toolbar=0');
	}

    function SetLabelID(Label) {
		var LabelID = "";
		if (Label == "WareHouseDischargeDate") {
			LabelID = "Fecha de Descarga";
		} 		
		if (Label == "TripRequestDate") {
			LabelID = "Fecha de Requerimiento";
		} 	
        

		return LabelID;
	}

    function abrir(Label){
		var DateSend, Subject;
		if (parseInt(navigator.appVersion) < 5) {
			DateSend = document.forma(Label).value;
		} else {

            var labelid = SetLabelID(Label)
			DateSend = document.getElementById(labelid).value;
		}
		Subject = '';	
		window.open('Agenda.asp?Action=1&Label=' + Label + '&DateSend=' + DateSend + '&Subj=' + Subject,'Seleccionar','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,width=170,height=160,top=250,left=250');

        return false;
	}

    /*
	function abrir(Label){

        try {      
        
        var DateSend, Subject;
            console.log(navigator.appVersion); 	
                     
			if (parseInt(navigator.appVersion) < 5) {
				DateSend = document.forma(Label).value;
			} else {
				var LabelID = SetLabelID(Label);

                alert(LabelID);

				DateSend = document.getElementById(LabelID).value;
			}
        }
        catch(err) {
            console.log(err); 								
        }

        Subject = '';	
        window.open('Agenda.asp?Action=1&Label=' + Label + '&DateSend=' + DateSend + '&Subj=' + Subject,'Seleccionar','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,width=170,height=160,top=150,left=350');
	}
    */

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

    function Desglose(OID, CD, CT){
	window.open('InsertData.asp?GID=36&OID='+OID+'&CD='+CD+'&CT='+CT,'EDataD','height=650,width=600,menubar=0,resizable=1,scrollbars=1,toolbar=0');
    }

</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="resizeTo(600,680);self.focus();">
	<div id="myProgress">
      <div id="myBar">10%</div>
    </div>
    <FORM name="forma" action="InsertData.asp" method="post">
    <TABLE cellspacing=0 cellpadding=2 width=400 align=center>
	<INPUT name="Action" type=hidden value=0>
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
	<INPUT name="OID" type=hidden value="<%=ObjectID%>">
	<INPUT name="CD" type=hidden value="<%=CreatedDate%>">
	<INPUT name="CT" type=hidden value="<%=CreatedTime%>">
	<INPUT name="Expired" type=hidden value="0">
	<INPUT name="EID" type=hidden value="<%=EXID%>">
	<INPUT name="ET" type=hidden value="<%=ExType%>">
	<INPUT name="BLType" type=hidden value="<%=BLType%>">
	<INPUT name="EC" type=hidden value="<%=CountryDes%>">
	<INPUT name="ReservationDate" type=hidden value="<%=ReservationDate%>">	
    <INPUT name="Contact" type=hidden value="<%=ContactSignature%>">
    <INPUT name="CodeReferenceValid" type=hidden value="<%=CodeReferenceValid%>">
    <INPUT name="fleet" type=hidden value="<%=fleet%>">
    <INPUT name="cambio" type=hidden value="">
	
	<%select Case ExType
    case 0,1,2%>
    <INPUT name="WareHouseDischargeDate" type=hidden value="<%=WareHouseDischargeDate%>">
    <INPUT name="Contener" type=hidden value="<%=Contener%>">
	<INPUT name="Weights" type=hidden value="<%=Weight%>">
	<INPUT name="Volumes" type=hidden value="<%=Volume%>">
	<INPUT name="NoOfPieces" type=hidden value="<%=TotNoOfPieces%>">
	<INPUT name="Agents" type=hidden value="<%=AgentData%>">
	<INPUT name="BLs" type=hidden value="<%=BL%>">
	<INPUT name="MBLs" type=hidden value="<%=MBL%>">
	<INPUT name="Container" type=hidden value="<%=Container%>">
    <%case 4,5,6,7,9,10,11,12,13,14 %>
    <INPUT name="WareHouseDischargeDate" type=hidden value="<%=WareHouseDischargeDate%>">
    <INPUT name="Contener" type=hidden value="<%=Contener%>">
	

        <%if ExType <> 4 then %>

        <INPUT name="Weights" type=hidden value="<%=Weight%>">
	    <INPUT name="Volumes" type=hidden value="<%=Volume%>">

        <%end if %>


	<INPUT name="NoOfPieces" type=hidden value="<%=TotNoOfPieces%>">
	<INPUT name="Agents" type=hidden value="<%=AgentData%>">
	<INPUT name="BLs" type=hidden value="<%=BL%>">
    <INPUT name="RefHBLNumber" type=hidden value="<%=RefHBLNumber%>">
    <INPUT name="RefBLID" type=hidden value="<%=RefBLID%>">
    <INPUT name="RosClientID" type=hidden value="<%=RosClientID%>">


    <%case 99 %>
    <INPUT name="TransactionSetPurposeCode" type=hidden value="<%=TransactionSetPurposeCode%>">
    <INPUT name="ShipmentIDNumber" type=hidden value="<%=ShipmentIDNumber%>">
    <INPUT name="tch_BLDetailID" type=hidden value="<%=tch_BLDetailID%>">
    <INPUT name="gS_MessageID" type=hidden value="<%=gS_MessageID%>">          
	<%end select %>

    <INPUT name="Countries" type=hidden value="<%=Countries%>">	
	<INPUT name="ItemCurrs" type=hidden value="<%=aList2Values%>">
	<INPUT name="ItemIDs" type=hidden value="<%=aList3Values%>">
	<INPUT name="ItemVals" type=hidden value="<%=aList4Values%>">
	<INPUT name="ItemLocs" type=hidden value="<%=aList5Values%>">
	<INPUT name="ItemNames" type=hidden value="<%=aList6Values%>">
	<INPUT name="ItemOVals" type=hidden value="<%=aList7Values%>">
	<INPUT name="ItemPPCCs" type=hidden value="<%=aList8Values%>">
	<INPUT name="ItemServIDs" type=hidden value="<%=aList9Values%>">
	<INPUT name="ItemServNames" type=hidden value="<%=aList10Values%>">
    <INPUT name="ItemInvoices" type=hidden value="<%=aList11Values%>">
    <INPUT name="ItemCalcInBls" type=hidden value="<%=aList12Values%>">
    <INPUT name="ItemIntercompanyIDs" type=hidden value="<%=aList13Values%>">
    <INPUT name="CantItems" type=hidden value="<%=CountList1Values%>">
	<INPUT name="Freight" type=hidden value="<%=Freight%>">
	<INPUT name="Freight2" type=hidden value="<%=Freight2%>">
	<INPUT name="Insurance" type=hidden value="<%=Insurance%>">
	<INPUT name="Insurance2" type=hidden value="<%=Insurance2%>">
	<INPUT name="AnotherChargesCollect" type=hidden value="<%=AnotherChargesCollect%>">
	<INPUT name="AnotherChargesPrepaid" type=hidden value="<%=AnotherChargesPrepaid%>">
	<INPUT name="PolicyNo" type=hidden value="<%=PolicyNo%>">
		<%'if BLNumber<>"" then %>
        <!--<TR><TD class=label align=center colspan=2><b><font color=red>Nota: No puede Modificar Datos porque ya tiene asignada la CP Hija<br><%=BLNumber %></font></b></TD></TR> -->
		<%'end if %>



        <TR style="display:none"><TD class=label align=right><b>File :</b></TD><TD class=label align=left>
            <input class="label" name="file" id="file" readonly value="<%=file%>" />
            </TD></TR> 

        <%
        
        'response.write "(" & ExType & ")(" & TransactionSetPurposeCode & ")(" & CheckNum(ObjectID) & ")(" & EXID & ")(" & gS_MessageID & ")(" & CheckNum(Request("EID")) & ")<br>"

        
        If ExType = 99 and (TransactionSetPurposeCode = "02" or CheckNum(ObjectID) = 0) then%>

            <tr><td colspan=2 align=center>
            
                <%If TransactionSetPurposeCode = "02" then%>

                    <a href="Utils.asp?tch_pk=<%=CheckNum(Request("EID"))%>&tch_aceptar=ACEPTAR&tch_BLDetailID=<%=tch_BLDetailID%>&SetPurposeCode=02&gS_MessageID=<%=gS_MessageID%>&Result=1" onclick="return confirm('Esta seguro de Aceptar la cancelacion de Registro Colgate?')" class="button listCancelar">CANCELACION <%=EXID_Ant & " - " & BL%></a>
                    <br />
                    <a href="Utils.asp?tch_pk=<%=EXID%>&tch_rechazar=RECHAZAR&gS_MessageID=<%=gS_MessageID%>&Result=2" onclick="return confirm('Esta seguro de Rechazar el registro Colgate?')" class="button <%=Iif(TransactionSetPurposeCode = "00","listOriginal","listRemplazo")%>">RECHAZAR Message ID <%=gS_MessageID%></a>

                <%Else%>

                    <a href="Utils.asp?tch_pk=<%=EXID%>&tch_rechazar=RECHAZAR&gS_MessageID=<%=gS_MessageID%>&Result=2" onclick="return confirm('Esta seguro de Rechazar el registro Colgate?')" class="button <%=Iif(TransactionSetPurposeCode = "00","listOriginal","listRemplazo")%>">RECHAZAR Message ID <%=gS_MessageID%></a>

                <%End If%>
  
            </tr></td>

        <%End If%>

        <%If ColgateData<>"" then%>
                <font color=blue><b><%=ColgateData%></b></font><br>
        <%End If%>



        <TR><TD class=label align=right><b>C&oacute;digo:</b></TD><TD class=label align=left><%if ObjectID <> 0 then response.write ObjectID End if%></TD></TR> 
		
		<%If ReservationDate <> "" then%>
		<TR><TD class=label align=right><b>Fecha Solicitud:</b></TD><TD class=label align=left><%=WareHouseDischargeDate & "&nbsp;" & ReservationDate%></TD></TR> 
		<%Else%>
            <TR><TD class=label align=right><b>Fecha Descarga:</b></TD><TD class=label align=left>
            <%if ExType<>8 and ExType<>15 and ExType<>99 Then%>
    		<%=WareHouseDischargeDate%>
	    	<%Else%>
		    <INPUT type="text" readonly name="WareHouseDischargeDate" value="<%=WareHouseDischargeDate%>" size=23 maxLength=19 class=label id="Fecha de Descarga">&nbsp;<a href="#" onClick="JavaScript:abrir('WareHouseDischargeDate');return (false);" class="menu"><font color="FFFFFF"><b>Seleccionar</b></font></a>
            <%End If%>
            </TD></TR> 
		<%End If%>
		
		<TR><TD class=label align=right valign=top><b>Cliente:</b></TD><TD class=label align=left>
		    <%if Name="" and ClientsTemp<>"" then%>
                <font color=blue><b><%=ClientsTemp%></b></font><br>
            <%end if%>
		    <%if ClientsAddTemp<>"" then%>
                <font color=blue><b><%=ClientsAddTemp%></b></font><br>
            <%end if%>
			
            <INPUT TYPE=text class=label name="Clients" id="Cliente" value="<%=Name%>" maxlength="200" size="35" readonly ><%="<b> ID: " & ClientID & "<b>"%>
			
            <%Select Case Extype 
            Case 8,15,99%>
            <a class="submenu" href="http://10.10.1.20/catalogo_admin/login.php" target="_blank"><font color="FFFFFF"><b>Nuevo</b></font></a>
			<a href="#" onClick="Javascript:GetData(11,'<%=ClientsTemp%>');return (false);" class="menu"><font color="FFFFFF"><b>Buscar</b></font></a>
            <%End Select %>
		<INPUT name="ClientsID" type=hidden value="<%=ClientID%>">
		<INPUT name="AddressesID" type=hidden value="<%=CheckNum(AddressID)%>">
        <INPUT name="ClientCollectID" type=hidden value="<%=ClientCollectID%>">
        <INPUT name="ClientsCollect" type=hidden value="<%=ClientsCollect%>">
		<INPUT name="ClientColoader" type=hidden value="<%=CheckNum(ClientColoader)%>">
        <input name="NotifyPartyID" type=hidden value="<%=NotifyPartyID%>"/>
        <input name="NotifyPartyAddrID" type=hidden value="<%=NotifyPartyAddrID%>"/>
        <input name="NotifyParty" type=hidden value="<%=NotifyParty%>"/>


        <INPUT name="Routing_Seg" type=hidden value="<%=Routing_Seg%>">
        <INPUT name="Seguro" type=hidden value="<%=Seguro%>">
        <INPUT name="Poliza_Seguro" type=hidden value="<%=Poliza_Seguro%>">
        <INPUT name="RSeg_Borrado" type=hidden value="<%=RSeg_Borrado%>">
        <INPUT name="RSeg_Poliza" type=hidden value="<%=RSeg_Poliza%>">
        <INPUT name="RSeg_BLID" type=hidden value="<%=RSeg_BLID%>">
        
        <INPUT name="Routing_Adu" type=hidden value="<%=Routing_Adu%>">
        <INPUT name="RAdu_Activo" type=hidden value="<%=RAdu_Activo%>">
        <INPUT name="RAdu_Borrado" type=hidden value="<%=RAdu_Borrado%>">
        <INPUT name="RAdu_NDUA" type=hidden value="<%=RAdu_NDUA%>">

		</TD></TR>









        <% if fleet <> "NA" AND fleet <> "" then%>
<!-------------------------------------------2020-09-08--->
        <TR><TD class=label align=center colspan=2 style="border:1px solid yellow"><i>Informacion requerida para "<%=fleet%>"</i></TD></TR> 

        <TR><TD class=label align=right><b>Tipo Referencia:</b></TD><TD class=label align=left>
            <select name="TypeReference" id="TypeReference" class="label">
					<option value="Shipment">Shipment</option>
			</select>
        </TD></TR> 

       <TR><TD class=label align=right><b>Codigo Referencia:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="CodeReference" id="Codigo Referencia" value="<%=CodeReference%>" maxlength="10" size="15" onKeyUp="res(this,numb);"  <%=IIf(CodeReferenceValid = 1," readonly style='background:#eee'","") %> >
       
            <% 'if fleet <> "NA" AND fleet <> "" AND CheckNum(CodeReference) = 0 then%>
            
            <% if CheckNum(CodeReference) = 0 then%>

            <INPUT name=enviar type=button onClick="JavaScript:validar(9);" value="&nbsp;&nbsp;Validar Codigo Referencia&nbsp;&nbsp;" class=label />
      
            <% else %>
            
            <INPUT name=enviar type=button onClick="JavaScript:validar(8);" value="&nbsp;&nbsp;Limpiar Codigo Referencia&nbsp;&nbsp;" class=label />
      
            <% end if %>

       </TD></TR> 
 

        <TR><TD class=label align=right><b>Tipo Viaje:</b></TD><TD class=label align=left>
            <select name="TripType" id="TripType" class="label">
					<option value="Replenishment">Replenishment</option>
                    <option value="Crossdock">Crossdock</option>
			</select>
        </TD></TR> 


        <TR><TD class=label align=right><b>Bodega:</b></TD><TD class=label align=left>
            <select name="ClientWarehouse" id="ClientWarehouse" class="label">
					<option value="GT10">GT10</option>
					<option value="GT21">GT21</option>
					<option value="GT10-GT21">GT10-GT21</option>
			</select>
        </TD></TR> 


        <TR><TD class=label align=right><b>Prioridad:</b></TD><TD class=label align=left>
            <select name="TripPriority" id="TripPriority" class="label">
					<option value="Alta">Alta</option>
					<option value="Super Alta">Super Alta</option>
			</select>
        </TD></TR> 

        <TR><TD class=label align=right><b>Fecha de Requerimiento:</b></TD>
  
            <TD class="label  label1" align=left colspan=3 nowrap>
		        <INPUT name="TripRequestDate" id="Fecha de Requerimiento" type=text value="<%=FormatDateTime(TripRequestDate)%>" size=12 maxLength=19 class=label readonly>		
                <INPUT type=image onClick="return abrir('TripRequestDate');" src="img/calendar.png">
		    </TD>
        </TR> 
           

        <TR><TD class=label align=right><b>Otros Documentos:</b></TD><TD class=label align=left>
            <textarea class="label" cols="55" rows="2" name="OtherDocs" id="OtherDocs"><%=OtherDocs%></textarea></TD></TR> 

        <TR><TD class=label align=right colspan=2 style="border:1px solid yellow"></TD></TR> 

<!---------------------------------------------->
        <% end if%>

		
		<%select Case ExType
		  case 0,1,2,4,5,6,7,8,9,10,11,12,13,14,15
			if InStr(1,Session("Countries"),"CR")>0 or InStr(1,Session("Countries"),"NI")>0 or InStr(1,Session("Countries"),"SV")>0 then%>
            <tr><TD class=label align=right>&nbsp;</TD><td class=label align=left colspan=2><select name='AIMAR' class=label>
			<option value="">INCLUIR AIMAR?</option>
			<option value=" / AIMAR NICARAGUA">AIMAR NICARAGUA</option>
			<option value=" / AIMAR LOGISTIC S.A. DE C.V.">AIMAR LOGISTIC S.A. DE C.V.</option>
			<option value=" / AIMAR GUATEMALA ALSERSA">AIMAR GUATEMALA ALSERSA</option>
			</select></td></tr>
		<%	end if
		  end select%>
		  
		<TR><TD class=label align=right valign=top><b>Exportador:</b></TD><TD class=label align=left>
		    <%if ShipperData="" and AgentsTemp<>"" then%>
                <font color=blue><b><%=AgentsTemp%></b></font><br>
            <%end if%>
		    
            <%if AgentsAddTemp<>"" then%>
                <font color=blue><b><%=AgentsAddTemp%></b></font><br>
            <%end if%>            
			
            <INPUT TYPE=text class=label name="Shippers" id="Exportador" value="<%=ShipperData%>" maxlength="200" size="35" readonly><%="<b> ID: " & ShipperID & "<b>"%>
			
            <%Select Case Extype 
            Case 8,15,99%>
                <a class="submenu" href="http://10.10.1.20/catalogo_admin/login.php" target="_blank"><font color="FFFFFF"><b>Nuevo</b></font></a>
			    <a href="#" onClick="Javascript:GetData(31,'<%=AgentsTemp%>');return (false);" class="menu"><font color="FFFFFF"><b>Buscar</b></font></a>
            <%End Select %>
		<INPUT name="ShippersID" type=hidden value="<%=ShipperID%>">
		<INPUT name="ShippersAddrID" type=hidden value="<%=CheckNum(ShipperAddrID)%>">
        <INPUT name="ShipperColoader" type=hidden value="<%=CheckNum(ShipperColoader)%>">
		</TD></TR>
        <TR><TD class=label align=right valign=top><b>Coloader:</b></TD><TD class=label align=left>
		    <INPUT TYPE=text class=label name="Coloaders" id="Text3" value="<%=ColoaderData%>" maxlength="200" size="35" readonly><%="<b> ID: " & ColoaderID & "<b>"%>
			<%Select Case Extype 
            Case 8,15,99%>
                <a class="submenu" href="http://10.10.1.20/catalogo_admin/login.php" target="_blank"><font color="FFFFFF"><b>Nuevo</b></font></a>
			    <a href="#" onClick="Javascript:GetData(34,'');return (false);" class="menu"><font color="FFFFFF"><b>Buscar</b></font></a>
            <%End Select%>
		<INPUT name="ColoadersID" type=hidden value="<%=ColoaderID%>">
		<INPUT name="ColoadersAddrID" type=hidden value="<%=CheckNum(ColoaderAddrID)%>">
        </TD></TR>
		<TR><TD class=label align=right><b>Agente:</b></TD><TD class=label align=left>
        <%if ExType<>8 and ExType<>15 and ExType<>99 then %>
        <%=AgentData%><%="<b> ID: " & AgentID & "<b>"%>
        <%else %>
        	<INPUT TYPE=text class=label name="Agents" id="Agente" value="<%=AgentData%>" maxlength="200" size="35" readonly>
            <%="<b> ID: " & AgentID & "<b>"%>
			<a class="submenu" href="http://10.10.1.20/catalogo_admin/login.php" target="_blank"><font color="FFFFFF"><b>Nuevo</b></font></a>
			<a href="#" onClick="Javascript:GetData(33,'');return (false);" class="menu"><font color="FFFFFF"><b>Buscar</b></font></a>
		<%end if %>
        <INPUT name="AgentsID" type=hidden value="<%=AgentID%>">
    	<INPUT name="AgentsAddrID" type=hidden value="<%=CheckNum(AgentAddrID)%>">
        <INPUT name="AgentNeutral" type=hidden value="<%=CheckNum(AgentNeutral)%>">
        </TD></TR> 

        <TR><TD class=label align=right valign=top><b>Producto:</b></TD>		
		<TD class=label align=left class="style4">
		    <%if Commodity="" and DiceContenerTemp<>"" then%>
                <font color=blue><b><%=DiceContenerTemp%></b></font><br>
            <%end if%>
        	<INPUT TYPE=text class=label name="DiceContener" id="Producto" value="<%=Commodity%>" maxlength="200" size="35" readonly><%="<b> ID: " & CommodityCode & "<b>"%>
            <a href="#" onClick="Javascript:IData(9);return (false);" class="submenu" target=_blank><font color="FFFFFF"><b>Nuevo</b></font></a>
			<a href="#" onClick="Javascript:GetData(9,'<%=DiceContenerTemp%>');return (false);" class="menu"><font color="FFFFFF"><b>Buscar</b></font></a>
		<INPUT name="CommoditiesID" type=hidden value="<%=CommodityCode%>">
		</TD></TR> 


		<%if ExType<>8 and ExType<>15 and ExType<>99 then %>

		<TR><TD class=label align=right><b>Tipo Paquete:</b></TD><TD class=label align=left><%=Contener%></TD></TR> 


    		<%if ExType = 4 then %>

		    <TR><TD class=label align=right><b>Peso:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Weights" id="Text5" value="<%=Weight%>" maxlength="200" size="15" onKeyUp="res(this,numb);"></TD></TR> 
		    <TR><TD class=label align=right><b>Volumen:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Volumes" id="Text6" value="<%=Volume%>" maxlength="200" size="15" onKeyUp="res(this,numb);"></TD></TR> 

            <%else %>

		    <TR><TD class=label align=right><b>Peso:</b></TD><TD class=label align=left><%=Weight%></TD></TR> 
		    <TR><TD class=label align=right><b>Volumen:</b></TD><TD class=label align=left><%=Volume%></TD></TR> 

            <%end if %>

		<TR><TD class=label align=right><b>Bultos:</b></TD><TD class=label align=left><%=TotNoOfPieces%></TD></TR> 
        
        <%else %>
        <TR><TD class=label align=right><b>Tipo Paquete:</b></TD><TD class=label align=left>
            <select name="Contener" id="Tipo Paquete" class="label">
					<option value="-1">Seleccionar</option>
                    <%for i=0 to CountList2Values%>
                     <option value="<%=aTableValues(0,i) %>"><%=aTableValues(0,i)%></option>
                    <%next 
                    Set aTableValues = Nothing
                    %>
			</select>
        </TD></TR> 
		<TR><TD class=label align=right><b>Peso:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Weights" id="Peso" value="<%=Weight%>" maxlength="200" size="15" onKeyUp="res(this,numb);"></TD></TR> 
		<TR><TD class=label align=right><b>Volumen:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Volumes" id="Volumen" value="<%=Volume%>" maxlength="200" size="15" onKeyUp="res(this,numb);"></TD></TR> 
		<TR><TD class=label align=right><b>Bultos:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="NoOfPieces" id="Bultos" value="<%=TotNoOfPieces%>" maxlength="200" size="15" onKeyUp="res(this,numb);"></TD></TR> 
        <%end if %>

        <%'if GroupID = 33 then 2020-10-08 %>
            <TR><TD class=label align=right><b>Empresa:</b></TD><TD class=label align=left>
			<select name="CountrySession" id="Empresa" class="label"> <!-- cambio CountrySession x Empresa 2020-10-13 -->
				<option value='-1'>Seleccionar</option> <!-- se agrego -1 2020-10-13 -->
			    <%DisplayCountries Request("CTR"), 1 'se cambio de 2 a 1 2020-10-08 %>
			</select>
			</TD></TR>
        <%'end if%>

		<%select Case ExType
		  case 0,1,2,4,5,8,99%>
			<TR><TD class=label align=right><b>Origen:</b></TD><TD class=label align=left>
			<%Select Case Extype 
            Case -1%>
                <INPUT type="text" name="PaisO" class="label" value="<%=TranslateCountry(CountryOrigen)%>" readonly/>
                <input type="hidden" name="CountryOrigen" class="label" value="<%=CountryOrigen%>"/>
            <%Case Else %>
                <select name="CountryOrigen" id="Pais Origen" class="label">
				    <option value="-1">Seleccionar</option>
                    <!--#include file=Countries.asp-->
			    </select>
            <%End Select %>
			</TD></TR>
            <TR><TD class=label align=right><b>Destino Final:</b></TD>
			<TD class=label align=left>
			<%Select Case Extype 
            Case -1%>
                <INPUT type="text" name="PaisD" class="label" value="<%=TranslateCountry(FinalDes)%>" readonly/>
                <input type="hidden" name="CountriesFinalDes" class="label" value="<%=FinalDes%>"/>
			<%Case Else%>
				<select name="CountriesFinalDes" id="Pais Destino Final" class="label">
					<option value="-1">Seleccionar</option>
                    <!--#include file=Countries.asp-->
				</select>
			<%End Select%>
			</TD></TR>
		<%case 6,7%>
                <INPUT name="CountryOrigen" type=hidden value="<%=CountryOrigen%>">
			    <INPUT name="CountriesFinalDes" type=hidden value="<%=CountryOrigen%>">
        <%case 9,10,11,12,13,14%>
            <!--<INPUT name="CountryOrigen" type=hidden value="<%=CountryOrigen%>">-->
            <INPUT name="CountryOrigen" type=hidden value="<%=FinalDes%>">
			<INPUT name="CountriesFinalDes" type=hidden value="<%=FinalDes%>">
        <%case 15%>
            <INPUT name="CountryOrigen" type=hidden value="<%=CountryDes%>">
			<INPUT name="CountriesFinalDes" type=hidden value="<%=CountryDes%>">
		<%End Select%>

        <TR><TD class=label align=right><b>Vendedor o Usuario:</b></TD><TD class=label align=left><%=ContactSignature%></TD></TR> 
		<TR><TD class=label align=right><b>PO:</b></TD><TD class=label align=left>
			<INPUT TYPE=text class=label name="PO" id="PO" value="<%=PO%>" maxlength="200" size="35" <%If Extype <> 8 or Extype <> 15 then %> readonly <%End If %>>
		</TD></TR>
		<TR><TD class=label align=right><b>Tipo de Carga:</b></TD><TD class=label align=left>
			<select name="ChargeType" id="Tipo de Carga" class="label">
				<option value="-1">Seleccionar</option>
				<option value="0">FISCAL</option>
				<option value="1">GENERAL</option>
			</select>
		</TD></TR>

        <%Select Case ExType
        Case 8,9,10,11,12,13,14,15,99 %>
        <TR><TD class=label align=right><b>Incoterms:</b></TD><TD class=label align=left>
            <select name="IncotermsID" id="Tipo de Incoterms" class="label">
            	<option value="-1">Seleccionar</option>
                <%for i=0 to CountTableValues2%>
                    <option value="<%=aTableValues2(0,i) %>"><%=aTableValues2(1,i)%></option>
                <%next 
                Set aTableValues2 = Nothing
                %>
			</select>
            <INPUT TYPE=hidden class=label name="Incoterms" value="<%=Incoterms%>">
        </TD></TR>
        <%Case else %>
        <TR><TD class=label align=right><b>Incoterms:</b></TD><TD class=label align=left><%=Incoterms%>
        <INPUT TYPE=hidden class=label name="IncotermsID" value="<%=IncotermsID%>">
        <INPUT TYPE=hidden class=label name="Incoterms" value="<%=Incoterms%>">
        </TD></TR> 
        <%end Select %>

		<%select Case ExType
            case 0,1,2%>
		<TR><TD class=label align=right><b>Contenedor:</b></TD><TD class=label align=left><%=Container%></TD></TR> 
		<TR><TD class=label align=right><b>Master BL:</b></TD><TD class=label align=left><%=MBL%></TD></TR> 
        <%case 9,10,11,12,13,14 %>
        <TR><TD class=label align=right><b>Contenedor:</b></TD><TD class=label align=left><%=Container%><INPUT TYPE=hidden name="Container" value="<%=Container%>"></TD></TR> 
		<TR><TD class=label align=right><b>Master BL:</b></TD><TD class=label align=left><%=MBL%><INPUT TYPE=hidden name="MBLs" value="<%=MBL%>"></TD></TR> 
        <%case else %>
        <TR><TD class=label align=right><b>Contenedor:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Container" id="Contenedor" value="<%=Container%>" maxlength="200" size="35"></TD></TR> 
		<TR><TD class=label align=right><b>Master BL:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="MBLs" id="MBL" value="<%=MBL%>" maxlength="200" size="35"></TD></TR> 
        <%end select %>

		<%select Case ExType
		  case 4,5,6,7%>
		<TR><TD class=label align=right><b>RO:</b></TD><TD class=label align=left><%=BL%>&nbsp;
        <a href="#" onClick="Javascript:window.open('http://10.10.1.20/ventasV2/vendedores/detalle_routing.php?id_routing=<%=EXID %>&ref=<%=CountryDes %>', 'routing_ver', 'height=600, width=700, menubar=0, resizable=1, scrollbars=1, toolbar=0');return (false);" class="menu"><font color="FFFFFF"><b>Ver RO</b></font></a></TD></TR>
		<%case 8,15%>
		<TR><TD class=label align=right><b>House BL:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="BLs" id="House BL" value="<%=BL%>" maxlength="200" size="35"></TD></TR>
		<%case else%>
        <TR><TD class=label align=right><b>House BL:</b></TD><TD class=label align=left><%=BL%><INPUT TYPE=hidden name="BLs" id="Text4" value="<%=BL%>" ></TD></TR>
        <%end select%>
		
		<%
        select Case ExType
		  case 0,1,2,4,5,8,99%>
        <TR><TD class=label align=right><b>Endoso Aduanal-RO:</b></TD><TD class=label align=left>
			<select name="Endorse" id="Endoso Aduanal-RO" class="label">
				<option value="-1">Seleccionar</option>
				<option value="0">NO</option>
				<option value="1">SI</option>
			</select>
		</TD></TR>
		<TR><TD class=label align=right><b>Tipo Endoso Aduanal-RO:</b></TD><TD class=label align=left>
			<select name="EndorseType" id="Tipo Endoso Aduanal-RO" class="label">
				<option value="-1">Seleccionar</option>
				<option value="0">COPIA</option>
				<option value="1">ORIGINAL</option>
				<option value="2">N/A</option>
			</select>
		</TD></TR>
		<TR><TD class=label align=right><b>Declaracion de Aduana:</b></TD><TD class=label align=left>
			<select name="Declaration" id="Declaracion de Aduana" class="label">
				<option value="-1">Seleccionar</option>
				<option value="0">NO</option>
				<option value="1">SI</option>
			</select>
		</TD></TR>
		<TR>
		  <TD class=label align=right><b>Tipo Declaracion de Aduana:</b></TD>
		  <TD class=label align=left>
			<select name="DeclarationType" id="Tipo Declaracion de Aduana" class="label">
				<option value="-1">Seleccionar</option>
				<option value="0">COPIA</option>
				<option value="1">ORIGINAL</option>
				<option value="2">N/A</option>
			</select>
		</TD></TR>
		<TR><TD class=label align=right><b>Requerimiento de Partida:</b></TD><TD class=label align=left>
			<select name="RequestNo" id="Requerimiento de Partida" class="label">
				<option value="-1">Seleccionar</option>
				<option value="0">NO</option>
				<option value="1">SI</option>
			</select>
		</TD></TR>
		<TR><TD class=label align=right><b>Tipo&nbsp;Requerimiento&nbsp;de&nbsp;Partida:</b></TD><TD class=label align=left>
			<select name="RequestType" id="Tipo Requerimiento de Partida" class="label">
				<option value="-1">Seleccionar</option>
				<option value="0">COPIA</option>
				<option value="1">ORIGINAL</option>
				<option value="2">N/A</option>
			</select>
		</TD></TR>
		<TR><TD class=label align=right><b>Tipo de BL:</b></TD><TD class=label align=left>
			<select name="BLsType" id="Tipo de BL" class="label">
				<option value="-1">Seleccionar</option>
				<option value="0">COPIA</option>
				<option value="1">ORIGINAL</option>
				<option value="2">N/A</option>
			</select>
		</TD></TR>
		<TR><TD class=label align=right><b>Tipo de Factura:</b></TD><TD class=label align=left>
			<select name="BillType" id="Tipo de factura" class="label">
				<option value="-1">Seleccionar</option>
				<option value="0">COPIA</option>
				<option value="1">ORIGINAL</option>
				<option value="2">N/A</option>
			</select>
		</TD></TR>
		<TR><TD class=label align=right><b>No.Factura:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Bill" id="No.Factura" value="<%=Bill%>" maxlength="1000" size="56"></TD></TR>
		<TR><TD class=label align=right><b>Tipo de Lista de Empaque:</b></TD><TD class=label align=left>
			<select name="PackingListType" id="Tipo de Lista de Empaque" class="label">
				<option value="-1">Seleccionar</option>
				<option value="0">COPIA</option>
				<option value="1">ORIGINAL</option>
				<option value="2">N/A</option>
			</select>
		</TD></TR>
		<TR><TD class=label align=right><b>No.de Lista de Empaque:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="PackingList" id="No.de Lista de Empaque" value="<%=PackingList%>" maxlength="45"></TD></TR>
		<%case 6,7,9,10,11,12,13,14,15%>
		<TR><TD class=label align=right><b>Fecha Servicio:</b></TD>
		<TD class=label align=left>
        	<INPUT name="BLArrivalDate" type=hidden value="<%=BLArrivalDate%>">
			<select name="Yr" class=label id="A&Ntilde;O">
                <%=mesdia(Year(Now) - 10, Year(Now), 4, -1, "A&Ntilde;O")%>
			</select>/
			<select name="Mh" class=label id="Mes">
				<%=mesdia(1,12,2,-1,"MES")%>
			</select>/
			<select name="Dy" class=label id="Dia">
				<%=mesdia(1,31,2,-1,"DIA")%>
			</select>-
			<select name="Hr" class=label id="Hora">						    
                <%=mesdia(0,23,2,-1,"HORA")%>	
			</select>:<select name="Mn" class=label id="Minuto">
                <%=mesdia(0,59,2,-1,"MINUTO")%>	
			</select>
		</TD></TR>
        <TR><TD class=label align=right><b>Recolecci&oacute;n:</b></TD><TD class=label align=left>
        <textarea class="label" cols="55" rows="2" name="SenderData" id="Recoleccion"><%=SenderData%></textarea></TD></TR> 
		<TR><TD class=label align=right><b>Entrega:</b></TD><TD class=label align=left>
        <textarea class="label" cols="55" rows="2" name="ConsignerData" id="Entrega"><%=ConsignerData%></textarea></TD></TR> 
		<TR><TD class=label align=right><b>Contacto:</b></TD><TD class=label align=left>
        <textarea class="label" cols="55" rows="2" name="AgentSignature"id="Contacto"><%=AgentSignature%></textarea></TD></TR> 
		<TR><TD class=label align=right><b>Telefono Contacto:</b></TD><TD class=label align=left>
        <textarea class="label" cols="55" rows="2" name="Phone1" id="Telefono Contacto"><%=Phone1%></textarea></TD></TR> 
		<%end select%>

		<TR><TD class=label align=right><b>Observaciones:</b></TD><TD class=label align=left><textarea class="label" cols="55" rows="4" name="Observations"><%=Observations%></textarea></TD></TR>
		<TR><TD class=label align=right><b>Notificar a:</b></TD><TD class=label align=left><%=CheckNum(NotifyPartyID) & " " & NotifyParty%></TD></TR>
        <TR><TD class=label align=right></TD><TD class=label align=left><textarea class="label" cols="55" rows="2" id="Notificar a" name="Notify"><%=Notify%></textarea></TD></TR>
		
        <%if ExType=8 then %>
        <TR><TD class=label align=right><b>Declaracion Aduanera:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="CIFBrokerIn" id="Text1" value="<%=CIFBrokerIn%>" maxlength="200" size="35"></TD></TR>
		<TR><TD class=label align=right><b>Valor Flete Terrestre:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="CIFLandFreight" id="Text2" value="<%=CIFLandFreight%>" maxlength="200" size="35"></TD></TR>
		<%end if %>

        <TD colspan="2" class=label align=center>

        <% '"(" & CheckNum(CodeReference) & ")(" & fleet & ")(" & (fleet <> "NA" AND fleet <> "" AND CheckNum(CodeReference) = 0) & ")"%>

        <% if fleet <> "NA" AND fleet <> "" AND CheckNum(CodeReference) = 0 then%>
        <% 'if CheckNum(CodeReference) = 0 then%>

            <INPUT name=enviar type=button onClick="JavaScript:validar(9);" value="&nbsp;&nbsp;Validar Codigo Referencia&nbsp;&nbsp;" class=label />
      
        <% else %>

			<TABLE cellspacing=0 cellpadding=2 width=200>
		  	<TR>


            
            <%if iMensaje <> "" then%>
                
                  <TD class=label align=center colspan=2>
                  
                        <font color=red><%=iMensaje%></font>"

                  </TD>
           
			<%else%>

			<%if CountTableValues = -1 then%>
                 <%select case ExType
                 Case 6,7
                    if SenderData = "" or ConsignerData = "" Then
                        JavaMsg = "No se puede grabar la informacion, el RO esta incompleto, favor de revisar que tenga Fecha Servicio, Datos Entrega y Recoleccion"
                    End if                    
                 End Select
                 
                 if JavaMsg = "" then%>


                        <TD class=label align=center colspan=2>
                        
                        <%If ExType = 99 AND TransactionSetPurposeCode = "02" then%>
                        
                        <!-- EN ESTE CASO NO DEBE MOSTRAR EL BOTON INSERT -->
                        <INPUT name=enviar type=button value="&nbsp;&nbsp;Agregar&nbsp;&nbsp;" class=label disabled>
           
                        <%Else%>
                        <INPUT name=enviar type=button onClick="JavaScript:validar(1);" value="&nbsp;&nbsp;Agregar&nbsp;&nbsp;" class=label>
                        <%End If%>

                        </TD>
                 
                 <%else %>
                 <SCRIPT LANGUAGE="JavaScript">alert('<%=JavaMsg %>');</script>
                 <%end if %>
			<%else%>


                <%If ExType = 99 and TransactionSetPurposeCode = "02" then%>
            
			    <%else%>

				 <TD class=label align=center><INPUT name=enviar type=button onClick="JavaScript:validar(2);" value="&nbsp;&nbsp;Actualizar&nbsp;&nbsp;" class=label></TD>
                 <%if Week <> 0 then %>
                    <TD class=label align=center><INPUT name=enviar type=button onClick="JavaScript:window.open('ItineraryCharges.asp?OID=<%=ObjectID%>&GID=29','Cargos','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=990,height=480,top=150,left=150');" value="&nbsp;&nbsp;Cargos&nbsp;&nbsp;" class=label></TD>
                 <%end if%>
                 <TD class=label align=center><INPUT name=enviar type=button onClick="JavaScript:top.close();" value="&nbsp;&nbsp;Cerrar&nbsp;&nbsp;" class=label></TD>
                 </TR>
                 </TABLE>
                 <TABLE cellspacing=0 cellpadding=2 width=200>
                 <TR>
                 <TD class=label align=center><input name=rep1 type=button onClick="Javascript:window.open('InsertData.asp?OID=<%=ObjectID%>&GID=15&CD=<%=CreatedDate%>&CT=<%=CreatedTime%>','CobrosyDocumentos','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=1000,height=650,left=500');return false;" value="Cobros&nbsp;y&nbsp;Documentos" class=label></TD>
                 <%if BLID>0 then %>
                 <TD class=label align=center><input name=rep2 type=button onClick="Javascript:window.open('InsertData.asp?BLID=<%=BLID%>&GID=23&AT=0&CD=<%=CreatedDate%>&CT=<%=CreatedTime%>&CID=<%=ClientID%>','ManifiestoIngreso','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750,height=600');return false;" value="Rastreo" class=label></TD>
                 <%end if %>
                 <TR>
                 <%select Case ExType
                    case 3,4,5,8%>
                 </TR>
			     </TABLE>
                 <TABLE cellspacing=0 cellpadding=2 width=200>
		  	     <TR>
                 <TD class=label align=center><input name=rep1 type=button onClick="Javascript:window.open('Reports.asp?GID=<%=GroupID%>&AT=3&OID=<%=ObjectID%>&CID=<%=ClientID%>&AID=<%=ShipperID%>&SEP=<%=Sep%>','EndosoAduana','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750,height=600');return false;" value="Carta&nbsp;Endoso&nbsp;Aduanal" class=label></TD>
                 <TD class=label align=center><input name=rep2 type=button onClick="Javascript:window.open('Reports.asp?GID=13&AT=3&OID=<%=ObjectID%>&CID=<%=ClientID%>&AID=<%=ShipperID%>&SEP=<%=Sep%>','ManifiestoIngreso','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750,height=600');return false;" value="Manifiesto&nbsp;Ingreso" class=label></TD>
                 <%end select %>

                 <%end if %>
                 
			<%end if%>
			<%end if%>
			</TR>


            
            <%If ExType = 99 and TransactionSetPurposeCode = "02" then%>
            
			<%else%>
            <tr>
                <td colspan=3 align=center>
                <%if CountTableValues > -1 then%>
                    <input name=rep2 type=button onClick="JavaScript:Desglose('<%=ObjectID%>','<%=CreatedDate%>','<%=CreatedTime%>');return (false);" value="Desglose" class=label></TD> 
                <%end if %>
                </td>
            </tr>
			<%end if%>
			
            </TABLE>

        <%end if%>
			
		<TD>
		</TR>
	</TABLE>
	</FORM>
</BODY>
</HTML>
<script>

try {
                selecciona('forma.TypeReference','<%=TypeReference%>');
                selecciona('forma.TripType','<%=TripType%>');
                selecciona('forma.ClientWarehouse','<%=ClientWarehouse%>');
                selecciona('forma.TripPriority','<%=TripPriority%>');
}
catch(err) {

}



<%
Select Case ExType
Case 0,1,2,4,5,8,99
	if CountTableValues >= 0 then%>
		var PreviousClientsName = '<%=NameES%>';
		if (PreviousClientsName.search(/ \/ AIMAR NICARAGUA/i)>=0){
			selecciona('forma.AIMAR',' / AIMAR NICARAGUA');
		};
		if (PreviousClientsName.search(/ \/ AIMAR LOGISTIC S.A. DE C.V./i)>=0){
			selecciona('forma.AIMAR',' / AIMAR LOGISTIC S.A. DE C.V.');
		};
		if (PreviousClientsName.search(/ \/ AIMAR GUATEMALA ALSERSA/i)>=0){
			selecciona('forma.AIMAR',' / AIMAR GUATEMALA ALSERSA');
		};
		
		selecciona('forma.ChargeType','<%=ChargeType%>');
		selecciona('forma.Endorse','<%=Endorse%>');
		selecciona('forma.EndorseType','<%=EndorseType%>');
		selecciona('forma.Declaration','<%=Declaration%>');
		selecciona('forma.DeclarationType','<%=DeclarationType%>');
		selecciona('forma.RequestNo','<%=RequestNo%>');
		selecciona('forma.RequestType','<%=RequestType%>');
		selecciona('forma.BLsType','<%=BLsType%>');
		selecciona('forma.BillType','<%=BillType%>');
		selecciona('forma.PackingListType','<%=PackingListType%>');
        selecciona('forma.CountrySession','<%=Countries%>');
		<%if Action=1 or Action=2 then%>
		top.opener.location.reload();
		<%end if%>
	<%else%>
		<%if Pos>=0 then%>
		selecciona('forma.ChargeType','<%=TranslateCondition(HandlingInformation(0,0))%>');
		selecciona('forma.BillType','<%=TranslateCondition(HandlingInformation(1,0))%>');
		selecciona('forma.PackingListType','<%=TranslateCondition(HandlingInformation(2,0))%>');
		<%end if
	end if%>
        selecciona('forma.CountriesFinalDes','<%=FinalDes%>');
        selecciona('forma.CountryOrigen','<%=CountryOrigen%>');
        <%if ExType<>8 and ExType<>99 then %>
		<%else %>
        selecciona('forma.Contener','<%=Contener%>');
        selecciona('forma.IncotermsID','<%=IncotermsID%>');
        <%end if %>
<%Case 6,7,9,10,11,12,13,14,15
        if Year(BLArrivalDate)>=Year(Now) then%>
		selecciona('forma.Yr','<%=Year(BLArrivalDate)%>');
		selecciona('forma.Mh','<%=TwoDigits(Month(BLArrivalDate))%>');
		selecciona('forma.Dy','<%=TwoDigits(Day(BLArrivalDate))%>');
		selecciona('forma.Hr','<%=TwoDigits(Hour(BLArrivalDate))%>');
		selecciona('forma.Mn','<%=TwoDigits(Minute(BLArrivalDate))%>');		
		<%end if%>
        var PreviousClientsName = '<%=NameES%>';
		if (PreviousClientsName.search(/ \/ AIMAR NICARAGUA/i)>=0){
			selecciona('forma.AIMAR',' / AIMAR NICARAGUA');
		};
		if (PreviousClientsName.search(/ \/ AIMAR LOGISTIC S.A. DE C.V./i)>=0){
			selecciona('forma.AIMAR',' / AIMAR LOGISTIC S.A. DE C.V.');
		};
		if (PreviousClientsName.search(/ \/ AIMAR GUATEMALA ALSERSA/i)>=0){
			selecciona('forma.AIMAR',' / AIMAR GUATEMALA ALSERSA');
		};
        <%select case ExType
        case 9,10,11,12,13,14,15 %>
        selecciona('forma.IncotermsID','<%=IncotermsID%>');
        <%end select %>
        <%if ExType=15 then %>
            selecciona('forma.Contener','<%=Contener%>');
            selecciona('forma.ChargeType','<%=ChargeType%>');            
        <%end if %>

		<%if Action=1 or Action=2 then%>
		top.opener.location.reload();
		<%end if%>
<%End Select
if AlertSpecialClient <> "" then%>
    alert("<%=AlertSpecialClient %>");
<%end if%>
</script>