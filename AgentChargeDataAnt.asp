<%@ Language=VBscript %>
<%Option Explicit%>
<!-- #INCLUDE file="../admin/utils.asp" -->
<%
Dim XMLData, Conn, rs
Dim AgentID, AgentName, AgentUsr, AgentPwd, WithAccess
Dim xmlRequest, xnodelistAgent, xnodelistCharge, objItem, XMLResponse, sep
Dim DataAction, ClientName, ShipperName, Commodity, PackageType, Weight, Volume, Countries
Dim NoOfPieces, CountryOrigen, CountryDestiny, Container, MBL, BL, Observations, Notify, AgentNeutral, ColoadersID, ColoadersAddrID, Coloaders, EXRouting
Dim ResponseCode, ResponseDescription, ResponseCodeDescription, InTransit, BLDetailID, CreatedDate, CreatedTime
	
    XMLData = Request.Form("XMLChargeData")
    WithAccess = 0
    XMLResponse = "<ResultChargeData>"
    ResponseCode = "OK"
    
    'Cargando el XML Recibido
    If Len(XMLData) <> 0 Then  
		Set xmlRequest = CreateObject("MSXML2.DOMDocument")
		xmlRequest.async = false  
        'Cargando el XML, si al cargarlo presenta algun error se sale y notifica
        On Error Resume Next
        xmlRequest.loadXml XMLData  
        If Err.Number = 0 Then
            if InStr(XMLData,"<AgentData ")<>0 and InStr(XMLData,"</AgentData>")<>0 then
                On Error Resume Next
                Set xnodelistAgent = xmlRequest.documentElement.selectNodes("//AgentData")
			    If Err.Number = 0 Then
                    For Each objItem In xnodelistAgent   
				        'Validando que si vengan todos los nodos del XML
                        On Error Resume Next
                        AgentID = objItem.selectSingleNode("AgentID").Text
                        If Err.Number <> 0 Then
                            ResponseCode = "ERROR"
                            ResponseCodeDescription = ResponseCodeDescription & sep & "1"
                            ResponseDescription = ResponseDescription & sep & "nodo AgentID no existe en XML"
                            sep = "|"
                        End If
                        On Error Resume Next
                        AgentUsr = objItem.selectSingleNode("AgentUsr").Text
                        If Err.Number <> 0 Then
                            ResponseCode = "ERROR"
                            ResponseCodeDescription = ResponseCodeDescription & sep & "2"
                            ResponseDescription = ResponseDescription & sep & "nodo AgentUsr no existe en XML"
                            sep = "|"
                        End If
                        On Error Resume Next
                        AgentPwd = objItem.selectSingleNode("AgentPwd").Text
                        If Err.Number <> 0 Then
                            ResponseCode = "ERROR"
                            ResponseCodeDescription = ResponseCodeDescription & sep & "3"
                            ResponseDescription = ResponseDescription & sep & "nodo AgentPwd no existe en XML"
                            sep = "|"
                        End If                
			        Next
                    AgentID = CheckNum(AgentID)
                    AgentUsr = Trim(AgentUsr)
                    AgentPwd = Trim(AgentPwd)
                    'Si vienen los nodos de autenticacion continua el proceso
                    if ResponseCode = "OK" then
                        openConnMaster Conn
                            Set rs = Conn.Execute("select agente_id from accesos where es_agent=true and usr='" & AgentUsr & "' and pwd='" & AgentPwd & "' and agente_id=" & AgentID)
		                    If Not rs.EOF Then
                                If CheckNum(rs(0)) <> 0 then
                                    WithAccess = 1
                                End If
                            End If
                            'pais_terrestre_auto es el pais asignado al Agente para ingresar su carga automatica
                            if WithAccess = 1 then
                                CloseOBJ rs
                                Set rs = Conn.Execute("select agente, pais_terrestre_auto, es_neutral from agentes where agente_id=" & AgentID)
                                if Not rs.EOF then
                                    AgentName = rs(0)
                                    Countries = rs(1)
                                    AgentNeutral = CheckNum(rs(2))
                                End if
                            End if
                        CloseOBJs rs, Conn

                        'Si tiene acceso se procede a analizar los datos de la Carga
                        if WithAccess = 1 then
                            FormatTime CreatedDate, CreatedTime
                                
                            Set xnodelistCharge = xmlRequest.documentElement.selectNodes("//RowData")
                            For Each objItem In xnodelistCharge
				                CreatedTime = CreatedTime + 1
                                sep = ""

                                'Validacion de Formato y Longitudes
                                ResponseCode = "OK"
                                ResponseDescription = ""
                                ResponseCodeDescription = ""
                                InTransit = -1
                                BLDetailID = 0
                    
                                'Validando que si vengan todos los nodos del XML
                                On Error Resume Next
                                DataAction = objItem.selectSingleNode("DataAction").Text
                                If Err.Number <> 0 Then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "4"
                                    ResponseDescription = ResponseDescription & sep & "nodo DataAction no existe en XML"
                                    sep = "|"
                                End If
                                On Error Resume Next
                                ClientName = objItem.selectSingleNode("ClientName").Text
                                If Err.Number <> 0 Then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "5"
                                    ResponseDescription = ResponseDescription & sep & "nodo ClientName no existe en XML"
                                    sep = "|"
                                End If
                                On Error Resume Next
                                ShipperName = objItem.selectSingleNode("ShipperName").Text
				                If Err.Number <> 0 Then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "6"
                                    ResponseDescription = ResponseDescription & sep & "nodo ShipperName no existe en XML"
                                    sep = "|"
                                End If
                                On Error Resume Next
                                Commodity = objItem.selectSingleNode("Commodity").Text
				                If Err.Number <> 0 Then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "7"
                                    ResponseDescription = ResponseDescription & sep & "nodo Commodity no existe en XML"
                                    sep = "|"
                                End If
                                On Error Resume Next
                                PackageType = UCase(objItem.selectSingleNode("PackageType").Text)
				                If Err.Number <> 0 Then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "8"
                                    ResponseDescription = ResponseDescription & sep & "nodo PackageType no existe en XML"
                                    sep = "|"
                                End If
				                On Error Resume Next
                                Weight = objItem.selectSingleNode("Weight").Text
				                If Err.Number <> 0 Then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "9"
                                    ResponseDescription = ResponseDescription & sep & "nodo Weight no existe en XML"
                                    sep = "|"
                                End If
				                On Error Resume Next
                                Volume = objItem.selectSingleNode("Volume").Text
				                If Err.Number <> 0 Then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "10"
                                    ResponseDescription = ResponseDescription & sep & "nodo Volume no existe en XML"
                                    sep = "|"
                                End If
				                On Error Resume Next
                                NoOfPieces = objItem.selectSingleNode("NoOfPieces").Text
				                If Err.Number <> 0 Then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "11"
                                    ResponseDescription = ResponseDescription & sep & "nodo NoOfPieces no existe en XML"
                                    sep = "|"
                                End If
				                On Error Resume Next
                                CountryOrigen = UCase(objItem.selectSingleNode("CountryOrigen").Text)
				                If Err.Number <> 0 Then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "12"
                                    ResponseDescription = ResponseDescription & sep & "nodo CountryOrigen no existe en XML"
                                    sep = "|"
                                End If
                                On Error Resume Next
                                CountryDestiny = UCase(objItem.selectSingleNode("CountryDestiny").Text)
                                If Err.Number <> 0 Then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "13"
                                    ResponseDescription = ResponseDescription & sep & "nodo CountryDestiny no existe en XML"
                                    sep = "|"
                                End If
                                On Error Resume Next
                                Container = objItem.selectSingleNode("Container").Text
                                If Err.Number <> 0 Then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "14"
                                    ResponseDescription = ResponseDescription & sep & "nodo Container no existe en XML"
                                    sep = "|"
                                End If
				                On Error Resume Next
                                MBL = UCase(objItem.selectSingleNode("MBL").Text)
				                If Err.Number <> 0 Then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "15"
                                    ResponseDescription = ResponseDescription & sep & "nodo MBL no existe en XML"
                                    sep = "|"
                                End If
				                On Error Resume Next
                                BL = UCase(objItem.selectSingleNode("BL").Text)
                                If Err.Number <> 0 Then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "16"
                                    ResponseDescription = ResponseDescription & sep & "nodo BL no existe en XML"
                                    sep = "|"
                                End If
                                On Error Resume Next
                                Observations = objItem.selectSingleNode("Observations").Text
                                If Err.Number <> 0 Then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "17"
                                    ResponseDescription = ResponseDescription & sep & "nodo Observations no existe en XML"
                                    sep = "|"
                                End If
                                On Error Resume Next
                                Notify = objItem.selectSingleNode("Notify").Text
                                If Err.Number <> 0 Then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "18"
                                    ResponseDescription = ResponseDescription & sep & "nodo Notify no existe en XML"
                                    sep = "|"
                                End If

                                On Error Resume Next '2019-08-12
                                ColoadersID = objItem.selectSingleNode("Coloader").Text
                                If Err.Number <> 0 Then
									ColoadersID = 0
                                    'ResponseCode = "ERROR"
                                    'ResponseCodeDescription = ResponseCodeDescription & sep & "19"
                                    'ResponseDescription = ResponseDescription & sep & "nodo Coloader no existe en XML"
                                    'sep = "|"
                                End If
								
								
                                On Error Resume Next '2019-09-02
                                EXRouting = objItem.selectSingleNode("Routing").Text
                                If Err.Number <> 0 Then
									EXRouting = ""
                                    'ResponseCode = "ERROR"
                                    'ResponseCodeDescription = ResponseCodeDescription & sep & "19"
                                    'ResponseDescription = ResponseDescription & sep & "nodo EXRouting no existe en XML"
                                    'sep = "|"
                                End If
								
                                'Validando formatos y longitudess de cada nodo del XML
                                if DataAction<>0 and DataAction<>1 and DataAction<>2 then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "19"
                                    ResponseDescription = ResponseDescription & sep & "DataAction invalido"
                                    sep = "|"
                                end if
                                if Len(ClientName)>200 then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "20"
                                    ResponseDescription = ResponseDescription & sep & "ClientName excede longitud permitida (200)"
                                    sep = "|"
                                end if
                                if Len(ShipperName)>200 then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "21"
                                    ResponseDescription = ResponseDescription & sep & "ShipperName excede longitud permitida (200)"
                                    sep = "|"
                                end if
                                if Len(Commodity)>500 then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "22"
                                    ResponseDescription = ResponseDescription & sep & "Commodity excede longitud permitida (500)"
                                    sep = "|"
                                end if
                                openConnMaster Conn
                                    Set rs = Conn.Execute("select tipo from tipo_paquete where tipo='" & PackageType & "'")
		                            If Not rs.EOF Then
                                        if rs(0)="" then
                                            ResponseCode = "ERROR"
                                            ResponseCodeDescription = ResponseCodeDescription & sep & "23"
                                            ResponseDescription = ResponseDescription & sep & "PackageType Invalido"
                                            sep = "|"
                                        end if
                                    else
                                        ResponseCode = "ERROR"
                                        ResponseCodeDescription = ResponseCodeDescription & sep & "23"
                                        ResponseDescription = ResponseDescription & sep & "PackageType Invalido"
                                        sep = "|"
                                    End If
                                CloseOBJs rs, Conn
                                if Not IsNumeric(Weight) then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "24"
                                    ResponseDescription = ResponseDescription & sep & "Weight no es numerico"
                                    sep = "|"
                                end if
                                if Not IsNumeric(Volume) then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "25"
                                    ResponseDescription = ResponseDescription & sep & "Volume no es numerico"
                                    sep = "|"
                                end if
                                if Not IsNumeric(NoOfPieces) then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "26"
                                    ResponseDescription = ResponseDescription & sep & "NoOfPieces no es numerico"
                                    sep = "|"
                                end if
                                if Len(CountryOrigen)<>2 then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "27"
                                    ResponseDescription = ResponseDescription & sep & "CountryOrigen no cumple longitud requerida (2)"
                                    sep = "|"
                                end if
                                if IsNumeric(CountryOrigen) then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "28"
                                    ResponseDescription = ResponseDescription & sep & "CountryOrigen debe ser string"
                                    sep = "|"
                                end if
                                if Len(CountryDestiny)<>2 then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "29"
                                    ResponseDescription = ResponseDescription & sep & "CountryDestiny no cumple longitud requerida (2)"
                                    sep = "|"
                                end if
                                if IsNumeric(CountryDestiny) then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "30"
                                    ResponseDescription = ResponseDescription & sep & "CountryDestiny debe ser string"
                                    sep = "|"
                                end if
                                if Len(Container)>60 then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "31"
                                    ResponseDescription = ResponseDescription & sep & "Container excede longitud permitida (60)"
                                    sep = "|"
                                end if
                                if Len(MBL)>45 then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "32"
                                    ResponseDescription = ResponseDescription & sep & "MBL excede longitud permitida (45)"
                                    sep = "|"
                                end if
                                if Len(BL)>45 then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "33"
                                    ResponseDescription = ResponseDescription & sep & "BL excede longitud permitida (45)"
                                    sep = "|"
                                end if
                                if Len(BL)<5 then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "34"
                                    ResponseDescription = ResponseDescription & sep & "BL no cumple con longitud minima (5)"
                                    sep = "|"
                                end if
                                if Len(Observations)>300 then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "35"
                                    ResponseDescription = ResponseDescription & sep & "Observations excede longitud permitida (300)"
                                    sep = "|"
                                end if
				                if Len(Notify)>300 then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "36"
                                    ResponseDescription = ResponseDescription & sep & "Notify excede longitud permitida (300)"
                                    sep = "|"
                                end if

								if Len(ColoadersID)>5 then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "37"
                                    ResponseDescription = ResponseDescription & sep & "Coloader excede longitud permitida (5)"
                                    sep = "|"
                                end if

								if Len(EXRouting)>50 then
                                    ResponseCode = "ERROR"
                                    ResponseCodeDescription = ResponseCodeDescription & sep & "38"
                                    ResponseDescription = ResponseDescription & sep & "Routing excede longitud permitida (50)"
                                    sep = "|"
                                end if								
								
                                'Si cumple con la validacion de formatos y longitudes se revisa el DataAction, si es 0 (insert) revisa repetidos
                                if ResponseCode = "OK" then
                        
									if (ColoadersID = "GT") then
										ColoadersID = "GTAIM"
									end if
									
									select Case ColoadersID
									'AIMAR
									case "GTAIM" 
										ColoadersID = 40396										
										
									'GRUPO TLA
									case "GTTLA"
										ColoadersID = 29457 '61920 este es el que se ha usado

									case "SVTLA"
										ColoadersID = 65768
										
									case "HNTLA"
										ColoadersID = 93049

									case "NITLA"
										ColoadersID = 77222 
										
									case "N1TLA"
										ColoadersID = 63709 '67728 

									case "CRTLA"
										ColoadersID = 43421
																				
									case "PATLA"
										ColoadersID = 75002

									'LATIN FREIGHT
									case "GTLTF"
										ColoadersID = 7052
										
									case else
										ColoadersID = 0
									end select

									ColoadersAddrID = 0
									Coloaders = ""
										
									If ColoadersID > 0 Then	
										openConnMaster Conn
										Set rs = Conn.Execute("SELECT a.id_cliente, a.nombre_cliente, b.id_direccion FROM clientes a, direcciones b WHERE a.id_cliente = " & ColoadersID & " AND a.id_cliente = b.id_cliente LIMIT 1")
										If Not rs.EOF Then											
											ColoadersID = CheckNum(rs(0)) 
											Coloaders = rs(1)									
											ColoadersAddrID = CheckNum(rs(2))
										End If			
										CloseOBJs rs, Conn
									End If									
						
                                    openConnLand Conn
                                    select Case DataAction
                                    case 0
                                        'Creando Script para Insertar la Informacion
                                        'Valida que no exista el BL para evitar duplicados
                                        Set rs = Conn.Execute("select BLs from BLDetail where Expired=0 and BLs='" & BL & "' and ShippersID=" & AgentID & " and Countries='" & Countries & "'")
                                        if rs.EOF then
											
											On Error Resume Next  
											
											Conn.Execute("insert into BLDetail (Clients, Agents, DiceContener, Priority, ExType, CreatedDate, CreatedTime, ShippersID, " & _
                                            "Shippers, ClientsTemp, AgentsTemp, DiceContenerTemp, " & _
                                            "ClassNoOfPieces, Weights, Volumes, NoOfPieces, CountryOrigen, CountriesFinalDes, " & _
                                            "Container, MBLs, BLs, Observations, Notify, Countries, EXDBCountry, AgentNeutral, ColoadersID, ColoadersAddrID, Coloaders, EXRouting) values ('', '', '', 1, 8, " & _
                                            "'" & CreatedDate & "', " & _
                                            CreatedTime & ", " & _
                                            AgentID & ", " & _
                                            "'" & AgentName & "', " & _
                                            "'" & ClientName & "', " & _
                                            "'" & ShipperName & "', " & _
                                            "'" & Commodity & "', " & _
                                            "'" & PackageType & "', " & _
                                            Weight & ", " & _
                                            Volume & ", " & _
                                            NoOfPieces & ", " & _
                                            "'" & CountryOrigen & "', " & _
                                            "'" & CountryDestiny & "', " & _
                                            "'" & Container & "', " & _
                                            "'" & MBL & "', " & _
                                            "'" & BL & "', " & _
                                            "'" & Observations & "', " & _
                                            "'" & Notify & "', " & _
                                            "'" & Countries & "', " & _
                                            "'" & Countries & "', " & _
AgentNeutral & ", " & ColoadersID & ", " & ColoadersAddrID & ", '" & Coloaders  & "', '" & EXRouting & "')")
                                
											If Err.Number <> 0 Then
												ResponseCodeDescription = Err.Number
												ResponseDescription = Err.description
												sep = "|"
											else
												ResponseCodeDescription = ResponseCodeDescription & sep & "37"
												ResponseDescription = ResponseDescription & sep & "BL ingresado exitosamente"
												sep = "|"
											end if	

                                        else
                                            ResponseCode = "ERROR"
                                            ResponseCodeDescription = ResponseCodeDescription & sep & "38"
                                            ResponseDescription = ResponseDescription & sep & "BL ya existe en sistema de AIMAR, solo puede actualizarse enviando DataAction=1"
                                            sep = "|"
                                        end if
                                        CloseOBJ rs
                                    Case 1
                                        'Creando Script para Actualizar la Informacion
                                        Set rs = Conn.Execute("select InTransit, BLDetailID from BLDetail where Expired=0 and BLs='" & BL & "' and ShippersID=" & AgentID & " and Countries='" & Countries & "'")
                                        if Not rs.EOF then
                                            Intransit = CheckNum(rs(0))
                                            BLDetailID = rs(1)
                                        end if
                                        CloseOBJ rs
                            
                                        'Si InTransit=0 la carga todavia no esta asignada a un Itinerario y puede recibir modificacion automatica
                                        'Se guarda ActionTemp=1 para indicar que es actualizacion del lado del Agente
                                        'Se limpian los datos de Cliente, Shipper y Producto por si el agente actualizo esos datos y asi trafico vuelva
                                        'a reasignar los datos correspondientes
                                        select Case Intransit
                                        Case -1
                                            ResponseCode = "ERROR"
                                            ResponseCodeDescription = ResponseCodeDescription & sep & "39"
                                            ResponseDescription = ResponseDescription & sep & "BL no existe en sistema de AIMAR para poder actualizarlo"
                                            sep = "|"
                                        Case 0
											
											On Error Resume Next
                                            Conn.Execute("update BLDetail set ActionTemp=1, " & _
                                            "Clients='', ClientsID=0, AddressesID=0, " & _
                                            "Agents='', AgentsID=0, AgentsAddrID=0, " & _
                                            "DiceContener='', CommoditiesID=0, " & _
                                            "ClientsTemp='" & ClientName & "', " & _
                                            "AgentsTemp='" & ShipperName & "', " & _
                                            "DiceContenerTemp='" & Commodity & "', " & _
                                            "ClassNoOfPieces='" & PackageType & "', " & _
                                            "Weights=" & Weight & ", " & _
                                            "Volumes=" & Volume & ", " & _
                                            "NoOfPieces=" & NoOfPieces & ", " & _
                                            "CountryOrigen='" & CountryOrigen & "', " & _
                                            "CountriesFinalDes='" & CountryDestiny & "', " & _
                                            "Container='" & Container & "', " & _
                                            "MBLs='" & MBL & "', " & _
                                            "Observations='" & Observations & "', " & _
                                            "Notify='" & Notify & "', " & _
											"ColoadersID=" & ColoadersID & ", ColoadersAddrID=" & ColoadersAddrID & ", Coloaders='" & Coloaders & "', EXRouting = '" & EXRouting & "' where BLDetailID=" & BLDetailID)

											If Err.Number <> 0 Then
												ResponseCodeDescription = Err.Number
												ResponseDescription = Err.description
												sep = "|"
											else
												ResponseCodeDescription = ResponseCodeDescription & sep & "40"
												ResponseDescription = ResponseDescription & sep & "BL actualizado exitosamente"
												sep = "|"
											end if	

                                        Case Else
                                            ResponseCode = "ERROR"
                                            ResponseCodeDescription = ResponseCodeDescription & sep & "41"
                                            ResponseDescription = ResponseDescription & sep & "BL no puede actualizarse porque el documento ya tiene estado ASIGNADO, debe comunicarse con el personal de Transporte Terrestre de AIMAR"
                                            sep = "|"
                                        End Select
                                    Case 2
                                        'Creando Script para Desactivar(Eliminar) la Informacion
                                        Set rs = Conn.Execute("select InTransit, BLDetailID from BLDetail where Expired=0 and BLs='" & BL & "' and ShippersID=" & AgentID & " and Countries='" & Countries & "'")
                                        if Not rs.EOF then
                                            Intransit = CheckNum(rs(0))
                                            BLDetailID = rs(1)
                                        end if
                                        CloseOBJ rs

                                        'Si InTransit=0 la carga todavia no esta asignada a un Itinerario y puede recibir eliminacion automatica
                                        'Se guarda ActionTemp=2 para indicar que es eliminacion del lado del Agente
                                        'Se limpian los datos de Cliente, Shipper y Producto por si el agente actualizo esos datos
                                        select Case Intransit
                                        Case -1
                                            ResponseCode = "ERROR"
                                            ResponseCodeDescription = ResponseCodeDescription & sep & "42"
                                            ResponseDescription = ResponseDescription & sep & "BL no existe en sistema de AIMAR para poder eliminarlo"
                                            sep = "|"
                                        Case 0
                                            Conn.Execute("update BLDetail set ActionTemp=2, Expired=1, " & _
                                            "Clients='', ClientsID=0, AddressesID=0, " & _
                                            "Agents='', AgentsID=0, AgentsAddrID=0, " & _
                                            "DiceContener='', CommoditiesID=0, " & _ 
											"ColoadersID = 0, ColoadersAddrID = '', Coloaders  = '' " & _
                                            " where BLDetailID=" & BLDetailID)

                                            ResponseCodeDescription = ResponseCodeDescription & sep & "43"
                                            ResponseDescription = ResponseDescription & sep & "BL eliminado exitosamente"
                                            sep = "|"
                                        Case Else
                                            ResponseCode = "ERROR"
                                            ResponseCodeDescription = ResponseCodeDescription & sep & "44"
                                            ResponseDescription = ResponseDescription & sep & "BL no puede eliminarse porque el documento ya tiene estado ASIGNADO, debe comunicarse con el personal de Transporte Terrestre de AIMAR"
                                            sep = "|"
                                        End Select
                                    End Select
                                end if
                                CloseOBJ Conn
                                
                                XMLResponse = XMLResponse & _
                                    "<ResultRowData>" & _
                                    "<BL>" & BL & "</BL>" & _
                                    "<ResultCode>" & ResponseCode & "</ResultCode>" & _
                                    "<ResultDescription>" & ResponseDescription & "</ResultDescription>" & _
                                    "<ResultDescriptionCode>" & ResponseCodeDescription & "</ResultDescriptionCode>" & _
                                    "</ResultRowData>"
			                Next
                        Else
                            XMLResponse = XMLResponse & _
                                "<ResultRowData>" & _
                                "<BL></BL>" & _
                                "<ResultCode>ERROR</ResultCode>" & _
                                "<ResultDescription>Acceso Invalido</ResultDescription>" & _
                                "<ResultDescriptionCode>45</ResultDescriptionCode>" & _
                                "</ResultRowData>"
                        End if
                    Else
                        XMLResponse = XMLResponse & _
                            "<ResultRowData>" & _
                            "<BL>" & BL & "</BL>" & _
                            "<ResultCode>" & ResponseCode & "</ResultCode>" & _
                            "<ResultDescriptionCode>" & ResponseCodeDescription & "</ResultDescriptionCode>" & _
                            "</ResultRowData>"
                    End if
                Else
                    XMLResponse = XMLResponse & _
                        "<ResultRowData>" & _
                        "<BL></BL>" & _
                        "<ResultCode>ERROR</ResultCode>" & _
                        "<ResultDescription>la estructura del XML no esta correcta</ResultDescription>" & _
                        "<ResultDescriptionCode>46</ResultDescriptionCode>" & _
                        "</ResultRowData>"
                End if
		    Else
                XMLResponse = XMLResponse & _
                    "<ResultRowData>" & _
                    "<BL></BL>" & _
                    "<ResultCode>ERROR</ResultCode>" & _
                    "<ResultDescription>nodo AgentData no existe en XML</ResultDescription>" & _
                    "<ResultDescriptionCode>47</ResultDescriptionCode>" & _
                    "</ResultRowData>"				    
            End If
            set xmlRequest = nothing
        Else
            XMLResponse = XMLResponse & _
                "<ResultRowData>" & _
                "<BL></BL>" & _
                "<ResultCode>ERROR</ResultCode>" & _
                "<ResultDescription>la estructura del XML no esta correcta</ResultDescription>" & _
                "<ResultDescriptionCode>46</ResultDescriptionCode>" & _
                "</ResultRowData>"				    
        End If
	Else
        XMLResponse = XMLResponse & _
            "<ResultRowData>" & _
            "<BL></BL>" & _
            "<ResultCode>ERROR</ResultCode>" & _
            "<ResultDescription>XML Invalido</ResultDescription>" & _
            "<ResultDescriptionCode>48</ResultDescriptionCode>" & _
            "</ResultRowData>"        
	End if
    XMLResponse = XMLResponse & "</ResultChargeData>"
    response.write "<?xml version=""1.0"" encoding=""UTF-8""?>" & XMLResponse
%>