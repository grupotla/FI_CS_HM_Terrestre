<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<% 
if Session("OperatorID") <> 0 then
Checking "0|1|2"
Dim JavaMsg, SQLQuery, DisplayPost, CountTableValues, GroupID, SQLQuery2, BusinessGID, CountTableValues2
Dim Conn, Connx, rs, rs2, rs3, Action, rsFilter, ObjectID, CanDisplayInfo, aTableValues, HasSecure, UnitNo, BLSize, Marchamo, aTableValues2
Dim TableName, ObjectName, QuerySelect, CreatedDate, CreatedTime, i, j, RangeID, RangeESID, EDIStatusID, SenderAddrID, ShipperAddrID, ConsignerAddrID
Dim CarrierID, AirportID, TerminalFeeCS, TerminalFeePD, CustomFee, FuelSurcharge, SecurityFee, Comment, Comment2, Comment4
Dim CountList1Values, CountList2Values, CountList3Values, CountList4Values, CountList5Values, CountList6Values, CountList7Values, CountListEValues, ClientID, Passport, NotifyPartyID, NotifyPartyAddrID, NotifyParty, NotifyPartyColoader
Dim aList1Values, aList2Values, aList3Values, aList4Values, aList5Values, aList6Values, aList7Values, aList8Values, aList9Values, aList10Values, aList11Values, aList12Values, aList13Values, aList14Values
Dim Val, NameES, NameEN, TypeVal, CommodityCode, CurrencyCode, Xchange, Logo, Footer, Estate, BillName, AddressID, Region, DepartureDate
Dim ProviderID, TruckNo, License, CodProv, Model, Mark, SenderID, ConsignerID, ShipperID, WhereSQL, PhoneID, AttnID, BLArrivalHour, BLArrivalMin, DeliveryDate
Dim SenderData, CountryDep, CountryDes, CouDesValue, BrokerID, PilotID, TruckID, Bail, ntr, AgentID, ReqAuth, AccountNo, IATANo, isConsigneer
Dim ChargePlace, FinalDes, Container, TotVolume, Freight, Insurance, ContainerDep, Freight2, Insurance2, ReservationDate, isShipper
Dim Address, Address2, Phone1, Phone2, Attn, Expired, AirportCode, Name, Week, Consolidated, Closed, BLIDs, ClientIDs, DTIObservations
Dim OtherChargesPrintType, BLNumber, ShipperData, BLType, AnotherChargesCollect, Contener, AnotherChargesPrepaid
Dim ConsignerData, AgentData, HandlingInformation, Observations, HTMLCode, BLExitDate, BLRealExitDate, BLArrivalDate, WType, Pos
Dim TotNoOfPieces, TotWeight, TotPrepaid, TotCollect, ContactSignature, AgentSignature, Countries, SearchOption, TotDiceContenerValue
Dim LtAcceptNumber, LtAcceptDate, BrokerRecepID, BLID, LtEndorseDate, Chassis, DTI, Email, ChargeType, DestinyType
Dim Arancel_GT, Arancel_SV, Arancel_HN, Arancel_NI, Arancel_CR, Arancel_PA, Arancel_BZ, BLEstArrivalDate, Motor, Axes, Tara
Dim BLDetailID(), BLIDTransit(), NoOfPieces(), GuiaRemisionDet(), ClassNoOfPieces(), CommoditiesID(), DiceContener(), DiceContenerValue(), CountriesOrigen()
Dim UpdateBLTransit(), ClientsID(), AddressesID(), Clients(), BLs(), DischargeDate(), CountriesFinalDes(), InTransit()
Dim DetailInsert(), DetailUpdate(), AgentsID(), Agents(), Volumes(), Weights(), HBLNumber(), AgentsAddrID(), Seps(), HaveInvoices(), DetailInvoices(), Posi()
Dim UserCreate, UserModify, CreatedIn, EXID, EXType, EXType99, BLDetailID99, HBLNumber99, WareHouseDischargeDate, Commodity, Weight, Volume, AgentAddrID
Dim BL, MBL, Endorse, EndorseType, Declaration, DeclarationType, RequestNo, RequestType, BLsType, BillType, Bill, Routing_Seg, Seguro, Poliza_Seguro, RSeg_Borrado, RSeg_Poliza, RSeg_BLID, Routing_Adu, RAdu_Activo, RAdu_Borrado, RAdu_NDUA, RAdu_BLID
Dim PackingListType, PackingList, Sep, CountryOrigen, ArrivalNotes, LtArrivalDate, LtArrivalDeliveryDocs, Notify
Dim DeliveryPolicyDate, DeliveryPolicyHour, DeliveryPolicyMin, PolicyNo, ItineraryType, BLFinishHour, BLFinishMin
Dim EndorseObservations, CPDocType, ManifestDocType, EndorseDocType, DTIDocType, RoutingID, BLDispatchDate
Dim WarehouseID, StartValue, ActualValue, FinishValue, BagValue, UserID, CIFLandFreight, CIFBrokerIn, PO, DBVentas, CorreoTranshipper
Dim ClientsTemp, AgentsTemp, DiceContenerTemp, PilotInstructions, ColoaderData, ColoaderID, ColoaderAddrID
Dim ClientColoader, ShipperColoader, AgentNeutral, GuiaRemision, Incoterms, IncotermsID, AlertSpecialClient, ClientCollectID, ClientsCollect, NumeroCerosGuiaRemision, CountGuiaRemision
Dim encValues, CountListValues
Dim FreightColoader, FreightColoader2, InsuranceColoader, InsuranceColoader2, AnotherChargesColoader, AnotherChargesColoader2
Dim RefHBLNumber, RefBLID, RosClientID
Dim BLsFreight, BLsFreight2, BLsInsurance, BLsInsurance2, BLsAnotherChargesPrepaid, BLsAnotherChargesCollect

	GroupID = CheckNum(Request("GID")) 
	ObjectID = CheckNum(Request("OID"))
	SQLQuery = ""
	CountTableValues = -1
	Action = CheckNum(Request("Action"))
	CreatedDate = CheckTxt(Request("CD"))
	CreatedTime = CheckNum(Request("CT"))
	BLType = CheckNum(Request("AT"))
	SearchOption = CheckNum(Request("SO"))
	AddressID=0
    Set rs2 = CreateObject("ADODB.Recordset")

    FormatTime CreatedDate, CreatedTime
	

	If GroupID >= 1 And GroupID <= 37 Then
			GetTableData GroupID, TableName, ObjectName, QuerySelect, BLType
			'Preparando el Query de Seleccion
			Select Case GroupID
			Case 2, 3, 4, 9, 11, 20
				openConn2 Conn 'Abriendo la conexion a BBDD Master
			Case Else
				WhereSQL = "CreatedDate='" & CreatedDate & "' and CreatedTime="
				openConn Conn 'Abriendo la conexion a BBDD
			End Select

			If Action >= 1 And Action <= 3 Then
                 'obteniendo los parametros para hacer las operaciones de Insert, Update o Delete
                 'Creando los filtros para cada opcion de almacenamiento
				Select Case GroupID
				Case -2
					RsFilter = "select * from " & TableName & " where " & ObjectName & "=" & ObjectID '"id_cliente=" & ObjectID
					WhereSQL = ""
				'Case 1 'Awb
				'	BLNumber = request.Form("BLNumber")
				'	Countries = request.Form("Countries")
				'	RsFilter = " BLNumber='" & BLNumber & "' and Countries='" & Countries & "'"
				Case 7, 8, 10
 					Name = request.Form("Name")
					'Countries = request.Form("Countries")
					'RsFilter = " Name='" & Name & "' and Countries='" & Countries & "'"
					RsFilter = " Name='" & Name & "'"
				Case 2, 3, 4, 11, 20
					Name = request.Form("Name")
					'RsFilter = " nombre_cliente='" & Name & "'"
					RsFilter = "select * from " & TableName & " where " & ObjectName & "=" & ObjectID '"id_cliente=" & ObjectID
					WhereSQL = "fecha_creacion='" & CreatedDate & "' and hora_creacion="
				Case 5
					Name = request.Form("Name")
					Countries = request.Form("Countries")
				  	'ProviderID = CheckNum(request.Form("ProviderID"))
				  	RsFilter = " Name='" & Name & "' and Countries='" & Countries & "'"'and ProviderID=" & ProviderID
				'Case 6
				'	TruckNo = request.Form("TruckNo")
				'	Countries = request.Form("Countries")
				'  	ProviderID = CheckNum(request.Form("ProviderID"))
				'  	RsFilter = " TruckNo='" & TruckNo & "' and Countries='" & Countries & "' and ProviderID=" & ProviderID
				Case 7
 					Name = request.Form("Name")
					Countries = request.Form("Countries")
					CodProv = request.Form("CodProv")
					RsFilter = " Name='" & Name & "' and Countries='" & Countries & "' or CodProv='" & CodProv & "'"
				Case 9
					RsFilter = "select * from " & TableName & " where " & ObjectName & "=" & ObjectID
					WhereSQL = "createddate='" & CreatedDate & "' and createdtime="
				Case 14, 15
 					RsFilter = ObjectName & "=" & ObjectID
				Case 23
					FormatTime "", 0
					RsFilter = ObjectName & "=" & ObjectID & " and CreatedTime=" & CreatedTime
				Case 27, 28, 33, 37
					RsFilter = "(" & ObjectName & "=" & ObjectID & " and CreatedTime=" & CreatedTime & ")"
					if CheckNum(Request("EXID"))<>0 then
						RsFilter = RsFilter & " or (EXID=" & CheckNum(Request("EXID")) & " and ExType=" & CheckNum(Request("ET")) & ")"
					end if
				Case Else
					RsFilter = ObjectName & "=" & ObjectID & " and CreatedTime=" & CreatedTime
				End Select

        				  'response.write RsFilter & "<br>"
        On Error Resume Next      
                   
				  Select Case GroupID
				  Case 2, 3, 4, 9, 11, 20
					  Set rs = Server.CreateObject("ADODB.Recordset")
				      rs.Open RsFilter, Conn, 2, 3, 1
				  Case Else
					  openTable Conn, TableName, rs 'Abriendo la base de Datos
					  rs.Filter = RsFilter
				  End Select
				  'response.write RsFilter & "<br>"

        If Err.Number<>0 then
	        response.write "InsertData :" & Err.Number & " - " & Err.Description & "<br>"  
        end if
                 
				  Select Case Action
                  Case 1 ' Insert
                        If rs.EOF Then 'Si no existe el atributo, puede ingresarlo


                            if GroupID = 34 then
                            
                                response.write "Aqui nunca deberia proceder"

                            else
                                


                                SaveInfo Conn, rs, Action, GroupID, CreatedDate, CreatedTime, BLType
                                closeOBJ rs
                                Select Case GroupID
                                Case 27, 28, 33
                                    Set rs = Conn.Execute("select " & ObjectName & ", EXID, EXType, EXDBCountry from " & TableName & " a where " & WhereSQL & CreatedTime)
                                Case Else
                                    'response.write("select " & ObjectName & " from " & TableName & " where " & WhereSQL & CreatedTime)
                                    Set rs = Conn.Execute("select " & ObjectName & " from " & TableName & " where " & WhereSQL & CreatedTime)
                                End Select
                                If Not rs.EOF Then
                                    ObjectID = CheckNum(rs(0))
                                    Select Case GroupID
                                    Case 27, 28, 33
                                        EXID = CheckNum(rs(1))
                                    End Select
								    Select Case GroupID 'Actualizando datos de la tabla master (Clientes, Shippers, Exporters)
								    Case 3, 4, 11, 20
									    AddressID = SaveMaster (Conn, ObjectID)
								    End Select
                                    Select Case GroupID
                                    Case 28
                                        Select Case rs(2)
                                        Case 4,5,6,7
                                            OpenConn2 Connx
                                                Set rs2 = Connx.Execute("select routing_seg, routing_adu from routings where id_routing = " & EXID & "")
                                                'response.write("select routing_seg, routing_adu from routings where id_routing = " & EXID & "")
                                                'response.write("<br>")
                                                if Not rs2.EOF then
                                                    if (rs2(0) = 0 and rs2(1) = 0) then
                                                        Connx.Execute("update routings set bl_id=" & ObjectID & " where id_routing = " & EXID & " and borrado=false ")
                                                        'Response.write("update routings set bl_id=" & ObjectID & " where id_routing = " & EXID & " and borrado=false ")
                                                    elseif (rs2(0) <> 0 and rs2(1) = 0) then
                                                        Connx.Execute("update routings set bl_id=" & ObjectID & " where ((id_routing = " & EXID & " and borrado=false) or id_routing = " & rs2(0) & ") ")
                                                        'Response.write("update routings set bl_id=" & ObjectID & " where ((id_routing = " & EXID & " and borrado=false) or id_routing = " & rs2(0) & ") ")
                                                    elseif (rs2(0) = 0 and rs2(1) <> 0) then
                                                        Connx.Execute("update routings set bl_id=" & ObjectID & " where ((id_routing = " & EXID & " and borrado=false) or id_routing = " & rs2(1) & ") ")
                                                        'response.write("update routings set bl_id=" & ObjectID & " where ((id_routing = " & EXID & " and borrado=false) or id_routing = " & rs2(1) & ") ")
                                                    elseif (rs2(0) <> 0 and rs2(1) <> 0) then
                                                        Connx.Execute("update routings set bl_id=" & ObjectID & " where id_routing in (" & EXID & "," & rs2(0) & "," & rs2(1) & ")")
                                                        'response.write("update routings set bl_id=" & ObjectID & " where id_routing in (" & EXID & "," & rs2(0) & "," & rs2(1) & ")")
                                                    End if
                                                End if
                                            CloseOBJs Connx, rs2
                                        End Select
                                    Case 27
                                        Select Case rs(2)
                                            Case 1,2,12,13
                                                openConnOcean Connx, "ventas_" & LCase(rs(3))
                                                    'response.write("ventas_" & LCase(rs(3)))
                                                    'response.write("update bill_of_lading set ref_id = " & ObjectID & " where bl_id = " & EXID & " ")
                                                    Set rs2 = Connx.Execute("update bill_of_lading set ref_id = " & ObjectID & " where bl_id = " & EXID & " ")
                                                CloseOBJs Connx, rs2
                                            Case 0,11
                                                openConnOcean Connx, "ventas_" & LCase(rs(3))
                                                    Set rs2 = Connx.Execute("update bl_completo set ref_id = " & ObjectID & " where bl_id = " & EXID & " ")
                                                CloseOBJs Connx, rs2
                                        End Select
                                    End Select
                                End If


                            end if
                        Else
                            JavaMsg = "La informacion ya existe " & ObjectID & "-" & RsFilter
                        End If
                  Case 2 'Update
                        CreatedTime = CreatedTime + 1 
						If Not rs.EOF Then 'Si existe el atributo, puede actualizarlo
                            'JavaMsg = "Save Info en Update"
                            
                            
                            if GroupID = 34 then
                            
                                Dim fec1,fec2,fec3,fec4, sql                           

                                'TripRequestDate
                                fec1 = NewServerDate(Request.Form("TripRequestDate"), "", "","TripRequestDate34 Fechas Operativas", CheckNum(Request.Form("OID")), Request.Form("CodeReference"))
        
                                'TripLoadDate
                                fec2 = NewServerDate(Request.Form("TripLoadDate"), Request.Form("TripLoadDate_h"), Request.Form("TripLoadDate_i"), "TripLoadDate", CheckNum(Request.Form("OID")), Request.Form("CodeReference"))
       
                                'TripUnloadPartialDate
                                fec3 = NewServerDate(Request.Form("TripUnloadPartialDate"), Request.Form("TripUnloadPartialDate_h"), Request.Form("TripUnloadPartialDate_i"), "TripUnloadPartialDate", CheckNum(Request.Form("OID")), Request.Form("CodeReference"))
      
                                'TripUnloadDate
                                fec4 = NewServerDate(Request.Form("TripUnloadDate"), Request.Form("TripUnloadDate_h"), Request.Form("TripUnloadDate_i"), "TripUnloadDate", CheckNum(Request.Form("OID")), Request.Form("CodeReference"))

                                sql = "UPDATE BLDetail SET CreatedTime = '" & CreatedTime & "', TripRequestDate = '" & fec1 & "', TripLoadDate = '" & fec2 & "', TripUnloadPartialDate = '" & fec3 & "', TripUnloadDate = '" & fec4 & "' WHERE BLDetailID = " & ObjectID
                                response.write sql & "<br>"
                                Conn.Execute(sql)
                                'result = WsInsertData(sql, "terrestre")

                            else

                                SaveInfo Conn, rs, Action, GroupID, CreatedDate, CreatedTime, BLType

							    Select Case GroupID 'Actualizando datos de la tabla master (Clientes, Shippers, Exporters)
							    Case 3, 4, 11, 20
								    AddressID = SaveMaster (Conn, ObjectID)
							    End Select

                            end if


                        Else
                            JavaMsg = "La informacion no existe"
                        End If
                  Case 3 'Delete
                        If Not rs.EOF Then 'Si existe el atributo, puede borrarlo
							'al Borrarse la CP Master se libera su detalle
							Select Case GroupID
                            Case 1
                                'response.write("select a." & ObjectName & ", b.EXID from " & TableName & " a inner join BLDetail b on a.BLID = b.BLID where a.CreatedDate='" & CreatedDate & "' and a.CreatedTime=" & CreatedTime)
                                Set rs2 = Conn.Execute("select a." & ObjectName & ", b.EXID, b.EXType, EXDBCountry from " & TableName & " a inner join BLDetail b on a.BLID = b.BLID where a.CreatedDate='" & CreatedDate & "' and a.CreatedTime=" & CreatedTime)
                                If Not rs2.EOF Then
                                    encValues = rs2.GetRows
                                    CountListValues = rs2.RecordCount-1
                                End If
                                if CountListValues >= 0 Then
                                For i = 0 To CountListValues
                                    EXType = 0
                                    EXID = CheckNum(encValues(1,i))
                                    EXType = encValues(2,i)
                                    if EXID <> 0 then
                                        Select Case EXType
                                        Case 4, 5, 6
                                            OpenConn2 Connx
                                            'response.write("select routing_seg from routings where id_routing = " & EXID & "")
                                            Set rs3 = Connx.Execute("select routing_seg, routing_adu from routings where id_routing = " & EXID & "")
                                            If Not rs3.EOF Then
                                                if (rs3(0) <> 0 and rs3(1) = 0) then
                                                    'response.write("update routings set no_bl = '' where id_routing in (" & EXID & "," & rs3(0) & ")")
                                                    Connx.Execute("update routings set no_bl = '', activo=true where id_routing in (" & EXID & "," & rs3(0) & ")")
                                                    Connx.Execute("update routings set no_embarque = '' where id_routing in (" & EXID & ")")
                                                elseif (rs3(0) = 0 and rs3(1) <> 0) then
                                                    'response.write("update routings set no_bl = '' where id_routing in (" & EXID & "," & rs3(1) & ")")
                                                    Connx.Execute("update routings set no_bl = '', activo=true where id_routing in (" & EXID & "," & rs3(1) & ")")
                                                    Connx.Execute("update routings set no_embarque = '' where id_routing in (" & EXID & ")")
                                                elseif (rs3(0) <> 0 and rs3(1) <> 0) then
                                                    'response.write("update routings set no_bl = '' where id_routing in (" & EXID & "," & rs3(0) & "," & rs3(1) & ")")
                                                    Connx.Execute("update routings set no_bl = '', activo=true where id_routing in (" & EXID & "," & rs3(0) & "," & rs3(1) & ")")
                                                    Connx.Execute("update routings set no_embarque = '' where id_routing in (" & EXID & ")")
                                                else    
                                                    'response.write("update routings set no_bl = '' where id_routing = " & EXID & "")
                                                    Connx.Execute("update routings set no_bl = '', activo=true, no_embarque = '' where id_routing = " & EXID & "")
                                                end if
                                            End If
                                            CloseOBJ rs3
                                        Case 1,2,12,13
                                            openConnOcean Connx, "ventas_" & LCase(encValues(3,i))
                                            Set rs3 = Connx.Execute("update bill_of_lading set ref_doc = '' where bl_id = " & EXID & " ")
                                            CloseOBJs Connx, rs3
                                        Case 0,11
                                            openConnOcean Connx, "ventas_" & LCase(encValues(3,i))
                                            Set rs3 = Connx.Execute("update bl_completo set ref_doc = '' where bl_id = " & EXID & " ")
                                            CloseOBJs Connx, rs3
                                        End Select
                                    End if
                                    'Desasignando Registros Nuevos, borra el HBLNumber, para obtener el que corresponda cuando vuelva a ser asignado
								    Conn.Execute("update BLDetail set BLID=-1, Pos=0, HBLNumber='--' where BLID=" & ObjectID & " and BLIDTransit=0")
								    'Desasignando Registros en Transito, no borra el HBLNumber porque se debe mantener por la poliza de seguro
								    Conn.Execute("update BLDetail set BLID=-1, Pos=0 where BLID=" & ObjectID & " and BLIDTransit<>0")
                                    Conn.Execute("update BLs set Expired = 1 where BLID = " & encValues(0,i))
                                Next
                                End if
                                'rs.Delete
							Case 23
                                'response.write("update Tracking set Expired = 1 where TrackingID = " & ObjectID)
                                Conn.Execute("update Tracking set Expired = 1 where TrackingID = " & ObjectID)
                                'rs.Delete
                            Case Else
                                CreatedTime = CreatedTime + 1 
                                select case GroupID
		                        case 2
			                        rs("hora_creacion") = CreatedTime 
                                    rs("activo") = 0
                                case 3, 4, 11, 20
                                    rs("hora_creacion") = CreatedTime 
                                    rs("id_estatus") = 0
		                        case else
			                        rs("Expired") = 1
                                    rs("CreatedTime") = CreatedTime
		                        end select                                
                                rs.Update
                            end Select
						    ObjectID = 0
							Name = ""
							Countries = ""
							ProviderID = 0
							TruckNo = ""
							CodProv = ""
                        Else
                            JavaMsg = "La informacion no existe"
                        End If
                  End Select
                  closeOBJ rs
        	End If
		
			if JavaMsg <> "" then
				Response.write "<SCRIPT>alert('" & JavaMsg & "');</SCRIPT>"
			end if
			
			If Action <> 4 then

				select case GroupID

				Case 3, 4, 11, 20
					if AddressID = 0 then
						AddressID=CheckNum(Request("AID"))
						if AddressID = 0 then
							AddressID = CheckNum(Request("AddressID"))
						end if
					end if
					SQLQuery = QuerySelect & " and c." & ObjectName & "=" & ObjectID & " and d.id_direccion=" & AddressID
				'case 14
				'	SQLQuery = QuerySelect & ObjectID & " group by c.LtEndorseDate"
				case 2, 9, 14, 15, 34, 35, 37
					SQLQuery = QuerySelect & TableName & " where " & ObjectName & "=" & ObjectID
				case else
					SQLQuery = QuerySelect & TableName & " where " & ObjectName & "=" & ObjectID & " and " & WhereSQL & CreatedTime & " and Expired = 0"
				end select
				
				'response.write SQLQuery & "<br>"
				'SQLQuery = "select * from BLDetail where BLDetailID=0"
                
				Set rs = Conn.Execute(SQLQuery)
				If Not rs.EOF Then
        			aTableValues = rs.GetRows
        			CountTableValues = rs.RecordCount
    			End If
    			closeOBJs rs, Conn
			
				Select Case GroupID
				   Case 1 'Cartas Porte %>
					<!--#include file=BLs.asp--> 
				<% Case 2 'Remitentes %>
					<!--#include file=Senders.asp--> 
				<% Case 3, 4, 11, 20 'Destinatarios - Consigneers, Shippers %>
					<!--#include file=Master.asp-->
				<% 'Case 3, 20 'Clients.asp, Shippers.asp %>
				<% 'Case 4 'Consignatario Consigners.asp%>
				<% Case 5 'Transportes - Pilotos %>
					<!--#include file=Pilots.asp--> 
				<% Case 6 'Transporte - Cabezales  %>
					<!--#include file=Trucks.asp--> 	
				<% Case 7 'Transporte - Proveedor %>
					<!--#include file=Providers.asp--> 
				<% Case 8 'Aduana %>
					<!--#include file=Brokers.asp--> 
				<% Case 9 'Productos %>
					<!--#include file=Commodities.asp--> 
				<% Case 10 'Agentes %>
					<!--#include file=Letters.asp--> 
				<% Case 12 'Itinerario %>
					<!--#include file=RepItinerary.asp--> 
				<% Case 13 'Fianzas %>
					<!--#include file=RepBail.asp--> 
				<% Case 14, 15 'Carta Endoso Detalle Bls %>
					<!--#include file=RepEndorse.asp--> 
				<% Case 16 'Carta Aceptacion %>
					<!--#include file=RepAccept.asp--> 
				<% Case 17 'Carta Entrega %>
					<!--#include file=RepDelivery.asp--> 
				<% Case 18 'Carta Recoleccion %>
					<!--#include file=RepCollect.asp--> 
				<% Case 19 'Solicitud Movimiento %>
					<!--#include file=RequestMov.asp-->
				<% Case 21 'Bodega / WareHouse %>
					<!--#include file=Warehouses.asp-->	 
				<% Case 22 'Agrupacion de BLs %>
					<!--#include file=BLGroups.asp-->
				<% Case 23 'Rastreo / Tracking %>
					<!--#include file=Tracking.asp-->		 
				<% Case 25 'DTI %>
					<!--#include file=DTI.asp-->		  
				<% Case 26 'Plantillas para DTI %>
					<!--#include file=DTITemplates.asp-->		 
				<% Case 27, 28, 33, 37 'Carga Transito o General para Itinerario %>
					<!--#include file=ItineraryAdds.asp-->		 
				<% Case 32 'Carga Transito o General para Itinerario %>
					<!--#include file=Marchamos.asp-->
                <% Case 34 'fechas operativas Trip Area 2020-09-08 %>
					<!--#include file=fechas_operativas.asp-->	
                <% Case 35 'fechas aduanas Trip Area 2020-09-08 %>
                    <!--#include file=fechas_aduanas.asp-->	  
                <% Case 36 'desglose 2021-02-19 %>
                    <!--#include file=Product_Split.asp-->	                               
   				<% end Select
				Set aTableValues = Nothing
			Else
				Select Case GroupID
				Case 2
					SearchSimilars "agente", request.Form("Name"), GroupID, " ", SearchOption
				Case 5, 7, 8, 10
					SearchSimilars "Name", request.Form("Name"), GroupID, " ", SearchOption
				Case 3, 4, 11, 20
					SearchSimilars "nombre_cliente", request.Form("Name"), GroupID, " ", SearchOption
				Case 6
					SearchSimilars "TruckNo", request.Form("TruckNo"), GroupID, " ", SearchOption
				Case 9
					SearchSimilars "namees", request.Form("namees"), GroupID, " ", SearchOption
				End Select %>
				<!--#include file=Similars.asp--> 
			<% End If
    Elseif GroupID = 35 Then
        
    End If 
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>
