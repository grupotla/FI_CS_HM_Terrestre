<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
Server.ScriptTimeout = 1000
if Session("OperatorID") <> 0 then
Checking "0|1|2"
Dim GroupID, HTMLCode, HTMLCode2, HTMLTitle, HTMLTitle2, Table, OrderName, QuerySelect, DateFrom, DateTo, CD, MoreOptions, SUBQUERY
Dim Option1, Option2, Option3, Option4, Option5, Option6, Option7, Option8, Option9, Option10, Option11, ReportTitle
Dim Name, ProviderID, Attn, CurrencyCode, Val, TruckNo, Countries, TruckType, DTI, ReportType, ShipperName
Dim BLNumber, BLType, Week, Week2, BrokerID, PilotID, TruckID, CountryDep, CountryDes, HBL, MBL, CNT, DownFile, IT
Dim JavaMsg, i, j, elements, PageCount, AbsolutePage, HTMLHidden, Status, WarehouseID, BagValue, ItineraryType, rs1, Correos, CodeReference

GroupID = CheckNum(Request("GID"))
	
if GroupID >= 1 and GroupID <=37 then
	AbsolutePage = CheckNum(Request.Form("P"))
	if AbsolutePage = 0 then
		 AbsolutePage = 1
	end if
	elements = 5
	PageCount = 0
    Select case GroupID
	case 1, 12, 13, 14, 15, 16, 25, 29, 34, 35 'BL y Reportes
		 BLType = CheckNum(Request.Form("BLType"))
		 ReportType = CheckNum(Request.Form("ReportType"))
		 BLNumber = Request.Form("BLNumber")
		 Week = CheckNum(Request.Form("Week"))
		 CountryDep = Request.Form("CountryDep")
		 CountryDes = Request.Form("CountryDes")
		 BrokerID = CheckNum(Request.Form("BrokerID"))
		 PilotID = CheckNum(Request.Form("PilotID"))
		 TruckID = CheckNum(Request.Form("TruckID"))
		 DTI = Request.Form("DTI")
		 HBL = Request.Form("HBL")
		 CodeReference = Request.Form("CodeReference")

		 select Case GroupID
		 case 14
			 OrderName = " order by BLID desc, ClientsID, AgentsID"
			 QuerySelect = ""'como utiliza SQL UNION comprende 2 Selects que se unifican mas abajo
			 HTMLTitle = "<tr><td class=titlelist><b>Fecha</td><td class=titlelist><b>Pais</td><td class=titlelist><b>No. de Carta Porte Individual</td><td class=titlelist><b>Consignatario</td><td class=titlelist><b>Status</td></tr>"
		 case 16
			 if BLType >= 0 then
				OrderName = " order by a.BLID Desc"
				QuerySelect = "select a.BLID, a.CreatedTime, a.CreatedDate, a.Countries, a.BLNumber, a.Expired, a.LtAcceptDate from BLs a"
			 else
				OrderName = " order by a.BLGroupID Desc"
				QuerySelect = "select a.BLGroupID, a.CreatedTime, a.CreatedDate, a.Countries, a.BLNumber, a.Expired, a.LtAcceptDate from BLGroups a"
			 end if
			 HTMLTitle = "<tr><td class=titlelist><b>Fecha</td><td class=titlelist><b>Paises</td><td class=titlelist width=50><b>CP</td><td class=titlelist width=50><b>DOCS</td><td class=titlelist width=50><b>MF</td><td class=titlelist width=50><b>IT</td><td class=titlelist><b>No. de Carta Aceptaci&oacute;n</td><td class=titlelist><b>Status</td></tr>"
		 case 25
			 OrderName = " order by a.BLID Desc"
			 QuerySelect = "select a.BLID, a.CreatedTime, a.CreatedDate, a.Countries, a.BLNumber, a.Expired, a.DTI from BLs a"
			 HTMLTitle = "<tr><td class=titlelist><b>Fecha</td><td class=titlelist><b>Paises</td><td class=titlelist><b>No. de Carta Porte</td><td class=titlelist><b>Status</td></tr>"
		 
         
        case 34, 35
			 
             OrderName = " order by BLDetailID Desc"
			 
             'QuerySelect = "select a.BLID, a.CreatedTime, a.CreatedDate, a.Countries, a.BLNumber, a.Expired FROM "

             'QuerySelect = "select a.HBLNumber, a.Clients, a.Weights, a.Volumes, a.Shippers, a.CountriesFinalDes, a.BLs, a.MBLs, a.ClientsID, a.EXID from BLDetail a"
	       
             QuerySelect = "select BLDetailID, CreatedTime, CreatedDate, Countries, HBLNumber, CodeReference from BLDetail "  

			 HTMLTitle = "<tr><td class=titlelist><b>Fecha</td><td class=titlelist><b>Paises</td><td class=titlelist><b>No. de Carta Porte</td><td class=titlelist><b>Codigo Referencia</td></tr>"

         case else
			 OrderName = " order by a.BLID Desc"
			 QuerySelect = "select a.BLID, a.CreatedTime, a.CreatedDate, a.Countries, a.BLNumber, a.Expired"
			 HTMLTitle = "<tr><td class=titlelist><b>Fecha</td><td class=titlelist><b>Paises</td><td class=titlelist><b>No. de Carta Porte</td><td class=titlelist><b>Status????</td></tr>"
			 select case GroupID
			 case 1, 29
				 QuerySelect = "select a.BLID, a.CreatedTime, a.CreatedDate, a.Countries, a.BLNumber, a.Closed, a.LtAcceptDate"
			 case 12
				 QuerySelect = QuerySelect & ", a.BLArrivalDate"
			 case 13
				 QuerySelect = QuerySelect & ", a.Bail"
			 end select
			 QuerySelect = QuerySelect & " from BLs a"
		 end select
					 
		 select case GroupID
		 case 12, 29
			'Option1 = " (a.Countries in " & Session("Countries") & " or a.CountryDes in " & Session("Countries") & ") and a.BLType=" & BLType & " "				
			Option1 = " a.BLType=" & BLType & " "				
		 case 14, 15
			'Option1 = " a.CountryDes in " & Session("Countries") & " and a.LtAcceptDate<>'' and a.BLType=" & BLType & " "
			'Option1 = " a.CountryDes in " & Session("Countries") & " and b.BLIDTransit=0 and a.BLType=" & BLType & " "
			'Option1 = " a.CountryDes in " & Session("Countries") & " and a.BLArrivalDate<>'' and a.BLType=" & BLType & " "
			Option1 = "(b.CountriesFinalDes in " & Session("Countries") & " or b.Countries in " & Session("Countries") & " or b.CountryOrigen in " & Session("Countries") & ") and a.BLType=" & BLType & " "
			'Option1 = " a.BLType=" & BLType & " "
		 case 16
			if BLType >= 0 then
				'Option1 = " a.CountryDes in " & Session("Countries") & " and Closed=1 and a.Bail<>'' and a.BLType=" & BLType & " "
				'Option1 = " a.CountryDes in " & Session("Countries") & " and a.Closed=1 and a.BLType=" & BLType & " "
				Option1 = " a.Closed=1 and a.BLType=" & BLType & " "
			else
				'Option1 = " a.CountryDes in " & Session("Countries") & " "
			end if

         case 34, 35 

            Option1 = " HBLNumber like '%" & BLNumber & "%' "

            IF CodeReference <> "" then           
                Option1 = Option1 & " AND CodeReference like '%" & CodeReference & "%' "
            end if

            Option1 = Option1 & " AND ClientsID IN (" & GroupList() & ")" '2020-09-11  

		 case else
			Select Case Session("Login") 
            Case "sv-intermodal", "cesar-sanchez"
                Option1 = " a.BLType=" & BLType & " "
            Case "vanessa-cruz"
                Option1 = " (a.CountryDep = 'HN' or a.CountryDes = 'HN' or a.CountryDep = 'HNLTF' or a.CountryDes = 'HNLTF') and a.BLType=" & BLType & " "
            Case "cr-intermodaltransport", "margarita-rodriguez"
                Option1 = " (a.CountryDep = 'CR' or a.CountryDes = 'CR' or a.CountryDep = 'CRLTF' or a.CountryDes = 'CRLTF') and a.BLType=" & BLType & " "
            Case Else
                Option1 = " a.Countries in " & Session("Countries") & " and a.BLType=" & BLType & " and a.Expired = 0"
            End Select
		end select


		 if BLNumber <> "" and GroupID <> 34 and GroupID <> 35 then
				if GroupID<>14 then
					Option2 = " a.BLNumber like '%" & BLNumber & "%' "
				else
					'Option2 = " b.HBLNumber like '%" & BLNumber & "%' or a.BLNumber like '%" & BLNumber & "%' "
					Option2 = " b.HBLNumber like '%" & BLNumber & "%' "
				end if
		 end if
		 if Week <> 0 then
				Option3 = " a.Week=" & Week & " "
		 end if
		 if CountryDep <> "" then
				Option4 = " a.CountryDep='" & CountryDep & "' "
		 end if
		 if CountryDes <> "" then
				Option5 = " a.CountryDes='" & CountryDes & "' "
		 end if
		 if BrokerID <> 0 then
				Option6 = " a.BrokerID=" & BrokerID & " "
		 end if
		 if PilotID <> 0 then
				Option7 = " a.PilotID=" & PilotID & " "
		 end if
		 if TruckID <> 0 then
				Option8 = " a.TruckID=" & TruckID & " "
		 end if
		 if DTI <> "" then
				Option9 = " a.DTI=" & DTI & " "
		 end if
		 if HBL <> "" then
		 		Option9 = " (b.BLs like '%" & HBL & "%' or b.MBLs like '%" & HBL & "%') "
		 end if
		
		 HTMLHidden = HTMLHidden & "<INPUT name='BLType' type=hidden value='" & BLType & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='BLNumber' type=hidden value='" & BLNumber & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='Week' type=hidden value='" & Week & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='CountryDep' type=hidden value='" & CountryDep & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='CountryDes' type=hidden value='" & CountryDes & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='BrokerID' type=hidden value='" & BrokerID & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='PilotID' type=hidden value='" & PilotID & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='TruckID' type=hidden value='" & TruckID & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='DTI' type=hidden value='" & DTI & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='HBL' type=hidden value='" & HBL & "'>"


	case 37

		elements=7
		 
        OrderName = " order by tch_fecha DESC "
        
        QuerySelect = GetSQLSearch (GroupID)

		HTMLTitle = "<tr><td class=titlelist><b>Fecha</td><td class=titlelist><b>ShipmentIDNumber</td><td class=titlelist><b>Carta Porte</td><td class=titlelist><b>TransactionSetPurposeCode</td><td class=titlelist><b>MessageID</td><td class=titlelist><b>Estado</td></tr>"

		Option1 = "" '" a.tch_estado = 1 "

	case 2

		 elements=6
		 OrderName = " order by a.agente"
		 QuerySelect = GetSQLSearch (GroupID)
			 HTMLTitle = "<tr><td class=titlelist><b>Fecha</td><td class=titlelist><b>Codigo</td><td class=titlelist><b>Nombre</td><td class=titlelist><b>Contacto</td><td class=titlelist><b>Status</td></tr>"
		 Name = Request.Form("Name")
		 if Name <> "" then
				Option1 = " a.agente ilike '%" & Name & "%' "
		 end if
		 HTMLHidden = HTMLHidden & "<INPUT name='Name' type=hidden value='" & Name & "'>"

	case 7, 8, 21
		 OrderName = " order by a.Countries, a.Name"
		 QuerySelect = GetSQLSearch (GroupID)
		 Select Case GroupID
		 case 2
			 HTMLTitle = "<tr><td class=titlelist><b>Fecha</td><td class=titlelist><b>Paises</td><td class=titlelist><b>Nombre Remitente</td><td class=titlelist><b>Status</td></tr>"
		 case 7
			 HTMLTitle = "<tr><td class=titlelist><b>Fecha</td><td class=titlelist><b>Paises</td><td class=titlelist><b>Nombre Proveedor</td><td class=titlelist><b>Status</td></tr>"
		 case 8
			 HTMLTitle = "<tr><td class=titlelist><b>Fecha</td><td class=titlelist><b>Paises</td><td class=titlelist><b>Nombre Aduana</td><td class=titlelist><b>Status</td></tr>"
		 case 21
			 HTMLTitle = "<tr><td class=titlelist><b>Fecha</td><td class=titlelist><b>Paises</td><td class=titlelist><b>Nombre Bodega</td><td class=titlelist><b>Status</td></tr>"
		 end select
		 Name = Request.Form("Name")
		 Attn = Request.Form("Attn")
		 Countries = Request.Form("Countries")
		 'Option1 = " a.Countries in " & Session("Countries") & " "
		 if Name <> "" then
				Option1 = " a.Name like '%" & Name & "%' "
		 end if
		 if Attn <> "" then
				Option2 = " a.Attn like '%" & Attn & "%' "
		 end if
		 if Countries <> "" then
				Option3 = " a.Countries like '%" & Countries & "%' "
		 end if
		 HTMLHidden = HTMLHidden & "<INPUT name='Name' type=hidden value='" & Name & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='Attn' type=hidden value='" & Attn & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='Countries' type=hidden value='" & Countries & "'>"
	case 3
		 elements = 6
		 OrderName = " order by p.codigo, a.nombre_cliente"
		 QuerySelect = GetSQLSearch (GroupID)
		 HTMLTitle = "<tr><td class=titlelist><b>Fecha</td><td class=titlelist><b>Pais</td><td class=titlelist><b>Codigo</td><td class=titlelist><b>Embarcador</td><td class=titlelist><b>Status</td></tr>"
		 Name = Request.Form("Name")
		 Option1 = " a.id_cliente = d.id_cliente " & _
						"and d.id_nivel_geografico = n.id_nivel " & _
						"and n.id_pais = p.codigo " & _
						"and a.es_shipper = true "' & _
						'"and p.codigo in " & Session("Countries") & " "
		 if Name <> "" then
				Option2 = " a.nombre_cliente ilike '%" & Name & "%' "
		 end if
		 HTMLHidden = HTMLHidden & "<INPUT name='Name' type=hidden value='" & Name & "'>"
	case 4, 11	
		 elements = 6
		 OrderName = " order by p.codigo, a.nombre_cliente"
		 QuerySelect = GetSQLSearch (GroupID)
		 HTMLTitle = "<tr><td class=titlelist><b>Fecha</td><td class=titlelist><b>Pais</td><td class=titlelist><b>Codigo</td><td class=titlelist><b>Destinatario</td><td class=titlelist><b>Status</td></tr>"
		 Name = Request.Form("Name")
		 Option1 = " a.id_cliente = d.id_cliente " & _
						"and d.id_nivel_geografico = n.id_nivel " & _
						"and n.id_pais = p.codigo "' & _
						'"and a.es_consigneer = true " & _
						'"and p.codigo in " & Session("Countries") & " "
		 if Name <> "" then
				Option2 = " a.nombre_cliente ilike '%" & Name & "%' "
		 end if
		 HTMLHidden = HTMLHidden & "<INPUT name='Name' type=hidden value='" & Name & "'>"
	case 5, 6
		 QuerySelect = GetSQLSearch (GroupID)
		 Select Case GroupID
		 case 5
			 HTMLTitle = "<tr><td class=titlelist><b>Fecha</td><td class=titlelist><b>Paises</td><td class=titlelist><b>Nombre Piloto</td><td class=titlelist><b>Status</td></tr>"
			 OrderName = " order by a.Countries, a.Name"
		 case 6
			 HTMLTitle = "<tr><td class=titlelist><b>Fecha</td><td class=titlelist><b>Paises</td><td class=titlelist><b>Numero Cabezal</td><td class=titlelist><b>Status</td></tr>"
			 OrderName = " order by a.Countries, a.TruckNo"
		 end select
		 Countries = Request.Form("Countries")
		 ProviderID = CheckNum(Request.Form("ProviderID"))
		 Name = Request.Form("Name")
		 TruckNo = Request.Form("TruckNo")
		 TruckType = Request.Form("TruckType")
		 
		 if Countries <> "" then
				Option1 = " a.Countries like '%" & Countries & "%' "
		 end if
		 if ProviderID <> 0 then
				Option2 = " a.ProviderID =" & ProviderID & " "
		 end if
		 if Name <> "" then
				Option3 = " a.Name like '%" & Name & "%' "
		 end if
		 if TruckNo <> "" then
				Option4 = " a.TruckNo like '%" & TruckNo & "%' "
		 end if
		 if TruckType <> "" then
				Option5 = " a.TruckType = " & TruckType & " "
		 end if
		 HTMLHidden = HTMLHidden & "<INPUT name='ProviderID' type=hidden value='" & ProviderID & "'>"		 
		 HTMLHidden = HTMLHidden & "<INPUT name='Name' type=hidden value='" & Name & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='TruckNo' type=hidden value='" & TruckNo & "'>"		 
	case 9
		 OrderName = " order by a.NameES"
		 QuerySelect = 	GetSQLSearch (GroupID)
		 HTMLTitle = "<tr><td class=titlelist><b>Fecha</td><td class=titlelist><b>C&oacute;digo SCR</td><td class=titlelist><b>Producto</td><td class=titlelist><b>Status</td></tr>"
		 Name = Request.Form("Name")
		 if Name <> "" then
				Option1 = " a.NameES ilike '%" & Name & "%' "
		 end if
		 HTMLHidden = HTMLHidden & "<INPUT name='Name' type=hidden value='" & Name & "'>"
	case 10
		 OrderName = " order by a.LetterID"
		 QuerySelect = 	"select a.LetterID, a.CreatedTime, a.CreatedDate, a.Countries, a.Name, a.Expired from Letters a"
		 HTMLTitle = "<tr><td class=titlelist><b>Fecha</td><td class=titlelist><b>Paises</td><td class=titlelist><b>Nombre</td><td class=titlelist><b>Status</td></tr>"
		 Option1 = " a.LetterID>0"
	case 17
		 BLType = CheckNum(Request.Form("BLType"))
         ReportType = CheckNum(Request.Form("ReportType"))
		 Val = Request.Form("Yr")
		 Week = CheckNum(Request.Form("Week"))
         Week2 = CheckNum(Request.Form("Week2"))
         DownFile = CheckNum(Request.Form("DownFile"))
         Select Case ReportType
         Case 0
            if DownFile=1 then
                'Para guardar los resultados en excel
                'Response.Buffer = False
                Server.ScriptTimeout = 5000
                Response.ContentType = "application/vnd.ms-excel"
                Response.AddHeader "Content-Disposition", "filename=CargaCAM.xls"
             end if
             OrderName = " group by a.HBLNumber order by a.Week, a.CountriesFinalDes"
             if BLType <> 2 then
             QuerySelect = "select a.HBLNumber, a.Clients, a.Weights, a.Volumes, a.Shippers, a.CountriesFinalDes, a.BLs, a.MBLs, a.ClientsID, a.EXID from BLDetail a"
	         HTMLTitle = "<tr><td class=label colspan=2><b>Reporte de Carga CAM</b></td><td class=label align=right colspan=6><input class=label onclick='javascript:document.forma.DownFile.value=1;document.forma.submit();' type=button value='Bajar a Excel'></td></tr>" & _
                "<tr><td class=titlelist><b>CP Hija</b></td><td class=titlelist><b>Cliente</b></td><td class=titlelist><b>Vendedor</b></td><td class=titlelist><b>Peso</b></td><td class=titlelist><b>Volumen</b></td><td class=titlelist><b>Agente</b></td><td class=titlelist><b>Destino Final<b></td><td class=titlelist colspan=2><b>BL/RO<b></td><td class=titlelist><b>MBL/RO<b></td class=titlelist><td class=titlelist><b>No. Factura</b></td><td class=titlelist><b>Tipo Moneda</b></td><td class=titlelist><b>Monto</b></td><td class=titlelist colspan='2'><b>Rubro</b></td>"
             else
             QuerySelect = "select a.HBLNumber, a.Clients, a.Weights, a.Volumes, a.Shippers, a.CountriesFinalDes, a.BLs, a.MBLs, a.ClientsID, a.EXID, c.Name, d.TruckNo from BLDetail a INNER JOIN BLs b ON b.BLID = a.BLID INNER JOIN Pilots c ON c.PilotID = b.PilotID INNER JOIN Trucks d ON d.TruckID = b.TruckID"
	         HTMLTitle = "<tr><td class=label colspan=2><b>Reporte de Carga CAM</b></td><td class=label align=right colspan=6><input class=label onclick='javascript:document.forma.DownFile.value=1;document.forma.submit();' type=button value='Bajar a Excel'></td></tr>" & _
                "<tr><td class=titlelist><b>CP Hija</b></td><td class=titlelist><b>Cliente</b></td><td class=titlelist><b>Piloto</b></td><td class=titlelist><b>Placa Camion</b></td><td class=titlelist><b>Vendedor</b></td><td class=titlelist><b>Peso</b></td><td class=titlelist><b>Volumen</b></td><td class=titlelist><b>Agente</b></td><td class=titlelist><b>Destino Final<b></td><td class=titlelist><b>BL/RO<b></td><td class=titlelist><b>MBL/RO<b></td class=titlelist><td class=titlelist><b>No. Factura</b></td><td class=titlelist><b>Tipo Moneda</b></td><td class=titlelist colspan='2'><b>Monto, </b><b>Rubro</b></td>" 'Imprime Piloto y Placa camión en Reporte carga CAM si es carga local
             end if
    		 Option1 = " a.BLType=" & BLType & " and ((a.Countries in " & Session("Countries") & " and (instr(a.Countries,Mid(HBLNumber,2,2)))) or a.CountriesFinalDes IN " & Session("Countries") & ") and a.HBLNumber <> '--' "
         Case 1
             OrderName = ""
    	     QuerySelect = "select a.BLID, a.BLNumber, a.BLRealExitDate, a.BLExitDate, a.BLArrivalDate, a.BLEstArrivalDate, a.Comment, a.CountryDep, a.CountryDes, a.ChargeType from BLs a"
	         HTMLTitle = "<tr><td class=label colspan=10><b>Reporte de Tiempos de Rutas</b></td></tr>" & _
                "<tr><td class=titlelist><b>Ruta</b></td><td class=titlelist><b>Paises</b></td><td class=titlelist><b>Carta&nbsp;Porte</b></td><td class=titlelist><b>Salida Estimada</b></td><td class=titlelist><b>Salida</b></td><td class=titlelist><b>Cumplio<b></td><td class=titlelist><b>Llegada Estimada<b></td><td class=titlelist><b>Llegada<b></td><td class=titlelist><b>Cumplio<b></td><td class=titlelist><b>Observaciones</b></td></tr>"
    		 Option1 = " a.BLType=" & BLType & " "
            HTMLTitle2 = "<tr><td class=label colspan=10><b>Reporte de Tiempos que no se encuentran en Rutas Establecidas</b></td></tr>" & _
                "<tr><td class=titlelist><b>Paises</b></td><td class=titlelist><b>Carta&nbsp;Porte</b></td><td class=titlelist><b>Salida Estimada</b></td><td class=titlelist><b>Salida</b></td><td class=titlelist><b>Cumplio<b></td><td class=titlelist><b>Llegada Estimada<b></td><td class=titlelist><b>Llegada<b></td><td class=titlelist><b>Cumplio<b></td><td class=titlelist><b>Observaciones</b></td></tr>"            
		 Case 2,4
             if DownFile=1 then
                'Para guardar los resultados en excel
                'Response.Buffer = False
                Server.ScriptTimeout = 3600
                Response.ContentType = "application/vnd.ms-excel"
                Response.AddHeader "Content-Disposition", "filename=PorcentajeUtilizacionRutaCAM.xls"
             end if
             OrderName = ""
             'se separan totales cuando:
             '1. No esta Ruteado y no es Coloader
             '2. No esta ruteado y si es Coloader 
    	     '3. Cuando esta Ruteado

             if ReportType=2 then
                 'Query para las cargas generadas tanto en origen como transito
                 'En este query las posiciones 6,7 del query toman el total de peso y volumen directo del BL
                 QuerySelect = "select a.TruckID, '', round(coalesce(b.Weight,0)+coalesce(i.Weight,0),2) as Peso_Total, round(coalesce(b.Volume,0)+coalesce(i.Volume,0),2) as CBMS_Total, 0, 0, round(sum(a.TotWeight),2) as Peso_Utilizado, round(sum(a.TotVolume),2) as CBMS_Utilizados, " & _
                 "coalesce((select round(sum(Weights),2) from BLDetail c where c.BLID in (#SUBQUERY# #SUBQUERY2#) and c.ExType not in (4,5,6,7) and c.ColoadersID=0), 0) as Peso_Agente, " & _
                 "coalesce((select round(sum(Volumes),2) from BLDetail d where d.BLID in (#SUBQUERY# #SUBQUERY2#) and d.ExType not in (4,5,6,7) and d.ColoadersID=0), 0) as CBMS_Agente, " & _
                 "coalesce((select round(sum(Weights),2) from BLDetail e where e.BLID in (#SUBQUERY# #SUBQUERY2#) and e.ExType not in (4,5,6,7) and e.ColoadersID<>0), 0) as Peso_Coloader, " & _
                 "coalesce((select round(sum(Volumes),2) from BLDetail f where f.BLID in (#SUBQUERY# #SUBQUERY2#) and f.ExType not in (4,5,6,7) and f.ColoadersID<>0), 0) as CBMS_Coloader, " & _
                 "coalesce((select round(sum(Weights),2) from BLDetail g where g.BLID in (#SUBQUERY# #SUBQUERY2#) and g.ExType in (4,5,6,7)), 0) as Peso_RO_Aimar, " & _
                 "coalesce((select round(sum(Volumes),2) from BLDetail h where h.BLID in (#SUBQUERY# #SUBQUERY2#) and h.ExType in (4,5,6,7)), 0) as CBMS_RO_Aimar, " & _
                 "a.BLExitDate " & _
                 "from ((BLs a left join Trucks b on a.TruckID=b.TruckID and b.TruckType in (0,2)) left join Trucks i on a.Container=i.TruckID and i.TruckType in (1))"
                 ReportTitle = ""
             else
                 'query solo para las cargas generadas en origen, no toma en transito
                 'En este query las posiciones 6,7 son la suma respectiva de peso y volumen solo de las cargas generadas en origen
                 QuerySelect = "select a.TruckID, '', round(coalesce(b.Weight,0)+coalesce(i.Weight,0),2), round(coalesce(b.Volume,0)+coalesce(i.Volume,0),2), 0, 0, " & _
                 "round(" & _
                 "coalesce((select round(sum(Weights),2) from BLDetail wc where wc.BLID in (#SUBQUERY# #SUBQUERY2#) and wc.ExType not in (4,5,6,7) and wc.Countries=SubStr(wc.HBLNumber,2,2) and wc.ColoadersID=0), 0)+" & _
                 "coalesce((select round(sum(Weights),2) from BLDetail we where we.BLID in (#SUBQUERY# #SUBQUERY2#) and we.ExType not in (4,5,6,7) and we.Countries=SubStr(we.HBLNumber,2,2) and we.ColoadersID<>0), 0)+" & _
                 "coalesce((select round(sum(Weights),2) from BLDetail wg where wg.BLID in (#SUBQUERY# #SUBQUERY2#) and wg.ExType in (4,5,6,7) and wg.Countries=SubStr(wg.HBLNumber,2,2)), 0)" & _
                 ",2), " & _
                 "round(" & _
                 "coalesce((select round(sum(Volumes),2) from BLDetail vd where vd.BLID in (#SUBQUERY# #SUBQUERY2#) and vd.ExType not in (4,5,6,7) and vd.Countries=SubStr(vd.HBLNumber,2,2) and vd.ColoadersID=0), 0)+" & _
                 "coalesce((select round(sum(Volumes),2) from BLDetail vf where vf.BLID in (#SUBQUERY# #SUBQUERY2#) and vf.ExType not in (4,5,6,7) and vf.Countries=SubStr(vf.HBLNumber,2,2) and vf.ColoadersID<>0), 0)+" & _
                 "coalesce((select round(sum(Volumes),2) from BLDetail vh where vh.BLID in (#SUBQUERY# #SUBQUERY2#) and vh.ExType in (4,5,6,7) and vh.Countries=SubStr(vh.HBLNumber,2,2)), 0)" & _
                 ",2), " & _
                 "coalesce((select round(sum(Weights),2) from BLDetail c where c.BLID in (#SUBQUERY# #SUBQUERY2#) and c.ExType not in (4,5,6,7) and c.Countries=SubStr(c.HBLNumber,2,2) and c.ColoadersID=0), 0), " & _
                 "coalesce((select round(sum(Volumes),2) from BLDetail d where d.BLID in (#SUBQUERY# #SUBQUERY2#) and d.ExType not in (4,5,6,7) and d.Countries=SubStr(d.HBLNumber,2,2) and d.ColoadersID=0), 0), " & _
                 "coalesce((select round(sum(Weights),2) from BLDetail e where e.BLID in (#SUBQUERY# #SUBQUERY2#) and e.ExType not in (4,5,6,7) and e.Countries=SubStr(e.HBLNumber,2,2) and e.ColoadersID<>0), 0), " & _
                 "coalesce((select round(sum(Volumes),2) from BLDetail f where f.BLID in (#SUBQUERY# #SUBQUERY2#) and f.ExType not in (4,5,6,7) and f.Countries=SubStr(f.HBLNumber,2,2) and f.ColoadersID<>0), 0), " & _
                 "coalesce((select round(sum(Weights),2) from BLDetail g where g.BLID in (#SUBQUERY# #SUBQUERY2#) and g.ExType in (4,5,6,7) and g.Countries=SubStr(g.HBLNumber,2,2)), 0), " & _
                 "coalesce((select round(sum(Volumes),2) from BLDetail h where h.BLID in (#SUBQUERY# #SUBQUERY2#) and h.ExType in (4,5,6,7) and h.Countries=SubStr(h.HBLNumber,2,2)), 0), " & _
                 "BLExitDate " & _
                 "from ((BLs a left join Trucks b on a.TruckID=b.TruckID and b.TruckType in (0,2)) left join Trucks i on a.Container=i.TruckID and i.TruckType in (1))"
                 ReportTitle = " Solo Origenes"
             end if
             
             HTMLTitle = "<tr><td class=label colspan=2><b>Reporte de Carga CAM " & ReportTitle & "</b></td><td class=label align=right colspan=15><input class=label onclick='javascript:document.forma.DownFile.value=1;document.forma.submit();' type=button value='Bajar a Excel'></td></tr>" & _
                "<tr><td class=label colspan=10><b>Porcentaje de Utilizacion de Rutas" & ReportTitle & "</b></td></tr>" & _
                "<tr><td class=titlelist colspan=4><b>Detalle por Rutas</b></td><td class=titlelist colspan=2><b>Capacidad de Equipo</b></td><td class=titlelist colspan=2><b>Dato Real Transportado</b></td><td class=titlelist colspan=3><b>% de Utilizacion<b></td><td class=titlelist colspan=3><b>Agente<b></td><td class=titlelist colspan=3><b>Coloader<b></td><td class=titlelist colspan=3><b>RO Aimar<b></td></tr>" & _
                "<tr><td class=titlelist><b>Ruta</b></td><td class=titlelist><b>Paises</b></td><td class=titlelist><b>Carta&nbsp;Porte</b></td><td class=titlelist><b>Fecha ETD</b></td><td class=titlelist><b>Peso</b></td><td class=titlelist><b>CBM</b></td><td class=titlelist><b>Peso<b></td><td class=titlelist><b>CBM<b></td><td class=titlelist><b>Peso %<b></td><td class=titlelist><b>CBM %<b></td><td class=titlelist><b>Disponible %<b></td><td class=titlelist><b>Peso<b></td><td class=titlelist><b>CBM<b></td><td class=titlelist><b>% Utilizacion<b></td><td class=titlelist><b>Peso<b></td><td class=titlelist><b>CBM<b></td><td class=titlelist><b>% Utilizacion<b></td><td class=titlelist><b>Peso<b></td><td class=titlelist><b>CBM<b></td><td class=titlelist><b>% Utilizacion<b></td></tr>"
    		 Option1 = " a.BLType=" & BLType & " "
             SUBQUERY = "select BLID from BLs aa, Trucks b where aa.TruckID=b.TruckID and aa.TruckID=a.TruckID and " & Option1
		 Case 3
             OrderName = "order by b.Login, a.BLNumber, c.CreatedDate, c.CreatedTime, a.CountryDep, a.CountryDes"
             QuerySelect = "Select b.Login, a.BLNumber, c.CreatedDate, c.CreatedTime, a.CountryDep, a.CountryDes from Tracking c, Operators b, BLs a "
	         HTMLTitle = "<tr><td class=label colspan=10><b>Reporte de Status Ingresados por usuario por CP</b></td></tr>" & _
                "<tr><td class=titlelist><b>Usuario</b></td><td class=titlelist><b>Carta&nbsp;Porte</b></td><td class=titlelist><b>Fecha</b></td><td class=titlelist><b>Origen</b></td><td class=titlelist><b>Destino</b></td></tr>"
    		 Option1 = " c.OperatorID=b.OperatorID and c.BLID=a.BLID and a.BLType=" & BLType & " "
		 End Select

         'Si solo viene una semana indicada, en el rango la ultima semana se coloca el mismo valor para que el query funcione
         'con una semana o rango de semanas
		 if Week2=0 then
            Week2 = Week
         end if
         
         if Week <> 0 then
				Option2 = " a.Week>=" & Week & " "
                SUBQUERY = SUBQUERY & " and aa.Week>=" & Week & " "
		 end if

         if Week2 <> 0 then
				Option3 = " a.Week<=" & Week2 & " "
                SUBQUERY = SUBQUERY & " and aa.Week<=" & Week2 & " "
		 end if

		 if Val <> "" then
				Option4 = " Year(a.CreatedDate)=" & Val & " "
                SUBQUERY = SUBQUERY  & " and Year(aa.CreatedDate)=" & Val & " "
		 end if

         QuerySelect = Replace(QuerySelect, "#SUBQUERY#", SUBQUERY)

		 HTMLHidden = HTMLHidden & "<INPUT name='BLType' type=hidden value='" & BLType & "'>"
         HTMLHidden = HTMLHidden & "<INPUT name='ReportType' type=hidden value='" & ReportType & "'>"
         HTMLHidden = HTMLHidden & "<INPUT name='Week' type=hidden value='" & Week & "'>"
         HTMLHidden = HTMLHidden & "<INPUT name='Week2' type=hidden value='" & Week2 & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='Yr' type=hidden value='" & Val & "'>"
	case 22
		 OrderName = " order by a.BLGroupID Desc"
		 QuerySelect = "select a.BLGroupID, a.CreatedTime, a.CreatedDate, a.Countries, a.BLNumber, a.Expired from BLGroups a"
		 HTMLTitle = "<tr><td class=titlelist><b>Fecha</td><td class=titlelist><b>Paises</td><td class=titlelist><b>No. de Carta Porte</td><td class=titlelist><b>Status</td></tr>"

		 BLNumber = Request.Form("BLNumber")
		 Week = CheckNum(Request.Form("Week"))

		 Option1 = " a.Countries in " & Session("Countries") & " "
		 if BLNumber <> "" then
				Option2 = " a.BLNumber like '%" & BLNumber & "%' "
		 end if
		 if Week <> 0 then
				Option3 = " a.Week=" & Week & " "
		 end if

		 HTMLHidden = HTMLHidden & "<INPUT name='BLNumber' type=hidden value='" & BLNumber & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='Week' type=hidden value='" & Week & "'>"
	case 23 'Rastreo
		 elements = 9
		 Week = CheckNum(Request.Form("Week"))
		 BLNumber = Request.Form("BLNumber")
		 MBL = Request.Form("MBL")
		 Name = Request.Form("Name")
         ShipperName = Request.Form("ShipperName")
		 BLType = CheckNum(Request.Form("BLType"))
		 CountryDes = Request.Form("CountryDes")
		 CountryDep = Request.Form("CountryDep")
		 HTMLTitle = "<tr><td class=titlelist><b>Fecha Salida</td><td class=titlelist width=120><b>Carta Porte</td><td class=titlelist><b>Semana</td><td class=titlelist><b>Consignatario</td><td class=titlelist><b>No.Bultos</td><td class=titlelist><b>Carga</td><td class=titlelist><b>Correos</td><td class=titlelist><b>Status</td></tr>"

         select case BLType
         case -1 'Grupo
             OrderName = " order by a.BLGroupID Desc"
			 QuerySelect = "select a.BLGroupID, a.CreatedTime, c.BLExitDate, a.BLNumber, a.Week, b.Clients, b.NoOfPieces, b.DiceContener, '' as correos, (select c.BLID from Tracking c where c.BLID=a.BLGroupID and c.ClientID=-1 limit 1), b.ClientsID, b.AgentsID, b.ShippersID, b.Pos, b.EXID, b.EXDBCountry, b.EXType, b.ColoadersID, b.HBLNumber, b.BLs, replace(b.MBLs, ' ', '-') as MBLs from BLs c, BLDetail b, BLGroups a, BLGroupDetail d "
	
			 'option1 = " c.BLID=b.BLID and c.BLID=d.BLID and a.BLGroupID=d.BLGroupID and (a.Countries in " & Session("Countries") & " or a.CountryDes in " & Session("Countries") & ") "
			 option1 = " c.BLID=b.BLID and c.BLID=d.BLID and a.BLGroupID=d.BLGroupID and b.Expired=0 "
			 
			 if CountryDep <> "" then
					Option2 = " c.Countries='" & CountryDep & "' "
			 end if
         case -2 'CIF Ingreso
             OrderName = " order by b.BLDetailID Desc"
			 QuerySelect = "select b.BLDetailID, b.CreatedTime, b.DischargeDate, b.BLs, b.Week, b.Clients, b.NoOfPieces, b.DiceContener, '' as correos, (select c.BLID from Tracking c where c.BLID=b.BLDetailID and c.ClientID=-2 limit 1), b.ClientsID, b.AgentsID, b.ShippersID, b.Pos, b.EXID, b.EXDBCountry, b.EXType, b.ColoadersID, b.HBLNumber, b.BLs, replace(b.MBLs, ' ', '-') as MBLs from BLs a right join BLDetail b on a.BLID=b.BLID "
	         
			 option1 = " b.ExType=8 and b.Expired=0 "
             
			 if CountryDep <> "" then
					Option2 = " b.CountryOrigen='" & CountryDep & "' "
			 end if
         case Else 'Consolidado, Express
             
			 OrderName = " order by a.BLID Desc"
			 QuerySelect = "select a.BLID, a.CreatedTime, a.BLExitDate, a.BLNumber, a.Week, b.Clients, b.NoOfPieces, b.DiceContener, '' as correos, (select c.BLID from Tracking c where c.BLID=a.BLID and c.ClientID=0 limit 1), b.ClientsID, b.AgentsID, b.ShippersID, b.Pos, b.EXID, b.EXDBCountry, b.EXType, b.ColoadersID, b.HBLNumber, b.BLs, replace(b.MBLs, ' ', '-') as MBLs from BLs a, BLDetail b "
	         'option1 = " a.BLID=b.BLID and (a.Countries in " & Session("Countries") & " or b.CountriesFinalDes in " & Session("Countries") & ") "
			 option1 = " a.BLID=b.BLID and b.Expired=0 "
             
			 if CountryDep <> "" then
					Option2 = " a.CountryDep='" & CountryDep & "' "
			 end if
         End Select
		
		 if Week <> 0 then
				Option3 = " a.Week=" & Week & " "
		 end if
		 if BLNumber <> "" then
				Option4 = " (a.BLNumber like '%" & BLNumber & "%' or b.HBLNumber like '%" & BLNumber & "%') "
		 end if
		 if MBL <> "" then
				Option5 = " b.BLs like '%" & MBL & "%' "
		 end if
		 if Name <> "" then
				Option6 = " b.Clients like '%" & Name & "%' "
		 end if
		 if BLType >= 0 and BLType<>4 then
                Option7 = " a.BLType=" & BLType & " "
		 end if
		 if CountryDes <> "" then
				Option8 = " a.CountryDes='" & CountryDes & "' "
		 end if
         if ShipperName <> "" then
				Option9 = " b.Agents like '%" & ShipperName & "%' "
		 end if

		 HTMLHidden = HTMLHidden & "<INPUT name='Week' type=hidden value='" & Week & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='BLNumber' type=hidden value='" & BLNumber & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='MBL' type=hidden value='" & MBL & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='Name' type=hidden value='" & Name & "'>"
         HTMLHidden = HTMLHidden & "<INPUT name='ShipperName' type=hidden value='" & ShipperName & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='BLType' type=hidden value='" & BLType & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='CountryDep' type=hidden value='" & CountryDep & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='CountryDes' type=hidden value='" & CountryDes & "'>"
	case 26 'DTI Template
		 OrderName = " order by a.Countries, a.DTITemplateID"
		 QuerySelect = "select a.DTITemplateID, a.CreatedTime, a.CreatedDate, a.Countries, a.Name, a.Expired from DTITemplates a"

		 Name = Request.Form("Name")
		 Countries = Request.Form("Countries")
		 HTMLTitle = "<tr><td class=titlelist><b>Fecha</td><td class=titlelist><b>Pais</td><td class=titlelist><b>Nombre</td><td class=titlelist><b>Status</td></tr>"

		 if Name <> "" then
				Option1 = " a.Name like '%" & Name & "%' "
		 end if
		 if CountryDes <> "" then
				Option2 = " a.Countries='" & CountryDes & "' "
		 end if

		 HTMLHidden = HTMLHidden & "<INPUT name='Name' type=hidden value='" & Name & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='Countries' type=hidden value='" & Countries & "'>"
	case 27 'Carga Maritima en Transito
         elements = 9
		 HTMLTitle = "<tr><td class=titlelist width='auto'><b>Descarga Puerto</td><td class=titlelist width='auto'><b>House</td><td class=titlelist width='auto'><b>MBL</td><td class=titlelist width='auto'><b>Contenedor</td><td class=titlelist width='auto'><b>Origen</td><td class=titlelist width='auto'><b>Destino Final</td><td class=titlelist width='auto'><b>Descarga Almacen</td><td class=titlelist width='auto'><b>Status</td></tr>"
		 BLType = CheckNum(Request("BLType"))
		 HBL = Request("HBL")
		 MBL = Request("MBL")
		 CNT = Request("CNT")
		 Countries = Request("Countries")
		 CountryDes = Request("CountriesSearch")
		 Val = CheckNum(Request("DIV"))
		 BLNumber = CheckNum(Request("BLID"))
         IT = CheckNum(Request("IT"))

		 select case BLType
		 case 0 'FCL
             OrderName = " order by B.bl_id desc"
             if  IT = 1 then
                ItineraryType = 0 'Intermodal
             else
                ItineraryType = 11 'Local
             end if
			 QuerySelect = "select C.contenedor_id, " & ItineraryType & ", B.fecha_descarga, B.no_bl, B.mbl, coalesce(C.no_contenedor, ''), B.pais_origen_carga, B.id_pais_final, '', B.cerrado, 0 from bl_completo B, contenedor_completo C"
			 Option1 = "B.bl_id = C.bl_id and B.activo and C.activo and extract(year from B.fecha_ingreso_sistema) >= (extract(year from NOW()) - 1) "
			 if HBL <> "" then
					Option2 = " B.no_bl ilike '%" & HBL & "%' "
			 end if
			 if MBL <> "" then
					Option3 = " B.mbl ilike '%" & MBL & "%' "
			 end if
			 if CNT <> "" then
					Option4 = " C.no_contenedor ilike '%" & CNT & "%' "
			 end if
		 case 1 'LCL
             OrderName = " order by B.bl_id desc"
			 select case Val
			 case 0 'Sin division
                 if  IT = 1 then
                    ItineraryType = 1 'Intermodal
                 else
                    ItineraryType = 12 'Local
                 end if
				 QuerySelect = "select B.bl_id, " & ItineraryType & ", VC.fecha_descarga, B.no_bl, VC.mbl, coalesce(VC.no_contenedor, ''), d.origen, e.destino, '', B.cerrado, B.dividido from bill_of_lading b inner join viaje_contenedor vc on vc.viaje_contenedor_id=b.viaje_contenedor_id inner join viajes v on v.viaje_id=vc.viaje_id left join dblink('dbname=master-aimar port=5432 host=10.10.1.20 user=dbmaster password=aimargt','select a.unlocode_id, a.pais from unlocode a')as d(unlocode_id int8 , origen varchar) on d.unlocode_id=v.id_puerto_origen left join dblink('dbname=master-aimar port=5432 host=10.10.1.20 user=dbmaster password=aimargt','select a.unlocode_id, a.pais from unlocode a')as e(unlocode_id int8 , destino varchar) on e.unlocode_id=v.id_puerto_desembarque"
				 Option1 = "B.viaje_contenedor_id = VC.viaje_contenedor_id and VC.viaje_id = V.viaje_id and B.activo and VC.activo and V.activo and extract(year from B.fecha_ingreso_sistema) >= (extract(year from NOW()) - 1) "
			 case 1 'Con Division
                 if  IT = 1 then
                    ItineraryType = 2 'Intermodal
                 else
                    ItineraryType = 13 'Local
                 end if
				 QuerySelect = "select DB.division_id, " & ItineraryType & ", VC.fecha_descarga, DB.no_bl, VC.mbl, coalesce(VC.no_contenedor, ''), d.origen, e.destino, '', B.cerrado, 0 from bill_of_lading B inner join divisiones_bl DB on DB.bl_asoc=B.bl_id inner join viaje_contenedor vc on vc.viaje_contenedor_id=b.viaje_contenedor_id inner join viajes v on v.viaje_id=vc.viaje_id left join dblink('dbname=master-aimar port=5432 host=10.10.1.20 user=dbmaster password=aimargt','select a.unlocode_id, a.pais from unlocode a')as d(unlocode_id int8 , origen varchar) on d.unlocode_id=v.id_puerto_origen left join dblink('dbname=master-aimar port=5432 host=10.10.1.20 user=dbmaster password=aimargt','select a.unlocode_id, a.pais from unlocode a')as e(unlocode_id int8 , destino varchar) on e.unlocode_id=v.id_puerto_desembarque"
				 Option1 = "B.viaje_contenedor_id = VC.viaje_contenedor_id and VC.viaje_id = V.viaje_id and B.activo and VC.activo and V.activo and extract(year from B.fecha_ingreso_sistema) >= (extract(year from NOW()) - 1) "
			 end select
 			 if HBL <> "" then
					Option2 = " B.no_bl ilike '%" & HBL & "%' "
			 end if
			 if MBL <> "" then
					Option3 = " VC.mbl ilike '%" & MBL & "%' "
			 end if
			 if CNT <> "" then
					Option4 = " VC.no_contenedor ilike '%" & CNT & "%' "
			 end if
			 if BLNumber > 0 then
					Option5 = " B.bl_id=" & BLNumber & " "
			 end if
		 case 2 'AEREO
             OrderName = " order by a.HAWBNumber"
             if  IT = 1 then
                ItineraryType = 10 'Intermodal
             else
                ItineraryType = 9 'Local
             end if
			 QuerySelect = "select a.AWBID, " & ItineraryType & ", a.AWBDate, a.HAWBNumber, a.AWBNumber, a.Voyage, b.Country, c.Country, '', a.Expired, 0 from Awbi a inner join Airports b on b.AirportID=a.AirportDepID inner join Airports c on c.AirportID=a.AirportDesID"
			 Option1 = "a.HAWBNumber<>'' "
			 if HBL <> "" then
					Option2 = " a.HAWBNumber like '%" & HBL & "%' "
			 end if
			 if MBL <> "" then
					Option3 = " a.AWBNumber like '%" & MBL & "%' "
			 end if
			 if CNT <> "" then
					Option4 = " a.Voyage like '%" & CNT & "%' "
			 end if
         case 3 'TERRESTRE INTERMODAL
             OrderName = " order by a.CreatedDate Desc"
             ItineraryType = 14 'Local
			 QuerySelect = "select a.BLDetailID, " & ItineraryType & ", a.DischargeDate, a.HBLNumber, b.BLNumber, '', CountryOrigen, CountriesFinalDes, '', a.Expired, 0 from BLDetail a, BLs b"
			 Option1 = "a.BLID=b.BLID and b.BLType in (0,1) "
			 if HBL <> "" then
					Option2 = " a.HBLNumber like '%" & HBL & "%' "
			 end if
			 if MBL <> "" then
					Option3 = " a.BLNumber like '%" & MBL & "%' "
			 end if
			 'if CNT <> "" then
			 '		Option4 = " a.Voyage like '%" & CNT & "%' "
			 'end if
         case 4 'ALMACEN
             OrderName = " order by a.HAWBNumber"
             if  IT = 1 then
                ItineraryType = 10 'Intermodal
             else
                ItineraryType = 9 'Local
             end if
			 QuerySelect = "select a.AWBID, " & ItineraryType & ", a.AWBDate, a.HAWBNumber, a.AWBNumber, a.Voyage, b.Country, c.Country, '', a.Expired, 0 from Awbi a inner join Airports b on b.AirportID=a.AirportDepID inner join Airports c on c.AirportID=a.AirportDesID"
			 Option1 = "a.HAWBNumber<>'' "
			 if HBL <> "" then
					Option2 = " a.HAWBNumber like '%" & HBL & "%' "
			 end if
			 if MBL <> "" then
					Option3 = " a.AWBNumber like '%" & MBL & "%' "
			 end if
			 if CNT <> "" then
					Option4 = " a.Voyage like '%" & CNT & "%' "
			 end if
		 end select
		 HTMLHidden = HTMLHidden & "<INPUT name='BLType' type=hidden value='" & BLType & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='HBL' type=hidden value='" & HBL & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='MBL' type=hidden value='" & MBL & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='CNT' type=hidden value='" & CNT & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='CountriesSearch' type=hidden value='" & CountryDes & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='Countries' type=hidden value='" & Countries & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='DIV' type=hidden value='" & Val & "'>"
		 HTMLHidden = HTMLHidden & "<INPUT name='BLID' type=hidden value='" & BLNumber & "'>"
         HTMLHidden = HTMLHidden & "<INPUT name='IT' type=hidden value='" & IT & "'>"
	case 32
		 QuerySelect = GetSQLSearch (GroupID)
		 HTMLTitle = "<tr><td class=titlelist><b>Fecha</td><td class=titlelist><b>Bodega</td><td class=titlelist><b>Bolsa</td><td class=titlelist><b>Status</td></tr>"
		 OrderName = " order by a.WarehouseID, BagValue"
		 WareHouseID = CheckNum(Request.Form("WareHouseID"))
		 BagValue = CheckNum(Request.Form("BagValue"))
		 
		 option1 = " a.WarehouseID=b.WarehouseID "
		 if WareHouseID <> 0 then
				Option2 = " a.WareHouseID =" & WareHouseID & " "
		 end if
		 if BagValue <> 0 then
				Option3 = " a.BagValue =" & BagValue & " "
		 end if
		 HTMLHidden = HTMLHidden & "<INPUT name='WareHouseID' type=hidden value='" & WareHouseID & "'>"		 
		 HTMLHidden = HTMLHidden & "<INPUT name='BagValue' type=hidden value='" & BagValue & "'>"	 
	end select

	'Construyendo el Query segun los parametros de busqueda seleccionados en la pagina anterior
	DateFrom = ConvertDate(Request.Form("DateFrom"),3)
	DateTo = ConvertDate(Request.Form("DateTo"),3)
	
	if DateFrom <> "" then
		 Option10 = " a.CreatedDate>='" & DateFrom & "' "
	end if	
	if DateTo <> "" then
		 Option11 = " a.CreatedDate<='" & DateTo & "' "
	end if
	HTMLHidden = HTMLHidden & "<INPUT name=DateFrom type=hidden value='" & DateFrom & "'>"
	HTMLHidden = HTMLHidden & "<INPUT name=DateTo type=hidden value='" & DateTo & "'>"
	MoreOptions = 0
    CreateSearchQuery QuerySelect, Option1, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option2, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option3, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option4, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option5, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option6, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option7, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option8, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option9, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option10, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option11, MoreOptions, " and "
	
	select case GroupID
	case 14
		'QuerySelect = "select b.BLDetailID, a.CreatedTime, a.CreatedDate, b.HBLNumber, b.Clients, a.BLID, a.Consolidated, b.LtEndorseDate, b.ClientsID, b.AgentsID, b.Seps from BLDetail b, BLs a " & _
		'			QuerySelect & " and a.BLID=b.BLID and a.Consolidated=1 UNION " & _
		'			"select a.BLID, a.CreatedTime, a.CreatedDate, b.HBLNumber, b.Clients, a.BLID, a.Consolidated, b.LtEndorseDate, b.ClientsID, b.AgentsID, b.Seps from BLs a, BLDetail b " & _
		'			QuerySelect & " and a.BLID=b.BLID and a.Consolidated=0 group by b.BLID, b.ClientsID, b.AgentsID, b.Seps "					
		elements = 6
		QuerySelect = "select b.BLDetailID, a.CreatedTime, a.CreatedDate, b.Countries, b.HBLNumber, b.Clients, a.BLID, a.Consolidated, b.LtEndorseDate, b.ClientsID, b.AgentsID, b.Seps from BLs a, BLDetail b " & _
					QuerySelect & " and a.BLID=b.BLID and b.Expired = 0 group by b.BLID, b.ClientsID, b.AgentsID, b.Seps "					
	end select
	
	QuerySelect = QuerySelect & OrderName
	'response.write GroupID & "<br>"
	'response.write QuerySelect & "<br>"
    'response.write CountryDes & "<br>"
	HTMLCode = ""
    HTMLCode2 = ""
	if GroupID<>17 then
		DisplaySearchAdminResults HTMLCode
	else
	    Select Case ReportType
        Case 0
		    DisplayCAMCharge HTMLCode, HTMLTitle
	    Case 1
		    DisplayTimesReport HTMLCode, HTMLCode2, HTMLTitle, HTMLTitle2
        Case 2, 4
            DisplayPercentReport HTMLCode, HTMLTitle
        Case 3
            DisplayTrackingStatusReport HTMLCode, HTMLTitle
	    end Select
	end if
%>

<HTML><HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language=javascript src="img/config.js"></SCRIPT>
<SCRIPT src="img/mainLib.js" language="JavaScript"></SCRIPT>
<SCRIPT language="javascript" src="img/validaciones.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
function NextPage(PageNo) {
				 document.forma.P.value = PageNo;
				 document.forma.submit();
}
</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
	<%if JavaMsg <> "" then
			 Response.Write "<SCRIPT>alert('" & JavaMsg & "');</SCRIPT>"
		end if
	%>
	<FORM name="forma" action="Search_ResultsAdmin.asp" method="post" target="_self">
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
  <INPUT name="Action" type=hidden value=1>
  <INPUT name="DownFile" type=hidden value=0>
<INPUT name="P" type=hidden value=1>
  <%=HTMLHidden%>
	<TABLE cellspacing=0 cellpadding=2 width=100% align=center>
		<TR>
		<TD colspan=2 class=label align=right valign=top>
				<TABLE cellspacing=5 cellpadding=2 width=100% align=center>
							 <%=HTMLTitle%>
							 <%=HTMLCode%>
				</TABLE>
                 <%if (GroupID = 17) then%>
                     <TABLE cellspacing=5 cellpadding=2 width=100% align=center>
							 <%=HTMLTitle2%>
                             <%=HTMLCode2%>
				    </TABLE>
                 <%end if%>
		</TD>
	  </TR>
<% if PageCount > 1 then%>
		<TR>
		<TD width=40% colspan=2 class=label align=right valign=top>
				<TABLE cellspacing=5 cellpadding=2 width=100% align=center>
				<TR>
				<TD class=label align=left valign="top" width=15%>
				<%if AbsolutePage > 1 then%>&nbsp;
								<a class=label onclick=JavaScript:NextPage("<%=(AbsolutePage-1)%>"); href=# target=_self><u><< Anterior</u></a>&nbsp;
				<%else%>
								<a class=label href="Search_admin.asp?GID=<%=GroupID%>" target=_self><u><< Regresar</u></a>&nbsp;
				<%end if%>&nbsp;
				</TD>
				<TD class=label align=center>
							 <%
							 for i = 1 to PageCount
							 		 Response.write "&nbsp;<a class=label onclick=JavaScript:NextPage(" & i & ") href=#><u>" & i & "</u></a>&nbsp;"
							 		 if i <> PageCount then
							 		 		Response.write "<font class=label>|</font>" 
							 		 end if
									 if (i mod 20) = 0 then
									 		Response.write "<br>"
									 end if
							 next
							 %>
				</TD>
				<TD class=label align=right valign="top" width=15%>&nbsp;
				<%if PageCount <> AbsolutePage then%> 
						 <a class=label onclick=JavaScript:NextPage("<%=(AbsolutePage+1)%>"); href=# target=_self><u>Siguiente >></u></a>
				<%end if%>&nbsp;
				</TD>
				</TR>
				</TABLE>
		</TD>
	  </TR>
<%else
    if DownFile=0 then%>
		<TR>
		<TD width=40% colspan=2 class=label align=right valign=top>
				<TABLE cellspacing=5 cellpadding=2 width=100% align=center>
				<TR>
				<TD class=label align=left>
				<a class=label onclick=JavaScript:history.back(); href=# target=_self><u><< Regresar</u></a>
				</TD>
				</TR>
				</TABLE>
		</TD>
	  </TR>
<%  end if
 end if%>
		</TABLE>
  </FORM>				
</BODY>
</HTML>
<%
end if
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>