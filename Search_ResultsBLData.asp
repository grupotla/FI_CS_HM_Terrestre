<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"
Dim GroupID, HTMLCode, HTMLTitle, Table, QuerySelect, MoreOptions, CountriesFinalDes, PolicesNo
Dim Option1, Option2, Option3, Option4, Conn, rs, JavaMsg, i, link, Countries, Routing, CountriesOrigen
Dim Name, OrderName, Attn, CountryDep, TransitList, BLNumber, DiceContenerValue, BLIDTransit, ListColor, CountryAL
Dim elements, PageCount, AbsolutePage, HTMLHidden, Consolidated, ConsignerID, HBLNumber, ItemPos, Identifier
Dim BLDetailID, NoOfPieces, ClassNoOfPieces, CommoditiesID, DiceContener, Volumes, BLType, Seps
Dim Weights, ClientsID, Clients, AgentsID, AgentsAddrID, Agents, BLs, DischargeDate, AddressesID, QueryLimit

GroupID = CheckNum(Request("GID"))
QueryLimit = ""

if GroupID >= 2 and GroupID <=36 then
	AbsolutePage = CheckNum(Request.Form("P"))
	if AbsolutePage = 0 then
		 AbsolutePage = 1
	end if
	elements = 5
	PageCount = 0
    Select case GroupID
	case 2, 33
			 OrderName = " order by a.agente"
			 QuerySelect = "select a.agente_id, a.agente, a.direccion, a.telefono, a.fax, a.contacto, a.es_neutral from agentes a"
			 HTMLTitle = "<td class=titlelist><b>Codigo</td><td class=titlelist><b>Nombre Agente</td>"

			 Name = Request.Form("Name")
			 Option1 = "activo=true "
			 if Name <> "" then
			 		Option2 = " a.agente ilike '%" & Name & "%' "
			 end if
			 HTMLHidden = HTMLHidden & "<INPUT name='Name' type=hidden value='" & Name & "'>"
	case 3, 4, 11, 20, 31, 34, 35, 36
			 OrderName = "order by p.codigo, a.nombre_cliente"
			 QuerySelect = 	"select a.id_cliente, p.codigo, a.nombre_cliente, d.id_direccion, a.es_coloader from clientes a, direcciones d, niveles_geograficos n, paises p "						
			 
			 Name = Request.Form("Name")
			 Countries = Request.Form("Countries")
			 BLType = CheckNum(Request.Form("BTP"))

			 if GroupID=4 and BLType=0 and InStr(1,Session("Countries"),"CR")>0 then
			 	HTMLTitle = "<tr><td class=label align=right><b><input type=checkbox name='STA' class=label></td><td class=label><b>INCLUIR AIMAR NICARAGUA</td></tr>"
				HTMLTitle = "<tr><td class=label align=left colspan=2><select name='STA' class=label>" & _
				"<option value=-1>INCLUIR AIMAR?</option>" & _
				"<option value=1>AIMAR NICARAGUA</option>" & _
				"<option value=2>AIMAR LOGISTIC S.A. DE C.V.</option>" & _
				"</select></td></tr>"
			end if

			 HTMLTitle = HTMLTitle & "<tr><td class=titlelist><b>Pais</td><td class=titlelist><b>Codigo</td><td class=titlelist><b>Nombre</td></tr>"

			 Option1 = " a.id_cliente = d.id_cliente " & _
							"and d.id_nivel_geografico = n.id_nivel " & _
							"and n.id_pais = p.codigo " & _
							"and a.id_estatus in (1,2)"
			 Select Case GroupID
			 case 3, 20, 31
			 	Option1 = Option1 & "and a.es_shipper = true "
			 case 4, 11, 35
			 	Option1 = Option1 & "and a.es_consigneer = true "
             case 34, 36
			 	Option1 = Option1 & "and a.es_coloader = true "
			 End Select
							
			 if Name <> "" then
			 		Option2 = " a.nombre_cliente ilike '%" & Name & "%' "
			 end if

 			 if Countries <> "" then
			 		Option3 = " p.codigo ilike '%" & Countries & "%' "
			 end if
			 HTMLHidden = HTMLHidden & "<INPUT name='Name' type=hidden value='" & Name & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='Countries' type=hidden value='" & Countries & "'>"
	case 9
			 OrderName = " order by a.NameES"
			 QuerySelect = 	"select a.CommodityID, TRIM(REPLACE(a.NameES,'	','')), ReqAuth from Commodities a"
             QueryLimit = " limit 150"
			 HTMLTitle = "<td class=titlelist><b>Producto</td>"
			 Name = Request.Form("NameES")
			 Option1 = "Expired=0 "
			 if Name <> "" then
			 		Option2 = " a.NameES ilike '%" & Name & "%' "
			 end if
			 HTMLHidden = HTMLHidden & "<INPUT name='NameES' type=hidden value='" & Name & "'>"
	case 14 'BL Detail
			 OrderName = " order by a.CountriesFinalDes, a.Pos"
			 QuerySelect = 	"select a.BLDetailID, a.NoOfPieces, a.ClassNoOfPieces, a.CommoditiesID, a.DiceContener, a.Volumes, a.Weights, a.ClientsID, a.AddressesID, a.Clients, a.BLs, a.DischargeDate, a.CountriesFinalDes, a.AgentsID, a.Agents, a.DiceContenerValue, a.HBLNumber, a.AgentsAddrID, a.BLIDTransit, a.Seps, a.CountryOrigen, a.PolicyNo, a.CodeReference, a.Container, a.MBLs from BLDetail a "
			 
             
            HTMLTitle = "<tr><td colspan=6 align=center>Busqueda&nbsp;<input type=text name=txtbuscar>&nbsp;<input type=submit></td></tr>" 
          
             HTMLTitle = HTMLTitle & "<td class=titlelist><input type=checkbox name='Set' onClick='Javascript:SetAll();'></td><td class=titlelist><b>Destino Final</b></td><td class=titlelist><b>MBL</b></td><td class=titlelist><b>Producto</b></td><td class=titlelist><b>Cliente</b></td><td class=titlelist><b>Container</b></td>"
			 CountryDep = Request("CTD")
			 Consolidated = CheckNum(Request("CSD"))
			 ConsignerID = CheckNum(Request("CID"))
			 BLType = CheckNum(Request("BTP"))
             'Select Case CountryDep
             '   Case "GT","GTLTF"
             '       CountryAL = "GT','GTLTF'"
             '   Case "SV","SVLTF"
             '       CountryAL = "SV','SVLTF"
             '   Case "HN","HN1","HNLTF"
             '       CountryAL = "HN','HN1','HNLTF"
             '   Case "NI","NILTF"
             '       CountryAL = "NI','NILTF"
             '   Case "CR","CRLTF"
             '       CountryAL = "CR','CRLTF"
             '   Case "PA","PALTF"
             '       CountryAL = "PA','PALTF"
             '   Case "MX","MXLTF"
             '       CountryAL = "MX','MXLTF"
             'End Select
             
             CountryAL = Left(CountryDep,2)

			 'Option1 = " b.BLArrivalDate<>'' and a.InTransit=2 and b.Consolidated=" & Consolidated & " and a.BLID=b.BLID and b.BLType=" & BLType & " and b.CountryDes in " & Session("Countries") & " " 
			 'Option1 = " b.BLArrivalDate<>'' and a.InTransit=2 and b.Consolidated=" & Consolidated & " and a.BLID=b.BLID and b.BLType=" & BLType
			 if BLType<>2 then
                'Option1 = " a.Expired=0 and a.InTransit=1 and a.BLID=-1 and a.BLType=" & BLType & " and a.Countries in ('" & CountryAL & "') "
                Option1 = " a.Expired=0 and a.InTransit=1 and a.BLID=-1 and a.BLType=" & BLType & " and substr(a.Countries,1,2) = '" & CountryAL & "' "
             else
                Option1 = " a.Expired=0 and a.InTransit=1 and a.BLID=-1 and a.BLType=" & BLType & " and a.Countries in " & Session("Countries") & " "
             end if


             if Len(Trim(Request.Form("txtbuscar"))) > 0 then
                Option1 = Option1 & " and UPPER(a.DiceContener) LIKE '%" & Trim(Request.Form("txtbuscar")) & "%' "
             end if

			 'if Consolidated = 0 then
			 '	Option2 = " a.ClientsID=" & ConsignerID & " "
			 'end if
			 HTMLHidden = HTMLHidden & "<INPUT name='CTD' type=hidden value='" & CountryDep & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='CSD' type=hidden value='" & Consolidated & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='CID' type=hidden value='" & ConsignerID & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='BTP' type=hidden value='" & BLType & "'>"
	case 28 'Carga en Transito
			 OrderName = " order by a.id_routing desc"
			 QuerySelect = 	"select a.id_routing, a.routing from routings a left join seguros b on (a.id_routing=b.id_routing and b.anulado=false) left join routing_terrestre c on a.id_routing = c.id_routing inner join transporte d on a.id_transporte = d.id_transporte "
             HTMLTitle = "<td class=titlelist><b>Routing</b></td>"
			 Routing = Request.Form("RO")
			 BLType = CheckNum(Request.Form("BLType"))
			 Countries = Request.Form("Countries")
			 CountryDep = Request.Form("CountriesSearch")
			 CountriesFinalDes = Request.Form("CountriesFinalDes")
			 Option1 = " a.id_transporte=" & BLType & " and a.id_routing_type=2 and ((a.seguro=true and (a.activo=true or a.activo=false)) or (a.activo=true and a.seguro=false)) and a.bl_id = 0 and a.borrado=false "'Transporte Terrestre (Consol o Express), Routings Tipo Internos
			 
			 if Routing <> "" then
			 		Option2 = " a.routing ilike '%" & Routing & "%' "
			 end if

			 if CountryDep <> "" then
			 		Option3 = " a.id_pais = '" & CountryDep & "' "
			 end if

			 HTMLHidden = HTMLHidden & "<INPUT name='RO' type=hidden value='" & Routing & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='BLType' type=hidden value='" & BLType & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='Countries' type=hidden value='" & Countries & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='CountriesSearch' type=hidden value='" & CountryDep & "'>"
	case 29
			 OrderName = "order by a.desc_rubro_es"
			 QuerySelect = 	"select a.id_rubro, a.desc_rubro_es from rubros a "
			 HTMLTitle = "<td class=titlelist><b>ID</b></td><td class=titlelist><b>Routing</b></td>"
			 ItemPos = CheckNum(Request("N"))
			 Name = Request.Form("Name")
			 Option1 = " a.id_estatus=1 "'Rubro Activo
			 if Name <> "" then
			 		Option2 = " a.desc_rubro_es ilike '%" & Name & "%' "
			 end if
			 HTMLHidden = HTMLHidden & "<INPUT name='Name' type=hidden value='" & Name & "'>" 
			 HTMLHidden = HTMLHidden & "<INPUT name='N' type=hidden value='" & ItemPos & "'>"
	case 30
			 Agents = Request.Form("ST")
			 ItemPos = CheckNum(Request("N"))
			 Name = Request.Form("Name")
				 
			 select Case Agents
			 Case 0 'Linea Aereo
				 OrderName = "order by a.name"
				 QuerySelect = 	"select a.carrier_id, a.name, b.es_afecto, 0 from carriers a, regimen_tributario b "
				 HTMLTitle = "<td class=titlelist><b>Codigo</td><td class=titlelist><b>L&iacute;nea&nbsp;Aerea</b></td>"
				 Option1 = " a.tiporegimen=b.id_regimen and trim(a.name)<>'' "
				 if Name <> "" then
						Option2 = " a.name ilike '%" & Name & "%' "
				 end if
			 Case 1 'Agentes
				 OrderName = "order by a.agente"
				 QuerySelect = 	"select a.agente_id, a.agente, b.es_afecto, a.es_neutral from agentes a, regimen_tributario b "
				 HTMLTitle = "<td class=titlelist><b>Codigo</td><td class=titlelist><b>Agente</b></td>"
				 Option1 = " a.tiporegimen=b.id_regimen and a.activo=true and trim(a.agente)<>'' "
				 if Name <> "" then
						Option2 = " a.agente ilike '%" & Name & "%' "
				 end if
			 Case 2 'Naviera
				 OrderName = "order by a.nombre"
				 QuerySelect = 	"select a.id_naviera, a.nombre, b.es_afecto, 0 from navieras a, regimen_tributario b "
				 HTMLTitle = "<td class=titlelist><b>Codigo</td><td class=titlelist><b>Naviera</b></td>"
				 Option1 = " a.tiporegimen=b.id_regimen and a.activo=true and trim(a.nombre)<>'' "
				 if Name <> "" then
						Option2 = " a.nombre ilike '%" & Name & "%' "
				 end if
			 Case 3	'Proveedores (Otros)
				 OrderName = "order by a.nombre"
				 QuerySelect = 	"select a.numero, a.nombre, b.es_afecto, 0 from proveedores a, regimen_tributario b "
				 HTMLTitle = "<td class=titlelist><b>Codigo</td><td class=titlelist><b>Proveedor</b></td>"
				 Option1 = " a.tiporegimen=b.id_regimen and a.status in (0,1) and trim(a.nombre)<>'' "
				 if Name <> "" then
						Option2 = " a.nombre ilike '%" & Name & "%' "
				 end if
			 End Select
			 HTMLHidden = HTMLHidden & "<INPUT name='Name' type=hidden value='" & Name & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='ST' type=hidden value='" & Agents & "'>"
			 HTMLHidden = HTMLHidden & "<INPUT name='N' type=hidden value='" & ItemPos & "'>"
	end select

	MoreOptions = 0
	CreateSearchQuery QuerySelect, Option1, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option2, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option3, MoreOptions, " and "
	CreateSearchQuery QuerySelect, Option4, MoreOptions, " and "
	QuerySelect = QuerySelect & OrderName & QueryLimit
	'response.write GroupID & "<br>" 
	'response.write QuerySelect & "<br>"
    'QuerySelect = "select a.numero, a.nombre, b.es_afecto from proveedores a, regimen_tributario b where a.tiporegimen=b.id_regimen and a.status in (0,1) and trim(a.nombre)<>'' order by a.nombre"
	
    HTMLCode = ""
    select Case GroupID
	case 2, 3, 4, 9, 11, 20, 28, 29, 30, 31, 33, 34, 35, 36
		OpenConn2 Conn
	case else
		OpenConn Conn
	end select

	'Buscando los archivos que coinciden con el query de Busqueda
	Set rs = Conn.Execute(QuerySelect)
	if Not rs.EOF then
	'if False then
		'Obteniendo la cantidad de resultados por busqueda
		rs.PageSize = 5
		'Saltando a la pagina seleccionada
  	  	rs.AbsolutePage = AbsolutePage
		PageCount = rs.PageCount
		'Desplegando los resultados de la pagina
		select Case GroupID
		'a.agente_id, a.agente, a.direccion, a.telefono, a.fax, contacto from agentes a"
		
		case 2
			for i=1 to rs.PageSize			
				if CheckNum(rs(6))=0 then
                    ListColor = "list"
                    Identifier = ""
                else
                    ListColor = "listwarning"
                    Identifier = "[Neutral]"
                end if
                
                link = "<td class=" & ListColor & "><a class=labellist href=# onclick=" & _
				"""top.opener.document.forms[0].SenderData.value = '" & rs(1) & "\n" & rs(2) & "\n"
				link = link & rs(3) & "    " & rs(4)
				if rs(5) <> "" then 'Atencion
					link = link & "\nATTN:" & rs(5)
				end if
				link = link & "';top.opener.document.forms[0].SenderID.value=" & rs(0) & ";" & _
                "top.opener.document.forms[0].AgentNeutral.value = '"& CheckNum(rs(6)) & "';top.close();"">"
				
				HTMLCode = HTMLCode & "<tr>" & link & rs(0) & "</td>" & link & rs(1) & "&nbsp;" & Identifier & "</a></td></tr>"

				rs.MoveNext
				If rs.EOF Then Exit For 
   	    	next 
		case 3, 4, 11, 20, 31, 34, 35, 36
			for i=1 to rs.PageSize
                if CheckNum(rs(4))=0 then
                    ListColor = "list"
                    Identifier = ""
                else
                    ListColor = "listwarning"
                    Identifier = "[Coloader]"
                end if

				HTMLCode = HTMLCode & "<tr><td class=" & ListColor & "><a class=labellist href=# onclick=Javascript:SetMaster(" & GroupID & "," & rs(0) & "," & rs(3) & ");>" & rs(1) & "</a></td>" & _
				"<td class=" & ListColor & "><a class=labellist href=# onclick=Javascript:SetMaster(" & GroupID & "," & rs(0) & "," & rs(3) & ");>" & rs(0) & "</a></td>" & _
				"<td class=" & ListColor & "><a class=labellist href=# onclick=Javascript:SetMaster(" & GroupID & "," & rs(0) & "," & rs(3) & ");>" & rs(2) & "&nbsp;" & Identifier & "</a></td></tr>"
				rs.MoveNext
				If rs.EOF Then Exit For 
   	    	next 
		case 9
			for i=1 to rs.PageSize
				'link = "<td class=list><a class=labellist href=# onclick=" & _
				'"""if (top.opener.document.forms[0].BLDetailID.value != '') {com='|';} else {com='';}; " & _
				'"top.opener.document.forms[0].BLIDTransit.value = top.opener.document.forms[0].BLIDTransit.value + com + '0';" & _
				'"top.opener.document.forms[0].BLDetailID.value = top.opener.document.forms[0].BLDetailID.value + com + '0';" & _
				'"top.opener.document.forms[0].InTransit.value = top.opener.document.forms[0].InTransit.value + com + '0';" & _
				'"top.opener.EnumDiceContener('" & rs(1) & "', " & rs(0) & ");ReqAuth("&rs(2)&");top.close();"">"
				
				link = "<td class=list><a class=labellist href=# onclick=" & _
				"""top.opener.document.forms[0].CommoditiesID.value='" & rs(0) & "';" & _
				"top.opener.document.forms[0].DiceContener.value='" & mid(rs(1),1,100) & "';" & _
				"ReqAuth(" & rs(2) & ");top.close();"">"
				
				HTMLCode = HTMLCode & "<tr>" & link & rs(1) & "</a></td></tr>"
	
				rs.MoveNext
				If rs.EOF Then Exit For 
   	    	next
		case 14

            dim fleet, flag


			for i = 1 to rs.PageSize


                On Error Resume Next   
                
                        flag = 0   
    
                        fleet = GroupData( CheckNum(rs(7)) ) 

                If Err.Number<>0 then
	                response.write "list :" & Err.Number & " - " & Err.Description & "<br>"  
                end if
    
                 if fleet <> "NA" AND fleet <> "" then
                  
                    if CheckNum(rs(22)) = 0 then

                        flag = 1 'indica que el code reference no se ingreso y que el cliente es tipo colgate

                    end if

                 end if


				link = "<td class=list><input type=checkbox name='Pos" & i & "' onclick='return ValidaCodeReference(this," & flag & ")'></td>" & _ 
				"<td class=list><a href=# class=labellist onclick=Javascript:SetList('Pos" & i & "'," & flag & ");>" & rs(12) & "</a></td>" & _
				"<td class=list><a href=# class=labellist onclick=Javascript:SetList('Pos" & i & "'," & flag & ");>" & rs(23) & "</a></td>" & _
				"<td class=list><a href=# class=labellist onclick=Javascript:SetList('Pos" & i & "'," & flag & ");>" & rs(4) & "</a></td>" & _
				"<td class=list><a href=# class=labellist onclick=Javascript:SetList('Pos" & i & "'," & flag & ");>" & rs(9) & "</a></td>" & _
				"<td class=list><a href=# class=labellist onclick=Javascript:SetList('Pos" & i & "'," & flag & ");>" & rs(24) & "</a></td>"

				HTMLCode = HTMLCode & "<tr>" & link & "</tr>"
	
				BLDetailID = BLDetailID & "BLDetailID[" & i & "]=" & rs(0) & ";" & vbCrLf
				NoOfPieces = NoOfPieces & "NoOfPieces[" & i & "]='" & rs(1) & "';" & vbCrLf
				ClassNoOfPieces = ClassNoOfPieces & "ClassNoOfPieces[" & i & "]='" & rs(2) & "';" & vbCrLf
				CommoditiesID = CommoditiesID & "CommoditiesID[" & i & "]=" & rs(3) & ";" & vbCrLf
				DiceContener = DiceContener & "DiceContener[" & i & "]='" & rs(4) & "';" & vbCrLf
				Volumes = Volumes & "Volumes[" & i & "]='" & rs(5) & "';" & vbCrLf
				Weights = Weights & "Weights[" & i & "]='" & rs(6) & "';" & vbCrLf
				ClientsID = ClientsID & "ClientsID[" & i & "]=" & rs(7) & ";" & vbCrLf
				AddressesID = AddressesID & "AddressesID[" & i & "]=" & rs(8) & ";" & vbCrLf
				Clients = Clients & "Clients[" & i & "]='" & rs(9) & "';" & vbCrLf
				BLs = BLs & "BLs[" & i & "]='" & rs(10) & "';" & vbCrLf
				DischargeDate = DischargeDate & "DischargeDate[" & i & "]='" & rs(11) & "';" & vbCrLf
				CountriesFinalDes = CountriesFinalDes & "CountriesFinalDes[" & i & "]='" & rs(12) & "';" & vbCrLf
				AgentsID = AgentsID & "AgentsID[" & i & "]=" & rs(13) & ";" & vbCrLf
				Agents = Agents & "Agents[" & i & "]='" & rs(14) & "';" & vbCrLf
				DiceContenerValue = DiceContenerValue & "DiceContenerValue[" & i & "]='" & rs(15) & "';" & vbCrLf
				HBLNumber = HBLNumber & "HBLNumber[" & i & "]='" & rs(16) & "';" & vbCrLf
				AgentsAddrID = AgentsAddrID & "AgentsAddrID[" & i & "]=" & rs(17) & ";" & vbCrLf
				BLIDTransit = BLIDTransit & "BLIDTransit[" & i & "]=" & rs(18) & ";" & vbCrLf
				Seps = Seps & "Seps[" & i & "]=" & rs(19) & ";" & vbCrLf
				CountriesOrigen = CountriesOrigen & "CountriesOrigen[" & i & "]='" & rs(20) & "';" & vbCrLf
				PolicesNo = PolicesNo & "PolicesNo[" & i & "]='" & rs(21) & "';" & vbCrLf
                rs.MoveNext
				If rs.EOF Then Exit For
   	    	next
			if i>rs.PageSize then
				i=i-1
			end if
			HTMLCode = HTMLCode & "<tr><td colspan=4 align=center><input class=label type=button value='Asignar Transito' onclick='Javascript:SetTransit();return false;'></td></tr>" 
		case 28
            for i=1 to rs.PageSize
				link = "<td class=list><a class=labellist href='InsertData.asp?GID=28&EID=" & rs(0) & "&ET=" & BLType & "&CTR=" & Countries & "&CTR2=" & CountryDep & "'>"

				HTMLCode = HTMLCode & "<tr>" & _
				link & rs(1) & "</a></td></tr>"
	
				rs.MoveNext
				If rs.EOF Then Exit For 
   	    	next
		case 29
			for i=1 to rs.PageSize
				link = "<td class=list><a class=labellist href=# onclick=" & _
				"""top.opener.document.forms[0].I" & ItemPos & ".value = "& rs(0) & ";" & _
				"top.opener.document.forms[0].N" & ItemPos & ".value = '"& trim(rs(1)) & "';" & _
				"top.close();"">"
				
				HTMLCode = HTMLCode & "<tr>" & _
				link & rs(0) & "</a></td>" & _
				link & trim(rs(1)) & "</a></td></tr>"
	
				rs.MoveNext
				If rs.EOF Then Exit For 
   	    	next 
		case 30
			for i=1 to rs.PageSize
                if CheckNum(rs(3))=0 then
                    ListColor = "list"
                    Identifier = ""
                else
                    ListColor = "listwarning"
                    Identifier = "[Neutral]"
                end if

				link = "<td class=" & ListColor & "><a class=labellist href=# onclick=" & _
				"""top.opener.document.forms[0].SI" & ItemPos & ".value = "& rs(0) & ";" & _
				"top.opener.document.forms[0].SN" & ItemPos & ".value = '"& trim(replace(rs(1),"""","",1,-1)) & "';" & _
				"top.opener.document.forms[0].SAF" & ItemPos & ".value = "& rs(2) & ";" & _
                "top.opener.document.forms[0].SNEU" & ItemPos & ".value = "& rs(3) & ";" & _
                "top.opener.ValidarDoble("& ItemPos & ");top.close();"">"
				HTMLCode = HTMLCode & "<tr>" & _
				link & rs(0) & "</a></td>" & _
				link & trim(rs(1)) & "&nbsp;" & Identifier & "</a></td></tr>"
	
				rs.MoveNext
				If rs.EOF Then Exit For 
   	    	next 
         case 33
			for i=1 to rs.PageSize
				if CheckNum(rs(6))=0 then
                    ListColor = "list"
                    Identifier = ""
                else
                    ListColor = "listwarning"
                    Identifier = "[Neutral]"
                end if
                link = "<td class=" & ListColor & "><a class=labellist href=# onclick=" & _
				"""top.opener.document.forms[0].AgentsID.value = "& rs(0) & ";" & _
				"top.opener.document.forms[0].AgentsAddrID.value = 0;" & _
				"top.opener.document.forms[0].Agents.value = '"& rs(1) & "';" & _
				"top.opener.document.forms[0].AgentNeutral.value = '"& CheckNum(rs(6)) & "';top.close();"">"
				HTMLCode = HTMLCode & "<tr>" & _
				link & rs(0) & "</a></td>" & _
				link & trim(rs(1)) & "&nbsp;" & Identifier & "</a></td></tr>"	
				rs.MoveNext
				If rs.EOF Then Exit For 
   	    	next 
		end select
	else
		If GroupID = 28 Then
            i = 0
            CloseOBJs rs, Conn
            OpenConn Conn
            QuerySelect = "select BLs, Countries, Week from BLDetail where BLs = '" & Routing & "' and Expired = 0 order by BLDetailID LIMIT 1"
            Set rs = Conn.Execute(QuerySelect)
	        if Not rs.EOF then
                if CheckNum(rs(2)) > 0 then
                    JavaMsg = "Este RO ya fue cargado al Sistema de Tráfico Terrestre y se encuentra en Asignados de la empresa " & TranslateCountry(rs(1)) & " Semana " & rs(2)
                else
                    JavaMsg = "Este RO ya fue cargado al Sistema de Tráfico Terrestre y se encuentra en Pendientes de la empresa " &TranslateCountry(rs(1))
                end if
            else
                JavaMsg = "El RO que busca no existe o ha sido eliminado. Por favor póngase en contacto con Sales Support para más información."
            end if
        End If
	end if
CloseOBJs rs, Conn
%>

<HTML><HEAD><TITLE>Aimar - Terrestre</TITLE></HEAD>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/config.js"></SCRIPT>
<SCRIPT src="img/mainLib.js" language="JavaScript"></SCRIPT>
<SCRIPT language="javascript" src="img/validaciones.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
var ntr = "";
var com = "";
function NextPage(PageNo) {
				 document.forma.P.value = PageNo;
				 document.forma.submit();
}
function ReqAuth (val) { //Para Commodities
	if (val==1) {
		alert("Aviso: este producto requiere tramite de Autorizacion");
	}	
}

function SetMaster(GID, OID, AID) {
	var STA = 0;
	<%if GroupID=4 and BLType=0 and InStr(1,Session("Countries"),"CR")>0 then%>
		if (document.forma.STA.value > 0) {
			STA = document.forma.STA.value;
		}
	<%end if%>
	document.location.href = "SetMaster.asp?GID=" + GID + "&OID=" + OID + "&AID=" + AID + "&STA=" + STA
}

<%
if GroupID = 14 then
	response.write "var BLDetailID= new Array();" & vbCrLf & _
				"var NoOfPieces = new Array();" & vbCrLf & _
				"var ClassNoOfPieces = new Array();" & vbCrLf & _
				"var CommoditiesID = new Array();" & vbCrLf & _
				"var DiceContener = new Array();"& vbCrLf & _
				"var Volumes = new Array();" & vbCrLf & _
				"var Weights = new Array();" & vbCrLf & _
				"var ClientsID = new Array();" & vbCrLf & _
				"var AddressesID = new Array();" & vbCrLf & _
				"var Clients = new Array();" & vbCrLf & _
				"var BLs = new Array();" & vbCrLf & _
				"var DischargeDate = new Array();" & vbCrLf & _
				"var CountriesFinalDes = new Array();" & vbCrLf & _
				"var AgentsID = new Array();" & vbCrLf & _
				"var AgentsAddrID = new Array();" & vbCrLf & _
				"var Agents = new Array();" & vbCrLf & _
				"var DiceContenerValue = new Array();" & vbCrLf & _
				"var HBLNumber = new Array();" & vbCrLf & _
				"var BLIDTransit = new Array();" & vbCrLf & _
				"var Seps = new Array();" & vbCrLf & _
				"var CountriesOrigen = new Array();" & vbCrLf & _
				"var PolicesNo = new Array();" & vbCrLf
				
				
	response.write BLDetailID & vbCrLf & _
				NoOfPieces & vbCrLf & _
				ClassNoOfPieces & vbCrLf & _
				CommoditiesID & vbCrLf & _
				DiceContener & vbCrLf & _
				Volumes & vbCrLf & _
				Weights & vbCrLf & _
				ClientsID & vbCrLf & _
				AddressesID & vbCrLf & _
				Clients & vbCrLf & _
				BLs & vbCrLf & _
				DischargeDate & vbCrLf & _
				CountriesFinalDes & vbCrLf & _
				AgentsID & vbCrLf & _
				AgentsAddrID & vbCrLf & _
				Agents & vbCrLf & _
				DiceContenerValue & vbCrLf & _
				HBLNumber & vbCrLf & _
				BLIDTransit & vbCrLf & _
				Seps & vbCrLf & _
				CountriesOrigen & vbCrLf & _
				PolicesNo & vbCrLf & _
				"var Consolidated = " & Consolidated & ";" & vbCrLf
%>
function SetList(Pos, flag) {

    var e = document.forma.elements[Pos];
    
    var res = true;

    if (flag == 1)    
    res = ValidaCodeReference(e, flag)

    //alert(res);

    if (res) {

        if (e.checked){
		    e.checked = false;
	    } else {
		    e.checked = true;
	    }
    }
}

function SetTransit() {
	var Checked=0; //Para validar que al menos un registro fue seleccionado
	var CheckedPos=0;
	//Eliminar ceros como |0|0|...
	var Path = "|"+top.opener.document.forms[0].BLDetailID.value;
	Path = Path.replace(/\|0/g,"");
	//Eliminar el primer |
	if (Path.substr(0,1) == "|") {
		Path = Path.substr(1,Path.length);
	}
	
	if (Path == ""){
		var Patrn = new RegExp("-1");
	} else {
		var Patrn = new RegExp(Path);
	}	

	//Verificando que al menos un registro este seleccionado
	for (var i=1; i<=<%=i%>; i++) {
		if (document.forma.elements["Pos" + i].checked) {
			Checked++;
		}
	}
	if (Checked==0) {
		alert("Debe seleccionar al menos un registro para Asignar a la CP");
		return(false);
	}
	
	<%if BLType=2 or BLType=3 then 'Solo aplica para Trafico Local, solo se puede Asignar un registro%>
	if (Checked>1) {
		alert("En Trafico Local solo puede seleccionar un registro para Asignar a la CP");
		return(false);
	}
	<%end if%>

	//Asignando los Valores
	for (var i=1; i<=<%=i%>; i++) {
		if (document.forma.elements["Pos" + i].checked) {
			<%if BLType=2 or BLType=3 then 'Solo aplica para Trafico Local, solo se puede Asignar un registro o varios registros de un mismo cliente%>
			CheckedPos = i;
			if ((top.opener.document.forms[0].ConsignerID.value=="") || (top.opener.document.forms[0].ConsignerID.value==ClientsID[CheckedPos])) {
			} else {
				alert("En Trafico Local solo puede asignar carga de un mismo cliente en cada CP");
				return(false);
			};		
			<%end if%>
			//Si el BLID no esta asignado se puede asignar
			if (!Patrn.test(BLDetailID[i])) {
				if (top.opener.document.forms[0].BLDetailID.value != '') {ntr='\n'; com='|';} else {ntr=''; com='';};
				top.opener.document.forms[0].BLDetailID.value = top.opener.document.forms[0].BLDetailID.value + com + BLDetailID[i];
				top.opener.document.forms[0].BLIDTransit.value = top.opener.document.forms[0].BLIDTransit.value + com + BLIDTransit[i];
				top.opener.document.forms[0].NoOfPieces.value = top.opener.document.forms[0].NoOfPieces.value + ntr + NoOfPieces[i];
				top.opener.document.forms[0].ClassNoOfPieces.value = top.opener.document.forms[0].ClassNoOfPieces.value + ntr + ClassNoOfPieces[i];
				top.opener.document.forms[0].CommoditiesID.value = top.opener.document.forms[0].CommoditiesID.value + com + CommoditiesID[i];
				top.opener.document.forms[0].DiceContener.value = top.opener.document.forms[0].DiceContener.value + com + DiceContener[i];
				top.opener.document.forms[0].Volumes.value = top.opener.document.forms[0].Volumes.value + ntr + Volumes[i];
				top.opener.document.forms[0].Weights.value = top.opener.document.forms[0].Weights.value + ntr + Weights[i];
				top.opener.document.forms[0].BLs.value = top.opener.document.forms[0].BLs.value + ntr + BLs[i];
				top.opener.document.forms[0].DischargeDate.value = top.opener.document.forms[0].DischargeDate.value + ntr + DischargeDate[i];
				top.opener.document.forms[0].CountriesFinalDes.value = top.opener.document.forms[0].CountriesFinalDes.value + ntr + CountriesFinalDes[i];
				top.opener.document.forms[0].AgentsID.value = top.opener.document.forms[0].AgentsID.value + com + AgentsID[i];
				top.opener.document.forms[0].AgentsAddrID.value = top.opener.document.forms[0].AgentsAddrID.value + com + AgentsAddrID[i];
				top.opener.document.forms[0].Agents.value = top.opener.document.forms[0].Agents.value + ntr + Agents[i];
				top.opener.document.forms[0].DiceContenerValue.value = top.opener.document.forms[0].DiceContenerValue.value + ntr + DiceContenerValue[i];
				top.opener.document.forms[0].HBLNumber.value = top.opener.document.forms[0].HBLNumber.value + com + HBLNumber[i];
				top.opener.document.forms[0].Seps.value = top.opener.document.forms[0].Seps.value + ntr + Seps[i];
				top.opener.document.forms[0].CountriesOrigen.value = top.opener.document.forms[0].CountriesOrigen.value + ntr + CountriesOrigen[i];
				top.opener.document.forms[0].ClientsID.value = top.opener.document.forms[0].ClientsID.value + com + ClientsID[i];
				top.opener.document.forms[0].AddressesID.value = top.opener.document.forms[0].AddressesID.value + com + AddressesID[i];
				top.opener.document.forms[0].Clients.value = top.opener.document.forms[0].Clients.value + ntr + Clients[i];
				top.opener.document.forms[0].InTransit.value = top.opener.document.forms[0].InTransit.value + com + '1';
				<%if BLType=2 then%>
				top.opener.document.forms[0].PolicyNo.value = top.opener.document.forms[0].PolicyNo.value + com + PolicesNo[i];
                //En entrega local todos los Paises deben ser el mismo pais
                top.opener.document.forms[0].CountryDep.value = CountriesFinalDes[i];
                top.opener.document.forms[0].CountryDes.value = CountriesFinalDes[i];
                //top.opener.document.forms[0].Countries.value = CountriesFinalDes[i];
                //alert (CountriesFinalDes[i]);
				<%end if%>
				top.opener.document.forms[0].HaveInvoices.value = top.opener.document.forms[0].HaveInvoices.value + com + '0';
			}
		}
	}
	top.opener.EnumDiceContener("", "");
	top.opener.SumVals(top.opener.document.forms[0].NoOfPieces, top.opener.document.forms[0].TotNoOfPieces);
	top.opener.SumVals(top.opener.document.forms[0].Volumes, top.opener.document.forms[0].TotVolume);	
	top.opener.SumVals(top.opener.document.forms[0].Weights, top.opener.document.forms[0].TotWeight);	
	top.opener.SumVals(top.opener.document.forms[0].DiceContenerValue, top.opener.document.forms[0].TotDiceContenerValue);
	
	<%if BLType=2 or BLType=3 then 'Solo aplica para Trafico Local, solo se puede Asignar un registro%>
		document.location.href = "SetMaster.asp?GID=14&OID=" + ClientsID[CheckedPos] + "&AID=" + AddressesID[CheckedPos] + "&SBLID=" + BLDetailID[CheckedPos]
	<%else%>
		top.close();
	<%end if%>
}

function resizeOuterTo(w,h) {
 if (parseInt(navigator.appVersion)>3) {
   if (navigator.appName=="Netscape") {
    top.outerWidth=w;
    top.outerHeight=h;
   }
   else top.resizeTo(w,h);
 }
}

function ValidaCodeReference(e, flag) {

    var res = false;
    //alert( flag + " " + e.checked);

    if (flag == 1) {  //indica que el code reference no se ingreso y que el cliente es tipo colgate
        e.checked = false; 
        alert("Codigo Referencia no tiene valor, porfavor dirigirse a cargas cif a completar datos.");
    } else {
        res = true;
    }

    return  res;//e.checked;
}

function SetAll() {
	if (document.forma.Set.checked) {
		for (var i=1; i<=<%=i%>; i++) {
			document.forma.elements["Pos" + i].checked = true;
		}
	} else {
		for (var i=1; i<=<%=i%>; i++) {
			document.forma.elements["Pos" + i].checked = false;
		}

	}
}
<%
end if
%>
</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="JavaScript:self.focus();">
	<%if JavaMsg <> "" then
			 Response.Write "<SCRIPT>alert('" & JavaMsg & "');</SCRIPT>"
		end if
	%>
	<FORM name="forma" action="Search_ResultsBLData.asp" method="post" target=_self>
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
  	<INPUT name="Action" type=hidden value=1>
	<INPUT name="P" type=hidden value=1>
	
  <%=HTMLHidden%>
	<TABLE cellspacing=0 cellpadding=2 width=100% align=center>
		<TR>
		<TD colspan=2 class=label align=right valign=top>
			<TABLE cellspacing=5 cellpadding=2 width=100% align=center>
				<%=HTMLTitle%>
				<%=HTMLCode%>
			</TABLE>
		</TD>
	  </TR>
<% if PageCount > 1 then%>
		<TR>
		<TD width=100% colspan=2 class=label align=right valign=top>
				<TABLE cellspacing=5 cellpadding=2 width=100% align=center>
				<TR>
				<TD class=label align=left valign="top" width=65>
				<%if AbsolutePage > 1 then%>&nbsp;
								<a class=label onclick=JavaScript:NextPage("<%=(AbsolutePage-1)%>"); href=# target=_self><u><< Anterior</u></a>&nbsp;
				<%end if%>&nbsp;
				</TD>
				<TD class=label align=center width="300">
							 <%
							 for i = 1 to PageCount
							 		 Response.write " <a class=label onclick=JavaScript:NextPage(" & i & ") href=#><u>" & i & "</u></a> "
							 		 if i <> PageCount then
							 		 		Response.write "<font class=label>|</font>" 
							 		 end if
									 'if (i mod 12) = 0 then
									 	'	Response.write "<br>"
									 'end if
							 next
							 %>
				</TD>
				<TD class=label align=right valign="top" width=65>&nbsp;
				<%if PageCount <> AbsolutePage then%> 
						 <a class=label onclick=JavaScript:NextPage("<%=(AbsolutePage+1)%>"); href=# target=_self><u>Siguiente >></u></a>
				<%end if%>&nbsp;
				</TD>
				</TR>
				</TABLE>
		</TD>
	  </TR>
<%else%>
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
<%end if%>		
		</TABLE>
  </FORM>				
</BODY>
<%if GroupID = 14 then%>
<SCRIPT>resizeOuterTo(500,400)</SCRIPT>
<%end if%>
</HTML>
<%
end if
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>