<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
Checking "0|1|2"

Dim Conn, rs, ObjectID, Name, Address, Phone1, Phone2, Attn, QuerySelect, GroupID, Countries, AddressID, BLType, SetAimar
Dim SBLID, PickUpData, DeliveryData, ClientContact, EXID, ChargeWH, DeliveryWH, ClientType

ObjectID = CheckNum(Request("OID"))
GroupID = CheckNum(Request("GID"))
AddressID = CheckNum(Request("AID"))
SetAimar = CheckNum(Request("STA"))
SBLID = CheckNum(Request("SBLID"))

if ObjectID <> 0 then
	'Obteniendo los datos del Cliente en la Master
	QuerySelect = "select a.nombre_cliente, d.direccion_completa, d.phone_number, p.codigo, a.es_coloader " & _
							"from clientes a, direcciones d, niveles_geograficos n, paises p " & _
							" where a.id_cliente = d.id_cliente" & _
							" and d.id_nivel_geografico = n.id_nivel" & _
							" and n.id_pais = p.codigo" & _
							" and a.id_cliente = " & ObjectID  & _
							" and d.id_direccion = " & AddressID
	OpenConn2 Conn
	Set rs = Conn.Execute(QuerySelect)
	if Not rs.EOF then
		Name = rs(0)
		Address = PurgeData(rs(1))
		Phone1 = rs(2)
		Countries = rs(3)
        ClientType = CheckNum(rs(4))
	end if
	CloseOBJ rs

	set rs = Conn.Execute("select numero_telefono from cli_telefonos where id_cliente=" & ObjectID)
	if Not rs.EOF then
		Phone2 = rs(0)
	end if
	CloseOBJ rs

	set rs = Conn.Execute("select nombres from contactos where id_cliente=" & ObjectID)
	if Not rs.EOF then
		Attn = rs(0)
	end if
	CloseOBJs rs, Conn
	
	select Case SetAimar
	case 1
		Name = Name & " / AIMAR DE NICARAGUA S.A."
	case 2
		Name = Name & " / AIMAR LOGISTIC S.A. DE C.V."
	end select
	
	if GroupID=14 then
		'si el Grupo=14, Obteniendo la direccion de contacto, recoleccion y entrega
		OpenConn Conn
		Set rs = Conn.Execute("select PickUpData, DeliveryData, Notify, EXID from BLDetail where BLDetailID=" & SBLID)
		if Not rs.EOF then
			PickUpData = PurgeData2(rs(0))
			DeliveryData = PurgeData2(rs(1))
			ClientContact = PurgeData2(rs(2))
			EXID = CheckNum(rs(3))
		end if
		CloseOBJs rs, Conn
		
		ChargeWH = -1
		DeliveryWH = -1
		if EXID <> 0  then
			OpenConn2 Conn
			Set rs = Conn.Execute("select id_almacen_recoleccion, id_almacen_entrega from routing_terrestre where id_routing=" & EXID)
			if Not rs.EOF then
				ChargeWH = SetWarehouseID(CheckNum(rs(0)))
				DeliveryWH = SetWarehouseID(CheckNum(rs(1)))
			end if
			CloseOBJs rs, Conn
		end if
	end if
%>
<html>
    <head>
        <title>AWB - Aimar - Administración</title>
    </head>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <script type="text/javascript" LANGUAGE="JavaScript">
    var ntr = "";
    var com = "";
    <%Select Case GroupID
    Case 2%>
	    //top.opener.document.forms[0].SenderData.value = '<%=Name%>\n<%=Address%>\n<%=Phone1%><%if Phone2 <> "" then%> / <%=Phone2%><%end if%><%if Attn <> "" then%>\nATTN: <%=Attn%><%end if%>';
	    //top.opener.document.forms[0].SenderID.value=<%=ObjectID%>;
	    //top.opener.document.forms[0].SenderAddrID.value=<%=AddressID%>;
    <%Case 3%>
	    top.opener.document.forms[0].ShipperData.value = '<%=Name%>\n<%=Address%>\n<%=Phone1%><%if Phone2 <> "" then%> / <%=Phone2%><%end if%><%if Attn <> "" then%>\nATTN: <%=Attn%><%end if%>';
	    top.opener.document.forms[0].ShipperID.value=<%=ObjectID%>;
	    top.opener.document.forms[0].ShipperAddrID.value=<%=AddressID%>;
        top.opener.document.forms[0].ShipperColoader.value=<%=ClientType%>;  
    <%Case 4%>

		if (top.opener.document.forma2) {

			if (top.opener.document.forma.ConsignerID)
				top.opener.document.forma.ConsignerID.value = <%=ObjectID%>;

            if (top.opener.document.forma.ConsignerData)
                top.opener.document.forma.ConsignerData.value = '<%=Name%>';

			//top.opener.document.forma.Action.value = 10;

			//top.opener.document.forma.submit();

            top.close();

			//return false;

		} else {

			if (top.opener.document.forms[0].ConsignerData) {
				if (top.opener.document.forms[0].Attn)
					top.opener.document.forms[0].ConsignerData.value = '<%=Name%>\n<%=Address%>\n<%=Phone1%><%if Phone2 <> "" then%> / <%=Phone2%><%end if%>';
				else
					top.opener.document.forms[0].ConsignerData.value = '<%=Name%>';
			}

			if (top.opener.document.forms[0].Attn)
				top.opener.document.forms[0].Attn.value = '<%if Attn <> "" then%><%=Name%>\nATTN: <%=Attn%><%end if%>';

			if (top.opener.document.forms[0].ConsignerID)
				top.opener.document.forms[0].ConsignerID.value =<%=ObjectID %>;

			if (top.opener.document.forms[0].ConsignerAddrID)
				top.opener.document.forms[0].ConsignerAddrID.value =<%=AddressID %>;

			if (top.opener.document.forms[0].ConsignerColoader)
				top.opener.document.forms[0].ConsignerColoader.value =<%=ClientType %>;

			if (top.opener.document.forms[0].CountryConsignee)
				top.opener.CountryConsignee = '<%=Countries%>';

			if (top.opener.document.forms[0].Consignee)
				top.opener.Consignee = '<%=Name%>';

			top.opener.document.forms[0].Action.value = 10;
			//submit para validar si el cliente es colgate 2020-10-10
			top.opener.document.forms[0].submit();
		}
    <%Case 11%>
	    top.opener.document.forms[0].ClientsID.value = '<%=ObjectID%>';
	    top.opener.document.forms[0].AddressesID.value = '<%=AddressID%>';
	    top.opener.document.forms[0].Clients.value = '<%=Name%>';
        top.opener.document.forms[0].ClientColoader.value=<%=ClientType%>;

        if (top.opener.document.forms[0].CodeReferenceValid) { //2020-09-15 cuando es ItineraryAdds.asp si seleccionan cliente realiza submit
            
            if (top.opener.document.forms[0].OID.value > 0) //2020-09-22 cuando es update y el cliente era de colgate debe reiniciar valores en ItineraryAdds.asp
                top.opener.document.forms[0].cambio.value=1;
            
            top.opener.document.forms[0].submit();
        }

    <%Case 14%>
	    top.opener.document.forms[0].ConsignerData.value = '<%=Name%>\n<%=Address%>\n<%=Phone1%><%if Phone2 <> "" then%> / <%=Phone2%><%end if%>';
	    top.opener.document.forms[0].Attn.value='<%=ClientContact%>';
	    top.opener.document.forms[0].ConsignerID.value=<%=ObjectID%>;
	    top.opener.document.forms[0].ConsignerAddrID.value=<%=AddressID%>;
	    top.opener.document.forms[0].ChargePlace.value='<%=PickUpData%>';
	    top.opener.document.forms[0].FinalDes.value='<%=DeliveryData%>';
	    top.opener.CountryConsignee='<%=Countries%>';
	    top.opener.Consignee='<%=Name%>';	
	    <%if ChargeWH <> 0 then%>
	    top.opener.document.forma.ChargeType.value ='<%=ChargeWH%>';
	    <%end if%>
	    <%if DeliveryWH <> 0  then%>
	    top.opener.document.forma.DestinyType.value ='<%=DeliveryWH%>';
	    <%end if%>
    <%Case 20%>
	    if (top.opener.document.forms[0].Agents.value != '') {ntr='\n'; com='|';} else {ntr=''; com='';};
	    top.opener.document.forms[0].AgentsID.value = top.opener.document.forms[0].AgentsID.value + com + '<%=ObjectID%>';
	    top.opener.document.forms[0].AgentsAddrID.value = top.opener.document.forms[0].AgentsAddrID.value + com + '<%=AddressID%>';
	    top.opener.document.forms[0].Agents.value = top.opener.document.forms[0].Agents.value + ntr + '<%=Name%>';
    <%Case 31%>
	    top.opener.document.forms[0].ShippersID.value = '<%=ObjectID%>';
	    top.opener.document.forms[0].ShippersAddrID.value = '<%=AddressID%>';
	    top.opener.document.forms[0].Shippers.value = '<%=Name%>';	
        top.opener.document.forms[0].ShipperColoader.value=<%=ClientType%>;
    <%Case 34,36%>
	    top.opener.document.forms[0].ColoadersID.value = '<%=ObjectID%>';
	    top.opener.document.forms[0].Coloaders.value = '<%=Name%>';
        <%If GroupID = 36 Then %>
            top.opener.document.forms[0].Consultar.disabled = false;
        <%Else %>
            top.opener.document.forms[0].ColoadersAddrID.value = '<%=AddressID%>';
        <%End If %>
    <%Case 35%>
	    top.opener.document.forms[0].ClientCollectID.value = '<%=ObjectID%>';
	    top.opener.document.forms[0].ClientsCollect.value = '<%=Name%>';
    <%End Select%>	 
	    top.close();
    </script>
    </body>
</html>
<%else%>
<script type="text/javascript" language="JavaScript">
	top.close();
</script>
<%end if%>		
