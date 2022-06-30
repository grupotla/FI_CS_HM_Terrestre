<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="Utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"

Dim TipoConta, SelectBodegas, ActividadComercial, CondicionPago, ObservacionesErp, iSelectBodegas, iActividadComercial, iCondicionPago, iObservacionesErp, facturar_a, facturar_a_nombre, Pedido_Erp, Pedido_Msg, RoutingID
Dim ArrAwbType, aTableValues, CountTableValues, aList6Values, CountList6Values, CantItems, Movimiento, homos, CountryExactus
Dim Msg, facturacion, ItemsPedidos, esquema, PedidoCliente, PedidoRubro
Dim ConsignerID, ConsignerData, Countries, AWBNumber, HBLNumber, BLType, CountryOrigen, CountriesFinalDes
Dim Conn, rs, QuerySelect, TipoCarga, ObjectID, AwbType, i, Currencies, result, iLink, orden, DUA          

AWBType = Request("AT")
TipoCarga = Request("TC")
ObjectID = Request("OID")
Movimiento = Iif(AWBType = "1","EXPORT","IMPORT")

ConsignerID = Request("ConsignerID")
ConsignerData = Request("ConsignerData")
PedidoCliente = Request("PedidoCliente")
PedidoRubro = Request("PedidoRubro")
orden = Request("orden")
DUA = Request("DUA")

facturar_a = ""
facturar_a_nombre = ""
CountTableValues = -1
CountList6Values = -1
RoutingID = 0
CantItems = 10




    OpenConn Conn

    '                       0           1           2          3           4           5           6               7               8       9           10      11              12              13          14              15                  16          17       18
    QuerySelect = "SELECT BLDetailID, CreatedDate, CreatedTime, BLs, HBLNumber, Countries, Clients,         Volumes,      Weights,      '',     BLType,  CountryOrigen, CountriesFinalDes,  '',         '',             '',             RoClientID, Agents, ClientsID FROM BLDetail WHERE BLDetailID=" & ObjectID
    'response.write QuerySelect & "<br>"
    Set rs = Conn.Execute(QuerySelect)
    If Not rs.EOF Then
        RoutingID = CheckNum(rs("RoClientID"))
        AWBNumber = rs("BLs")
        HBLNumber = rs("HBLNumber")
        Countries = rs("Countries")
        ConsignerID = rs("ClientsID")
        ConsignerData = rs("Clients")
        BLType = rs("BLType")	
        CountryOrigen = rs("CountryOrigen")	
        CountriesFinalDes = rs("CountriesFinalDes")	
    End If
    CloseOBJ rs


    if orden = "" then
        orden = "id_cliente DESC, ItemID"
    end if

    '                       0       1           2       3       4       5       6           7           8           9           10          11      12      13  14  15      16      17              18                      19                          20                          21                          22                  23                          24        25  26 27  28
    QuerySelect = "Select UserID, ItemName, ItemID, Currency, Value, OverSold, Local, PrepaidCollect, ServiceID, ServiceName, InvoiceID, CalcInBL, DocType, '', '', 0, ChargeID, '', COALESCE(id_cliente,0), COALESCE(id_pedido,0), TRIM(COALESCE(pedido_erp,'')), COALESCE(cliente_nombre,''), COALESCE(Regimen,''), COALESCE(TarifaPricing,''), COALESCE(TarifaTipo,''), 0, 0, '', '' FROM ChargeItems WHERE Expired=0 and SBLID=" & ObjectID & " ORDER BY " & orden 
    'response.write QuerySelect & "<br>"
    Set rs = Conn.Execute(QuerySelect)
    If Not rs.EOF Then
	    aTableValues = rs.GetRows
	    CountTableValues = rs.RecordCount - 1
    End If
    CloseOBJ rs

    CountryExactus = Session("OperatorCountry")
    
    QuerySelect = "select 'EXPORT' as tipo, 0 as Intransit from BLDetail where BLDetailID = " & ObjectID & " and substring(Countries,1,2) in ('" & left(CountryExactus,2) & "') and substring(CountriesFinalDes,1,2)  not in ('" & left(CountryExactus,2) & "') and BLType in (0,1) " & _
    "union " & _
    "select 'IMPORT' as tipo, Intransit from BLDetail where BLDetailID = " & ObjectID & " and substring(CountriesFinalDes,1,2)  in ('" & left(CountryExactus,2) & "') and BLType in (0,1)  " & _
    "union  " & _
    "select 'IMPORT' as tipo, 0 as Intransit from BLDetail where BLDetailID = " & ObjectID & " and substring(Countries,1,2)  in ('" & left(CountryExactus,2) & "') and substring(CountriesFinalDes,1,2)  in ('" & left(CountryExactus,2) & "') and BLType in (2) " 
    'response.write QuerySelect & "<br>"
	Set rs = Conn.Execute(QuerySelect)
	if Not rs.EOF then
        Movimiento = rs(0)
        if CheckNum(rs(1)) = 1 and Movimiento = "IMPORT" then
            Movimiento = ""
            response.write "<p><font face=verdana color=red >Esta carta porte es IMPORT pero aun esta en transito, Debe ir a ""Salidas y Llegadas"", ingresar fecha de llegada a CP master, de igual forma, esta master debe estar cerrada en pais creacion.</font></p>"
        end if
	end if

    CloseOBJs rs, Conn

OpenConn2 Conn

    QuerySelect = "SELECT pais_iso, CASE WHEN COALESCE(vencimiento - CURRENT_DATE,0) > 0 THEN '1' ELSE '0' END FROM empresas WHERE activo = 't' AND pais_iso IN ('" & CountriesFinalDes & "','" & CountryOrigen & "')" 
    'response.write QuerySelect & "<br>"
	Set rs = Conn.Execute(QuerySelect)
	if Not rs.EOF then           
            
	    Do While Not rs.EOF
            if rs(0) = CountriesFinalDes and  Movimiento = "IMPORT"  then
                if rs(1) = "1" then
                    CountryExactus = CountriesFinalDes
                end if
            end if

            if rs(0) = CountryOrigen and  Movimiento = "EXPORT"  then
                if rs(1) = "1" then
                    CountryExactus = CountryOrigen
                end if
            end if
                    
            rs.MoveNext
	    Loop

	end if
    CloseOBJ rs

        'response.write AWBNumber & " " & HBLNumber & " " & Countries & "<br>"

    SelectBodegas = ""     
    CondicionPago = ""
    ActividadComercial = ""
    ObservacionesErp = ""
    Pedido_Erp = ""
    Pedido_Msg = ""
    esquema = ""

    On Error Resume Next

        '////////////// PARAMETROS DE LA EMPRESA A FACTURAR
		if Request("TipoConta") <> "" then
            TipoConta = Request("TipoConta")
        end if

        result = WsEvaluaPedidos(Iif(HBLNumber = "",AWBNumber,HBLNumber), ObjectID, "2", CountryExactus, Pedido_Msg)

        if CheckNum(result(0)) = 1 then
            Msg = result(1)
            Pedido_Erp = result(2)
            TipoConta = result(3)

            if ubound(result) > 4 then 
                esquema = result(4)
            end if

        else
            response.write "Verifique conexion a Pedidos<br>"  
        end if

        iSelectBodegas = Request("SelectBodegas")  
        iCondicionPago = Request("CondicionPago")
        iActividadComercial = Request("ActividadComercial")
        iObservacionesErp = Request("ObservacionesErp")

	    if TipoConta = "EXACTUS" then

            if Len(iSelectBodegas) = 0 then 
                iSelectBodegas = "BOSE"            
            end if 

            if Len(iCondicionPago) = 0 then 
                iCondicionPago = "00"            
            end if 

            if Len(iActividadComercial) = 0 then 
                iActividadComercial = "602001"            
            end if 

            '/////////////////////// LEE EL CATALOGO DE BODEGAS EXACTUS 
            result = WsExactusCatalogos("BODEGA", esquema, "1")
            SelectBodegas = result(1)

            '/////////////////////// LEE EL CATALOGO DE CONDICION_PAGO EXACTUS 
            result = WsExactusCatalogos("CONDICION_PAGO", esquema, "1")
            CondicionPago = result(1)

            '/////////////////////// LEE EL CATALOGO DE ACTIVIDAD_COMERCIAL 2021-08-09
            result = WsExactusCatalogos("ACTIVIDAD_COMERCIAL", esquema, "1")
            ActividadComercial = result(1)

            ObservacionesErp = "<span class=erpLab>OBSERVACIONES : </span><textarea name=ObservacionesErp id='Observaciones para facturacion' class=erpFil style='width:100%' rows=3>" & iObservacionesErp & "</textarea>"

        end if

    If Err.Number <> 0 Then

        'response.write "<br>WsExactusCatalogos Error : " & Err.Number & " - " & Err.description & "<br>"  
        response.write "Verifique conexion a Catalogos<br>" 

    end if

    if TipoConta = "BAW" then     
        response.write "<font family=verdana color=navy>Pais " & CountryExactus & " tiene Contabilidad BAW</font><br>" 
	end if

    if TipoConta = "" then     
        response.write "<font family=verdana color=navy>No hay Tipo Conta definida para : " & CountryExactus & "</font><br>" 
	end if

    if esquema = ""  then
        response.write "<font family=verdana color=navy>No hay Esquema definido para : " & CountryExactus & "</font><br>" 
    end if





    'OpenConn2 Conn

    if CheckNum(RoutingID) > 0 then

        '                       0           1               2           3           4                       5                           6
        QuerySelect = "SELECT seguro, poliza_seguro, routing_seg, routing_adu, routing_ter, COALESCE(a.id_facturar,0), COALESCE(b.nombre_cliente,''), routing FROM routings a LEFT JOIN clientes b ON b.id_cliente = a.id_facturar WHERE a.id_routing = " & RoutingID
        'response.write QuerySelect & "<br>"
        Set rs = Conn.Execute(QuerySelect)
        if Not rs.EOF then
            'Seguro = rs(0)
		    'routing_seg = rs(2)
            'routing_adu = rs(3)
            'routing_ter = rs(4)
            RoutingID = rs(7)
            ConsignerID = rs(5)
            ConsignerData = rs(6)        
        end if
        CloseOBJ rs

    else

        QuerySelect = "SELECT b.nombre_cliente FROM clientes b WHERE b.id_cliente = " & ConsignerID
        'response.write QuerySelect & "<br>"
        Set rs = Conn.Execute(QuerySelect)
        if Not rs.EOF then
            ConsignerData = rs("nombre_cliente")
        end if
        CloseOBJ rs

    end if


    QuerySelect = "select distinct simbolo from monedas where pais in ('" & CountryExactus & "') order by simbolo"
	Set rs = Conn.Execute(QuerySelect)
    'response.write QuerySelect & "<br>"
    Currencies = ""
    if Not rs.EOF then
	    Do While Not rs.EOF
		    Currencies = Currencies & "<option value=" & rs("simbolo") & ">" & rs("simbolo") & "</option>"
		    rs.MoveNext
	    Loop
    end if

    CloseOBJ rs

    'response.write "Finalizo<br>"
    'response.end    


Dim FacID, FacType, FacStatus

ItemsPedidos = 0
homos = ""

for i=0 to CountTableValues

    if homos <> "" then
        homos = homos & ","
    end if

    homos = homos & "'" & aTableValues(8,i) & "-" & aTableValues(2,i) & "-" & aTableValues(22,i) & "'"

	FacID = CheckNum(aTableValues(10,i))    'InvoiceID        
    FacType = CheckNum(aTableValues(12,i))  'DocType
    FacStatus = 0

    aTableValues(17,i) = FacID

    if (FacID = 0 OR FacType = 9) then 'si no tiene factura ó tipo doc es pedido
        ItemsPedidos = ItemsPedidos + 1
    end if

    if FacID<>0 then


        Select case FacType

        case 9,10

            aTableValues(10,i) = 0

            '                                                           0                           1                               2                           3                            4                               5                      6
            QuerySelect = "SELECT DISTINCT TRIM(COALESCE(a.pedido_erp,'')) as a, COALESCE(a.estado,0) as b, TRIM(COALESCE(b.fc_numero,'')) as c, COALESCE(b.fc_estado,0) as d, COALESCE(b.fc_saldo,0) as e, COALESCE(c.nc_numero,'') as f, replace(regexp_replace(COALESCE(pedido,''), E'<[^>]+>', '', 'gi'),'&#124;','&#124;') as g FROM exactus_pedidos a LEFT JOIN exactus_pedidos_fc b ON a.id_pedido = b.id_pedido  LEFT JOIN exactus_pedidos_nc c ON a.id_pedido = c.id_pedido WHERE a.id_pedido = " & aTableValues(19,i) & " "
            'response.write QuerySelect & "<br>"
            set rs = Conn.Execute(QuerySelect)
			if Not rs.EOF then
            
                FacStatus = 90 

                if rs("a") <> "" then 'pedido_erp
                    'aTableValues(13,i) = FacID & " - " & rs("a")                     
                end if

                if rs("c") <> "" then
                    aTableValues(10,i) = CheckNum(FacID)
                    aTableValues(13,i) = FacID & " - " & rs("c") 
                    'FacStatus = 91 'facturada
                end if


                select case rs("b") 'estado
                    case "1"
                        aTableValues(14,i) = "<font color=red>ERROR 1</font>"
                    case "2"
                        aTableValues(14,i) = "<font color=red>ERROR 2</font>"
                    case "3"
                        aTableValues(14,i) = "<font color=green>ENVIADO</font>"
                    case "4"
                        aTableValues(14,i) = "<font color=blue>FACTURADO</font>"
                    case "5"
                        aTableValues(14,i) = "<font color=gray>INACTIVO</font>"
                end select

               if aTableValues(13,i) = "" then
                    aTableValues(13,i) = rs("g") 
               end if

            end if
            CloseOBJ rs

        end select 



        select Case FacStatus
        case 2
            aTableValues(14,i) = "<font color=blue>ABONADO</font>"
        case 4
            aTableValues(14,i) = "<font color=blue>PAGADO</font>"

        case 90 '2021-08-06
            
        case Else
            aTableValues(14,i) = "<font color=red>PENDIENTE</font>"

        End Select

    end if

next

    'AND a.d2 = 'A' 
    QuerySelect = "SELECT a.codigo, COALESCE(eh_erp_codigo,'') as eh_erp_codigo, a.id_servicio, a.id_rubro, a.d3 FROM vw_rubros_combinaciones a " & _
    "LEFT JOIN exactus_homologaciones ON codigo = eh_codigo AND eh_erp_categoria = '06' AND eh_estado = 1 AND eh_erp_esquema = '" & esquema & "' " & _ 
    "WHERE a.d1 = '" & Movimiento & "' AND a.id_servicio || '-' || a.id_rubro || '-' || a.d3 IN (" & homos & ") " & _ 
    "LIMIT 100"
    'response.write QuerySelect & "<br>"

    if homos <> "" then 
            
    Set rs = Conn.Execute(QuerySelect)
	Do While Not rs.EOF

        for i=0 to CountTableValues

            if aTableValues(8,i) & "-" & aTableValues(2,i) & "-" & aTableValues(22,i) = rs("id_servicio") & "-" & rs("id_rubro") & "-" & rs("d3") then

                aTableValues(27,i) = rs("codigo")
                aTableValues(28,i) = rs("eh_erp_codigo")

            end if

        next
	
        rs.MoveNext
	Loop
    CloseOBJ rs

    end if

CloseOBJ Conn

    'response.write "Finalizo<br>"
    'response.end 
%>
<html>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<LINK REL="stylesheet" type="text/css" HREF="img/estilos.css">
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>




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
.style4 {
	font-size:10px;
	color: #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
}

.readonly { border:0px;
            background:silver;
            color:navy;
            font-size: 10px; 
            font-family: Verdana, Arial, Helvetica, sans-serif;  
            width:auto; }

.ids    {   border:0px;
            color:navy;
            font-weight:normal;
            background:silver;
            font-size: 8px;
            width:auto; 
            }


-->


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

.erpLab {
    color:white;background-color:gray;height:20px;display:block;padding:2px;
}

.erpFil {
    background-color:rgb(255,232,159);
}

     
</style>
<body onbeforeunload="return check()">


<div id="myProgress">
  <div id="myBar">10%</div>
</div>


<form name="forma2" action="ItineraryChargesSave.asp" method="post" target="operations" onsubmit="move();">

	<INPUT name="ItemServIDs" type=hidden value="">
	<INPUT name="ItemServNames" type=hidden value="">
	<INPUT name="ItemNames" type=hidden value="">
	<INPUT name="ItemIDs" type=hidden value="">
	<INPUT name="ItemCurrs" type=hidden value="">
	<INPUT name="ItemVals" type=hidden value="">
	<INPUT name="ItemOVals" type=hidden value="">
	<INPUT name="ItemLocs" type=hidden value="">
	<INPUT name="ItemPPCCs" type=hidden value="">
	<INPUT name="ItemInvoices" type=hidden value="-1">
	<INPUT name="ItemDocType" type=hidden value="-1">
	<INPUT name="ItemCalcInBLs" type=hidden value="-1">
    <INPUT name="ItemInRO" type=hidden value="-1">
    <INPUT name="ItemChargeID" type=hidden value="">
    <INPUT name="ItemCli" type=hidden value="-1">
    <INPUT name="ItemCliNom" type=hidden value="">
    <INPUT name="ItemPedErp" type=hidden value="">
    <INPUT name="ItemRegimen" type=hidden value="">
    <INPUT name="ItemTarifaPrice" type=hidden value="">
    <INPUT name="ItemTarifaTipo" type=hidden value="">
    <INPUT name="ItemAgent" type=hidden value="">
	<INPUT name="CantItems" type=hidden value="-1">
	<INPUT name="Pos" type=hidden value="-1">

	<INPUT name="Action" type=hidden value=0>	
	<INPUT name="OID" type=hidden value="<%=ObjectID%>">
	<INPUT name="AT" type=hidden value="<%=AwbType%>">
	<INPUT name="esquema" type=hidden value="<%=esquema%>">

</form>


<form name="forma" action="ItineraryCharges.asp" method="post" onsubmit="move();">

	<INPUT name="Action"    type=hidden value=0>	
	<INPUT name="OID"       type=hidden value="<%=ObjectID%>">
	<INPUT name="AT"        type=hidden value="<%=AwbType%>">
    
	<INPUT name="Pedido_Erp"        type=hidden value="" <%'=Pedido_Erp%> >
	<INPUT name="esquema"           type=hidden value="<%=esquema%>">			   	
    <INPUT name="PedidoCliente"     type=hidden value="">
    <INPUT name="PedidoRubro"       type=hidden value="">
	<INPUT name="CountryExactus"         type=hidden value="<%=CountryExactus%>">
	<INPUT name="AWBNumber"         type=hidden value="<%=AWBNumber%>">
	<INPUT name="HBLNumber"        type=hidden value="<%=HBLNumber%>">
	<INPUT name="Movimiento"        type=hidden value="<%=Movimiento%>">
	<INPUT name="TipoCarga"         type=hidden value="<%=TipoCarga%>">
	<INPUT name="Main"              type=hidden value="1">
	<INPUT name="orden"         type=hidden value="<%=orden%>">
		   	



<table width="96%" border="0" cellpadding="2" cellspacing="0" align="center">

  <!---- ////////////////////////////////////////TIPO CARGO - BUSQUEDA DE CLIENTES///////////////////////////////////////////////////////////  --->
  <tr>
    <td colspan="4" align="center"  valign="top" style="border:0px" cellpadding=0 cellspacing=0>

  
        <table width="100%" border="0" align="left">
		<TR>
            <TH class=label align=right>SBLID&nbsp;:&nbsp;</TH>
            <TD class=label align=left><%=ObjectID%></TD>
        </TR> 
		<TR>
            <TH class=label align=right>MBLNumber&nbsp;:&nbsp;</TH>
            <TD class=label align=left><%=AWBNumber%></TD>
        </TR> 
		<TR>
            <TH class=label align=right>HBLNumber&nbsp;:&nbsp;</TH>
            <TD class=label align=left><%=HBLNumber%></TD>
        </TR> 
		<TR>
            <TH class=label align=right>Countries&nbsp;:&nbsp;</TH>
            <TD class=label align=left><%=Iif(CountryExactus = Countries, "<font style='background:silver'>" & Countries & "</font>", Countries)%></TD>
        </TR> 
		<TR>
            <TH class=label align=right>Esquema&nbsp;:&nbsp;</TH>
            <TD class=label align=left><%=esquema%></TD>
        </TR> 
		
		<TR>
            <TH class=label align=right>CountryOrigen&nbsp;:&nbsp;</TH>
            <TD class=label align=left><%=Iif(CountryExactus = CountryOrigen, "<font style='background:silver'>" & CountryOrigen & "</font>", CountryOrigen)%></TD>
        </TR> 
        </table>	


    </td>


    <td colspan="4" align="center" valign="top" style="border:0px" cellpadding=0 cellspacing=0>

  
        <table width="100%" border="0" align="left">
		<TR>
            <TH class=label align=right>ConsignerID&nbsp;:&nbsp;</TH>
            <TD class=label align=left><%=ConsignerID%></TD>
        </TR> 
		<TR>
            <TH class=label align=right>ConsignerData&nbsp;:&nbsp;</TH>
            <TD class=label align=left width="300px"><%=ConsignerData%></TD>
        </TR> 

		<TR>
            <TH class=label align=right>Movimiento&nbsp;:&nbsp;</TH>
            <TD class=label align=left><%=Movimiento%></TD>
        </TR> 

		<TR>
            <TH class=label align=right>BLType&nbsp;:&nbsp;</TH>
            <TD class=label align=left><%=BLType%></TD>
        </TR> 



		<TR>
            <TH class=label align=right>Tipo Carga&nbsp;:&nbsp;</TH>
            <TD class=label align=left>
                <input type="text" name="TC" readonly value="<%=TipoCarga%>"/>
                <!--
                    <select class="style10" name="TC" id="Select1" disabled>
                    <option value="-1">---</option> 
			    </select>	
                    -->
			    <%
                'OpenConn3 Conn
                'QuerySelect = "SELECT ""tpt_codigo"", ""tpt_descripcion"", ""tpt_pk"" FROM ""ti_pricing_tipo"" WHERE ""tpt_tipo"" = 'TIPO_CARGA' AND ""tpt_tps_fk"" = '1' ORDER BY ""tpt_descripcion"""
	            'Set rs = Conn.Execute(QuerySelect)
	            'Do While Not rs.EOF
		        '    response.write "<option value=" & rs("tpt_codigo") & ">" & rs(0) & "</option>"
		        '    rs.MoveNext
	            'Loop
                'CloseOBJs rs, Conn                        
                %>
            </TD>
        </TR> 
		<TR>
            <TH class=label align=right>CountriesFinalDes&nbsp;:&nbsp;</TH>
            <TD class=label align=left><%=Iif(CountryExactus = CountriesFinalDes, "<font style='background:silver'>" & CountriesFinalDes & "</font>", CountriesFinalDes)%></TD>
        </TR> 
		</table>	


    </td>

    <td colspan="4"  valign="top" style="border:0px" cellpadding=0 cellspacing=0>

        <table width="100%" border="0" align="left">
		<TR>
            <TH class=label align=left>Dua&nbsp;:&nbsp;</TH>
        </TR> 
		<TR>
            <TD class=label align=left width="300px">                
                <input type="text" name="DUA" size="10" value ="<%=DUA%>" />
            </TD>
        </TR> 

        <TR>
            <TH class=label align=left>&nbsp;</TH>
        </TR> 

		<TR>
            <TH class=label align=left>Routing&nbsp;:&nbsp;</TH>
        </TR> 
		<TR>
            <TD class=label align=left nowrap><%=RoutingID%></TD>
        </TR> 

		</table>	

    </td>

    <td colspan="4"  valign="top" style="border:0px" cellpadding=0 cellspacing=0>

		    <table width="100%" border="0" align="left">

				<TR>
                    <TD class=label align=center colspan=3>
                        <table>
                        <TR>
                            <TH class=label align=right>Asignar Cliente:</TH>
                            <TD class=label align=center>
                            
                            <button onClick="Javascript:GetData(4);return (false);" title="Buscar Cliente"><img src="img/Search16.png" /></button>

                            </TD>
                            
                            <TD class=label align=center><INPUT name="ConsignerID" size="10" type=text readonly value="<%=ConsignerID%>"></TD>
                            <td>
                                <button onClick="Javascript:Asignar();return (false);" title="Asignar Cliente"><img src="img/dispatch_console.gif" /></button>
                            </td>
                        </TR> 
                        </table>
                    </TD>
                </TR> 

			    <TR>
                    <TD class=label align=center colspan=3>                      
                        <textarea class="style10" name="ConsignerData" rows="4" cols="40" id="Consignee / Consignatario"  readonly="readonly"><%=ConsignerData%></textarea>	                      
                    </TD>
                </TR> 

		    </table>
                       
            	
    </td>
    <td colspan="4"  valign="top" style="border:0px" cellpadding=0 cellspacing=0>

            <!-- //////////////ACCIONES PEDIDOS/////////////////////// --->
            <table align=center>
                <tr>
                    <td colspan=3>
                        <label class=label><b>Cliente:</b></label><br />
                        <select name='CmbClientes' id='CmbClientes' style='width:200px'></select><br />
                    </td>
                </tr>
                <tr>
                <td align=center>
                    <button onClick="Javascript:ClienteUpdate();return (false);" title="Transmitir Rubros"><img src="img/inter.gif" /></button><br />
                    <label class=label>Transmitir</label><br />
                </td>
                <td align=center>
                    <button onClick="Javascript:ClienteFree();return (false);" title="Liberar Rubros"><img src="img/remove.gif" /></button><br />                    
                    <label class=label>Liberar</label><br />
                </td>
                <td align=center>
                    <button onClick="Javascript:Asignar2();return (false);" title="Asignar Cliente"><img src="img/dispatch_console.gif" /></button><br />
                    <label class=label>Asignar</label><br />
                </td>
                </tr>
            </table>

            <input  name='CmbRubrosTra' id='CmbRubrosTra' type="hidden" size="50" />
            <input  name='CmbRubrosLib' id='CmbRubrosLib' type="hidden" size="50" />
            <input  name='CmbRubrosBlock' id='CmbRubrosBlock' type="hidden" size="50" />
            <input  name='CmbRubrosCon' id='CmbRubrosCon' type="hidden" size="50" />
            
            <% 
            'iLink = "GID=0&ObjectID=" & ObjectID & "&DocTyp=" & Iif(Movimiento = "EXPORT", 0, 1) & "&HBLNumber=" & HBLNumber & "&AWBNumber=" & AWBNumber & "&esquema=" & esquema
            %>
            <center>
                <!--
                <INPUT type=button onClick="Javascript:window.open('Awb-Facturacion.asp?<%=iLink%>','AWBData','height=400,width=1100,menubar=0,resizable=1,scrollbars=1,toolbar=0');return false;"" value="&nbsp;&nbsp;Articulos / Pedidos / Facturas&nbsp;&nbsp;" class="Boton cRed">
                -->
                <input type="button" onclick="move();opener.location.reload();" value="Refresh" />


            </center>

    </td>
  </tr>
  </table>

    <div style="border:0px;height:300px;overflow:auto;border:1px solid orange">

        <table width="100%" border="0" cellpadding="2" cellspacing="0" align="center">
        <tr>
        <td class="style4" align="center" colspan="16">Cargos</td>
        </tr>
        <tr>
	        <td align="center" class="menu activeMain"></td>
	        <td align="center" class="menu activeMain"></td>
            <td align="center" class="menu activeMain"><a href="#" class="menu activeMain" onclick="return Ordenamiento(1);">Cliente</a></td>
	        <td align="center" class="menu activeMain"><a href="#" class="menu activeMain" onclick="return Ordenamiento(2);">AgentType</a></td>
	        <td align="center" class="menu activeMain">Int/Loc</td>
	        <td align="center" class="menu activeMain">Pos</td>
	        <td align="center" class="menu activeMain">Servicio</td>
	        <td align="center" class="menu activeMain">Rubro</td>
	        <td align="center" class="menu activeMain"></td>
	        <td align="center" class="menu activeMain">Tarifa</td>
	        <!-- <td align="center" class="menu activeMain">Tarifa Tipo</td> -->
	        <!-- <td align="center" class="menu activeMain">Regimen</td> -->
	        <td align="center" class="menu activeMain">Moneda</td>
	        <td align="center" class="menu activeMain">Monto</td>
	        <td align="center" class="menu activeMain">CC/PP</td>
	        <td align="center" class="menu activeMain"></td>
	        <td align="center" class="menu activeMain">Pedido/Factura</td>
	        <td align="center" class="menu activeMain">Estado</td>
	        <td align="center" class="menu activeMain">Homologado</td>
        </tr>


        <%for i=0 to CantItems %>
		  
          <tr id="Row<%=i%>">

			<td align="center" class="style4" nowrap>

				<input type="checkbox"  id="CHK<%=i%>" name="CHK" value="0">

                <% if i <= CountTableValues then  %>
                <input type="hidden"      id="CID<%=i%>" name="CID<%=i%>" size="2" readonly>
                <input type="hidden"      id="PID<%=i%>" name="PID<%=i%>" size="1" readonly>
                <input type="hidden"      id="PER<%=i%>" name="PER<%=i%>" size="1" readonly>
                <% end if  %>

			</td>

            <td align="right" class="style4" nowrap>

                <% if i <= CountTableValues then  %>
                <button onClick="Javascript:LiberarRubro(<%=i%>);return (false);" id="CR1<%=i%>" title="Liberar de Pedido"><img src="img/remove.gif" /></button>
                <% end if  %>

                <% if i <= CountTableValues then  %>
			    <button onClick="Javascript:DelCharge(<%=i%>);return (false);" id="DE<%=i%>" title="Borrar Rubro"><img src="img/delete.gif" /></button>
                <% end if  %>

			</td>
            <td align="left" class="style4" nowrap>
                <!-- <div id="CLI1_<%=i%>"></div> -->
				<input type="text" size="15" class="style10" name="CLI1_<%=i%>" readonly>
				<input type="hidden" size="5" class="style10" name="CLI<%=i%>" readonly>
				<input type="hidden" size="15" class="style10" name="CNO<%=i%>" readonly>
			</td>
			<td align="right" class="style4">      <!-- cargos tipos --> 
				<select class='style10' name='CT<%=i%>' id="Tipo de Cargo" style="width:50px">
				<option value='-1'>---</option>
				<option value='0'>TRANSPORTISTA</option>
				<option value='1'>AGENTE</option>
                <% 'if AWBType <> 1 then %>
				<!-- <option value='2'>OTROS</option> -->
                <% 'end if %>
			 	</select>
			</td>
			<td align="right" class="style4">
				<select class='style10' name='T<%=i%>' id="Tipo de Cobro">
				<option value='-1'>---</option>
				<option value='0'>INT</option>
				<option value='1'>LOC</option>
			 	</select>
			</td>
			<td align="right" class="style4">   <!-- Pos -->
                <div id="POS<%=i%>"></div>
			</td>

			<td align="right" class="style4">   <!-- servicio -->
				<input type="text" size="15" class="style10" name="SVN<%=i%>" value="" readonly>
				<input type="hidden" name="SVI<%=i%>" value="">
			</td>
			<td align="right" class="style4">   <!-- rubro -->
				<input type="text" size="15" class="style10" name="N<%=i%>" value="" readonly>
				<input type="hidden" name="I<%=i%>" value="">
			</td>
			<td align="right" class="style4">
				<div id=DR<%=i%> style="VISIBILITY: visible;">
                <button onClick="Javascript:SearchCharge(<%=i%>);return (false);" title="Buscar Rubro"><img src="img/Search16.png" /></button>
				</div>
			</td>
			<td align="right" class="style4">
                <div id="TT1_<%=i%>"></div>
                <div id="TP1_<%=i%>"></div>
                
                <input type="hidden" size="8" class="style10" name="TP<%=i%>" value="" readonly>
                <input type="hidden" size="8" class="style10" name="R<%=i%>" value="" readonly>
                <input type="hidden" size="8" class="style10" name="TT<%=i%>" value="" readonly>
            </td>
              <!--
			<td align="right" class="style4"></td> -->
           <!--<td align="left" class="style4">
                <div id="R1_<%=i%>"></div>                
			</td>-->
			<td align="right" class="style4">
				<select class='style10' name='C<%=i%>' id="Moneda">
				<option value='-1'>---</option>
				<%=Currencies%>
				</select>
			</td>
			<td align="center" class="style4">   <!-- monto -->
				<input type="text" size="8" class="style10" name="V<%=i%>" value="" onKeyUp="res(this,numb);" id="Monto">
			</td>
			<td align="right" class="style4">
				<select class='style10' name='TC<%=i%>' id="Forma de Pago" style="width:50px">
				<option value='-1'>---</option>
				<option value='0'>PREPAID</option>
				<option value='1'>COLLECT</option>
			 	</select>
			</td>
            <td>
                <button onClick="Javascript:SaveCharge(<%=i%>);return (false);" id="SV<%=i%>" title="Grabar Rubro"><img src="img/floppy.gif" /></button>
            </td>
            <td align="right" class="style4">
				<input type="text" size="25" class="style10" name="FAC<%=i%>" value="" readonly>
			</td>

            <td align="left" class="style4">
				<div id="STATFAC<%=i%>" style="VISIBILITY: visible;white-space:nowrap"></div>
			</td>

			<input type="hidden" name="INV<%=i%>" value="0">
			<input type="hidden" name="DTY<%=i%>" value="0">
            <input type="hidden" name="IRO<%=i%>" value="0">
            <input type="hidden" name="ICID<%=i%>" value="0">

            <td align="left" class="style4">
                <input type="hidden" name="HOM<%=i%>" value="0">
				<div id="HOM_1<%=i%>" style="white-space:nowrap"></div>
			</td>

		  </tr>
		  <%next%>
		</table>

    </div>





    <table width="100%" border="0" cellpadding="2" cellspacing="0" align="center">
		<% if Pedido_Erp = "" then %>
		<TR>             
				<TD class=label align=right ><b>Pedido Abierto:</b></TD><TD class=label align=left ><input type="text" value="<%=Pedido_Erp%>" size="30" readonly style="background-color:silver"></TD>
				<td>
					<input name=enviar type=button onClick="JavaScript:Solicitar(4);"  value="&nbsp;&nbsp;Solicitar&nbsp;&nbsp;" class="Boton cBlue">
				</td>
		</TR>
		<% end if %>                 

		<TR><TD colspan=5><%=ObservacionesErp%></TD></TR>

		<TR>		 
				<TD class=label align=center>             
				<input type="hidden" name="TipoConta" value="<%=TipoConta%>"  />          
				<% if SelectBodegas <> "" then %>
				<span class=erpLab>BODEGA&nbsp;:&nbsp;</span><select name="SelectBodegas" id="Bodega" class=erpFil><%=SelectBodegas%></select>
				<% end if %>
				</TD>

				<TD class=label align=center>             
				<% if ActividadComercial <> "" then %>
				<span class=erpLab>ACTIVIDAD COMERCIAL&nbsp;:&nbsp;</span><select name="ActividadComercial" id="Actividad Comercial" class=erpFil><%=ActividadComercial%></select>
				<% end if %>
				</TD>

		</TR>

        <tr>

				<TD class=label align=center>             
				<% if CondicionPago <> "" then %>
				<span class=erpLab>CONDICION DE PAGO&nbsp;:&nbsp;</span><select name="CondicionPago" id="Condicion de Pago" class=erpFil><%=CondicionPago%></select>
				<% end if %>
				</TD>

				<TD class=label align=center>         
				<% if TipoConta = "EXACTUS" then %>
					 
				<% if Pedido_Erp = "" then %>
						
					<% if ItemsPedidos = 0 then %>
                    <!--
							<input name=enviar type=button onclick="alert('No hay rubros para Transmitir, Solicite Pedido Abierto')" value="&nbsp;&nbsp;Transmitir&nbsp;&nbsp;" class=label>
                            -->
					<% else %>
					<!--
							<input name=enviar type=button onclick="JavaScript:FacturarAbierto();" value="&nbsp;&nbsp;Transmitir Abierto&nbsp;&nbsp;" class=label>
				            -->
					<% end if %>
						
						
				<% else %>
                    <!--
						<input name=enviar type=button onClick="JavaScript:Facturar();" value="&nbsp;&nbsp;Transmitir Pedido&nbsp;&nbsp;" class="Boton cBlue">
                    -->
				<% end if %>

				<% end if %>         
				</TD>                
        </tr>

		<TR><TD colspan=4></TD></TR>

	</table>

</form>


    
</body>


<script type="text/javascript">

    window.addEventListener('beforeunload', function (e) {
        e.preventDefault();
        e.returnValue = '';
    });

    function check() {
        return "Are you sure you want to exit this page?";
    }

    function Ordenamiento(tipo) {

        var url = 'ItineraryCharges.asp?OID=<%=ObjectID%>&TC=<%=TipoCarga%>&AT=<%=AwbType%>';
        var orden = "";
        switch (tipo) {

            case 1: //clientes

                orden = "id_cliente DESC, AgentTyp, Pos, ItemID";

                break;

            case 2: //agent type

                orden = " AgentTyp, Pos, ItemID";

                break;
        }

        url += '&orden=' + orden;

        //alert(url);

        location.href = url;


        return false;
    }



    function move() {
        window.location = '#';
        //document.awb_frame.style.display = "none";
        document.getElementById('myBar').style.display = "block";
        var elem = document.getElementById("myBar");
        var width = 10;
        var id = setInterval(frame, 65);
        function frame() {
            if (width >= 100) {
                clearInterval(id);
                document.getElementById('myBar').style.display = "none";
            } else {
                width++;
                elem.style.width = width + '%';
                elem.innerHTML = width * 1 + '%';
            }
        }
    }

    function SearchCharge(ChargePos) {

        if (!valSelec(document.forma.elements["CT"+ChargePos])){return false};
        if (!valSelec(document.forma.elements["T" + ChargePos])) { return false };
        if (!valSelec(document.forma.elements["TC" + ChargePos])) { return false };

        var servicio = document.forma.elements['SVI' + ChargePos].value;
        var rubro = document.forma.elements['I' + ChargePos].value;        
        var iNo = '0'; //document.getElementById('No').value;
        var ChargeName = document.forma.elements['N' + ChargePos].value;
        var ChargeMoneda = document.forma.elements['C' + ChargePos].value;
        var CargoTipo = document.forma.elements['CT' + ChargePos].value;
        var Regimen = document.forma.elements['R' + ChargePos].value;
        var TC = document.forma.elements['TC' + ChargePos].value;
        var IL = document.forma.elements["T" + ChargePos].value;

        window.open('Search_Charges.asp?PG=1&GID=29&OID=' + <%=ObjectID %> + '&C=' + '<%=CountryExactus%>' + '&N=' + ChargePos + '&T=<%=BLType%>&IL=' + IL + '&CM=' + ChargeMoneda + '&No=' + iNo + '&ServiceID=' + servicio + '&ItemID=' + rubro + '&esquema=<%=esquema%>&impex=<%=Movimiento%>' + '&RegimenID=' + Regimen + '&TC=' + TC, 'BLData', 'height=300,width=450,menubar=0,resizable=1,scrollbars=1,toolbar=0');

        return false;
    }


    function DelCharge(Pos) {

        if (!ValidarCharge(Pos)) return false;

        if (!confirm('Confirme borrar rubro?')) return false;

        SubmitIframe('borrar');

        return false;
    }


	 function SaveCharge(Pos) {

         if (!document.forma.elements["CID" + Pos]) {

             document.forma.elements['CLI' + Pos].value = document.forma.elements['ConsignerID'].value;

             document.forma.elements['CNO' + Pos].value = document.forma.elements['ConsignerData'].value;

         }

         var text = 'Confirme Grabar rubro asignado a cliente ' + document.forma.elements['ConsignerID'].value + ' ' + document.forma.elements['ConsignerData'].value.substring(0,30) + ' ?';

         if (!ValidarCharge(Pos)) return false;

         if (!confirm(document.forma.elements["CID"+Pos] ? 'Confirme Actualizar ? ' : text)) return false;

        SubmitIframe(document.forma.elements["CID"+Pos] ? 'update' : 'insert');

		return false;	 
	 }


     function ValidarCharge(Pos){
     
        var sep = '';
		var CantItems=-1;
  		document.forma2.ItemServIDs.value = "";
		document.forma2.ItemServNames.value = "";
  		document.forma2.ItemNames.value = "";
  		document.forma2.ItemIDs.value = "";
  		document.forma2.ItemCurrs.value = "";
  		document.forma2.ItemVals.value = "";
  		document.forma2.ItemOVals.value = "";
  		document.forma2.ItemLocs.value = "";
  		document.forma2.ItemPPCCs.value = "";
		document.forma2.ItemInvoices.value = "";
		document.forma2.ItemDocType.value = "";
		document.forma2.ItemCalcInBLs.value = "";
        document.forma2.ItemInRO.value = "";
        document.forma2.ItemChargeID.value = "";
        document.forma2.ItemCli.value = "";
        document.forma2.ItemCliNom.value = "";
		document.forma2.ItemPedErp.value = "";
		document.forma2.ItemRegimen.value = "";
        document.forma2.ItemTarifaPrice.value = "";
        document.forma2.ItemTarifaTipo.value = "";
        document.forma2.ItemAgent.value = "";
		
        document.forma2.Pos.value = Pos;

		for (i=0; i<=<%=CantItems%>;i++) {

            if ((Pos > -1 && Pos == i) || (Pos == -1)) 
            {

				if (!valSelec(document.forma.elements["CT"+i])){return false};
				if (!valSelec(document.forma.elements["T"+i])){return false};

			    if (document.forma.elements["N"+i].value == '') {
                    alert('Aun no hay rubro');
                    return false
                }
				    
				if (!valSelec(document.forma.elements["C"+i])){return false};
				if (!valTxt(document.forma.elements["V"+i], 1, 5)){return false};
				if (!valSelec(document.forma.elements["TC"+i])){return false};
			
            }	    


            if (document.forma.elements["SVI"+i].value != "" && document.forma.elements["SVN"+i].value != "" && document.forma.elements["N"+i].value != "") {
                //if (!valSelec(document.forma.elements["CCBL"+i])){return false};
				//if (document.forma.elements["OV"+i].value == '') {document.forma.elements["OV"+i].value = 0};
				//document.forma2.ItemCalcInBLs.value = document.forma2.ItemCalcInBLs.value + sep + document.forma.elements["CCBL"+i].value;

				//if (document.forma.elements["SVI"+i].value!="") {
					document.forma2.ItemServIDs.value = document.forma2.ItemServIDs.value + sep + document.forma.elements["SVI"+i].value;
					document.forma2.ItemServNames.value = document.forma2.ItemServNames.value + sep + document.forma.elements["SVN"+i].value;
				//} else {
				//	document.forma2.ItemServIDs.value = "0" + sep + document.forma.elements["SVI"+i].value;
				//	document.forma2.ItemServNames.value = " " + sep + document.forma.elements["SVN"+i].value;
				//}

				document.forma2.ItemNames.value = document.forma2.ItemNames.value + sep + document.forma.elements["N"+i].value;
				document.forma2.ItemIDs.value = document.forma2.ItemIDs.value + sep + document.forma.elements["I"+i].value;
				document.forma2.ItemCurrs.value = document.forma2.ItemCurrs.value + sep + document.forma.elements["C"+i].value;
				document.forma2.ItemVals.value = document.forma2.ItemVals.value + sep + document.forma.elements["V"+i].value;				    
				document.forma2.ItemLocs.value = document.forma2.ItemLocs.value + sep + document.forma.elements["T"+i].value;
				document.forma2.ItemPPCCs.value = document.forma2.ItemPPCCs.value + sep + document.forma.elements["TC"+i].value;
				document.forma2.ItemInvoices.value = document.forma2.ItemInvoices.value + sep + document.forma.elements["INV"+i].value;
				document.forma2.ItemDocType.value = document.forma2.ItemDocType.value + sep + document.forma.elements["DTY"+i].value;
                document.forma2.ItemInRO.value = document.forma2.ItemInRO.value + sep + document.forma.elements["IRO"+i].value;                    
                document.forma2.ItemCli.value = document.forma2.ItemCli.value + sep + document.forma.elements["CLI"+i].value;
                document.forma2.ItemCliNom.value = document.forma2.ItemCliNom.value + sep + document.forma.elements["CNO"+i].value;
                    
                if (document.forma.elements["PER"+i])
                    document.forma2.ItemPedErp.value = document.forma2.ItemPedErp.value + sep + document.forma.elements["PER"+i].value;
                else
                    document.forma2.ItemPedErp.value = document.forma2.ItemPedErp.value + sep + '';
				    
                document.forma2.ItemRegimen.value = document.forma2.ItemRegimen.value + sep + document.forma.elements["R"+i].value;
                document.forma2.ItemTarifaPrice.value = document.forma2.ItemTarifaPrice.value + sep + document.forma.elements["TP"+i].value;
                document.forma2.ItemTarifaTipo.value = document.forma2.ItemTarifaTipo.value + sep + document.forma.elements["TT"+i].value;
                document.forma2.ItemAgent.value = document.forma2.ItemAgent.value + sep + document.forma.elements["CT"+i].value;

                if (document.forma.elements["CID"+i])
                    document.forma2.ItemChargeID.value = document.forma2.ItemChargeID.value + sep + document.forma.elements["CID"+i].value;
                else
                    document.forma2.ItemChargeID.value = document.forma2.ItemChargeID.value + sep + '0';

                CantItems++;
				sep = "|";

            }
		}
	    document.forma2.CantItems.value = CantItems;     
        
        return true;
     }


    function SubmitIframe(Action) {
        //move();        
        console.log('Ejecuta submit operations : ' + Action);
        document.forma2.Action.value = Action;
        //document.forma2.target = "operations";        
        //document.forma2.Action = "AwbSaves.asp";
        document.forma2.submit();     
    }


    /////////////////////////////////////////////      
    function GetData(GID) {
        window.open('Search_BLData.asp?GID=' + GID + '&BTP=<%=BLType%>', 'BLData', 'height=200,width=450,menubar=0,resizable=1,scrollbars=1,toolbar=0,status=1');
    }

    function Asignar2(){

        if (document.forma.CmbClientes.options[document.forma.CmbClientes.selectedIndex].value != '') {
            var temp1 = document.forma.CmbClientes.options[document.forma.CmbClientes.selectedIndex].text;
            var temp2 = temp1.split('-');

            document.forma.ConsignerID.value = temp2[0].trim();
            document.forma.ConsignerData.value = temp2[1].trim();
        }
	    return false;	 
	}

    function Asignar(){
     
        var forma = document.forma;
        var chk = forma.elements['CHK'];
        var v1 = 0; //valida que exista un check true
        var v2 = 0; //valida que no seleccione el mismo cliente
        var i;

        for (i = 0; i < chk.length; i++) {
            chk[i].value = '';
            if (chk[i].checked) {
                if (forma.elements['PID'+i].value == 0) {  
                    chk[i].value = forma.elements['CID'+i].value + '|' + forma.elements["PID"+i].value + '|' + forma.elements["PER"+i].value + '|' + forma.elements["ConsignerID"].value + '|' + forma.elements["ConsignerData"].value;
                    v1 = v1 + 1;
                    if (forma.elements['CLI'+i].value == forma.elements['ConsignerID'].value)
                        v2 = v2 + 1;
                }             
            }            
        }     
     
        if (v1 == 0) {
            alert('No se ha seleccionado ninguna casilla!');
            return false;        
        }

        if (v2 > 0) {
            alert('Intenta asignar mismo cliente!');
            return false;        
        }

        if (confirm('Confirmar Asignar Cliente?')) {
            //SubmitIframe('asignar');

            console.log('Ejecuta submit operations : ' + 'asignar');
            document.forma.Action.value = "asignar";
            document.forma.target = "operations";
            document.forma.action = "ItineraryChargesSave.asp";
            document.forma.submit();

        }

        return false
    } 





     function ClienteUpdate(){
	    ClienteProc(1);
    	return false;	 
	 }

     function ClienteFree(){
		ClienteProc(2);
        return false;	 
	 }


    function BlockedRubros(tipo, cliente_id, rubros) {

		//alert(cliente_id + ' ' + rubros);

        var combo = '', tmp = '', conteo, total = 0;

        switch (tipo) {
            case 0:
                combo = document.forma.CmbRubrosBlock.value.split('#*#');
                break;

            case 1:
                combo = document.forma.CmbRubrosTra.value.split('#*#');
                break;

            case 2:
                combo = document.forma.CmbRubrosLib.value.split('#*#');
                break;
        }

        //////////////// rubros no homologados se suman a los rubros existentes
        for (var i = 0; i < combo.length; i++) {

            var row = combo[i].split('|');

            if (cliente_id == row[0]) {

                tmp = row[1].slice(0, -1);

                if (rubros != '' && tmp != '') rubros += ",";

                rubros += tmp;

                break;
            }
        }

        alert(rubros);

        console.log('Rubros=' + rubros);

        if (tipo == 0) {

            ///////////////// cantidad de rubros totales por cliente
            conteo = document.forma.CmbRubrosCon.value.split('#*#');
            for (var i = 0; i < conteo.length; i++) {

                var row = conteo[i].split('|');

                if (cliente_id == row[0]) {

                    total = row[1];

                    break;
                }
            }

            conteo = "";

            if (rubros != "") {
                conteo = rubros.split(",");
            }

            console.log('Total=' + total + ' Bloqueados=' + conteo.length);

            if (conteo.length >= total) { //si h

                var r = HayPedido(cliente_id);

                if (r) {

                    console.log('Hay Pedido Sin Facturar!');

                } else {

                    rubros = "*";
                }

            }

        }

        //alert('Final:'+rubros);

        return rubros;

    }


    function HayPedido(cliente_id) {

        var forma = document.forma;
        var chk = forma.elements['CHK'];
        var i, j;

        j = false;
        for (i = 0; i < chk.length; i++) {

            if (forma.elements["CLI" + i].value == cliente_id) { //cliente

                if (forma.elements["PID" + i].value > 0 && forma.elements["PER" + i].value != "") { //hay pedido

                    if (forma.elements["DTY" + i].value != 10) { //si no esta facturado

                        j = true;
                        break;
                    }
                }
            }
        }

        return j;
    }


    function HayCheck(cliente_id) {

        var forma = document.forma;
        var chk = forma.elements['CHK'];
        var i, hay_check_t = false, hay_check_f = false, bloqued = "", continuar = false;

        for (i = 0; i < chk.length; i++) {

            if (forma.elements["CLI" + i].value == cliente_id) { //cliente

                continuar = false;

                if (chk[i].disabled) { //si esta deshabilitado

                    if (forma.elements["PID" + i].value > 0 && forma.elements["PER" + i].value != "") { //hay pedido
                        continuar = true;
                    }

                } else {

                    if (chk[i].checked == true) {

                        hay_check_t = true;

                    }
                }
                if (chk[i].checked == false) {
                    continuar = true;
                }

                if (continuar == true) {

                    hay_check_f = true;

                    if (bloqued != "")
                        bloqued += ",";

                    bloqued += forma.elements["CID" + i].value;
                }

            }
        }

        if (hay_check_t == true && hay_check_f == true) { //debe bloquear rubros no seleccionados

        } else {

            bloqued = "";

        }

        return bloqued;
    }


    function ClienteProc(tipo) {

        var select = document.getElementById('CmbClientes');
        var cliente_id = select.options[select.selectedIndex].value;
        var cliente_str = select.options[select.selectedIndex].text;

        if (cliente_id == '') {
            alert('Seleccione Cliente');
            return false;
        }

        var Bloquear = "";

        if (tipo == 1) { //si va transmitir solo un rubro

            Bloquear = HayCheck(cliente_id);

            //alert("Bloquear : " + Bloquear);
        }

        try {

            switch (tipo) {
                case 1: //transmitir rubros
                    if (Bloquear == "")
                        document.forma.ObservacionesErp.value = "Transmitir Rubros Cliente " + cliente_str;
                    else
                        document.forma.ObservacionesErp.value = "Transmitir Rubro(s) Seleccionado(s) Cliente " + cliente_str;
                    break;

                case 2: //liberar rubros
                    document.forma.ObservacionesErp.value = "Liberar Rubros Cliente " + cliente_str;
                    break;
            }

        } catch (err) {
            alert('Verifique conectividad con server ERP');
            return false;
        }

        console.clear();
        console.log('-----' + 1);
        var rubros = BlockedRubros(tipo, cliente_id, '');

        if (tipo == 2) { //liberar si no hay rubros para liberar no hace nada

            if (rubros == '') {
                alert('No hay rubros disponibles para liberar');
                return false;
            }

        }

        ///////////////////////// BLOCKED
        rubros = BlockedRubros(0, cliente_id, rubros);

        if (rubros == '*') { //si trae asterisco es porque limpio rubros
            switch (tipo) {
                case 1: //transmitir rubros
                    alert('No hay rubros disponibles para transmitir');
                    return false;
                    break;

                case 2: //liberar rubros
                    alert('No hay rubros disponibles para liberar');
                    return false;
                    break;
            }
        }

        if (Bloquear != "") {

            if (rubros != "")
                rubros += ",";

            rubros += Bloquear;  //cuando hay seleccionado un rubro, bloque los demas para no enviarlos

            alert("Rubros 2:" + rubros)
        }

        if (confirm('Confirme ' + document.forma.ObservacionesErp.value + ' ?')) {

            document.forma.PedidoCliente.value = cliente_id;

            document.forma.PedidoRubro.value = rubros;


            ////////////////////////////////// validar datos para pedido

            if (!valTxt(document.forma.elements["ObservacionesErp"], 1, 5)){return false};
            if (!valSelec(document.forma.elements["SelectBodegas"])){return false};
            if (!valSelec(document.forma.elements["ActividadComercial"])){return false};
            if (!valSelec(document.forma.elements["CondicionPago"])){return false};

            //if (confirm('Confirme Transmitir ?')) {

            document.forma.action = "ItineraryChargesPedidos.asp";

            //move();
            document.forma.Action.value = 5;
            document.forma.submit();

            document.forma.action = "ItineraryCharges.asp";

            //}

            document.forma.PedidoCliente.value = '';
            document.forma.PedidoRubro.value = '';
        }

        return false;
    }



    function LiberarRubro(Pos) {  //liberacion individual a seleccion 

        var temp, temp1;
        temp = document.forma.elements["CLI"+Pos]; 
        temp1 = temp.value
        document.forma.PedidoCliente.value = temp1;

        temp = document.forma.elements["CID"+Pos]; 
        temp1 = temp.value
        document.forma.PedidoRubro.value = temp1;

        if (document.forma.elements["PER"+Pos].value == '')
            document.forma.ObservacionesErp.value = "Liberar Rubro " + document.forma.PedidoRubro.value + ' Pedido CS ' + document.forma.elements["PID"+Pos].value + " Cliente " + document.forma.PedidoCliente.value;
        else
            document.forma.ObservacionesErp.value = "Liberar Rubro " + document.forma.PedidoRubro.value + ' Pedido ERP ' + document.forma.elements["PER"+Pos].value + " Cliente " + document.forma.PedidoCliente.value;



        ///////////////////////// BLOCKED
        var rubros = BlockedRubros(0, document.forma.PedidoCliente.value, document.forma.PedidoRubro.value);


        if (confirm('Confirme ' + document.forma.ObservacionesErp.value + ' ?')) {

            if (document.forma.elements["PER"+Pos].value == ''){        
            
                if (ProcessRubros(Pos)) {

                    console.log('Ejecuta submit operations : ' + 'liberar');
                    document.forma.Action.value = "liberar";
                    document.forma.target = "operations";
                    document.forma.action = "ItineraryChargesSave.asp";
                    document.forma.submit();

                }

            } else {
                //Facturar();

                ////////////////////////////////// validar datos para pedido

                if (!valTxt(document.forma.elements["ObservacionesErp"], 1, 5)){return false};
                if (!valSelec(document.forma.elements["SelectBodegas"])){return false};
                if (!valSelec(document.forma.elements["ActividadComercial"])){return false};
                if (!valSelec(document.forma.elements["CondicionPago"])){return false};

                document.forma.PedidoRubro.value = rubros;
                document.forma.Pedido_Erp.value = document.forma.elements["PER" + Pos].value;

                document.forma.action = "ItineraryChargesPedidos.asp";
                document.forma.Action.value = 5;
                document.forma.submit();

                document.forma.action = "ItineraryCharges.asp";
            }
        }

        document.forma.PedidoCliente.value = '';
        document.forma.PedidoRubro.value = '';
        return false;
    }






    function ProcessRubros(Pos) {

        // 998 delete
        // 997 liberar

        var forma = document.forma;
        var chk = forma.elements['CHK'];
        var v1 = 0;
        var i;

        for (i = 0; i < chk.length; i++) {

            if (!chk[i].disabled) {
                chk[i].checked = false;  //limpia todos los checkbox             
            }

            if (i == Pos) {

                if (chk[i].disabled) {
                    chk[i].disabled = false;  //libera para transportar
                }

                chk[i].checked = true;

                if (forma.elements["CID"+i])
                    chk[i].value = forma.elements["CID"+i].value;                           //  0
                else
                    chk[i].value = '';

                if (forma.elements["PID"+i])
                    chk[i].value = chk[i].value + '|' + forma.elements["PID"+i].value;      //  1
                else
                    chk[i].value = chk[i].value + '|' + '';

                if (forma.elements["PER"+i])
                    chk[i].value = chk[i].value + '|' + forma.elements["PER"+i].value;      //  2
                else
                    chk[i].value = chk[i].value + '|' + '';
                    
                if (forma.elements["CLI"+i])
                    chk[i].value = chk[i].value + '|' + forma.elements["CLI"+i].value;      //  3
                else
                    chk[i].value = chk[i].value + '|' + '';
                    
                if (forma.elements["CNO"+i])
                    chk[i].value = chk[i].value + '|' + forma.elements["CNO"+i].value;      //  4
                else
                    chk[i].value = chk[i].value + '|' + '';

                //alert(i + ' ' + chk[i].value);
            }

            if (forma.elements["PID"+Pos])
                if (forma.elements["PID"+Pos].value != '0')
                    v1 = v1 + 1;
        }

        if (chk[Pos])
            chk[Pos].value = chk[Pos].value + '|' + v1; //registro a ser liberado o borrado     //  5

        //SubmitSaves(Action);

        return true;

    }


    function AlertaIframe() {


        alert('Cliente no esta homologado');

    }


    function CheckDoble(RubID, ServiceID) {

        var forma = document.forma;
        var chk = forma.elements['CHK'];
        var v1 = 0;
        var i;


        for (i = 0; i < chk.length; i++) {



            alert(document.forma.elements["SVI" + i].value + ' ' +
                document.forma.elements["I" + i].value + ' ' +
                document.forma.elements["DTY" + i].value + ' ' +
                document.forma.elements["PID" + i].value + ' ' +
                document.forma.elements["PER" + i].value);


            if ((document.forma.elements["SVI" + i].value == ServiceID) && (document.forma.elements["I" + i].value == RubID)) {

                if (document.forma.elements["DTY" + i].value != 10) { //10 si esta facturado

                    if ((document.forma.elements["PID" + i].value != 0) && (document.forma.elements["PER" + i].value != "")) { //si ya tiene pedido

                    } else {

                        alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero o no ");

                        document.forma.elements["SVI" + i].value = '';
                        document.forma.elements["SVN" + i].value = '';

                        document.forma.elements["N" + i].value = '';
                        document.forma.elements["I" + i].value = '';

                        return (false);

                    }

                }

            }

        }

        return true;

    }



    selecciona('forma.SelectBodegas', '<%=iSelectBodegas%>');
    selecciona('forma.CondicionPago','<%=iCondicionPago%>');
    selecciona('forma.ActividadComercial','<%=iActividadComercial%>');
    //document.forma.ObservacionesErp.value = '<%=Request("ObservacionesErp")%>';            
    //selecciona('forma.TC','<%=TipoCarga%>');


    var tmp = '', rub_opt_con = '', con = 0, homo = false, cli_ant = 0, cli = 0, cli_nom = '', c = 0; bg = '', cli_opt = '', rub_opt_block = '', rub_opt_lib = '', rub_opt_tra = '', DocType = 0, pedido_cs = 0, pedido_erp = '', Invoice = 0, IRO = 0;

    cli_opt += "<option value=''>-- Seleccione --</option>";

<%for i=0 to CountTableValues%>

    cli = '<%=aTableValues(18,i)%>';
    cli_nom = '<%=Left(PurgeData(aTableValues(21,i)),30)%>';

    //alert(cli_nom);

    DocType = '<%=aTableValues(12,i)%>';
    pedido_cs = '<%=aTableValues(19,i)%>';
    pedido_erp = '<%=aTableValues(20,i)%>';
    Invoice = '<%=aTableValues(10,i)%>';
    IRO = '<%=aTableValues(15,i)%>';

    if (cli != cli_ant) {

        c++;

        if (cli_ant > 0) {
            if (rub_opt_con != '') rub_opt_con += '#*#';
            rub_opt_con += cli_ant + '|' + con;
            con = 0;
        }

        cli_ant = cli;

        if (cli != 0) 
            cli_opt += "<option value='" + cli + "'>" + cli + " - " +  cli_nom + "</option>";

        if (rub_opt_lib != '') rub_opt_lib += '#*#';
        rub_opt_lib += cli + '|';

        if (rub_opt_tra != '') rub_opt_tra += '#*#';
        rub_opt_tra += cli + '|';

        if (rub_opt_block != '') rub_opt_block += '#*#';
        rub_opt_block += cli + '|';
    }

    con++;

	document.getElementById('Row<%=i%>').style.background = ((c % 2 == 0) ? 'lightblue' : 'orange');
    document.forma['CLI1_<%=i%>'].value = '<%=aTableValues(16,i)%>' + ' ' + cli + ' ' + cli_nom;
    document.forma['CLI<%=i%>'].value = cli; //id_cliente
    document.forma['CNO<%=i%>'].value = cli_nom; //cliente_nombre

    document.forma['PID<%=i%>'].value = pedido_cs;
	document.forma['PER<%=i%>'].value = pedido_erp;
    document.forma['CID<%=i%>'].value = '<%=aTableValues(16,i)%>';     //ChargeID

    document.forma['R<%=i%>'].value = '<%=aTableValues(22,i)%>';       //Regimen
    document.forma['TP<%=i%>'].value = '<%=aTableValues(23,i)%>';       //TarifaPricing
    document.forma['TT<%=i%>'].value = '<%=aTableValues(24,i)%>';      //TarifaTipo

    //document.getElementById('R1_'+<%=i%>).innerHTML = '<%=split(aTableValues(22,i)," ")(0)%>';        //
    document.getElementById('TP1_<%=i%>').innerHTML = '<%=aTableValues(23,i)%>';       //
    document.getElementById('TT1_<%=i%>').innerHTML = '<%=split(aTableValues(24,i)," ")(0)%>';       //

    selecciona('forma.CT<%=i%>','<%=aTableValues(25,i)%>');

    document.getElementById('POS<%=i%>').innerHTML = '<%=aTableValues(26,i)%>';          //Pos

    document.forma['HOM<%=i%>'].value = '<%=aTableValues(28,i)%>';                            //homologado


    /////////////////////////////////////////////////////////////
    tmp = '<%=aTableValues(27,i)%>';        //cs - exactus

    if ('<%= aTableValues(28, i) %>' == '') {
        document.getElementById('HOM_1<%=i%>').style.color = 'red';
        homo = false;
        tmp += ' (' + 'NO HOMOLOGADO' + ')';     //exactus
    } else {
        document.getElementById('HOM_1<%=i%>').style.color = 'gray';
        tmp += ' (' + '<%=aTableValues(28,i)%>' + ')';     //exactus
        homo = true;
    }

    tmp += ' ' + '<%=aTableValues(22,i)%>';     //regimen

    document.getElementById('HOM_1<%=i%>').innerHTML = tmp;
    /////////////////////////////////////////////////////////////


    document.getElementById('CHK<%=i%>').title = '<%=aTableValues(16,i)%>' + ' CS:' + '<%=aTableValues(19,i)%>' + ' ERP:' + '<%=aTableValues(20,i)%>';


    //alert('DocType=' + DocType + ' ' + 'pedido_erp=' + pedido_erp + ' ' + 'pedido_cs=' + pedido_cs + ' ' + 'Invoice=' + Invoice + ' ' + 'IRO=' + IRO);


    if (DocType == 9 && homo) {

        rub_opt_lib += document.forma['CID<%=i%>'].value + ','; //ChargeID

    } else {

        if ((DocType == 9 || DocType == 0) && homo) {

            rub_opt_lib += document.forma['CID<%=i%>'].value + ','; //ChargeID

        }
    }


    if ((DocType == 9 || DocType == 0) && homo) {


    } else {

        rub_opt_block += document.forma['CID<%=i%>'].value + ','; //ChargeID
    }

    //if ((Invoice == 0 && DocType == 0 && pedido_erp != '') || pedido_cs == 0 || DocType == 10) {

    if (pedido_erp != '' && pedido_cs > 0) {

        //        rub_opt_lib += document.forma['CID<%=i%>'].value + ','; //ChargeID

        if (pedido_erp == '') { //no hay pedido_erp

            document.getElementById('CHK<%=i%>').checked = false;
            document.getElementById('CHK<%=i%>').value = '0';

        } else { //si tiene pedido_erp

            document.getElementById("DE<%=i%>").style.display = "none";    //delete
            document.getElementById("SV<%=i%>").style.display = "none";    //insert update

            document.getElementById('CHK<%=i%>').checked = true;
            document.getElementById('CHK<%=i%>').value = document.forma['CID<%=i%>'].value + ','; //ChargeID
            document.getElementById('CHK<%=i%>').disabled = true;
        }

    }


    if ((pedido_cs == 0 && pedido_erp == '') || DocType == 10) {

        document.getElementById("CR1<%=i%>").style.display = "none";   //free pedido erp   

    }

    document.forma['N<%=i%>'].value = '<%=aTableValues(1,i)%>';
    document.forma['I<%=i%>'].value = '<%=aTableValues(2,i)%>';
	selecciona('forma.C<%=i%>','<%=aTableValues(3,i)%>');
    document.forma['V<%=i%>'].value = '<%=aTableValues(4,i)%>';    //monto
	//document.forma['OV<%=i%>'].value = '<%=aTableValues(5,i)%>';
    selecciona('forma.T<%=i%>','<%=aTableValues(6,i)%>');
	selecciona('forma.TC<%=i%>','<%=aTableValues(7,i)%>');
	document.forma['SVI<%=i%>'].value = '<%=aTableValues(8,i)%>';
    document.forma['SVN<%=i%>'].value = '<%=aTableValues(9,i)%>';
	
	//document.forma.CCBL<%=i%>'].value = '<%=aTableValues(11,i)%>';    
    document.forma['FAC<%=i%>'].value = '<%=aTableValues(13,i)%>';
    document.forma['FAC<%=i%>'].title = '<%=aTableValues(13,i)%>';
    document.getElementById('STATFAC<%=i%>').innerHTML = '<%=aTableValues(14,i)%>';
    	
    document.forma['INV<%=i%>'].value = '<%=aTableValues(17,i)%>';
    document.forma['DTY<%=i%>'].value = '<%=aTableValues(12,i)%>';
    document.forma['IRO<%=i%>'].value = '<%=aTableValues(15,i)%>';
    document.forma['ICID<%=i%>'].value = '<%=aTableValues(16,i)%>';


    if (Invoice == 0 && IRO == 0) {
        document.forma['N<%=i%>'];
        document.forma['I<%=i%>'];
        document.forma['C<%=i%>'];
        document.forma['V<%=i%>'];
        //document.forma.OV<%=i%>;
        document.forma['T<%=i%>'];
        document.forma['TC<%=i%>'];
        document.forma['SVI<%=i%>'];
        document.forma['SVN<%=i%>'];
        //document.forma.CCBL<%=i%>;
        document.getElementById("DE<%=i%>");
        document.getElementById("DR<%=i%>");
    }

    if ((Invoice == 0 && IRO != 0) || Invoice != 0) {
        document.forma['N<%=i%>'].disabled = 'false';
        document.forma['I<%=i%>'].disabled = 'false';
        document.forma['C<%=i%>'].disabled = 'false';
        document.forma['V<%=i%>'].disabled = 'false';
        //document.forma.OV<%=i%>.disabled = 'false';
        document.forma['T<%=i%>'].disabled = 'false';
        document.forma['TC<%=i%>'].disabled = 'false';
        document.forma['SVI<%=i%>'].disabled = 'false';
        document.forma['SVN<%=i%>'].disabled = 'false';
        //document.forma.CCBL<%=i%>;
        document.getElementById("DE<%=i%>").style.display = "none";
        document.getElementById("DR<%=i%>").style.visibility = "hidden";
        document.forma['CLI<%=i%>'].disabled = 'false'; //id_cliente
        document.forma['CNO<%=i%>'].disabled = 'false'; //cliente_nombre
    }

<% next %>

    if (cli_ant > 0) {
        if (rub_opt_con != '') rub_opt_con += '#*#';
        rub_opt_con += cli_ant + '|' + con;
        con = 0;
    }

        //alert(rub_opt_con);

    document.getElementById('CmbClientes').innerHTML = cli_opt;
    document.getElementById('CmbRubrosLib').value = rub_opt_lib;
    document.getElementById('CmbRubrosTra').value = rub_opt_tra;
    document.getElementById('CmbRubrosBlock').value = rub_opt_block;
    document.getElementById('CmbRubrosCon').value = rub_opt_con;

<%for i=CountTableValues+1 to CantItems%>
    if (document.getElementById('CHK<%=i%>'))
        document.getElementById('CHK<%=i%>').style.display = 'none';                      
<%next%>

</script>

<%
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>




<iframe src="about:blank" height="150" width="99%" name="operations" border="1"> <!-- style='height:100px;border:0px'> -->
  <p>Su navegador no es compatible con iframes</p>
</iframe>
