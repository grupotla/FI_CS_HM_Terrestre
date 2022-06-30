<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
Checking "0|1|2"
Dim ObjectID, Conn, rs, Action, aTableValues, CountTableValues, aTableValues1, CountTableValues1, Currencies, GroupID, CountriesFinalDes, TipoConta, result, QuerySelect, SelectBodegas, ActividadComercial, CondicionPago, iSelectBodegas, iActividadComercial, iCondicionPago, ObservacionesErp, EstadoErp, FechaErp, CountryOrigen, Countries, RoClientID, facturar_a, facturar_a_nombre, Movimiento, CountryExactus, ClientsID, Pedido_Erp

Dim Name, Volume, Weight, Agent, HBLNumber, BL, i, FisBillID, FinBillID, CantItems
Dim Freight, Freight2, Insurance, Insurance2, AnotherChargesCollect, AnotherChargesPrepaid
Dim FacID, FacType, FacStatus, BLType, esquema, ConsignerID, ConsignerData, ItemsPedidos  
Dim PedidoCliente, PedidoRubro

ObjectID = CheckNum(Request("OID"))
GroupID = CheckNum(Request("GID"))
Action = CheckNum(Request("Action"))
esquema = Request("esquema")
ConsignerID = Request("ConsignerID")
ConsignerData = Request("ConsignerData")
PedidoCliente = Request("PedidoCliente")
PedidoRubro = Request("PedidoRubro")

CountTableValues = -1
CantItems = 30

OpenConn Conn

    QuerySelect = "select Clients, Volumes, Weights, Agents, HBLNumber, BLs, FisBillID, FinBillID, BLType, CountriesFinalDes, Countries, CountryOrigen, RoClientID, ClientsID from BLDetail where BLDetailID=" & ObjectID
    'response.write QuerySelect & "<br>"
	Set rs = Conn.Execute(QuerySelect)
	if Not rs.EOF then
		Name = rs(0)
		Volume = rs(1)
		Weight = rs(2)
		Agent = rs(3)
		HBLNumber = rs(4)
		BL = rs(5)
		FisBillID = rs(6)
		FinBillID = rs(7)	
        BLType = rs(8)	
        CountriesFinalDes = rs(9)	
        Countries = rs(10)	
        CountryOrigen = rs(11)	
        RoClientID = rs(12)
        ClientsID = rs(13)
	end if
	CloseOBJ rs

    CountryExactus = Session("OperatorCountry")

    'and Intransit IN (2) 

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

    CloseOBJ rs

    '                       0       1           2       3       4       5       6           7           8           9           10          11      12      13  14  15      16      17              18                      19                          20                          21                          22
    QuerySelect = "SELECT UserID, ItemName, ItemID, Currency, Value, OverSold, Local, PrepaidCollect, ServiceID, ServiceName, InvoiceID, CalcInBL, DocType, '', '', InRO, ChargeID, '', COALESCE(id_cliente,0), COALESCE(id_pedido,0), TRIM(COALESCE(pedido_erp,'')), COALESCE(cliente_nombre,''), COALESCE(Regimen,'') FROM ChargeItems WHERE Expired=0 AND SBLID=" & ObjectID & " AND InterProviderType<>5 AND InterChargeType<>2 " & _
    "ORDER BY COALESCE(id_cliente,0) DESC, COALESCE(pedido_erp,''), COALESCE(id_pedido,0) DESC, ChargeID"

    '"ORDER BY ChargeID, PrepaidCollect, Local, Currency, ServiceName, ItemName" 'InvoiceID Desc, 
    
    'response.write QuerySelect & "<br>"
	Set rs = Conn.Execute(QuerySelect)
	if Not rs.EOF then
		aTableValues = rs.GetRows
		CountTableValues = rs.RecordCount-1
	end if

CloseOBJs rs, Conn






dim ConnMaster

OpenConn2 ConnMaster

        'QuerySelect = "SELECT codigo, oficina_aimar FROM paises WHERE codigo IN ('" & CountriesFinalDes & "','" & CountryOrigen & "')" 
        QuerySelect = "SELECT pais_iso, CASE WHEN COALESCE(vencimiento - CURRENT_DATE,0) > 0 THEN '1' ELSE '0' END FROM empresas WHERE activo = 't' AND pais_iso IN ('" & CountriesFinalDes & "','" & CountryOrigen & "')" 
        'response.write QuerySelect & "<br>"
	    Set rs = ConnMaster.Execute(QuerySelect)
	    if Not rs.EOF then           
            
	        Do While Not rs.EOF
		            
                'response.write "(" & rs(1) & ")<br>"

                if rs(0) = CountriesFinalDes and  Movimiento = "IMPORT"  then
                    
                    if rs(1) = "1" then
                               
                        'response.write "1(SI)<br>"

                        CountryExactus = CountriesFinalDes
                                
                    end if
                    
                end if

                if rs(0) = CountryOrigen and  Movimiento = "EXPORT"  then
                    
                    if rs(1) = "1" then

                        'response.write "2(SI)<br>"
                    
                        CountryExactus = CountryOrigen

                    end if
                    
                end if
                    
                rs.MoveNext
	        Loop

	    end if

        'response.write "(" & CountryExactus & ")<br>"




	'Obteniendo Monedas
	Set rs = ConnMaster.Execute("select distinct simbolo from monedas order by simbolo")
	Do While Not rs.EOF
		Currencies = Currencies & "<option value=" & rs(0) & ">" & rs(0) & "</option>"
		rs.MoveNext
	Loop

    QuerySelect = "SELECT COALESCE(a.id_facturar,0), COALESCE(b.nombre_cliente,'') FROM routings a INNER JOIN clientes b ON b.id_cliente = a.id_facturar WHERE a.id_routing = " & RoClientID
	'response.write QuerySelect & "<br>"
    Set rs = ConnMaster.Execute(QuerySelect)
	if Not rs.EOF then           
        facturar_a = rs(0)
        facturar_a_nombre = rs(1)
    end if


    SelectBodegas = ""     
    CondicionPago = ""
    ActividadComercial = ""
    ObservacionesErp = ""
    EstadoErp = -1
    Pedido_Erp = ""
    esquema = ""

    On Error Resume Next

        '////////////// PARAMETROS DE LA EMPRESA A FACTURAR
        'TipoConta = "BAW"

        if Request("TipoConta") <> "" then

            TipoConta = Request("TipoConta")

        end if

        Dim Msg, Pedido_Msg

        result = WsEvaluaPedidos(HBLNumber, ObjectID, "2", CountryExactus, Pedido_Msg)

        if CheckNum(result(0)) = 1 then
            Msg = result(1)
            pedido_erp = result(2)
            TipoConta = result(3)

            if ubound(result) > 4 then 
                esquema = result(4)
            end if
        else
            response.write "Verifique conexion a Pedidos<br>"  
        end if

        'if TipoConta = "" then
        '   TipoConta = "BAW" 
        'end if

        
         if TipoConta = "EXACTUS" and Movimiento = "" then
            TipoConta = "" 
            response.write "<p><font face=verdana color=red >No se puede facturar esta carga en oficina " & CountryExactus & ".</font></p>"
         end if

        iSelectBodegas = Request("SelectBodegas")  
        iCondicionPago = Request("CondicionPago")
        iActividadComercial = Request("ActividadComercial")
       
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
            result = WsExactusCatalogos("BODEGA", "1")
            SelectBodegas = result(1)

            '/////////////////////// LEE EL CATALOGO DE CONDICION_PAGO EXACTUS 
            result = WsExactusCatalogos("CONDICION_PAGO", "1")
            CondicionPago = result(1)

            '/////////////////////// LEE EL CATALOGO DE ACTIVIDAD_COMERCIAL 2021-08-09
            result = WsExactusCatalogos("ACTIVIDAD_COMERCIAL", "1")
            ActividadComercial = result(1)

            'ActividadComercial = "<option value=''>- Seleccione -</option><option value='630901'>AGENCIAS ADUANALES</option><option value='602001'>SERVICIO TRANSPORTE CARGA</option><option value='701004'>ALQUILER DE EDIFICIOS</option>"
        
            ObservacionesErp = "<span class=""menu"" style=""display:block;color:White;padding:3px;"">OBSERVACIONES : </span><textarea name=ObservacionesErp id='Observaciones para facturacion' style='background-color:rgb(255,232,159);width:100%' rows=3>" & Request("ObservacionesErp") & "</textarea>"

        end if



    If Err.Number <> 0 Then
        'response.write "<br>WsExactusCatalogos Error : " & Err.Number & " - " & Err.description & "<br>"  
        response.write "Verifique conexion a Catalogos<br>"  
    end if

    'if esquema = ""  then
    '    esquema = "TRANSIT"
    'end if

    if TipoConta = "BAW" then     
        response.write "<font family=verdana color=navy>Pais " & CountryExactus & " tiene Contabilidad BAW</font>" 
	end if

    if TipoConta = "" then     
        response.write "<font family=verdana color=navy>No hay Tipo Conta definida para : " & CountryExactus & "</font><br>" 
	end if

    if esquema = ""  then
        response.write "<font family=verdana color=navy>No hay Esquema definido para : " & CountryExactus & "</font><br>" 
    end if

'    response.write "(" & TipoConta & ")(" & esquema & ")<br>" 


'CloseOBJs rs, ConnMaster

'OpenConn Conn

'CloseOBJs rs, Conn

'OpenConn2 ConnMaster



'Seleccion de Serie, Correlativo y Estado de Pago de facturas/ND del BAW
openConnBAW Conn


    'response.write "(CountTableValues=" & CountTableValues & ")"

ItemsPedidos = 0

for i=0 to CountTableValues
	FacID = CheckNum(aTableValues(10,i))    'InvoiceID        
    FacType = CheckNum(aTableValues(12,i))  'DocType
    FacStatus = 0

    aTableValues(17,i) = FacID

    'id_cliente  18
    'id_pedido   19
    'pedido_erp  20

    'response.write "-(" & i & ")(" & FacID & ")(" & FacType & ")<br>" 

    if (FacID = 0 OR FacType = 9) then 'si no tiene factura ó tipo doc es pedido
        ItemsPedidos = ItemsPedidos + 1
    end if


    if FacID<>0 then
	    Select case FacType
        case 1
            set rs = Conn.Execute("select tfa_serie, tfa_correlativo, tfa_ted_id from tbl_facturacion where tfa_id=" & FacID)
			    aTableValues(13,i) = "FC-" & rs(0) & "-" & rs(1)
                FacStatus = CheckNum(rs(2))
		    CloseOBJ rs

        case 4
            set rs = Conn.Execute("select tnd_serie, tnd_correlativo, tnd_ted_id from tbl_nota_debito where tnd_id=" & FacID)
			    aTableValues(13,i) = "ND-" & rs(0) & "-" & rs(1)
                FacStatus = CheckNum(rs(2))
		    CloseOBJ rs

        'case 0 'enviado por pedido exactus

            'set rs = ConnMaster.Execute("select a.pedido_erp, a.estado from exactus_pedidos a where a.id_pedido=" & FacID)
			'    aTableValues(13,i) = "PE-" & rs(0)
            '    FacStatus = CheckNum(rs(1))
		    'CloseOBJ rs

        case 9,10 'recibido por pedido exactus

            aTableValues(10,i) = 0

            '                                                           0                           1                               2                           3                            4                               5                      6
            QuerySelect = "SELECT DISTINCT TRIM(COALESCE(a.pedido_erp,'')) as a, COALESCE(a.estado,0) as b, TRIM(COALESCE(b.fc_numero,'')) as c, COALESCE(b.fc_estado,0) as d, COALESCE(b.fc_saldo,0) as e, COALESCE(c.nc_numero,'') as f, replace(regexp_replace(COALESCE(pedido,''), E'<[^>]+>', '', 'gi'),'&#124;','&#124;') as g FROM exactus_pedidos a LEFT JOIN exactus_pedidos_fc b ON a.id_pedido = b.id_pedido  LEFT JOIN exactus_pedidos_nc c ON a.id_pedido = c.id_pedido WHERE a.id_pedido = " & aTableValues(19,i) & " "
            'response.write QuerySelect & "<br>"
            set rs = ConnMaster.Execute(QuerySelect)
			if Not rs.EOF then

                FacStatus = 90 

                if rs("a") <> "" then
                    'aTableValues(13,i) = FacID & " - " & rs("a")                     
                end if

                if rs("c") <> "" then
                    aTableValues(10,i) = CheckNum(FacID)
                    aTableValues(13,i) = FacID & " - " & rs("c") 
                    'FacStatus = 91 'facturada
                end if


                select case rs("b") 'estado
                    case "1"
                        aTableValues(14,i) = "<font color=red>ERROR</font>"
                    case "2"
                        aTableValues(14,i) = "<font color=red>ERROR</font>"
                    case "3"
                        aTableValues(14,i) = "<font color=green>CORRECTO</font>"
                    case "4"
                        aTableValues(14,i) = "<font color=blue>FACTURADO</font>"
                    case "5"
                        aTableValues(14,i) = "<font color=gray>INACTIVO</font>"
                end select

                                
               'if  rs(5) <> "*" then
               '     aTableValues(13,i) = "NC-" & rs(5) 
               '     FacStatus = 92 'anulada
               'end if

               if aTableValues(13,i) = "" then
                    aTableValues(13,i) = rs("g") 
               end if



            end if
		    CloseOBJ rs

        end Select



    End If

        'Indicando el Estado de Pago de la Factura/ND
        select Case FacStatus
        case 2
            aTableValues(14,i) = "<font color=blue>ABONADO</font>"
        case 4
            aTableValues(14,i) = "<font color=blue>PAGADO</font>"

        case 90 '2021-08-06
            
            'i = 0
            'aTableValues(14,i) = "<font color=blue>ENVIADO</font>"

        'case 91 '2021-08-16
        '    aTableValues(14,i) = "<font color=blue>FACTURADO</font>"

        'case 92 '2021-08-06
        '    aTableValues(14,i) = "<font color=blue>CANCELADO</font>"

        'case 93
        '    aTableValues(14,i) = "<font color=red>REVISAR JSON</font>"

        case Else
            aTableValues(14,i) = "<font color=red>PENDIENTE</font>"

            'if aTableValues(13,i) = "" and aTableValues(20,i) <> "" then
                    'aTableValues(14,i) = "<font color=red>" & aTableValues(20,i) & "</font>"
            'end if


        End Select


next
CloseOBJ Conn
CloseOBJs ConnMaster
%>
<HTML>
<HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
    function move() {
        document.forma.style.display = "none";
        document.getElementById('myBar').style.display = "block";
        var elem = document.getElementById("myBar");
        var width = 10;
        var id = setInterval(frame, 45);
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

    function Solicitar(Action) {

        if (confirm('Solicitar Pedido Abierto a Exactus ?')) {
            move();
            document.forma.action = "ItineraryChargesPedidos.asp";
            document.forma.Action.value = 4;
            document.forma.submit();
        }
     }
	 
    
     function Facturar() {

        if (!validar(3,-1)) {        
            alert('Asegurese de presionar el boton Actualizar antes de Transmitir');
            return false;                  
        }

        if (!valTxt(document.forma.elements["ObservacionesErp"], 1, 5)){return false};
        if (!valSelec(document.forma.elements["SelectBodegas"])){return false};
        if (!valSelec(document.forma.elements["ActividadComercial"])){return false};
        if (!valSelec(document.forma.elements["CondicionPago"])){return false};

        if (confirm('Confirme Transmitir ?')) {

            document.forma.action = "ItineraryChargesPedidos.asp";

            move();
            document.forma.Action.value = 3;
            document.forma.submit();

            document.forma.action = "ItineraryCharges.asp";

        }
     }
	 

     
     
     function FacturarAbierto() {

        if (!validar(3,-1)) {        
            alert('Asegurese de presionar el boton Actualizar antes de Transmitir');
            return false;                  
        }

        if (!valTxt(document.forma.elements["ObservacionesErp"], 1, 5)){return false};
        if (!valSelec(document.forma.elements["SelectBodegas"])){return false};
        if (!valSelec(document.forma.elements["ActividadComercial"])){return false};
        if (!valSelec(document.forma.elements["CondicionPago"])){return false};

        if (confirm('Aun no tiene un no. de Pedido ERP, Desea Transmitir ?')) {

            document.forma.action = "ItineraryChargesPedidos.asp";

            move();
            document.forma.Action.value = 3;
            document.forma.submit();

            document.forma.action = "ItineraryCharges.asp";
        }
     }



	function validar(Action,Pos) {
		
        ColeccionRows(Pos);

        if (CantItems == -1) //si no hay nada
            return false;

        if (Action == 3) //si es facturacion                 
            return true;                        

        if (!confirm( Action == 3 ? 'Confirme Facturar ? ' : 'Confirme Actualizar datos'  )) //confirmar
	        return false;

	    //document.forma.Action.value = Action;
        //move();
	    //document.forma.submit();

        SubmitSaves(Action);


	 }


	 function SaveCharge(Pos) {

            var Action = 0;

			if (document.forma.elements["N"+Pos].value != '') {
				if (!valSelec(document.forma.elements["N"+Pos])){return false};
				if (!valSelec(document.forma.elements["C"+Pos])){return false};
				if (!valTxt(document.forma.elements["V"+Pos], 1, 5)){return false};
				if (!valSelec(document.forma.elements["T"+Pos])){return false};
				if (!valSelec(document.forma.elements["TC"+Pos])){return false};
				if (!valSelec(document.forma.elements["CCBL"+Pos])){return false};
                
                if (confirm('Confirme grabar rubro?')) {

                    var chk = forma.elements['CHK'];

                    chk[Pos].checked = true;                        

                    if (document.forma.elements['CID'+Pos]) {
                        
                        Action = 994; //update

                    } else {
                        
                        forma.elements["CLI"+Pos].value = forma.elements["ConsignerID"].value;
                        forma.elements["CNO"+Pos].value = forma.elements["ConsignerData"].value;;

                        Action = 995; //insert
                    
                    }

                    ColeccionRows(Pos);

                    ProcessRubros(Action, Pos) 
                }

            } else {
                alert('Nada que guardar');
            }

		return false;	 
    }

	 function DelCharge(Pos) {

        if (confirm('Confirme borrar rubro?')) {
            ProcessRubros(998, Pos);
        }
		return false;	 
	 }
     
	 function FreePedido(Pos) {

        if (confirm('Confirme liberar rubro de pedido CS?')) {
            ProcessRubros(996, Pos);
        }
		return false;	 
	 }

	 function LiberarRubro(Pos) {

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

        if (confirm('Confirme ' + document.forma.ObservacionesErp.value + ' ?')) {

            if (document.forma.elements["PER"+Pos].value == '')        
                ProcessRubros(996, Pos);
            else
                Facturar();
        }

        document.forma.PedidoCliente.value = '';
        document.forma.PedidoRubro.value = '';
		return false;	 
	 }

	 function ActualizaPedido(Pos) {

        var temp, temp1;
        temp = document.forma.elements["CLI"+Pos]; 
        temp1 = temp.value
        document.forma.PedidoCliente.value = temp1;

        if (document.forma.elements["PER"+Pos].value != "")
            document.forma.ObservacionesErp.value = "Actualizar Rubros Pedido ERP " + document.forma.elements["PER"+Pos].value + " Cliente " + document.forma.PedidoCliente.value;
        else
            document.forma.ObservacionesErp.value = "Transmitir Rubros Cliente " + document.forma.PedidoCliente.value;

        if (confirm('Confirme ' + document.forma.ObservacionesErp.value + ' ?')) {
            Facturar();
        }

        document.forma.PedidoCliente.value = '';
		return false;	 
	 }

     function ClienteUpdate(){
	    ClienteProc(1);
    	return false;	 
	 }

     function ClienteFree(){
		ClienteProc(2);
        return false;	 
	 }

     function ClienteProc(tipo){

        var select = document.getElementById('CmbClientes');
        var cliente_id = select.options[select.selectedIndex].value;
        var cliente_str = select.options[select.selectedIndex].text;

        if (cliente_id == '') {
            alert('Seleccione Cliente');
            return false;
        } 
        
        try {        
            if (tipo == 1) //transmitir rubros
                document.forma.ObservacionesErp.value = "Transmitir Rubros Cliente " + cliente_str;
            if (tipo == 2) //liberar rubros
                document.forma.ObservacionesErp.value = "Liberar Rubros Cliente " + cliente_str;
        } catch(err) {
            alert('Verifique conectividad con server ERP');
            return false;
        }

        var clientes;
        
        if (tipo == 1)
            clientes = document.forma.CmbRubrosTra.value.split('#*#');

        if (tipo == 2)
            clientes = document.forma.CmbRubrosLib.value.split('#*#');

        var row, rubros = '', cliente = '';

        for(var i = 0; i < clientes.length; i++){       
            row = clientes[i].split('|');        

            if (cliente_id == row[0]) {
                cliente = row[0];
                rubros = row[1].slice(0, -1);       
                break;
            }
        }

        if (cliente != '') {
        
            //alert('No hay cliente disponible');
        }

        if (rubros != '') {

            if (confirm('Confirme ' + document.forma.ObservacionesErp.value + ' ?')) {

                document.forma.PedidoCliente.value = cliente;

                if (tipo == 2)
                    document.forma.PedidoRubro.value = rubros; //envia rubros cuando va liberar

                Facturar();

                document.forma.PedidoCliente.value = '';
                document.forma.PedidoRubro.value = '';
            }

        } else {
        
            alert('No hay rubros disponibles');
        }

		return false;	 
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
     
        var forma = document.forms[0];
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
            SubmitSaves(999);
        }

        return false
     } 




     function ProcessRubros(Action, Pos) {

        // 998 delete
        // 997 liberar

        var forma = document.forms[0];
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
     
        SubmitSaves(Action);
             
     }


     
     function SubmitSaves(Action) {
        move();        
        console.log('Ejecuta submit saves : ' + Action);
        document.forma.Action.value = Action;
        document.forma.target = "saves";        
        document.forma.action = "ItineraryChargesSave.asp";
        document.forma.submit();     
     }


	 
	 function AddCharge(Pos) {
        if (document.forma.elements["T"+Pos].value != -1) {

            var iNo = ''; //document.getElementById('No').value;
            var ChargeMoneda = '';
            var servicio = ''
            var rubro = '';
    		window.open('Search_Charges.asp?PG=1&GID=29&N='+Pos+'&T=<%=BLType%>&IL='+(document.forma.elements["T"+Pos].value*1+1)+'&CM='+ChargeMoneda+'&No='+iNo+'&ServiceID='+servicio+'&ItemID='+rubro+'&esquema=<%=esquema%>&impex=<%=Movimiento%>','BLData','height=300,width=450,menubar=0,resizable=1,scrollbars=1,toolbar=0');
        } else {
            alert('Por favor indique el tipo de este cobro INT o LOC');
            document.forma.elements["T"+Pos].focus();
        }
		return false;	 
	 }
	 
	 function ValidarDoble(Pos) {
	 	for (i=0; i<=<%=CantItems%>;i++) {
			if  (i!= Pos) {
				if ((document.forma.elements["SVI"+i].value==document.forma.elements["SVI"+Pos].value) && 
				(document.forma.elements["SVN"+i].value==document.forma.elements["SVN"+Pos].value) &&
				(document.forma.elements["N"+i].value==document.forma.elements["N"+Pos].value) &&
				(document.forma.elements["I"+i].value==document.forma.elements["I"+Pos].value) &&
				(document.forma.elements["INV"+i].value=='0')) {
					alert("No puede repetir el mismo Rubro y Servicio si el anterior no ha sido facturado primero");
					DelCharge(Pos);
					return (false);
				}
			}			
		}
	 }


    function ClearRow(Pos){

		document.forma.elements["CLI"+Pos].value='0';   //id_cliente     
		document.forma.elements["CNO"+Pos].value='';    //descripcion
		document.forma.elements["PID"+Pos].value='0';   //id_pedido CS
		document.forma.elements["PER"+Pos].value='';    //pedido_erp
		document.forma.elements["CID"+Pos].value='0';   //ChargeID

		document.forma.elements["SVI"+Pos].value='';
		document.forma.elements["SVN"+Pos].value='';
		document.forma.elements["N"+Pos].value='';
		document.forma.elements["I"+Pos].value='';
		document.forma.elements["C"+Pos].value='-1';
		document.forma.elements["V"+Pos].value='';
		document.forma.elements["OV"+Pos].value='';
		document.forma.elements["T"+Pos].value='-1';
		document.forma.elements["TC"+Pos].value='-1';
		document.forma.elements["INV"+Pos].value='0';
		document.forma.elements["DTY"+Pos].value='0';
		document.forma.elements["CCBL"+Pos].value='-1';
    }


     function ColeccionRows(Pos){
     
        var sep = '';
		CantItems=-1;
  		document.forma.ItemServIDs.value = "";
		document.forma.ItemServNames.value = "";
  		document.forma.ItemNames.value = "";
  		document.forma.ItemIDs.value = "";
  		document.forma.ItemCurrs.value = "";
  		document.forma.ItemVals.value = "";
  		document.forma.ItemOVals.value = "";
  		document.forma.ItemLocs.value = "";
  		document.forma.ItemPPCCs.value = "";
		document.forma.ItemInvoices.value = "";
		document.forma.ItemDocType.value = "";
		document.forma.ItemCalcInBLs.value = "";
        document.forma.ItemInRO.value = "";
        document.forma.ItemChargeID.value = "";
        document.forma.ItemCli.value = "";
        document.forma.ItemCliNom.value = "";
		document.forma.ItemPedErp.value = "";
		document.forma.ItemRegimen.value = "";
		
		for (i=0; i<=<%=CantItems%>;i++) {

            if ((Pos > -1 && Pos == i) || (Pos == -1)) {

			    if (document.forma.elements["N"+i].value != '') {
				    if (!valSelec(document.forma.elements["N"+i])){return false};
				    if (!valSelec(document.forma.elements["C"+i])){return false};
				    if (!valTxt(document.forma.elements["V"+i], 1, 5)){return false};
				    if (!valSelec(document.forma.elements["T"+i])){return false};
				    if (!valSelec(document.forma.elements["TC"+i])){return false};
				    if (!valSelec(document.forma.elements["CCBL"+i])){return false};
				    if (document.forma.elements["OV"+i].value == '') {document.forma.elements["OV"+i].value = 0};
				    if (document.forma.elements["SVI"+i].value!="") {
					    document.forma.ItemServIDs.value = document.forma.ItemServIDs.value + sep + document.forma.elements["SVI"+i].value;
					    document.forma.ItemServNames.value = document.forma.ItemServNames.value + sep + document.forma.elements["SVN"+i].value;
				    } else {
					    document.forma.ItemServIDs.value = "0" + sep + document.forma.elements["SVI"+i].value;
					    document.forma.ItemServNames.value = " " + sep + document.forma.elements["SVN"+i].value;
				    }
				    document.forma.ItemNames.value = document.forma.ItemNames.value + sep + document.forma.elements["N"+i].value;
				    document.forma.ItemIDs.value = document.forma.ItemIDs.value + sep + document.forma.elements["I"+i].value;
				    document.forma.ItemCurrs.value = document.forma.ItemCurrs.value + sep + document.forma.elements["C"+i].value;
				    document.forma.ItemVals.value = document.forma.ItemVals.value + sep + document.forma.elements["V"+i].value;
				    document.forma.ItemOVals.value = document.forma.ItemOVals.value + sep + document.forma.elements["OV"+i].value;
				    document.forma.ItemLocs.value = document.forma.ItemLocs.value + sep + document.forma.elements["T"+i].value;
				    document.forma.ItemPPCCs.value = document.forma.ItemPPCCs.value + sep + document.forma.elements["TC"+i].value;
				    document.forma.ItemInvoices.value = document.forma.ItemInvoices.value + sep + document.forma.elements["INV"+i].value;
				    document.forma.ItemDocType.value = document.forma.ItemDocType.value + sep + document.forma.elements["DTY"+i].value;
				    document.forma.ItemCalcInBLs.value = document.forma.ItemCalcInBLs.value + sep + document.forma.elements["CCBL"+i].value;
                    document.forma.ItemInRO.value = document.forma.ItemInRO.value + sep + document.forma.elements["IRO"+i].value;
                    document.forma.ItemChargeID.value = document.forma.ItemChargeID.value + sep + document.forma.elements["ICID"+i].value;
                    document.forma.ItemCli.value = document.forma.ItemCli.value + sep + document.forma.elements["CLI"+i].value;
                    document.forma.ItemCliNom.value = document.forma.ItemCliNom.value + sep + document.forma.elements["CNO"+i].value;
                    
                    if (document.forma.elements["PER"+i])
                        document.forma.ItemPedErp.value = document.forma.ItemPedErp.value + sep + document.forma.elements["PER"+i].value;
				    
                    //if (document.forma.elements["R"+i])
                        document.forma.ItemRegimen.value = document.forma.ItemRegimen.value + sep + document.forma.elements["R"+i].value;
				    
                    CantItems++;
				    sep = "|";
			    }
            }
		}
	    document.forma.CantItems.value = CantItems;     
     }

function GetData(GID){
	window.open('Search_BLData.asp?GID='+GID+'&BTP=<%=BLType%>','BLData','height=200,width=450,menubar=0,resizable=1,scrollbars=1,toolbar=0,status=1');
}


</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<style type="text/css">
<!--
.style10 {font-size: 10px; font-family: Verdana, Arial, Helvetica, sans-serif; }
.style4 {	font-size:10px;
	color: #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
}
.style8 {	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 10px;
	font-weight: bold;
	color: #999999;
}
-->

input[attr=readonly] { background-color:silver; }

/*button { background-color:transparent; }*/

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
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="self.focus();">

<div id="myProgress">
  <div id="myBar">10%</div>
</div>

<iframe src="about:blank" height="150" width="900" name="saves" style='height:0px;border:0px'>
  <p>Su navegador no es compatible con iframes</p>
</iframe>


<TABLE cellspacing=0 cellpadding=2 width=400 align=center border=0>



	<FORM name="forma" action="ItineraryCharges.asp" method="post">
	<INPUT name="Action" type=hidden value=0>
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
	<INPUT name="OID" type=hidden value="<%=ObjectID%>">
	<INPUT name="ItemServIDs" type=hidden value="">
	<INPUT name="ItemServNames" type=hidden value="">
	<INPUT name="ItemNames" type=hidden value="">
	<INPUT name="ItemIDs" type=hidden value="">
	<INPUT name="ItemCurrs" type=hidden value="">
	<INPUT name="ItemVals" type=hidden value="">
	<INPUT name="ItemOVals" type=hidden value="">
	<INPUT name="ItemLocs" type=hidden value="">
	<INPUT name="ItemPPCCs" type=hidden value="">
	<INPUT name="CantItems" type=hidden value="-1">
	<INPUT name="ItemInvoices" type=hidden value="-1">
	<INPUT name="ItemDocType" type=hidden value="-1">
	<INPUT name="ItemCalcInBLs" type=hidden value="-1">
    <INPUT name="ItemInRO" type=hidden value="-1">
    <INPUT name="ItemChargeID" type=hidden value="">
    <INPUT name="ItemCli" type=hidden value="-1">
    <INPUT name="ItemCliNom" type=hidden value="">
    <INPUT name="ItemPedErp" type=hidden value="">
    <INPUT name="ItemRegimen" type=hidden value="">

	<INPUT name="Main" type=hidden value="1">
	<INPUT name="HBLNumber" type=hidden value="<%=HBLNumber %>">
	<INPUT name="CountryOrigen" type=hidden value="<%=CountryOrigen %>">
	<INPUT name="CountriesFinalDes" type=hidden value="<%=CountriesFinalDes %>">
	<INPUT name="CountryExactus" type=hidden value="<%=CountryExactus %>">
	<INPUT name="Movimiento" type=hidden value="<%=Movimiento %>">
	<INPUT name="Pedido_Erp" type=hidden value="<%=Pedido_Erp %>">
	<INPUT name="BLType" type=hidden value="<%=BLType %>">
    <INPUT name="esquema" type=hidden value="<%=esquema %>">
    <INPUT name="PedidoCliente" type=hidden value="">
    <INPUT name="PedidoRubro" type=hidden value="">


		<TD colspan="2" class=label align=center>

        <table width="100%" border="0" align="left">

<%
            if Msg <> "" then                  
                'response.write "<tr><td colspan=3>" & Msg & "</td></tr>"
            end if   
%>

		<tr>
        <td width="25%" valign=top>
        	
            <table width="90%" border="0" align="center">
            	<TR><TD class=label align=right><b>BL ID:</b></TD><TD class=label align=left colspan=2><%=ObjectID%></TD></TR> 
			    <TR><TD class=label align=right><b>Carta Porte:</b></TD><TD class=label align=left colspan=2><%=HBLNumber%></TD></TR> 
			    <TR><TD class=label align=right><b>Consignatario:</b></TD><TD class=label align=left colspan=2><%=ClientsID & " - " & Name%></TD></TR>
			    <TR><TD class=label align=right><b>Volumen:</b></TD><TD class=label align=left colspan=2><%=Volume%></TD></TR> 
			    <TR><TD class=label align=right><b>Peso:</b></TD><TD class=label align=left colspan=2><%=Weight%></TD></TR> 
			    <TR><TD class=label align=right><b>Shipper:</b></TD><TD class=label align=left colspan=2><%=Agent%></TD></TR> 
			    <TR><TD class=label align=right><b>BL o RO:</b></TD><TD class=label align=left><%=BL%></TD><TD class=label align=right>
                
                    <% if TipoConta = "BAW" then %>
                    <a href="#" onClick="window.open('ItineraryInterCharges.asp?OID=<%=ObjectID%>&GID=29','CargosIntercompany','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=1100,height=480,top=170,left=170');" class="menu"><font color="FFFFFF"><b>Cargos Intercompany</b></font></a>
                    <% end if %>         
                
                </TD></TR> 


		    </table>

        </td>
        <td width="25%" valign=top align=left>

		    <table width="100%" border="0" align="left">

			    <TR><TD class=label align=right width=100px><b>Movimiento:</b></TD><TD class=label align=left colspan=2><%=Movimiento%></TD></TR> 
			    <TR><TD class=label align=right><b>Pais&nbsp;Creacion:</b></TD><TD class=label align=left colspan=2><%=Iif(CountryExactus = Countries, "<font style='background:silver'>" & Countries & " " & esquema & "</font>", Countries)%></TD></TR> 
			    <TR><TD class=label align=right><b>Pais&nbsp;Origen:</b></TD><TD class=label align=left colspan=2><%=Iif(CountryExactus = CountryOrigen, "<font style='background:silver'>" & CountryOrigen & " " & esquema & "</font>", CountryOrigen)%></TD></TR>
			    <TR><TD class=label align=right><b>Pais&nbsp;Destino:</b></TD><TD class=label align=left colspan=2><%=Iif(CountryExactus = CountriesFinalDes, "<font style='background:silver'>" & CountriesFinalDes & " " & esquema & "</font>", CountriesFinalDes)%></TD></TR> 
			    <TR><TD class=label align=right><b>RO&nbsp;Facturar:</b></TD><TD class=label align=left colspan=2><%=facturar_a & " " & facturar_a_nombre%></TD></TR> 

                <TR>
                <TD align=center colspan=2>
                <!--
                    <INPUT name=enviar type=button onClick="JavaScript:validar(2,-1)" value="&nbsp;&nbsp;Actualizar&nbsp;&nbsp;" class=label>
               -->
               <!--
                </TD>
                <td class=label align=center>
                -->
<% 
Dim iLink
iLink = "GID=0&ObjectID=" & ObjectID & "&DocTyp=" & Iif(Movimiento = "EXPORT", 0, 1) & "&HAWBNumber=" & HBLNumber & "&AWBNumber=" & HBLNumber & "&BLType=" & BLType & "&esquema=" & esquema
%>


<button onClick="Javascript:window.open('ItineraryCharges-Facturacion.asp?<%=iLink%>','AWBData','height=400,width=1100,menubar=0,resizable=1,scrollbars=1,toolbar=0');return false;" title="Buscar Rubro">Articulos / Pedidos / Facturas <img src="img/relocation_console.gif" /></button>

                </td>
                </TR> 

		    </table>

        </td>

        <%
'response.write "(" & facturar_a & ")(" & facturar_a_nombre & ")<br>" 

if ConsignerID = "" and ConsignerData = "" then 
    if facturar_a = "" and facturar_a_nombre = "" then 
	    ConsignerID = ClientsID
	    ConsignerData = Name	
    else  
	    ConsignerID = facturar_a
	    ConsignerData = facturar_a_nombre		
end if
end if

'response.write "(" & ConsignerID & ")(" & ConsignerData & ")<br>" 

        %>
        <td width="25%" valign=top align=left>
               
		    <table width="100%" border="0" align="left">
                <TR>
                    <% if Pedido_Erp = "" then %>
                    <TD class=label align=right nowrap><b>Pedido Abierto:</b></TD>
                    <TD nowrap class=label align=left width=100 style="background-color:silver;color:black"><%=Pedido_Erp%></TD>
                    <TD>
                    <% if Pedido_Erp = "" and TipoConta = "EXACTUS" and ItemsPedidos = 0 then %>
                     <input name=enviar type=button onClick="JavaScript:Solicitar(4);"  value="&nbsp;&nbsp;Solicitar&nbsp;&nbsp;" class=label>
                    <% else %>
                     <input name=enviar type=button onclick="alert('Ya hay rubros, hasta abajo presione Transmitir');document.getElementById('Observaciones para facturacion').focus();"  value="&nbsp;&nbsp;Solicitar&nbsp;&nbsp;" class=label>
                    <% end if %>
                     </TD>
                    <% else %>
                        <td><h1></h1></td>
                    <% end if %>
                </TR> 

                <TR><TD class=label align=center colspan=2>


                </TD>
                <td>
              
                </td>
                </TR> 

                <TR><TD class=label align=center colspan=2>
				
                </TD></TR> 

                <TR>
                    <TD class=label align=center colspan=3>
                        <table>
                        <TR>
                            <TH class=label align=center>Asignar Cliente:</TH>
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

        <td width="25%" valign=top align=left nowrap>
                
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

                <input  name='CmbRubrosTra' id='CmbRubrosTra' type="text" />
                <br />
                <input  name='CmbRubrosLib' id='CmbRubrosLib' type="text" />
                
        </td>

        </tr>
        </table>

		 <%'="(" & ItemsPedidos & ")(" & CountTableValues & ")"%>

		<table width="80%" border="0">
		  <tr><td class="submenu" colspan="15"></td></tr>
          <tr><td class="style4" colspan="15" align="center">CARGOS A CLIENTES</td></tr>
          <tr><td class="submenu" colspan="15"></td></tr>
		  <tr>
			<td align="center" class="menu activeMain">
			</td>
			<td align="center" class="menu activeMain" nowrap>
				
			</td>
            <td align="center" class="menu activeMain">
				Cliente  
			</td>
			<td align="center" class="menu activeMain">
				Tipo
			</td>
			<td align="center" class="menu activeMain">
				Servicio
			</td>
			<td align="center" class="menu activeMain">
				Rubro
			</td>
			<td align="center" class="menu activeMain">&nbsp;</td>
			<td align="center" class="menu activeMain">
				Regimen
			</td>
			<td align="center" class="menu activeMain">
				Moneda
			</td>
			<td align="center" class="menu activeMain">
				Monto
			</td>
			<td align="center" class="menu activeMain">
				Sobre Venta
			</td>
			<td align="center" class="menu activeMain">
				Pago
			</td>
			<td align="center" class="menu activeMain">
				Calcular en HBL?
			</td>
			<td align="center" class="menu activeMain">
			</td>
			<td align="center" class="menu activeMain">
				Pedido/Factura
			</td>
            <td align="center" class="menu activeMain">
				Estado  
			</td>
		  </tr>
		  <%for i=0 to CantItems %>
		  
          <tr id="Row<%=i%>">

			<td align="center" class="style4" nowrap>

				<input type="checkbox"  id="CHK<%=i%>" name="CHK"   alt="hilmar" title="hans"    value="0">

                <% if i <= CountTableValues then  %>
                <input type="hidden"      id="CID<%=i%>" name="CID<%=i%>" size="2" readonly>
                <input type="hidden"      id="PID<%=i%>" name="PID<%=i%>" size="1" readonly>
                <input type="hidden"      id="PER<%=i%>" name="PER<%=i%>" size="1" readonly>
                <% end if  %>

			</td>

            <td align="right" class="style4" nowrap>

                <% 'if i <= CountTableValues then  %>
			    <!-- <a href="#" onClick="Javascript:FreePedido(<%=i%>);" id="CR3<%=i%>" style="display:inline" title="Liberar Pedido CS"><img src="img/editPassword.gif" /></a> -->
                <% 'end if  %>

                <% if i <= CountTableValues then  %>
                <button onClick="Javascript:LiberarRubro(<%=i%>);return (false);" id="CR1<%=i%>" title="Liberar de Pedido ERP"><img src="img/remove.gif" /></button>
                <% end if  %>

                <% 'if i <= CountTableValues then  %>
                <!-- <button onClick="Javascript:ActualizaPedido(<%=i%>);return (false);" id="CR2<%=i%>" title="Actualiza Pedido ERP"><img src="img/inter.gif" /></button> -->
                <% 'end if  %>

                <% if i <= CountTableValues then  %>
			    <button onClick="Javascript:DelCharge(<%=i%>);return (false);" id="DE<%=i%>" title="Borrar Rubro"><img src="img/delete.gif" /></button>
                <% end if  %>

			</td>
            <td align="left" class="style4" nowrap>
				<input type="text" size="5" class="style10" name="CLI<%=i%>" readonly>
				<input type="text" size="15" class="style10" name="CNO<%=i%>" readonly>
			</td>
			<td align="right" class="style4">
				<select class='style10' name='T<%=i%>' id="Tipo de Cobro">
				<option value='-1'>---</option>
				<option value='0'>INT</option>
				<option value='1'>LOC</option>
			 	</select>
			</td>
			<td align="right" class="style4">
				<input type="text" size="15" class="style10" name="SVN<%=i%>" value="" readonly>
				<input type="hidden" name="SVI<%=i%>" value="">
			</td>
			<td align="right" class="style4">
				<input type="text" size="15" class="style10" name="N<%=i%>" value="" readonly>
				<input type="hidden" name="I<%=i%>" value="">
			</td>
			<td align="right" class="style4">
				<div id=DR<%=i%> style="VISIBILITY: visible;">
                <button onClick="Javascript:AddCharge(<%=i%>);return (false);" title="Buscar Rubro"><img src="img/Search16.png" /></button>
				</div>
			</td>
			<td align="right" class="style4"><input type="text" size="15" class="style10" name="R<%=i%>" value="" readonly>
			</td>
			<td align="right" class="style4">
				<select class='style10' name='C<%=i%>' id="Moneda">
				<option value='-1'>---</option>
				<%=Currencies%>
				</select>
			</td>
			<td align="center" class="style4">
				<input type="text" size="8" class="style10" name="V<%=i%>" value="" onKeyUp="res(this,numb);" id="Monto">
			</td>
			<td align="center" class="style4">
				<input type="text" size="4" class="style10" name="OV<%=i%>" value="" onKeyUp="res(this,numb);" id="SobreVenta">
			</td>
			<td align="right" class="style4">
				<select class='style10' name='TC<%=i%>' id="Forma de Pago">
				<option value='-1'>---</option>
				<option value='0'>PREPAID</option>
				<option value='1'>COLLECT</option>
			 	</select>
			</td>
            <td align="right" class="style4">
				<select class='style10' name='CCBL<%=i%>' id="Calcular en BL">
				<option value='-1'>---</option>
				<option value='0'>NO</option>
				<option value='1'>SI</option>
			 	</select>
			</td>
            <td>
                <button onClick="Javascript:SaveCharge(<%=i%>);return (false);" id="SV<%=i%>" title="Buscar Rubro"><img src="img/floppy.gif" /></button>
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
		  </tr>
		  <%next%>
		</table>

		<TABLE cellspacing=0 cellpadding=2 width=70% border=0 align=left>

        <TR><TD colspan=4><%=ObservacionesErp%></TD></TR>

        <TR>		 
             <TD class=label align=center>   
             <input type="hidden" name="TipoConta" value="<%=TipoConta%>" />          
             <% if SelectBodegas <> "" then %>
             <span class="menu" style="display:block;color:White;padding:3px;">BODEGA&nbsp;:&nbsp;</span><select name="SelectBodegas" id="Bodega" style="max-width:280px;"><%=SelectBodegas%></select>
             <% end if %>
             </TD>

             <TD class=label align=center>             
             <% if ActividadComercial <> "" then %>
             <span class="menu" style="display:block;color:White;padding:3px;">ACTIVIDAD COMERCIAL&nbsp;:&nbsp;</span><select name="ActividadComercial" id="Actividad Comercial" style="max-width:280px;"><%=ActividadComercial%></select>
             <% end if %>
             </TD>

             <TD class=label align=center>             
             <% if CondicionPago <> "" then %>
             <span class="menu" style="display:block;color:White;padding:3px;">CONDICION DE PAGO&nbsp;:&nbsp;</span></span><select name="CondicionPago" id="Condicion de Pago" style="max-width:280px;"><%=CondicionPago%></select>
             <% end if %>
             </TD>

             <TD class=label align=center>         
             <% if TipoConta = "EXACTUS" then %>

                <% if Pedido_Erp = "" then %>
                
                    <% if ItemsPedidos = 0 then %>

                         <input name=enviar type=button onclick="alert('No hay rubros para Transmitir, Solicite Pedido Abierto')" value="&nbsp;&nbsp;Transmitir&nbsp;&nbsp;" class=label>

                    <% else %>
                
                         <input name=enviar type=button onclick="JavaScript:FacturarAbierto();" value="&nbsp;&nbsp;Transmitir Abierto&nbsp;&nbsp;" class=label>
        
                    <% end if %>
                
                
                <% else %>

                     <input name=enviar type=button onClick="JavaScript:Facturar();" value="&nbsp;&nbsp;Transmitir Pedido&nbsp;&nbsp;" class="Boton cBlue">

                <% end if %>

             <% end if %>   
                   
             </TD>
		</TR>
		
        
        <TR><TD colspan=4>&nbsp;</TD></TR>


        </TABLE>
		
		<TD>
		</TR>
	</FORM>
	</TABLE>
</BODY>
</HTML>
<script>

<% if (iSelectBodegas <> "") then %>
    document.forma.SelectBodegas.value = '<%=iSelectBodegas%>';            
<% end if %>

<% if (iCondicionPago <> "") then %>
    document.forma.CondicionPago.value = '<%=iCondicionPago%>';            
<% end if %>

<% if (iActividadComercial <> "") then %>
    document.forma.ActividadComercial.value = '<%=iActividadComercial%>';            
<% end if %>

<% if (Request("ObservacionesErp") <> "") then %>
    document.forma.ObservacionesErp.value = '<%=Request("ObservacionesErp")%>';            
<% end if %>

var cli_ant = 0, cli = 0, c = 0; bg = '', cli_opt = '', rub_opt_lib = '', rub_opt_tra = '', DocType = 0, pedido_cs = 0, pedido_erp = '', Invoice = 0;

cli_opt += "<option value=''>-- Seleccione --</option>";

<%for i=0 to CountTableValues%>

    cli = '<%=aTableValues(18,i)%>';
    DocType = '<%=aTableValues(12,i)%>';
    pedido_cs = '<%=aTableValues(19,i)%>';
    pedido_erp = '<%=aTableValues(20,i)%>';
    Invoice = '<%=aTableValues(10,i)%>';

    if (cli != cli_ant) {
        c++;    
        cli_ant = cli;
        cli_opt += "<option value='" + cli + "'>" + cli + " - " +  '<%=aTableValues(21,i)%>' + "</option>";

        if (rub_opt_lib != '') rub_opt_lib += '#*#';
        rub_opt_lib += cli + '|';

        if (rub_opt_tra != '') rub_opt_tra += '#*#';
        rub_opt_tra += cli + '|';
    }

    /*if (c % 2 == 0) {
        bg = 'lightblue'    
    } else {    
        bg = 'orange';    
    }*/

	document.getElementById('Row<%=i%>').style.background = ((c % 2 == 0) ? 'lightblue' : 'orange');
	document.forma.CLI<%=i%>.value = cli; //id_cliente     
    document.forma.CNO<%=i%>.value = '<%=aTableValues(21,i)%>'; //cliente_nombre     
	document.forma.PID<%=i%>.value = pedido_cs;
	document.forma.PER<%=i%>.value = pedido_erp;
    document.forma.CID<%=i%>.value = '<%=aTableValues(16,i)%>'; //ChargeID
    document.forma.R<%=i%>.value = '<%=aTableValues(22,i)%>'; //Regimen
    document.getElementById('CHK'+<%=i%>).title = '<%=aTableValues(16,i)%>' + ' CS:' + '<%=aTableValues(19,i)%>' + ' ERP:' + '<%=aTableValues(20,i)%>';
    
    //document.getElementById("CR3<%=i%>").style.display = "none";   //free pedido_id	        
    //document.getElementById("CR2<%=i%>").style.display = "none";   //se traslado para el encabezado

    /*
    if (DocType == 9) { //si no esta facturado
            rub_opt_lib += '<%=aTableValues(16,i)%>' + ',';
    } 
    if (DocType == 10) { //si no esta facturado   
    } else {
         if (DocType == 9 && pedido_erp == '') {
            rub_opt_tra += '<%=aTableValues(16,i)%>' + ',';
        }
    }
    if (DocType == 0) {
        rub_opt_tra += '<%=aTableValues(16,i)%>' + ',';
    }    
    */

    //alert('DocType=' + DocType + ' ' + 'pedido_erp=' + pedido_erp + ' ' + 'pedido_cs=' + pedido_cs);

    if (DocType == 9 || DocType == 0) {
        rub_opt_tra += document.forma.CID<%=i%>.value + ','; //ChargeID
    }    

    if ((Invoice == 0 && DocType == 0 && pedido_erp != '') || pedido_cs == 0 || DocType == 10) {
        document.getElementById("CR1<%=i%>").style.display = "none";   //free pedido erp   
    } else {
        rub_opt_lib += document.forma.CID<%=i%>.value + ','; //ChargeID
    }

    <% if aTableValues(10,i) = 0 and aTableValues(12,i) = 0 and aTableValues(20,i) <> "" then %> // INVOICE - DOCTYPE - PEDIDO_ERP

	        //document.getElementById("CR1<%=i%>").style.display = "none";   //free pedido erp

    <% else %>

        if (document.forma.PER<%=i%>.value == '') { //no hay pedido_erp

            document.getElementById('CHK'+<%=i%>).checked = false;
            document.getElementById('CHK'+<%=i%>).value = '0';    

	        //document.getElementById("CR2<%=i%>").style.display = "none";   //update pedido erp all pedidos id
        
            //if (document.forma.PID<%=i%>.value == '0') { //no hay pedido_id
    	        //document.getElementById("CR1<%=i%>").style.display = "none";   //free pedido erp        	    
            //}

        } else { //si tiene pedido_erp

	        document.getElementById("DE<%=i%>").style.display = "none";    //delete
	        document.getElementById("SV<%=i%>").style.display = "none";    //insert update

            document.getElementById('CHK'+<%=i%>).checked = true;
            document.getElementById('CHK'+<%=i%>).value = document.forma.CID<%=i%>.value + ','; //ChargeID
            document.getElementById('CHK'+<%=i%>).disabled = true;         
        }

    <% end if %>

    if ('<%=aTableValues(12,i)%>' == 10) {  //si no ha sido facturado   
        //document.getElementById("CR1<%=i%>").style.display = "none";   //free pedido erp   
    }

	document.forma.N<%=i%>.value = '<%=aTableValues(1,i)%>';
	document.forma.I<%=i%>.value = '<%=aTableValues(2,i)%>';
	selecciona('forma.C<%=i%>','<%=aTableValues(3,i)%>');
	document.forma.V<%=i%>.value = '<%=aTableValues(4,i)%>';
	document.forma.OV<%=i%>.value = '<%=aTableValues(5,i)%>';
	selecciona('forma.T<%=i%>','<%=aTableValues(6,i)%>');
	selecciona('forma.TC<%=i%>','<%=aTableValues(7,i)%>');
	document.forma.SVI<%=i%>.value = '<%=aTableValues(8,i)%>';
	document.forma.SVN<%=i%>.value = '<%=aTableValues(9,i)%>';
	document.forma.INV<%=i%>.value = '<%=aTableValues(17,i)%>';
	document.forma.DTY<%=i%>.value = '<%=aTableValues(12,i)%>';
	document.forma.CCBL<%=i%>.value = '<%=aTableValues(11,i)%>';    
    document.forma.FAC<%=i%>.value = '<%=aTableValues(13,i)%>';
    document.forma.FAC<%=i%>.title = '<%=aTableValues(13,i)%>';
    document.getElementById('STATFAC<%=i%>').innerHTML = '<%=aTableValues(14,i)%>';
    document.forma.IRO<%=i%>.value = '<%=aTableValues(15,i)%>';
    document.forma.ICID<%=i%>.value = '<%=aTableValues(16,i)%>';
    <% if ((aTableValues(10,i) = 0) and (aTableValues(15,i) = 0)) then %>
	    document.forma.N<%=i%>;
	    document.forma.I<%=i%>;
	    document.forma.C<%=i%>;
	    document.forma.V<%=i%>;
	    document.forma.OV<%=i%>;
	    document.forma.T<%=i%>;
	    document.forma.TC<%=i%>;
	    document.forma.SVI<%=i%>;
	    document.forma.SVN<%=i%>;
	    document.forma.CCBL<%=i%>;
	    document.getElementById("DE<%=i%>");
	    document.getElementById("DR<%=i%>");
    <% elseif ((aTableValues(10,i) = 0) and (aTableValues(15,i) <> 0)) then %>
	    document.forma.N<%=i%>.disabled = 'false';
	    document.forma.I<%=i%>.disabled = 'false';
	    document.forma.C<%=i%>.disabled = 'false';
	    document.forma.V<%=i%>.disabled = 'false';
	    document.forma.OV<%=i%>.disabled = 'false';
	    document.forma.T<%=i%>;
	    document.forma.TC<%=i%>.disabled = 'false';
	    document.forma.SVI<%=i%>.disabled = 'false';
	    document.forma.SVN<%=i%>.disabled = 'false';
	    document.forma.CCBL<%=i%>;
	    document.getElementById("DE<%=i%>").style.display = "none";
	    document.getElementById("DR<%=i%>").style.visibility = "hidden";

        document.forma.CLI<%=i%>.disabled = 'false'; //id_cliente     
        document.forma.CNO<%=i%>.disabled = 'false'; //cliente_nombre     

    <% elseif (aTableValues(10,i) <> 0) then %>
	    document.forma.N<%=i%>.disabled = 'false';
	    document.forma.I<%=i%>.disabled = 'false';
	    document.forma.C<%=i%>.disabled = 'false';
	    document.forma.V<%=i%>.disabled = 'false';
	    document.forma.OV<%=i%>.disabled = 'false';
	    document.forma.T<%=i%>.disabled = 'false';
	    document.forma.TC<%=i%>.disabled = 'false';
	    document.forma.SVI<%=i%>.disabled = 'false';
	    document.forma.SVN<%=i%>.disabled = 'false';
	    document.forma.CCBL<%=i%>;
	    document.getElementById("DE<%=i%>").style.display = "none";
	    document.getElementById("DR<%=i%>").style.visibility = "hidden";

        document.forma.CLI<%=i%>.disabled = 'false'; //id_cliente     
        document.forma.CNO<%=i%>.disabled = 'false'; //cliente_nombre     

    <% end if %>

    
<%next%>

    //console.log(rub_opt); 
           
	document.getElementById('CmbClientes').innerHTML = cli_opt;
	document.getElementById('CmbRubrosLib').value = rub_opt_lib;
	document.getElementById('CmbRubrosTra').value = rub_opt_tra;

<%for i=CountTableValues+1 to CantItems%>
    if (document.getElementById('CHK'+<%=i%>))

        document.getElementById('CHK'+<%=i%>).style.display = 'none';         
        
        //document.getElementById('CHK'+<%=i%>).disabled = true;         
<%next
Set aTableValues = Nothing
if Action=1 or Action=2 then%>
	top.opener.location.reload();
<%end if%>



</script>