<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
Checking "0|1|2"

Dim ObjectID, DUA, GroupID, WsRes, Action, Action2, CountriesFinalDes, TipoConta, result
Dim SelectBodegas, ActividadComercial, CondicionPago, ObservacionesErp, iSelectBodegas, iActividadComercial, iCondicionPago, CountryOrigen, Countries, Movimiento, CountryExactus, Pedido_Erp, HBLNumber, FacID, Pedido_Msg, iLink, BLType, esquema, PedidoCliente, PedidoRubro

ObjectID = CheckNum(Request("OID"))
GroupID = CheckNum(Request("GID"))
Action = CheckNum(Request("Action"))
FacID = CheckNum(Request("Main"))
SelectBodegas = Request("SelectBodegas")
ActividadComercial = Request("ActividadComercial")
CondicionPago = Request("CondicionPago")
ObservacionesErp = Request("ObservacionesErp")              
HBLNumber = Request("HBLNumber")
CountryOrigen = Request("CountryOrigen")
CountriesFinalDes = Request("CountriesFinalDes")
CountryExactus = Request("CountryExactus")
Movimiento = Request("Movimiento")
Pedido_Erp = Request("Pedido_Erp")
Pedido_Msg = Request("Pedido_Msg")
BLType = Request("BLType")
esquema = Request("esquema")
PedidoCliente = Request("PedidoCliente")
PedidoRubro = Request("PedidoRubro")

DUA = Request("DUA")

if InStr(1,Request("Pedido_Msg"),",") > 0 then
    r = Split(Request("Pedido_Msg"), ",")
    Pedido_Msg = r(0)
else
    Pedido_Msg = Request("Pedido_Msg")
end if
          
    response.write "(Action=" & Action & ")(" & FacID & ")(" & PedidoCliente & ")(" & PedidoRubro & ")<br>"

    if Pedido_Msg <> "" then 
        On Error Resume Next

    'response.write "<br>*********************************<br>" & Request("Pedido_Msg") & "<br>***********************<br>"

            Pedido_Msg = Base64Decode2(Pedido_Msg)
    
    'response.write "<br>*********************************<br>" & Pedido_Msg & "<br>***********************<br>"
            
            result = split(Pedido_Msg,"#*#")
            
            if IFNULL(result(1)) <> "" then 'Stat
                response.write "<span class=Textos>" & result(1) & "<span class=Textos>"    'Msg
            end if

            if IFNULL(result(4)) <> "" then 'Error
                response.write "<span class=Textos>" & result(2) & "<span class=Textos>"    'Msg
            end if

            Pedido_Msg = ""           
        If Err.Number <> 0 Then
            response.write "<br>Pedido_Msg Error : " & Err.Number & " - " & Err.description & "<br>"  
        end if
    end if    

    WsRes = 0

if FacID = 0 then       
    On Error Resume Next
        if Action = 3 or Action = 4 then '- 3 Pedido Normal ""    - 4 Pedido Abierto "1"     
		    result = WsExactusSetPedidos(ObjectID, Movimiento, SelectBodegas, ActividadComercial, CondicionPago, ObservacionesErp, CountryExactus, Session("Login"), Request.ServerVariables("REMOTE_ADDR"), "1", Iif(Action = 3, "", "1"), PedidoCliente, PedidoRubro, DUA, Pedido_Erp)
            Pedido_Msg = Base64Encode2("" & result(0) & "#*#" & result(1) & "#*#" & result(2) & "#*#" & result(3) & "#*#" & result(4) & "")
            
            Action = 9

            if CheckNum(result(0)) = 1 then
                WsRes = 1
            else
                WsRes = 2                
            end if

        end if
    If Err.Number <> 0 Then
        response.write "<br>WsExactusSetPedidos Error : " & Err.Number & " - " & Err.description & "<br>"  
    end if
end if
%>

<body onload="Notificar();" >

    <div id="content"><span id="Leyenda" class="Textos"></span></div>

	<div id=Retorno style="display:none">
        <center><h1 style="font-family:Arial">PEDIDOS EXACTUS</h1></center>

		<form name=forma1 action="ItineraryChargesPedidos.asp" method=post>    
    
	        <INPUT name="GID" type=hidden value="<%=GroupID%>">
	        <INPUT name="CountryOrigen" type=hidden value="<%=CountryOrigen %>">
	        <INPUT name="CountriesFinalDes" type=hidden value="<%=CountriesFinalDes %>">
	        <INPUT name="Action" type=hidden value="3">
            <INPUT name="esquema" type=hidden value="<%=esquema%>">
	        <INPUT name="Pedido_Msg" type=hidden value="<%=Pedido_Msg%>">
            <INPUT name="PedidoCliente" type=hidden value="<%=PedidoCliente%>">
            <INPUT name="PedidoRubro" type=hidden value="<%=PedidoRubro%>">
            <INPUT name="Pedido_Erp" type=hidden value="<%=Pedido_Erp%>">

            <table width=50% align=center>

            <% if Action = 4 then %>
                <tr><th>Pedido ERP</th><td> <%=Pedido_Erp%>
                <%
                if Action = 9 or Action = 10 then 
                    response.write "*"
                end if
                %>          
                </td></tr>            
            <% end if %>

            <tr><th>BL ID</th><td> <INPUT name="OID" type=text readonly value="<%=ObjectID%>"></td></tr>
            <tr><th>CountryExactus</th><td> <INPUT name="CountryExactus" type=text readonly value="<%=CountryExactus%>"></td></tr>
            <tr><th>HBLNumber</th><td> <INPUT name="HBLNumber" type=text readonly value="<%=HBLNumber%>"></td></tr>
            <tr><th>Movimiento</th><td> <INPUT name="Movimiento" type=text readonly value="<%=Movimiento%>"></td></tr>
            <tr><th>Bodega</th><td> <INPUT name="SelectBodegas" type=text readonly value="<%=SelectBodegas%>"></td></tr>
            <tr><th>Actividad Comercial</th><td> <INPUT name="ActividadComercial" type=text readonly value="<%=ActividadComercial%>"></td></tr>
            <tr><th>Condicion de Pago</th><td> <INPUT name="CondicionPago" type=text readonly value="<%=CondicionPago%>"></td></tr>
            <tr><th>Observaciones</th><td>             
                <textarea name="ObservacionesErp" style="width:100%" rows=3 readonly><%=ObservacionesErp%></textarea>            
            </td></tr>
            <tr><th>Dua:</th><td> <INPUT name="DUA" type=text readonly value="<%=DUA%>"></td></tr>
            </table>
	 
            <table width=50% align=center border=0>
            <tr>
            <td width=50% valign=top align=center>
            <br />
            <% if Action = 9 then %>
                    <input type="submit" value="<%=Iif(Action = 3, "Transmitir", "Solicitar Pedido Abierto") %>" disabled>
            <% else %>
                    <input type="submit" value="<%=Iif(Action = 10 or Action = 3, "Transmitir", "Solicitar Pedido Abierto") %>" onclick="return Loading('<%=Iif(Action = 3 or Action = 10, "Esta seguro de Transmitir Pedido ?", "Esta seguro de Solicitar Pedido Abierto ?") %>', '<%=Iif(Action = 3 or Action = 10, "Transmitiendo", "Solicitando") %>');" >
            <% end if %>
            </td>
            <td nowrap>
                 <% 
                 iLink = "GID=0&ObjectID=" & ObjectID & "&DocTyp=" & Iif(Movimiento = "EXPORT", 0, 1) & "&HAWBNumber=" & HBLNumber & "&AWBNumber=" & HBLNumber & "&BLType=" & BLType & "&esquema=" & esquema
                 %>
                <br />
                <a href="#" onClick="Javascript:window.open('ItineraryCharges-Facturacion.asp?<%=iLink%>','AWBData','height=400,width=1100,menubar=0,resizable=1,scrollbars=1,toolbar=0');return false;">Articulos / Pedidos / Facturas</a>
            </td>

        </form>

            <td width=50% valign=top align=center> 

			<br />
                    <form name="forma" action="ItineraryCharges.asp">	
	                <INPUT name="CountryExactus" type=hidden value="<%=CountryExactus%>">
	                <INPUT name="HBLNumber" type=hidden value="<%=HBLNumber%>">                    
	                <INPUT name="GID" type=hidden value="<%=GroupID%>">
	                <INPUT name="OID" type=hidden value="<%=ObjectID %>">
	                <INPUT name="Pedido_Erp" type=hidden value="<%=Pedido_Erp%>">	                
                    <INPUT name="Movimiento" type=hidden value="<%=Movimiento%>">
	                <INPUT name="SelectBodegas" type=hidden value="<%=SelectBodegas%>">
	                <INPUT name="ActividadComercial" type=hidden value="<%=ActividadComercial%>">
	                <INPUT name="CondicionPago" type=hidden value="<%=CondicionPago %>">
	                <INPUT name="ObservacionesErp" type=hidden value="<%=ObservacionesErp %>">
                    <INPUT name="esquema"        type=hidden value="<%=esquema%>">			   	
                    <INPUT name="DUA"           type=hidden value="<%=DUA%>">			   	
         
                    <input type="hidden" value="Retornar" onclick="return Loading('Esta seguro de retornar ?','Retornando');" >
           
			        <input type="button" value="Retornar" onclick="if (confirm('Confirme Retornar?')) {location.href = 'ItineraryCharges.asp?GID=<%=GroupID%>&OID=<%=ObjectID%>';} return false;" >

                    </form>
            </td>
            </tr>            
            </table>




            <div style="background-color:white;color:white;">
<%
    On Error Resume Next

        'response.write "<textarea class=Textos readonly>" & result(0) & "</textarea><br>"   
        'response.write "<textarea class=Textos readonly>" & result(1) & "</textarea><br>"   
        'response.write "<textarea class=Textos readonly>" & result(2) & "</textarea><br>"   
        'response.write "<textarea class=Textos readonly>" & result(3) & "</textarea><br>"   
        'response.write "<textarea class=Textos readonly>" & result(4) & "</textarea><br>"   

        if IFNULL(result(0)) = "" then 

            'response.write "<span class=Textos>NULL-</span>"  
        else
                if result(0) > 1 then 'Stat

                    if IFNULL(result(3)) <> "" then 
                        response.write "<H3>LOG : </H3>"  
                        response.write result(3) & "<HR>"   
                    end if

                    'if IFNULL(result(4)) <> "" then
                    '    response.write "<H3>ERROR : </H3>"  
                    '    response.write result(4) & "<HR>"   
                    'end if
            
                end if
        end if

    If Err.Number <> 0 Then
        'response.write "<br>Pedido_Msg Error : " & Err.Number & " - " & Err.description & "<br>"  
    end if
%>
	</div>

</body>

<script type="text/javascript">

    <% if Action = 9 then  %> // trae respuesta del ws

        <% if WsRes = 1 then %> // tuvo exito la transmision

			location.href = 'ItineraryCharges.asp?GID=<%=GroupID%>&OID=<%=ObjectID%>';

        <% else %> // hay errores
        
            document.forma1.Action.value = 10;                      
            document.forma1.submit();

        <% end if %>

    <% end if %>  

    function Notificar() {
        document.getElementById("Retorno").style.display = "inline";
        document.body.style.backgroundColor = "white";
        document.body.style.backgroundImage = "url('')";
        document.getElementById("content").style.display = "none";
        document.getElementById("Leyenda").innerHTML = '';
    }


    function Loading(pregunta, mensaje) {

        if (confirm(pregunta)) {

            var a = parseInt("194");
            var b = parseInt("223");
            var c = parseInt("239");
            document.body.style.backgroundColor = "rgb(" + [a, b, c].join() + ")";

            document.getElementById("Retorno").style.display = "none";
            document.body.style.backgroundImage = "url('img/loader.gif')";
            document.getElementById("content").style.display = "inline";

            mensaje = mensaje + '<br><br>' + document.forma.CountryExactus.value + '-' + document.forma.OID.value + '-' + document.forma.Movimiento.value + '<br><br> HBLNumber: ' + document.forma.HBLNumber.value;

            //if (document.forma.Pedido_Erp.value != '')
            //mensaje = mensaje + '<br><br>PedidoERP: ' + document.forma.Pedido_Erp.value;

            document.getElementById("Leyenda").innerHTML = mensaje;

        } else return false;

    }

</script>



<style type="text/css">

	body {
		background-color:rgb(194,223,239);	
		background-image:url(img/loader.gif);
        background-repeat: no-repeat;
        background-attachment: fixed;
        background-position: center;  	
	}
	
	.Textos {	
		font-family:Arial;
		font-size:medium;
		color:Navy;
	}
	
	#content {
        margin:auto;
        height: 173px;
        width: 173px;
        position:fixed;
        top:0;
        bottom:0;
        left:0;
        right:0;
        background:rgb(194,223,239);
        vertical-align:middle;
        text-align:center;
   }
   
   th  
   {
       font-family:Arial;
       background-color:Navy;
       color:White;
       }
	
	td  
	{	    
	    font-family:Arial;
        color:Navy;
    
        
	}
	
	input[type=text]
    {
        width:100%; 
	    background-color:silver;
     }
</style>

