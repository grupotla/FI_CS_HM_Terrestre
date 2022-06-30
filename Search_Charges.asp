<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"
Dim GroupID, rs, Conn, i, ServiceID, ItemID, TC, aList1Values, CountList1Values, aList2Values, CountList2Values, aList3Values, CountList3Values, Country, txtbusqueda, RegimenID
Dim Transport, InternationalLocal, IntercompanyFilter, Filter, esquema, impex, SQLQuery, ObjectID
	GroupID = CheckNum(Request("GID"))
	ServiceID = CheckNum(Request("ServiceID"))
	ItemID = CheckNum(Request("ItemID"))
	CountList1Values=-1
	CountList2Values=-1
    CountList3Values=-1
    'Obteniendo el pais de Operacion del esquema
    Country = SetCountryBAW(Session("OperatorCountry"))
    'Transporte =3 Consolidado, 4=Express, 5=Local
    Transport = CheckNum(Request("T"))
	InternationalLocal = CheckNum(Request("IL"))
    IntercompanyFilter = CheckNum(Request("INTR"))
    txtbusqueda = Trim(Request("txtbusqueda"))
    esquema = Trim(Request("esquema"))
    impex = Trim(Request("impex"))
    RegimenID = Request("RegimenID")
    TC = Request("TC")

    if InternationalLocal = 0 then
        InternationalLocal = "1,2"
    end if

    if IntercompanyFilter = 1 then
        IntercompanyFilter = " and c.id_servicio<>14 "
    else
        IntercompanyFilter = ""
    end if    

    'Filtrado de Rubros, solo en la pagina de cargos al cliente
    'PG=0=Costos Cliente, PG=1=Cargos Cliente, PG=2=Cargos Intercompany
    if CheckNum(Request("PG"))=1 then
        Filter = "and b.id_rubro not in ('486')"
    end if
    
    'Si el usuario tiene pais de Operacion se le presentan los servicios autorizados para cobrar/pagar de ese pais
	if Country>0 then
        'Obteniendo listado de Servicios
	    'Terrestre=3,4,5
        'el query suma 1 al valor InternationalLocal para hacer equivalencia entre INT(0)/LOC(1) Terrstre con INT(1)/LOC(2) Master

        OpenConn2 Conn

        SQLQuery = "select a.id_servicio, c.nombre_servicio " & _
        "from empresas_transportes_servicios as a " & _
        "inner join transporte as b on (b.id_transporte=a.id_transporte) " & _
        "inner join servicios as c on (c.id_servicio=a.id_servicio) " & _
        "inner join empresas as d on (d.id_empresa=a.id_empresa and d.activo=true) " & _
        "where ( d.id_empresa=" & Country & " and a.id_transporte=" & SetTransport(Transport) & " and a.activo=true) and a.cargo_int_loc in (1," & InternationalLocal & ",3) " & _
        IntercompanyFilter & _
        "order by c.nombre_servicio"
        'response.write SQLQuery & "<br>"
        Set rs = Conn.Execute(SQLQuery)
        If Not rs.EOF Then
		    aList1Values = rs.GetRows
		    CountList1Values = rs.RecordCount-1
	    End If
	    CloseOBJ rs

        if ServiceID > 0 then

            SQLQuery = "SELECT er_regimen, er_abreviatura, er_descripcion " & _ 
            "FROM exactus_regimen " & _ 
            "WHERE er_esquema = '" & esquema & "' AND er_status = '1' "
            'response.write SQLQuery & "<br>"
            Set rs = Conn.Execute(SQLQuery)
            If Not rs.EOF Then
		        aList3Values = rs.GetRows
		        CountList3Values = rs.RecordCount-1
	        End If
	        CloseOBJ rs

            if CountList3Values = -1 then
                RegimenID = "IV"
            end if

        else
            RegimenID = ""
        end if


        if RegimenID <> "" and RegimenID <> "-1" then 'or CountList3Values = 999 then

            if txtbusqueda <> "" then
                Filter = Filter & " AND UPPER(a.desc_rubro_es) LIKE '%" & UCase(txtbusqueda) & "%' "
            end if


            if esquema <> "" and impex <> "" then

                SQLQuery = "SELECT c.id_rubro, c.desc_rubro_es, a.codigo, COALESCE(eh_erp_codigo,''), COALESCE(eh_estado,0), COALESCE(eh_otros,'') " & vbCrLf & _ 

                "FROM rubros c " & vbCrLf & _ 

                "INNER JOIN rubros_servicios b ON (c.id_rubro=b.id_rubro AND b.activo = 1) " & vbCrLf & _ 

                "INNER JOIN vw_rubros_combinaciones a ON a.id_servicio = b.id_servicio AND a.id_rubro = c.id_rubro " & vbCrLf & _ 

                "AND a.d1 = '" & impex & "' AND (a.descripcion ILIKE '%terrestre%' OR a.serv ILIKE '%terrestre%') AND a.d3 = '" & RegimenID & "' AND a.d2 = '" & Iif(Transport = 0,"LL",Iif(Transport = 1,"LE","LC")) & "' " & vbCrLf & _ 

                "LEFT JOIN exactus_homologaciones ON codigo = eh_codigo AND eh_erp_categoria = '06' AND eh_estado = 1 AND eh_erp_esquema = '" & esquema & "' " & vbCrLf & _ 

                "WHERE b.in_conta_baw=1 AND c.id_estatus=1 AND b.id_servicio=" & ServiceID & " " & Filter & " ORDER BY c.desc_rubro_es" 

            else
            
                SQLQuery = "SELECT c.id_rubro, c.desc_rubro_es, '', '', 0, '' " & _ 
                "FROM rubros c  " & _ 
                "INNER JOIN rubros_servicios b ON (c.id_rubro=b.id_rubro AND b.activo = 1)  " & _ 
                "WHERE b.in_conta_baw=1 AND c.id_estatus=1 AND b.id_servicio=" & ServiceID & " " & Filter & " ORDER BY c.desc_rubro_es"

            end if

            'Obteniendo listado de rubros asignados al Servicio
            
            'response.write SQLQuery & "<br><br>"
            
            Set rs = Conn.Execute(SQLQuery)
		    If Not rs.EOF Then
			    aList2Values = rs.GetRows
			    CountList2Values = rs.RecordCount-1
		    End If
		    CloseOBJ rs
	    end if
        CloseOBJ Conn
    end if

%>
 
<HTML><HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="javascript">
    function Asign() {
        var Pos = document.forma.N.value;
        if (!valSelec(document.forma.ServiceID)) { return (false) };
        if (!valSelec(document.forma.ItemID)) { return (false) };

        var result = CheckDoble(document.forma.ServiceID.value, document.forma.ItemID.value, Pos);

        if (!result) return false;  

        top.opener.document.forma.elements["SVI" + Pos].value = document.forma.ServiceID.value;
        top.opener.document.forma.elements["SVN" + Pos].value = document.forma.ServiceID.options[document.forma.ServiceID.selectedIndex].text;
        top.opener.document.forma.elements["I" + Pos].value = document.forma.ItemID.value;

        var rubro_text = document.forma.ItemID.options[document.forma.ItemID.selectedIndex].text;
        var rubro_des = rubro_text.split('[');
        top.opener.document.forma.elements["N" + Pos].value = rubro_des[0] + '(' + document.forma.ItemID.value + ')';

        try {
            if (document.forma.RegimenID)
                top.opener.document.forma.elements["R" + Pos].value = document.forma.RegimenID.options[document.forma.RegimenID.selectedIndex].value;
        } catch (err) {
            top.opener.document.forma.elements["R" + Pos].value = ''; // '<%=RegimenID%>';        
        }

        //top.opener.ValidarDoble(Pos);
        top.close();
    }


    function CheckDoble(ServiceID, RubID, Pos) {

        var forma = top.opener.document.forma;
        var chk = forma.elements['CHK'];
        var v1 = 0, sigue = false;
        var i;


        for (i = 0; i < chk.length; i++) {

            //alert(ServiceID + ' ' + RubID);

            if ((forma.elements["SVI" + i].value == ServiceID) && (forma.elements["I" + i].value == RubID)) {

                
                //alert(forma.elements["SVI" + i].value + ' ' +
                    //forma.elements["I" + i].value + ' ' +
                    //forma.elements["DTY" + i].value + ' ' + '');
                    //forma.elements["PID" + i].value + ' ' +
                    //forma.elements["PER" + i].value);


                sigue = false;
                if (forma.elements["PID" + i] && forma.elements["PER" + i]) { //existen los campos

                    //alert(forma.elements["SVI" + i].value + ' ' +
                      //  forma.elements["I" + i].value + ' ' +
                        //forma.elements["DTY" + i].value + ' ' + 
                        //forma.elements["PID" + i].value + ' ' +
                        //forma.elements["PER" + i].value);

                    if ((forma.elements["PID" + i].value != 0) && (forma.elements["PER" + i].value != "")) { //si ya tiene pedido
                        sigue = false;
                    }

                } else {
                    //sigue = false;
                }


                if (forma.elements["DTY" + i].value != 10) { //10 si no esta facturado entra a validar

                    //alert(sigue);
                    if (sigue == true) { //si ya tiene pedido


                    } else {
                        //alert(Pos + ' ' + i);
                        if (Pos == i) { //si es el mismo si puede seleccionarse nuevmente

                        } else {

                            if (forma.elements["TC" + i].value == document.forma.TC.value) { // forma de pago

                                alert("No puede repetir el mismo Rubro y Servicio (Forma Pago) si el anterior no ha sido facturado");

                                return (false);
                            }
                        }
                    }

                }

            }

        }

        return true;

    }


</SCRIPT>


<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="JavaScript:self.focus()">
	<FORM name="forma" action="Search_Charges.asp" method="post" target=_self>
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
    <INPUT name="OID" type=hidden value="<%=ObjectID%>">
    <INPUT name="IL" type=hidden value="<%=CheckNum(Request("IL"))%>">
    <INPUT name="T" type=hidden value="<%=Transport%>">
    <INPUT name="INTR" type=hidden value="<%=CheckNum(Request("INTR"))%>">
    <INPUT name="PG" type=hidden value="<%=CheckNum(Request("PG"))%>">
    <INPUT name="esquema" type=hidden value="<%=esquema%>">
    <INPUT name="impex" type=hidden value="<%=impex%>">
	<INPUT name="N" type=hidden value="<%=Request("N")%>">
	<INPUT name="TC" type=hidden value="<%=TC%>">
	<br>
	<TABLE cellspacing=0 cellpadding=2 width=100% align=center border=0>
		<TR>
		<TD class=label align=center colspan="2"><b>Rubros</b></TD>
		</TR> 
		<TR>
		<TD class=label align=right width="120px"><b>Servicio:</b></TD>
		<TD class=label align=left>
		<select name="ServiceID" id="Servicio" class=label onChange="document.forma.submit();">
			<option value='-1'>Seleccionar</option>
			<%for i=0 to CountList1Values%>
			<option value='<%=aList1Values(0,i)%>'><%=aList1Values(1,i) & " (" & aList1Values(0,i) & ")"%></option>
			<%next%>
		</select>


		<%'="(" & ServiceID & ")(" & CountList3Values & ")(" & RegimenID & ")<br>" %>

		</TD>
		</TR> 

		<%if ServiceID > -1 and CountList3Values > -1 then%>
		    <TR>
		    <TD class=label align=right width="120px"><b>Regimen:</b></TD>
		    <TD class=label align=left>
		        <select name="RegimenID" id="RegimenID" class=label onChange="document.forma.submit();">
			        <option value='-1'>Seleccionar</option>
			        <%for i=0 to CountList3Values%>
			        <option value='<%=aList3Values(1,i)%>'><%=aList3Values(1,i) & " - " & aList3Values(2,i) %></option>
			        <%next%>
		        </select>
		    </TD>
		    </TR> 
        <% else %>

            <INPUT name="RegimenID" type=hidden value="<%=RegimenID%>">

        <%end if%>


		<%if RegimenID <> "" and RegimenID <> "-1" then%>
        <!--
		<TR>
		<TD class=label align=right><b>Busqueda:</b></TD>
		<TD class=label align=left>
            <input type="text" id="txtbusqueda" name="txtbusqueda" value="<%=txtbusqueda%>" />

            <input type="submit" value="Buscar" />

            <input type="submit" value="Limpiar" onclick="document.getElementById('txtbusqueda').value = '';" />
		</TD>
		</TR> 
        -->
		<TR>
		<TD class=label align=right><b>Rubro:</b></TD>
		<TD class=label align=left>
		<select name="ItemID" id="Rubro" class=label style="width:300px">
			<option value='-1'>Seleccionar</option>
			<%for i=0 to CountList2Values%>
			<option value='<%=aList2Values(0,i)%>' 
            <%
            if aList2Values(3,i) <> "" then
                response.write " style='background-color:rgb(255,241,193)' "
            end if
            %>
            >
            <%
            response.write aList2Values(1,i) 

            if aList2Values(2,i) <> "" then
                response.write " [" & aList2Values(2,i) & "]"
            end if
            
            if aList2Values(3,i) <> "" then
                response.write " " & aList2Values(3,i) & ""
            else
                response.write " NO HOMOLOGADO"
            end if

            if aList2Values(5,i) <> "" then
                response.write " (" & aList2Values(5,i) & ")"
            end if
            %>
            </option>
			<%next%>
		</select>
		</TD>
		</TR> 
		<%end if%>

	<TR>
		 <td></td><TD class=label align=left><INPUT name=enviar type=button onClick="JavaScript:Asign();" value="&nbsp;&nbsp;Asignar&nbsp;&nbsp;" class=label></TD>
	</TR>
	</TABLE>
	</FORM>
<script>
    selecciona('forma.ServiceID', '<%=ServiceID%>');
    selecciona('forma.RegimenID', '<%=RegimenID%>');
    selecciona('forma.ItemID', '<%=ItemID%>');
</script>
</BODY>
</HTML>
<%
    Set aList1Values = Nothing
    Set aList2Values = Nothing
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>
