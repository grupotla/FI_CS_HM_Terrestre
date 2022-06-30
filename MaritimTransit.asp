<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
Server.ScriptTimeout=360000
if Session("OperatorID") <> 0 then
Checking "0|1|2"
Dim Action, Conn, rs, CountListValues, aListValues, i, j, Countries(15), Country, CountryStart, CountryEnd
Dim DateFrom, DateTo, HTMLResult, SQLQuery, CountryName, BLType, TransitDays, ArrivalDate

    Action = Request.Form("Action")

    'Busqueda de los Resultados
    if Action=1 then

        Countries(0) = "ventas_gt"
	    Countries(1) = "ventas_sv"
	    Countries(2) = "ventas_hn"
	    Countries(3) = "ventas_ni"
	    Countries(4) = "ventas_ni_grh"
	    Countries(5) = "ventas_cr"
	    Countries(6) = "ventas_pa"
        Countries(7) = "ventas_bz"
        Countries(8) = "ventas_mx"
        Countries(9) = "ventas_gtltf"
        Countries(10) = "ventas_svltf"
        Countries(11) = "ventas_hnltf"
        Countries(12) = "ventas_niltf"
        Countries(13) = "ventas_crltf"
        Countries(14) = "ventas_paltf"
        Countries(15) = "ventas_bzltf"

        BLType = Request.Form("BLType")
        Country = CheckNum(Request.Form("Countries"))

        'Obteniendo las fechas y convirtiendolas al formato yyyy-mm-dd
        DateFrom = replace(Request.Form("DateFrom"),"/","-",1,-1)
        DateTo = replace(Request.Form("DateTo"),"/","-",1,-1)
        'response.write(DateFrom & "<br>")
        'response.write(DateTo & "<br>")
        
        if Country = "-1" then
            'Buscara en todos los paises
            CountryStart=0
            CountryEnd=15
        else
            'Buscara solo en un pais
            CountryStart=Country
            CountryEnd=Country
        end if

        'Busqueda de Resultados
        For i=CountryStart to CountryEnd
            'del nombre ventas_gt, se extrae el pais ISO = GT
            CountryName = Ucase(Mid(Countries(i),8,6))

            select Case BLType
            Case 0 'FCL
                SQLQuery = "select c.no_contenedor,b.no_bl,cc.nombre_cliente,b.etd,b.eta,b.fecha_arribo,b.fecha_descarga,b.id_pais_final2,b.id_origen_carga " & _                  
                    "from bl_completo as b " & _
                        "left join dblink('dbname=master-aimar port=5432 host=" & TrafficDBServer & "  user=dbmaster password=aimargt','select id_cliente,nombre_cliente from clientes') " & _
                        "        as cc(id_cliente int8 , nombre_cliente varchar ) on (cc.id_cliente=b.id_cliente) " & _
                        "inner join contenedor_completo c on(c.bl_id=b.bl_id and c.activo=true) " & _
                        "where (b.activo and b.import_export=true) and b.eta between '" & DateFrom & "' AND '" & DateTo & "' " & _
                        "AND trim(b.id_pais_final2)<>'' AND b.en_intermodal=true " & _
                        "order by b.fecha_arribo,b.bl_id,c.no_contenedor"
            Case 1 'LCL (sin y con division)
                SQLQuery = "select vc.no_contenedor,b.no_bl,c.nombre_cliente,v.etd,v.eta,v.fecha_arribo,vc.fecha_descarga,b.id_pais_final2,v.pais_orig_id " & _
                            "from bill_of_lading b " & _
                            "left join dblink('dbname=master-aimar port=5432 host=" & TrafficDBServer & "  user=dbmaster password=aimargt','select id_cliente,nombre_cliente from clientes') " & _
                            "        as c(id_cliente int8 , nombre_cliente varchar ) on (c.id_cliente=b.id_cliente), " & _
                            "viaje_contenedor vc, " & _
                            "viajes v " & _
                            "where " & _
                            "b.dividido=false AND " & _
                            "b.activo AND " & _
                            "vc.activo AND " & _
                            "v.activo AND v.import_export=true and " & _
                            "b.viaje_contenedor_id = vc.viaje_contenedor_id AND " & _
                            "vc.viaje_id = v.viaje_id AND v.eta BETWEEN '" & DateFrom & "' and '" & DateTo & "' " & _
                            "AND trim(b.id_pais_final2)<>'' AND b.en_intermodal=true " & _
                            " UNION " & _
                            "select vc.no_contenedor,b.no_bl || ' (' || db.no_bl || ')',c.nombre_cliente,v.etd,v.eta,v.fecha_arribo,vc.fecha_descarga,b.id_pais_final2,v.pais_orig_id " & _
                            "from bill_of_lading b " & _
                            "left join dblink('dbname=master-aimar port=5432 host=" & TrafficDBServer & "  user=dbmaster password=aimargt','select id_cliente,nombre_cliente from clientes') " & _
                            "        as c(id_cliente int8 , nombre_cliente varchar ) on (c.id_cliente=b.id_cliente), " & _
                            "viaje_contenedor vc, " & _
                            "viajes v, " & _
                            "divisiones_bl db " & _
                            "where " & _
                            "b.dividido=true AND " & _
                            "b.activo AND " & _
                            "vc.activo AND " & _
                            "v.activo AND v.import_export=true and " & _
                            "db.bl_asoc=b.bl_id and " & _
                            "b.viaje_contenedor_id = vc.viaje_contenedor_id AND " & _
                            "vc.viaje_id = v.viaje_id AND v.eta BETWEEN '" & DateFrom & "' and '" & DateTo & "' " & _
                            "AND trim(b.id_pais_final2)<>'' AND b.en_intermodal=true " & _
                    "order by etd,eta,fecha_arribo,fecha_descarga,id_pais_final2"
            Case 2 'ROs
                SQLQuery = "select '',b.routing,cc.nombre_cliente,'','',b.fecha,'',b.id_pais_destino,b.id_pais_origen  " & _
                    "from routings b, clientes cc " & _
                        "where b.id_cliente=cc.id_cliente and b.activo=true and b.borrado=false and b.fecha between '" & DateFrom & "' AND '" & DateTo & "' " & _
                        "and b.id_pais='" & SetCountryRO(CountryName)  & "' "  & _
                        "and b.id_routing_type=2 and trim(b.id_pais_destino)<>'' and id_transporte in (4,5) " & _
                        "order by b.fecha,b.id_routing"
            Case 3 'MX
                SQLQuery = "select '',BLs,Clients,'','',DischargeDate,'',CountriesFinalDes,CountryOrigen  " & _
                    "from BLDetail " & _
                        "where Expired=0 and ExType=8 and DischargeDate is not null " & _
                        "and STR_TO_DATE(DischargeDate, '%d/%m/%Y') between '" & DateFrom & "' AND '" & DateTo & "' " & _
                        "and Countries='" & SetCountryRO(CountryName)  & "' "  & _
                        "order by DischargeDate,BLDetailID "
            End Select
            'response.write SQLQuery & "<br><br>"
	    
            CountListValues = -1

                Select Case BLType
                Case 0,1
                    openConnOcean Conn, Countries(i)
                Case 2
                    openConn2 Conn
                Case 3
                    OpenConn Conn
                End Select
            
                Set rs = Conn.Execute(SQLQuery)
                if Not rs.EOF then
                    aListValues = rs.GetRows
			        CountListValues = rs.RecordCount-1
                end if
            CloseOBJ Conn
            
            'Dibujando los resultados
            for j=0 to CountListValues
                ArrivalDate = ""
                OpenConn Conn
                    'response.write("select a.BLArrivalDate from BLs a, BLDetail b where a.BLID=b.BLID and b.BLs='" & aListValues(1,j) & "' and substr(a.CountryDes,1,2)='" & Left(Ucase(aListValues(7,j)),2) & "'")
                    set rs = Conn.Execute("select a.BLArrivalDate from BLs a, BLDetail b where a.BLID=b.BLID and b.BLs='" & aListValues(1,j) & "' and substr(a.CountryDes,1,2)='" & Left(Ucase(aListValues(7,j)),2) & "'")
                    if Not rs.EOF then
                        ArrivalDate = rs(0)
                    End if

                    'En el caso de RO (Carga General) y CIF (MX) se toma la fecha de salida, la fecha de despacho de la CP Master Inicial
                    Select Case BLType
                    Case 2,3
                        CloseOBJ rs
                        set rs = Conn.Execute("select a.BLDispatchDate from BLs a, BLDetail b where a.BLID=b.BLID and b.BLs='" & aListValues(1,j) & "' and a.CountryDep='" & Ucase(aListValues(8,j)) & "'")
                        if Not rs.EOF then
                            aListValues(5,j) = rs(0)
                        else 
                            aListValues(5,j) = ""
                        End if
                    End Select
                CloseOBJs rs, Conn                

                Select Case BLType
                Case 0,1
                    'Calculando los Dias en Transito
                    if ArrivalDate<>"" and aListValues(5,j)<>"" then
                        TransitDays = DateDiff("d",ConvertDate(aListValues(5,j),2),ConvertDate(ArrivalDate,3)) '& "-" & ArrivalDate
                        ArrivalDate = ConvertDate(ArrivalDate,1)
                    else
                        TransitDays = ""
                    end if
                    HTMLResult = HTMLResult & "<tr><td class=label>" & aListValues(0,j) & "</td> " & _
                        "<td class=label>" & aListValues(1,j) & "</td> " & _
                        "<td class=label>" & aListValues(2,j) & "</td> " & _
                        "<td class=label>" & ConvertDate2(aListValues(3,j),4) & "</td> " & _
                        "<td class=label>" & ConvertDate2(aListValues(4,j),4) & "</td> " & _
                        "<td class=label>" & ConvertDate2(aListValues(5,j),4) & "</td> " & _
                        "<td class=label>" & ArrivalDate & "</td> " & _
                        "<td class=label>" & TransitDays & "</td> " & _
                        "<td class=label>" & ConvertDate2(aListValues(6,j),4) & "</td> " & _
                        "<td class=label>" & CountryName & "</td> " & _
                        "<td class=label>" & Ucase(aListValues(7,j)) & "</td></tr>"
                Case 2,3
                    'Calculando los Dias en Transito
                    if ArrivalDate<>"" and aListValues(5,j)<>"" then
                        TransitDays = DateDiff("d",ConvertDate(aListValues(5,j),3),ConvertDate(ArrivalDate,3))
                        ArrivalDate = ConvertDate(ArrivalDate,1)
                    else
                        TransitDays = ""
                    end if
                    HTMLResult = HTMLResult & "<tr><td class=label>" & aListValues(1,j) & "</td> " & _
                        "<td class=label>" & aListValues(2,j) & "</td> " & _
                        "<td class=label>" & aListValues(5,j) & "</td> " & _
                        "<td class=label>" & ArrivalDate & "</td> " & _
                        "<td class=label>" & TransitDays & "</td> " & _
                        "<td class=label>" & aListValues(8,j) & "</td> " & _
                        "<td class=label>" & Ucase(aListValues(7,j)) & "</td></tr>"
                end Select
            Next
        Next
    end if
%>

<HTML><HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="javaScript">
    function abrir(Label) {
        var DateSend, Subject;
        
        try {

            DateSend = document.forma(Label).value;

        } catch (e) {
            //DateSend = document.getElementById(Label).value;

            //alert(label);

            var labelid = SetLabelID(Label)
            DateSend = document.getElementById(labelid).value;
        }
        Subject = '';
        window.open('Agenda.asp?Action=1&Label=' + Label + '&DateSend=' + DateSend + '&Subj=' + Subject, 'Seleccionar', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,width=170,height=160,top=250,left=250');
    }

    function SetLabelID(Label) {
		var LabelID = "";
		if (Label == "DateFrom") {
			LabelID = "Fecha Inicial";
		} 		
		if (Label == "DateTo") {
			LabelID = "Fecha Final";
		} 	
        

		return LabelID;
	}

    function validar() {
        if (!valTxt(document.forma.DateFrom, 1, 3)) { return (false) };
        if (!valTxt(document.forma.DateTo, 3, 5)) { return (false) };
        document.forma.Action.value = 1;
        document.forma.submit();
    }	 	 
</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="JavaScript:self.focus()">
	<%if Action=0 then%>
    <FORM name="forma" action="MaritimTransit.asp" method="post" target=_self>
	<INPUT name="Action" type=hidden value="0">
	<TABLE cellspacing=0 cellpadding=2 width=100% align=center>
		<TR>
		<TD class=label align=center colspan=2><br><br><b>Carga en Transito:</b><br><br></TD>
		</TR>
	    <TR>
	    <TD class=label align=right width=40%><b>Tipo de Carga:</b></TD>
	    <TD class=label align=left width=60%>
	    <select name="BLType" class=label>
		    <!--<option value=0>FCL</option>-->
		    <option value=1>LCL</option>
            <option value=2>RO</option>
            <option value=3>CIF</option>
	    </select>
	    </TD>
	    </TR>
	      <TR>
		    <TD width=40% class=label align=right valign=top><b>Rango de fechas:</b><br>(dd-mm-yyyy)</TD>
		    <TD width=60% class=label align=left>Desde:<br><INPUT  readonly="readonly" name="DateFrom" id="Fecha Inicial" type=text value="" size=23 maxLength=19 class=label>&nbsp;
			    <INPUT type=button value="Seleccionar" onClick="JavaScript:abrir('DateFrom');" class=label><br>
			    Hasta:<br><INPUT  readonly="readonly" name="DateTo", id="Fecha Final" type=text value="" size=23 maxLength=19 class=label>&nbsp;
			    <INPUT type=button value="Seleccionar" onClick="JavaScript:abrir('DateTo');" class=label><br>
		    </TD>
	      </TR>
	    <TR>
	    <TD class=label align=right width=40%><b>Pais:</b></TD>
	    <TD class=label align=left width=60%>
	    <select name="Countries" class=label>
            <option value='-1'>TODOS</option>
            <option value='0'>GUATEMALA</option>
            <option value='1'>EL SALVADOR</option>
            <option value='2'>HONDURAS</option>
            <option value='3'>NICARAGUA</option>
            <option value='4'>NICARAGUA (GRH)</option>
            <option value='5'>COSTA RICA</option>
            <option value='6'>PANAMA</option>
            <option value='7'>BELICE</option>
            <option value='8'>MEXICO</option>
            <option value='9'>LATIN FREIGHT GT</option>
            <option value='10'>LATIN FREIGHT SV</option>
            <option value='11'>LATIN FREIGHT HN</option>
            <option value='12'>LATIN FREIGHT NI</option>
            <option value='13'>LATIN FREIGHT CR</option>
            <option value='14'>LATIN FREIGHT PA</option>
            <option value='15'>LATIN FREIGHT BZ</option>
	    </select>
	    </TD>
	    </TR>
	    </TABLE>
		<TABLE cellspacing=0 cellpadding=2 width=400 align=center>
		<TR>
		<TD colspan="2" class=label align=center>
			<TABLE cellspacing=0 cellpadding=2 width=200>
			<TR>
				<TD class=label align=center colspan=2><INPUT name=enviar onclick="validar();" type=button value="&nbsp;&nbsp;Buscar&nbsp;&nbsp;" class=label></TD>
			</TR>
			</TABLE>
		<TD>
		</TR>
	</TABLE>
	</FORM>
    <%else %>
    <TABLE cellspacing=5 cellpadding=2 width=100% align=center>
    <%Select Case BLType
        Case 0,1%>
    <TR>
		<TD class=label colspan=9 align=left><b>Carga Maritima General en Transito:</b></TD>		
	</TR>
    <TR>
		<TD class=titlelist align=left><b>Contenedor:</b></TD>
		<TD class=titlelist align=left><b>BL:</b></TD>
		<TD class=titlelist align=left><b>Cliente:</b></TD>
		<TD class=titlelist align=left><b>ETD:</b></TD>
		<TD class=titlelist align=left><b>ETA:</b></TD>
		<TD class=titlelist align=left><b>Arribo Puerto:</b></TD>
		<TD class=titlelist align=left><b>Arribo Destino Final:</b></TD>
		<TD class=titlelist align=left><b>Dias Transito:</b></TD>
		<TD class=titlelist align=left><b>Descarga:</b></TD>
		<TD class=titlelist align=left><b>Procedencia:</b></TD>
		<TD class=titlelist align=left><b>Destino Final:</b></TD>			
	</TR>
     <%Case Else%>
     <TR>
		<TD class=label colspan=9 align=left><b>Carga General en Transito:</b></TD>		
	</TR>
    <TR>
		<TD class=titlelist align=left><b>RO:</b></TD>
		<TD class=titlelist align=left><b>Cliente:</b></TD>
		<TD class=titlelist align=left><b>Despacho:</b></TD>
		<TD class=titlelist align=left><b>Arribo Destino Final:</b></TD>
		<TD class=titlelist align=left><b>Dias Transito:</b></TD>
		<TD class=titlelist align=left><b>Procedencia:</b></TD>
		<TD class=titlelist align=left><b>Destino Final:</b></TD>			
	</TR>
     <%End Select%>
    <%=HTMLResult %>
    </TABLE>
    <%end if %>
</BODY>
</HTML>
<%
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>
