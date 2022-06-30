<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"
Dim GroupID, rs, Conn, i, ServiceID, aList1Values, CountList1Values, aList2Values, CountList2Values, Country
Dim Transport, InternationalLocal, IntercompanyFilter, Filter

	GroupID = CheckNum(Request("GID"))
	ServiceID = Request("ServiceID")
	CountList1Values=-1
	CountList2Values=-1
    'Obteniendo el pais de Operacion del Usuario
    Country = SetCountryBAW(Session("OperatorCountry"))
    'Transporte =3 Consolidado, 4=Express, 5=Local
    Transport = CheckNum(Request("T"))
	InternationalLocal = CheckNum(Request("IL"))
    IntercompanyFilter = CheckNum(Request("INTR"))


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

        dim SQLQuery
        SQLQuery = "SELECT tipo, descripcion FROM tmp_tipo_documento_exactus WHERE subtipo = 0 ORDER BY tipo" 
        'response.write SQLQuery & "<br>"
        OpenConn2 Conn
        Set rs = Conn.Execute(SQLQuery)

        If Not rs.EOF Then
		    aList1Values = rs.GetRows
		    CountList1Values = rs.RecordCount-1
	    End If
	    CloseOBJ rs

        if ServiceID <> "" then
            SQLQuery = "SELECT subtipo, descripcion FROM tmp_tipo_documento_exactus WHERE tipo = '" + ServiceID + "' AND subtipo <> 0 ORDER BY subtipo"  
            Set rs = Conn.Execute(SQLQuery)
            'response.write SQLQuery & "<br>"
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

        top.opener.document.forma.elements["SVI_" + Pos].value = document.forma.ServiceID.value;
        top.opener.document.forma.elements["SVN_" + Pos].value = document.forma.ServiceID.options[document.forma.ServiceID.selectedIndex].text;
        top.opener.document.forma.elements["I_" + Pos].value = document.forma.ItemID.value;
        top.opener.document.forma.elements["N_" + Pos].value = document.forma.ItemID.options[document.forma.ItemID.selectedIndex].text;
        top.opener.ValidarDoble(Pos);
        top.close();
    }	 	 
</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="JavaScript:self.focus()">
	<FORM name="forma" action="Search_TipoDoc.asp" method="post" target=_self>
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
    <INPUT name="T" type=hidden value="<%=Transport%>">
    <INPUT name="IL" type=hidden value="<%=CheckNum(Request("IL"))%>">
    <INPUT name="INTR" type=hidden value="<%=CheckNum(Request("INTR"))%>">
    <INPUT name="PG" type=hidden value="<%=CheckNum(Request("PG"))%>">
	<br>
	<TABLE cellspacing=0 cellpadding=2 width=100% align=center>
		<TR>
		<TD class=label align=center colspan="2"><b>Tipos de Documentos</b></TD>
		</TR> 
		<TR>
		<TD class=label align=right><b>Tipo:</b></TD>
		<TD class=label align=left>
		<select name="ServiceID" id="Tipo" class=label onChange="document.forma.submit();">
			<option value='-1'>Seleccionar</option>
			<%for i=0 to CountList1Values%>
			<option value='<%=aList1Values(0,i)%>'><%=aList1Values(1,i)%></option>
			<%next%>
		</select>
		<INPUT name="N" type=hidden value="<%=Request("N")%>">
		</TD>
		</TR> 
		<%if ServiceID <> "" then%>
		<TR>
		<TD class=label align=right><b>SubTipo:</b></TD>
		<TD class=label align=left>
		<select name="ItemID" id="SubTipo" class=label>
			<option value='-1'>Seleccionar</option>
			<%for i=0 to CountList2Values%>
			<option value='<%=aList2Values(0,i)%>'><%=aList2Values(1,i) & " (" & aList2Values(0,i) & ")" %></option>
			<%next%>
		</select>
		</TD>
		</TR> 
		<%end if%>
	</TABLE>
	<TABLE cellspacing=0 cellpadding=2 width=100%>
	<TR>
		 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:Asign();" value="&nbsp;&nbsp;Asignar&nbsp;&nbsp;" class=label></TD>
	</TR>
	</TABLE>
	</FORM>
<script>
    selecciona('forma.ServiceID', '<%=ServiceID%>');
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
