<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"
Dim GroupID, Conn, rs, rs1, CountList1Values, CountList2Values, CountList3Values, CountList4Values, ItineraryType, tempo1, tempo2
Dim aList1Values, aList2Values, aList3Values, aList4Values, i

	GroupID = CheckNum(Request("GID")) 'Revisando que el Grupo sea 1 = Categorias, 2 = Noticias , 3 = Mensajes, 4 = Usuarios
	ItineraryType = CheckNum(Request("IT"))
	CountList1Values = -1
	CountList2Values = -1
	CountList3Values = -1
	CountList4Values = -1
%>

<HTML><HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javaScript" src="img/config.js"></SCRIPT>
<!--<SCRIPT language="javaScript" src="img/mainLib.js" ></SCRIPT>-->
<SCRIPT language="javascript" src="img/validaciones.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="javaScript">
	function abrir(Label){
		var DateSend, Subject;
		if (parseInt(navigator.appVersion) < 5) {
			DateSend = document.forma(Label).value;
		} else {
			DateSend = document.getElementById(Label).value;
		}
		Subject = '';	
		window.open('Agenda.asp?Action=1&Label=' + Label + '&DateSend=' + DateSend + '&Subj=' + Subject,'Seleccionar','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,width=170,height=160,top=250,left=250');
	}
	
	function validate(){
		<%if GroupID=17 then%>
		if (!valTxt(document.forma.Week, 1, 3)){return (false)};	
		if (!valSelec(document.forma.Yr)){return (false)};			
        if (document.forma.Week2.value != "") {
            alert("Aviso: Al seleccionar un rango de Semanas el reporte necesitara de varios minutos para procesarse");
        }
	  	<%end if%>
        document.forma.submit();
	}
	 	 
</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="JavaScript:self.focus()">
	<FORM name="forma" action="Search_ResultsAdmin.asp" method="post" target=_self>
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
	<TABLE cellspacing=0 cellpadding=2 width=100% align=center>
	  <%if GroupID<>27 and GroupID<>28 and GroupID<>17 and GroupID<>34 and GroupID<>35 then%>
	  <TR>
		<TD width=40% class=label align=right valign=top><b>Rango de fechas:</b><br>(dd-mm-yyyy)</TD>
		<TD width=60% class=label align=left>Desde:<br><INPUT  readonly="readonly" name="DateFrom" id="DateFrom" type=text value="" size=23 maxLength=19 class=label>&nbsp;
			<INPUT type=button value="Seleccionar" onClick="JavaScript:abrir('DateFrom');" class=label><br>
			Hasta:<br><INPUT  readonly="readonly" name="DateTo", id="DateTo" type=text value="" size=23 maxLength=19 class=label>&nbsp;
			<INPUT type=button value="Seleccionar" onClick="JavaScript:abrir('DateTo');" class=label><br>
		</TD>
	  </TR>
	  <%end if%>
		<%select case GroupID
		Case 1, 12, 13, 14, 15, 16, 17, 25, 29, 34, 35 'BL y Reportes
			if GroupID <>17 and GroupID <> 34  and GroupID <> 35 then
				OpenConn Conn
				'Obteniendo listado de Aduanas
				Set rs = Conn.Execute("select BrokerID, Name, Countries from Brokers where Expired = 0 and Countries in " & Session("Countries") & " order by Name, Countries")
				If Not rs.EOF Then
					aList1Values = rs.GetRows
					CountList1Values = rs.RecordCount-1
				End If
				CloseOBJ rs
			
				'Obteniendo listado de Pilotos
				Set rs = Conn.Execute("select PilotID, Name, License, Countries from Pilots where Expired = 0 and Countries in " & Session("Countries") & " order by Name, Countries")
				If Not rs.EOF Then
					aList2Values = rs.GetRows
					CountList2Values = rs.RecordCount-1
				End If
				CloseOBJ rs
			
				'Obteniendo listado de Cabezales
				Set rs = Conn.Execute("select TruckID, TruckNo, Countries, TruckType, Mark, Model from Trucks where Expired = 0 and Countries in " & Session("Countries") & " order by TruckNo, Countries")
				If Not rs.EOF Then
					aList3Values = rs.GetRows
					CountList3Values = rs.RecordCount-1
				End If
				CloseOBJs rs, Conn
			end if
		%>
		<TR>
		<TD class=label align=right width=40%><b>Tipo de Carta Porte:</b></TD>
		<TD class=label align=left width=60%>
		<select name="BLType" class=label id="Tipo de Carta Porte">
			<%Select Case ItineraryType
			Case 1%>
			<option value=0>CONSOLIDADO</option>
			<option value=1>EXPRESS</option>
			<%Case 2%>
			<option value=2>LOCAL</option>
			<%end Select%>
		</select>
		</TD>
		</TR>
		<%if GroupID=17 then%>
        <TR>
		<TD class=label align=right width=40%><b>Tipo de Reporte:</b></TD>
		<TD class=label align=left width=60%>
		<select name="ReportType" class=label id="Tipo de Reporte">
			<option value=0>CARGA EN CAM</option>
			<option value=1>TIEMPOS DE RUTAS</option>
            <option value=2>PORCENTAJE USO DE RUTAS</option>
            <option value=4>PORCENTAJE USO DE RUTAS ORIGEN</option>
            <option value=3>INGRESO DE STATUS POR USUARIO</option>

		</select>
		</TD>
		</TR>
        <%end if%>
        <%if GroupID<>34 and GroupID<>35 then%>
        <TR>
		<TD class=label align=right width=40%><b>De la Semana</b></TD>
		<TD class=label align=left width=60%><INPUT name="Week" id="De la Semana" type=text value="" size=30 maxLength=50 class=label onKeyUp="res(this,numb);"></TD>
		</TR>        
        <TR>
		<TD class=label align=right width=40%><b>Hasta la Semana</b></TD>
		<TD class=label align=left width=60%><INPUT name="Week2" id="Hasta la Semana" type=text value="" size=30 maxLength=50 class=label onKeyUp="res(this,numb);"></TD>
		</TR>   
        <%end if%>    
        <%if GroupID=17 then
			'Obteniendo listado de Anios
			OpenConn Conn
			Set rs = Conn.Execute("select distinct Year(CreatedDate) as YY from BLs order by YY desc ")
			If Not rs.EOF Then
				aList4Values = rs.GetRows
				CountList4Values = rs.RecordCount-1
			End If
			CloseOBJs rs, Conn
		%>
		<TR>
		<TD class=label align=right width=40%><b>A&ntilde;o</b></TD>
		<TD class=label align=left width=60%>
            <select name="Yr" class=label id="Anio">
                <%for i=0 to CountList4Values%>
                <option value="<%=aList4Values(0,i)%>"><%=aList4Values(0,i)%></option>
                <%next%>
            </select>        
        </TD>
		</TR>        
        <%end if%>
        <%if GroupID <> 17 then%>
		<TR>
		<TD class=label align=right width=40%><b>Numero Carta Porte <%if GroupID=14 or GroupID=15 then%>Hija<%end if%></b></TD>
		<TD class=label align=left width=60%><INPUT name="BLNumber" type=text value="" size=30 maxLength=50 class=label></TD>
		</TR>
        <%end if%>
               
        <%if GroupID = 34 or GroupID = 35 then%>
		<TR>
		<TD class=label align=right width=40%><b>Codigo Referencia</b></TD>
		<TD class=label align=left width=60%><INPUT name="CodeReference" type=text value="" size=30 maxLength=10 class=label  onKeyUp="res(this,numb);"></TD>
		</TR>
        <%end if%>               
         
		<%if GroupID=25 then%>
		<TR>
		<TD class=label align=right width=40%><b>Numero DTI</b></TD>
		<TD class=label align=left width=60%><INPUT name="DTI" type=text value="" size=30 maxLength=50 class=label></TD>
		</TR>
		<%end if%>
		<%if GroupID=14 then%>
		<TR>
		<TD class=label align=right width=40%><b>Numero HBL o RO</b></TD>
		<TD class=label align=left width=60%><INPUT name="HBL" type=text value="" size=30 maxLength=50 class=label></TD>
		</TR>
		<%end if%>
		<%if GroupID <> 14 and GroupID <> 15 and GroupID<>17 and GroupID<>34 and GroupID<>35 then%>
		<TR>
		<TD class=label align=right width=40%><b>Pais de Destino:</b></TD>
		<TD class=label align=left width=60%>
		<select name="CountryDes" class=label>
			<option value=''>Seleccionar</option>
			<%DisplayCountries "", 2%>
		</select>
		</TD>
		</TR>
		<%end if%>
        <%if GroupID <>17 and GroupID<>34 and GroupID<>35 then%>
		<TR>
		<TD class=label align=right width=40%><b>Pais de Origen:</b></TD>
		<TD class=label align=left width=60%>
		<select name="CountryDep" class=label>
			<option value=''>Seleccionar</option>
			<!--#include file=Countries.asp--> 
		</select>
		</TD>
		</TR>
		<TR>
		<TD class=label align=right width=40%><b>Aduana:</b></TD>
		<TD class=label align=left width=60%>
		<select name="BrokerID" class="label">
			<option value="">Seleccionar</option>
			<%
				For i = 0 To CountList1Values
			%>
			<option value="<%=aList1Values(0,i)%>"><%response.write aList1Values(1,i) & " - " & aList1Values(2,i)%></option>
			<%
				Next
			%>
		</select>
		</TD>
		</TR>
		<TR>
		<TD class=label align=right width=40%><b>Conductor:</b></TD>
		<TD class=label align=left width=60%>
		<select name="PilotID" class="label">
			<option value="">Seleccionar</option>
			<%
				For i = 0 To CountList2Values
			%>
			<option value="<%=aList2Values(0,i)%>"><%response.write aList2Values(1,i) & " - " & aList2Values(3,i)%></option>
			<%
				Next
			%>
		</select>
		</TD>
		</TR>
		<TR>
		<TD class=label align=right width=40%><b>Matricula:</b></TD>
		<TD class=label align=left width=60%>
		<select name="TruckID" class="label">
			<option value="">Seleccionar</option>
			<%
				For i = 0 To CountList3Values
			%>
			<option value="<%=aList3Values(0,i)%>"><%response.write aList3Values(1,i) & " " & aList3Values(4,i) & " " & aList3Values(5,i) & " - " & aList3Values(2,i)%></option>
			<%
				Next
			%>
		</select>
		</TD>
		</TR>
        <%end if%>		
		<% Case 7, 8, 10, 21 'Remitente, Embarcadores, Destinatarios, Proveedor, Aduana, Bodega%>
		<TR>
		<TD class=label align=right width=40%><b>Nombre:</b></TD>
		<TD class=label align=left width=60%><INPUT name="Name" type=text value="" size=30 maxLength=50 class=label></TD>
		</TR>
		<TR>
		<TD class=label align=right width=40%><b>Attn:</b></TD>
		<TD class=label align=left width=60%><INPUT name="Attn" type=text value="" size=30 maxLength=50 class=label></TD>
	  	</TR>
		<TR>
		<TD class=label align=right><b>Pais:</TD>
		<TD class=label align=left>
			<select name="Countries" id="Pais" class="label" required>
				<option value="">Seleccionar</option>
				<%if GroupID<>3 then%>
					<%DisplayCountries "", 2%>
				<%else%>
					<!--#include file=Countries.asp--> 
				<%end if%>
			</select>	
		</TD>
		</TR>
		<% Case 2, 3, 4, 11 'Consigners, Agents, Shippers	%>
		<TR>
		<TD class=label align=right><b>Nombre:</b></TD>
		<TD class=label align=left><INPUT TYPE=text class=label name="Name" value="" size=30></TD>
		</TR>
 		<% Case 5, 6 'Piloto, Cabezal
			GetProviders 'Procedimiento para obtener el listado de proveedores de pilotos y cabezales	
		%>
		<TR>
		<TD class=label align=right width=40%><b>Proveedor</b></TD>
		<TD class=label align=left width=60%>
		    <select name="ProviderID" class=label>
		    <option value="">Seleccionar</option>
		    <%
			    For i = 0 To CountList1Values-1
		    %>
		    <option value="<%=aList1Values(0,i)%>" title="<%=aList1Values(1,i) & " - " & aList1Values(3,i) & " - " & aList1Values(4,i)%>"><%=Left(aList1Values(1,i),50) & " - " & aList1Values(2,i)%></option>

		    <%
    		    Next
		    %>
		    </select>
	    </TR>
		<TR>
		<%if GroupID = 5 then%>
		<TD class=label align=right width=40%><b>Nombre:</b></TD>
		<TD class=label align=left width=60%><INPUT name="Name" type=text value="" size=30 maxLength=50 class=label></TD>
		<%else%>
		</TR>
		<TR>
		<TD class=label align=right width=40%><b>Tipo:</b></TD>
		<TD class=label align=left width=60%>
			<select name="TruckType" class=label>
				<option value="">Seleccionar</option>
				<option value="0">CABEZAL</option>
				<option value="1">FURGON</option>
                <option value="2">CAMION</option>
			</select>
		</TD>
		</TR>
		<TR>
		<TD class=label align=right width=40%><b>Matricula:</b></TD>
		<TD class=label align=left width=60%><INPUT name="TruckNo" type=text value="" size=30 maxLength=50 class=label></TD>
		<%end if%>
	  	</TR>
		<TR>
		<TD class=label align=right><b>Pais:</TD>
		<TD class=label align=left>
			<select name="Countries" id="Pais" class="label">
				<option value="">Seleccionar</option>
				<%if GroupID<>3 then%>
					<%DisplayCountries "", 2%>
				<%else%>
					<!--#include file=Countries.asp--> 
				<%end if%>
			</select>	
		</TD>
		</TR>
		<% Case 9 'Productos %>
		<TR>
		<TD class=label align=right><b>Nombre:</TD>
		<TD class=label align=left><INPUT TYPE=text class=label name="Name" value="" size=30></TD>
		</TR>
		<% Case 18 'Solicitud De Movimiento %>
		<TR>
		<TD class=label align=right><b>Valor:</TD>
		<TD class=label align=left><INPUT TYPE=text class=label name="Tax" value="" size=30></TD>
		</TR>
		<% Case 22 'Solicitud De Movimiento %>
		<TR>
		<TD class=label align=right width=40%><b>Semana</b></TD>
		<TD class=label align=left width=60%><INPUT name="Week" type=text value="" size=30 maxLength=50 class=label onKeyUp="res(this,numb);"></TD>
		</TR>
		<TR>
		<TD class=label align=right width=40%><b>Numero Carta Porte Grupo</b></TD>
		<TD class=label align=left width=60%><INPUT name="BLNumber" type=text value="" size=30 maxLength=50 class=label></TD>
		</TR>
		<% Case 23 'Rastreo o Tracking %>
		<TR>
		<TD class=label align=right width=40%><b>Tipo de Carta Porte:</b></TD>
		<TD class=label align=left width=60%>
		<select name="BLType" class=label id="Tipo de Carta Porte">
			<%Select Case ItineraryType
			Case 1%>
			<option value=-1>GRUPO</option>
			<option value=0 selected>CONSOLIDADO</option>
			<option value=1>EXPRESS</option>
            <option value=-2>CIF INGRESO</option>
			<%Case 2%>
			<option value=2>LOCAL</option>
			<%End Select%>
		</select>
		</TD>
		</TR>
		<TR>
		<TD class=label align=right width=40%><b>Semana</b></TD>
		<TD class=label align=left width=60%><INPUT name="Week" type=text value="" size=30 maxLength=50 class=label onKeyUp="res(this,numb);"></TD>
		</TR>
		<TR>
		<TD class=label align=right width=40%><b>Numero Carta Porte</b></TD>
		<TD class=label align=left width=60%><INPUT name="BLNumber" type=text value="" size=30 maxLength=50 class=label></TD>
		</TR>
		<TR>
		<TD class=label align=right width=40%><b>Numero BL/RO</b></TD>
		<TD class=label align=left width=60%><INPUT name="MBL" type=text value="" size=30 maxLength=50 class=label></TD>
		</TR>
		<TR>
		<TD class=label align=right width=40%><b>Consignatario</b></TD>
		<TD class=label align=left width=60%><INPUT name="Name" type=text value="" size=30 maxLength=50 class=label></TD>
		</TR>
        <TR>
		<TD class=label align=right width=40%><b>Exportador</b></TD>
		<TD class=label align=left width=60%><INPUT name="ShipperName" type=text value="" size=30 maxLength=50 class=label></TD>
		</TR>
		<%if ItineraryType=1 then%>
		<TR>
		<TD class=label align=right width=40%><b>Pais de Destino:</b></TD>
		<TD class=label align=left width=60%>
		<select name="CountryDes" class=label>
			<option value=''>Seleccionar</option>
			<%DisplayCountries "", 2%>
		</select>
		</TD>
		</TR>
		<TR>
		<TD class=label align=right width=40%><b>Pais de Origen:</b></TD>
		<TD class=label align=left width=60%>
		<select name="CountryDep" class=label>
			<option value=''>Seleccionar</option>
			<!--#include file=Countries.asp--> 
		</select>
		</TD>
		</TR>
		<%end if%>
		<% Case 26 'Plantilla DTI %>
		<TR>
		<TD class=label align=right width=40%><b>Nombre</b></TD>
		<TD class=label align=left width=60%><INPUT name="Name" type=text value="" size=30 maxLength=50 class=label></TD>
		</TR>
		<TR>
		<TD class=label align=right width=40%><b>Pais de Destino:</b></TD>
		<TD class=label align=left width=60%>
		<select name="Countries" class=label>
			<option value=''>Seleccionar</option>
			<%DisplayCountries "", 2%>
		</select>
		</TD>
		</TR>
		<%Case 27 'Carga Maritima en Transito %>
		<TR>
		<TD class=label align=center colspan=2><br><br><b>Carga en Transito:</b><br><br></TD>
		</TR>
		<TR>
		<TD class=label align=right width=40%><b>Tipo de Carga:</b></TD>
		<TD class=label align=left width=60%>
		<select name="BLType" class=label>
			<option value=0>FCL</option>
			<option value=1>LCL</option>
            <option value=2>AEREO</option>
            <option value=3>INTERMODAL</option>
		</select>
		</TD>
		</TR>
		<TR>
		<TD class=label align=right width=40%><b>House BL/AWB/CP:</b></TD>
		<TD class=label align=left width=60%><INPUT name="HBL" type=text value="" size=30 maxLength=50 class=label></TD>
		</TR>
		<TR>
		<TD class=label align=right width=40%><b>Master BL/AWB/CP:</b></TD>
		<TD class=label align=left width=60%><INPUT name="MBL" type=text value="" size=30 maxLength=50 class=label></TD>
		</TR>
		<TD class=label align=right width=40%><b>Contenedor:</b></TD>
		<TD class=label align=left width=60%><INPUT name="CNT" type=text value="" size=30 maxLength=50 class=label></TD>
		</TR>
		</TR>
		<TD class=label align=right width=40%><b>Pais:</b></TD>
		<TD class=label align=left width=60%>
		<select name="CountriesSearch" class=label>
			<option value=''>Seleccionar</option>
			<%DisplayCountries Request("CTR"), 1%>
		</select>
		<input type="hidden" name="Countries" value="<%=Request("CTR")%>">
        <input type="hidden" name="IT" value="<%=ItineraryType%>">
		</TD>
		</TR>
		<%Case 32 'Marchamos 
			OpenConn Conn
			'listado de bodegas para asignar bolsas de marchamos
			Set rs = Conn.Execute("select WarehouseID, Countries, Name from Warehouses where Expired=0 order by Name, Countries")
			If Not rs.EOF Then
				aList1Values = rs.GetRows
				CountList1Values = rs.RecordCount-1
			End If
			CloseOBJs rs, Conn		
		%>	
		<TR>
		<TD class=label align=right width=40%><b>Bodegas:</b></TD>
		<TD class=label align=left width=60%>
		<select class="label" name="WarehouseID" id="Bodega">
		<option value="0">Seleccionar</option>
		<%		
			For i = 0 To CountList1Values
		%>
		<option value="<%=aList1Values(0,i)%>"><%response.write aList1Values(2,i) & " - " & aList1Values(1,i)%></option>
		<%
   			Next
		%>
		</select>
		</TD>
		</TR>
		<TR>
		<TD class=label align=right width=40%><b>Bolsa</b></TD>
		<TD class=label align=left width=60%><INPUT name="BagValue" type=text value="" size=30 maxLength=50 class=label onKeyUp="res(this,numb);"></TD>
		</TR>


		<% end select %>
		</TABLE>
		<TABLE cellspacing=0 cellpadding=2 width=400 align=center>
		<TR>
		<TD colspan="2" class=label align=center>
			<TABLE cellspacing=0 cellpadding=2 width=200>
			<TR>
				<TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validate();" value="&nbsp;&nbsp;Buscar&nbsp;&nbsp;" class=label></TD>
			</TR>
			</TABLE>
		<TD>
		</TR>
	</TABLE>
	</FORM>
<%if GroupID=3 or GroupID=7 or GroupID=8 or GroupID=21 then%>
<script>
    selecciona('forma.Countries', '<%=SetDefaultCountry%>');
</script>
<%end if%>
</BODY>
</HTML>
<%
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>
