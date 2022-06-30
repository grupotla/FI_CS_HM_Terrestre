<%
Checking "0|1"
'4, 11 Consigners
'3, 20 Shippers
if CountTableValues >= 0 then
	CreatedDate = ConvertDate(aTableValues(1, 0),2)
    CreatedTime = aTableValues(2, 0)
	Expired = aTableValues(3, 0)
	Name = aTableValues(4, 0)
	BillName = aTableValues(5, 0)	
	Address = aTableValues(6, 0)
	AddressID = aTableValues(7, 0)
	Phone1 = aTableValues(8, 0)
	Countries = aTableValues(9, 0)
	Region = aTableValues(10, 0)
	isConsigneer = CheckNum(aTableValues(11, 0))
	isShipper = CheckNum(aTableValues(12, 0))
	BusinessGID = CheckNum(aTableValues(13, 0))
	UserCreate = CheckNum(aTableValues(14, 0))
	UserModify = CheckNum(aTableValues(15, 0))
	CreatedIn = aTableValues(16, 0)
else
	Expired = Request.Form("Expired")
	if Expired = "" then
		Expired = -1
	end if
	Name = Request.Form("Name")
	BillName = Request.Form("BillName")
	Address = Request.Form("Address")
	AddressID = Request.Form("AddressID")
	Phone1 = Request.Form("Phone1")
	Countries = Request.Form("Countries")
	Region = Request.Form("Region")
	isConsigneer = Request.Form("isConsigneer")
	isShipper = Request.Form("isAgent")
	BusinessGID = Request.Form("BGID")
end if
Set aTableValues = Nothing
if Trim(Request.Form("Countries")) <> "" then
	Countries = Trim(Request.Form("Countries"))
	Region = Trim(Request.Form("Region"))
End If

'Monitoreo
if Action=1 or Action=2 then
	OpenConn Conn
	Conn.Execute("insert into Logs (UserID, UserName, LogData, Action) values (" & Session("OperatorID") & ", '" & Session("OperatorName") & "', '" & ObjectID & " - " & Name & " - " & AddressID  & " - " & Countries & " - " & Region & "', " & Action & ")")
	CloseOBJ Conn
end if

OpenConn2 Conn	
	set rs = Conn.Execute("select id_estatus, descripcion from estatus")
	if Not rs.EOF then
		aTableValues = rs.GetRows
		CountTableValues = rs.RecordCount-1
	end if
	CloseOBJ rs
	
	if Countries <> "" then
		Dim aTableRegions, CountTableRegions
		CountTableRegions = -1
		set rs = Conn.Execute("select id_nivel, nivel1 from niveles_geograficos where id_pais='" & Countries & "' order by nivel1")
		if Not rs.EOF then
			aTableRegions = rs.GetRows
			CountTableRegions = rs.RecordCount-1
		end if
		CloseOBJ rs
	end if

	if ObjectID <> 0 then
		set rs = Conn.Execute("select id_telefono, numero_telefono from cli_telefonos where id_cliente=" & ObjectID)
		if Not rs.EOF then
			PhoneID = rs(0)
			Phone2 = rs(1)
		end if
		CloseOBJ rs
	
		set rs = Conn.Execute("select contacto_id, nombres from contactos where id_cliente=" & ObjectID)
		if Not rs.EOF then
			AttnID = rs(0)
			Attn = rs(1)
		end if
		CloseOBJ rs
	
		set rs = Conn.Execute("select no_cuenta, no_iata from clientes_aereo where id_cliente=" & ObjectID)
		if Not rs.EOF then
			AccountNo = rs(0)
			IATANo = rs(1)
		end if
		CloseOBJ rs		

		if UserCreate <> 0 then
			set rs = Conn.Execute("select pw_gecos from usuarios_empresas where id_usuario=" & UserCreate)
			if Not rs.EOF then
				UserCreate = UCASE(rs(0))
			end if
			CloseOBJ rs
		else
			UserCreate = ""
		end if
		
		if UserModify <> 0 then
			set rs = Conn.Execute("select pw_gecos from usuarios_empresas where id_usuario=" & UserModify)
			if Not rs.EOF then
				UserModify = UCASE(rs(0))
			end if
			CloseOBJ rs
		else
			UserModify = ""
		end if
	end if
	if Session("OperatorLevel") = 0 then
		'Obteniendo el listado de Grupos
		set rs = Conn.Execute("select id_grupo, nombre_grupo from grupos where id_estatus=1 order by nombre_grupo")
		if Not rs.EOF then
			aList3Values = rs.GetRows
			CountList3Values = rs.RecordCount-1
		end if
		CloseOBJ rs
	end if
CloseOBJ Conn
%>
<HTML>
<HEAD><TITLE>AWB - Aimar - Administración</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
var ntr = "";
var com = "";

	function validar(Action) {
  		if (Action != 3) {
			if (!valSelec(document.forma.Expired)){return (false)};
			if (!valSelec(document.forma.Countries)){return (false)};
			<%if CountTableRegions >= 0 then%>
			if (!valSelec(document.forma.Region)){return (false)};
			<%end if%>
			if (!valTxt(document.forma.Name, 3, 5)){return (false)};
			if (!valTxt(document.forma.BillName, 3, 5)){return (false)};
			if (!valTxt(document.forma.Address, 3, 5)){return (false)};
			<%if Len(Session("Countries")) > 6 then%>
			if (!valSelec(document.forma.CreatedIn)){return (false)};
			<%end if%>
		}
	    document.forma.Action.value = Action;
		document.forma.submit();			 
	 }
</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="Javascript:self.focus();">
	<TABLE cellspacing=0 cellpadding=2 width=400 align=center>
	<FORM name="forma" action="InsertData.asp" method="post">
	<INPUT name="Action" type=hidden value=0>
	<INPUT name="SO" type=hidden value="<%=SearchOption%>">
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
	<INPUT name="OID" type=hidden value="<%=ObjectID%>">
	<INPUT name="CD" type=hidden value="<%=CreatedDate%>">
	<INPUT name="CT" type=hidden value="<%=CreatedTime%>">
	<INPUT name="AddressT" type=hidden value="<%=Address%>">
	<INPUT name="Phone1T" type=hidden value="<%=Phone1%>">
	<INPUT name="Phone2T" type=hidden value="<%=Phone2%>">
	<INPUT name="AttnT" type=hidden value="<%=Attn%>">
	<INPUT name="AccountNoT" type=hidden value="<%=AccountNo%>">
	<INPUT name="IATANoT" type=hidden value="<%=IATANo%>">
	<INPUT name="AddressID" type=hidden value="<%=AddressID%>">
	<INPUT name="PhoneID" type=hidden value="<%=PhoneID%>">
	<INPUT name="AttnID" type=hidden value="<%=AttnID%>">
		<%if SearchOption = 1 then%>
		<TR><TD class=label align=center colspan="2"><b>Consignatarios:</b></TD></TR> 
		<%end if%>
		<TR><TD class=label align=right><b>Código:</b></TD><TD class=label align=left><%if ObjectID <> 0 then response.write ObjectID & "--" & AddressID End if%></TD></TR> 
		<TR><TD class=label align=right><b>Fecha Creación:</b></TD><TD class=label align=left><%=CreatedDate%></TD></TR> 
        <TR><TD class=label align=right><b>Estado:</b></TD><TD class=label align=left>
			<select name="Expired" id="Estado" class="label">
				<option value="-1">SELECCIONAR</option>
				<%for i=0 to CountTableValues%>
				<option value="<%=aTableValues(0,i)%>"><%=aTableValues(1,i)%></option>
				<%next%>
			</select>	
		<TR><TD class=label align=right><b>Pais:</b></TD><TD class=label align=left colspan=2>
			<select name="Countries" id="Pais" class="label" onChange="Javascript:document.forma.Region.value='';document.forma.submit();">
				<option value="-1">SELECCIONAR</option>
				<!--#include file=Countries.asp-->
			</select>	
		</TD></TR>
		<TR><TD class=label align=right><b>Region:</b></TD><TD class=label align=left colspan=2>
			<select name="Region" id="Region" class="label">
				<option value="-1">SELECCIONAR</option>
				<%if Countries <> "" then
				for i=0 to CountTableRegions%>
				<option value="<%=aTableRegions(0,i)%>"><%=aTableRegions(1,i)%></option>
				<%next
				end if%>
			</select>	
		</TD></TR>
		<TR><TD class=label align=right><b>Nombre:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Name" id="Nombre Destinatario" value="<%=Name%>"></TD></TR> 
		<TR><TD class=label align=right><b>Nombre&nbsp;Facturar:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="BillName" id="Nombre a Facturar" value="<%=BillName%>"></TD></TR> 
		<TR><TD class=label align=right><b>Es Cliente:</b></TD><TD class=label align=left><INPUT name=isConsigneer TYPE=checkbox class=label <%If isConsigneer Then response.write " checked"  End If%>></TD></TR>
		<TR><TD class=label align=right><b>Es Remitente(Shipper):</b></TD><TD class=label align=left><INPUT name=isShipper TYPE=checkbox class=label <%If isShipper Then response.write " checked"  End If%>></TD></TR>
		<%if Session("OperatorLevel") = 0 then%>
		<TR><TD class=label align=right valign="top"><b>Grupo:</b></TD><TD class=label align=left>
		<select name="BGID" class="label">
			<option value="0">Seleccionar</option>
			<%for i=0 to CountList3Values%>
				<option value="<%=aList3Values(0,i)%>"><%=aList3Values(1,i)%></option>
			<%next%>
		</select>
		</TD></TR>
		<%else%>
		<input type="hidden" value="<%=BusinessGID%>" name="BGID">
		<%end if%>
		<TR><TD class=label align=right><b>Dirección:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Address" id="Direccion" value="<%=Address%>" maxlength="250"></TD></TR>
		<TR><TD class=label align=right><b>Telefono 1:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Phone1" id="Telefono 1" value="<%=Phone1%>" maxlength="45"></TD></TR>
		<TR><TD class=label align=right><b>Telefono 2:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Phone2" id="Telefono 2" value="<%=Phone2%>" maxlength="45"></TD></TR>
		<TR><TD class=label align=right><b>No. de Cuenta:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="AccountNo" value="<%=AccountNo%>" maxlength="45"></TD></TR>
		<%if GroupID=8 then%>
		<TR><TD class=label align=right><b>No. IATA:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="IATANo" value="<%=IATANo%>" maxlength="45"></TD></TR>
		<%end if%>
		<TR><TD class=label align=right><b>Attn:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Attn" value="<%=Attn%>" maxlength="45"></TD></TR>
		<TR><TD class=label align=right><b>Creado En:</b></TD><TD class=label align=left colspan=2>
			<%if Len(Session("Countries")) > 6 then%>
			<select name="CreatedIn" id="Pais donde el registro es creado" class="label">
				<option value="-1">Seleccionar</option>
				<%DisplayCountries CreatedIn, 1%>
			</select>	
			<%else%>
				<%=TranslateCountry(CreatedIn)%>
				<input type="hidden" name="CreatedIn" value="<%=SetDefaultCountry%>">
			<%end if%>
		</TD></TR>
		<TR><TD class=label align=right><b>Creado por:</b></TD><TD class=label align=left><%=UserCreate%></TD></TR>
		<TR><TD class=label align=right><b>Modificado por:</b></TD><TD class=label align=left><%=UserModify%></TD></TR>
		<TR>
		<TD colspan="2" class=label align=center>
			<TABLE cellspacing=0 cellpadding=2 width=200>
			<TR>
			<%if ObjectID=0 then%>
					 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(4)" value="&nbsp;&nbsp;Agregar&nbsp;&nbsp;" class=label></TD>
			<%else%>
					 <%if SearchOption = 1 then
					 select case GroupID
					 case 3%>									 
					 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="top.opener.document.forms[0].ShipperData.value = '<%=Name%>\n<%=Address%>\n<%=Phone1%><%if Phone2 <> "" then%> / <%=Phone2%><%end if%><%if Attn <> "" then%>\nATTN: <%=Attn%><%end if%>';top.opener.document.forms[0].ShipperID.value=<%=ObjectID%>;top.opener.document.forms[0].ShipperAddrID.value=<%=AddressID%>;top.close();" value="&nbsp;&nbsp;Asignar&nbsp;&nbsp;" class=label></TD>
					 <%case 20%> 
					 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="top.opener.document.forms[0].Shippers.value = '<%=Name%>';top.opener.document.forms[0].ShippersID.value=<%=ObjectID%>;top.opener.document.forms[0].ShippersAddrID.value=<%=AddressID%>;top.close();" value="&nbsp;&nbsp;Asignar&nbsp;&nbsp;" class=label></TD>
					 <%case 4%>
					 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="top.opener.document.forms[0].ConsignerData.value = '<%=Name%>\n<%=Address%>\n<%=Phone1%><%if Phone2 <> "" then%> / <%=Phone2%><%end if%>';top.opener.document.forms[0].Attn.value='<%if Attn <> "" then%><%=Name%>\nATTN: <%=Attn%><%end if%>';top.opener.document.forms[0].ConsignerID.value=<%=ObjectID%>;top.opener.document.forms[0].ConsignerAddrID.value=<%=AddressID%>;top.opener.CountryConsignee='<%=Countries%>';top.opener.Consignee='<%=Name%>';top.close();" value="&nbsp;&nbsp;Asignar&nbsp;&nbsp;" class=label></TD>
					 <%case 11%>
					 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="top.opener.document.forms[0].ClientsID.value='<%=ObjectID%>';top.opener.document.forms[0].Clients.value='<%=Name%>';top.opener.document.forms[0].AddressesID.value='<%=AddressID%>';top.close();" value="&nbsp;&nbsp;Asignar&nbsp;&nbsp;" class=label></TD>
					 <%end select
					 end if%>
					 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(2)" value="&nbsp;&nbsp;Actualizar&nbsp;&nbsp;" class=label></TD>
					 <!--<TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(3)" value="&nbsp;&nbsp;Eliminar&nbsp;&nbsp;" class=label></TD>-->
			<%end if%>
			</TR>
			</TABLE>
		<TD>
		</TR>
	</FORM>
	</TABLE>
<script>
selecciona('forma.Expired',<%=Expired%>);
selecciona('forma.Countries','<%=Countries%>');
<%if Region <> "" then%>
selecciona('forma.Region','<%=Region%>');
<%else%>
document.forma.Region.focus();
<%end if%>
<%if Session("OperatorLevel")=0 then%>
selecciona('forma.BGID','<%=BusinessGID%>');
<%end if%>
</script>
</BODY>
</HTML>