<%
Checking "0|1"
'Dim IATANo, DefaultVal
if CountTableValues >= 0 then
	CreatedDate = ConvertDate(aTableValues(1, 0),2)
    CreatedTime = aTableValues(2, 0)
	Expired = aTableValues(3, 0)
	Name = aTableValues(4, 0)
	Address = aTableValues(5, 0)
	Phone1 = aTableValues(6, 0)
	Phone2 = aTableValues(7, 0)
	Email = aTableValues(8, 0)
	Attn = aTableValues(9, 0)
	BusinessGID = aTableValues(10, 0)
	UserCreate = CheckNum(aTableValues(11, 0))
	UserModify = CheckNum(aTableValues(12, 0))	
end if

if Session("OperatorLevel") = 0 then
	OpenConn2 Conn	
	'Obteniendo el listado de Grupos
	set rs = Conn.Execute("select id_grupo, nombre_grupo from grupos where id_estatus=1 order by nombre_grupo")
	if Not rs.EOF then
		aList3Values = rs.GetRows
		CountList3Values = rs.RecordCount-1
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
	CloseOBJ Conn
end if

%>
<HTML>
<HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
	function validar(Action) {
  		if (Action != 3) {
			//if (!valSelec(document.forma.Countries)){return (false)};
			if (!valTxt(document.forma.Name, 3, 5)){return (false)};
			//if (!valTxt(document.forma.Address, 3, 5)){return (false)};
			//if (!valEmail(document.forma.Email)){return (false)};
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
		<%if SearchOption = 1 then%>
		<TR><TD class=label align=center colspan="2"><b>Remitentes / Exportadores:</b></TD></TR> 
		<%end if%>
		<TR><TD class=label align=right><b>Código:</b></TD><TD class=label align=left><%if ObjectID <> 0 then response.write ObjectID End if%></TD></TR> 
		<TR><TD class=label align=right><b>Fecha Creación:</b></TD><TD class=label align=left><%=CreatedDate%></TD></TR> 
        <TR><TD class=label align=right><b>Activa:</b></TD><TD class=label align=left><INPUT name=Expired TYPE=checkbox class=label <%If Expired = 1 Then response.write " checked"  End If%>></TD></TR>
		<TR><TD class=label align=right><b>Nombre:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Name" id="Nombre" value="<%=Name%>" maxlength="45" size="50"></TD></TR> 
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
		<TR><TD class=label align=right><b>Dirección:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Address" id="Dirección" value="<%=Address%>" maxlength="60" size="50"></TD></TR>
		<TR><TD class=label align=right><b>Telefono:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Phone1" value="<%=Phone1%>" maxlength="45"></TD></TR>
		<TR><TD class=label align=right><b>Fax:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Phone2" value="<%=Phone2%>" maxlength="45"></TD></TR>
		<TR><TD class=label align=right><b>Contacto:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Attn" id="a quien se puede dirigir (Attn)" value="<%=Attn%>" maxlength="45"></TD></TR>
		<TR><TD class=label align=right><b>Email:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Email" id="Email de Notificaciones" value="<%=Email%>" maxlength="45"></TD></TR>
		<TR><TD class=label align=right><b>Creado por:</b></TD><TD class=label align=left><%=UserCreate%></TD></TR>
		<TR><TD class=label align=right><b>Modificado por:</b></TD><TD class=label align=left><%=UserModify%></TD></TR>
		<TR>
		<TD colspan="2" class=label align=center>
			<TABLE cellspacing=0 cellpadding=2 width=200>
			<TR>
				<%if CountTableValues = -1 then%>
					<TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(4)" value="&nbsp;&nbsp;Agregar&nbsp;&nbsp;" class=label></TD>
				<%else%>
					<%if SearchOption = 1 then
					 select case GroupID
					 case 3%>
					 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="top.opener.document.forms[0].SenderData.value='<%=Name%>\n<%=Address%>\n<%=Phone1%>&nbsp;&nbsp;&nbsp;&nbsp;<%=Phone2%><%if Attn <> "" then response.Write "\nATTN: " & Attn end if%>';top.opener.document.forms[0].SenderID.value=<%=ObjectID%>;top.close();" value="&nbsp;&nbsp;Asignar&nbsp;&nbsp;" class=label></TD>
					 <%case 20%>
					 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="if (top.opener.document.forms[0].Agents.value != '') {ntr='\n'; com='|';} else {ntr=''; com='';};top.opener.document.forms[0].AgentsID.value = top.opener.document.forms[0].AgentsID.value + com + '<%=ObjectID%>';top.opener.document.forms[0].Agents.value = top.opener.document.forms[0].Agents.value + ntr + '<%=Name%>';top.close();" value="&nbsp;&nbsp;Asignar&nbsp;&nbsp;" class=label></TD>
					 <%end select
					 end if%>
					<TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(2)" value="&nbsp;&nbsp;Actualizar&nbsp;&nbsp;" class=label></TD>
					<TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(3)" value="&nbsp;&nbsp;Eliminar&nbsp;&nbsp;" class=label></TD>
				<%end if%>
			</TR>
			</TABLE>
		<TD>
		</TR>
	</FORM>
	</TABLE>
<script>
<%if Session("OperatorLevel")=0 then%>
selecciona('forma.BGID','<%=BusinessGID%>');
<%end if%>
</script>
</BODY>
</HTML>