<%
Checking "0|1|2"
'Dim TaxNo, Address, Phone1, Phone2, AccountNo, Attn, Expired
if CountTableValues >= 0 then
	CreatedDate = ConvertDate(aTableValues(1, 0),2)
    CreatedTime = aTableValues(2, 0)
	BLNumber = aTableValues(3, 0)
	DTI = aTableValues(4, 0)
	Comment4 = aTableValues(5, 0)
end if

Set aTableValues = Nothing
%>
<HTML>
<HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
	function validar(Action) {
	   	if (!valTxt(document.forma.DTI, 3, 10)){return (false)};
		document.forma.Action.value = Action;
		document.forma.submit();			 
	 }
</script>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
	<TABLE cellspacing=0 cellpadding=2 width=400 align=center>
	<FORM name="forma" action="InsertData.asp" method="post">
	<INPUT name="Action" type=hidden value=0>
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
	<INPUT name="OID" type=hidden value="<%=ObjectID%>">
	<INPUT name="CD" type=hidden value="<%=CreatedDate%>">
	<INPUT name="CT" type=hidden value="<%=CreatedTime%>">
		<TR><TD class=label align=right><b>No. Carta Porte:</b></TD><TD class=label align=left><%=BLNumber%></TD></TR> 
		<TR><TD class=label align=right><b>Fecha Creación:</b></TD><TD class=label align=left><%=CreatedDate%></TD></TR> 
		<TR><TD class=label align=right><b>Código:</b></TD><TD class=label align=left><%=ObjectID%></TD></TR> 
		<TR><TD class=label align=right><b>No.DTI:</b></TD><TD class=label align=left>
		<INPUT name="DTI" id="DTI" type=text value="<%=DTI%>" size=23 maxLength=45 class=label></TD></TR>
		<TR><TD class=label align=right><b>Comentario:</b></TD><TD class=label align=left><Textarea name="Comment4" id="Comentario" cols="30" rows="5"><%=Comment4%></Textarea></TD></TR> 
		<TD colspan="2" class=label align=center>
			<TABLE cellspacing=0 cellpadding=2 width=200>
			<TR>
				<%if DTI <> "" then%>
				 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="Javascript:window.open('DTIPrint.asp?GID=<%=GroupID%>&OID=<%=ObjectID%>','DTI','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750,height=600');return false;" value="&nbsp;&nbsp;Previsualizar&nbsp;DTI&nbsp;&nbsp;" class=label></TD>
				<%end if%> 
				 <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(2)" value="&nbsp;&nbsp;Actualizar&nbsp;&nbsp;" class=label></TD>
			</TR>
			</TABLE>
		<TD>
		</TR>
	</FORM>
	</TABLE>
</BODY>
</HTML>