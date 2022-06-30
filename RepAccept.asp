<%
Checking "0|1|2"
	Dim BLNumbers, LenBLNumbers
	if CountTableValues >= 0 then
		CreatedDate = ConvertDate(aTableValues(1, 0),2)
		CreatedTime = aTableValues(2, 0)
		Expired = aTableValues(3, 0)
		LtAcceptDate = aTableValues(4, 0)
		BrokerRecepID = aTableValues(5, 0)
		CountryDes = aTableValues(6, 0)
		BLNumber = aTableValues(7, 0)	
	end if
	Set aTableValues = Nothing

	OpenConn Conn
	'Actualizando las CP que forman un Grupo
	if BLType < 0 and Action=2 then
		Conn.Execute("update BLs a, BLGroupDetail b set a.LtAcceptDate='" & LtAcceptDate & "', a.BrokerRecepID=" & BrokerRecepID & " where a.BLID=b.BLID and b.BLGroupID=" & ObjectID)
	end if
	
	'Obteniendo listado de Aduanas
	Set rs = Conn.Execute("select BrokerID, Name, Countries from Brokers where Expired=0 order by Countries, Name")
	If Not rs.EOF Then
   		aList1Values = rs.GetRows
       	CountList1Values = rs.RecordCount-1
    End If
		
	CloseOBJs rs, Conn
%>
<HTML>
<HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
	function validar(Action) {
		if (!valTxt(document.forma.LtAcceptDate, 3, 5)){return (false)};
		if (!valSelec(document.forma.BrokerRecepID)){return (false)};
	    document.forma.Action.value = Action;
		document.forma.submit();			 
	 }
	 
	 function SetLabelID(Label) {
		var LabelID = "";
		if (Label == "LtAcceptDate") {
			LabelID = "Fecha Carta de Aceptación";
		} 
		return LabelID;
	}
	 
	 
	function abrir(Label){
	var DateSend, Subject;
		if (parseInt(navigator.appVersion) < 5) {
			DateSend = document.forma(Label).value;
		} else {
			var LabelID = SetLabelID(Label);
			DateSend = document.getElementById(LabelID).value;
		}
		Subject = '';	
		window.open('Agenda.asp?Action=1&Label=' + Label + '&DateSend=' + DateSend + '&Subj=' + Subject,'Seleccionar','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,width=170,height=160,top=250,left=250');
	}
	
_editor_url = "Javascripts/";                     // URL to htmlarea files
var win_ie_ver = parseFloat(navigator.appVersion.split("MSIE")[1]);
if (navigator.userAgent.indexOf('Mac')        >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Windows CE') >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Opera')      >= 0) { win_ie_ver = 0; }
if (win_ie_ver >= 5.5) {
  document.write('<scr' + 'ipt src="' +_editor_url + 'editor.js"');
  document.write(' language="Javascript1.2"></scr' + 'ipt>');  
} else { document.write('<scr'+'ipt>function editor_generate() { return false; }</scr'+'ipt>'); }
</script>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<style type="text/css">
.style10 {font-size: 10px; font-family: Verdana, Arial, Helvetica, sans-serif; }
</style>
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
	<TABLE cellspacing=0 cellpadding=2 width=400 align=center>
	<FORM name="forma" action="InsertData.asp" method="post">
	<INPUT name="Action" type=hidden value=0>
	<INPUT name="Expired" type=hidden value="on">
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
	<INPUT name="OID" type=hidden value="<%=ObjectID%>">
	<INPUT name="AT" type=hidden value="<%=BLType%>">
	<INPUT name="CD" type=hidden value="<%=CreatedDate%>">
	<INPUT name="CT" type=hidden value="<%=CreatedTime%>">
	<INPUT name="CountryDes" type=hidden value="<%=CountryDes%>">
		<TR><TD class=label align=right><b>No. Carta Porte:</b></TD><TD class=label align=left><%=BLNumber%></TD></TR> 
		<TR><TD class=label align=right><b>Fecha Creación:</b></TD><TD class=label align=left><%=CreatedDate%></TD></TR> 
		<TR><TD class=label align=right><b>Código:</b></TD><TD class=label align=left><%if ObjectID <> 0 then response.write ObjectID End if%></TD></TR> 
		<!--<TR><TD class=label align=right><b>No.Carta de Aceptación:</b></TD><TD class=label align=left><%'=LtAcceptNumber%></TD></TR>-->
		<TR><TD class=label align=right><b>Fecha Carta de Aceptación:</b></TD><TD class=label align=left>
		<INPUT readonly="readonly" name="LtAcceptDate" id="Fecha Carta de Aceptación" type=text value="<%=LtAcceptDate%>" size=23 maxLength=19 class=label>&nbsp;
			<INPUT type=button value="Seleccionar" onClick="JavaScript:abrir('LtAcceptDate');" class=label></TD></TR>
		<TR><TD class=label align=right><b>Aduana de Recepción:</b></TD><TD class=label align=left>
		<select class="style10" name="BrokerRecepID" id="Aduana de Recepción">
		<option value="-1">Seleccionar</option>
		<%
			For i = 0 To CountList1Values
		%>
		<option value="<%=aList1Values(0,i)%>"><%response.write aList1Values(2,i) & " - " & aList1Values(1,i)%></option>
		<%
			Next
		%>
		</select>
		</TD></TR>	
	<TR><TD colspan="2" class=label align=center>
			<TABLE cellspacing=0 cellpadding=2 width=200>
		    <TR>
			<%if LtAcceptDate <> "" then%>
				<TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="Javascript:window.open('Reports.asp?GID=<%=GroupID%>&OID=<%=ObjectID%>&AT=<%=BLType%>','RepAccept','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750,height=600');return false;" value="&nbsp;&nbsp;Previsualizar&nbsp;Carta&nbsp;de&nbsp;Aceptación&nbsp;&nbsp;" class=label></TD>
			<%end if%>
				<TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(2)" value="&nbsp;&nbsp;Actualizar&nbsp;&nbsp;" class=label></TD>
			</TR>
			</TABLE>
		<TD>
		</TR>
	</FORM>
	</TABLE>
</BODY>
<script language="javascript1.2">
selecciona('forma.BrokerRecepID','<%=BrokerRecepID%>');
</SCRIPT>

</HTML>