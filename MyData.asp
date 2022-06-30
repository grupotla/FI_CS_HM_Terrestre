<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"
Dim Conn, Action, rs, Active, Checked, JavaMsg, CreatedDate, CreatedTime
Dim OperatorID, Pwd, FirstName, LastName, Email, Phone, Position, OperatorLevel
Dim StartTime, FinishTime, ProviderName, Sign

	OperatorID = CheckNum(Session("OperatorID"))
	FirstName = Request.Form("FirstName")
	LastName = Request.Form("LastName")
	Email = Request.Form("Email")
	Phone = Request.Form("Phone")
	Action = CheckNum(Request.Form("Action"))
	Sign = Request.Form("Sign")
	ProviderName = ""
	StartTime = 0
	FinishTime = 0

    'Seleccionando todas las categorias que existen actualmente
	OpenConn Conn
    if Action = 2 then
		'obteniendo los parametros para hacer las operaciones de Insert, Update o Delete
		FormatTime CreatedDate, CreatedTime				 
		JavaMsg = ""
		OpenTable Conn, "Operators", rs
			rs.Filter = "OperatorID=" & OperatorID						 
			if Not rs.EOF Then 'Si existe el atributo, puede actualizarlo
				'Guardando el nombre de la nueva columna para futuras verificaciones
				SaveData rs, Action, Array("FirstName", FirstName, "LastName", LastName, "Email", Email, "Phone", Phone, "Position", Position, "OperatorLevel", OperatorLevel, "CreatedDate", CreatedDate, "CreatedTime", CreatedTime, "Sign", Sign)
			else
				JavaMsg = "El Operador no existe"	
			end if
		CloseOBJ rs
	end if

	Set rs = Conn.Execute("select OperatorID, FirstName, LastName, Email, Phone, StartTime, FinishTime, Sign from Operators where OperatorID=" & OperatorID & " and Login='" & Session("Login") & "'")
	If Not rs.EOF Then
		 OperatorID = rs(0)
		 FirstName = rs(1)
		 LastName = rs(2)
		 Email = rs(3)
		 Phone = rs(4)
		 StartTime = rs(5)
		 FinishTime = rs(6)
		 Sign = rs(7)
    End If
    closeOBJs rs, Conn
	Session("Sign") = Sign
	Session("OperatorEmail") = Email
	Session("OperatorName") = FirstName & " " & LastName
%>

<HTML><HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language=javascript src="img/config.js"></SCRIPT>
<SCRIPT src="img/mainLib.js" language="JavaScript"></SCRIPT>
<SCRIPT language="javascript" src="img/validaciones.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
	 function validar(Action) {
			 if (document.forma.FirstName.value == "") {
			 		alert("Debe Ingresar Nombre del Editor");
			 		document.forma.FirstName.focus();
			 		return false;			 		
			 }
			 if (document.forma.Email.value == "") {
			 		alert("Debe Ingresar Email del Editor");
			 		document.forma.Email.focus();
			 		return false;			 		
			 }

			 document.forma.Action.value = Action;
			 document.forma.submit();
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
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
	<%if JavaMsg <> "" then
			 Response.Write "<SCRIPT>alert('" & JavaMsg & "');</SCRIPT>"
		end if
	%> 
	<FORM name="forma" action="MyData.asp" method="post" target=_self>
	<INPUT name="Action" type=hidden value=0>
	<INPUT name="OID" type=hidden value=<%=OperatorID%>>

	<TABLE cellspacing=0 cellpadding=2 width=400 align=center>
		<TR>
		<TD class=label align=right width=20%><b>Login:</b></TD>
		<TD class=label align=left width=75%><b><%=Session("Login")%></b></TD>
	  </TR>
		<%if Session("CRM") and (Session("OperatorLevel")=3 or Session("OperatorLevel")=4) then%>
		<TR>
		<TD class=label align=right width=20%><b>Proveedor:</b></TD>
		<TD class=label align=left width=75%><%=ProviderName%></TD>
	  </TR>
		<TR>
		<TD class=label align=right width=20%><b>Horario:</b></TD>
		<TD class=label align=left width=75%>
		<%	OperatorSchedule "", "", StartTime, 2
				Response.write "&nbsp;-&nbsp;"
				OperatorSchedule "", "", FinishTime, 2
		%>
		</TD>
	  </TR>
		<%end if%>
		<TR>
		<TD class=label align=right width=20%><b>Nombre:</b></TD>
		<TD class=label align=left width=75%><INPUT name="FirstName" type=text value="<%=FirstName%>" size=40 maxLength=255 class=label></TD>
	  </TR>
		<TR>
		<TD class=label align=right width=20%><b>Apellido:</b></TD>
		<TD class=label align=left width=75%><INPUT name="LastName" type=text value="<%=LastName%>" size=40 maxLength=255 class=label></TD>
	  </TR>
		<TR>
		<TD class=label align=right width=20%><b>Email:</b></TD>
		<TD class=label align=left width=75%><INPUT name="Email" type=text value="<%=Email%>" size=40 maxLength=255 class=label></TD>
	  </TR>
		<TR>
		<TD class=label align=right width=20%><b>Teléfono:</b></TD>
		<TD class=label align=left width=75%><INPUT name="Phone" type=text value="<%=Phone%>" size=30 maxLength=255 class=label></TD>
	  </TR>
	  <TR>
		<TD class=label align=right width=42%><b>Firma:</b></TD>
		<TD class=label align=left width=58%><Textarea name="Sign" id="Firma" cols="30" rows="5"><%=Sign%></Textarea></TD>
	  </TR> 
	  
		</TABLE>
		<TABLE cellspacing=0 cellpadding=2 width=200 align=center>
		<TR>
							<TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(2)" value="&nbsp;&nbsp;Actualizar&nbsp;&nbsp;" class=label></TD>
		</TR>
		</TABLE>
	</FORM>
</BODY>
<script language="javascript1.2">
editor_generate('Sign');
</SCRIPT>
</HTML>
<%
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>
