<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1"
Dim Conn, Action, rs, Active, Checked, JavaMsg, CreatedDate, CreatedTime, Sign
Dim OperatorID, Login, FirstName, LastName, Email, Phone, Position, OperatorLevel, Countries
Dim StartTime, FinishTime, ValidCountries, aList1Values, CountList1Values, i

		if Session("OperatorLevel") = 0 then
			ValidCountries = PtrnCountries
		else
			ValidCountries = replace(replace(Session("Countries"),"(","",1,-1),")","",1,-1)
		end if
        
		Checking "0|1"

	  	OperatorID = CheckNum(Request("OID"))
		Login = PurgeData(Request.Form("Login"))
		FirstName = PurgeData(Request.Form("FirstName"))
		LastName = PurgeData(Request.Form("LastName"))
		Email = PurgeData(Request.Form("Email"))
		Phone = PurgeData(Request.Form("Phone"))
		Position = PurgeData(Request.Form("Position"))
		OperatorLevel = CheckNum(Request.Form("OL"))
	 	Active = Request.Form("Active")
		Action = CheckNum(Request.Form("Action"))
		StartTime = CheckNum(Request.Form("StartTime"))
		FinishTime = CheckNum(Request.Form("FinishTime"))
		Countries = Request.Form("Countries")
		Sign = Request.Form("Sign")
		Checked = ""
		If Active = "on" Then
    	 	Active = 1
		Else
			Active = 0
		End If

	OpenConn Conn
		if (Action = 1) or (Action = 2) or (Action = 3) then
				 'obteniendo los parametros para hacer las operaciones de Insert, Update o Delete
				 FormatTime CreatedDate, CreatedTime				 
				 JavaMsg = ""
				 OpenTable Conn, "Operators", rs
						select case Action
						case 1 ' Insert
							 rs.Filter = "Login='" & Login & "'" 
							 if rs.EOF Then 'Si no existe el lenguaje, puede ingresarlo
									'Guardando el nombre de la nueva columna para futuras verificaciones
									SaveData rs, Action, Array("Login", Login, "OperatorID", OperatorID, "FirstName", FirstName, "LastName", LastName, "Email", Email, "Phone", Phone, "Position", Position, "OperatorLevel", OperatorLevel, "CreatedDate", CreatedDate, "CreatedTime", CreatedTime, "Active", Active, "StartTime", StartTime, "FinishTime", FinishTime, "Countries", Countries, "Sign", Sign)
									CloseOBJ rs
									Set rs = Conn.Execute("select OperatorID from Operators where Login='" & Login & "' and OperatorLevel>=" & Session("OperatorLevel"))
									OperatorID = rs(0)
							 else
							 		 JavaMsg = "El Usuario ya existe"
							 end if
						 case 2 'Update
						 	 rs.Filter = "OperatorID=" & OperatorID	
						 	 if Not rs.EOF Then 'Si existe el atributo, puede actualizarlo
									'Guardando el nombre de la nueva columna para futuras verificaciones
									SaveData rs, Action, Array("Login", Login, "OperatorID", OperatorID, "FirstName", FirstName, "LastName", LastName, "Email", Email, "Phone", Phone, "Position", Position, "OperatorLevel", OperatorLevel, "CreatedDate", CreatedDate, "CreatedTime", CreatedTime, "Active", Active, "StartTime", StartTime, "FinishTime", FinishTime, "Countries", Countries, "Sign", Sign)
							 else
							 		 JavaMsg = "El Usuario no existe"	
							 end if
						 case 3 'Delete
						   rs.Filter = "OperatorID=" & OperatorID
							 if Not rs.EOF Then 'Si existe el atributo, puede borrarlo
							 		 'Eliminando el nombre de la columna en la tabla de verificacion (Attributes)
									 rs.Delete
							 else
							 		 JavaMsg = "El Usuario no existe"
							 end if
						end select
						CloseOBJ rs
		end if
		
	'Obteniendo los datos personales del Editor
	Set rs = Conn.Execute("select OperatorID, Login, FirstName, LastName, Email, Phone, Position, OperatorLevel, Active, StartTime, FinishTime, Countries, Sign from Operators where OperatorID=" & OperatorID & " and OperatorLevel>=" & Session("OperatorLevel"))
    If Not rs.EOF Then
			 OperatorID = rs(0)
			 Login = rs(1)
			 FirstName = rs(2)
			 LastName = rs(3)
			 Email = rs(4)
			 Phone = rs(5)
			 Position = rs(6)
			 OperatorLevel = rs(7)
	 		 Active = rs(8)
			 StartTime = rs(9)
			 FinishTime = rs(10)
			 Countries = rs(11)
			 Sign = rs(12)
		else
			 OperatorID = 0
			 Login = ""
			 FirstName = ""
			 LastName = ""
			 Email = ""
			 Phone = ""
			 Position = ""
			 OperatorLevel = -1
	 		 Active = 0
			 Countries = ""
			 Sign = ""
    End If
    closeOBJs rs, Conn
	
	OpenConn2 Conn
	Set rs = Conn.Execute("select id_usuario, pw_name, dominio from usuarios_empresas order by pw_name, dominio")
	If Not rs.EOF Then
   		aList1Values = rs.GetRows
       	CountList1Values = rs.RecordCount-1
    End If
    closeOBJs rs, Conn	

	If Active = 1 Then
   	 	Checked = "checked"
	End If
%>

<HTML><HEAD><TITLE>Site - Aimar - Administración</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
	function validar(Action) {		
		if (Action == 1 || Action == 2) {
			//JoinCountries (document.forma, "'GT','SV','HN','NI','CR','PA','BZ','GT2','SV2','NI2','CR2','PA2','BZ2'");
			JoinCountries (document.forma, "<%=ValidCountries%>");
			if (!valSelec(document.forma.OID)){return (false)};
			if (!valTxt(document.forma.FirstName, 2, 5)){return (false)};
			if (!valTxt(document.forma.LastName, 3, 5)){return (false)};
			if (!valRept(document.forma.FirstName, document.forma.LastName)){return (false)};
			if (!valEmail(document.forma.Email)){return (false)};
			document.forma.Login.value = Logins[document.forma.OID.value];
		};
		document.forma.Action.value = Action;
		//alert(document.forma.Countries.value);
        document.forma.submit();			 
	 }
</SCRIPT>
<script language="Javascript1.2"><!-- // load htmlarea
_editor_url = "Javascripts/";                     // URL to htmlarea files
var win_ie_ver = parseFloat(navigator.appVersion.split("MSIE")[1]);
if (navigator.userAgent.indexOf('Mac')        >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Windows CE') >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Opera')      >= 0) { win_ie_ver = 0; }
if (win_ie_ver >= 5.5) {
  document.write('<scr' + 'ipt src="' +_editor_url + 'editor.js"');
  document.write(' language="Javascript1.2"></scr' + 'ipt>');  
} else { document.write('<scr'+'ipt>function editor_generate() { return false; }</scr'+'ipt>'); }
// -->
var Logins = new Array();
<%for i=0 to CountList1Values%>
	Logins[<%=aList1Values(0,i)%>] = "<%=aList1Values(1,i)%>"
<%next%>
</script>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
	<%if JavaMsg <> "" then
			 Response.Write "<SCRIPT>alert('" & JavaMsg & "');</SCRIPT>"
		end if
	%> 
	<FORM name="forma" action="Operators.asp" method="post" target=_self>
	<INPUT name="Action" type=hidden value=0>
	<INPUT name="Login" type=hidden value=<%=Login%>>

	<TABLE cellspacing=0 cellpadding=2 width=400 align=center>
		<TR>
		<TD class=label align=right width=20%><b>Activo:</b></TD>
		<TD class=label align=left width=75%><INPUT name="Active" maxlength=250 size=30 maxlength=200 TYPE=checkbox class=label <%=Checked%>>
		</TD>
	  </TR>
		<TR>
		<TD class=label align=right width=20%><b>Tipo:</b></TD>
		<TD class=label align=left width=75%>
		<select name="OL" class=label>
		<%Select Case session("OperatorLevel")%>
		<%case 0%>
					  <option value="0"<% If OperatorLevel = 0 then response.write " selected"%>>Root</option>
					  <option value="1"<% If OperatorLevel = 1 then response.write " selected"%>>Admin</option>
					  <option value="2"<% If OperatorLevel = 2 then response.write " selected"%>>Editor</option>
		<%case 1%>
					  <option value="2"<% If OperatorLevel = 2 then response.write " selected"%>>Editor</option>
		<%end select%>
		</select>
		</TD>
	  </TR>
		</TABLE>
		<TABLE cellspacing=0 cellpadding=2 width=400 align=center>
		<TR>
		<TD class=label align=right width=42%></TD>
		<TD class=label align=left width=58%><br>
		  <b><u>Datos Personales:</u></b></TD>
	    </TR>
		<TR>
		<TD class=label align=right width=42%><b>Login:</b></TD>
		<TD class=label align=left width=58%>
		<select name="OID" class="label" id="Login ">
			<option value="-1">Seleccionar</option>
			<%for i=0 to CountList1Values%>
				<option value="<%=aList1Values(0,i)%>"><%=aList1Values(1,i) & "@" & aList1Values(2,i)%></option>
			<%next%>
		</select>
		</TD>
	    </TR>
		<TR>
		<TD class=label align=right width=42%><b>Nombre:</b></TD>
		<TD class=label align=left width=58%><INPUT name="FirstName" id="Nombre" type=text value="<%=FirstName%>" size=40 maxLength=255 class=label></TD>
	  </TR>
		<TR>
		<TD class=label align=right width=42%><b>Apellido:</b></TD>
		<TD class=label align=left width=58%><INPUT name="LastName" id="Apellido" type=text value="<%=LastName%>" size=40 maxLength=255 class=label></TD>
	  </TR>
		<TR>
		<TD class=label align=right width=42%><b>Email:</b></TD>
		<TD class=label align=left width=58%><INPUT name="Email" id="Email" type=text value="<%=Email%>" size=40 maxLength=255 class=label></TD>
	  </TR>
		<TR>
		<TD class=label align=right width=42%><b>Teléfono:</b></TD>
		<TD class=label align=left width=58%><INPUT name="Phone" id="Telefono" type=text value="<%=Phone%>" size=30 maxLength=255 class=label></TD>
	  </TR>
		<TR>
		<TD class=label align=right width=42%><b>Puesto:</b></TD>
		<TD class=label align=left width=58%><INPUT name="Position" id="Puesto" type=text value="<%=Position%>" size=30 maxLength=255 class=label>
		<INPUT name="Countries" type=hidden value="<%=Countries%>" size=30 maxLength=255 class=label>
		</TD>
	  </TR>
		<TR>
		<TD class=label align=right width=42%><b>Firma:</b></TD>
		<TD class=label align=left width=58%><Textarea name="Sign" id="Firma" cols="30" rows="5"><%=Sign%></Textarea></TD>
	  </TR>
	  <TR><TD colspan="2" align="left"> 
	  <% ListCountries Countries, ValidCountries%>
	  </TD></TR>
		</TABLE>
		<TABLE cellspacing=0 cellpadding=2 width=200 align=center>
		<TR>
			<%if CheckNum(OperatorID)=0 then%>
			<TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(1)" value="&nbsp;&nbsp;Agregar&nbsp;&nbsp;" class=label></TD>
			<%else%>
			<TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(2)" value="&nbsp;&nbsp;Actualizar&nbsp;&nbsp;" class=label></TD>
			<%end if%>
			<TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(3)" value="&nbsp;&nbsp;Eliminar&nbsp;&nbsp;" class=label></TD>
		</TR>
		</TABLE>
	</FORM>
</BODY>
<script language="javascript1.2">
editor_generate('Sign');
selecciona('forma.OID','<%=OperatorID%>');
</SCRIPT>
</HTML>
<%
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>