<%@ Language=VBScript %>
<%Option Explicit%>
<%
Dim MS, MSG
MS = Trim(Request.QueryString("MS"))
MSG = ""
	 Select Case MS
	 				Case "1"
							 MSG = "No tiene permisos de Acceso"
					Case "2" 
							 MSG = "usuario o Pwd Incorrecto"
					Case "3"
							 MSG = "Ingresar Usr o Pwd"
					Case "4"
							 MSG = "La Sesión ha expirado"
                    Case "5"
							 MSG = "La contraseña ha vencido, favor de actualizarla <a href= http://10.10.1.20/catalogo_admin/cambio/index.php >AQUI</a>"
   End Select
%>
<HTML>
<HEAD>
<TITLE>Sistema Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<meta http-equiv="X-UA-Compatible" content="IE=9" />
</HEAD>
<LINK REL="stylesheet" type="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="javascript:document.forma.login.focus();">
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0 class=menu>
    <TR>
    <TD class=menu vAlign=center align=right><font class=activeMain>Sistema  Terrestre</font></TD>
    <TD vAlign=center align=middle width=1><IMG src="img/spacer.gif" height="31" border="0">
		</TD>
		</TR>
		<TR>
    <TD colspan=2 class=submenu vAlign=center align=right><font class=activeMain>&nbsp;</font></TD>
		</TD>
		</TR>
</TABLE>
<br>
<br>
<form name=forma method=post action=validator.asp>
<table cellSpacing=0 cellPadding=0 width="300" border=0 align=center>
  <tr>
   	<td align=right>
				<font class=label>Login: </font>
		</td>
		<td>		
				<input style="font-family: Verdana; font-size: 10px" type=text name=login  value="">
		</td>
	</tr>
	<tr>	
		<td align=right>		
				<font class=label>Password: </font><br>
		</td>
		<td>		
				<input style="font-family: Verdana; font-size: 10px" type=password name=pwd value="">
		</td>
	</tr>
	<tr>
    <td colspan=2 align=center>
				<input style="font-family: Verdana; font-size: 10px" type=submit value=ingresar>
		</td>
	</tr>
	<% if MSG <> "" then Response.Write "<tr><td colspan=2 align=center><br><font style='font-family: Verdana; font-size: 10px; color:red;'>" & MSG & "</font></td></tr>" end if%>
  <tr>
    <td colspan=2 align=center>
				<br><br><br><br><br><font class=label>Funciona en MSIE / Windows OS</font><br><br>
				<A href="http://www.microsoft.com/windows/ie/default.asp" target="_blank"><img src="img/ie.gif" border=0></a>
		</td>
	</tr>
</table>
</form>
</BODY>
</HTML>

