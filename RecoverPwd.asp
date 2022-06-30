<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="../admin/utils.asp" -->
<%
Dim MS, MSG
MS = CheckNum(Request.QueryString("MS"))
if MS = 1 then MSG = "El usuario es incorrecto" end if
%>
<HTML>
<HEAD>
<TITLE>Terra Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
</HEAD>
<LINK REL="stylesheet" type="text/css" HREF="/admin/img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0 class=menu>
    <TR>
    <TD class=menu vAlign=center align=right><font class=activeMain><!--Web Presence Tools 2.0&nbsp;&nbsp;--></font></TD>
    <TD vAlign=center align=middle width=79><IMG src="/admin/img/logo.gif">
		</TD>
		</TR>
		<TR>
    <TD colspan=2 class=submenu vAlign=center align=right><font class=activeMain>&nbsp;</font></TD>
		</TD>
		</TR>
</TABLE>
<br>
<br>
<form name=forma method=post action=RecoverPwd1.asp>
<table cellSpacing=0 cellPadding=0 width="300" align=center>
   <tr>
   	<td align=right>
				<font class=label>Para recordar su contraseña ingrese su Login:</font>
		</td>
	</tr>
	<tr>	
		<td align=center>		
				<input style="font-family: Verdana; font-size: 10px" type=text name=Login>
		</td>
	</tr>
	<tr>
    <td colspan=2 align=center>
				<input style="font-family: Verdana; font-size: 10px" type=submit value=ingresar>
		</td>
	</tr>
	<% if MSG <> "" then Response.Write "<tr><td colspan=2 align=center><br><font style='font-family: Verdana; font-size: 10px; color:red;'>" & MSG & "</font></td></tr>" end if%>
</table>
</form>

</BODY>
</HTML>

