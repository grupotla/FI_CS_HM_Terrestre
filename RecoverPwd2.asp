<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="../admin/utils.asp" -->
<%
Dim MS, MSG, Login, Conn, rs, Answer, UserID, Pwd
Login = CheckTxt(Request("Login"))
UserID = CheckNum(Request("UID"))
Answer = PurgeData(Request("Answer"))
if Login <> "" then
	 		OpenConn Conn
   		Set rs = Conn.Execute("select Pwd from Users where User='" & Login & "' and UserID=" & UserID & " and RememberAnswer='" & Answer & "'")
	 		if Not rs.EOF then
	 			 Pwd = rs(0)
	 		end if
			CloseOBJs rs, Conn
end if
if Pwd <> "" then
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
<form name=forma method=post action=RecoverPwd2.asp>
<table cellSpacing=0 cellPadding=0 width="300" align=center>
<input type=hidden name=UID value="<%=UserID%>">
<input type=hidden name=Login value="<%=Login%>">
   <tr>
   	<td align=center>
				<font class=label>La contraseña del Usuario <b><%=Login%></b> es:</font>
		</td>
	</tr>
   <tr>
   	<td align=center>
				<select name=Question class=label>
				<option>Ver contraseña</option>
				<option><%=Pwd%></option>
				</select>				
		</td>
	</tr>
</table>
</form>

</BODY>
</HTML>
<%
else
		Response.Redirect ("RecoverPwd1.asp?MS=1&Login=" & Login)
end if
%>
