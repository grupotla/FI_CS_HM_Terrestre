<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1"
Dim Conn, rs, Operators, CantOperators, HTMLCode, i, j, MatchCountries, Match, SQLQuery, Symbol
 
	CantOperators = -1
	Set MatchCountries = FRegExp(PtrnCountries, Session("Countries"),  "", 1)
	For Each Match in MatchCountries
		if i = 0 then
				SQLQuery = " a.Countries like ""%" & Match.value & "%"" "
		else
				SQLQuery = SQLQuery & "or a.Countries like ""%" & Match.value & "%"" "
		end if
		i=1 
	Next
	SQLQuery = "(" & SQLQuery & ")"
	
	if Session("OperatorLevel") = 0 then
		Symbol = " >= "' Solo para Root, puede crear todo tipo de usuarios (Root, Admin u Operador)
	else
		Symbol = " > "' Para ADmin, solo puede crear Operadores
	end if
		
	OpenConn Conn
	Set rs = Conn.Execute("select a.OperatorID, a.Login, a.LastName, a.FirstName, a.OperatorLevel, a.Countries, a.Active from Operators a where a.OperatorLevel" & Symbol & Session("OperatorLevel") & " and " & SQLQuery & " order by a.OperatorLevel, a.Login, a.LastName, a.FirstName")
	If Not rs.EOF Then
        Operators = rs.GetRows
		CantOperators = UBound(Operators,2)
    End If
    closeOBJs rs, Conn

	  if CantOperators >= 0 then
		 for i=0 to CantOperators
		 		 HTMLCode = HTMLCode & "<tr>" 
				 for j = 0 to 6 
				 		 HTMLCode = HTMLCode & "<td class=list><a class=labellist href=Operators.asp?OID=" & CheckNum(Operators(0,i)) & ">"
						 select case j
						 case 6
						 	if Operators(j,i) = 1 then
						 		HTMLCode = HTMLCode & "Activo</a></td>"
							else
								HTMLCode = HTMLCode & "Inactivo</a></td>"
							end if
						 Case 4
						 	Select Case Operators(j,i)
							Case 0
						 		HTMLCode = HTMLCode & "Root</a></td>"
							Case 1
								HTMLCode = HTMLCode & "Admin</a></td>"
							Case Else
								HTMLCode = HTMLCode & "Operator</a></td>"
							end Select
						 Case Else
						 	if (CheckNum(Operators(j,i)) > 0) Then
                                HTMLCode = HTMLCode & CheckNum(Operators(j,i)) & "</a></td>"
                            Else
                                HTMLCode = HTMLCode & Operators(j,i) & "</a></td>"
                            End If
						 end Select						 
				 next
				 HTMLCode = HTMLCode & "</tr>" 
		 next
	  end if
	
%>


<HTML><HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language=javascript src="img/config.js"></SCRIPT>
<SCRIPT src="img/mainLib.js" language="JavaScript"></SCRIPT>
<SCRIPT language="javascript" src="img/validaciones.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
	<TABLE cellspacing=0 cellpadding=2 width=100% align=center>
	<TR>
	<TD width=40% colspan=2 class=label align=right valign=top>
		<TABLE cellspacing=5 cellpadding=2 width=100% align=center>
		<tr>
		<td class=titlelist><b>Código</b></td><td class=titlelist><b>Login</b></td><td class=titlelist><b>Apellido</b></td><td class=titlelist><b>Nombre</b></td><td class=titlelist><b>Tipo</b></td><td class=titlelist><b>Paises</b></td><td class=titlelist><b>Status</b></td>
		</tr>
		<%=HTMLCode%>
		</TABLE>
	</TD>
	</TR>
	</TABLE>		
</BODY>
</HTML>
<%
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>

