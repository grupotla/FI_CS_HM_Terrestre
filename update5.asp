<%@ Language=VBScript %>
<!-- #INCLUDE file="utils.asp" -->
<%
	CountTableValues = -1
	OpenConn Conn
	Set rs = Conn.Execute("select AgentsID, Agents, BLDetailID from BLDetail where BLDetailID in (2089, 2090, 2091, 2092, 2093, 2094, 2105) order by Agents")
	
	If Not rs.EOF Then
    	aTableValues = rs.GetRows
    	CountTableValues = rs.RecordCount-1
	End If
	CloseOBJs rs, Conn

	response.write "<table border=1>"
	OpenConn2 Conn
	for i=0 to CountTableValues
		'ANames = Split(aTableValues(1,i),chr(13))
		'Names = Replace(aTableValues(2,i),chr(13)&chr(10),"<br>",1,-1)
		'Names = Left(ANames(0),8)
		Names = aTableValues(1,i)
		Set rs = Conn.Execute("select a.id_cliente, a.nombre_cliente, b.id_direccion from clientes a, direcciones b where a.id_cliente = b.id_cliente and a.nombre_cliente ilike '%" & UCase(Names) & "%'")

		If Not rs.EOF Then
    		'response.write "<tr><td>" & aTableValues(0,i) & "</td><td>" & aTableValues(1,i) & "</td><td>" & rs(0) & "</td><td>" & rs(1) & "</td><td>" & rs(2) & "</td><td>" & rs(3) & "</td><td>" & rs(4) & "</td><td>" & rs(5) & "</td><td>" & rs(6) & "</td></tr>"
			response.write "<tr><td>" & aTableValues(2,i) & "</td><td>" & aTableValues(0,i) & "</td><td>" & aTableValues(1,i) & "</td><td>" & rs(0) & "</td><td>" & rs(1) & "</td><td>" & rs(2) & "</td></tr>"
			'Conn.Execute("insert into agentes (hora_creacion, fecha_creacion, agente, activo, direccion, telefono) values ('"&rs(1)&"', "&rs(2)&",'"&rs(3)&"',1,'"&rs(5)&"','"&rs(6)&"')")
			'response.Write("insert into agentes (fecha_creacion, hora_creacion, agente, activo, direccion, telefono) values ('"&rs(1)&"', "&CheckNum(rs(2))&",'"&rs(3)&"----- -',true,'"&rs(5)&"','"&rs(6)&"');<br>")
		Else
    		'response.write "<tr><td>" & aTableValues(0,i) & "</td><td>" & aTableValues(1,i) & "</td><td>" & aTableValues(2,i) & "</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
    		'response.write "<tr><td>" & aTableValues(0,i) & "</td><td>" & aTableValues(1,i) & "</td><td>" & Names(0) & "</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
    		response.write "<tr><td>" & aTableValues(2,i) & "</td><td>" & aTableValues(0,i) & "</td><td>" & aTableValues(1,i) & "</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
		End If
		CloseOBJ rs
	next
	CloseOBJ Conn
	Set aTableValues=Nothing
	response.write "</table>"
%>
listo