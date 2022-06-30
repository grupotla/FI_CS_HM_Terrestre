<%@ Language=VBScript %>
<!-- #INCLUDE file="utils.asp" -->
<%
	CountTableValues = -1
	OpenConn Conn1
	Set rs = Conn1.Execute("select BLDetailID, CommoditiesID, DiceContener from BLDetail")
	
	If Not rs.EOF Then
    	aTableValues = rs.GetRows
    	CountTableValues = rs.RecordCount-1
	End If
	CloseOBJ rs

	response.write "<table border=1>"
	OpenConn2 Conn
	for i=0 to CountTableValues
		'Names = Split(aTableValues(2,i),chr(13))
		Names = Replace(aTableValues(2,i),chr(13),"<br>",1,-1)
		'BLS-Shippers
		'Set rs = Conn.Execute("select BLID from BLs where ShipperID=" & aTableValues(0,i))
		'BLDetail-Shippers
		'Set rs = Conn.Execute("select BLDetailID from BLDetail where AgentsID=" & aTableValues(0,i))
		'BLS-Consigners
		'Set rs = Conn.Execute("select BLID from BLs where ConsignerID=" & aTableValues(0,i))
		'BLDetail-Shippers
		Set rs = Conn.Execute("select commodityid, namees from commodities where namees='" & Ucase(aTableValues(2,i)) & "'")
		If Not rs.EOF Then
    		'Conn1.Execute("update BLDetail set CommoditiesID=" & rs(0) & " where BLDetailID=" & aTableValues(0,i))
			response.write "update BLDetail set CommoditiesID=" & rs(0) & " where BLDetailID=" & aTableValues(0,i)
			response.write "<tr><td>" & aTableValues(0,i) & "</td><td>" & aTableValues(1,i) & "</td><td>" & Names & "</td><td>" & rs(0) & "</td><td>" & rs(1) & "</td></tr>"
		Else
    		'response.write "<tr><td>" & aTableValues(0,i) & "</td><td>" & aTableValues(1,i) & "</td><td>" & aTableValues(2,i) & "</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
    		'response.write "<tr><td>" & aTableValues(0,i) & "</td><td>" & aTableValues(1,i) & "</td><td>" & Names(0) & "</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
    		response.write "<tr><td>" & aTableValues(0,i) & "</td><td>" & aTableValues(1,i) & "</td><td>" & Names & "</td><td>&nbsp;</td><td>&nbsp;</td></tr>"
		End If
		CloseOBJ rs
	next
	CloseOBJ Conn
	CloseOBJ Conn1
	Set aTableValues=Nothing
	response.write "</table>"
%>
listo