<%@ Language=VBScript %>
<!-- #INCLUDE file="utils.asp" -->
<%
	CountTableValues = -1
	OpenConn Conn
	Set rs = Conn.Execute("select BLID, BLNumber from BLs order by BLID desc")
	If Not rs.EOF Then
    	aTableValues = rs.GetRows
    	CountTableValues = rs.RecordCount-1
	End If
	CloseOBJ rs
	
	for i=0 to CountTableValues
		Set rs = Conn.Execute("select AgentsID, ClientsID, BLDetailID from BLDetail where BLID=" & aTableValues(0,i))
		Do While Not rs.EOF
			Conn.Execute("update BLDetail set HBLNumber='" & aTableValues(1,i) & "-" & FiveDigits(CheckNum(rs(0))+CheckNum(rs(1))) & "' where BLDetailID=" & rs(2))
			response.write "update BLDetail set HBLNumber='" & aTableValues(1,i) & "-" & FiveDigits(CheckNum(rs(0))+CheckNum(rs(1))) & "' where BLDetailID=" & rs(2) & "<br>"
			rs.MoveNext
		loop
		CloseOBJ rs
	next
	CloseOBJ Conn
	Set aTableValues=Nothing
%>
listo