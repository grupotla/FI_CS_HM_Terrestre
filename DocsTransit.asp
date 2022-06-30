<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
Checking "0|1"
Dim ObjectID, Conn, rs, aListValues, CountListValues, i, FilesID

	ObjectID = CheckNum(Request("OID"))
	CountListValues = -1
	OpenConn Conn	
	Set rs = Conn.Execute("select a.FileID, a.FileName, b.CountryDes, a.CountriesFinalDes from Files a, BLs b where a.InTransit=2 and a.BLID=b.BLID and b.BLType=(select BLType from BLs where BLID=" & ObjectID & ") and b.CountryDes in " & Session("Countries"))
	If Not rs.EOF Then
   		aListValues = rs.GetRows
       	CountListValues = rs.RecordCount-1
    End If
	CloseOBJs rs, Conn

%>
<HTML>
<HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<LINK REL="stylesheet" type="text/css" HREF="img/estilos.css">
<style type="text/css">
.style4 {
	font-size:10px;
	color: #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
}
</style>
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0"  onLoad="JavaScript:self.focus()">
	<TABLE cellspacing=1 cellpadding=2 width=100% align=center>
	<form name="forma" action="Docs.asp" method="post">
	<input type="hidden" name="FilesID" value="">
	<input type="hidden" name="OID" value="<%=ObjectID%>">
	<input type="hidden" name="Action" value="0">
	<tr><td class=titlelist></td><td class=titlelist><b>Pais&nbsp;Transito</b></td><td class=titlelist><b>Pais&nbsp;Destino</b></td><td class=titlelist><b>Archivo</b></td></tr>
	<%
	for i=0 to CountListValues
		FilesID = FilesID & "FilesID[" & i & "]=" & aListValues(0,i) & ";" & vbCrLf%>
		<tr>
		<td class=list><input type=checkbox name=Pos<%=i%>></td>
		<td class=list><a class=labellist href="<%=Session("VirtualPath") & aListValues(1,i)%>" target="_blank"><%=aListValues(2,i)%></a></td>
		<td class=list><a class=labellist href="<%=Session("VirtualPath") & aListValues(1,i)%>" target="_blank"><%=aListValues(3,i)%></a></td>
		<td class=list><a class=labellist href="<%=Session("VirtualPath") & aListValues(1,i)%>" target="_blank"><%=aListValues(1,i)%></a></td>
		</tr>
	<%next

	Set aListValues = Nothing
	%>
	<tr><td colspan="4" align="center"><input class="label" type="button" value="Asignar Documentos" onClick="javascript:validar();"></td></tr>
	</form>
	</TABLE>
<SCRIPT LANGUAGE="JavaScript">
var FilesID = new Array();
var ntr = "";
<%=FilesID%>

	function validar(){
		for (var i=0; i<<%=i%>; i++) {
			if (document.forma.elements["Pos" + i].checked) {
				document.forma.FilesID.value += ntr + FilesID[i];
				ntr = "|";
			}
		}
		if (document.forma.FilesID.value != "") {
			document.forma.Action.value = 4;
		}
		document.forma.submit();
	}
</SCRIPT>
</BODY>
</HTML>