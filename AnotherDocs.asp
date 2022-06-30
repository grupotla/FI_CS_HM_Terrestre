<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then

Dim ObjectID, CountList1Values, aTableValues, Conn, rs, i
Dim BLNumber, Action
Dim FreightValue, ExworksValue, CorrectionReason, WhereSays, ShouldSays, LetterType, AnotherDocsID

FreightValue = Request("Freight")
Action = CheckNum(Request("Action"))
BLNumber = Request("CP")
ObjectID = CheckNum(Request("OID"))
CountList1Values = -1
Set aTableValues = nothing

OpenConn Conn
'response.write("SELECT FreightValue, ExworksValue, CorrectionReason, WhereSays, ShouldSays, LetterType, AnotherDocsID FROM AnotherDocs WHERE BLDetailID = " & ObjectID)
Set rs = Conn.Execute("SELECT FreightValue, ExworksValue, CorrectionReason, WhereSays, ShouldSays, LetterType, AnotherDocsID FROM AnotherDocs WHERE BLDetailID = " & ObjectID)
If Not rs.EOF Then
    aTableValues = rs.GetRows
	CountList1Values = rs.RecordCount-1    
End If
CloseOBJs Conn, rs

If CountList1Values >= 0 Then
    FreightValue = CheckNum(aTableValues(0, 0))
    ExworksValue = CheckNum(aTableValues(1, 0))
    CorrectionReason = CheckTxt(aTableValues(2, 0))
    WhereSays = aTableValues(3, 0)
    ShouldSays = aTableValues(4, 0)
    LetterType = CheckNum(aTableValues(5, 0))
    AnotherDocsID = CheckNum(aTableValues(6, 0))
End If

if Action=1 then
    SaveAnotherDocs Conn, AnotherDocsID
end if

%>
<HTML>
<HEAD>
	<TITLE>Generaci&oacute;n de cartas</TITLE>
	<META http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
    <SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
    <SCRIPT LANGUAGE="JavaScript">
        function validar(Action) {
            document.forma.Action.value = Action;
            document.forma.submit();
        }
        function prueba(value)
        {
            if (value == 1) {
                document.getElementById("flete").style.display = "block";
                document.getElementById("exworks").style.display = "none";
                document.getElementById("mcorreccion").style.display = "none";
                document.getElementById("dondedice").style.display = "none";
                document.getElementById("debedecir").style.display = "none";
            }
            else if (value == 2) {
                document.getElementById("flete").style.display = "none";
                document.getElementById("exworks").style.display = "block";
                document.getElementById("mcorreccion").style.display = "none";
                document.getElementById("dondedice").style.display = "none";
                document.getElementById("debedecir").style.display = "none";
            }
            else if (value == 3) {
                document.getElementById("flete").style.display = "none";
                document.getElementById("exworks").style.display = "none";
                document.getElementById("mcorreccion").style.display = "block";
                document.getElementById("dondedice").style.display = "block";
                document.getElementById("debedecir").style.display = "block";
            }
            else if (value == 4) {
                document.getElementById("flete").style.display = "none";
                document.getElementById("exworks").style.display = "none";
                document.getElementById("mcorreccion").style.display = "none";
                document.getElementById("dondedice").style.display = "none";
                document.getElementById("debedecir").style.display = "none";
            }
        }
    </SCRIPT>
    <LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
    <style type="text/css">
    .style10 {font-size: 10px; font-family: Verdana, Arial, Helvetica, sans-serif; }
        .style11
        {
            font-family: Verdana, Arial, Helvetica, sans-serif;
            font-size: 7.6pt;
            color: #000000;
            font-weight: none;
            text-transform: none;
            text-decoration: none;
            height: 18px;
        }
    </style>
</HEAD>
<body>
    <FORM name="forma" action="AnotherDocs.asp" method="post">
        <INPUT name="Action" type="hidden" value="0">
        <INPUT name="OID" type="hidden" value="<%=ObjectID%>">
        <table id="AnotherLetters" align="center" class="label">
	        <tr>
		        <td colspan="2" align="center" style="font-weight:bold; font-size:15px;">
			        <%=Request("Client")%>
		        </td>
	        </tr>
            <tr><td><br /></td></tr>
	        <tr>
		        <td>Seleccionar Tipo de Carta</td>
		        <td>
			        <select name="LetterType" id="LetterType" onchange="Javascript:prueba(this.options[this.selectedIndex].value);">
				        <option value="0" selected>Seleccionar...</option>
				        <option value="1">Confirmación de flete</option>
				        <option value="2">Gastos exworks</option>
				        <option value="3">Corrección de datos</option>
			        </select>
		        </td>
	        </tr>
	        <tr id="flete" style="display: none;">
		        <td>
			        Flete Terrestre USD
		        </td>
		        <td>
			        <input name="FreightValue" type="text" onkeyup="res(this,numb);" onkeydown="res(this,numb);" value="<%=FreightValue%>"/>
		        </td>
	        </tr>
	        <tr id="exworks" style="display: none;">
		        <td>
			        Exworks USD
		        </td>
		        <td>
			        <input name="ExworksValue" type="text" onkeyup="res(this,numb);" onkeydown="res(this,numb);" value="<%=ExworksValue%>" />
		        </td>
	        </tr>
	        <tr id="mcorreccion" style="display: none;">
		        <td>
			        Motivo corrección
		        </td>
		        <td>
			        <input type="text" name="CorrectionReason" value="<%=CorrectionReason%>"/>
		        </td>
	        </tr>
	        <tr id="dondedice" style="display: none;">
		        <td>
			        Donde dice
		        </td>
		        <td>
			        <textarea name="WhereSays" style="width:300; height:100;"><%=WhereSays%></textarea>
		        </td>
	        </tr>
	        <tr id="debedecir" style="display: none;">
		        <td>
			        Debe decir
		        </td>
		        <td>
			        <textarea name="ShouldSays" style="width:300; height:100;"><%=ShouldSays%></textarea>
		        </td>
	        </tr>
	        <tr>

                <td colspan="2" align="center" class="label">
			        <input name="rep1" type="button" value="Grabar/Imprimir" onClick="Javascript:if(document.getElementById('LetterType').value > 0){validar(1);window.open('Reports.asp?GID=<%=35%>&TC='+document.getElementById('LetterType').value+'&CP=<%=BLNumber%>&Client=<%=Request("Client")%>&OID=<%=ObjectID%>','RepEndorse','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750,height=700,left=400');window.close();return false;}else{alert('Por favor seleccione el tipo de carta');}" class="label" />
		        </td>
	        </tr>
        </table>
    </FORM>
</body>
</HTML>
<script>
    document.forma.FreightValue.value = '<%=FreightValue%>';
    document.forma.OID.value = '<%=ObjectID%>';
    document.forma.ExworksValue.value = '<%=ExworksValue%>';
    document.forma.CorrectionReason.value = '<%=CorrectionReason%>';
    document.forma.WhereSays.value = '<%=WhereSays%>';
    document.forma.ShouldSays.value = '<%=ShouldSays%>';
</script>
<%
Else
    Response.Redirect "redirect.asp?MS=4"
end if%>
