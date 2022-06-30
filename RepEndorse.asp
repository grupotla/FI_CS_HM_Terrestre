<%
Checking "0|1|2"

Dim TotalFreight, Master

Action = CheckNum(Request("Action"))

if CountTableValues >= 0 then
	BLID = aTableValues(1, 0)
    ConsignerData = aTableValues(2, 0)
	LtEndorseDate = aTableValues(3, 0)
	DTI = aTableValues(4, 0)
	Pos = aTableValues(5, 0) + 1
	Freight = aTableValues(6, 0)
	Freight2 = aTableValues(7, 0)
	ClientID = aTableValues(8, 0)
	Insurance = aTableValues(9, 0)
	Insurance2 = aTableValues(10, 0)
	AnotherChargesCollect = aTableValues(11, 0)
	AnotherChargesPrepaid = aTableValues(12, 0)
	DTIObservations = aTableValues(13, 0)
	Address = aTableValues(14, 0)
	Comment4 = aTableValues(15, 0)
	AgentID = aTableValues(16, 0)
	SenderData = aTableValues(17, 0)
	Sep = aTableValues(18, 0)
	BLNumber = aTableValues(19, 0)
	LtArrivalDate = aTableValues(20, 0)
	LtArrivalDeliveryDocs = aTableValues(21, 0)
	EndorseObservations = aTableValues(22, 0)
	CPDocType = aTableValues(23, 0)
	ManifestDocType = aTableValues(24, 0)
	EndorseDocType = aTableValues(25, 0)
	DTIDocType = aTableValues(26, 0)
	BLsType = aTableValues(27, 0)
	BillType = aTableValues(28, 0)
	Bill = aTableValues(29, 0)
    Mark = aTableValues(30, 0)
    ExType = aTableValues(31, 0)
    EXID = aTableValues(32, 0)
    FreightColoader = aTableValues(33, 0)
    FreightColoader2 = aTableValues(34, 0)
    InsuranceColoader = aTableValues(35, 0)
    InsuranceColoader2 = aTableValues(36, 0)
    AnotherChargesColoader = aTableValues(37, 0)
    AnotherChargesColoader2 = aTableValues(38, 0)
    ColoaderID = aTableValues(39, 0)
    Countries = aTableValues(40, 0)
    ColoaderData = aTableValues(41, 0)
end if
Set aTableValues = Nothing
CountTableValues = -1
Master = "--"

OpenConn Conn
	Set rs = Conn.Execute("select BLType from BLs where BLID=" & BLID)
	if Not rs.EOF then
		BLType = rs(0)
	end if
	CloseOBJ rs

    SQLQuery = "select a.BLDetailID, a.BLs, a.Clients, a.EXDBCountry, b.BLNumber from BLDetail a left join BLs b on b.BLID=a.BLID where a.BLID=" & BLID & " and a.ClientsID=" & ClientID & " and a.AgentsID=" & AgentID & " and a.Seps=" & Sep & " and a.HBLNumber='" & BLNumber & "' and a.Expired = 0"	
	'response.write(SQLQuery & "<br>")
	Set rs = Conn.Execute(SQLQuery)
	if Not rs.EOF then
		aTableValues = rs.GetRows
		CountTableValues = rs.RecordCount-1
	end if

    If CountTableValues >= 0 Then
        Master = aTableValues(4,0)
    End If

    if Action = 9 then

        On Error Resume Next     
         
            'SQLQuery = "CREATE TEMPORARY TABLE IF NOT EXISTS table2 AS (SELECT * FROM BLDetail WHERE BLDetailID=" & ObjectID & "); " & _ 
            '"UPDATE table2 SET BLDetailID=0, BLID=-1, HBLNumber='--', Expired=0, RefHBLNumber='', RefBLID=0, CreatedDate=NOW(), Week=0 WHERE BLDetailID=" & ObjectID & "; " & _ 
            '"INSERT INTO BLDetail SELECT * FROM table2 WHERE BLDetailID=0;" 
            '2021-03-02

            SQLQuery = "CREATE TABLE IF NOT EXISTS table2 AS (SELECT * FROM BLDetail WHERE BLDetailID=" & ObjectID & ");"
                        Conn.Execute(SQLQuery)
            SQLQuery = "UPDATE table2 SET BLDetailID=0, BLID=-1, HBLNumber='--', Expired=0, RefHBLNumber='', RefBLID=0, CreatedDate=NOW(), Week=0 WHERE BLDetailID=" & ObjectID & ";" 
                        Conn.Execute(SQLQuery)
            SQLQuery = "INSERT INTO BLDetail SELECT * FROM table2 WHERE BLDetailID=0;" 
                        Conn.Execute(SQLQuery)
            SQLQuery = "DROP TABLE table2;"
                        Conn.Execute(SQLQuery)

            Dim last_id
            SQLQuery = "SELECT LAST_INSERT_ID();"	
            Set rs = Conn.Execute(SQLQuery)
	        if Not rs.EOF then
		        last_id = rs(0)
	        end if
      	
            'response.write(SQLQuery & "<br>")
            response.write("<font color=green>La copia fue realizada correctamente con No. Registro " & last_id & ".  Anote este numero para referencias futuras.</font><br>")

        If Err.Number<>0 then
            Err.Number = 0
	        response.write "Copy :" & Err.Number & " - " & Err.Description & "<br>"  
        end if


    end if

CloseOBJs rs, Conn

TotalFreight = 0

If (FreightColoader > 0 or FreightColoader2 > 0) then
    TotalFreight = FreightColoader + FreightColoader2
Else
    TotalFreight = Freight + Freight2
End If

%>
<HTML>
<HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">


    function validar(Action) {
        
        if (Action == 2)    
    		if (!valTxt(document.forma.LtEndorseDate, 3, 5)){return (false)};

        move();
        document.forma.Action.value = Action;
		document.forma.submit();			 
	 }
	 
	 function SetLabelID(Label) {
		var LabelID = "";
		if (Label == "LtEndorseDate") {
			LabelID = "Fecha Carta de Endoso";
		} else if (Label == "LtArrivalDate") {
			LabelID = "Fecha Nota de Arribo";
		} else if (Label == "LtArrivalDeliveryDocs") {
			LabelID = "Fecha Libre para Entrega de Documentos";
		} 
		return LabelID;
	}
	 
	function abrir(Label){
	var DateSend, Subject;
		if (parseInt(navigator.appVersion) < 5) {
			DateSend = document.forma(Label).value;
		} else {
			var LabelID = SetLabelID(Label);
			DateSend = document.getElementById(LabelID).value;
		}
		Subject = '';	
		window.open('Agenda.asp?Action=1&Label=' + Label + '&DateSend=' + DateSend + '&Subj=' + Subject,'Seleccionar','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,width=170,height=160,top=250,left=250');
	}
	function Open(ObjectID){
		window.open('ItineraryCharges.asp?OID=' + ObjectID + '&GID=29','Cargos','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=1240,height=450,top=100,left=50');
}

function move() {
    document.forma.style.display = "none";
    document.getElementById('myBar').style.display = "block";
    var elem = document.getElementById("myBar");
    var width = 10;
    var id = setInterval(frame, 65);
    function frame() {
        if (width >= 100) {
            clearInterval(id);
        } else {
            width++;
            elem.style.width = width + '%';
            elem.innerHTML = width * 1 + '%';
        }
    }
}
</script>
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
    
#myBar {
    width: 10%;
    height: 15px;
    background-color: #4CAF50;
    text-align: center; /* To center it horizontally (if you want) */
    line-height: 15px; /* To center it vertically */
    color: white;
    font-weight: bold;
    display: none;
}    
</style>
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">

	<div id="myProgress">
      <div id="myBar">10%</div>
    </div>

	<TABLE cellspacing=0 cellpadding=2 width=400 align=center>
	<FORM name="forma" action="InsertData.asp" method="post">
	<INPUT name="Action" type=hidden value=0>
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
    <INPUT name="CTR" type=hidden value="<%=Countries%>">
	<INPUT name="OID" type=hidden value="<%=ObjectID%>">
	<INPUT name="CID" type=hidden value="<%=ClientID%>">
	<INPUT name="AID" type=hidden value="<%=AgentID%>">
	<INPUT name="BLID" type=hidden value="<%=BLID%>">
		<%If GroupID = 14 or GroupID = 15 Then%>
            <TR><TD class=label align=right width="40%"><b>No. Carta Porte Master:</b></TD><TD class=label align=left colspan="2"><%=Master%></TD></TR> 
            <TR><TD class=label align=right width="40%"><b>No. Carta Porte Hija:</b></TD><TD class=label align=left colspan="2"><%=BLNumber%></TD></TR> 
        <%Else %>
            <TR><TD class=label align=right width="40%"><b>No. Carta Porte:</b></TD><TD class=label align=left colspan="2"><%=BLNumber%></TD></TR> 
        <%End If %>
		<TR><TD class=label align=right><b>Consignatario:</b></TD><TD class=label align=left colspan="2"><%=ConsignerData & "<b> - ID: " & ClientID & "<b>" %></TD></TR> 
		<TR><TD class=label align=right><b>Exportador:</b></TD><TD class=label align=left colspan="2"><%=SenderData & "<b> - ID: " & AgentID & "<b>" %></TD></TR> 
        <TR><TD class=label align=right><b>Coloader:</b></TD><TD class=label align=left colspan="2"><%=ColoaderData & "<b> - ID: " & ColoaderID & "<b>" %></TD></TR> 
		<TR><TD class=label align=right><b>Código:</b></TD><TD class=label align=left colspan="2"><%if ObjectID <> 0 then response.write ObjectID End if%></TD></TR> 
		<TR><TD class=label align=right><b>Fecha Carta de Endoso:</b></TD>
			<TD class=label align=left colspan="2">
			<INPUT readonly="readonly" name="LtEndorseDate" id="Fecha Carta de Endoso" type=text value="<%=LtEndorseDate%>" size=23 maxLength=19 class=label>&nbsp;
			<INPUT type=button value="Seleccionar" onClick="JavaScript:abrir('LtEndorseDate');" class=label>
			</TD>
		</TR>
		<TR>
		  <TD class=label align=right><b>Fecha de Notificación:</b></TD>
		  <TD class=label align=left colspan="2">
			<INPUT readonly="readonly" name="LtArrivalDate" id="Fecha Nota de Arribo" type=text value="<%=LtArrivalDate%>" size=23 maxLength=19 class=label>&nbsp;
			<INPUT type=button value="Seleccionar" onClick="JavaScript:abrir('LtArrivalDate');" class=label>
			</TD></TR>
		<TR><TD class=label align=right><b>Fecha Libre para Entrega Documentos:</b></TD>
		<TD class=label align=left colspan="2">
			<INPUT readonly="readonly" name="LtArrivalDeliveryDocs" id="Fecha Libre para Entrega de Documentos" type=text value="<%=LtArrivalDeliveryDocs%>" size=23 maxLength=19 class=label>&nbsp;
			<INPUT type=button value="Seleccionar" onClick="JavaScript:abrir('LtArrivalDeliveryDocs');" class=label>
		  </TD></TR>
		<TR><TD class=label align=right><b>Flete al Cobro:</b></TD><TD class=label align=left>
			<INPUT name="Freight" id="Flete al Cobro" type=text value="<%=Freight%>" size=23 maxLength=45 class=label onKeyUp="res(this,numb);" readonly>
			</TD>
			<TD class=label align=left valign="top" rowspan="6">
				<TABLE cellspacing=0 cellpadding=2 width=100% border="1">
				<TR>
					<TD class=label align=left colspan="2"><b>Detalle de cobros:</b></TD>
				</TR>
				<TR>
					<TD class=style11 align=left><b>BL/RO:</b></TD>
					<TD class=style11 align=left><b>Consignatario:</b></TD>
				</TR>
				<%for i=0 to CountTableValues%>
				<TR>
					<TD class=label align=left><a href="Javascript:Open(<%=aTableValues(0,i)%>)"><%=aTableValues(1,i)%></a>
                     <%select Case ExType
		              case 4,5,6,7%>
		                <br><br><a href="#" onClick="Javascript:window.open('http://10.10.1.20/ventasV2/vendedores/detalle_routing.php?id_routing=<%=EXID %>&ref=<%=aTableValues(3,i)%>', 'routing_ver', 'height=600, width=700, menubar=0, resizable=1, scrollbars=1, toolbar=0');return (false);" class="menu"><font color="FFFFFF"><b>Ver RO</b>
                     <%end select %>
                    </TD>
					<TD class=label align=left><a href="Javascript:Open(<%=aTableValues(0,i)%>)"><%=aTableValues(2,i)%></a>

                    <% Dim baw                    
                    baw = "http://10.10.1.7:8181/Default.aspx?login_terrestre=1&login_user=" & Session("Login") & "&bl_no=" & BLNumber & "&blid=" & ObjectID & "&paisid=" & aTableValues(3,i) %>                    
                    <br><br><a href="#" onClick="Javascript:window.open('<%=baw%>', 'routing_ver', 'height=600, width=700, menubar=0, resizable=1, scrollbars=1, toolbar=0');return (false);" class="menu"><font color="FFFFFF"><b>Facturacion</b>                    
                    
                    </TD>

				</TR>
				<%next%>
                </TABLE>
                <%select Case BLType
                Case 2,3
                Case Else %>    
                <table cellspacing="0" cellpadding="2" border="1">
                <tr align="center">
                    <br />
                    <td class="label" align="center"><input name="rep2" type="button" onclick="Javascript:window.open('AnotherDocs.asp?Client=<%=ConsignerData%>&Freight=<%=TotalFreight%>&CP=<%=BLNumber%>&OID=<%=ObjectID%>','SIData27','height=420,width=530,menubar=0,resizable=1,scrollbars=1,toolbar=0')" value="Cartas" class="label" style="background-color: #996600; color: White; font-weight: bold;"/></td>
                </tr>
                </table>	
                <%End Select %>		
			</TD>
		</TR>
		<TR><TD class=label align=right><b>Flete Prepagado:</b></TD><TD class=label align=left>
			<INPUT name="Freight2" id="Flete Prepagado" type=text value="<%=Freight2%>" size=23 maxLength=45 class=label onKeyUp="res(this,numb);" readonly>
			</TD></TR>
		<TR><TD class=label align=right><b>Seguro al Cobro:</b></TD><TD class=label align=left>
			<INPUT name="Insurance" id="Flete al Cobro" type=text value="<%=Insurance%>" size=23 maxLength=45 class=label onKeyUp="res(this,numb);" readonly>
			</TD></TR>
		<TR><TD class=label align=right><b>Seguro Prepagado:</b></TD><TD class=label align=left>
			<INPUT name="Insurance2" id="Flete Prepagado" type=text value="<%=Insurance2%>" size=23 maxLength=45 class=label onKeyUp="res(this,numb);" readonly>
			</TD></TR>
		<TR><TD class=label align=right><b>Otros Cargos al Cobro:</b></TD><TD class=label align=left>
			<INPUT name="AnotherChargesCollect" id="Flete al Cobro" type=text value="<%=AnotherChargesCollect%>" size=23 maxLength=45 class=label onKeyUp="res(this,numb);" readonly>
			</TD></TR>
		<TR><TD class=label align=right><b>Otros Cargos Prepagado:</b></TD><TD class=label align=left>
			<INPUT name="AnotherChargesPrepaid" id="Flete Prepagado" type=text value="<%=AnotherChargesPrepaid%>" size=23 maxLength=45 class=label onKeyUp="res(this,numb);" readonly>
		    </TD></TR>
        <TR><TD class=label align=right><b>Flete al Cobro Coloader:</b></TD><TD class=label align=left>
			<INPUT name="FreightColoader" id="FreightColoader" type=text value="<%=FreightColoader%>" size=23 maxLength=45 class=label onKeyUp="res(this,numb);">
			</TD></TR>
        <TR><TD class=label align=right><b>Flete Prepagado Coloader:</b></TD><TD class=label align=left>
			<INPUT name="FreightColoader2" id="FreightColoader2" type=text value="<%=FreightColoader2%>" size=23 maxLength=45 class=label onKeyUp="res(this,numb);">
			</TD></TR>
		<TR><TD class=label align=right><b>Seguro al Cobro Coloader:</b></TD><TD class=label align=left>
			<INPUT name="InsuranceColoader" id="InsuranceColoader" type=text value="<%=InsuranceColoader%>" size=23 maxLength=45 class=label onKeyUp="res(this,numb);">
			</TD></TR>
		<TR><TD class=label align=right><b>Seguro Prepagado Coloader:</b></TD><TD class=label align=left>
			<INPUT name="InsuranceColoader2" id="InsuranceColoader2" type=text value="<%=InsuranceColoader2%>" size=23 maxLength=45 class=label onKeyUp="res(this,numb);">
			</TD></TR>
		<TR><TD class=label align=right><b>Otros Cargos al Cobro Coloader:</b></TD><TD class=label align=left>
			<INPUT name="AnotherChargesColoader" id="AnotherChargesColoader" type=text value="<%=AnotherChargesColoader%>" size=23 maxLength=45 class=label onKeyUp="res(this,numb);">
			</TD></TR>
		<TR><TD class=label align=right><b>Otros Cargos Prepagado Coloader:</b></TD><TD class=label align=left>
			<INPUT name="AnotherChargesColoader2" id="AnotherChargesColoader2" type=text value="<%=AnotherChargesColoader2%>" size=23 maxLength=45 class=label onKeyUp="res(this,numb);">
			</TD></TR>
		<TR><TD class=label align=right><b>As Agreed:</b></TD><TD class=label align=left>
			<select name="AsAgreed" id="As Agreed" class="label">
				<option value="0">NO</option>
				<option value="1">SI</option>
			</select>
			</TD></TR>
		<TR><TD class=label align=right><b>Observaciones para Carta Endoso:</b></TD><TD class=label align=left colspan="2">
			<textarea class="style10" cols="45" rows="2" name="EndorseObservations"><%=EndorseObservations%></textarea>
			</TD></TR>
		<TR><TD class=label align=right><b>Carta Porte:</b></TD><TD class=label align=left>
			<select name="CPDocType" id="Tipo de Carta Porte" class="label">
				<option value="-1">Seleccionar</option>
				<option value="0">COPIA</option>
				<option value="1">ORIGINAL</option>
				<option value="2">N/A</option>
			</select>
		</TD></TR><TR><TD class=label align=right><b>Manifiesto:</b></TD><TD class=label align=left>
			<select name="ManifestDocType" id="Tipo de Manifiesto" class="label">
				<option value="-1">Seleccionar</option>
				<option value="0">COPIA</option>
				<option value="1">ORIGINAL</option>
				<option value="2">N/A</option>
			</select>
		</TD></TR>
		</TD></TR><TR><TD class=label align=right><b>Carta de Endoso:</b></TD><TD class=label align=left>
			<select name="EndorseDocType" id="Tipo de Carta de Endoso" class="label">
				<option value="-1">Seleccionar</option>
				<option value="0">COPIA</option>
				<option value="1">ORIGINAL</option>
				<option value="2">N/A</option>
			</select>
		</TD></TR>
		<TR><TD class=label align=right><b>BL:</b></TD><TD class=label align=left>
			<select name="BLsType" id="Tipo de BL" class="label">
				<option value="-1">Seleccionar</option>
				<option value="0">COPIA</option>
				<option value="1">ORIGINAL</option>
				<option value="2">N/A</option>
			</select>
		</TD></TR>
		<TR><TD class=label align=right><b>Factura:</b></TD><TD class=label align=left>
			<select name="BillType" id="Tipo de factura" class="label">
				<option value="-1">Seleccionar</option>
				<option value="0">COPIA</option>
				<option value="1">ORIGINAL</option>
				<option value="2">N/A</option>
			</select>
		</TD></TR>
		<TR><TD class=label align=right><b>DTI:</b></TD><TD class=label align=left>
			<select name="DTIDocType" id="Tipo de DTI" class="label">
				<option value="-1">Seleccionar</option>
				<option value="0">COPIA</option>
				<option value="1">ORIGINAL</option>
				<option value="2">N/A</option>
			</select>
		</TD></TR>
		<TR><TD class=label align=right><b>No. DTI:</b></TD><TD class=label align=left colspan="2">
			<INPUT name="DTI" id="Declaración Tránsito" type=text value="<%=DTI%>" size=23 maxLength=45 class=label>
			</TD></TR>
		<TR><TD class=label align=right><b>Marcas de Expedicion, Nos. Contenedor,<br>dimensiones para inciso 28 de DTI:</b></TD><TD class=label align=left colspan="2">
			<textarea class="style10" cols="45" rows="2" name="DTIObservations"><%=DTIObservations%></textarea>
			</TD></TR>
		<TR><TD class=label align=right><b>Observaciones para DTI:</b></TD><TD class=label align=left colspan="2">
			<textarea class="style10" cols="45" rows="2" name="Comment4"><%=Comment4%></textarea>
			</TD></TR>
	<TR><TD colspan="3" class=label align=center>
			<TABLE cellspacing=0 cellpadding=2 width=200>
		    <TR>
				<TD class=label align=center><input name=rep1 type=button onClick="Javascript:window.open('Reports.asp?GID=<%=GroupID%>&BLID=<%=BLID%>&OID=<%=ObjectID%>&CID=<%=ClientID%>&AID=<%=AgentID%>&SEP=<%=Sep%>','RepEndorse','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750,height=600');return false;" value="Ver&nbsp;Carta&nbsp;Endoso" class=label></TD>
                <%select Case BLType
                Case 2,3 %>
                    <TD class=label align=center><input name=rep2 type=button onClick="Javascript:window.open('BLPrintConditions.asp?BLID=<%=BLID%>&BTP=<%=BLType%>&CTR=<%=Countries%>','SBLPrint','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750');return false;" value="Ver&nbsp;Carta&nbsp;Porte" class=label></TD>
				<%case Else %>
                    <TD class=label align=center><input name=rep2 type=button onClick="Javascript:window.open('BLPrintConditions.asp?GID=<%=GroupID%>&BLID=<%=BLID%>&SBLID=<%=ObjectID%>&BTP=<%=BLType%>&CID=<%=ClientID%>&AID=<%=AgentID%>&SEP=<%=Sep%>&CTR=<%=Countries%>','SBLPrint','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=850');return false;" value="Ver&nbsp;CP&nbsp;Individual" class=label></TD>

                    <% if Right(Countries,3) = "LTF" then %>
                    <TD class=label align=center style="border:1px solid orange;background-color:yellow;"> NEW <input name=rep2 type=button onClick="Javascript:window.open('BLPrintConditions.asp?GID=<%=GroupID%>&BLID=<%=BLID%>&SBLID=<%=ObjectID%>&BTP=<%=BLType%>&CID=<%=ClientID%>&AID=<%=AgentID%>&SEP=<%=Sep%>&CTR=<%=Countries%>&id_routing=<%=EXID%>','SBLPrint','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=850');return false;" value="&nbsp;CP&nbsp;LATIN&nbsp;" class=label></TD>
                    <% end if %>

                    <TD class=label align=center><input name=rep2 type=button onClick="Javascript:window.open('MultipleDocs.asp?GID=<%=GroupID%>&BLID=<%=BLID%>&SBLID=<%=ObjectID%>&BTP=<%=BLType%>&CID=<%=ClientID%>&AID=<%=AgentID%>&SEP=<%=Sep%>&Typ=2','SBLPrealert','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750');return false;" value="PreAlerta&nbsp;Individual" class=label></TD>
				<%End Select %>
                <!--<TD class=label align=center><input name=rep3 type=button onClick="Javascript:window.open('BLPrint.asp?GID=<%=GroupID%>&BLID=<%=BLID%>&SBLID=<%=ObjectID%>&BTP=5&CID=<%=ClientID%>&AID=<%=AgentID%>&SEP=<%=Sep%>','SBLPrint','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750');return false;" value="Ver&nbsp;Carta&nbsp;Seguro" class=label></TD>-->
				<TD class=label align=center><input name=rep4 type=button onClick="Javascript:window.open('Reports.asp?GID=13&AT=1&BLID=<%=BLID%>&OID=<%=ObjectID%>&CID=<%=ClientID%>&AID=<%=AgentID%>&SEP=<%=Sep%>','RepManif','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750,height=600');return false;" value="Ver&nbsp;Manifiesto&nbsp;Individual" class=label></TD>
				<TD class=label align=center><input name=rep5 type=button onClick="Javascript:window.open('DTIPrint.asp?GID=<%=GroupID%>&OID=<%=BLID%>&CID=<%=ClientID%>&CAID=<%=Address%>&AID=<%=AgentID%>&SEP=<%=Sep%>','DTI','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750,height=600');return false;" value="Ver&nbsp;DTI&nbsp;Individual" class=label></TD>
				<TD class=label align=center><input name=rep6 type=button onClick="Javascript:window.open('ArrivalPrint.asp?OID=<%=BLID%>&CID=<%=ClientID%>&CAID=<%=Address%>&AID=<%=AgentID%>&SEP=<%=Sep%>','Arrival','menubar=1,resizable=1,scrollbars=1,toolbar=1,width=750,height=600');return false;" value="Ver&nbsp;Nota Arribo" class=label></TD>
				<TD class=label align=center><INPUT name=enviar type=button onClick="JavaScript:validar(2)" value="&nbsp;&nbsp;Actualizar&nbsp;&nbsp;" class=label></TD>
                <TD class=label align=center><INPUT name=enviar type=button onClick="JavaScript:validar(9)" value="&nbsp;&nbsp;Copiar&nbsp;&nbsp;" class=label></TD>
			</TR>
			</TABLE>
		<TD> 
		</TR>
	</FORM>
	</TABLE>
</BODY>
<script>
		selecciona('forma.CPDocType','<%=CPDocType%>');
		selecciona('forma.ManifestDocType','<%=ManifestDocType%>');
		selecciona('forma.EndorseDocType','<%=EndorseDocType%>');
		selecciona('forma.BLsType','<%=BLsType%>');
		selecciona('forma.BillType','<%=BillType%>');
		selecciona('forma.DTIDocType', '<%=DTIDocType%>');
		selecciona('forma.AsAgreed', '<%=Mark%>');
</script>
</HTML>
<%Set aTableValues = Nothing%> 