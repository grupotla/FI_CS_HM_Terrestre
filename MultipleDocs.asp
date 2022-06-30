<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
Checking "0|1|2"
Dim BLID, Conn, rs, aTableValues, CountTableValues, i, HTMLCode, CIDS, AIDS, SEPS, SBLIDS, SBLID, MAILS, CTRS, EXDB
Dim Typ, GroupID, BLType, MailList, Sep, Sepp, ETA, SubQuery, Qry, CLIENTS

GroupID = CheckNum(Request("GID"))
BLType = CheckNum(Request("BTP"))
BLID = CheckNum(Request("BLID"))
Typ = CheckNum(Request("Typ")) '0=Seguros, 1=CP individual, 2=PreAlertas
SBLID = CheckNum(Request("SBLID"))

if SBLID<>0 then
    SubQuery = "and BLDetailID=" & SBLID
end if

if BLID <> 0 then
	CountTableValues = -1
	i = -1
	OpenConn Conn

    Qry = "select ClientsID, CountriesFinalDes, Clients, BLDetailID, AgentsID, Agents, Seps, ShippersID, ColoadersID, '', '', '', Countries, EXDBCountry from BLDetail where BLID=" & BLID & " " & SubQuery & " and Expired = 0 group by ClientsID, AgentsID, Seps"
	'response.write(Qry & "<br><br>")
    Set rs = Conn.Execute(Qry)

	if Not rs.EOF then
		aTableValues = rs.GetRows
		CountTableValues = rs.RecordCount-1
	End If
	closeOBJs rs, Conn

	if CountTableValues >= 0 then
        for i=0 to CountTableValues
			CIDS = CIDS & "CIDS[" & i & "]=" & aTableValues(0,i) & ";" & vbCrLf
			SBLIDS = SBLIDS & "SBLIDS[" & i & "]=" & aTableValues(3,i) & ";" & vbCrLf
			AIDS = AIDS & "AIDS[" & i & "]=" & aTableValues(4,i) & ";" & vbCrLf
			SEPS = SEPS & "SEPS[" & i & "]=" & aTableValues(6,i) & ";" & vbCrLf
            EXDB = EXDB & "EXDB[" & i & "]='" & aTableValues(13,i) & "';" & vbCrLf

			HTMLCode = HTMLCode & "<tr><td class=list><input type=checkbox name='Pos" & i & "'></td>" & _
				"<td class=list><a href=# class=labellist onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(1,i) & "</a></td>" & _
				"<td class=list><a href=# class=labellist onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(2,i) & "</a></td>" & _
				"<td class=list><a href=# class=labellist onclick=Javascript:SetList('Pos" & i & "');>" & aTableValues(5,i) & "</a></td>"

            'Para el caso de PrepAlerta se busca los correos que se debe notificar
            if Typ=2 then
                'Obteniendo el ETA para enviarlo en el Prealerta
                OpenConn Conn
	            Set rs = Conn.Execute("select BLEstArrivalDate from BLs where BLID=" & BLID)
	            if Not rs.EOF then
		            ETA = rs(0)
	            End If
	            closeOBJs rs, Conn

                OpenConn2 Conn
                MailList = ""
                Sep = ""
                'Si el BL tiene Coloader se notifica unicamente a el, caso contrario se notifica al Cliente, Shipper y Agente
                if aTableValues(8,i)<>0 then
                    'Contactos del Coloader
                    Set rs = Conn.Execute("select email from contactos where id_cliente=" & CheckNum(aTableValues(8,i)) & " and activo=true and character_length(email)>5 UNION select email from clientes where character_length(email)>5 and id_cliente=" & CheckNum(aTableValues(8,i)))
                    Sepp = ""
                    Do While Not rs.EOF
                        aTableValues(9,i) = aTableValues(9,i) & Sepp & rs(0)
                        rs.MoveNext
                        Sepp = ","
                    Loop
                    MailList = aTableValues(9,i)
                else
                    'Contactos del Cliente
                    Set rs = Conn.Execute("select email from contactos where id_cliente=" & CheckNum(aTableValues(0,i)) & " and activo=true and character_length(email)>5 UNION select email from clientes where character_length(email)>5 and id_cliente=" & CheckNum(aTableValues(0,i)))
                    'response.write "select email from contactos where id_cliente=" & CheckNum(aTableValues(0,i)) & " and activo=true and character_length(email)>5<br>"
                    Sepp = ""
                    if Not rs.EOF then
                        Do While Not rs.EOF
                            aTableValues(9,i) = aTableValues(9,i) & Sepp & rs(0)
                            rs.MoveNext
                            Sepp = ","
                        Loop
                        MailList = aTableValues(9,i)
                        Sep = ","
                    end if
                    CloseOBJ rs

                    'Contactos del Shipper
                    Set rs = Conn.Execute("select email from contactos where id_cliente=" & CheckNum(aTableValues(4,i)) & " and activo=true and character_length(email)>5 UNION select email from clientes where character_length(email)>5 and id_cliente=" & CheckNum(aTableValues(4,i)))
                    'response.write "select email from contactos where id_cliente=" & CheckNum(aTableValues(4,i)) & " and activo=true and character_length(email)>5<br>"
                    Sepp = ""
                    if Not rs.EOF then
                        Do While Not rs.EOF
                            aTableValues(10,i) = aTableValues(10,i) & Sepp & rs(0)
                            rs.MoveNext
                            Sepp = ","
                        Loop
                        MailList = MailList & Sep & aTableValues(10,i)
                        Sep = ","
                    end if
                    CloseOBJ rs

                    'Contactos del Agente
                    Set rs = Conn.Execute("select correo from agentes where agente_id=" & CheckNum(aTableValues(7,i)) & " and character_length(correo)>5")
                    'response.write "select correo from agentes where agente_id=" & CheckNum(aTableValues(7,i)) & " and character_length(correo)>5<br>"
                    Sepp = ""
                    if Not rs.EOF then
                        Do While Not rs.EOF
                            aTableValues(11,i) = aTableValues(11,i) & Sepp & rs(0)
                            rs.MoveNext
                            Sepp = ","
                        Loop
                        MailList = MailList & Sep & aTableValues(11,i)
                    end if
                    CloseOBJ rs
                end if
                CloseOBJ Conn
                
                MAILS = MAILS & "MAILS[" & i & "]='" & LCase(MailList) & "';" & vbCrLf
                CTRS = CTRS & "CTRS[" & i & "]='" & aTableValues(12,i) & "';" & vbCrLf
                CLIENTS = CLIENTS & "CLIENTS[" & i & "]='" & aTableValues(2,i) & "';" & vbCrLf
        
                'Si es Coloader solo se envia a el
                if aTableValues(8,i)<>0 then
                    HTMLCode = HTMLCode & "<td class=list><a href=# class=labellist onclick=Javascript:SetList('Pos" & i & "');>" & _
                    "COLOADER:" & LCase(aTableValues(9,i)) & "</a></td>"
                else
                    HTMLCode = HTMLCode & "<td class=list><a href=# class=labellist onclick=Javascript:SetList('Pos" & i & "');>" & _
                    "CLIENTE: " & LCase(aTableValues(9,i)) & "<br>" & _
                    "SHIPPER: " & LCase(aTableValues(10,i)) & "<br>" & _
                    "AGENTE: " & LCase(aTableValues(11,i)) & "</a></td>"
                end if                
            end if
            HTMLCode = HTMLCode & "</tr>"
		next
		
        'Si es envio de Prealerta se agrega link para accesar al catalogo en caso necesiten actualizar datos
        if Typ=2 then
            HTMLCode = HTMLCode & "<tr><td colspan=5 align=center class=label>Para actualizar o agregar datos del Catalogo, puede ingresar <A class='titlelist' href='http://10.10.1.20/catalogo_admin/login.php' target='_blank'><b>&nbsp;Aqu&iacute;&nbsp;</b></A></td></tr>"
            HTMLCode = HTMLCode & "<tr><td colspan=5 align=center><input class=label type=button value='Enviar Prealertas' onclick='Javascript:SetIDs();'></td></tr>"
        else
            HTMLCode = HTMLCode & "<tr><td colspan=5 align=center><input class=label type=button value='Ver Documentos' onclick='Javascript:SetIDs();'></td></tr>"
        end if        
	end if
%>

<HTML><HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/config.js"></SCRIPT>
<SCRIPT src="img/mainLib.js" language="JavaScript"></SCRIPT>
<SCRIPT language="javascript" src="img/validaciones.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
var CIDS = new Array();
var SBLIDS = new Array();
var AIDS = new Array();
var EXDB = new Array();
var SEPS = new Array();
var MAILS = new Array();
var CTRS = new Array();
var CLIENTS = new Array();
<%=CIDS%>
<%=SBLIDS%>
<%=AIDS%>
<%=EXDB%>
<%=SEPS%>
<%=MAILS%>
<%=CTRS%>
<%=CLIENTS%>
function SetList(Pos) {
	if (document.forma.elements[Pos].checked){
		document.forma.elements[Pos].checked = false;
	} else {
		document.forma.elements[Pos].checked = true;
	}
}

function SetIDs() {
	var sep = "";
	document.forma.CIDS.value = "";
	document.forma.SBLIDS.value = "";
	document.forma.AIDS.value = "";
    document.forma.EXDB.value = "";
	document.forma.SEPS.value = "";
    document.forma.MAILS.value = "";
    document.forma.CTRS.value = "";
    document.forma.CLIENTS.value = "";
	for (var i=0; i<<%=i%>; i++) {
		if (document.forma.elements["Pos" + i].checked) {
			document.forma.CIDS.value = document.forma.CIDS.value + sep + CIDS[i];
			document.forma.SBLIDS.value = document.forma.SBLIDS.value + sep + SBLIDS[i];
			document.forma.AIDS.value = document.forma.AIDS.value + sep + AIDS[i];
            document.forma.EXDB.value = document.forma.EXDB.value + sep + EXDB[i];
			document.forma.SEPS.value = document.forma.SEPS.value + sep + SEPS[i];     
            document.forma.MAILS.value = document.forma.MAILS.value + sep + MAILS[i];
            document.forma.CTRS.value = document.forma.CTRS.value + sep + CTRS[i];
            document.forma.CLIENTS.value = document.forma.CLIENTS.value + sep + CLIENTS[i];
			sep = "|"
		}
	}
	document.forma.submit();
}

function SetAll() {
	if (document.forma.Set.checked) {
		for (var i=0; i<<%=i%>; i++) {
			document.forma.elements["Pos" + i].checked = true;
		}
	} else {
		for (var i=0; i<<%=i%>; i++) {
			document.forma.elements["Pos" + i].checked = false;
		}

	}
}
</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="JavaScript:self.focus()">
	<FORM name="forma" action="DisplayMultipleDocs.asp" method="post" target=_self>
  	<input type="hidden" name="CIDS" value="">
	<input type="hidden" name="SBLIDS" value="">
    <input type="hidden" name="CLIENTS" value="">
  	<input type="hidden" name="AIDS" value="">
    <input type="hidden" name="EXDB" value="">
	<input type="hidden" name="SEPS" value="">
    <input type="hidden" name="MAILS" value="">
    <input type="hidden" name="CTRS" value="">
    <input type="hidden" name="ETA" value="<%=ETA%>">
	<input type="hidden" name="BLID" value="<%=BLID%>">
	<input type="hidden" name="Typ" value="<%=Typ%>">
	<input type="hidden" name="GID" value="<%=GroupID%>">
	<input type="hidden" name="BTP" value="<%=BLType%>">
	<TABLE cellspacing=0 cellpadding=2 width=100% align=center>
		<TR>
		<TD colspan=2 class=label align=right valign=top>
				<TABLE cellspacing=5 cellpadding=2 width=100% align=center>
				<tr>
                <td class=titlelist><input type=checkbox name='Set' onClick='Javascript:SetAll();'></td>
                <td class=titlelist><b>Destino Final</b></td>
                <td class=titlelist><b>Cliente</b></td>
                <td class=titlelist><b>Exportador</b></td>
                <%if Typ=2 then %>
                <td class=titlelist><b>Correos</b></td>
                <%end if %>
                </tr>
				 <%=HTMLCode%>
				</TABLE>
		</TD>
	  </TR>
		</TABLE>
  </FORM>				
</BODY>
</HTML>
<%
	Set aTableValues = Nothing
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>