<!-- 
    METADATA 
    TYPE="typelib" 
    UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  
    NAME="CDO for Windows 2000 Library" 
-->
<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
Checking "0|1|2"

On Error Resume Next 


Dim BLID, Typ, CIDS, AIDS, SEPS, SBLIDS, CLIENTS, MAILS, CantDocs, i, GroupID, BLType, HTMLConditions, HTMLBL, ETA, CTRS, EXDB
Dim StartPositionNoBL, FinishPositionNoBL, LenNoBL, BL, Company

GroupID = CheckNum(Request("GID"))
BLID = CheckNum(Request("BLID"))
BLType = CheckNum(Request("BTP"))
Typ = CheckNum(Request("Typ")) '0=Seguros, 1=CP individual, 2=Prealerta
ETA = Request("ETA")
CIDS = Split(request.Form("CIDS"),"|")
SBLIDS = Split(request.Form("SBLIDS"),"|")
AIDS = Split(request.Form("AIDS"),"|")
EXDB = Split(request.Form("EXDB"),"|")
SEPS = Split(request.Form("SEPS"),"|")
MAILS = Split(request.Form("MAILS"),"|")
CTRS = Split(request.Form("CTRS"),"|")
CLIENTS = Split(request.Form("CLIENTS"),"|")
CantDocs = ubound(CIDS)
'response.write "[" & Typ & "]"
%>
<HTML>
<HEAD>
<SCRIPT language="javascript"> 
function imprimir()
{ if ((navigator.appName == "Netscape")) { window.print() ; 
} 
else
{ var WebBrowser = '<OBJECT ID="WebBrowser1" WIDTH=0 HEIGHT=0 CLASSID="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2"></OBJECT>'; 
document.body.insertAdjacentHTML('beforeEnd', WebBrowser); WebBrowser1.ExecWB(7, -1); WebBrowser1.outerHTML = "";
}
}
</SCRIPT> 
</HEAD>
<!--<BODY onload="imprimir();">-->
<BODY>
<style type="text/css">
<!--
body {
	margin: 0px;
}
.style4 {
	font-size:10px;
	color: #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	border-bottom-style:solid;
	border-left-style:solid;
	border-right-style:solid;
	border-top-style:solid;
	border-collapse:collapse;
	border-width: 1px;
}
-->
</style>

<%

if Session("OperatorID") = 1237 then
    response.write "Entro DisplayMultipleDocs<br>"
end if

Select Case Typ
Case 0
	for i=0 to CantDocs%>
		<iframe src="BLPrint.asp?GID=<%=GroupID%>&BLID=<%=BLID%>&SBLID=<%=SBLIDS(i)%>&BTP=<%=BLType%>&CID=<%=CIDS(i)%>&AID=<%=AIDS(i)%>&SEP=<%=SEPS(i)%>" frameborder="0" framespacing="0" scrolling="auto" width="750" height="1300"></iframe>
        <iframe src="Conditions.asp?EXDBCountry=<%=EXDB(i)%>" frameborder="0" framespacing="0" scrolling="auto" width="750" height="1300"></iframe>
<%	next
Case 1
	for i=0 to CantDocs%>
		<iframe src="Reports.asp?GID=13&AT=<%=16-GroupID%>&BLID=<%=BLID%>&OID=<%=SBLIDS(i)%>&CID=<%=CIDS(i)%>&AID=<%=AIDS(i)%>&SEP=<%=SEPS(i)%>" frameborder="0" framespacing="0" scrolling="auto" width="800" height="960"></iframe>
<%	next
Case 2

    response.write "<br><br><br><table align=center>"

    'Condiciones de Carga
    HTMLConditions = GetHTMLSource ( "http://10.10.1.21:8181/terrestre/admin/Conditions.asp?EXDBCountry=" & EXDB(i) )
    'HTMLConditions = GetHTMLSource ( "Conditions.asp?EXDBCountry=" & EXDB(i) )

    'Se obtiene cada CP Hija y se notifica a sus contactos
    for i=0 to CantDocs
        if InStr(1,CTRS(i),"LTF") then
            Company = "LATIN FREIGHT"
        else
            Company = "AIMAR"
        end if
        HTMLBL = GetHTMLSource ( "http://10.10.1.21:8181/terrestre/admin/BLPrint.asp?GID=" & GroupID & "&BLID=" & BLID & "&SBLID=" & SBLIDS(i) & "&BTP=" & BLType & "&CID=" & CIDS(i) & "&AID=" & AIDS(i) & "&SEP=" & SEPS(i) & "&EXDB=" & EXDB(i) )
        'HTMLBL = GetHTMLSource ( "BLPrint.asp?GID=" & GroupID & "&BLID=" & BLID & "&SBLID=" & SBLIDS(i) & "&BTP=" & BLType & "&CID=" & CIDS(i) & "&AID=" & AIDS(i) & "&SEP=" & SEPS(i) & "&EXDB=" & EXDB(i) )
        'en la impresion del BL se encuentran los tags ocultos <!--BLN--> y <!--/BLN--> para poder obtener el No de CP Hija
        StartPositionNoBL = InStr(1,HTMLBL,"<!--BLN-->")+10
        FinishPositionNoBL = InStr(1,HTMLBL,"<!--/BLN-->")
        LenNoBL = FinishPositionNoBL-StartPositionNoBL
        
        if LenNoBL > 0 then
            BL = Mid(HTMLBL,StartPositionNoBL,LenNoBL)

                Select Case EXDB(i)
                    Case "GTLTF","SVLTF","HNLTF","NILTF","CRLTF","PALTF","MXLTF","BZLTF"
                        HTMLBL = "</b>Estimado Cliente:<br><br>Es un gusto informarle que su carga esta siendo despachada, a continuación se detalla carta de porte incluyendo términos y condiciones del servicio contratado<br><br>" & _
                            HTMLBL & "<br>" & HTMLConditions & "<br><br>" & _
                            "Para solicitar un usuario y password para acceso al tracking puede comunicarse con nuestro representante de Servicio al Cliente<br><br>" & _
                            "Estamos para servirle, atentamente,<br><br>" & _
                            "Latin Freight"

                    Case "GTTLA","SVTLA","HNTLA","NITLA","CRTLA","PATLA","MXTLA","BZTLA"
                        HTMLBL = "</b>Estimado Cliente:<br><br>Es un gusto informarle que su carga esta siendo despachada, a continuación se detalla carta de porte incluyendo términos y condiciones del servicio contratado<br><br>" & _
                            HTMLBL & "<br>" & HTMLConditions & "<br><br>" & _
                            "Estamos para servirle, atentamente,<br><br>" & _
                            "GRUPO TLA"

                    Case Else
                        HTMLBL = "</b>Estimado Cliente:<br><br>Es un gusto informarle que su carga esta siendo despachada, a continuación se detalla carta de porte incluyendo términos y condiciones del servicio contratado<br><br>" & _
                            HTMLBL & "<br>" & HTMLConditions & "<br><br>" & _
                            "Para solicitar un usuario y password para acceso al tracking puede comunicarse con nuestro representante de Servicio al Cliente<br><br>" & _
                            "Estamos para servirle, atentamente,<br><br>" & _
                            "Aimar Group"
                End Select

                Dim subject, Chain, iI, LenChain 
                subject = "Prealerta " & BL & " con ETA: " & ETA & " / " & CLIENTS(i)

                'response.write "(" & MAILS(i) & ")<br>"

                Chain = split(MAILS(i), ",")
                LenChain = ubound(Chain)
                for iI = 1 to LenChain
                    'response.write "(" & Trim(Chain(iI)) & ")<br>"
                    SendMail HTMLBL, Trim(Chain(iI)), subject, EXDB(i)
                next

                'response.write "//////////////////////////////////////////////<br>"
            
                SendMail HTMLBL, Session("OperatorEmail"), subject, EXDB(i) 

            response.write "<tr><td class=style4>Se envio prealerta de CP: <b>" & BL & "</b></td></tr>"
        else
            response.write "<tr><td class=style4>No se pudo enviar prealerta BLDetailID: <font color=red>" & SBLIDS(i) & "</font></td></tr>"
        end if
    next
    response.write "</table>"




End Select%>


<%
If Err.Number<>0 then
	response.write "DisplayMultipleDocs :" & Err.Number & " - " & Err.Description & "<br>"  
    Err.Number = 0
end if
%>
</BODY>
</HTML>

