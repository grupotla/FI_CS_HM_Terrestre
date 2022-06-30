<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then

Dim ObjectID, CountList1Values, CountList2Values, aTableValues, aTableValues2, Conn, rs
Dim BLNumber, Action, i, PD
Dim Trafico, UserID, Routing, Titulo, Observaciones, Usuario

Action = CheckNum(Request("Action"))
ObjectID = CheckNum(Request("id_routing"))
UserID = Session("OperatorID")
Trafico = CheckNum(Request("id_trafico"))
PD = CheckNum(Request("PD"))
Observaciones = Request("ObservacionesValue")
'Titulo = CheckNum(Request("Titulo"))
CountList1Values = -1
Set aTableValues = nothing

OpenConn2 Conn
'response.write("SELECT FreightValue, ExworksValue, CorrectionReason, WhereSays, ShouldSays, LetterType, AnotherDocsID FROM AnotherDocs WHERE BLDetailID = " & ObjectID)
Set rs = Conn.Execute("select a.routing, coalesce(c.titulo,''), coalesce(b.observaciones,''), coalesce(d.pw_name,'') from routings a left join routings_errores b on b.id_routing=a.id_routing and b.activo=true left join routings_errores_clasifica c on c.id_error_clase=b.id_error_clase left join usuarios_empresas d on d.id_usuario=b.id_usuario where a.id_routing = " & ObjectID)
If Not rs.EOF Then
    aTableValues = rs.GetRows
	CountList1Values = rs.RecordCount-1    
End If
CloseOBJ rs

Set rs = Conn.Execute("select id_error_clase, titulo from routings_errores_clasifica order by linea")
If Not rs.EOF Then
    aTableValues2 = rs.GetRows
    CountList2Values = rs.RecordCount-1
End If

CloseOBJs Conn, rs

If CountList1Values >= 0 Then
    Routing = aTableValues(0, 0)
    Usuario = aTableValues(3, 0)
End If

If (Usuario = "") Then
    Usuario = Session("Login")
End If

if Action=1 then
    RoutingError Conn, Request("OID"), Request("Titulo"), UserID, 3, Observaciones
end if

%>
<HTML>
<HEAD>
	<TITLE>Reportar Routings</TITLE>
	<META http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <style>
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
    <SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
    <SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
    <SCRIPT LANGUAGE="JavaScript">
        function validar(Action) {
            if (document.forma.Titulo.value == 0) {
                alert("Debe seleccionar el Tipo de incidente en el error.");
                return (false);
            }
            document.forma.Action.value = Action;
            document.forma.PD.value = 1;
            move();
            document.forma.submit();
            window.close();
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
        .auto-style1 {
            height: 28px;
        }
    </style>
</HEAD>
<body>
    <FORM name="forma" action="RoutingError.asp" method="post">
        <div id="myProgress">
            <div id="myBar">10%</div>
        </div>
        <INPUT name="Action" type="hidden" value="0">
        <INPUT name="PD" type="hidden" value="0">
        <INPUT name="OID" type="hidden" value="<%=ObjectID%>">
        <table id="RoutingError" align="center" class="label">
	        <tr>
		        <td colspan="2" align="center" style="font-weight:bold; font-size:15px;">
			        ROUTING REPORTADO EN SISTEMA TERRESTRE
		        </td>
	        </tr>
            <tr><td><br /></td></tr>
            <tr id="Routing">
		        <td class="auto-style1">
			        Routing:
		        </td>
		        <td class="auto-style1">
			        <input name="RoutingValue" type="text" value="<%=Routing%>" readonly/>
		        </td>
	        </tr>
	        <tr>
		        <td>Incidente:</td>
		        <td>
			        <select name="Titulo" id="Titulo">
				        <option value="0" selected="selected">Seleccione...</option>
                        <%
		                    For i = 0 To CountList2Values
	                    %>
	                    <option value="<%=aTableValues2(0,i)%>"><%response.write aTableValues2(1,i)%></option>
	                    <%
   		                    Next
	                    %>
			        </select>
		        </td>
	        </tr>
	        <tr id="Observaciones">
		        <td>
			        Observaciones
		        </td>
		        <td>
			        <textarea name="ObservacionesValue" style="width:300; height:100;"><%=Observaciones%></textarea>
		        </td>
	        </tr>
	        <tr id="Usuario">
		        <td>
			        Reportado por:
		        </td>
		        <td>
			        <input name="UsuarioValue" type="text" value="<%=Usuario%>" readonly/>
		        </td>
	        </tr>
	        <tr>

                <td colspan="2" align="center" class="label">
			        <input name="rep1" type="button" value="Grabar" onClick="Javascript: if (document.getElementById('Titulo').value > 0 && document.getElementById('ObservacionesValue').value != '') { validar(1); return false; } else { alert('Por favor llene todos los campos.'); }" class="label" />
		        </td>
	        </tr>
        </table>
    </FORM>
</body>
</HTML>
<script>
    document.forma.Routing.value = '<%=Routing%>';
    document.forma.OID.value = '<%=ObjectID%>';
    document.forma.Titulo.value = '<%=Titulo%>';
    document.forma.Observaciones.value = '<%=Observaciones%>';
    document.forma.Usuario.value = '<%=Usuario%>';
</script>
<%
Else
    Response.Redirect "redirect.asp?MS=4"
end if%>
