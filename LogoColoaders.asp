<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Dim CountTableValues, aTableValues, Conn, rs 
Dim ColoaderID, ColoaderData, Action, PathLogo

Action = Request("Action")
ColoaderID = Request("ColoadersID")
ColoaderData = Request("Coloaders")
PathLogo = ""
CountTableValues = -1
Set aTableValues = nothing

Select Case Action
    Case 1
        DatosLogo()        
    Case 2
        OpenConn Conn
        Conn.Execute("UPDATE LogosColoader SET Expired = 1 WHERE ColoaderID = " & ColoaderID & " AND Expired = 0")
        Conn.Execute("INSERT INTO LogosColoader (ColoaderID, PathLogo, UserID) VALUES (" & ColoaderID & ", 'http://www.aimargroup.com/logos/" & Request("FileName1") & "', " & Session("OperatorID") & ")")
        CloseOBJs Conn, rs
        DatosLogo()
    Case 3
        OpenConn Conn
        Conn.Execute("UPDATE LogosColoader SET Expired = 1 WHERE ColoaderID = " & ColoaderID & " AND Expired = 0")
        CloseOBJs Conn, rs
        'CountTableValues = -1
        ColoaderID = 0
        ColoaderData = ""
        DatosLogo()
End Select

Set aTableValues = Nothing

Sub DatosLogo()
    If CheckNum(ColoaderID) > 0 Then
        OpenConn Conn
        Set rs = Conn.Execute("SELECT PathLogo FROM LogosColoader WHERE ColoaderID = " & ColoaderID & " and Expired = 0 ")
        If Not rs.EOF Then
            aTableValues = rs.GetRows
	        CountTableValues = rs.RecordCount-1    
        End If
        if CountTableValues >= 0 then
	        PathLogo = aTableValues(0, 0)
        end if
        CloseOBJs Conn, rs
    End If
End Sub

%>
<html>
    <head>
        <title>Logos Coloader</title>
    </head>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <script type="text/javascript" language="javascript" src="img/validaciones.js"></script>
    <script type="text/javascript" language="javascript" src="img/vals.js"></script>
    <script type="text/javascript" language="JavaScript">
        function validar(Action) 
        {
            if (Action != 2) {
                if (Action == 3) {
                    if (!confirm("¿Eliminar logo para el coloader " + document.forma.ColoadersID.value + "?")) {
                        return (false);
                    }
                }
                document.forma.Action.value = Action;
	            document.forma.submit();
	        }
	        else{
                var fileName = document.getElementById("fileName").value;
                if (fileName == "") {
                    alert("No ha seleccionado ningún logo para cargar");
                    return (false);
                } else {
                    document.forma2.FileName1.value = document.forma2.image.value;
                }
                if (!confirm("¿Guardar logo para el coloader " + document.forma.ColoadersID.value + "?")) {
                    return (false);
                } else {
                    document.forma.Action.value = Action;
                    document.forma2.submit();
                }
	  	    }
        }
        function GetData(GID) {
            window.open('Search_BLData.asp?GID=' + GID, 'BLData', 'height=400,width=460,menubar=0,resizable=1,scrollbars=1,toolbar=0,status=1');
        }
        function validateFileType() {
            var fileName = document.getElementById("fileName").value;
            var idxDot = fileName.lastIndexOf(".") + 1;
            var extFile = fileName.substr(idxDot, fileName.length).toLowerCase();
            if (extFile == "jpg" || extFile == "jpeg" || extFile == "png") {
                document.forma2.agregar.disabled = false;
                //TO DO
            } else {
                alert("Solo es posible cargar archivos de tipo Imagen .jpg/.jpeg/.png");
                window.forma.reset();
            }
        }
    </script>
    <link rel="stylesheet" type="text/css" href="img/estilos.css" />
    <body text="#000000" vLink="#000000" aLink="#000000" link="#000000" bgColor="#ffffff" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="JavaScript:self.focus()" style="text-align:center; width:auto;">
        <style type="text/css">
            .style1
            {
                font-family: Verdana, Arial, Helvetica, sans-serif;
                font-size: 7.6pt;
                color: #000000;
                font-weight: none;
                text-transform: none;
                text-decoration: none;
                width: auto;
            }
            .style2
            {
                font-family: Verdana, Arial, Helvetica, sans-serif;
                font-size: 7.6pt;
                color: #000000;
                font-weight: none;
                text-transform: none;
                text-decoration: none;
                width: auto;
            }
        </style>
	    <form name="forma" action="LogoColoaders.asp" method="post">
	        <input name="Action" type=hidden value=0>
            <input name="ColoadersID" type="hidden" value="<%=ColoaderID%>" readonly>
            <table style="text-align:center; width:auto;">
		        <tr>
                    <td class=style2 colspan=1><b>Coloader:</b></td>
                    <TD class=style1 colspan=1><INPUT TYPE=text class=label name="Coloaders" id="Text3" value="<%=ColoaderData%>" maxlength="200" size="35" readonly></td>
                    <td class=label colspan=1><a class="submenu" href="http://10.10.1.20/catalogo_admin/login.php" target="_blank"><font color="FFFFFF"><b>Nuevo</b></font></a></td>
                    <td class=style1 colspan=1><a href="#" onClick="Javascript:GetData(36);return (false);" class="menu"><font color="FFFFFF"><b>Buscar</b></font></a></td>
                </tr>
                <tr>
                    <td class=label colspan=4>
                        <input name=Consultar type=button onClick="JavaScript:validar(1);" value="&nbsp;&nbsp;Consultar&nbsp;&nbsp;" class=label disabled>
                    </td>
                </tr>
                <tr>
		            <%if CountTableValues = -1 then%>
		            <%else%>
			            <td class=label colspan=4>
                            <img src="<%=PathLogo%>" border=1 />
                        </td>
		            <%end if%>
		        </tr>
	        </table>
	    </form>
        <form METHOD="POST" ENCTYPE="multipart/form-data" name="forma2" action="UploadBinaryFiles.asp?Path=1&ColoadersID=<%=ColoaderID%>&Coloaders=<%=ColoaderData%>">
            <input name="FileName1" type=hidden value="">
            <table style="text-align:center; width:auto;">
                <%if ColoaderID > 0 Then%>
                    <tr>
                        <td colspan=4>
                            <input name="image" type="file" id="fileName" accept=".jpg,.jpeg,.png" onchange="Javascript:validateFileType();" />
                        </td>
                    </tr>
                <%End If %>
                <%if CountTableValues = -1 then%>
                    <%if ColoaderID > 0 then %>
                        <tr>
                            <td class=label>
                                <input name=agregar type=button onClick="JavaScript:validar(2)" value="&nbsp;&nbsp;Agregar&nbsp;&nbsp;" class=label disabled>
                            </td>
                        </tr>
                    <%end if %>
		        <%else%>
			        <tr>
                        <td class="label">
                            <input name=enviar type=button onClick="JavaScript:validar(2)" value="&nbsp;&nbsp;Actualizar&nbsp;&nbsp;" class=label>
                        </td>
                        <td class="label">
                            <input name=enviar type=button onClick="JavaScript:validar(3)" value="&nbsp;&nbsp;Eliminar&nbsp;&nbsp;" class=label>
                        </td>
                    </tr>
		        <%end if%>
            </table>
        </form>
    </body>
</html>
<%
Else
    Response.Redirect "redirect.asp?MS=4"
end if%>