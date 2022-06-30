<%
Checking "0|1|2"

Dim geocerca_mov, placa, grupo, fleet, geocerca, id_json, fecha, Action2, plate, fecha1

placa = Request("placa")
plate = Request("plate")
fleet = Request("fleet")
BLID = CheckNum(Request("BLID"))
BL = Request("BL")
OID = CheckNum(Request("OID"))

id_json = CheckNum(Request("js"))
Action2 = CheckNum(Request("Action2"))
GroupID = Request("GID")

fecha = Request("fecha")
geocerca = Request("geocerca")
geocerca_mov = Request("geocerca_mov")

CountList1Values = -1
CountList2Values = -1
CountList3Values = -1

if CountTableValues >= 0 then
	OID = aTableValues(0, 0)
    CreatedDate = aTableValues(1, 0)
    CreatedTime = aTableValues(2, 0)
    BL = aTableValues(53, 0)
    MBL = aTableValues(20, 0)
    BLID = CheckNum(aTableValues(52, 0))
    ConsignerID = aTableValues(6, 0)   
    ConsignerData = aTableValues(8, 0)
    CodeReference = aTableValues(71, 0)
end if


    'if CheckNum(CodeReference) = 0 then
    '    CodeReference = OID 
    'end if
  

	OpenConn Conn

    On Error Resume Next      

        if Action2 >=1 and Action2 <=4 then

            if Action2 < 4 then    
                           
                fecha = NewServerDate(Request("fecha"), Request.Form("fecha_h"), Request.Form("fecha_i"), "fecha_aduana", OID, CodeReference) & " "    

            end if

            QuerySelect = ""
            Select Case Action2
	        Case 1
                QuerySelect = disatel_json(0, placa, fecha, fleet, geocerca, geocerca_mov, BLID, OID, BL, CodeReference, ConsignerID, plate, 1)
	        Case 2
                QuerySelect = disatel_json(id_json, placa, fecha, fleet, geocerca, geocerca_mov, BLID, OID, BL, CodeReference, ConsignerID, plate, 1)
	        Case 3
    	        'QuerySelect = "DELETE FROM disatel_json WHERE id_json=" & id_json  
                QuerySelect = "UPDATE disatel_json SET UserID_modifica =" & CheckNum(Session("OperatorID")) & ", User_modifica = NOW(), activo = 0 WHERE id_json=" & id_json  
            End Select

            if QuerySelect = "" then

                if Action2 = 1 or Action2 = 2 then
                    response.write "<script" & ">alert('" & "Ya existe un registro con estos datos : """ & geocerca & """ - """ & geocerca_mov & """');<" & "/script>"
                end if

            else
                Conn.Execute(QuerySelect)  
            end if

            id_json = 0  
            fecha = ""
            geocerca = ""
            geocerca_mov = ""

        end if

    If Err.Number<>0 then
	    response.write "Msg :" & Err.Number & " - " & Err.Description & "<br>"  
    end if

    '//////////////////////////////////ADUANAS
    QuerySelect = "SELECT id_geocerca, geocerca FROM GeoCercas ORDER BY geocerca"
    Set rs = Conn.Execute(QuerySelect)
	if Not rs.EOF then
   		aList3Values = rs.GetRows
       	CountList3Values = rs.RecordCount-1	        		
	end if
    CloseOBJ rs	

      
dim filter2

        '//////////////////////////////////registros guardados
        On Error Resume Next      

            if CodeReference <> "" then
                filter2 = " Shipping = '" & CodeReference & "' "
            else
                filter2 = " HBLNumber = '" & IIf(BL = "--","SIN NUMERO DE CARTA PORTE",BL) & "' "
            end if
            
            filter2 = filter2 & " AND activo = 1 ORDER BY fecha DESC, geocerca_mov, id_json desc " 
                   
            aList1Values = GetPlate(filter2)
            CountList1Values = Ubound(aList1Values, 2)

        If Err.Number<>0 then
	        'response.write "disatel_json lists :" & Err.Number & " - " & Err.Description & "<br>"  
        end if
        
        '//////////////////////////////////registro actual       
        On Error Resume Next      
            filter2 = "id_json = " & id_json
            aList2Values = GetPlate(filter2)
            CountList2Values = Ubound(aList2Values, 2)
        If Err.Number<>0 then
	        'response.write "disatel_json :" & Err.Number & " - " & Err.Description & "<br>"  
        end if
                 

    '//////////////////////////////////registros guardados
                         '   0         1         2     3       4       5     6       7       8       9           10    11     12      13      14      15          16          17           18          19      20      21          22        23
'	QuerySelect = "SELECT id_json, creado_json, id, nombre, plate, fecha, evento, grupo, latitude, longitude, address, fleet, speed, course, heading, eventnemo, driver, UPPER(geocerca), geocerca_mov, BLID, BLDetailID, HBLNumber, UserID, Login FROM disatel_json LEFT JOIN Operators ON OperatorID = UserID WHERE "  
    'Set rs = Conn.Execute(QuerySelect & "1=1 order by id_json desc limit 10")       
    'Set rs = Conn.Execute(QuerySelect & "BLDetailID = " & OID & " AND activo = 1 ORDER BY fecha DESC, geocerca_mov, id_json desc")
'    Set rs = Conn.Execute(QuerySelect & "Shipping = " & CodeReference & " AND activo = 1 ORDER BY fecha DESC, geocerca_mov, id_json desc")
'	if Not rs.EOF then       
   '		aList1Values = rs.GetRows
    '   	CountList1Values = rs.RecordCount-1			
	'end if
	'CloseOBJ rs   

    '//////////////////////////////////registro actual
    'Set rs = Conn.Execute(QuerySelect & "id_json = " & id_json)
    'response.write QuerySelect & "id_json = " & id_json & "<br>"
	'if Not rs.EOF then     
    '    aList2Values = rs.GetRows
    '   	CountList2Values = rs.RecordCount-1	             		
	'end if
    'CloseOBJ rs	

    if CountList2Values >= 0 then
        fecha1 = aList2Values(5, 0)
        fecha = ConvertDate(aList2Values(5, 0),2)

        geocerca = aList2Values(17, 0)
        geocerca_mov = aList2Values(18, 0)
        placa = aList2Values(3, 0)	
        fleet = aList2Values(7, 0)	 
        plate = aList2Values(4, 0)	

        'fecha1 = TwoDigits(Day(fecha)) & "/" & TwoDigits(Month(fecha)) & "/" & Year(fecha)
    end if




    '//////////////////////////////////PLACA
    CountList2Values = -1
    QuerySelect = "SELECT a.BLNumber, b.Countries, b.TruckNo, b.Mark, b.Model, b.Motor FROM BLs a, Trucks b WHERE a.TruckID = b.TruckID and BLID=" & BLID
    'response.write QuerySelect & "<br>"
    Set rs = Conn.Execute(QuerySelect)
	if Not rs.EOF then
   		aList2Values = rs.GetRows
       	CountList2Values = rs.RecordCount-1	        		
	end if
    CloseOBJ rs	

    if  CountList2Values >= 0 then
        placa = Replace(aList2Values(2, 0), "-", "")
        plate = aList2Values(2, 0)	
    end if
    CloseOBJ Conn


    fleet = GroupData(ConsignerID)


%>
<HTML>
<HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
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
    .label1 
    {      
        background-color:#eee;
    }
    .labellist 
    {
        text-decoration:none;
    }
</style>
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">


    <% if iMensaje <> "" then %>
        alert('<%=iMensaje%>');
    <% else %>
	    function validar(Action) {

            if (Action == 3) {
                
                if (!confirm("Confirmar Borrar Registro ?")) return (false);
            }

            if (Action == 1 || Action == 2) {
                if (document.forma.placa.value == ""){ alert("No puede guardar sin numero de placa!"); return (false)};
                if (!valSelec(document.forma.geocerca_mov)){return (false)};
                if (!valSelec(document.forma.geocerca)){return (false)};            
                if (document.forma.fecha.value == ""){ alert("Seleccione Fecha"); return (false)};          
            }

            move();
            document.forma.Action2.value = Action;
            document.forma.submit();	
	     }
    <% end if %>


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

     
	function abrir(Label){
		var DateSend, Subject;
        console.log(navigator.appVersion); 	
        try {                   
			if (parseInt(navigator.appVersion) < 5) {
				DateSend = document.forma(Label).value;
			} else {
				var LabelID = SetLabelID(Label);
				DateSend = document.getElementById(LabelID).value;
			}
        }
        catch(err) {
            console.log(err); 								
        }
		Subject = '';	
		window.open('Agenda.asp?Action=1&Label=' + Label + '&DateSend=' + DateSend + '&Subj=' + Subject,'Seleccionar','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,width=170,height=160,top=150,left=350');
        return false;
	}

window.onload = function() {
  document.getElementById("Movimiento").focus();
}

</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">

<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="resizeTo(600,680);self.focus();">
	<div id="myProgress">
      <div id="myBar">10%</div>
    </div>
    <FORM name="forma" action="InsertData.asp" method="post">

    <INPUT name="Action" type=hidden value=0>
    <INPUT name="Action2" type=hidden value=0>
	<INPUT name="GID" type=hidden value="<%=GroupID%>">
	<INPUT name="OID" type=hidden value="<%=OID%>">
	<INPUT name="CD" type=hidden value="<%=CreatedDate%>">
	<INPUT name="CT" type=hidden value="<%=CreatedTime%>">	

    <INPUT name="fleet" type=hidden value="<%=fleet%>">	
    <INPUT name="placa" type=hidden value="<%=placa%>">	
    <INPUT name="plate" type=hidden value="<%=plate%>">	
    <INPUT name="js" type=hidden value="<%=id_json%>">
    <INPUT name="BLID" type=hidden value="<%=BLID%>">	
    <INPUT name="BL" type=hidden value="<%=BL%>">
    <INPUT name="CodeReference" type=hidden value="<%=CodeReference%>">

    <TABLE cellspacing=3 cellpadding=2 width=650 align=center border=0>
 


        <TR><TD class=label align=right><b>BLDetailID:</b></TD><TD class="label label1" align=left><%if OID <> 0 then response.write OID End if%></TD>
            <TD class=label align=right><b>Fecha:</b></TD><TD class="label label1" align=left><%=ConvertDate(CreatedDate,1)%></TD></TR>               
        <TR><TD class=label align=right><b>BLID:</b></TD><TD class="label label1" align=left><%=BLID%></TD> 
            <TD class=label align=right><b>C.P.H.:</b></TD><TD class="label label1" align=left><%=IIf(BL = "--","<font color=red>SIN NUMERO DE CARTA PORTE</font>",BL)%></TD></TR> 
		<TR><TD class=label align=right><b>MBL:</b></TD><TD class="label label1" align=left><%=MBL%></TD> 
		    <TD class=label align=right><b>Placa:</b></TD><TD class="label label1" align=left><%=IIf(placa = "","<font color=red>SIN PLACA</font>",plate)%></TD></TR>
        <TR><TD class=label align=right><b>Cliente:</b></TD><TD class="label label1" align=left><%=ConsignerID & " - " & ConsignerData%></TD>
            <TD class=label align=right><b>Grupo:</b></TD><TD class="label label1" align=left><%=fleet%></TD></TR>
         
        <TR><TD class=label align=right nowrap><b>Codigo Referencia:</b></TD><TD class="label label1" align=left><%=CodeReference%></TD>
            <TD class=label align=right><b></b></TD><TD class="label " align=left><%=%></TD></TR>

         

				
		<TR><TD class=label align=right height=20><b>&nbsp;</b></TD><TD class=label align=left></TD></TR> 

        <TR><TD class=label align=right><b>#:</b></TD><TD class="label label1" align=left><%=id_json%></TD>

            <TD class=label align=right><b>Movimiento:</b></TD><TD class="label label1" align=left>
            <select name="geocerca_mov" id="Movimiento" class="label" autofocus>
                    <option value="-1">Seleccionar</option>
					<option value="ENTRADA">ENTRADA</option>
                    <option value="SALIDA">SALIDA</option>
			</select>
        </TD>
        </TR> 

       <TR><TD class=label align=right height=1 ><b></b></TD><TD class=label align=left></TD></TR> 


        <TR> 
            <TD class=label align=right><b>Aduana:</b></TD><TD class="label label1" align=left>
            <select name="geocerca" id="Aduana" class="label">
					<option value="-1">Seleccionar</option>
                    <%for i=0 to CountList3Values  '<%=IIf(aList3Values(1,i) = geocerca," selected ","")%>%>
                     <option value="<%=aList3Values(1,i)%>" ><%=aList3Values(1,i)%></option>
                    <%next 
                    Set aList3Values = Nothing
                    %>
			</select>
            </TD>

            <TD class=label align=right><b>Fecha :</b></TD>
            <TD class="label label1" align=left colspan=3 nowrap>
		        <INPUT readonly="readonly" name="fecha" id="Fecha" type=text value="<%=FormatDateTime(fecha)%>" size=12 maxLength=19 class=label>		

               <INPUT type=image onClick="return abrir('fecha');" src="img/calendar.png">
  
			    <select name="fecha_h" class=label id="fecha_h">
                    <%=mesdia(0,23,2,-1,"HORA")%>	
			    </select>:<select name="fecha_i" class=label id="fecha_i">
                    <%=mesdia(0,59,2,-1,"MINUTO")%>	
			    </select>
		    </TD>
        </TR>

        
        		<TR><TD class=label align=right ><b>&nbsp;</b></TD><TD class=label align=left></TD></TR> 


        <TR>
        <td></td>
        <TD class=label align=center>
        <% if id_json > 0 then %>
        <INPUT name=enviar type=button onClick="JavaScript:validar(2);" value="&nbsp;&nbsp;Actualizar&nbsp;&nbsp;" class=label>
        </TD>
        <TD align=right>
        <INPUT name=enviar type=button onClick="JavaScript:validar(3);" value="&nbsp;&nbsp;Borrar&nbsp;&nbsp;" class=label>
        <% else %>
        <INPUT name=enviar type=button onClick="JavaScript:validar(1);" value="&nbsp;&nbsp;Grabar&nbsp;&nbsp;" class=label>
        <% end if %>
        </TD>
              <TD align=center>
        <INPUT name=enviar type=button onClick="JavaScript:validar(4);" value="&nbsp;&nbsp;Cancelar&nbsp;&nbsp;" class=label>
  
        </TR>

     

	</TABLE>

	            <TABLE cellspacing=0 cellpadding=2 width=100% align=center>
		            <TR>
		            <TD colspan=2 class=label align=right valign=top>
				            <TABLE cellspacing=5 cellpadding=2 width=100% align=center>
							        
<tr>
<td class="titlelist"><b></b></td><td class="titlelist"><b>#</b></td><td class="titlelist"><b>Fecha Aduana</b></td>
<td class="titlelist"><b>Aduana</b></td><td class="titlelist"><b>Movimiento</b></td><td class="titlelist"><b>Placa</b></td>
<td class="titlelist"><b>Evento</b></td><td class="titlelist"><b>Usuario</b></td><td class="titlelist"><b>Fecha Creacion</b></td>
</tr>
               
                    <%for i=0 to CountList1Values%>

<% 
dim onclick : onclick = ""
if CheckNum(aList1Values(20,i)) > 0 and aList1Values(17,i)  <> "" then
    onclick = "<a href='" & "InsertData.asp?OID=" & OID & "&GID=35&js=" & aList1Values(0,i) & "'><img src='img/edit.png'></a>" 
else
    onclick = ""
end if



'if CheckNum(Session("OperatorID")) = aList1Values(22,i) then
'onclick = "<a href='" & "InsertData.asp?OID=" & OID & "&GID=35&js=" & aList1Values(0,i) & "'><img src='img/edit.png'></a>" 
'else
'onclick = ""
'end if
%>



                        <tr>
                        <td class="listwarning labellist" align=center width=14><%=onclick%></td>
                        <td class="listwarning labellist"><%=aList1Values(0,i) %></td>                       
                        <td class="listwarning labellist"><%=FormatDateTime(aList1Values(5, i)) & " " & TwoDigits(Hour(aList1Values(5,i))) & ":" & TwoDigits(Minute(aList1Values(5,i))) %></td>
                        <td class="listwarning labellist"><%=aList1Values(17,i) %></td>
                        <td class="listwarning labellist"><%=aList1Values(18,i) %></td>
                        <td class="listwarning labellist"><%=aList1Values(4,i) %></td>
                        <td class="listwarning labellist"><%=aList1Values(15,i) %></td>
                        <td class="listwarning labellist"><%=IIf(aList1Values(22,i) = 0, "disatel",aList1Values(23,i)) %></td>
                        <td class="listwarning labellist"><%=aList1Values(1,i) %></td>
                        </tr>
                    <%next 
                    Set aList1Values = Nothing
                    %>

				            </TABLE>
                            
		            </TD>
	              </TR>
                  </TABLE>

                  <%="(" & fecha_h & ")"%>
	</FORM>
</BODY>
</HTML>
<script type="text/javascript">
    selecciona('forma.geocerca', '<%=geocerca%>');
    selecciona('forma.geocerca_mov', '<%=geocerca_mov%>');
    selecciona('forma.fecha_h', '<%=TwoDigits(Hour(fecha1))%>');
    selecciona('forma.fecha_i', '<%=TwoDigits(Minute(fecha1))%>');        
</script>

