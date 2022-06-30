<%
Checking "0|1|2"

Dim OID : OID = Request("OID")

Dim TripLoadDate, TripUnloadDate, TripUnloadPartialDate, TruckGps, TruckAvailable


Dim TripRequestDate1, TripLoadDate1, TripUnloadDate1, TripUnloadPartialDate1


GroupID = Request("GID")

TripRequestDate = Request.Form("TripRequestDate")
TripLoadDate = Request.Form("TripLoadDate")
TripUnloadDate = Request.Form("TripUnloadDate")
TripUnloadPartialDate = Request.Form("TripUnloadPartialDate")

if CountTableValues >= 0 then
	ObjectID = aTableValues(0, 0)
    CreatedDate = aTableValues(1, 0)
    CreatedTime = aTableValues(2, 0)
    BL = aTableValues(53, 0)
	MBL = aTableValues(20, 0)

    
    BLID = CheckNum(aTableValues(52, 0))
    ConsignerID = aTableValues(6, 0)   
    ConsignerData = aTableValues(8, 0)
    CodeReference = aTableValues(71, 0)
    


    OtherDocs = aTableValues(76, 0)

    TripRequestDate = aTableValues(75, 0)
    TripLoadDate = aTableValues(77, 0)
    TripUnloadDate = aTableValues(79, 0)
    TripUnloadPartialDate = aTableValues(78, 0)
end if

'TripRequestDate1 = TwoDigits(Day(TripRequestDate)) & "/" & TwoDigits(Month(TripRequestDate)) & "/" & Year(TripRequestDate)
'TripLoadDate1 = TwoDigits(Day(TripLoadDate)) & "/" & TwoDigits(Month(TripLoadDate)) & "/" & Year(TripLoadDate)
'TripUnloadDate1 = TwoDigits(Day(TripUnloadDate)) & "/" & TwoDigits(Month(TripUnloadDate)) & "/" & Year(TripUnloadDate)
'TripUnloadPartialDate1 = TwoDigits(Day(TripUnloadPartialDate)) & "/" & TwoDigits(Month(TripUnloadPartialDate)) & "/" & Year(TripUnloadPartialDate)

TripRequestDate1 = ConvertDate(TripRequestDate,2)
TripLoadDate1 = ConvertDate(TripLoadDate,2)
TripUnloadDate1 = ConvertDate(TripUnloadDate,2)
TripUnloadPartialDate1 = ConvertDate(TripUnloadPartialDate,2)



'function ReadPlate() 
    


'End Function


 'ReadPlate 



 

    if Action = 1 or Action = 2 then   
   
        'RowsCount = -1
        'id_json = 0
        'QuerySelect = ""

        On Error Resume Next      
            'Buscaba el registro de la placa en disatel_json 
            'filter2 = " Shipping = '" & CodeReference & "' and nombre <> '' AND activo IN (0,3) ORDER BY fecha DESC, id_json desc LIMIT 1"    
            'Rows = GetPlate(filter2)

            'On Error Resume Next      
            '    RowsCount = Ubound(Rows, 2) 
            'If Err.Number<>0 then
                'response.write "cierre :" & Err.Number & " - " & Err.Description & "<br>"  
            'end if

            'if RowsCount <> -1 then
            '    id_json = CheckNum(Rows(0, 0)) 'captura el id
            'end if  
            
            'response.write "(" & TripUnloadDate & ")(" & id_json & ")(" & RowsCount & ")<br>"
             
            if TripUnloadDate <> "" then 'si la fecha de descarga fue grabada debe generar el registro de cierre en disatel_json
                'dim test : test = Year(TripUnloadDate) & "-" & TwoDigits(Month(TripUnloadDate)) & "-" & TwoDigits(Day(TripUnloadDate)) & IIf(Len(TripUnloadDate) > 10, Right(TripUnloadDate,9), "")                           
                'QuerySelect = disatel_json(id_json, placa, test, "", "", "", BLID, OID, BL, CodeReference, ConsignerID, plate, 3)
            
                'habilita placa de gps

                'if TruckAvailable = "(OFF)" then 'TruckGps = "(GPS)" and 
                    ActivePlate BLID, 1
                'end if

            else 
                'si la fecha no fue grabada pone inactivo el registro si existe
                'QuerySelect = "UPDATE disatel_json SET UserID_modifica =" & CheckNum(Session("OperatorID")) & ", User_modifica = NOW(), activo = 0 WHERE id_json=" & id_json  

                'if TruckAvailable = "(ON)" then 'TruckGps = "(GPS)" and 
                    ActivePlate BLID, 0
                'end if

            end if

            
     'ReadPlate 

                 
            'If QuerySelect <> "" then
                'response.write "" & QuerySelect & "<br>"
            '    OpenConn Conn
	        '    Conn.Execute(QuerySelect)   
            'end if

        If Err.Number<>0 then
	        response.write "cierre :" & Err.Number & " - " & Err.Description & "<br>"  
        end if


    end if 





        

    TruckID  = 0
    CountryDes = ""
    plate = ""
    placa = ""
    TruckGps = ""
    TruckAvailable = ""
    BLNumber = ""




    On Error Resume Next      
        Dim Rows, RowsCount ', filter2 : filter2 = " Shipping = '" & CodeReference & "' and nombre <> '' AND activo = 1 ORDER BY fecha DESC, id_json desc LIMIT 1"    
        
        RowsCount = -1

        'response.write "(" & CodeReference & ")(" & BLID & ")"

        'if CodeReference <> "" then

            'Rows = GetPlate(filter2)
        
            Rows = GetPlate2(BLID)

                On Error Resume Next      
                    RowsCount = Ubound(Rows, 2) 
                If Err.Number<>0 then
                    'response.write "cierre :" & Err.Number & " - " & Err.Description & "<br>"  
                end if

            if RowsCount <> -1 then           
                TruckID = Rows(0, 0)
                placa = Rows(1, 0)	
                plate = Rows(2, 0)
                CountryDes = Rows(3, 0)     
                TruckGps = Rows(4, 0)
                TruckAvailable = Rows(5, 0)     
                BLNumber = Rows(6, 0)     
            end if

        'end if

    If Err.Number<>0 then

	    'response.write "fechas operativas :" & Err.Number & " - " & Err.Description & "<br>"  
    end if
           

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
</style>
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">


    <% if iMensaje <> "" then %>
        alert('<%=iMensaje%>');
    <% else %>
	    function validar(Action) {

            /*///////////////////////////////////////////////////////////2020-09-08
            if (!valTxt(document.forma.CodeReference, 1, 10)){return (false)};
            if (!valSelec(document.forma.TripRequestDate_y)){return (false)};
			if (!valSelec(document.forma.TripRequestDate_m)){return (false)};
			if (!valSelec(document.forma.TripRequestDate_d)){return (false)};
            ////////////////////////////////////////////////////////////*/
            move();
            document.forma.Action.value = Action;
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

	}

</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="resizeTo(600,680);self.focus();">
	<div id="myProgress">
      <div id="myBar">10%</div>
    </div>
    <FORM name="forma" action="InsertData.asp" method="post">
	    <INPUT name="Action" type=hidden value=0>
	    <INPUT name="GID" type=hidden value="<%=GroupID%>">
	    <INPUT name="OID" type=hidden value="<%=ObjectID%>">
	    <INPUT name="CD" type=hidden value="<%=CreatedDate%>">
	    <INPUT name="CT" type=hidden value="<%=CreatedTime%>">	 
	    <INPUT name="CodeReference" type=hidden value="<%=CodeReference%>">	 

    <TABLE cellspacing=3 cellpadding=2 width=650 align=center border=0>

        <TR><TD class=label align=right><b>BLDetailID:</b></TD><TD class="label label1" align=left><%if OID <> 0 then response.write OID End if%></TD>
            <TD class=label align=right><b>Fecha:</b></TD><TD class="label label1" align=left><%=ConvertDate(CreatedDate,2)%></TD></TR>               
       <TR><TD class=label align=right><b>BLID:</b></TD><TD class="label label1" align=left><%=BLID%></TD> 
            <TD class=label align=right><b>C.P.M.:</b></TD><TD class="label label1" align=left nowrap><%=BLNumber%></TD></TR> 


 		<TR><TD class=label align=right><b>MBL:</b></TD><TD class="label label1" align=left><%=MBL%></TD> 
		    <TD class=label align=right><b>Placa:</b></TD><TD class="label label1" align=left><%=IIf(placa = "","<font color=red>SIN PLACA</font>", placa & " " & TruckGps & " " & TruckAvailable)%></TD></TR>

        <TR><TD class=label align=right><b>Cliente:</b></TD><TD class="label label1" align=left width=350><%=ConsignerID & " - " & ConsignerData%></TD>
            <TD class=label align=right nowrap><b>Codigo Referencia:</b></TD><TD class="label label1" align=left><%=CodeReference%></TD></TR>
  
        <TR><TD class=label align=right><b>Pais Destino:</b></TD><TD class="label label1" align=left width=350><%=CountryDes%></TD>
            <TD class=label align=right><b>C.P.H.:</b></TD><TD class="label label1" align=left nowrap><%=IIf(BL = "--","<font color=red>SIN NUMERO DE CARTA PORTE</font>",BL)%></TD></TR> 

  
        <TR><TD class=label align=right height=30><b> </b></TD><TD class="label " align=left></TD>
            <TD class=label align=right><b></b></TD><TD class="label" align=left></TD></TR>



        <TR><TD class=label align=right nowrap><b>Fecha Requerimiento:</b></TD>
        <TD class=label align=left width=300 colspan=3>
		    <INPUT  name="TripRequestDate" id="Fecha Real de Salida" type=text value="<%=FormatDateTime(TripRequestDate1)%>" size=23 maxLength=19 class=label readonly>&nbsp;
		    <INPUT type=button value="Seleccionar" onClick="JavaScript:abrir('TripRequestDate');" class=label>
        </TD></TR>

        <TR><TD class=label align=right height=10><b> </b></TD><TD class="label " align=left></TD>
            <TD class=label align=right><b></b></TD><TD class="label " align=left></TD></TR>

        <TR><TD class=label align=right><b>Fecha Carga:</b></TD>
        <TD class=label align=left width=300 colspan=3>
		    <INPUT name="TripLoadDate" id="Text1" type=text value="<%=FormatDateTime(TripLoadDate1)%>" size=23 maxLength=19 class=label readonly>&nbsp;
		    <INPUT type=button value="Seleccionar" onClick="JavaScript:abrir('TripLoadDate');" class=label>
        </TD></TR>

        <TR><td></td>
		<TD class=label align=left colspan=2>
			
			<select name="TripLoadDate_h" class=label id="TripLoadDate_h">
                <%=mesdia(0,23,2,-1,"HORA")%>				
			</select>:<select name="TripLoadDate_i" class=label id="TripLoadDate_i">
                <%=mesdia(0,59,2,-1,"MINUTO")%>					
			</select>
		</TD></TR>

        <TR><TD class=label align=right height=10><b> </b></TD><TD class="label " align=left></TD>
            <TD class=label align=right><b></b></TD><TD class="label " align=left></TD></TR>

<% if CountryDes = "H2TLA" then %>	
            
        <TR><TD class=label align=right nowrap><b>Fecha Descarga Parcial:</b></TD>
        <TD class=label align=left width=300 colspan=3>
		    <INPUT  name="TripUnloadPartialDate" id="Text2" type=text value="<%=FormatDateTime(TripUnloadPartialDate1)%>" size=23 maxLength=19 class=label readonly>&nbsp;
		    <INPUT type=button value="Seleccionar" onClick="JavaScript:abrir('TripUnloadPartialDate');" class=label>
        </TD></TR>

        <tr><td></td>
		<TD class=label align=left colspan=2>
		
			<select name="TripUnloadPartialDate_h" class=label id="TripUnloadPartialDate_h">
                <%=mesdia(0,23,2,-1,"HORA")%>	
			</select>:<select name="TripUnloadPartialDate_i" class=label id="TripUnloadPartialDate_i">
                <%=mesdia(0,59,2,-1,"MINUTO")%>	
			</select>
		</TD></TR>
        
                <TR><TD class=label align=right height=10><b> </b></TD><TD class="label " align=left></TD>
            <TD class=label align=right><b></b></TD><TD class="label " align=left></TD></TR>

<% else %>

		    <INPUT  name="TripUnloadPartialDate" id="Text4" type=hidden value="" >


<% end if %>	
      
        <TR><TD class=label align=right><b>Fecha Descarga Final:</b></TD>
        <TD class=label align=left width=300 colspan=3>
		    <INPUT  name="TripUnloadDate" id="Text3" type=text value="<%=FormatDateTime(TripUnloadDate1)%>" size=23 maxLength=19 class=label readonly>&nbsp;
		    <INPUT type=button value="Seleccionar" onClick="JavaScript:abrir('TripUnloadDate');" class=label>
        </TD></TR>

        <tr><td></td>
		<TD class=label align=left colspan=2>
		
			<select name="TripUnloadDate_h" class=label id="Select1">
                <%=mesdia(0,23,2,-1,"HORA")%>	
			</select>:<select name="TripUnloadDate_i" class=label id="Select2">
                <%=mesdia(0,59,2,-1,"MINUTO")%>	
			</select>
		</TD></TR>
        
                <TR><TD class=label align=right height=30><b> </b></TD><TD class="label " align=left></TD>
            <TD class=label align=right><b></b></TD><TD class="label " align=left></TD></TR>


        <tr>
        <td></td>
        <TD class=label align=right><INPUT name=enviar type=button onClick="JavaScript:validar(2);" value="&nbsp;&nbsp;Actualizar&nbsp;&nbsp;" class=label></TD>
         </tr>

	</TABLE>
	</FORM>
</BODY>
</HTML>
<script>


	selecciona('forma.TripLoadDate_h','<%=TwoDigits(Hour(TripLoadDate))%>');
	selecciona('forma.TripLoadDate_i','<%=TwoDigits(Minute(TripLoadDate))%>');    

    selecciona('forma.TripUnloadDate_h','<%=TwoDigits(Hour(TripUnloadDate))%>');
	selecciona('forma.TripUnloadDate_i','<%=TwoDigits(Minute(TripUnloadDate))%>'); 
      
    selecciona('forma.TripUnloadPartialDate_h','<%=TwoDigits(Hour(TripUnloadPartialDate))%>');
    selecciona('forma.TripUnloadPartialDate_i','<%=TwoDigits(Minute(TripUnloadPartialDate))%>');

<%if AlertSpecialClient <> "" then%>
    alert("<%=AlertSpecialClient %>");
<%end if%>
</script>

