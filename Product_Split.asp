<%
Checking "0|1|2"

Dim InTransit2

BLID = CheckNum(Request("BLID"))
BL = Request("BL")
OID = CheckNum(Request("OID"))

id_json = CheckNum(Request("js"))
Action2 = CheckNum(Request("Action2"))
GroupID = Request("GID")
 

'if CheckNum(id_json) > 0 then

CommodityCode = Request("CommoditiesID") 
TotNoOfPieces = Request("NoOfPieces")
Contener = Request("Contener")
Commodity = Request("DiceContener")
Volume = Request("Volumes")
Weight = Request("Weights")
Countries = Request("CountrySession")
CountryOrigen = Request("CountryOrigen")
FinalDes = Request("CountriesFinalDes")
InTransit2 = 1 'Request("InTransit")

'end if

CountList1Values = -1
CountList2Values = -1
CountList3Values = -1

if CountTableValues >= 0 then

	OID = aTableValues(0, 0)
    CreatedDate = aTableValues(1, 0)
    CreatedTime = aTableValues(2, 0)
    'BL = aTableValues(53, 0)
    'MBL = aTableValues(20, 0)
    BLID = CheckNum(aTableValues(52, 0))
    ConsignerID = aTableValues(6, 0)   
    ConsignerData = aTableValues(8, 0)

	Container = aTableValues(19, 0)
	MBL = aTableValues(20, 0)
	BL = aTableValues(10, 0) 


CreatedDate = ConvertDate(aTableValues(1, 0),2)
WareHouseDischargeDate = aTableValues(5, 0)
ClientID = aTableValues(6, 0)     
AddressID = aTableValues(7, 0)
Name = aTableValues(8, 0)
'CommodityCode = aTableValues(9, 0)
'Commodity = aTableValues(10, 0)
'Weight = aTableValues(11, 0)
'Volume = aTableValues(12, 0)
'TotNoOfPieces = aTableValues(13, 0)
'FinalDes = aTableValues(14, 0)
ShipperID = aTableValues(15, 0)
ShipperAddrID = aTableValues(16, 0)
ShipperData = aTableValues(17, 0)
ContactSignature = aTableValues(18, 0)

BLID = aTableValues(52, 0)
'BLNumber = replace(Trim(aTableValues(53, 0)),"--","",1,-1)
BLNumber = aTableValues(53, 0)

end if

 

    On Error Resume Next      

        if Action2 >= 1 and Action2 <= 4 then


            'response.write "(" & id_json & ")(" & OID & ")"

            QuerySelect = ""
            Select Case Action2
	        Case 1
                QuerySelect = PutSplits(      0, OID, TotNoOfPieces, Contener, CommodityCode, Commodity, Volume, Weight, DischargeDate, Countries, CountryOrigen, FinalDes, InTransit2, 0) 

	        Case 2
                QuerySelect = PutSplits(id_json, OID, TotNoOfPieces, Contener, CommodityCode, Commodity, Volume, Weight, DischargeDate, Countries, CountryOrigen, FinalDes, InTransit2, 0) 
	        
            Case 3
                QuerySelect = PutSplits(id_json, OID, TotNoOfPieces, Contener, CommodityCode, Commodity, Volume, Weight, DischargeDate, Countries, CountryOrigen, FinalDes, InTransit2, 1) 

            End Select

            id_json = 0  

CommodityCode = ""
TotNoOfPieces = ""
Contener = ""
Commodity = ""
Volume = ""
Weight = ""
Countries = ""
CountryOrigen = ""
FinalDes = ""
InTransit2 = -1

            
        end if

    If Err.Number<>0 then
	    response.write "Msg :" & Err.Number & " - " & Err.Description & "<br>"  
    end if

    '//////////////////////////////////ADUANAS
'    QuerySelect = "SELECT id_geocerca, geocerca FROM GeoCercas ORDER BY geocerca"
'    Set rs = Conn.Execute(QuerySelect)
'	if Not rs.EOF then
   '		aList3Values = rs.GetRows
    '   	CountList3Values = rs.RecordCount-1	        		
	'end if
    'CloseOBJ rs	

      
'dim filter2

        '//////////////////////////////////registros guardados split
        On Error Resume Next                 
            filter2 = " AND BLID=" & OID                    
            aList1Values = GetSplits(filter2)
            CountList1Values = Ubound(aList1Values, 2)
        If Err.Number<>0 then
	        'response.write "disatel_json lists :" & Err.Number & " - " & Err.Description & "<br>"  
        end if
        
        '//////////////////////////////////registro actual split      
        On Error Resume Next      
            filter2 = "AND BLDetailID=" & id_json
            aList2Values = GetSplits(filter2)
            CountList2Values = Ubound(aList2Values, 2)
        If Err.Number<>0 then
	        'response.write "disatel_json :" & Err.Number & " - " & Err.Description & "<br>"  
        end if
                 	

    if CountList2Values >= 0 then

        TotNoOfPieces = aList2Values(2, 0)
        Contener = aList2Values(3, 0)
        CommodityCode = aList2Values(4, 0)
        Commodity = aList2Values(5, 0)
        Volume = aList2Values(6, 0)
        Weight = aList2Values(7, 0)
        Countries = aList2Values(9, 0)
        CountryOrigen = aList2Values(10, 0)
        FinalDes = aList2Values(11, 0)
        InTransit2 = aList2Values(12, 0)

    end if


    OpenConn2 Conn

    	'Obteniendo listado de tipos de paquete
        Set rsFilter = Conn.Execute("select distinct tipo from tipo_paquete order by tipo")
		if Not rsFilter.EOF then
            aTableValues = rsFilter.GetRows
        	CountList2Values = rsFilter.RecordCount -1
		end if
        CloseOBJ rsFilter

    CloseOBJs rsFilter, Conn


    '//////////////////////////////////PLACA
    'CountList2Values = -1
    'QuerySelect = "SELECT a.BLNumber, b.Countries, b.TruckNo, b.Mark, b.Model, b.Motor FROM BLs a, Trucks b WHERE a.TruckID = b.TruckID and BLID=" & BLID
    ''response.write QuerySelect & "<br>"
    'Set rs = Conn.Execute(QuerySelect)
	'if Not rs.EOF then
   	'	aList2Values = rs.GetRows
    '   	CountList2Values = rs.RecordCount-1	        		
	'end if
    'CloseOBJ rs	

    'if  CountList2Values >= 0 then
    '    placa = Replace(aList2Values(2, 0), "-", "")
    '    plate = aList2Values(2, 0)	
    'end if
    'CloseOBJ Conn


    'fleet = GroupData(ConsignerID)


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
                           
                if (!valTxt(document.forma.DiceContener, 3, 5)){return (false)};
                if (!valSelec(document.forma.Contener, 3, 5)){return (false)};
			    if (!valTxt(document.forma.Weights, 1, 10)){return (false)};
			    if (!valTxt(document.forma.Volumes, 1, 10)){return (false)};
			    if (!valTxt(document.forma.NoOfPieces, 1, 10)){return (false)};

                //if (!valSelec(document.forma.CountrySession)){return (false)}; 
			    if (!valSelec(document.forma.CountryOrigen)){return (false)};
			    if (!valSelec(document.forma.CountriesFinalDes)){return (false)};
                //if (!valSelec(document.forma.InTransit)){return (false)};
 
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

	function GetData(GID,DiceContenerTemp){
		window.open('Search_BLData.asp?GID='+GID+'&DiceContenerTemp='+DiceContenerTemp,'BLData','height=400,width=460,menubar=0,resizable=1,scrollbars=1,toolbar=0,status=1');
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

    <INPUT name="js" type=hidden value="<%=id_json%>">
    <INPUT name="BLID" type=hidden value="<%=BLID%>">	
    <INPUT name="BL" type=hidden value="<%=BL%>">
    <INPUT name="Container" type=hidden value="<%=Container%>">
    <INPUT name="MBL" type=hidden value="<%=MBL%>">
  
    <INPUT name="BLNumber" type=hidden value="<%=BLNumber%>">
    <INPUT name="ConsignerID" type=hidden value="<%=ConsignerID%>">
    <INPUT name="ConsignerData" type=hidden value="<%=ConsignerData%>">
    <INPUT name="CountrySession" type=hidden value="<%=Session("OperatorCountry")%>">
    
    

    <TABLE cellspacing=3 cellpadding=2 width=650 align=center border=0>
 


        <TR><TD class=label align=right><b>BLDetailID:</b></TD><TD class="label label1" align=left><%if OID <> 0 then response.write OID End if%></TD>
            <TD class=label align=right><b>Fecha:</b></TD><TD class="label label1" align=left><%=ConvertDate(CreatedDate,1)%></TD></TR>               
        <TR><TD class=label align=right><b>BLID:</b></TD><TD class="label label1" align=left><%=BLID%></TD> 
            <TD class=label align=right><b>CP Hija:</b></TD><TD class="label label1" align=left><%=IIf(BLNumber = "--","<font color=red>SIN NUMERO DE CARTA PORTE</font>",BLNumber)%></TD></TR> 
		<TR><TD class=label align=right><b>MBL:</b></TD><TD class="label label1" align=left><%=MBL%></TD> 		    
            <TD class=label align=right><b>Cliente:</b></TD><TD class="label label1" align=left><%=ConsignerID & " - " & ConsignerData%></TD></TR>



		<TR><TD class=label align=right><b>Contenedor:</b></TD><TD class="label label1" align=left><%=Container%></TD>
		    <TD class=label align=right><b>Descripcion de la Carga:</b></TD><TD class="label label1" align=left><%=BL%></TD></TR> 


<!--     
 		<TR><TD class=label align=right><b>Pais Session:</b></TD><TD class="label label1" align=left><%=Countries%></TD> 		    
            <TD class=label align=right><b></b></TD><TD class="label label1" align=left><%=%></TD></TR>
                    
		<TR><TD class=label align=right><b>Origen:</b></TD><TD class="label label1" align=left><%=CountryOrigen%></TD> 		    
            <TD class=label align=right><b>Destino Final:</b></TD><TD class="label label1" align=left><%=FinalDes%></TD></TR>

		<TR><TD class=label align=right><b>Fecha Descarga:</b></TD><TD class="label label1" align=left><%=ConvertDate(WareHouseDischargeDate,1)%></TD> 
        		    
            <TD class=label align=right><b>Commodity:</b></TD><TD class="label label1" align=left><%=CommodityCode & " - " & Commodity%></TD></TR>
-->	 
		<TR><TD class=label align=right height=20><b>&nbsp;</b></TD><TD class=label align=left></TD></TR> 


		<TR><TD class=label align=right height=20 colspan=3><b><h1>DESGLOSE</h1></b></TD><TD class=label align=left></TD></TR> 


        <TR><TD class=label align=right valign=top><b>Producto:</b></TD>		
		<TD class=label align=left class="style4" colspan=3>
        	<INPUT TYPE=text class=label name="DiceContener" id="Producto" value="<%=Commodity%>" maxlength="200" size="35" readonly><%="<b> ID: " & CommodityCode & "<b>"%>
            <a href="#" onClick="Javascript:IData(9);return (false);" class="submenu" target=_blank><font color="FFFFFF"><b>Nuevo</b></font></a>
			<a href="#" onClick="Javascript:GetData(9,'');return (false);" class="menu"><font color="FFFFFF"><b>Buscar</b></font></a>
		    <INPUT name="CommoditiesID" type=hidden value="<%=CommodityCode%>" >
		</TD></TR> 
        <TR><TD class=label align=right><b>Tipo Paquete:</b></TD><TD class=label align=left>
            <select name="Contener" id="Tipo Paquete" class="label">
					<option value="-1">Seleccionar</option>
                    <%for i=0 to CountList2Values%>
                     <option value="<%=aTableValues(0,i) %>"><%=aTableValues(0,i)%></option>
                    <%next 
                    Set aTableValues = Nothing
                    %>
			</select>
        </TD> 
		<TD class=label align=right><b>Peso:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Weights" id="Peso" value="<%=Weight%>" maxlength="200" size="15" onKeyUp="res(this,numb);"></TD></TR> 
		
        <TR><TD class=label align=right><b>Volumen:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="Volumes" id="Volumen" value="<%=Volume%>" maxlength="200" size="15" onKeyUp="res(this,numb);"></TD>
		    <TD class=label align=right><b>Bultos:</b></TD><TD class=label align=left><INPUT TYPE=text class=label name="NoOfPieces" id="Bultos" value="<%=TotNoOfPieces%>" maxlength="200" size="15" onKeyUp="res(this,numb);"></TD></TR> 
        
        <!--
        <TR><TD class=label align=right><b>En transito:</b></TD><TD class=label align=left >
			<select name="InTransit" id="En transito" class="label">
				<option value="-1">Seleccionar</option>
				<option value="0">No</option>
				<option value="1">Si</option>                        
			</select>
			</TD>
            -->

		<TR><TD class=label align=right><b>Origen:</b></TD><TD class=label align=left colspan=3>
            <select name="CountryOrigen" id="Pais Origen" class="label">
				<option value="-1">Seleccionar</option>
                <!--#include file=Countries.asp-->
			</select>
		    </TD></TR>
        <TR><TD class=label align=right><b>Destino Final:</b></TD><TD class=label align=left colspan=3>
			<select name="CountriesFinalDes" id="Pais Destino Final" class="label">
				<option value="-1">Seleccionar</option>
                <!--#include file=Countries.asp-->
			</select>
			</TD></TR>





        <TR>
        <td></td>
        <TD class=label align=center>
                <% if id_json > 0 then %>
                <INPUT name=enviar type=button onClick="JavaScript:validar(2);" value="&nbsp;&nbsp;Actualizar&nbsp;&nbsp;" class=label>
                </TD>
                <TD align=left>
                <INPUT name=enviar type=button onClick="JavaScript:validar(3);" value="&nbsp;&nbsp;Borrar&nbsp;&nbsp;" class=label>
                <% else %>
                <INPUT name=enviar type=button onClick="JavaScript:validar(1);" value="&nbsp;&nbsp;Grabar&nbsp;&nbsp;" class=label>
                <% end if %>
        </TD>
        <TD align=left>
            <INPUT name=enviar type=button onClick="JavaScript:validar(4);" value="&nbsp;&nbsp;Cancelar&nbsp;&nbsp;" class=label>
  
        </TR>

     

	</TABLE>

	            <TABLE cellspacing=0 cellpadding=2 width=100% align=center>
		            <TR>
		            <TD colspan=2 class=label align=right valign=top>
				            <TABLE cellspacing=5 cellpadding=2 width=100% align=center>
							        
<tr>
<td class="titlelist"><b></b></td><td class="titlelist"><b>#</b></td>
<td class="titlelist"><b>No. Bultos</b></td><td class="titlelist"><b>Clase Bultos</b></td><td class="titlelist"><b>Descripción de Carga</b></td>
<td class="titlelist"><b>Volumen (CBM)</b></td><td class="titlelist"><b>Peso Bruto(Kg)</b></td><td class="titlelist"><b>Origen</b></td>
<td class="titlelist"><b>Procedencia</b></td><td class="titlelist"><b>Destino</b></td><td class="titlelist"><b>Transito</b></td>
</tr>
             
                    <%for i=0 to CountList1Values%>
<% 
'dim onclick : onclick = ""
'if CheckNum(aList1Values(20,i)) > 0 and aList1Values(17,i)  <> "" then
    'onclick = "<a href='" & "InsertData.asp?OID=" & OID & "&GID=36&js=" & aList1Values(0,i) & "'><img src='img/edit.png'></a>" 

    onclick = "<a href='" & "InsertData.asp?GID=36&OID=" & OID & "&CD=" & CreatedDate & "&CT=" & CreatedTime & "&js=" & aList1Values(0,i) & "'><img src='img/edit.png'></a>" 

     
'else
'    onclick = ""
'end if
%>
                        <tr>
                        <td class="listwarning labellist" align=center width=14><%=onclick%></td>
                        <td class="listwarning labellist"><%=aList1Values(0,i) %></td>                       
                        <td class="listwarning labellist"><%=aList1Values(2, i) %></td>
                        <td class="listwarning labellist"><%=aList1Values(3,i) %></td>
                        <td class="listwarning labellist"><%=aList1Values(5,i) %></td>
                        <td class="listwarning labellist"><%=aList1Values(6,i) %></td>
                        <td class="listwarning labellist"><%=aList1Values(7,i) %></td>
                        <td class="listwarning labellist"><%=aList1Values(9,i) %></td>
                        <td class="listwarning labellist"><%=aList1Values(10,i) %></td>
                        <td class="listwarning labellist"><%=aList1Values(11,i) %></td>
                        <td class="listwarning labellist"><%=aList1Values(12,i) %></td>
                        </tr>
                    <%next 
                    Set aList1Values = Nothing
                    %>

				            </TABLE>
                            
		            </TD>
	              </TR>
                  </TABLE>

	</FORM>
</BODY>
</HTML>
<script type="text/javascript">
    selecciona('forma.Contener', '<%=Contener%>');
    //selecciona('forma.CountrySession', '<%=Countries%>'); 
    selecciona('forma.CountryOrigen', '<%=CountryOrigen%>');
    selecciona('forma.CountriesFinalDes', '<%=FinalDes%>');
    //selecciona('forma.InTransit', '<%=InTransit2%>');        
</script>

