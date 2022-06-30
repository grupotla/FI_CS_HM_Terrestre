<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
Checking "0|1|2"
Dim Conn, rs, i, aTableValues, CountTableValues, HTMLCode, Countries, Week, Yr, TempCountry, BLType, SQLFilter, qry
Dim SubTotalWeight, SubTotalVolume, SubTotalPieces, TotalWeight, TotalVolume, TotalPieces, Typ, ReportType, CtrsTemp

	CountTableValues = -1
	Week = CheckNum(Request("W"))
	Countries  = Request("CT")
	Yr = CheckNum(Request("YR"))
	Typ = CheckNum(Request("T"))
	BLType = CheckNum(Request("BT"))
    ReportType = CheckNum(Request("RT"))
    SQLFilter = Request("filter")

	'Al no indicarse el pais, se selecciona el primer pais asignado al usuario
	if Countries = "" then
		Countries = SetDefaultCountry
	end if
    
    'Select Case Countries
    'Case "GT","GTLTF"
    '    CtrsTemp = "GT','GTLTF"
    'Case "SV","SVLTF"
    '    CtrsTemp = "SV','SVLTF"
    'Case "HN","HN1","HNLTF"
    '    CtrsTemp = "HN','HN1','HNLTF"
    'Case "NI","NILTF"
    '    CtrsTemp = "NI','NILTF"
    'Case "CR","CRLTF"
    '    CtrsTemp = "CR','CRLTF"
    'Case "PA","PALTF"
    '    CtrsTemp = "PA','PALTF"
    'Case "MX","MXLTF"
    '    CtrsTemp = "MX','MXLTF"
    'End Select

    CtrsTemp = Left(Countries,2)
	
	'Si no se indica el anio, se toma el actual
	if Yr = 0 then
		Yr = Year(Now)
	end if

    
	
	OpenConn Conn
	if Typ = 0 then 'Impresion Itinerario
		If ReportType = 0 then 'Reporte por Pais
            SQLFilter = Replace(SQLFilter,"b.","")
            '                               0              1           2        3           4               5         6         7          8            9              10      11       12       13        14      15       16          17              18         19            20              21         22
            'set rs = Conn.Execute("select BLDetailID, BLIDTransit, InTransit, Clients, DischargeDate, DiceContener, Weights, Volumes, NoOfPieces, CountriesFinalDes, Notify, Agents, Shippers, Contact, Container, BLs, CreatedDate, CreatedTime, Observations, Countries, CountriesFinalDes, Incoterms, Coloaders from BLDetail where InTransit in (1,2) and Expired=0 and Countries in ('" & CtrsTemp & "') and Year(CreatedDate)=" & Yr & " and Week = " & Week & " and BLType=" & BLType & " Order by CountriesFinalDes, Pos")
            'set rs = Conn.Execute("select BLDetailID, BLIDTransit, InTransit, Clients, DischargeDate, DiceContener, Weights, Volumes, NoOfPieces, CountriesFinalDes, Notify, Agents, Shippers, Contact, Container, BLs, CreatedDate, CreatedTime, Observations, Countries, CountriesFinalDes, Incoterms, Coloaders from BLDetail where InTransit in (1,2) and Expired=0 and substr(Countries,1,2) = '" & CtrsTemp & "' and Year(CreatedDate)=" & Yr & " and Week = " & Week & " and BLType=" & BLType & " Order by CountriesFinalDes, Pos")
            qry = "select BLDetailID, BLIDTransit, InTransit, Clients, DischargeDate, DiceContener, Weights, Volumes, NoOfPieces, CountriesFinalDes, Notify, Agents, Shippers, Contact, Container, BLs, CreatedDate, CreatedTime, Observations, Countries, CountriesFinalDes, Incoterms, Coloaders from BLDetail where InTransit in (1,2) and Expired=0 and substr(Countries,1,2) = '" & CtrsTemp & "' and Year(CreatedDate)=" & Yr & SQLFilter & " and BLType=" & BLType & " Order by CountriesFinalDes, Pos"            

        Else 'Reporte Regional
            'response.write("select a.BLDetailID, a.BLIDTransit, a.InTransit, a.Clients, a.DischargeDate, a.DiceContener, a.Weights, a.Volumes, a.NoOfPieces, a.CountriesFinalDes, a.Notify, a.Agents, a.Shippers, a.Contact, a.Container, a.BLs, a.CreatedDate, a.CreatedTime, a.Observations, a.Countries, b.CountryDes, a.Incoterms, Coloaders from BLDetail a, BLs b where a.BLID=b.BLID and a.InTransit in (1,2) and a.Expired=0 and Year(a.CreatedDate)=" & Yr & " and a.Week = " & Week & " and a.BLType=" & BLType & " Order by a.CountriesFinalDes, a.Countries, b.CountryDes, a.Pos")
            '                                   0              1           2            3           4                 5             6         7             8               9               10         11         12        13           14        15         16            17              18             19           20            21         22
            'qry = "select a.BLDetailID, a.BLIDTransit, a.InTransit, a.Clients, a.DischargeDate, a.DiceContener, a.Weights, a.Volumes, a.NoOfPieces, a.CountriesFinalDes, a.Notify, a.Agents, a.Shippers, a.Contact, a.Container, a.BLs, a.CreatedDate, a.CreatedTime, a.Observations, a.Countries, b.CountryDes, a.Incoterms, Coloaders from BLDetail a, BLs b where a.BLID=b.BLID and a.InTransit in (1,2) and a.Expired=0 and Year(a.CreatedDate)=" & Yr & " and a.Week = " & Week & " and a.BLType=" & BLType & " Order by a.CountriesFinalDes, a.Countries, b.CountryDes, a.Pos")
            qry = "select a.BLDetailID, a.BLIDTransit, a.InTransit, a.Clients, a.DischargeDate, a.DiceContener, a.Weights, a.Volumes, a.NoOfPieces, a.CountriesFinalDes, a.Notify, a.Agents, a.Shippers, a.Contact, a.Container, a.BLs, a.CreatedDate, a.CreatedTime, a.Observations, a.Countries, b.CountryDes, a.Incoterms, Coloaders from BLDetail a, BLs b where a.BLID=b.BLID and a.InTransit in (1,2) and a.Expired=0 and Year(a.CreatedDate)=" & Yr & SQLFilter & " and a.BLType=" & BLType & " Order by a.CountriesFinalDes, a.Countries, b.CountryDes, a.Pos"
        
        End If
	else 'Impresion datos Adicionales
        SQLFilter = Replace(SQLFilter,"b.","")
		'response.write("select CountriesFinalDes, ChargeType, Clients, Endorse, EndorseType, Declaration, DeclarationType, RequestNo, RequestType, BLsType, BLs, BillType, Bill, PackingListType, PackingList, Observations from BLDetail where InTransit in (1,2) and Expired=0 and Countries='" & Countries & "' and Year(CreatedDate)=" & Yr & " and Week = " & Week & " and BLType=" & BLType & " Order by CountriesFinalDes, Pos")
        'set rs = Conn.Execute("select CountriesFinalDes, ChargeType, Clients, Endorse, EndorseType, Declaration, DeclarationType, RequestNo, RequestType, BLsType, BLs, BillType, Bill, PackingListType, PackingList, Observations from BLDetail where InTransit in (1,2) and Expired=0 and Countries in ('" & CtrsTemp & "') and Year(CreatedDate)=" & Yr & " and Week = " & Week & " and BLType=" & BLType & " Order by CountriesFinalDes, Pos")
        qry = "select CountriesFinalDes, ChargeType, Clients, Endorse, EndorseType, Declaration, DeclarationType, RequestNo, RequestType, BLsType, BLs, BillType, Bill, PackingListType, PackingList, Observations from BLDetail where InTransit in (1,2) and Expired=0 and substr(Countries,1,2) = '" & CtrsTemp & "' and Year(CreatedDate)=" & Yr & SQLFilter & " and BLType=" & BLType & " Order by CountriesFinalDes, Pos"
	
    end if

    'response.write qry & "<br>"
    set rs = Conn.Execute(qry)            

	If Not rs.EOF Then
		aTableValues = rs.GetRows
		CountTableValues = rs.RecordCount-1
	End If
	closeOBJs rs, Conn
    
	if CountTableValues >=0 then
		if Typ = 0 then
			for i=0 to CountTableValues
				if TempCountry <> aTableValues(9,i) then
					'Dibujando los Subtotales del Pais Anterior, ya que se estara abriendo el cuadro para un nuevo pais
					if SubTotalWeight <> 0 then
						HTMLCode = HTMLCode & "<tr><td class=labela colspan=3>&nbsp;</td>" & _
							"<td class=labela align=center><b>" & SubTotalWeight & "</b></td>" & _
							"<td class=labela align=center><b>" & SubTotalVolume & "</b></td>" & _
							"<td class=labela align=center><b>" & SubTotalPieces & "</b></td>" & _
							"<td class=labela colspan=10>&nbsp;</td></tr>"
					end if
					'Abriendo el cuadro para un nuevo pais
					HTMLCode = HTMLCode & "<tr><td colspan=17 align=center bgcolor=#000000></td></tr>" & _
						"<tr><td colspan=17 align=center bgcolor=#000000></td></tr>" & _
						"<tr><td class=style11 colspan=17 align=center><b>" & TranslateCountry(aTableValues(9,i)) & "</b></td></tr>"
					TempCountry = aTableValues(9,i)
					SubTotalWeight = 0
					SubTotalVolume = 0
					SubTotalPieces = 0				
				end if
				HTMLCode = HTMLCode & "<tr>" & _
					"<td class=labela><a href=# class=labela>" & aTableValues(3,i) & "</a></td>" & _
					"<td class=labela><a href=# class=labela>" & aTableValues(4,i) & "</a></td>" & _
					"<td class=labela><a href=# class=labela>" & aTableValues(5,i) & "</a></td>" & _
					"<td class=labela align=center><a href=# class=labela>" & aTableValues(6,i) & "</a></td>" & _
					"<td class=labela align=center><a href=# class=labela>" & aTableValues(7,i) & "</a></td>" & _
					"<td class=labela align=center><a href=# class=labela>" & aTableValues(8,i) & "</a></td>"
				
                if ReportType=1 then 'Reporte Regional
                    HTMLCode = HTMLCode & "<td class=labela align=center><a href=# class=labela><b>" & aTableValues(19,i) & "</b></a></td>" & _
                        "<td class=labela align=center><a href=# class=labela><b>" & aTableValues(20,i) & "</b></a></td>"
                end if

                HTMLCode = HTMLCode & _
                    "<td class=labela align=center><a href=# class=labela>" & aTableValues(9,i) & "</a></td>" & _
					"<td class=labela><a href=# class=labela>" & aTableValues(10,i) & "</a></td>" & _
					"<td class=labela><a href=# class=labela>" & aTableValues(11,i) & "</a></td>" & _
					"<td class=labela><a href=# class=labela>" & aTableValues(12,i) & "&nbsp;</a></td>" & _
					"<td class=labela><a href=# class=labela>" & aTableValues(22,i) & "&nbsp;</a></td>" & _
					"<td class=labela><a href=# class=labela>" & aTableValues(13,i) & "&nbsp;</a></td>" & _
					"<td class=labela><a href=# class=labela>" & aTableValues(14,i) & "&nbsp;</a></td>" & _
					"<td class=labela><a href=# class=labela>" & aTableValues(15,i) & "&nbsp;</a></td>" & _
					"<td class=labela><a href=# class=labela>" & aTableValues(18,i) & "&nbsp;</a></td>" & _
                    "<td class=labela><a href=# class=labela>" & aTableValues(21,i) & "&nbsp;</a></td>" & _
                    "</tr>"
				SubTotalWeight = SubTotalWeight + aTableValues(6,i)
				SubTotalVolume = SubTotalVolume + aTableValues(7,i)
				SubTotalPieces = SubTotalPieces + aTableValues(8,i)
				
				TotalWeight = TotalWeight + aTableValues(6,i)
				TotalVolume = TotalVolume + aTableValues(7,i)
				TotalPieces	= TotalPieces + aTableValues(8,i)
			next
			'Dibujando los Subtotales del ultimo pais y Totales Generales
			HTMLCode = HTMLCode & "<tr><td class=labela colspan=3>&nbsp;</td>" & _
				"<td class=labela align=center><b>" & SubTotalWeight & "</b></td>" & _
				"<td class=labela align=center><b>" & SubTotalVolume & "</b></td>" & _
				"<td class=labela align=center><b>" & SubTotalPieces & "</b></td>" & _
				"<td class=labela colspan=10>&nbsp;</td></tr>" & _
				"<tr><td colspan=17 align=center bgcolor=#000000></td></tr>" & _
				"<tr><td colspan=17 align=center bgcolor=#000000></td></tr>" & _
				"<tr><td class=labela colspan=3 align=right><b>TOTALES</b></td>" & _
				"<td class=labela align=center><b>" & TotalWeight & "</b></td>" & _
				"<td class=labela align=center><b>" & TotalVolume & "</b></td>" & _
				"<td class=labela align=center><b>" & TotalPieces & "</b></td>" & _
				"<td class=labela colspan=10>&nbsp;</td></tr>"
		else 'Si Typ=1, es para mostrar el Itinerario Adicional de documentos
			for i=0 to CountTableValues
				if TempCountry <> aTableValues(0,i) then
					if TempCountry <> "" then
						HTMLCode = HTMLCode & "<tr><td class=labela colspan=17>&nbsp;</td></tr>"
					end if
					'Abriendo el cuadro para un nuevo pais
					HTMLCode = HTMLCode & "<tr><td colspan=17 align=center bgcolor=#000000></td></tr>" & _
						"<tr><td colspan=17 align=center bgcolor=#000000></td></tr>" & _
						"<tr><td class=style11 colspan=17 align=center><b>" & TranslateCountry(aTableValues(0,i)) & "</b></td></tr>"
					TempCountry = aTableValues(0,i)
				end if
				HTMLCode = HTMLCode & "<tr>" & _
					"<td class=labela><a href=# class=labela>" & SetType(aTableValues(1,i),0) & "</a></td>" & _
					"<td class=labela><a href=# class=labela>" & aTableValues(2,i) & "</a></td>" & _
					"<td class=labela><a href=# class=labela>" & SetType(aTableValues(3,i),1) & "</a></td>" & _
					"<td class=labela><a href=# class=labela>" & SetType(aTableValues(4,i),2) & "</a></td>" & _
					"<td class=labela><a href=# class=labela>" & SetType(aTableValues(5,i),1) & "</a></td>" & _
					"<td class=labela><a href=# class=labela>" & SetType(aTableValues(6,i),2) & "</a></td>" & _
					"<td class=labela><a href=# class=labela>" & SetType(aTableValues(7,i),1) & "</a></td>" & _
					"<td class=labela><a href=# class=labela>" & SetType(aTableValues(8,i),2) & "</a></td>" & _
					"<td class=labela><a href=# class=labela>" & SetType(aTableValues(9,i),2) & "</a></td>" & _
					"<td class=labela><a href=# class=labela>" & aTableValues(10,i) & "</a></td>" & _
					"<td class=labela><a href=# class=labela>" & SetType(aTableValues(11,i),2) & "</a></td>" & _
					"<td class=labela><a href=# class=labela>" & aTableValues(12,i) & "</a></td>" & _
					"<td class=labela><a href=# class=labela>" & SetType(aTableValues(13,i),2) & "</a></td>" & _
					"<td class=labela><a href=# class=labela>" & aTableValues(14,i) & "</a></td>" & _
					"<td class=labela><a href=# class=labela>" & aTableValues(15,i) & "</a></td>" & _
					"</tr>"
			next
			'Dibujando los Subtotales del ultimo pais y Totales Generales
			HTMLCode = HTMLCode & "<tr><td class=labela colspan=17>&nbsp;</td></tr>" & _
				"<tr><td colspan=17 align=center bgcolor=#000000></td></tr>" & _
				"<tr><td colspan=17 align=center bgcolor=#000000></td></tr>" & _
				"<tr><td class=labela colspan=17>&nbsp;</td></tr>"
		end if
	end if
%>
<HTML>
<HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/config.js"></SCRIPT>
<SCRIPT src="img/mainLib.js" language="JavaScript"></SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<style type="text/css">
<!--
body {
	margin: 0px;
}
.style3 {
	font-size:10px;
	color: #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
}
.style4 {
	font-size:10px;
	color: #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
	border-bottom-style:solid;
	border-left-style:solid;
	border-right-style:solid;
	border-top-style:solid;
	border-collapse:collapse;
	border-width: 1px;
}
.style10 {font-size: 10px; font-family: Verdana, Arial, Helvetica, sans-serif; font-weight:normal;}
.style11 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 11px;
}	
.styleborder {
	border-bottom-style:solid;
	border-left-style:solid;
	border-right-style:solid;
	border-top-style:solid;
	border-width: 1px;
	border-collapse:collapse;
}
-->
</style>
<%
        Dim iResult, iLogo, iEdicion, iTitulo, iEmpresa, iDireccion, iObservaciones, iPlantilla    
        iResult = WsGetLogo(Countries, "TERRESTRE",  IIf(Typ=0,"18","19"),  "",  "")
        iLogo = iResult(20)
        iEdicion = iResult(2)
        iTitulo = iResult(3)
        iEmpresa = iResult(4)
        iDireccion = iResult(6)
        iObservaciones = iResult(1)
        iPlantilla = iResult(22)

    'dim aTableValues5
    'aTableValues5 = EmpresaParametros(Countries,  IIf(Typ=0,"18","19"), "TERRESTRE")
    'if aTableValues5(1,0) <> "" then
    '    iLogo = aTableValues5(20,0)
    '    iEdicion = aTableValues5(3,0)
    '    iTitulo = aTableValues5(4,0)
    '    iEmpresa = aTableValues5(5,0)
    '    iDireccion = aTableValues5(7,0)
    '    iObservaciones = aTableValues5(11,0)
    '    iPlantilla = aTableValues5(21,0)
    'end if    
 %>
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="JavaScript:self.focus();">
	<%if Typ=0 then%>
	<TABLE cellspacing=0 cellpadding=2 width=100% align=center border="0">
		<TR>
			<TD class=label align=left>
				<%=DisplayLogo(Countries, 0, 0, 0, iLogo)%>
			</TD>
			<TD class=label align=center>
				<b><%=IIf(iPlantilla = "", "STATUS DE CARGA SEMANA", iPlantilla)%>&nbsp;
                <%
                if ReportType=0 then 'Reporte Regional
                    response.write CtrsTemp & "-" & SetType(BLType,3) & "-" & Week  & "-" & Yr
                else
                    response.write Week  & "-" & Yr
                end if
                %>                
                </b><BR>				
                <%=IIf(iTitulo = "", "ITINERARIO DE CARGA CENTROAMERICANA", iTitulo)%>
			</TD>
			<TD class=label align=left>
				<b><%=IIf(iEdicion = "", "EDICION 1", iEdicion)%></b>
			</TD>
		</TR> 		
	</TABLE>
	<TABLE cellspacing=0 cellpadding=2 width=998 align=center class="styleborder" border="1">
		<TR>
			<TD class=labela align=left>&nbsp;<b>CLIENTE</b>&nbsp;</TD>
            <TD class=labela align=left>&nbsp;<b>FECHA&nbsp;DESCARGA</b>&nbsp;</TD>
            <TD class=labela align=left>&nbsp;<b>DESCRIPCION&nbsp;DE&nbsp;CARGA</b>&nbsp;</TD>
            <TD class=labela align=left>&nbsp;<b>PESO</b>&nbsp;</TD>
            <TD class=labela align=left>&nbsp;<b>CBM</b>&nbsp;</TD>
            <TD class=labela align=left>&nbsp;<b>BULTOS</b>&nbsp;</TD>
            <%if ReportType=1 then 'Reporte Regional%>
               <TD class=labela align=center>&nbsp;<b>PAIS TRANSITO SALIDA</b>&nbsp;</TD>
               <TD class=labela align=center>&nbsp;<b>PAIS TRANSITO DESTINO</b>&nbsp;</TD>
            <%end if%>
            <TD class=labela align=center>&nbsp;<b>PAIS DESTINO FINAL</b>&nbsp;</TD>
            <TD class=labela align=left>&nbsp;<b>CONTACTO</b>&nbsp;</TD>
            <TD class=labela align=left>&nbsp;<b>EXPORTADOR</b>&nbsp;</TD>
            <TD class=labela align=left>&nbsp;<b>AGENTE</b>&nbsp;</TD>
			<TD class=labela align=left>&nbsp;<b>COLOADER</b>&nbsp;</TD>
			<TD class=labela align=left>&nbsp;<b>OPERADOR</b>&nbsp;</TD>
			<TD class=labela align=left>&nbsp;<b>CONTENEDOR</b>&nbsp;</TD>
			<TD class=labela align=left>&nbsp;<b>BL/RO</b>&nbsp;</TD>
            <TD class=labela align=left>&nbsp;<b>OBSERVACIONES</b>&nbsp;</TD>
            <TD class=labela align=left>&nbsp;<b>INCOTERMS</b>&nbsp;</TD>
		</TR> 
		<%=HTMLCode%>
	</TABLE>
	<%else%>
	<TABLE cellspacing=0 cellpadding=2 width=998 align=center border="0">
		<TR>
			<TD class=label align=left>
                <%=DisplayLogo(Countries, 0, 0, 0, iLogo)%>
			</TD>
			<TD class=label align=left>
				<b><%=IIf(iPlantilla = "", "STATUS DE CARGA SEMANA", iPlantilla)%> &nbsp; <%=CtrsTemp & "-" & SetType(BLType,3) & "-" & Week  & "-" & Yr%></b><BR>				
                <%=IIf(iTitulo = "", "ENTREGA DE DOCUMENTOS RUTA CENTROAMERICANA", iTitulo)%>
			</TD>
			<TD class=label align=left>
                <b><%=IIf(iEdicion = "", "EDICION 1", iEdicion)%></b>
			</TD>
		</TR> 		
	</TABLE>
	<TABLE cellspacing=0 cellpadding=2 width=998 align=center class="styleborder" border="1">
		<TR>
			<TD class=labela align=center>&nbsp;<b>TIPO&nbsp;DE<br>CARGA</b>&nbsp;</TD><TD class=labela align=center>&nbsp;<b>CLIENTE</b>&nbsp;</TD><TD class=labela colspan=2 align=center>&nbsp;<b>ENDOSO<BR>ADUANAL-RO</b>&nbsp;</TD><TD class=labela colspan=2 align=center>&nbsp;<b>DECLARACION<br>DE&nbsp;ADUANA</b>&nbsp;</TD><TD class=labela colspan=2 align=center>&nbsp;<b>REQUERIMIENTO<br>DE&nbsp;PARTIDAS</b>&nbsp;</TD><TD class=labela colspan=2 align=center>&nbsp;<b>BL</b>&nbsp;</TD><TD class=labela colspan=2 align=center>&nbsp;<b>FACTURA</b>&nbsp;</TD><TD class=labela colspan=2 align=center>&nbsp;<b>LISTA<br>DE&nbsp;EMPAQUE</b>&nbsp;</TD><TD class=labela align=center>&nbsp;<b>OBSERVACIONES</b>&nbsp;</TD>
		</TR> 
		<%=HTMLCode%>
	</TABLE>
	<%end if%>
</BODY>
</HTML>
<%Set aTableValues = Nothing%>