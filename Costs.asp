<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
Checking "0|1|2"
Dim ObjectID, Conn, Conn2, rs, Action, Currencies, Name, Volume, Weight, Agent, HBLNumber, BL, i, j, k, CantItems, Countries, BAWResult, TipoConta, QuerySelect, ConnMaster, FreezeCosts, excel_str
Dim aTableValues, CountTableValues, aTableValues2, CountTableValues2, aTableValues3, CountTableValues3, aTableValues4, CountTableValues4, BLType, TC, Msg, OP, ProvID, Servicio 

ObjectID = CheckNum(Request("OID"))
Action = CheckNum(Request("Action"))
BAWResult = 0
Msg = ""
k=0
CountTableValues = -1
CountTableValues2 = -1
CountTableValues3 = -1
CantItems = 24
FreezeCosts = 0

OpenConn Conn
    
    if Action=3 then 'CIERRA COSTOS
        'QuerySelect = "UPDATE BLs SET FreezeCosts = 1 WHERE BLID = " & ObjectID
        'response.write QuerySelect & "<br>"
        'Conn.Execute(QuerySelect)
    end if

    if Action=4 then 'ABRE COSTOS
        'QuerySelect = "UPDATE BLs SET FreezeCosts = 0 WHERE BLID = " & ObjectID
        'response.write QuerySelect & "<br>"
        'Conn.Execute(QuerySelect)
    end if

    Dim CountryExactus, Movimiento

    Movimiento = "-"

    CountryExactus = Session("OperatorCountry")

QuerySelect = "select DISTINCT 'EXPORT' as tipo from BLDetail INNER JOIN BLs ON BLDetail.BLID = BLs.BLID where BLDetail.BLID = " & ObjectID & " and BLDetail.Countries in ('" & CountryExactus & "') and CountriesFinalDes not in ('" & CountryExactus & "') and BLDetail.BLType in (0,1) " & _
"union select DISTINCT 'IMPORT' as tipo from BLDetail INNER JOIN BLs ON BLDetail.BLID = BLs.BLID where BLDetail.BLID = " & ObjectID & " and Intransit IN (2) and CountriesFinalDes in ('" & CountryExactus & "') and BLDetail.BLType in (0,1) " & _
"union select DISTINCT 'IMPORT' as tipo from BLDetail INNER JOIN BLs ON BLDetail.BLID = BLs.BLID where BLDetail.BLID = " & ObjectID & " and BLDetail.Countries in ('" & CountryExactus & "') and CountriesFinalDes in ('" & CountryExactus & "') and BLDetail.BLType in (2) "

    'response.write QuerySelect & "<br>"
	Set rs = Conn.Execute(QuerySelect)
	if Not rs.EOF then
		Movimiento = rs(0)
	end if


    QuerySelect = "SELECT ConsignerData, TotVolume, TotWeight, ShipperData, BLNumber, BLType, Countries, FreezeCosts, CASE WHEN BLType = 0 THEN 'LL' WHEN BLType = 1 THEN 'LE' ELSE 'LC' END FROM BLs WHERE BLID=" & ObjectID
    'response.write QuerySelect & "<br>"
	Set rs = Conn.Execute(QuerySelect)
	if Not rs.EOF then
		Name = Split(rs(0),chr(13)&chr(10))
		Volume = rs(1)
		Weight = rs(2)
		Agent = Split(rs(3),chr(13)&chr(10))
		HBLNumber = rs(4)
		BLType = rs(5)
		Countries = rs(6)
        FreezeCosts = 0 'CheckNum(rs(7))
        Servicio = rs(8)
	end if
	CloseOBJ rs
    


    OP = BLType
	'Verificando que exista Tipo Cambio en el pais donde se va provisionar, caso contrario no permitira guardar los costos
	if Session("OperatorCountry")<>"" then
		OpenConnBAW Conn2
			Set rs = Conn2.Execute("select tca_tcambio from tbl_tipo_cambio where tca_fecha='" & Year(Now) & "-" & Month(Now) & "-" & Day (Now) & "' and tca_pai_id=" & SetCountryBAW(Session("OperatorCountry")))
            if Not rs.EOF then
				TC = CheckNum(rs(0))
			end if
		CloseOBJs rs, Conn2
	end if


    TC = 1 'temporal 


OpenConn2 ConnMaster
	'Obteniendo Monedas
	'Set rs = Conn.Execute("select distinct simbolo from monedas where pais in " & Session("Countries") & " order by simbolo")
    Set rs = ConnMaster.Execute("select distinct simbolo from monedas order by simbolo")
	Do While Not rs.EOF
		Currencies = Currencies & "<option value=" & rs(0) & ">" & rs(0) & "</option>"
		rs.MoveNext
	Loop


        '////////////// PARAMETROS DE LA EMPRESA A FACTURAR
        TipoConta = "BAW"

        if Request("TipoConta") <> "" then

            TipoConta = Request("TipoConta")
        else
            QuerySelect = "SELECT COALESCE(a.tipo_conta,'BAW') as tipo_conta FROM empresas_parametros a WHERE a.country = '" & Countries & "' " 
            'response.write QuerySelect & "<br>"
            Set rs = ConnMaster.Execute(QuerySelect)
	        if Not rs.EOF then                       
                TipoConta = rs(0)
            end if

        end if

         if TipoConta = "" then
            TipoConta = "BAW" 
         end if


        if TipoConta = "BAW" then
            response.write "<font family=verdana color=navy>Pais " & Countries & " tiene Contabilidad BAW</font><br>" 
        end if

	'Almacenando los Costos en Terrestre y BAW
    if Action=1 or Action=2 then

   

		    SaveCostItems Conn, ObjectID, Action
		    'Llamando al BAW para que genere las provisiones correspondientes a terrestre, SysID=3
		    if Session("OperatorCountry")<>"" then
			    'Conn.Execute("update Costs set ProvisionID=1 where BLID=" & ObjectID)
			    'Ultimo parametro es 0 en la funcion significa ir a provisionar costos
			    'response.write("ObjectID - " & ObjectID & " 3 - " & 3 & " OP - " & OP & " SessionOperator - " & Session("OperatorCountry") & " SessionLogin - " & Session("Login") & " SessionLogin - " & 0 & "<br>")
            
                if TipoConta = "BAW" then
                    BAWResult = SetBAWProvition (ObjectID, 3, OP, Session("OperatorCountry"), Session("Login"), 0)   
                end if
		
            end if

   
	end if




    if Action = 5 then    


        Dim DatosCorruptos, codigo_erp, moneda_erp, hay_costos, TipoIva, valor, base, impuesto, pais_id, costo_id 'tabla_exactus, categoria_exactus, 
		
        hay_costos = 0
        CountTableValues4 = -1
        DatosCorruptos = -1

QuerySelect = "SELECT z.SupplierType, z.SupplierID, z.SupplierName, z.Reference, z.Currency, z.ReferenceDate, z.TipoDocID, z.TipoDocName, z.SubTipoID, z.SubTipoName, CASE WHEN z.IsAffected = 1 THEN 'AFECTO' ELSE 'NO AFECTO' END FROM Costs z INNER JOIN ( " & _ 			
"SELECT * FROM ( " & _
"SELECT COUNT(*) as N, SupplierType, SupplierID, COALESCE(SupplierName,'') FROM ( " & _
"SELECT SupplierType, SupplierID, SupplierName, Reference, Currency, ReferenceDate, TipoDocID, SubTipoID, IsAffected " & _
"FROM Costs WHERE BLID = " & ObjectID & " AND Expired = 0 AND ProvisionID = 0 AND cxp_exactus_id = 0 GROUP BY SupplierType, SupplierID, SupplierName, Reference, Currency, ReferenceDate, TipoDocID, SubTipoID, IsAffected " & _
") x GROUP BY SupplierType, SupplierID, COALESCE(SupplierName,'') " & _
") y WHERE  N > 1 " & _
") b ON z.SupplierType = b.SupplierType AND z.SupplierID = b.SupplierID " & _
"WHERE z.BLID = " & ObjectID & " AND z.Expired = 0 AND z.cxp_exactus_id = 0 LIMIT 2 "

        'response.write(QuerySelect & "<br>")
	    Set rs = Conn.Execute(QuerySelect)
	    if Not rs.EOF then
            aTableValues4 = rs.GetRows
            CountTableValues4 = rs.RecordCount-1
	    end if


        Dim strA, strB, strC
        if CountTableValues4 > -1 then 

            msg = "Hay diferencias en uno o varios de los siguientes datos :"
            msg = msg & "<table border=0 cellpadding=3><tr bgcolor=silver><th class='style4'>SupplierID</th><th class='style4'>SupplierName</th><th class='style4'>Reference</th><th class='style4'>Currency</th><th class='style4'>ReferenceDate</th><th class='style4'>TipoDocID</th><th class='style4'>TipoDocName</th><th class='style4'>SubTipoID</th><th class='style4'>SubTipoName</th><th class='style4'>IsAffected</th></tr>"

	  	    for i=0 to CountTableValues4
           
                msg = msg & "<tr>"
            
                for j=1 to 10    
                    
                    strA = ""
                    strB = ""
                    strC = " style='font-weight:normal;background-color:"

                    On Error Resume Next
                        strA = aTableValues4(j,0)
                        strB = aTableValues4(j,1)
                    If Err.Number<>0 then                                '
	                    response.write Err.Number & " - " & Err.Description & "<br>"
                    end if

                    if strA <> strB then
                        strC = strC & "lightblue' "
                    else
                        strC = strC & "white' "
                    end if

                    msg = msg & "<td class=style4 " & strC & ">" & aTableValues4(j,i) & "</td>"
                next

                msg = msg & "</tr>"

            next 
        
            msg = msg & "</table>"
                               

        else



            CountTableValues4 = -1
                '                       0                       1                           2                          3                      4                        5                               6                                                7                         8                         9                       10          11                      12                                  13                          14
            QuerySelect = "SELECT COALESCE(SupplierType,0), COALESCE(SupplierID,0), COALESCE(SupplierName,''), COALESCE(Currency,''), COALESCE(SUM(Cost),0), COALESCE(Reference,''), COALESCE(CAST(ReferenceDate as char),CURRENT_DATE), COALESCE(TipoDocID,''), COALESCE(TipoDocName,''), COALESCE(SubTipoID,''), COALESCE(SubTipoName,''), IsAffected, GROUP_CONCAT(CostID SEPARATOR ','), GROUP_CONCAT(ItemName SEPARATOR '; '), COUNT(*) " & _   
            "FROM Costs " & _
            "WHERE BLID = '" & ObjectID & "' AND Expired = 0 AND ProvisionID = 0 AND cxp_exactus_id = 0 " & _
            "GROUP BY SupplierType, SupplierID, SupplierName, Currency, Reference, ReferenceDate, TipoDocID, TipoDocName, SubTipoID, SubTipoName, IsAffected " 
            'response.write(QuerySelect & "<br>")
	        Set rs = Conn.Execute(QuerySelect)
	        if Not rs.EOF then
		        aTableValues4 = rs.GetRows
		        CountTableValues4 = rs.RecordCount-1
	        end if

                
            if CountTableValues4 = -1 then 
                  
                msg = "No hay registros pendientes para enviar el email"

            else

                pais_id = 41

                QuerySelect = "SELECT id, iva+1 FROM paises_iva INNER JOIN empresas_parametros ON country = '" & Session("OperatorCountry") & "' WHERE codigo = '" & Left(Countries,2) & "'" 
                'response.write QuerySelect & "<br>"
                Set rs = ConnMaster.Execute(QuerySelect)
	            if Not rs.EOF then                  
                    pais_id = rs(0)
                    TipoIva = FormatNumber(CDbl(rs(1)), 2)
                end if


                
	  	        for i=0 to CountTableValues4

			        codigo_erp = "-1"
                    'tabla_exactus = "-1" 
                    'categoria_exactus = "-1"

                    'aTableValues4(0,i) = "0" or 
                    if aTableValues4(1,i) = "0" or aTableValues4(3,i) = "" or aTableValues4(4,i) = "0" or aTableValues4(5,i) = "" or aTableValues4(6,i) = "" or aTableValues4(6,i) = "0000-00-00" or aTableValues4(6,i) = "" or aTableValues4(7,i) = "" or aTableValues4(9,i) = "" then  
                                            
                        msg = "Uno o varios de los siguientes datos no tiene valor : Proveedor, Moneda, Costo, Documento, Fecha, TipoDoc, SubTipo, Regimen Rubro<br>Favor corregir antes de enviar el email<br>"	            
                        msg = msg & aTableValues4(1,i) & " " & aTableValues4(2,i) & " | " & aTableValues4(3,i) & " | " & aTableValues4(4,i) & " | " & aTableValues4(5,i) & " | " & aTableValues4(6,i) & " | " & aTableValues4(8,i) & " | " &  aTableValues4(10,i)  
                      
            'msg = "Uno o varios de los siguientes datos no tienen valor :"
            'msg = msg & "<table border=0 cellpadding=3><tr bgcolor=silver><th class='style4'>SupplierID</th><th class='style4'>SupplierName</th><th class='style4'>Currency</th><th class='style4'>Value</th><th class='style4'>ReferenceDate</th><th class='style4'>TipoDocID</th><th class='style4'>TipoDocName</th><th class='style4'>SubTipoID</th><th class='style4'>SubTipoName</th><th class='style4'>IsAffected</th></tr>"

                    else 

                        dim categoria, categoriaStr, codigoStr, codigoInt
                        categoria = ""
                        categoriaStr = ""
                        codigoStr = ""
                        codigoInt = CheckNum(aTableValues4(1,i))

                        Select Case CheckNum(aTableValues4(0,i))
	                    Case 0 'LINEA AEREA

                            categoria = "04"
                            categoriaStr = "LINEA AEREA"
                            QuerySelect = "SELECT carrier_id, name FROM carriers WHERE carrier_id = '" & codigoInt & "' " 
                            'response.write QuerySelect & "<br>"
                            Set rs = ConnMaster.Execute(QuerySelect)
	                        if Not rs.EOF then                       
                                codigoStr = rs(1)
                            end if

                        Case 1 'AGENTE 

                            categoria = "02"
                            categoriaStr = "AGENTE"
                            QuerySelect = "SELECT agente_id, agente FROM agentes WHERE agente_id = '" & codigoInt & "' " 
                            'response.write QuerySelect & "<br>"
                            Set rs = ConnMaster.Execute(QuerySelect)
	                        if Not rs.EOF then                       
                                codigoStr = rs(1)
                            end if
                	        
                        Case 2 'NAVIERA

                            categoria = "03"
                            categoriaStr = "NAVIERA"
                            QuerySelect = "SELECT id_naviera, nombre FROM navieras WHERE id_naviera = '" & codigoInt & "' " 
                            'response.write QuerySelect & "<br>"
                            Set rs = ConnMaster.Execute(QuerySelect)
	                        if Not rs.EOF then                       
                                codigoStr = rs(1)
                            end if
                	        
                        Case 3 'PROVEEDOR

                            categoria = "05"
                            categoriaStr = "PROVEEDOR"
                            QuerySelect = "SELECT numero, nombre FROM proveedores WHERE numero = '" & codigoInt & "' " 
                            'response.write QuerySelect & "<br>"
                            Set rs = ConnMaster.Execute(QuerySelect)
	                        if Not rs.EOF then                       
                                codigoStr = rs(1)
                            end if

                        End Select

                        
                        if categoria <> "" then
 
                            QuerySelect = "SELECT eh_erp_codigo, eh_erp_descripcion, eh_estado, eh_erp_categoria FROM exactus_homologaciones WHERE eh_codigo = '" & codigoInt & "' AND eh_categoria = '" & categoria & "'  AND eh_estado = '1' LIMIT 1"
                            'response.write QuerySelect & "<br>"
                            Set rs = ConnMaster.Execute(QuerySelect)
	                        if Not rs.EOF then                       
                                codigo_erp = rs(0)
                            end if
                                                  
                        end if
				
                               
                                
                        QuerySelect = "SELECT eh_erp_codigo, eh_erp_descripcion, eh_estado, eh_erp_categoria FROM exactus_homologaciones WHERE eh_codigo = '" & aTableValues4(3,i) & "' AND eh_categoria = '08'  AND eh_estado = '1' LIMIT 1" 
                        'response.write QuerySelect & "<br>"
                        Set rs = ConnMaster.Execute(QuerySelect)
	                    if Not rs.EOF then                       
                            moneda_erp = rs(0)
                        end if 


                        if codigo_erp <> "-1" then
                                                    
                            On Error Resume Next
                                impuesto = 0
                                valor = CDbl(aTableValues4(4,i))
                                base = valor
                                if CheckNum(aTableValues4(11,i)) = 1 then
                                    base = replace(FormatNumber(valor / TipoIva,2),",","")
                                    impuesto = valor - base
                                    'response.write "(" & TipoIva & ")(" & valor & ")(" & base & ")(" & impuesto & ")" & "<br>"
                                end if

                                QuerySelect = "INSERT INTO exactus_costos (pais_id, pais_iso, blmaster, blid, servicio, proveedor, documento, aplicacion, doc_fecha, moneda, valor, base, impuesto, esquema_erp, tipodoc_erp, subtipo_erp, numero_erp, proveedor_erp, moneda_erp, id_cargo_system, estado, usuario_cs, usuario_ip, movimiento) VALUES ( " & _ 
                                " " & pais_id & ", '" & Session("OperatorCountry") & "', '" & HBLNumber & "', '" & ObjectID & "', '" & Servicio & "', '" & codigoInt & "', '" & aTableValues4(5,i) & "',	'" & aTableValues4(13,i) & "', '" & aTableValues4(6,i) & "', '" & aTableValues4(3,i) & "', " & valor & ", " & base & ", " & impuesto & ",	'',	'" & aTableValues4(7,i) & "', '" & aTableValues4(9,i) & "', '', '" & codigo_erp & "', " & _ 
                                " '" & moneda_erp & "', 2,	1, '" & Session("Login") & "', '" & Request.ServerVariables("REMOTE_ADDR") & "', '" & Movimiento & "') RETURNING id_costo;"
                                'response.write QuerySelect & "<br>"
                                rs = null
                                costo_id = 0
                                Set rs = ConnMaster.Execute(QuerySelect)
                                if Not rs.EOF then                                    
                                    costo_id = CheckNum(rs(0))
                                    'response.write "(costo_id=" & costo_id & ")<br>"
                                end if

                                if costo_id > 0 then
                                     QuerySelect = "UPDATE Costs SET cxp_exactus_id = '" & costo_id & "' WHERE CostID IN (" & aTableValues4(12,i) & ")"
                                     'response.write QuerySelect & "<br>"
                                     Conn.Execute(QuerySelect)
                                     hay_costos = 1
                                else 
                                   msg = "Problemas para almacenar en excel<br>"                     	                           
                                   
                                 end if

                            If Err.Number<>0 then                                '
	                                    
                                msg = ""                       

                                QuerySelect = "SELECT COUNT(*), blmaster FROM exactus_costos WHERE proveedor = '" & codigoInt & "' AND documento = '" & aTableValues4(5,i) & "' AND estado_anulado = '0' GROUP BY blmaster " 
                                'response.write QuerySelect & "<br>"
                                Set rs = ConnMaster.Execute(QuerySelect)
	                            if Not rs.EOF then                       
                                    if CheckNum(rs(0)) > 0 then
                                        msg = categoriaStr & " " & codigoInt & " - " & codigoStr & " Y Documento:" & aTableValues4(5,i) & ", ya fueron digitados " & Iif(IFNULL(rs(1)) <> "","en " & IFNULL(rs(1)),"") & "<br>"                                   
                                    end if
                                end if 

                                if msg = "" then
                                    msg = Err.Number & " - " & Err.Description & "<br><br>"
                                    msg = msg & QuerySelect & "<br>"
                                end if

                            end if

                        else
                            msg = " " & categoriaStr & " NO HOMOLOGADO : " & aTableValues4(1,i) & " - " & codigoStr & "<br>"
                        end if

                    end if

                next 

            end if

        end if

        'pais_iso	blmaster	blid    servicio
        'Session("OperatorCountry") & "', '" & HBLNumber & "', '" & ObjectID & "', '" & Servicio & 

        if hay_costos = 1 then
            
            QuerySelect = "http" & "://10.10.1.20/tools/sendmail.php?id_cargo_system=2&pais_iso=" & Session("OperatorCountry") & "&servicio=" & Servicio & "&blid=" & ObjectID & "&send=1"

            'response.write QuerySelect & "<br>"

            result = GetHTMLSource(QuerySelect)

            'response.write result & "<br>"
    
            msg = "PROCESO CORRECTO!"

        else
            'response.write "&nbsp;<span style='font-size:16px;color:red'>" & result(1) & "</span>"
        end if



    end if



    'if Action = 6 then            
    '    On Error Resume Next
    '        QuerySelect = "http" & "://10.10.1.20/tools/sendmail.php?id_cargo_system=2&pais_iso=" & Session("OperatorCountry") & "&servicio=" & Servicio & "&blid=" & ObjectID & "&send=0"
    '        response.write QuerySelect & "<br>"
    '        GetHTMLSource(QuerySelect)
    '    If Err.Number <> 0 then
	'        response.write Err.Number & " - " & Err.Description & "<br>"
    '    end if
    'end if



    'response.write(Session("OperatorCountry") & " - OperatorCountry")

	'Obteniendo los Costos Master
                    '       0       1           2       3       4       5           6               7           8               9           10       11          12          13         14          15         16      17           18          19                         20                                          21                                                       22                  23                      24                  25                       26                   27
    QuerySelect = "SELECT UserID, ItemName, ItemID, Currency, Cost, SupplierType, SupplierID, SupplierName, Distribution, PurchaseOrder, CostID, ServiceID, ServiceName, Reference, ThirdParties, ProvisionID, '', IsAffected, Countries, SupplierNeutral, DATE_FORMAT(CreatedDate, '%d/%m/%Y'),  DATE_FORMAT(DATE_ADD(CreatedDate, INTERVAL 30 DAY), '%d/%m/%Y'), IFNULL(SubTipoID,''), IFNULL(SubTipoName,''), IFNULL(TipoDocID,''), IFNULL(TipoDocName,''), IFNULL(ReferenceDate,''), cxp_exactus_id, '' FROM Costs WHERE Expired=0 and BLID=" & ObjectID & " and SupplierType not in (5) ORDER BY CostID, ProvisionID Desc, Currency, ServiceName, ItemName"
    'response.write(QuerySelect & "<br>")
    Set rs = Conn.Execute(QuerySelect)
    if Not rs.EOF then
		aTableValues = rs.GetRows
		CountTableValues = rs.RecordCount-1
	end if
	CloseOBJ rs
    

	'Obteniendo los House de la Master
	'response.write "select BLDetailID, HBLNumber, Volumes, Weights from BLDetail where Expired=0 and BLID=" & ObjectID & " Order By BLDetailID<br>"
	Set rs = Conn.Execute("select BLDetailID, HBLNumber, Volumes, Weights from BLDetail where Expired=0 and BLID=" & ObjectID & " Order By BLDetailID")
	if Not rs.EOF then
		aTableValues2 = rs.GetRows
		CountTableValues2 = rs.RecordCount-1
	end if
	CloseOBJ rs

    'Obteniendo los Costos House
	'response.write "select a.CostID, a.SBLID, a.Cost from CostsDetail a, Costs b where a.CostID=b.CostID and b.Expired=0 and b.BLID=" & ObjectID & " Order By a.CostID, a.SBLID<br>"
	Set rs = Conn.Execute("select a.CostID, a.SBLID, a.Cost from CostsDetail a, Costs b where a.CostID=b.CostID and b.Expired=0 and b.BLID=" & ObjectID & " Order By a.CostID, a.SBLID")
	if Not rs.EOF then
		aTableValues3 = rs.GetRows
		CountTableValues3 = rs.RecordCount-1
	end if
CloseOBJs rs, Conn



dim fs,fo,tfile,filename,path

Dim attachments

'path = "C:\Logs\TerrestreCostosExcel" 
'filename = Countries & "_CP_" & HBLNumber & ".xls" se hara por medio de api php excel
'Set fs=Server.CreateObject("Scripting.FileSystemObject")

if Session("OperatorCountry")<>"" then
	openConnBAW Conn

    excel_str = ""
    attachments = ""



	for i=0 to CountTableValues
           
        'response.write(aTableValues(4,i) & "<br>")
		
        ProvID = CheckNum(aTableValues(15,i))
		if ProvID <> 0  then
            'response.write("select tpr_serie, tpr_correlativo from tbl_provisiones where tpr_prov_id=" & ProvID)
            set rs = Conn.Execute("select tpr_serie, tpr_correlativo from tbl_provisiones where tpr_prov_id=" & ProvID)
			if Not rs.EOF then
            	aTableValues(16,i) = rs(0) & "-" & rs(1)
                aTableValues(28,i) = "<font color=blue>REGISTRADO</font>"
            end if
			CloseOBJ rs

        else

		    ProvID = CheckNum(aTableValues(27,i))
		    if ProvID <> 0  then
                set rs = ConnMaster.Execute("select numero_erp, estado from exactus_costos where id_costo=" & ProvID)
			    if Not rs.EOF then
                    aTableValues(15,i) = aTableValues(27,i)
            	    aTableValues(16,i) = ProvID 
                    
                    if rs(0) = "" then 
                        aTableValues(28,i) = "<font color=blue>ENVIADO</font>"
                    else
                        aTableValues(16,i) = aTableValues(16,i) & " - " & rs(0) 
                        aTableValues(28,i) = "<font color=blue>REGISTRADO</font>"
                    end if

                    'select case CheckNum(rs(1))
                    'end select
                     
                end if
			    CloseOBJ rs
		    end if
		end if





	next
	CloseOBJ Conn


    CloseOBJs rs, ConnMaster


    

        
end if






'Si las variables no traen valor se les asigna 0
Volume = CheckNum(Volume)
Weight = CheckNum(Weight)


%>
<HTML>
<HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
var Volume = <%=Volume%>;
var Weight = <%=Weight%>;
var Volumes = new Array();
var Weights = new Array();

<%
    if Session("OperatorCountry")<>"" then
		if TC = 0 then
			response.write "alert('No hay tipo de cambio registrado para este dia, por lo cual no puede guardar costos, favor de comunicarse con Contabilidad');"
		else
			if BAWResult=1 then
				'Ultimo parametro es 2 en la funcion significa que si el resultado de la provision fue 1(exitoso), alerte si hay cobros no facturados aun
				'Ticket#2020061131000104 — INGRESO SDE COSTO POR SISTEMA TERRESTRE 
                'BAWResult = SetBAWProvition (ObjectID, 3, OP, Session("OperatorCountry"), "", 2)
                BAWResult = SetBAWProvition (ObjectID, 3, OP, Countries, "", 2)
				if BAWResult <> "" then
					response.write "alert('" & BAWResult & "');"
				end if
			else
				if BAWResult<>0 then
					'Ultimo parametro es 1 en la funcion significa ir a traer alerta cuando el resultado de la provision no fue 1(exitoso)
					Msg = SetBAWProvition (BAWResult, 3, 0, "", "", 1)
					response.write "alert('" & Msg & "');"
				end if
			end if
		end if	
	end if
%>

//Almacenando los Volumenes y Pesos de cada House, para utilizarlos en el prorrateo por Volumen y Peso
<%for i=0 to CountTableValues2%>
Volumes[<%=i%>] = <%=CheckNum(aTableValues2(2,i))%>;
Weights[<%=i%>] = <%=CheckNum(aTableValues2(3,i))%>;
<%next%>

	function Excel() {

        window.open("http" + "://10.10.1.20/tools/sendmail.php?id_cargo_system=2&pais_iso=<%=Session("OperatorCountry")%>&servicio=<%=Servicio%>&blid=<%=ObjectID%>&send=0");

    }

    function move() {
        document.forma.style.display = "none";
        document.getElementById('myBar').style.display = "block";
        var elem = document.getElementById("myBar");
        var width = 10;
        var id = setInterval(frame, 45);
        function frame() {
            if (width >= 100) {
                clearInterval(id);
                document.getElementById('myBar').style.display = "none";
            } else {
                width++;
                elem.style.width = width + '%';
                elem.innerHTML = width * 1 + '%';
            }
        }        
    }

    function Valida5(Action) {

        if (Action == 5)
            alert('Asegurese de presionar el boton Actualizar antes de Enviar el Email');

        return (false);
    }

	function validar(Action) {

		var sep = '';
		var CantItems=-1;

        document.forma.ItemServIDs_.value = "";
		document.forma.ItemServNames_.value = "";
		document.forma.ItemNames_.value = "";
  		document.forma.ItemIDs_.value = "";
        document.forma.ItemSPRsDate.value = "";

  		document.forma.ItemServIDs.value = "";
		document.forma.ItemServNames.value = "";
		document.forma.ItemNames.value = "";
  		document.forma.ItemIDs.value = "";
  		document.forma.ItemCurrs.value = "";
  		document.forma.ItemCosts.value = "";
  		document.forma.ItemSTypes.value = "";
  		document.forma.ItemSIDs.value = "";
		document.forma.ItemSNames.value = "";
        document.forma.ItemSAffected.value = "";
		document.forma.ItemDistribs.value = "";
		document.forma.ItemSPOs.value = "";
		document.forma.ItemSPRs.value = "";		
		document.forma.ItemThirdParties.value = "";
		document.forma.ItemProvitions.value = "";
        document.forma.ItemCountries.value = "";
        document.forma.ItemNeutrales.value = "";
        document.forma.ItemPos.value = "";

		for (i=0; i<=<%=CantItems%>;i++) {
			if (document.forma.elements["N"+i].value != '') {

				if (!valSelec(document.forma.elements["N"+i])){return Valida5(Action)};

                if (Action == 5)
                {      
                    if (!valTxt(document.forma.elements["SVN_"+i], 1, 10)){return false};
                    if (!valTxt(document.forma.elements["N_"+i], 1, 10)){return false};              
                }

				if (!valSelec(document.forma.elements["TP"+i])){return Valida5(Action)};
				if (!valSelec(document.forma.elements["C"+i])){return Valida5(Action)};
				if (!valTxt(document.forma.elements["CT"+i], 1, 10)){return Valida5(Action)};
				if (!valSelec(document.forma.elements["D"+i])){return Valida5(Action)};
				//Si la Distribucion indicada es Manual, debe existir al menos un registro ingresado y debe totalizar el valor del rubro
				//correspondiente en la Master
				<%if CountTableValues2>=0 then%>
				if (document.forma.elements["D"+i].value==3) {
					if (!CheckDistribution(3, i)){return Valida5(Action)};
				};
				<%end if%>
				if (!valSelec(document.forma.elements["ST"+i])){return Valida5(Action)};
				if (!valTxt(document.forma.elements["SN"+i], 1, 5)){return Valida5(Action)};


                if (Action == 5)
                {      
                    if (!valTxt(document.forma.elements["SPR"+i], 1, 10)){return false};
                    if (!valTxt(document.forma.elements["SPRDate"+i], 1, 10)){return false};                
                }


                if (!valSelec(document.forma.elements["SAF"+i])){return Valida5(Action)};
                if (!valTxt(document.forma.elements["PRVCTR"+i], 1, 5)){return Valida5(Action)};

				if (document.forma.elements["SVI"+i].value!="") {
					document.forma.ItemServIDs.value = document.forma.ItemServIDs.value + sep + document.forma.elements["SVI"+i].value;
					document.forma.ItemServNames.value = document.forma.ItemServNames.value + sep + document.forma.elements["SVN"+i].value;
				} else {
					document.forma.ItemServIDs.value = "0" + sep + document.forma.elements["SVI"+i].value;
					document.forma.ItemServNames.value = " " + sep + document.forma.elements["SVN"+i].value;
				}

                if (document.forma.elements["SVI_"+i].value!="") {
					document.forma.ItemServIDs_.value = document.forma.ItemServIDs_.value + sep + document.forma.elements["SVI_"+i].value;
					document.forma.ItemServNames_.value = document.forma.ItemServNames_.value + sep + document.forma.elements["SVN_"+i].value;
				} else {
					document.forma.ItemServIDs_.value = "0" + sep + document.forma.elements["SVI_"+i].value;
					document.forma.ItemServNames_.value = " " + sep + document.forma.elements["SVN_"+i].value;
				}

				document.forma.ItemNames.value = document.forma.ItemNames.value + sep + document.forma.elements["N"+i].value;
				document.forma.ItemIDs.value = document.forma.ItemIDs.value + sep + document.forma.elements["I"+i].value;

                document.forma.ItemNames_.value = document.forma.ItemNames_.value + sep + document.forma.elements["N_"+i].value;
				document.forma.ItemIDs_.value = document.forma.ItemIDs_.value + sep + document.forma.elements["I_"+i].value;

				document.forma.ItemThirdParties.value = document.forma.ItemThirdParties.value + sep + document.forma.elements["TP"+i].value;
				document.forma.ItemCurrs.value = document.forma.ItemCurrs.value + sep + document.forma.elements["C"+i].value;
				document.forma.ItemCosts.value = document.forma.ItemCosts.value + sep + document.forma.elements["CT"+i].value;
				document.forma.ItemSTypes.value = document.forma.ItemSTypes.value + sep + document.forma.elements["ST"+i].value;
				document.forma.ItemSIDs.value = document.forma.ItemSIDs.value + sep + document.forma.elements["SI"+i].value;
				document.forma.ItemSNames.value = document.forma.ItemSNames.value + sep + document.forma.elements["SN"+i].value;
				document.forma.ItemSAffected.value = document.forma.ItemSAffected.value + sep + document.forma.elements["SAF"+i].value;
                document.forma.ItemDistribs.value = document.forma.ItemDistribs.value + sep + document.forma.elements["D"+i].value;	
				document.forma.ItemProvitions.value = document.forma.ItemProvitions.value + sep + document.forma.elements["PRV"+i].value;	
                document.forma.ItemCountries.value = document.forma.ItemCountries.value + sep + document.forma.elements["PRVCTR"+i].value;	
                document.forma.ItemNeutrales.value = document.forma.ItemNeutrales.value + sep + document.forma.elements["SNEU"+i].value;	
                
                //Validando que para Latin Freight solo se puede provisionar Agentes Coloader = 1
                <%if FilterAimarLatin = 1 then%>
                if ((document.forma.elements["SNEU"+i].value==0) && (document.forma.elements["PRVCTR"+i].value.substr(2,3)=="LTF")) {
                    alert("En Latin Freight solo puede provisionar Agentes Neutrales");
                    document.forma.elements["SN"+i].focus();
                    return Valida5(Action);
                }
                <%end if %>
                if (document.forma.elements["SPO"+i].value=="") {
					document.forma.ItemSPOs.value = document.forma.ItemSPOs.value + sep + " ";
				} else {
					document.forma.ItemSPOs.value = document.forma.ItemSPOs.value + sep + document.forma.elements["SPO"+i].value;	
				}
				if (document.forma.elements["SPR"+i].value=="") {
					document.forma.ItemSPRs.value = document.forma.ItemSPRs.value + sep + " ";
				} else {
					document.forma.ItemSPRs.value = document.forma.ItemSPRs.value + sep + document.forma.elements["SPR"+i].value;	
				}

                if (document.forma.elements["SPRDate"+i].value=="") {
					document.forma.ItemSPRsDate.value = document.forma.ItemSPRsDate.value + sep + " ";
				} else {
					document.forma.ItemSPRsDate.value = document.forma.ItemSPRsDate.value + sep + document.forma.elements["SPRDate"+i].value;	
				}

				//indica la posicion de la fila(Master) y columa(House) donde se esta guardando el dato
				document.forma.ItemPos.value = document.forma.ItemPos.value + sep + i;
				CantItems++;
				sep = "|";
			}
		}

        if (CantItems == -1)
            return false;

        if (!confirm( Action == 5 ? 'Confirme Envio de Email ? ' : 'Confirme Actualizar datos'  )) 
			return false;
                
        move();
	    document.forma.CantItems.value = CantItems;
		document.forma.Action.value = Action;
        document.forma.submit();	

		//alert(document.forma.ItemNames.value);
		//alert(document.forma.ItemIDs.value);
		//alert(document.forma.ItemThirdParties.value);
		//alert(document.forma.ItemCurrs.value);
		//alert(document.forma.ItemCosts.value);
		//alert(document.forma.ItemSTypes.value);
		//alert(document.forma.ItemSIDs.value);
		//alert(document.forma.ItemSNames.value);
        //alert(document.forma.ItemSAffected.value);
		//alert(document.forma.ItemDistribs.value);
		//alert(document.forma.ItemSPOs.value);
		//alert(document.forma.ItemSPRs.value);
		//alert(document.forma.CantItems.value);
		//alert(document.forma.ItemProvitions.value);
        //alert(document.forma.ItemCountries.value);
		//alert(document.forma.CantHouses.value);
		//alert(document.forma.ItemNeutrales.value);
        //alert(document.forma.ItemPos.value);        
		 
	 }

	 function DelCharge(Pos) {
		document.forma.elements["SVI_"+Pos].value='';
		document.forma.elements["SVN_"+Pos].value='';
		document.forma.elements["N_"+Pos].value='';
		document.forma.elements["I_"+Pos].value='';
        document.forma.elements["SPRDate"+Pos].value='';

		document.forma.elements["SVI"+Pos].value='';
		document.forma.elements["SVN"+Pos].value='';
		document.forma.elements["N"+Pos].value='';
		document.forma.elements["I"+Pos].value='';
		document.forma.elements["TP"+Pos].value='-1';
		document.forma.elements["C"+Pos].value='-1';
		document.forma.elements["CT"+Pos].value=''; 
		document.forma.elements["ST"+Pos].value='-1';
		document.forma.elements["SN"+Pos].value='';
		document.forma.elements["SI"+Pos].value='';
        document.forma.elements["SAF"+Pos].value='-1';
        document.forma.elements["SNEU"+Pos].value='0';
        document.forma.elements["SPO"+Pos].value='';
		document.forma.elements["SPR"+Pos].value='';		
		document.forma.elements["D"+Pos].value='-1';
		document.forma.elements["PRV"+Pos].value='0';
        //document.forma.elements["PRVCTR"+Pos].value='';
		DelHouseCharges(Pos);
		return false; 
	 }
	
    function ResetDataprovider(Pos) {
        document.forma.elements["SN"+Pos].value='';
		document.forma.elements["SI"+Pos].value='';
        document.forma.elements["SAF"+Pos].value='-1';
        document.forma.elements["SNEU"+Pos].value='0';
    }
    
    function CheckThirdParties(Pos) {
        //Si indican que el rubro si viene en el BL, entonces el servicio solo puede ser de Terceros=14
        if ((document.forma.elements["TP"+Pos].value == 1) && (document.forma.elements["SVI"+Pos].value != 14))
        { 
            alert("Si el costo viene en el BL solo puede utilizar Rubros del servicio de Terceros");

            document.forma.elements["SVI_"+Pos].value='';
		    document.forma.elements["SVN_"+Pos].value='';
		    document.forma.elements["N_"+Pos].value='';
		    document.forma.elements["I_"+Pos].value='';

            document.forma.elements["SVI"+Pos].value='';
		    document.forma.elements["SVN"+Pos].value='';
		    document.forma.elements["N"+Pos].value='';
		    document.forma.elements["I"+Pos].value='';
            document.forma.elements["TP"+Pos].value='-1';
            return Valida5(Action);
        }
    }

	function DelHouseCharges(Pos) {
		for (i=0; i<=<%=CountTableValues2%>;i++) {
			document.forma.elements["H"+i+"C"+Pos].value='';
		}
		return false; 
	 }	 
	 
	 function SetDistribution(DType, Pos){
		 switch(DType) {
			case "1"://Peso				
				if (Weight==0) {
					alert("No se puede distribuir los costos por Peso, ya que el Peso indicado en la Master es 0");
					document.forma.elements["D"+Pos].focus();
					return (false);
				}
				for (i=0; i<=<%=CountTableValues2%>;i++) {
					document.forma.elements["H"+i+"C"+Pos].value=Math.round(document.forma.elements["CT"+Pos].value*(Weights[i]/Weight)*100)/100;
				}
				break;
			case "2"://Volumen
				if (Volume==0) {
					alert("No se puede distribuir los costos por Volumen, ya que el Volumen indicado en la Master es 0");
					document.forma.elements["D"+Pos].focus();
					return (false);
				}
				for (i=0; i<=<%=CountTableValues2%>;i++) {
					document.forma.elements["H"+i+"C"+Pos].value=Math.round(document.forma.elements["CT"+Pos].value*(Volumes[i]/Volume)*100)/100;
				}
				break;
			case "3"://Manual
				DelHouseCharges(Pos);
				break;
			case "4"://Sin Distribucion
				DelHouseCharges(Pos);
				break;
			case "5"://Promedio
				for (i=0; i<=<%=CountTableValues2%>;i++) {
					document.forma.elements["H"+i+"C"+Pos].value=Math.round(document.forma.elements["CT"+Pos].value/<%=CountTableValues2+1%>*100)/100;
				}
				break;
		 }
    }
	  
	function CheckDistribution(DType, Pos){
		var Verification=0;
        if (DType==3) { //Si la opcion es Manual, el total de valores en Houses debe sumar igual al total ingresado para el rubro correspondiente en la Master
			for (j=0; j<=<%=CountTableValues2%>;j++) {
				Verification += document.forma.elements["H"+j+"C"+Pos].value*1;
			}
			//Redondeo de 2 decimales, ya que la sumatoria del "for" genera mas de 2 decimales
			Verification = Math.round(Verification*100)/100;

                //alert(Verification + " " + document.forma.elements["CT"+Pos].value);
		    
            if (Verification!=document.forma.elements["CT"+Pos].value) {
				alert("La suma de costos en los Houses no concide con el total indicado por la Master en la columna "+(Pos+1));
				document.forma.elements["CT"+Pos].focus();
				return (false);
			}
		 }
		 return (true);
	 }	 
	 
	 function AddCharge(Pos) {
		window.open('Search_Charges.asp?GID=29&N='+Pos+'&T=<%=BLType%>','BLData','height=300,width=450,menubar=0,resizable=1,scrollbars=1,toolbar=0');
		return false;	 
	 }

	 function AddTipoDoc(Pos) {
		window.open('Search_TipoDoc.asp?GID=29&N='+Pos+'&T=<%=BLType%>','BLData','height=300,width=450,menubar=0,resizable=1,scrollbars=1,toolbar=0');
		return false;	 
	 }
     	 
	 function AddProvider(Pos) {
		if (!valSelec(document.forma.elements["ST"+Pos])){return (false)};
		window.open('Search_BLData.asp?GID=30&ST='+document.forma.elements["ST"+Pos].value+'&N='+Pos,'Supplier','height=300,width=450,menubar=0,resizable=1,scrollbars=1,toolbar=0');
		return false;
	 }
	 
	 function ValidarDoble(Pos) {
	 	for (i=0; i<=<%=CantItems%>;i++) {
			if  (i!= Pos) {
				if ((document.forma.elements["SVI"+i].value==document.forma.elements["SVI"+Pos].value) && 
				(document.forma.elements["SVN"+i].value==document.forma.elements["SVN"+Pos].value) &&
				(document.forma.elements["N"+i].value==document.forma.elements["N"+Pos].value) &&
				(document.forma.elements["I"+i].value==document.forma.elements["I"+Pos].value) &&

                (document.forma.elements["SVI_"+i].value==document.forma.elements["SVI_"+Pos].value) && 
				(document.forma.elements["SVN_"+i].value==document.forma.elements["SVN_"+Pos].value) &&
				(document.forma.elements["N_"+i].value==document.forma.elements["N_"+Pos].value) &&
				(document.forma.elements["I_"+i].value==document.forma.elements["I_"+Pos].value) &&

				(document.forma.elements["ST"+i].value==document.forma.elements["ST"+Pos].value) && 
				(document.forma.elements["SN"+i].value==document.forma.elements["SN"+Pos].value) &&
				(document.forma.elements["SI"+i].value==document.forma.elements["SI"+Pos].value) &&
                ((document.forma.elements["PRV"+i].value=='') || (document.forma.elements["PRV"+i].value==document.forma.elements["PRV"+Pos].value)) &&
				((document.forma.elements["SPR"+i].value=='') || (document.forma.elements["SPR"+i].value==document.forma.elements["SPR"+Pos].value)) &&

				((document.forma.elements["SPRDate"+i].value=='') || (document.forma.elements["SPRDate"+i].value==document.forma.elements["SPRDate"+Pos].value)) &&

				(document.forma.elements["SVI"+i].value!='') && 
				(document.forma.elements["SVN"+i].value!='') &&
				(document.forma.elements["N"+i].value!='') &&
				(document.forma.elements["I"+i].value!='') &&

				(document.forma.elements["SVI_"+i].value!='') && 
				(document.forma.elements["SVN_"+i].value!='') &&
				(document.forma.elements["N_"+i].value!='') &&
				(document.forma.elements["I_"+i].value!='') &&

				(document.forma.elements["ST"+i].value!='') && 
				(document.forma.elements["SN"+i].value!='') &&
				(document.forma.elements["SI"+i].value!=''))
				 {
					alert("No puede repetir el mismo Rubro, Servicio y Proveedor, si el anterior no ha sido facturado primero o tiene Referencia diferente");
					DelCharge(Pos);
					return (false);
				}
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



</SCRIPT>
<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">
<style type="text/css">
<!--
.style10 {font-size: 10px; font-family: Verdana, Arial, Helvetica, sans-serif; }
.style4 {	font-size:10px;
	color: #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
}
.style8 {	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 10px;
	font-weight: bold;
	color: #999999;
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

-->
</style>
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="self.focus();">
<div id="myProgress">
  <div id="myBar">10%</div>
</div>

<%if Session("OperatorCountry")<>"" then
	if TC = 0 then%>
    <TABLE cellspacing=0 align="center">
        <TR>
        <TD class=label align=center colspan=2><font color=red size="2">No hay tipo de cambio registrado para este dia, por lo cual no puede guardar costos, favor de comunicarse con Contabilidad</font></TD>
        </TR>
    </TABLE>
    <%end if
      if Msg <> "" then%>
    <TABLE cellspacing=0 align="center">
        <TR>
        <TD class=label align=center colspan=2><font color=red size="2"><%=Msg%></font></TD>
        </TR>
    </TABLE>
	<%end if
 end if%>

<div style="width:96%;display:block;border:0px solid red;text-align:center;padding-left:2%;" >


	<FORM name="forma" action="Costs.asp" method="post">
	<INPUT name="Action" type=hidden value=0>
	<INPUT name="OID" type=hidden value="<%=ObjectID%>">
	<INPUT name="ItemServIDs" type=hidden value="">
	<INPUT name="ItemServNames" type=hidden value="">
	<INPUT name="ItemNames" type=hidden value="">
	<INPUT name="ItemIDs" type=hidden value="">

	<INPUT name="ItemServIDs_" type=hidden value="">
	<INPUT name="ItemServNames_" type=hidden value="">
	<INPUT name="ItemNames_" type=hidden value="">
	<INPUT name="ItemIDs_" type=hidden value="">
    <INPUT name="ItemSPRsDate" type=hidden value="">

	<INPUT name="ItemCurrs" type=hidden value="">
	<INPUT name="ItemCosts" type=hidden value="">
	<INPUT name="ItemSIDs" type=hidden value="">
	<INPUT name="ItemSTypes" type=hidden value="">
	<INPUT name="ItemSNames" type=hidden value="">
    <INPUT name="ItemSAffected" type=hidden value="">
	<INPUT name="ItemDistribs" type=hidden value="">
	<INPUT name="ItemSPOs" type=hidden value="">
	<INPUT name="ItemSPRs" type=hidden value="">
	
	<INPUT name="CantItems" type=hidden value="-1">
	<INPUT name="CantHouses" type=hidden value="<%=CountTableValues2%>">
	<INPUT name="ItemPos" type=hidden value="">
	<INPUT name="ItemThirdParties" type=hidden value="">
	<INPUT name="ItemProvitions" type=hidden value="">
    <INPUT name="ItemCountries" type=hidden value="">
    <INPUT name="ItemNeutrales" type=hidden value="">



    	<table width="99%" align="center" border="0">
        <tr><td class=submenu colspan=2></td></tr>
	    <tr>
        <td>
		        <table width="90%" align="left">
			        <TR><TD class=label align=right><b>Carta Porte:</b></TD><TD class=label align=left><%=HBLNumber%></TD><TD class=label align=right><b>Volumen:</b></TD><TD class=label align=left><%=Volume%></TD></TR> 
			        <TR><TD class=label align=right><b>Consignatario:</b></TD><TD class=label align=left><%=Name(0)%></TD><TD class=label align=right><b>Peso:</b></TD><TD class=label align=left><%=Weight%></TD></TR>
			        <%select case BLType
			        case 0,1%>
			        <TR><TD class=label align=right><b>Agente:</b></TD><TD class=label align=left><%=Agent(0)%></TD><TD class=label align=right><b>Movimiento:</b></TD><TD class=label align=left><%=Movimiento%></TD></TR> 
			        <%end select%>
		        </table>		
        
        </td>        
        <td valign=bottom>
            <%if Session("OperatorCountry")<>"" then
			    if TC <> 0 then%>
				    <TABLE cellspacing=0 cellpadding=2 width=200 align=left border=0>
				    <TR>
					    <%if CountTableValues < 0 then%>
						    <TD class=label align=center><INPUT name=enviar type=button onClick="JavaScript:validar(1);" value="&nbsp;&nbsp;Guardar&nbsp;&nbsp;" class=label style="color:white0;background-color:rgb(153,102,0)0"></TD>
					    <%else%>
						    <TD class=label align=center><INPUT name=enviar type=button onClick="JavaScript:validar(2);" value="&nbsp;&nbsp;Actualizar&nbsp;&nbsp;" class=label style="color:white0;background-color:rgb(153,102,0)0"></TD>
					    <%end if%>

                    <%if TipoConta = "EXACTUS" and CountTableValues > -1 then %>

                        <TD class=label align=center>
                            <INPUT name=enviar type=button onClick="JavaScript:validar(5);" value="&nbsp;&nbsp;Enviar Email&nbsp;&nbsp;" class=label style="color:white0;background-color:rgb(102,0,0)0">
                        </TD>
                        <TD class=label align=center>
                            <INPUT name=enviar type=button onClick="JavaScript:Excel();" value="&nbsp;&nbsp;Ver Excel&nbsp;&nbsp;" class=label style="color:white0;background-color:rgb(102,0,0)0">
                        </TD>

                    <% else %>

                        <TD class=label align=center>
                            <INPUT name=enviar type=button  value="&nbsp;&nbsp;Enviar Email&nbsp;&nbsp;"  onclick="alert('Debe tener configurada contabilidad EXACTUS');" class=label style="color:white0;background-color:rgb(102,0,0)0">
                        </TD>
                        <TD class=label align=center>
                            <INPUT name=enviar type=button  value="&nbsp;&nbsp;Ver Excel&nbsp;&nbsp;" disabled class=label style="color:white0;background-color:rgb(102,0,0)0">
                        </TD>
    		
                    <%end if%>

				    </TR>
				    </TABLE> 

                <% else %>
                    
                    <TD class=label align=center colspan=2><INPUT name=enviar type=button value="&nbsp;&nbsp;Guardar&nbsp;&nbsp;" class=label disabled style="color:white;background-color:rgb(153,102,0)"></TD>
     
    		    <%end if


		    else%>            
                <TABLE cellspacing=0 cellpadding=2 width=200 align=left>
                <TR>
                    <%if CountTableValues < 0 then%>
                        <TD class=label align=center><INPUT name=enviar type=button onClick="JavaScript:validar(1);" value="&nbsp;&nbsp;Guardar&nbsp;&nbsp;" class=label style="color:white;background-color:rgb(153,102,0)"></TD>
                    <%else%>
                        <TD class=label align=center><INPUT name=enviar type=button onClick="JavaScript:validar(2);" value="&nbsp;&nbsp;Actualizar&nbsp;&nbsp;" class=label style="color:white;background-color:rgb(153,102,0)"></TD>
                    <%end if%>


                    <%if TipoConta = "EXACTUS" and CountTableValues > -1 then %>

                        <TD class=label align=center>
                            <INPUT name=enviar type=button onClick="JavaScript:validar(5);" value="&nbsp;&nbsp;Enviar Email&nbsp;&nbsp;" class=label style="color:white0;background-color:rgb(102,0,0)0">
                        </TD>
                        <TD class=label align=center>
                            <INPUT name=enviar type=button onClick="JavaScript:Excel();" value="&nbsp;&nbsp;Ver Excel&nbsp;&nbsp;" class=label style="color:white0;background-color:rgb(102,0,0)0">
                        </TD>

                    <% else %>

                        <TD class=label align=center>
                            <INPUT name=enviar type=button  value="&nbsp;&nbsp;Enviar Email&nbsp;&nbsp;" disabled class=label style="color:white0;background-color:rgb(102,0,0)0">
                        </TD>
                        <TD class=label align=center>
                            <INPUT name=enviar type=button  value="&nbsp;&nbsp;Ver Excel&nbsp;&nbsp;" disabled class=label style="color:white0;background-color:rgb(102,0,0)0">
                        </TD>
    		
                    <%end if%>          

                </TR>
                </TABLE>		
		    <%end if%>        
        </td>        
        </tr>
        </table>



		<div style="width:99%;height:350px;display:block;overflow:auto;" >
        <table width="80%" border="0" align="center">
		  <tr><td class=submenu colspan=15></td></tr>
		  <tr>
			<td align="center" class="style4">Servicio</td>
			<td align="center" class="style4">Rubro</td>
			<td align="center" class="style4">&nbsp;</td>

            <%if TipoConta = "EXACTUS" then%>

			<td align="center" class="style4">TipoDoc</td>
			<td align="center" class="style4">SubTipo</td>
			<td align="center" class="style4">&nbsp;</td>
            
            <% else %>

            <% end if %>

			<td align="center" class="style4">Viene en BL?</td>
			<td align="center" class="style4">Moneda</td>
			<td align="center" class="style4">Costo</td>
			<td align="center" class="style4">Prorrateo</td>
			<td align="center" class="style4">Proveedor</td>
			<td align="center" class="style4" colspan="2">Nombre&nbsp;Proveedor</td>
            <td align="center" class="style4">Regimen Rubro</td>
			<td align="center" class="style4">Orden&nbsp;de&nbsp;Compra</td>
			<td align="center" class="style4">Documento</td>
			<td align="center" class="style4">Fecha</td>
            <td align="center" class="style4"><%=Iif(TipoConta="BAW","Provision","Exactus Costo")%></td>
			<td align="center" class="style4">Estado</td>
			<td align="center" class="style4">Pais Provision</td>
			<td align="center" class="style4">&nbsp;</td>


		  </tr>
		  <%for i=0 to CantItems%>
		  <tr>


			<td align="right" class="style4">
				<%=i+1%>.&nbsp;<input type="text" size="20" class="style10" name="SVN<%=i%>" value="" readonly>
				<input type="hidden" name="SVI<%=i%>" value="">
			</td>
			<td align="right" class="style4">
				<input type="text" size="20" class="style10" name="N<%=i%>" value="" readonly>
				<input type="hidden" name="I<%=i%>" value="">
			</td>
			<td align="right" class="style4">
				<div id=DR<%=i%> s style="VISIBILITY: visible;"><a href="#" onClick="Javascript:AddCharge(<%=i%>);" class="menu"><font color="FFFFFF">Buscar</font></a></div>
			</td>

            <!-- TIPO DOCUMENTO EXACTOS -->

            <%if TipoConta = "EXACTUS" then%>

			<td align="right" class="style4">
				<input type="text" size="20" class="style10" id="Tipo Documento Exactus" name="SVN_<%=i%>" value="" readonly>
				<input type="hidden" name="SVI_<%=i%>" value="">
			</td>
			<td align="right" class="style4">
				<input type="text" size="20" class="style10" id="SubTipo Documento Exactus" name="N_<%=i%>" value="" readonly>
				<input type="hidden" name="I_<%=i%>" value="">
			</td>
			<td align="right" class="style4">
				<div id=DR_<%=i%> style="VISIBILITY: visible;"><a href="#" onClick="Javascript:AddTipoDoc(<%=i%>);" class="menu"><font color="FFFFFF">Buscar</font></a></div>
			</td>
            
            <% else %>

				<input type="hidden" name="SVN_<%=i%>" value="">
				<input type="hidden" name="SVI_<%=i%>" value="">
				<input type="hidden" name="N_<%=i%>" value="">
				<input type="hidden" name="I_<%=i%>" value="">

            <% end if %>

            <!-- -------------------- -->

			<td align="right" class="style4">
				<select class='style10' name='TP<%=i%>' id="Viene en BL" onChange="Javascript:CheckThirdParties(<%=i%>);">
				<option value='-1'>---</option>
				<option value='0'>NO</option>
				<option value='1'>SI</option>
				</select>
			</td>
			<td align="right" class="style4">
            	<select class='style10' name='C<%=i%>' id="Moneda" <%if Session("OperatorCountry")="NI" then  %>disabled<%end if %>>
				<option value='-1'>---</option>
				<%=Currencies%>
				</select>
			</td>
			<td align="center" class="style4">
				<input type="text" size="7" class="style10" name="CT<%=i%>" value="" onKeyUp="res(this,numb);" onBlur="SetDistribution(document.forma.D<%=i%>.value,<%=i%>);" id="Costos" <%if Session("OperatorCountry")="NI" then  %>readonly onclick="window.open('CostsConvert.asp?TC=<%=TC%>&N=<%=i%>','ConvertCosts','height=100,width=250,menubar=0,resizable=1,scrollbars=1,toolbar=0');return false;"<%end if %>>
			</td>
			<td align="right" class="style4">
				<select class='style10' name='D<%=i%>' id="Tipo de Prorrateo" onChange="Javascript:SetDistribution(this.value,<%=i%>);">
				<option value='-1'>---</option>
				<option value='1'>PESO</option>
				<option value='2'>VOLUMEN</option>
				<option value='3'>MANUAL</option>
				<option value='4'>SIN DISTRIBUCION</option>
				<option value='5'>PROMEDIO</option>
			 	</select>
			</td>
			<td align="right" class="style4">
				<select class='style10' name='ST<%=i%>' id="Tipo Proveedor" onchange="Javascript:ResetDataprovider(<%=i%>);">
				<option value='-1'>---</option>
				<option value='0'>LINEA AEREA</option>
				<option value='1'>AGENTE</option>
				<option value='2'>NAVIERA</option>
				<option value='3'>PROVEEDOR</option>
			 	</select>
			</td>
			<td align="right" class="style4">
				<input type="text" size="30" class="style10" name="SN<%=i%>" value="" id="Nombre Proveedor" readonly>
				<input type="hidden" name="SI<%=i%>" value="">
                <input type="hidden" name="SNEU<%=i%>" value="">
			</td>
            <td align="right" class="style4">
				<div id="DP<%=i%>" style="VISIBILITY: visible;"><a href="#" onClick="Javascript:AddProvider(<%=i%>);" class="menu"><font color="FFFFFF">Buscar</font></a></div>
			</td>
            <td align="right" class="style4">
				<select class='style10' name='SAF<%=i%>' id="Regimen Tributario">
				<option value='-1'>---</option>
				<option value='0'>NO AFECTO</option>
				<option value='1'>AFECTO</option>
				</select>
			</td>
            <td align="right" class="style4">
				<input type="text" size="17" class="style10" name="SPO<%=i%>" value="">
			</td>
			<td align="right" class="style4">
				<input type="text" size="12" class="style10" id="Documento" name="SPR<%=i%>" value="" onBlur="Javascript:ValidarDoble(<%=i%>);" maxlength=50>
			</td>

            <td align="right" class="style4" nowrap>
		       <INPUT readonly="readonly" name="SPRDate<%=i%>" id="Fecha Documento" type=text value="" size=10 maxLength=14 class=label>		
               <INPUT type=image onClick="return abrir('SPRDate<%=i%>');" src="img/calendar.png" id="SPRDate_<%=i%>">
			</td>

            <td align="right" class="style4">
				<input type="text" size="12" class="style10" name="PROVID<%=i%>" value="" disabled>
			</td>

            <td align="right" class="style4">
				<div class="style10" id="PROVES<%=i%>"></div>
			</td>

			 <td align="right" class="style4">
				<input type="text" size="12" class="style10" name="PRVCTR<%=i%>" id="Pais para Provision" value="<%=Session("OperatorCountry")%>" readonly>
			</td>
			<td align="right" class="style4">
				<div id="DE<%=i%>" style="VISIBILITY: visible;">
				<a href="#" onClick="Javascript:DelCharge(<%=i%>);" class="menu"><font color="FFFFFF">X</font></a>
				</div>
				<input type="hidden" name="PRV<%=i%>" value="0">
			</td>	
            
            			
		  </tr>
		  <%next%>
		</table>
        </div>

 

        <div style="width:99%;display:block;overflow:auto;float:left" >

		<table width="100%" border="0" align=left>
		<%if CountTableValues2>=0 then%>
		  <tr><td class=submenu colspan=15></td></tr>
		  <tr>
			<td align="center" class="style4">
				Cartas Porte House
			</td>
		  <%for i=0 to CantItems%>
			<td align="center" class="style4">
				<%=i+1%>.
			</td>
		  <%next%>
		  </tr>
		  <%for i=0 to CountTableValues2%>
		  <tr>
			<td align="center" class="style4">
				<input type="hidden" name="SBLID<%=i%>" value="<%=aTableValues2(0,i)%>">
				<div id="SBLIDN<%=i%>" style="VISIBILITY: visible; width:150px;"></div>
			</td>
			<%for j=0 to CantItems%>
			<td align="center" class="style4">
				<input type="text" size="7" class="style10" name="H<%=i%>C<%=j%>" value="" onKeyUp="res(this,numb);" id="Costos">
			</td>
		  	<%next%>
		  </tr>
		  <%next%>
		  <%else%>
		  <tr><td class=submenu colspan=14></td></tr>
		  <tr>
			<td align="center" class="style4">
				El Master seleccionado no tiene Houses asignados
			</td>
		</tr>
		  <%end if%>
		</table>

        </div>
       
        <input type="hidden" name="TipoConta" value="<%=TipoConta%>" /> 
</div>

</BODY>
</HTML>


<%
'Dim TestArray, fechaStr
'response.write "(CountTableValues=" & CountTableValues & ")<br>"

%>

<script>
<%'Seteando los costos de la Master


for i=0 to CountTableValues

    //response.write ConvertDate(aTableValues(26,i), 2) & "<br>"               
    //fechaStr = CStr(ConvertDate(aTableValues(26,i), 2))
    //response.write fechaStr & "<br>"
    //TestArray = Split(fechaStr,"/")
    //aTableValues(26,i) = TestArray(2) & "/" & TestArray(1) & "/" & TestArray(0)
    //response.write aTableValues(26,i) & "<br>"
%>

    try {

	document.forma.N<%=i%>.value = '<%=aTableValues(1,i)%>';
	document.forma.I<%=i%>.value = '<%=aTableValues(2,i)%>';
	selecciona('forma.TP<%=i%>','<%=aTableValues(14,i)%>');
	selecciona('forma.C<%=i%>','<%=aTableValues(3,i)%>');
	document.forma.CT<%=i%>.value = '<%=aTableValues(4,i)%>';
	selecciona('forma.ST<%=i%>','<%=aTableValues(5,i)%>');
	document.forma.SI<%=i%>.value = '<%=aTableValues(6,i)%>';
	document.forma.SN<%=i%>.value = '<%=aTableValues(7,i)%>';
	selecciona('forma.D<%=i%>','<%=aTableValues(8,i)%>');
	document.forma.SPO<%=i%>.value = '<%=aTableValues(9,i)%>';
	document.forma.SVI<%=i%>.value = '<%=aTableValues(11,i)%>';
	document.forma.SVN<%=i%>.value = '<%=aTableValues(12,i)%>';
	document.forma.SPR<%=i%>.value = '<%=aTableValues(13,i)%>';
	document.forma.PRV<%=i%>.value = '<%=aTableValues(15,i)%>';
	document.forma.PROVID<%=i%>.value = '<%=aTableValues(16,i)%>';
    selecciona('forma.SAF<%=i%>','<%=aTableValues(17,i)%>');
	document.forma.PRVCTR<%=i%>.value = '<%=aTableValues(18,i)%>';
    document.forma.SNEU<%=i%>.value = '<%=aTableValues(19,i)%>';


    document.forma.SVN_<%=i%>.value = '<%=aTableValues(25,i)%>';
	document.forma.SVI_<%=i%>.value = '<%=aTableValues(24,i)%>';
    document.forma.N_<%=i%>.value = '<%=aTableValues(23,i)%>';
    document.forma.I_<%=i%>.value = '<%=aTableValues(22,i)%>';

    <%if (aTableValues(26,i) <> "0000-00-00") then%>    
        document.forma.SPRDate<%=i%>.value = '<%=ConvertDate(aTableValues(26,i),6)%>';  
    <% end if %>

    document.getElementById('PROVES<%=i%>').innerHTML = '<%=aTableValues(28,i)%>';
    


    <% if aTableValues(15,i) <> 0 or aTableValues(27,i) then%>


        document.forma.N<%=i%>.disabled = 'false';
		document.forma.I<%=i%>.disabled = 'false';
		document.forma.TP<%=i%>.disabled = 'false';
		document.forma.C<%=i%>.disabled = 'false';
		document.forma.CT<%=i%>.disabled = 'false';
		document.forma.ST<%=i%>.disabled = 'false';
		document.forma.SI<%=i%>.disabled = 'false';
		document.forma.SN<%=i%>.disabled = 'false';
		document.forma.D<%=i%>.disabled = 'false';
		document.forma.SAF<%=i%>.disabled = 'false';
        document.forma.SPO<%=i%>.disabled = 'false';
		document.forma.SVI<%=i%>.disabled = 'false';
		document.forma.SVN<%=i%>.disabled = 'false';
		document.forma.SPR<%=i%>.disabled = 'false';
		document.forma.PRVCTR<%=i%>.disabled = 'false';
        document.getElementById("DE<%=i%>").style.visibility = "hidden";
		document.getElementById("DR<%=i%>").style.visibility = "hidden";
		document.getElementById("DP<%=i%>").style.visibility = "hidden";

		document.getElementById("DR_<%=i%>").style.visibility = "hidden";
		document.forma.N_<%=i%>.disabled = 'false';
		document.forma.I_<%=i%>.disabled = 'false';
		document.forma.SVI_<%=i%>.disabled = 'false';
		document.forma.SVN_<%=i%>.disabled = 'false';

		document.forma.SPRDate<%=i%>.disabled = 'false';
        document.getElementById("SPRDate_<%=i%>").style.visibility = "hidden";



	<%end if%>
		
            

	
<%  'Seteando los costos distribuidos a los Houses, cuando es no es costo directo (4=No Distribuido)
	if aTableValues(8,i) <> 4 and CountTableValues3>=0 then
        'Recorriendo el Numero de Houses que existen para presentar sus costos asignados
        for j=0 to CountTableValues2
			'Ingreso mientras el CostID de la Master es igual al CostID del House
			response.write "//COSTOS: " & i & "-" & j & "-" & k & "-" & aTableValues(10,i) & "-" & aTableValues3(0,k) & "-" & aTableValues2(0,j) & "-" & aTableValues3(1,k) & chr(13) & chr(10)
			
			if aTableValues(10,i)=aTableValues3(0,k) then 
                'Si el SBLID de CostsDetail es igual al BLDetailID de BLDetail, es decir el costo si corresponde al House
				if aTableValues2(0,j)=aTableValues3(1,k) then%>
					document.forma.H<%=j%>C<%=i%>.value = '<%=aTableValues3(2,k)%>';
					<%if aTableValues(15,i) <> 0 then%>
						document.forma.H<%=j%>C<%=i%>.disabled = 'false';
					<%end if%>
<%					if k < CountTableValues3 then
						k=k+1
					end if
				end if				
			end if
		next
	end if

%>


        } catch (err) {
            
            console.log('--------------------');
            console.log(err);
            console.log('--------------------');

        }
<%


next
  'Seteando el listado de HBLNumbers
  for i=0 to CountTableValues2%>
	document.getElementById("SBLIDN<%=i%>").innerHTML = "<%=aTableValues2(1,i)%>";
<%next
Set aTableValues = Nothing
Set aTableValues2 = Nothing
Set aTableValues3 = Nothing%>
	//Bloqueado los Costos que ya han sido pagados, esto va dentro del ciclo aTableValues
	/*	document.forma.N<%=i%>.disabled=true;
		document.forma.C<%=i%>.disabled=true;
		document.forma.V<%=i%>.disabled=true;
		document.forma.OV<%=i%>.disabled=true;
		document.forma.T<%=i%>.disabled=true;
		document.forma.TC<%=i%>.disabled=true;
		document.forma.CT<%=i%>.disabled=true;
		document.forma.PT<%=i%>.disabled=true;
		document.forma.PN<%=i%>.disabled=true;
		document.getElementById("DR<%=i%>").style.visibility = "hidden";
		document.getElementById("DP<%=i%>").style.visibility = "hidden";
		document.getElementById("DE<%=i%>").style.visibility = "hidden";
	*/
</script>