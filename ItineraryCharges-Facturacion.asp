<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"
Dim GroupID, ObjectID, DocTyp, CountList9Values, aList9Values, QuerySelect, rs, conn, ConnMaster, FacID, FacType, FacStatus, i, BLType, esquema

GroupID = CheckNum(Request("GID"))
ObjectID = CheckNum(Request("ObjectID"))
DocTyp = CheckNum(Request("DocTyp"))
BLType = CheckNum(Request("BLType"))
esquema = Request("esquema")

%>

<HTML><HEAD><TITLE>AWB - Aimar - Administración</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">

<LINK REL="stylesheet" TYPE="text/css" HREF="img/estilos.css">

<style type="text/css">
<!--
body {
	margin: 0px;
}
.style8 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 10px;
	font-weight: bold;
	color: #999999;
}
.style10 {font-size: 10px; font-family: Verdana, Arial, Helvetica, sans-serif; }
.style11 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 10px;
}
.style4 {
	font-size:10px;
	color: #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
}
.style5 {font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 10px;
	color: #000000;
}
.style12 {font-family: Arial, Helvetica, sans-serif; font-size: 10px; color: #FFFFFF; }
.style13 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-weight: bold; color: #FFFFFF; }

.ids    {   border:0px;
            color:white;
            font-weight:normal;
            background:gray;
            font-size: 10px;
            width:auto; 
            }
            
.readonly { border:0px;
            background:silver;
            color:navy;
            font-size: 10px; 
            font-family: Verdana, Arial, Helvetica, sans-serif;  
            width:auto; }            
-->

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

.erpLab {
    color:white;background-color:gray;height:20px;display:block;padding:2px;
}

.erpFil {
    background-color:rgb(255,232,159);
}

</style>

<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="Javascript:self.focus();">

<%
    OpenConn Conn
	    'Obteniendo listado de Rubros
	    CountList9Values = -1

        '                       0       1           2       3       4       5       6           7           8           9           10          11      12      13  14  15      16      17  18      19
        QuerySelect = "SELECT UserID, ItemName, ItemID, Currency, Value, OverSold, Local, PrepaidCollect, ServiceID, ServiceName, InvoiceID, CalcInBL, DocType, '', '', InRO, ChargeID, '', '', Regimen FROM ChargeItems WHERE Expired=0 AND SBLID=" & ObjectID & " AND InterProviderType<>5 AND InterChargeType<>2 ORDER BY ChargeID, PrepaidCollect, Local, Currency, ServiceName, ItemName" 'InvoiceID Desc, 

                          '     0           1       2           3     4     5       6       7           8               9           10          11      12     13  14  15  16
        'QuerySelect = "Select ItemID, AgentTyp, CurrencyID, Local, Value, ItemName, Pos, ServiceID, ServiceName, PrepaidCollect, InvoiceID, CalcInBL, DocType, '', '', '', '' from ChargeItems where Expired=0 and AwbID=" & ObjectID & " and DocTyp=" & DocTyp & " order by AgentTyp"
        'response.write QuerySelect & "<br>"
	    Set rs = Conn.Execute(QuerySelect)	
	    If Not rs.EOF Then
		    aList9Values = rs.GetRows
		    CountList9Values = rs.RecordCount - 1
	    End If

    'Dim TextoStr, Texto64
    'TextoStr = "http://localhost:1010/ItineraryChargesPedidos.asp?GID=29CountryOrigen=CRTLACountriesFinalDes=PAAction=10Pedido_Msg=1%29%28%3Cdiv+style%3D%27display%3Ablock%27%3E4131+%3Cdiv+style%3D%27font%3A+15px+Arial%2C+sans-serif%3Bbackground-color%3ASILVER%3Bcolor%3ANAVY%3Bdisplay%3Ainline%27%3ETIEX046194%3C%2Fdiv%3E+PROCESO+CORRECTO%3C%2Fdiv%3E%29%28TIEX046194%29%28%7CTraeTextos%7CSendPedido%7CCicloTextos%7CC%3D0%7Cpedido_obj%7Cresult_obj.PEDIDOTE-EX-LL%7C4131%7Csingle_resp%3D%5B%7B%27ASIENTO%27%3AnullCOD_COMPANIA%27%3AnullCOD_PAIS%27%3AnullCODIGO_ERROR%27%3A%2799%27ESTADO%27%3A%27CORRECTO%27MENSAJE%27%3A%27PROCESO+CORRECTO%27PEDIDO%27%3A%2712%7CTE-EX-LL%7C4131%27PEDIDO_EXACTUS%27%3A%27TIEX046194%27PEDIDO_RESP%27%3AnullCODIGO%27%3AnullDESCRIPCION%27%3Anull%7D%5D%7Cquery%3DUPDATE+exactus_pedidos+SET+estado+%3D+3%2C+json_cargo_system+%3D+%27%5B%7B%27EMPRESA%27%3A%27CRTLA%27MOVIMIENTO%27%3A%27EXPORT%27TIPO_CARGA%27%3A%27CONSOLIDADO%27PEDIDO%27%3A%2712%7CTE-EX-LL%7C4131%27PEDIDO_ERP%27%3A%27TIEX046194%27BODEGA%27%3A%27BOSE%27CLIENTE%27%3A%2701%7C60590%27RUTA%27%3A%27ND%27ZONA%27%3A%27ND%27PAIS%27%3A%27CRTLA%27NIVEL_PRECIO%27%3A%2708%7CCRC%27MONEDA%27%3A%2708%7CCRC%27VENDEDOR%27%3A%27ND%27COBRADOR%27%3A%27ND%27CONDICION_PAGO%27%3A%2700%27ESTADO%27%3A%27N%27FECHA_PEDIDO%27%3A%2705%2F27%2F2022%27FECHA_PROMETIDA%27%3A%2705%2F27%2F2022%27FECHA_PROX_EMBARQU%27%3A%2705%2F27%2F2022%27FECHA_ULT_EMBARQUE%27%3A%2705%2F27%2F2022%27FECHA_ULT_CANCELAC%27%3A%2705%2F27%2F2022%27ORDEN_COMPRA%27%3A%27%27FECHA_ORDEN%27%3A%2705%2F27%2F2022%27TARJETA_CREDITO%27%3A%27%27EMBARCAR_A%27%3A%270%27DIREC_EMBARQUE%27%3A%27ND%27DIRECCION_FACTURA%27%3A%270%27RUBRO1%27%3A%27%27RUBRO2%27%3A%27%27RUBRO3%27%3A%27%27RUBRO4%27%3A%27CCRTLA202128-0001-100707%27RUBRO5%27%3A%27%27OBSERVACIONES%27%3A%27Liberar+Rubro+526113+Pedido+ERP+TIEX046194+Cliente+60590%27COMENTARIO_CXC%27%3A%27CCRTLA202128-0001-100707%27TOTAL_MERCADERIA%27%3A%2775%27MONTO_ANTICIPO%27%3A%270%27MONTO_FLETE%27%3A%270%27MONTO_SEGURO%27%3A%270%27MONTO_DOCUMENTACIO%27%3A%270%27TIPO_DESCUENTO1%27%3A%27P%27TIPO_DESCUENTO2%27%3A%27P%27MONTO_DESCUENTO1%27%3A%270%27MONTO_DESCUENTO2%27%3A%270%27PORC_DESCUENTO1%27%3A%270%27PORC_DESCUENTO2%27%3A%270%27TOTAL_IMPUESTO1%27%3A%270%27TOTAL_IMPUESTO2%27%3A%270%27TOTAL_A_FACTURAR%27%3A%2775%27PORC_COMI_VENDEDOR%27%3A%270%27PORC_COMI_COBRADOR%27%3A%270%27TOTAL_CANCELADO%27%3A%270%27TOTAL_UNIDADES%27%3A%271%27IMPRESO%27%3A%27N%27USUARIO%27%3A%27csystem%27FECHA_HORA%27%3A%27%27DESCUENTO_VOLUMEN%27%3A%270%27TIPO_PEDIDO%27%3A%27N%27MONEDA_PEDIDO%27%3A%2708%7CCRC%27CLASE_PEDIDO%27%3A%27N%27TIPO_DOC_CXC%27%3A%27FAC%27SUBTIPO_DOC_CXC%27%3A%270%27VERSION_NP%27%3A%271%27AUTORIZADO%27%3A%27N%27DOC_A_GENERAR%27%3A%27F%27CLIENTE_ORIGEN%27%3A%270%27CLIENTE_CORPORAC%27%3A%270%27CLIENTE_DIRECCION%27%3A%270%27f%27%3A%27N%27DESCUENTO_CASCADA%27%3A%27N%27CONTRATO%27%3A%27%27PORC_INTCTE%27%3A%270%27NOTEEXISTSFLAG%27%3A%270%27RECORDDATE%27%3A%27%27CREATEDBY%27%3A%27csystem%27CREATEDATE%27%3A%27%27UPDATEDBY%27%3A%27csystem%27TIPO_CAMBIO%27%3A%270%27FIJAR_TIPO_CAMBIO%27%3A%27N%27ROWPOINTER%27%3A%270%27ORIGEN_PEDIDO%27%3A%27F%27DESC_DIREC_EMBARQUE%27%3A%27%27DIVISION_GEOGRAFICA1%27%3A%27%27DIVISION_GEOGRAFICA2%27%3A%27%27BASE_IMPUESTO1%27%3A%270%27BASE_IMPUESTO2%27%3A%270%27NOMBRE_CLIENTE%27%3A%27P%27FECHA_PROYECTADA%27%3A%27%27FECHA_APROBACION%27%3A%27%27TIPO_DOCUMENTO%27%3A%27P%27VERSION_COTIZACION%27%3A%27%27RAZON_CANCELA_COTI%27%3A%27%27DES_CANCELA_COTI%27%3A%27%27CAMBIOS_COTI%27%3A%27%27COTIZACION_PADRE%27%3A%27%27TASA_IMPOSITIVA%27%3A%27%27TASA_IMPOSITIVA_PORC%27%3A%270%27TASA_CREE1%27%3A%27%27TASA_CREE1_PORC%27%3A%270%27TASA_CREE2%27%3A%27%27TASA_CREE2_PORC%27%3A%270%27TASA_GAN_OCASIONAL_PORC%27%3A%270%27CONTRATO_AC%27%3A%27%27TIPO_CONTRATO_AC%27%3A%27%27PERIODICIDAD_CONTRATO_AC%27%3A%27%27FECHA_CONTRATO_AC%27%3A%27%27FECHA_INICIO_CONTRATO_AC%27%3A%27%27FECHA_PROXFAC_CONTRATO_AC%27%3A%27%27FECHA_FINFAC_CONTRATO_AC%27%3A%27%27FECHA_ULTAUMENTO_CONTRATO_AC%27%3A%27%27FECHA_PROXFACSIST_CONTRATO_AC%27%3A%27%27DIFERIDO_CONTRATO_AC%27%3A%27%27TOTAL_CONTRATO_AC%27%3A%27%27CONTRATO_REVENTA%27%3A%27N%27USR_NO_Af%27%3A%27%27FECHA_NO_Af%27%3A%27%27RAZON_DESAf%27%3A%27%27MODULO%27%3A%27%27CORREOS_ENVIO%27%3A%27%27CONTRATO_VIGENCIA_DESDE%27%3A%27%27CONTRATO_VIGENCIA_HASTA%27%3A%27%27USO_CFDI%27%3A%27%27FORMA_PAGO%27%3A%27%27CLAVE_REFERENCIA_DE%27%3A%27%27FECHA_REFERENCIA_DE%27%3A%27%27U_ENVIADO_TLA%27%3A%270%27TIPO_OPERACION%27%3A%27%27INCOTERMS%27%3A%27%27U_AD_WM_NUMERO_VENDEDOR%27%3A%27%27U_AD_WM_ENVIAR_GLN%27%3A%27%27U_AD_WM_NUMERO_RECEPCION%27%3A%27%27U_AD_WM_NUMERO_RECLAMO%27%3A%27%27U_AD_WM_FECHA_RECLAMO%27%3A%27%27U_AD_PC_NUMERO_VENDEDOR%27%3A%27%27U_AD_PC_ENVIAR_GLN%27%3A%27%27U_AD_GS_NUMERO_VENDEDOR%27%3A%27%27U_AD_GS_ENVIAR_GLN%27%3A%27%27U_AD_GS_NUMERO_RECEPCION%27%3A%27%27U_AD_GS_FECHA_RECEPCION%27%3A%27%27U_AD_AM_NUMERO_PROVEEDOR%27%3A%27%27U_AD_AM_ENVIAR_GLN%27%3A%27%27U_AD_AM_NUMERO_RECEPCION%27%3A%27%27U_AD_AM_NUMERO_RECLAMO%27%3A%27%27U_AD_AM_FECHA_RECLAMO%27%3A%27%27U_AD_AM_FECHA_RECEPCION%27%3A%27%27U_AD_CC_REMISION%27%3A%27%27U_AD_CC_FECHA_CONSUMO%27%3A%27%27U_AD_CC_HOJA_ENTRADA%27%3A%27%27U_IVA_CATEGORIA%27%3A%270%27ACTIVIDAD_COMERCIAL%27%3A%27602001%27MONTO_OTRO_CARGO%27%3A%2775%27CODIGO_REFERENCIA_DE%27%3A%27%27TIPO_REFERENCIA_DE%27%3A%27%27TIENE_RELACIONADOS%27%3A%27%27ES_FACTURA_REEMPLAZO%27%3A%27N%27FACTURA_ORIGINAL_REEMPLAZO%27%3A%27%27CONSECUTIVO_FTC%27%3A%27%27NUMERO_FTC%27%3A%27%27NIT_TRANSPORTADOR%27%3A%27%27NUM_OC_EXENTA%27%3A%27%27NUM_CONS_REG_EXO%27%3A%27%27NUM_IRSEDE_AGR_GAN%27%3A%27%27U_AD_GS_NUMERO_ORDEN%27%3A%270%27U_AD_GS_FECHA_RECLAMO%27%3A%27%27U_AD_GS_NUMERO_RECLAMO%27%3A%270%27U_AD_GS_FECHA_ORDEN%27%3A%27%27U_AD_WM_NUMERO_ORDEN%27%3A%270%27U_AD_WM_FECHA_ORDEN%27%3A%27%27U_AD_PM_ENVIAR_GLN%27%3A%270%27U_AD_MS_NUMERO_VENDEDOR%27%3A%270%27U_AD_MS_ENVIAR_GLN%27%3A%270%27U_AD_MS_NUMERO_RECEPCION%27%3A%270%27U_AD_MS_NUMERO_RECLAMO%27%3A%270%27U_AD_MS_FECHA_RECLAMO%27%3A%27%27TIPO_PAGO%27%3A%27%27TIPO_DESCUENTO_GLOBAL%27%3A%27%27TIPO_FACTURA%27%3A%27%27U_FECHA_DUA_AA%27%3A%27%27U_PAIS_ORIGEN_AA%27%3A%27%27U_DIAS_AF%27%3A%27%27U_EQUIPO_AI%27%3A%27%27U_CIF_AF%27%3A%27%27U_PESO_AF%27%3A%27%27U_IMPUESTOS_AF%27%3A%27%27U_VOLUMEN_AF%27%3A%27%27U_BULTOS_DUA_AF%27%3A%27%27U_TC_AF%27%3A%27%27U_BL_AF%27%3A%27%27U_MOVIMIENTO_AF%27%3A%27%27U_LIQUIDACION_TI%27%3A%27CCRTLA202128-0001-100707%27U_DUA_AF%27%3A%27%27U_SERVICIO_TI%27%3A%27LTL%27U_RED%27%3A%27%27U_AGENTE_TI%27%3A%27%27U_TARIFA_AF%27%3A%27%27U_TRAMITE_AA%27%3A%27%27U_ADUANA_AA%27%3A%27%27U_ASOCIAR_A_PEDIDO_AA%27%3A%27%27U_LIQUIDACION_AA%27%3A%27%27U_CLIENTE_TI%27%3A%2701%7C23216%27U_TRANSPORTISTA_TI%27%3A%27GRUPO+TLA+S.A.+COSTA+RICA%27U_REC_ANT_ORIGEN_AA%27%3A%27%27U_FECHA_MOVIMIENTO_AF%27%3A%27%27U_CLIENTE_AF%27%3A%27%27U_AGENCIA%27%3A%27%27U_AGENCIA_NOM%27%3A%27%27U_AGENCIA_NIT%27%3A%27%27U_CONSIG%27%3A%27%27U_CONSIG_NIT%27%3A%27%27NUMERO_REGISTRO_IVA%27%3A%27%27U_COPIA_PAIS%27%3A%27%27LINEAS%27%3A%5B%7B%27PEDIDO%27%3A%2712%7CTE-EX-LL%7C4131%27PEDIDO_LINEA%27%3A%271%27ARTICULO%27%3A%2709%7CTE15-EX-LL-42%27BODEGA%27%3A%27BOSE%27ESTADO%27%3A%27N%27FECHA_ENTREGA%27%3A%2705%2F27%2F2022%27LINEA_USUARIO%27%3A%270%27PRECIO_UNITARIO%27%3A%2775%27CANTIDAD_PEDIDA%27%3A%271%27CANTIDAD_A_FACTURA%27%3A%271%27CANTIDAD_FACTURADA%27%3A%270%27CANTIDAD_RESERVADA%27%3A%270%27CANTIDAD_BONIFICAD%27%3A%270%27CANTIDAD_CANCELADA%27%3A%270%27TIPO_DESCUENTO%27%3A%27P%27MONTO_DESCUENTO%27%3A%270%27PORC_DESCUENTO%27%3A%270%27DESCRIPCION%27%3A%27%27COMENTARIO%27%3A%27%27PEDIDO_LINEA_BONIF%27%3A%27%27LOTE%27%3A%27%27LOCALIZACION%27%3A%27%27UNIDAD_DISTRIBUCIO%27%3A%27%27FECHA_PROMETIDA%27%3A%2705%2F27%2F2022%27LINEA_ORDEN_COMPRA%27%3A%271%27NOTEEXISTSFLAG%27%3A%270%27RECORDDATE%27%3A%27%27CREATEDBY%27%3A%27csystem%27CREATEDATE%27%3A%27%27UPDATEDBY%27%3A%27csystem%27ROWPOINTER%27%3A%27%27PROYECTO%27%3A%27N%27FASE%27%3A%27N%27CENTRO_COSTO%27%3A%27%27CUENTA_CONTABLE%27%3A%27%27RAZON_PERDIDA%27%3A%27N%27TIPO_DESC%27%3A%270%27TIPO_IMPUESTO1%27%3A%27N%27TIPO_TARIFA1%27%3A%27N%27TIPO_IMPUESTO2%27%3A%27N%27TIPO_TARIFA2%27%3A%270%27PORC_EXONERACION%27%3A%270%27MONTO_EXONERACION%27%3A%270%27PORC_IMPUESTO1%27%3A%270%27PORC_IMPUESTO2%27%3A%270%27ES_OTRO_CARGO%27%3A%27N%27ES_CANASTA_BASICA%27%3A%27N%27PORC_EXONERACION2%27%3A%270%27MONTO_EXONERACION2%27%3A%270%27PORC_IMP1_BASE%27%3A%270%27PORC_IMP2_BASE%27%3A%270%27TIPO_DESCUENTO_LINEA%27%3A%27%27%7D%5D%7D%5D%27%2C+json_exactus+%3D+%27%5B%7B%27ASIENTO%27%3AnullCOD_COMPANIA%27%3AnullCOD_PAIS%27%3AnullCODIGO_ERROR%27%3A%2799%27ESTADO%27%3A%27CORRECTO%27MENSAJE%27%3A%27PROCESO+CORRECTO%27PEDIDO%27%3A%2712%7CTE-EX-LL%7C4131%27PEDIDO_EXACTUS%27%3A%27TIEX046194%27PEDIDO_RESP%27%3AnullCODIGO%27%3AnullDESCRIPCION%27%3Anull%7D%5D%27%2C+pedido+%3D+%27%3Cdiv+style%3D%27display%3Ablock%27%3E4131+%3Cdiv+style%3D%27font%3A+15px+Arial%2C+sans-serif%3Bbackground-color%3ASILVER%3Bcolor%3ANAVY%3Bdisplay%3Ainline%27%3ETIEX046194%3C%2Fdiv%3E+PROCESO+CORRECTO%3C%2Fdiv%3E%27%2C+codigo_consecutivo+%3D+%2712%7CTE-EX-LL%7C4131%27%2C+pedido_erp%3D%27TIEX046194%27%2C+esquema+%3D+%27TRANSIT%27+WHERE+id_pedido+%3D+4131%7CPostgres_.EjecutaQuery%3D1%29%28esquema=TRANSITPedidoCliente=60590PedidoRubro=526113Pedido_Erp=TIEX046194OID=100707CountryExactus=CRTLAHBLNumber=CCRTLA202128-0001-100707Movimiento=EXPORTSelectBodegas=BOSEActividadComercial=602001CondicionPago=00ObservacionesErp=Liberar+Rubro+526113+Pedido+ERP+TIEX046194+Cliente+60590"
    'Texto64 = Base64Encode2(TextoStr)
    'response.write Texto64 & "<hr>"


    openConnBAW Conn
    OpenConn2 ConnMaster

    for i=0 to CountList9Values

	    FacID = CheckNum(aList9Values(10,i))
        FacType = CheckNum(aList9Values(12,i))
        FacStatus = 0

        if FacID<>0 then
	        Select case FacType
            case 1
                set rs = Conn.Execute("select tfa_serie, tfa_correlativo, tfa_ted_id from tbl_facturacion where tfa_id=" & FacID)
                If Not rs.EOF Then
			        aList9Values(13,i) = "FC-" & rs(0) & "-" & rs(1)
                    FacStatus = CheckNum(rs(2))
                end if
		        CloseOBJ rs

            case 4
                set rs = Conn.Execute("select tnd_serie, tnd_correlativo, tnd_ted_id from tbl_nota_debito where tnd_id=" & FacID)
			        aList9Values(13,i) = "ND-" & rs(0) & "-" & rs(1)
                    FacStatus = CheckNum(rs(2))
		        CloseOBJ rs

            case 9,10 'recibido por pedido exactus
                QuerySelect = "SELECT DISTINCT COALESCE(a.pedido_erp,''), COALESCE(a.estado,0), COALESCE(b.fc_numero,''), COALESCE(b.fc_estado,0), COALESCE(b.fc_saldo,0), COALESCE(c.nc_numero,'')  FROM exactus_pedidos a LEFT JOIN exactus_pedidos_fc b ON a.id_pedido = b.id_pedido  LEFT JOIN exactus_pedidos_nc c ON a.id_pedido = c.id_pedido WHERE a.id_pedido = " & FacID & " "
                'response.write QuerySelect & "<br>"
                set rs = ConnMaster.Execute(QuerySelect)
                If Not rs.EOF Then
	    
                    if rs(0) <> "" then
                        'aList9Values(13,i) = "PE-" & rs(0) 
                        aList9Values(13,i) = FacID & " - " & rs(0) 
                        FacStatus = 90 'enviada
                    end if

                    if rs(2) <> "" then
                        'aList9Values(13,i) = "FC-" & rs(2) 
                        aList9Values(13,i) = FacID & " - " & rs(2) 
                        FacStatus = 91 'facturada
                    end if

                    'if  rs(5) <> "" then 'nunca entra aca
                    '    aList9Values(13,i) = "NC-" & rs(5) 
                    '    FacStatus = 92 'anulada
                    'end if

                end if		
		        CloseOBJ rs

            end Select

        End If


        'Indicando el Estado de Pago de la Factura/ND
        select Case FacStatus
        case 2
            aList9Values(14,i) = "<font color=blue>ABONADO</font>"
        case 4
            aList9Values(14,i) = "<font color=blue>PAGADO</font>"

        case 90 '2021-08-06
            aList9Values(14,i) = "<font color=blue>ENVIADO</font>"

        case 91 '2021-08-16
            aList9Values(14,i) = "<font color=blue>FACTURADO</font>"

        case 92 '2021-08-06
            aList9Values(14,i) = "<font color=blue>CANCELADO</font>"

        case Else
            aList9Values(14,i) = "<font color=red>PENDIENTE</font>"
        End Select


  
        'si funciona pero aun no esta autorizado mostrarlo solo 1237
        'if Session("OperatorID") = "1237" then 

            aList9Values(17,i) = "<img src='img/glyphicons_192_circle_remove1.png'>"

            QuerySelect = "SELECT a.codigo, COALESCE(eh_erp_codigo,''), COALESCE(eh_estado,0) " & _	
	"FROM vw_rubros_combinaciones a " & _
	"LEFT JOIN exactus_homologaciones ON codigo = eh_codigo AND eh_erp_categoria = '06' AND eh_estado = 1 AND eh_erp_esquema = '" & esquema & "' " & _     
    "WHERE a.id_servicio = " & aList9Values(8,i) & " AND a.id_rubro = " & aList9Values(2,i) & _ 
    " AND a.d1 = '" & Iif(DocTyp = "1","IMPORT","EXPORT") & "' AND (a.descripcion ILIKE '%terrestre%' OR a.serv ILIKE '%terrestre%') " & _
    " AND a.d2 = '" & Iif(BLType = 0,"LL",Iif(BLType = 1,"LE","LC")) & "' AND a.d3 = '" & aList9Values(19,i) & "' "	    
         
            'response.write QuerySelect & "<br>"
            set rs = ConnMaster.Execute(QuerySelect)
            If Not rs.EOF Then
                aList9Values(18,i) = rs(0) '"AE" & aList9Values(7,i) & "-IM-A-" & aList9Values(0,i)
       	        aList9Values(17,i) = "<font color=" & Iif(rs(2) = "1","green","white") & ">" & rs(1) & "</font>"
            end if
	        CloseOBJ rs
        
        'end if

        'response.write "(" & aList9Values(15,i) & ")(" & aList9Values(13,i) & ")(" & aList9Values(14,i) & ")" & "<br>"

    next



    Function IntLoc(num) 
		Select Case num 
		Case 0
			IntLoc = "INT"
		Case 1
			IntLoc = "LOC"
		Case Else
			IntLoc = "---"
		End Select
	End Function

    Function PrepColl(num) 
		Select Case num 
		Case 0
			PrepColl = "PREP"
		Case 1
			PrepColl = "COLL"
		Case Else
			PrepColl = "---"
		End Select
	End Function
%>



    <center><h3 class="menu" ><font color=white><%=ObjectID & " - " & Iif(Request("HAWBNumber") = "", "HBLNumber : " & Request("AWBNumber"), "HBLNumber : " & Request("HAWBNumber")) %>  &nbsp; - &nbsp; <%=Iif(DocTyp=0,"EXPORT","IMPORT")%></font></h3></center>

    <table width="80%" border="0">
        <tr>

            <%'if Session("OperatorID") = "1237" then %>
		    <td align="center" class="style4">
		        <font class="style8">Articulo</font>
            </td>

		    <td align="left" class="style4">
		        <font class="style8">Homologado</font>
            </td>          
            <%'end if %>

		    <td align="center" class="style4">
                <font class="style8">Servicio</font>
            </td>
		    <td align="center" class="style4">
                <font class="style8">Rubro</font>
            </td>
		    <td align="center" class="style4">
                <font class="style8">Moneda</font>
            </td>
		    <td align="center" class="style4">
                <font class="style8">Monto</font>
            </td>
		    <td align="center" class="style4">
		        <font class="style8">Int/Loc</font>
            </td>
		    <td align="center" class="style4">
		        <font class="style8">CC/PP</font>
            </td>


		    <td align="center" class="style4">
		        <font class="style8">Pedido / Factura / ND</font>
            </td>
		    <td align="center" class="style4">
		        <font class="style8">Estado</font>
            </td>


        </tr>


<%
if CountList9Values>=0 then

	for i=0 to CountList9Values

%>

    <tr bgcolor="">

           
        <%'if Session("OperatorID") = "1237" then %>
	        <td align="right" class="style4" nowrap><%=aList9Values(18,i)%></td>

	        <td align="right" class="style4" style="background-color:white;text-align:center">
		        <%=aList9Values(17,i)%>		
            </td>    
        <%'end if %>

		<td align="right" class="style4"> 
			<input type="text" size="18" class="style10" value="<%=aList9Values(8,i) & " - " & aList9Values(9,i)%>" id="SVNO1" readonly>
		</td>

		<td align="right" class="style4">
			<input type="text" class="style10" value="<%=aList9Values(2,i) & " - " & aList9Values(1,i)%>" size="25" readonly>		
        </td>
                      				
		<td align="right" class="style4">
			<input type="text" class="style10" value="<%=aList9Values(3,i)%>" size="5" readonly>	 <!-- moneda -->			
        </td>
                                            				
		<td align="right" class="style4">
			<input type="text" class="style10" value="<%=aList9Values(4,i)%>" size="20" readonly>	 <!-- valor -->	
        </td>
                      				
		<td align="right" class="style4">
			<input type="text" size="5" class="style10" value="<%=IntLoc(CheckNum(aList9Values(6,i)))%>"  readonly>		
        </td>
               
		<td align="right" class="style4">
			<input type="text" size="5" class="style10" value="<%=PrepColl(CheckNum(aList9Values(7,i)))%>"  readonly>		
        </td>

        <td align="right" class="style4">
			<input type="text" class="style10" value="<%=aList9Values(13,i)%>" size="25" readonly>		
        </td>

		<td align="right" class="style4" style="background-color:white">
			<%=aList9Values(14,i)%>		
        </td>




	</tr>

<%
    next

end if
%>
    </table>

</BODY>
</HTML>
<%
    'TextoStr = Base64Decode2(Texto64)
    'response.write TextoStr & "<hr>"
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>


