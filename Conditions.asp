<%@ Language=VBScript %>
<% 
'response.write Request("EXDBCountry") & "<br>"
'response.End 
%>

<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%


        Dim iResult, iEmpresa
        iResult = WsGetLogo(Request("EXDBCountry"), "TERRESTRE",  "12",  "",  "")
        iEmpresa = iResult(4)


'response.write "(" & Request("EXDBCountry") & ")(" & iEmpresa & ")"


    'dim Countries, Empresa, aTableValues5 
    '    Countries = Request("EXDBCountry")    
    'aTableValues5 = EmpresaParametros(Countries, "12", "TERRESTRE")           
    'Empresa = aTableValues5(5,0)
    
    'if (inStr(Countries, "LTF")) then
    '    Empresa = "LATIN FREIGHT"
    'else
    '    Empresa = "AIMAR GROUP"
    'end if
    
%>
<html>
<HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
body {
	margin: 0px;
}
.style4 {
	font-size:10px;
	color: #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;	
}
.style12 {font-size: 8px; font-family: Verdana, Arial, Helvetica, sans-serif; font-weight:normal;}
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
</head>
<body onLoad="JavaScript:self.focus();">
<h6 style = "text-align:center">TERMINOS Y CONDICIONES</h6>
<table width="641" cellpadding="2" cellspacing="0" align="center">
    <td align="left" class="style4" width="50%" valign="top">
	<span class="style12">
    Articulo 1. Los servicios llevados a cabo por <%=iEmpresa%> que en adelante se<br>
denominara EL TRANSPORTISTA ser&aacute;n exclusivamente regidos por estos<br>
T&eacute;rminos y Condiciones de la Carta de Porte, los cuales deben ser plenamente
aceptados en el momento de ordenar el servicio. Por este acto el cliente acepta
que estos T&eacute;rminos y Condiciones se apliquen a cualquier orden transmitida ya
sea verbalmente, por correo electr&oacute;nico por telefax, incluso y aun cuando no se haga ninguna referencia especifica estos T&eacute;rminos y Condiciones. Las
limitaciones de responsabilidad legal definidas en las estipulaciones de estos
T&eacute;rminos y Condiciones de la Carta de Porte se aplicaran axial mismo a toda
demanda de compensaci&oacute;n tambi&eacute;n como resultado de un acto il&iacute;cito.<br>
Articulo 2. RESPONSABILIDAD DE EL REMITENE (EMBARCADOR) EL<br>
REMITENTE ser&aacute; responsable de un embalaje apropiado, conforme a las<br>
exigencias propias de la naturaleza del Transporte.<br>
1. EL REMITENTE ser&aacute; responsable de informar con precisi&oacute;n a EL<br>
TRANSPORTISTA de su direcci&oacute;n, el lugar destinado a la entrega,<br>
la cantidad y tipo de bultos, el peso bruto, el contenido, el valor, el<br>
tiempo limite acordado para la entrega y el medio de transporte.<br>
2. EL REMITENTE deber&aacute; especificar a EL TRANSPORTISTA , la<br>
naturaleza de la mercader&iacute;a, su peso, su distribuci&oacute;n y si la<br>
mercanc&iacute;a es fr&aacute;gil. EL REMITENTE ser&aacute; responsable de un<br>
rotulado suficiente, y, si es necesario, de la numeraci&oacute;n de los<br>
bultos.<br>
3. EL REMITENTE ser&aacute; responsable de todos los gastos, multas,<br>
p&eacute;rdidas o da&ntilde;os causados por la falsedad, inexactitud, omisi&oacute;n o<br>
insuficiencia de los datos suministrados para la elaboraci&oacute;n de la<br>
Carta de Porte.<br>
4. EL REMITENTE ser&aacute; tambi&eacute;n responsable de cualquier cargo<br>
correspondiente a los camiones sufran alg&uacute;n retraso o retenci&oacute;n<br>
debido a cualquiera de las circunstancias arriba mencionadas.<br>
5. EL REMITENTE cumplir&aacute; con todas las leyes aplicables y las normas<br>
de los gobiernos de cualquier pa&iacute;s a, desde o a trav&eacute;s de cual<br>
pueda transportarse las mercader&iacute;as, incluyendo aquellas relativas<br>
a su embalaje, transporte o entrega y deber&aacute; proporcionar<br>
cualquier informaci&oacute;n y adjuntar a esta Carta de Porte tantos<br>
documentos como sean necesarios para cumplir con tales leyes y<br>
normas. EL TRANSPORTISTA no ser&aacute; responsable ante EL<br>
REMITENTE por cualquier p&eacute;rdida o gasto debido al no<br>
cumplimiento de dichos requisitos.<br>
6. Excepto cuando EL TRANSPORTISTA haya concedido cr&eacute;dito al<br>
destinatario sin el consentimiento por escrito de EL REMITENTE,<br>
este garantiza el pago de todos los cargos debidos por transporte y<br>
sus normas y leyes aplicables. (Incluyendo las leyes nacionales,<br>
reglamentos de los gobiernos, ordenes y requerimientos)<br>
Articulo 3. RESPONSABILIDAD<br>
1. EL TRANSPORTISTA ser&aacute; responsable de una ejecuci&oacute;n correcta y<br>
concienzuda de la orden.<br>
2. EL TRANSPORTISTA ser&aacute; responsable de cualquier da&ntilde;o imputable<br>
a <%=iEmpresa%> causado por el incumplimiento de las obligaciones del<br>
contrato.<br>
3. La responsabilidad de EL TRANSPORTISTA comenzara en el<br>
momento de transferir la mercanc&iacute;a a los empleados autorizados y<br>
terminara en el momento de la entrega de la mercanc&iacute;a al<br>
consignatario o su representante.<br>
4. Cualquier acci&oacute;n legal contra los empleados de EL TRANPORTISTA,<br>
ya sean estos permanentes o temporales por p&eacute;rdida o da&ntilde;o de la<br>
mercanc&iacute;a, solo ser&aacute; posible dentro de los l&iacute;mites considerados en<br>
los art&iacute;culos 3 y 4 siguientes.<br>
5. En caso de acci&oacute;n legal conjunta en contra de EL TRANPORTISTA<br>
y sus empleados ya sean fijos o temporales, la indemnizaci&oacute;n<br>
m&aacute;xima no exceder&aacute; los l&iacute;mites contemplados en el art&iacute;culo 4.<br>
Articulo 4. LIMITACIONES DE RESPONSABILIDAD<br>
1. La responsabilidad de EL TRANSPORTISTA ser&aacute; en cualquier caso<br>
limitada.<br>
2. No obstante, la compensaci&oacute;n no deber&aacute; exceder de: US$0.50 por<br>
Kilogramo de peso bruto, y hasta un maximo de US$500.00 con respecto a cualquier orden, incluyendo varios bultos. Cuando las cosas transportadas<br>
sean embaladas en contenedores, paletas y en general, en<br>
unidades selladas o precintadas, estas se consideran como una<br>
unidad de carga y deber&aacute;n de ser entregados por EL<br>
TRANSPORTISTA en el mismo estado en que se recibe el<br>
contenedor.<br>
Articulo 5. EXENCIONES<br>
1. EL TRANSPORTISTA no ser&aacute;, bajo ninguna circunstancia<br>
responsable de la perdida o da&ntilde;o de la mercanc&iacute;a, si estos han sido<br>
ocasionados por una mas de las siguientes circunstancias:<br>
a) La negligencia del cliente o su repres&eacute;ntate autorizado.<br>
b) Un embalaje, marcado o estilos incorrectos, o su total ausencia,<br>
siempre y cuando no haya sido EL TRANSPORTISTA quien haya<br>
realizado en embalaje, rotulado o estiba. EL TRANPORTISTA <br>
tampoco ser&aacute; responsable del embalaje de la mercanc&iacute;a, cuyo contenido no<br>
    </span>
    </td>
    <td align="left" class="style4" width="50%" valign="top">
	<span class="style12">
    pueda verificar.<br>
c) Guerra, rebeli&oacute;n, revoluci&oacute;n, insurrecci&oacute;n, poder usurpado o confiscaci&oacute;n,<br>
nacionalizaci&oacute;n o requisici&oacute;n por bajo las ordenes de cualquier gobierno o<br>
autoridad local publica.<br>
d) Da&ntilde;os causados por la energ&iacute;a nuclear<br>
e) Desastres naturales<br>
f) Casos de fuerza mayor<br>
g) Robo, asalto, secuestro<br>
h) Circunstancias en que EL TRANSPORTISTA no hubiese podido<br>
evitar y cuyas consecuencias no hubiera podido prever.<br>
2. EL TRANSPORTISTA no ser&aacute; bajo ninguna circunstancia,<br>
responsable, si la mercanc&iacute;a ha sido manejada por el cliente o su<br>
representante.<br>
3. EL TRANSPORTISTA no ser&aacute; responsable de las consecuencias de<br>
las operaciones de carga y descarga, que no haya realizado, a<br>
menos que se contrate el servicio espec&iacute;fico con el mismo.<br>
4. EL TRANSPORTISTA no ser&aacute; responsable del aumento del valor de<br>
la mercanc&iacute;a perdida o da&ntilde;ada.<br>
5. EL TRANPORTISTA no ser&aacute; responsable en relaci&oacute;n a cualquier<br>
perdida da&ntilde;o consecuente, tales como perdida de beneficios, lucro<br>
cesante, da&ntilde;o moral, perdida del cliente, reclamaciones por<br>
perdidas debidas a depreciaci&oacute;n y multas convencionales.<br>
Articulo 6. RESPONSABILIDAD EN CASO DE RETRASO<br>
El da&ntilde;o debido al retraso en la entrega no ser&aacute; indemnizado, excepto en el caso
en que la responsabilidad de EL TRANSPORTISTA a este respecto este<br>
debidamente estipulado por escrito por ambas partes. Adem&aacute;s, las condiciones
del Articulo 4 ?Limitaciones? y el Articulo 5 ?Exenciones? quedan expresamente
reservadas. En caso de indemnizaci&oacute;n como resultado de da&ntilde;os debidos a un
retraso en la entrega, la compensaci&oacute;n m&aacute;xima no deber&aacute; excedes la cantidad a
la que ascienda el flete.<br>
Articulo 7. RECLAMACIONES Y PRESCRIPCIONES<br>
1. Al momento de la entrega de las mercanc&iacute;as y antes de su retiro<br>
del lugar de descarga final, el consignatario o su representante<br>
deber&aacute;n revisar la cantidad, marcas y condiciones de los bultos. De<br>
existir irregularidades se dejara constancia escrita de estas<br>
irregularidades.<br>
2. La persona con derecho a la entrega debe reclamar por escrito al<br>
transporte<br>
I) en caso de da&ntilde;o evidente a las mercanc&iacute;as inmediatamente<br>
despu&eacute;s del descubrimiento del da&ntilde;o a m&aacute;s tardar dentro de las 24<br>
horas siguientes al recibo de la mercanc&iacute;a.<br>
II) de otros da&ntilde;os a las mercanc&iacute;as en un plazo m&aacute;ximo de tres<br>
d&iacute;as a partir de la recepci&oacute;n de las mismas.<br>
3. Cualquier derecho de indemnizaci&oacute;n contra EL<br>
TRANSPORTISTA prescribir&aacute; en el plazo de un (1) a&ntilde;o a partir de la fecha de<br>
llegada al punto de destino o desde la fecha la cual deber&iacute;a haber llegado la<br>
embarcaci&oacute;n o desde<br>
la fecha en que se suspendi&oacute; el transporte.<br>
4. Transcurrido el t&eacute;rmino expresado y pagado el porte, no se<br>
admitir&aacute; reclamaci&oacute;n alguna contra EL TRANSPORTISTA.<br>
5. Ning&uacute;n agente, empleado o representante de EL TRANSPORTISTA<br>
tiene autoridad para alterar, modificar o renunciar a cualquier<br>
disposici&oacute;n de estos t&eacute;rminos y condiciones. Las mercanc&iacute;as que<br>
viajan amparadas bajo esta orden de trabajo no son aseguradas<br>
por EL TRANSPORTISTA , a menos que medie una nota escrita del<br>
cliente/remitente/ o consignatario a tal efecto Registr&aacute;ndose este<br>
hecho, los bienes incluidos en esta gu&iacute;a son asegurados a trav&eacute;s de<br>
una p&oacute;liza abierta por la suma solicitada e indicada en la presente<br>
(limit&aacute;ndose al reintegro del valor de la mercanc&iacute;a perdida o<br>
da&ntilde;ada y seg&uacute;n los t&eacute;rminos de la p&oacute;liza abierta)<br>
6. Si desea asegurar su carga con <%=iEmpresa%> favor notificarlo por escrito de lo contrario
su embarque no contara con ninguna cobertura de seguro transporte.<br />
Con esto se exime a <%=iEmpresa%> de toda responsabilidad. <br>
7. Los T&eacute;rminos y Condiciones presentes quedan sujetos a cambio sin notificaci&oacute;n.<br>
    </span>
    </td>
   </tr>
</table>
</body>
</html>