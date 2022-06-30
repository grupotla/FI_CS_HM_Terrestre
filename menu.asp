<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
if Session("OperatorID") <> 0 then
Checking "0|1|2"
%>
<HTML><HEAD><TITLE>Sistema Terrestre</TITLE></HEAD>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language=javascript src="img/config.js"></SCRIPT>
<!--<SCRIPT src="img/mainLib.js" language="JavaScript"></SCRIPT>-->
<script>
var Menu1 = "";
var MenuColor1 = "";
var Menu2 = "";
var MenuColor2 = "";
var TitleMenu = "";

  var isNav = false;
  var isIE = false;
  var col1 = "";
  var styleObj = "";

  if (parseInt(navigator.appVersion) >= 4) {
     if(navigator.appName == "Netscape" ) {
		isNav = false;
     }
     else {
		isIE = true;
		col1 = "all.";
		styleObj = ".style";
     }	
  } else {
		styleObj = ".style";
  }
  //alert(navigator.appName);
  //alert(parseInt(navigator.appVersion));
  //alert(col1);
  //alert(styleObj);
  //alert(isIE);

  function getObject( obj ){
  var theObj;
  var x = typeof obj;
	//alert(obj);
	//alert(x);
	//alert(eval("document." + col1 + obj + styleObj));
	
	if (x == "string" ){
		if (parseInt(navigator.appVersion) < 5) {
			theObj = eval("document." + col1 + obj + styleObj );
		} else {
			theObj = eval(document.getElementById(obj).style);
		}
	}
	else
	{
		theObj = obj;
	}
	//alert(theObj);
	return theObj;
	
  }

	function showMenu( Main, MainColor, Level ){
	var ColorMenu;
	var ColorMain;
	var Menu;
	var MenuColor;	
	if (Level == 1) {
		 Menu = Menu1;
		 MenuColor = MenuColor1;
		 ColorMenu = "#996600";
		 ColorMain = "#660000";
		 if (Menu2 != "") {
		 		var objectMenu2 = getObject( Menu2 );
				var objectMenuColor2 = getObject( MenuColor2 );
				objectMenu2.visibility = "hidden";
  		  		objectMenuColor2.background = "#996600";
				if (TitleMenu != "") {
					 var objectTitleMenu = getObject( TitleMenu );
					 objectTitleMenu.visibility = "hidden";
				}
		 }
	};
	getObject( Main );
	var objectMain = getObject( Main );
	var objectMainColor = getObject( MainColor );
	if (Menu == "") {
		 objectMain.visibility = "visible";
		 objectMainColor.background = ColorMenu;
		 Menu = Main;
		 MenuColor = MainColor;		 
	} else {
  		if (Menu != Main) {
  		 	var objectMenu = getObject( Menu );
				var objectMenuColor = getObject( MenuColor );
				objectMenu.visibility = "hidden";
  		  objectMain.visibility = "visible";
				objectMenuColor.background = ColorMain;
  		  objectMainColor.background = ColorMenu;
				Menu = Main;				
				MenuColor = MainColor;
			} else {
						 if (Menu == Main) {
						 		var objectMenu = getObject( Menu );
								var objectMenuColor = getObject( MenuColor );
								objectMenu.visibility = "hidden";
								objectMenuColor.background = ColorMain;								
								Menu = "";
								MenuColor = "";
							}
			}
	}
	if (Level == 1) {
		 Menu1 = Menu;
		 MenuColor1 = MenuColor;
	};
	if (Level == 2) {
		 Menu2 = Menu;
		 MenuColor2 = MenuColor;
	};
};


/*In your javascript*/
var prevItem = null;
function activateItem(t) {
    if (prevItem != null) {
        prevItem.className = prevItem.className.replace(/activeItem/, "");
    }
    t.className += " activeItem";
    prevItem = t;
}

</script>
<LINK REL="stylesheet" type="text/css" HREF="img/estilos.css">
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">

<style>
.activeItem{
   background-color:rgb(102,0,0); /*make some difference for the active item here */
   color:orange;
   padding:3px;
}
</style>
<TABLE cellSpacing=0 cellPadding=0 width="100%" class=menu>
	<TR><TD colspan="10">
	<TABLE cellSpacing=0 cellPadding=2 width="100%" class=menu border=0>
	<TR>
		<TD class=titlea vAlign=center align=left width="20%">&nbsp;&nbsp;Sistema Terrestre<BR>&nbsp;&nbsp;<SPAN class=label><FONT color=#ffffff><%=Session("Date")%><!--Administrador BL 1.0--></FONT></SPAN></TD>
        <TD class=titlea vAlign=center align=right width="60%">


<iframe frameborder="0" height="30px" width="50%" src="ColgateRefresh.asp" scrolling="no">
<p>Su navegador no es compatible con iframes</p>
</iframe>

        </TD>       
        <TD class=titlea vAlign=center align=right width="20%">&nbsp;&nbsp;<SPAN class=label>        
        <FONT color=#ffffff>
        
        
        <%=Session("OperatorName")%> :: 
        
        <% 
        dim tempo, lentempo, i
        tempo = Split(Replace(Replace(Replace(Session("Countries"),"(",""),")",""),"'",""), ",")
        lentempo = ubound(tempo)
        response.write "<select style='background-color:rgb(102,0,0);color:white;border:0px;font-size:9px;'><option> Oficinas </option>" 
        for i = 0 to lentempo
            response.write "<option>" & tempo(i) & "</option>"
        next
        response.write "</select>" 
        %>

        <%'=Replace(Replace(Replace(Session("Countries"),"(",""),")",""),"'","")%>        
        
        <br />


        <script>document.write("Direccion IP: "+ip+"")</script></FONT></SPAN></TD>
	    <!-- <TD valign=right align=middle width="10%"><!--<IMG src="img/logo.gif">--></TD>
	</TR>
	</TABLE>
	<TABLE cellSpacing=0 cellPadding=0 width="100%" border="1">
	<TR>
		<TD class=border><IMG height=1 src="img/transparente.gif" width=1></TD>
	</TR>
	</TABLE>
	</TD></TR>
	<TR>
    <TD class=inactiveMain valign=center align=left>
		<table>
		<TR>
		<TD id=TDSetup0 class=menu vAlign=center align=left>&nbsp;<A class=activeMain href="javascript:showMenu('mainSetup0', 'TDSetup0', 1);">Tr&aacute;nsito</A>&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD id=TDSetup10 class=menu vAlign=center align=left>&nbsp;<A class=activeMain href="javascript:showMenu('mainSetup10', 'TDSetup10', 1)">Local</A>&nbsp;</TD>
		<%if Session("OperatorLevel")=0 or Session("OperatorLevel")=1 then%> 	
		<!--<TD class=separator vAlign=center align=left>|</TD>
		<TD id=TDSetup3 class=menu vAlign=center align=left>&nbsp;<A class=activeMain href="javascript:showMenu('mainSetup3', 'TDSetup3', 1)">DTI</A>&nbsp;</TD>-->
		<TD class=separator vAlign=center align=left>|</TD>
		<TD id=TDSetup1 class=menu vAlign=center align=left>&nbsp;<A class=activeMain href="http://10.10.1.20/catalogo_admin/login.php" target="_blank">Exportadores</A>&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD id=TDSetup8 class=menu vAlign=center align=left>&nbsp;<A class=activeMain href="http://10.10.1.20/catalogo_admin/login.php" target="_blank">Consignatarios&nbsp;/&nbsp;Embarcadores</A>&nbsp;</TD>
		<!--
		<TD class=separator vAlign=center align=left>|</TD>
		<TD id=TDSetup1 class=menu vAlign=center align=left>&nbsp;<A class=activeMain href="javascript:showMenu('mainSetup1', 'TDSetup1', 1)">Exportadores</A>&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD id=TDSetup8 class=menu vAlign=center align=left>&nbsp;<A class=activeMain href="javascript:showMenu('mainSetup8', 'TDSetup8', 1)">Consignatarios&nbsp;/&nbsp;Embarcadores</A>&nbsp;</TD>
		-->
		<TD class=separator vAlign=center align=left>|</TD>
		<TD id=TDSetup5 class=menu vAlign=center align=left>&nbsp;<A class=activeMain href="javascript:showMenu('mainSetup5', 'TDSetup5', 1)">Aduanas</A>&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD id=TDSetup4 class=menu vAlign=center align=left>&nbsp;<A class=activeMain href="javascript:showMenu('mainSetup4', 'TDSetup4', 1)">Transportes</A>&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD id=TDSetup7 class=menu vAlign=center align=left>&nbsp;<A class=activeMain href="javascript:showMenu('mainSetup7', 'TDSetup7', 1)">Productos</A>&nbsp;</TD>

		<TD class=separator vAlign=center align=left>|</TD>
		<TD id=TDSetup12 class=menu vAlign=center align=left>&nbsp;<A class=activeMain href="javascript:showMenu('mainSetup12', 'TDSetup12', 1)">Fechas</A>&nbsp;</TD>
		
        <%end if
		  if Session("OperatorLevel")=0 then%> 	
        <TD class=separator vAlign=center align=left>|</TD>
		<TD id=TDSetup11 vAlign=center align=left>&nbsp;<A class=activeMain href="javascript:showMenu('mainSetup11', 'TDSetup11', 1)">Logos</A>&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD id=TDSetup9 class=menu vAlign=center align=left>&nbsp;<A class=activeMain href="javascript:showMenu('mainSetup9', 'TDSetup9', 1)">Configuración&nbsp;</A></TD>
		<%end if%>
		</TR>
		</table>
		</TD>
		<TD class=menu vAlign=center width="100%" align=left>&nbsp;&nbsp;</TD>
		<TD width="1" class=separator vAlign=middle align=right>|</TD>
		<TD  width="50" id=MisDatos class=menu vAlign=center align=middle>&nbsp;&nbsp;<A class=activeMain onClick="javascript:showTitle('mainMyData');showMenu('mainMisDatos', 'MisDatos', 1)" href="MyData.asp" target=principal><B>&nbsp;&nbsp;&nbsp;Mis&nbsp;datos&nbsp;&nbsp;&nbsp;</B></A></TD>
    <%if Session("OperatorLevel")=0 then %>
		<TD width="1" class=separator vAlign=middle align=right>|</TD>
    	<TD width="50" id=Operators class=menu vAlign=center align=middle>
				<A class=activeMain href="javascript:showMenu('mainOperators', 'Operators', 1)">&nbsp;&nbsp;&nbsp;Administradores&nbsp;&nbsp;&nbsp;</A></A>
		</TD>
    <%end if%>
		<TD width="1" class=separator vAlign=middle align=right>|</TD>
    <TD width="50" class=menu vAlign=center align=middle>
				<A class=activeMain href="javascript:if(%20confirm('Esta%20seguro%20que%20desea%20salir')%20)%20document.location%20='LogOff.asp';">&nbsp;&nbsp;Salir&nbsp;&nbsp;</A>
		</TD>
</TR>
<!--</TABLE>-->

<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainSetup0 style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>	
		<TD class=separator vAlign=center width="1" align=right>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('ItinerarioPendiente');"  href="ItineraryPends.asp?IT=1" target=principal>Pendientes</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('ItinerarioAsignado');"  href="ItineraryAsigs.asp?IT=1" target=principal>Asignados</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<!--<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('ItinerarioCobros');"  href="Search_Admin.asp?GID=29" target=principal>Cobros</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>-->
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevoBL');"  href="InsertData.asp?GID=1&IT=1" target=principal>Nueva&nbsp;CP</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarBL');" href="Search_Admin.asp?GID=1&IT=1" target=principal>Editar&nbsp;CP</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<%'2021-11-04 if InStr(Session("Countries"), "GT")>0 or InStr(Session("Countries"), "SV")>0 or InStr(Session("Countries"), "HN")>0 or InStr(Session("Countries"), "CR")>0 or InStr(Session("Countries"), "PA")>0  then%>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevoBLG');"  href="InsertData.asp?GID=22" target=principal>Nuevo&nbsp;Grupo</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarBLG');" href="Search_Admin.asp?GID=22&IT=1" target=principal>Editar&nbsp;Grupo</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<%'2021-11-04 end if%>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('CartaEndoso');" href="Search_Admin.asp?GID=14&IT=1" target=principal>Cobros&nbsp;y&nbsp;Documentos&nbsp;</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('Costos');" href="Search_Admin.asp?GID=29&IT=1" target=principal>Costos</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('CartaAcept');" href="Search_Admin.asp?GID=16&IT=1" target=principal>Carta&nbsp;Aceptaci&oacute;n</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<!--<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('Fianzas');" href="Search_Admin.asp?GID=13" target=principal>Fianzas</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>-->
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('ItinerarioLlegadas');"  href="Search_Admin.asp?GID=12&IT=1" target=principal>Salidas&nbsp;y&nbsp;Llegadas</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('Rastreo');" href="Search_Admin.asp?GID=23&IT=1" target=principal>Rastreo</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('ReporteDias');" href="Search_Admin.asp?GID=17&IT=1" target=principal>Reportes</a>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<!--<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('SolMov');" href="Search_Admin.asp?GID=19" target=principal>Solicitud&nbsp;Movimiento</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>-->
		<TD class=submenu vAlign=center width="1000" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>

<%if Session("OperatorLevel")=0 or Session("OperatorLevel")=1 then %>
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainSetup1 style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>	
		<TD class=separator vAlign=center width="400" align=right>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevoRemitente');"  href="InsertData.asp?GID=2" target=principal>Nuevo</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarRemitente');" href="Search_Admin.asp?GID=2" target=principal>Editar</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center width="1000" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainSetup2 style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>	
		<TD class=separator vAlign=center width="17%" align=right>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevoEmbarcador');"  href="InsertData.asp?GID=3" target=principal>Nuevo</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarEmbarcador');" href="Search_Admin.asp?GID=3" target=principal>Editar</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center width="100%" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainSetup3 style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>	
		<TD class=separator vAlign=center width="11%" align=right>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('AsignarDTI');"  href="Search_Admin.asp?GID=25" target=principal>Asignar&nbsp;/&nbsp;Editar</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevaPlantillaDTI');" href="InsertData.asp?GID=26" target=principal>Nueva&nbsp;Plantilla</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarPlantillaDTI');" href="Search_Admin.asp?GID=26" target=principal>Editar&nbsp;Plantilla</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center width="1000" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainSetup4 style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>	
		<TD class=separator vAlign=center width="60%" align=right>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevoConductor');"  href="InsertData.asp?GID=5" target=principal>Nuevo&nbsp;Conductor</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarConductor');" href="Search_Admin.asp?GID=5" target=principal>Editar&nbsp;Conductor</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevoCabezal');"  href="InsertData.asp?GID=6" target=principal>Nuevo&nbsp;Transporte</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarCabezal');"  href="Search_Admin.asp?GID=6" target=principal>Editar&nbsp;Transporte</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevoProveedor');"  href="InsertData.asp?GID=7" target=principal>Nuevo&nbsp;Proveedor</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarProveedor');"  href="Search_Admin.asp?GID=7" target=principal>Editar&nbsp;Proveedor</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>		
		<TD class=submenu vAlign=center width="1000" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainSetup5 style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>	
		<TD class=separator vAlign=center width="38%" align=right>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevaAduana');"  href="InsertData.asp?GID=8" target=principal>Nuevo</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarAduana');" href="Search_Admin.asp?GID=8" target=principal>Editar</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center width="5000" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainSetup6 style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>	
		<TD class=separator vAlign=center width="58%" align=right>|</TD>		
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevoAgente');"  href="InsertData.asp?GID=10" target=principal>Nuevo</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarAgente');" href="Search_Admin.asp?GID=10" target=principal>Editar</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center width="100%" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainSetup7 style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>	
		<TD class=separator vAlign=center width="52%" align=right>|</TD>		
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevoProducto');"  href="InsertData.asp?GID=9" target=principal>Nuevo</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarProducto');" href="Search_Admin.asp?GID=9" target=principal>Editar</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center width="2000" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>


<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainSetup12 style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>	
		<TD class=separator vAlign=center width="52%" align=right>|</TD>		
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);"  href="Search_Admin.asp?GID=35&IT=1" target=principal>Aduanas</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);" href="Search_Admin.asp?GID=34&IT=1" target=principal>Operativas</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center width="2000" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>


<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainSetup11 style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>	
		<TD class=separator vAlign=center width="52%" align=right>|</TD>		
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevoLogo');"  href="LogoColoaders.asp" target=principal>Nuevo</A>&nbsp;&nbsp;</TD>
		<!--<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarProducto');" href="Search_Admin.asp?GID=9" target=principal>Editar</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>-->
		<TD class=submenu vAlign=center width="2000" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainSetup8 style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>	
		<TD class=separator vAlign=center width="31%" align=right>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevoCliente');"  href="InsertData.asp?GID=11" target=principal>Nuevo</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarCliente');" href="Search_Admin.asp?GID=11" target=principal>Editar</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center width="100%" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>
<% if Session("OperatorLevel")=0 then%> 	
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainSetup9 style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>	
		<TD class=separator vAlign=center width="55%" align=right>|</TD>		
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('Varios');"  href="Setup.asp" target=principal>Varios</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevaCarta');" href="InsertData.asp?GID=10" target=principal>Nueva&nbsp;Carta</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarCarta');" href="Search_ResultsAdmin.asp?GID=10" target=principal>Editar&nbsp;Carta</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevaBodega');" href="InsertData.asp?GID=21" target=principal>Nueva&nbsp;Bodega</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarBodega');" href="Search_Admin.asp?GID=21" target=principal>Editar&nbsp;Bodega</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center width="1000" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>
<% end if %>
<%
end if
%>
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainSetup10 style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>	
		<TD class=separator vAlign=center width="200" align=right>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('ItinerarioPendiente');"  href="ItineraryPends.asp?IT=2" target=principal>Pendientes</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('ItinerarioAsignado');"  href="ItineraryAsigs.asp?IT=2" target=principal>Asignados</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevoBL');"  href="InsertData.asp?GID=1&IT=2" target=principal>Nueva&nbsp;CP</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarBL');" href="Search_Admin.asp?GID=1&IT=2" target=principal>Editar&nbsp;CP</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('CartaEndoso');" href="Search_Admin.asp?GID=14&IT=2" target=principal>Cobros&nbsp;y&nbsp;Documentos&nbsp;</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('Costos');" href="Search_Admin.asp?GID=29&IT=2" target=principal>Costos</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('ItinerarioLlegadas');"  href="Search_Admin.asp?GID=12&IT=2" target=principal>Llegadas</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('Rastreo');" href="Search_Admin.asp?GID=23&IT=2" target=principal>Rastreo</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('NuevoMarchamo');"  href="InsertData.asp?GID=32&IT=2" target=principal>Nuevo&nbsp;Marchamo</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('EditarMarchamo');" href="Search_Admin.asp?GID=32&IT=2" target=principal>Editar&nbsp;Marchamo</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('Reportes');" href="LocalReports.asp" target=principal>Reporte&nbsp;Marchamos</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="activateItem(this);showTitle('ReporteDias');" href="Search_Admin.asp?GID=17&IT=2" target=principal>Tiempos</a>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center width="1000" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainMisDatos style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0>
<TR>	
		<TD class=separator vAlign=center align=left>&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>

<%if Session("OperatorLevel")=0 or Session("OperatorLevel")=1 then%>
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainOperators style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto; ">
<TABLE border=0 cellSpacing=0 cellPadding=0 class=submenu>
<TR>
		<TD class=separator vAlign=center width="88%" align=right>|</TD>		
		<TD id=TDOperator1 class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="javascript:showTitle('NuevoEditor');" href="OPerators.asp" target="principal">Nuevo</A>&nbsp;&nbsp;</TD>
		<TD class=separator vAlign=center align=left>|</TD>
		<TD id=TDOperator2 class=submenu vAlign=center align=left>&nbsp;&nbsp;<A class=activeMain onClick="javascript:showTitle('EditarEditor');" href="Search_Operators.asp" target="principal">Editar</A>&nbsp;&nbsp;</TD>
        <TD class=separator vAlign=center align=left>|</TD>
		<TD class=submenu vAlign=center width="2000" align=left>&nbsp;&nbsp;</TD>
</TR>
</TABLE>
</DIV>
</TD></TR>
<% end if%>
<!--Menu de encargado -->
<TR><TD colspan="10" bgcolor="#FFFFFF">
<DIV id=mainEditores style="LEFT: 25px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 65px; HEIGHT: auto">
<TABLE cellSpacing=0 cellPadding=0 bgcolor=#ff7900>

  <TR>
    <TD class=inactiveMain align=left>&nbsp;
	<A class=activeMain onmouseover="JavaScript:showMain('mainEditores')" onmouseout="JavaScript:hideMain('mainEditores')" href="#" target=principal target=principal>Nuevo</A> &nbsp;
	</TD></TR>
  <TR>
    <TD class=inactiveMain align=left>&nbsp;
	<A class=activeMain onmouseover="JavaScript:showMain('mainEditores')" onmouseout="JavaScript:hideMain('mainEditores')" href="#" target=principal target=principal>Editar</A> &nbsp;
	</TD></TR>
  <TR>
    <TD class=inactiveMain align=left>&nbsp;
	<A class=activeMain onmouseover="JavaScript:showMain('mainEditores')" onmouseout="JavaScript:hideMain('mainEditores')" href="#" target=principal target=principal>Eliminar</A> &nbsp;
	</TD></TR>
</TABLE>
</DIV>
</TD></TR>
</TABLE>

<DIV id=NuevoBL style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nueva Carta Porte</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=NuevoBLG style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nueva Carta Porte Grupo</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarBL style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Editar Carta Porte</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarBLG style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Editar Carta Porte Grupo</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=Itinerario style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Itinerario</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=Fianzas style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Fianzas</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=CartaAcept style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Carta Aceptaci&oacute;n</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=Rastreo style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Rastreo</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=Reportes style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Reportes</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=AsignarDTI style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>Asignar&nbsp;/&nbsp;Editar&nbsp;DTI</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=NuevaPlantillaDTI style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>Nueva&nbsp;Plantilla&nbsp;DTI</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarPlantillaDTI style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>Editar&nbsp;Plantilla&nbsp;DTI</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=CartaEndoso style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Cobros y Documentos</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=Costos style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Costos</TD>
	</TR>
</TABLE>
</DIV>


<DIV id=CartaEntrega style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Carta Entrega</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=CartaRecol style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Carta Recolecci&oacute;n</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=SolMov style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Solicitud de Movimiento</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=NuevoCliente style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nuevo Consignatario</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarCliente style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Editar Consignatario</TD>
	</TR>
</TABLE>
</DIV>



<DIV id=NuevoRemitente style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nuevo Exportador</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarRemitente style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Editar Exportador</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=NuevoEmbarcador style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nuevo Exportador / Embarcador</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarEmbarcador style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Editar Exportador / Embarcador</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=NuevoDestinatario style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nuevo Destinatario</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarDestinatario style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Editar Destinatario</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=NuevoConductor style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nuevo Conductor</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarConductor style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Editar Conductor</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=NuevoCabezal style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nuevo Cabezal</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarCabezal style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Editar Cabezal</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=NuevoProveedor style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nuevo Proveedor</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarProveedor style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Editar Proveedor</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=NuevaAduana style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nueva Aduana</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarAduana style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Editar Aduana</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=NuevoAgente style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nuevo Agente</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarAgente style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Editar Agente</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=NuevoProducto style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nuevo Producto</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarProducto style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Editar Producto</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=Varios style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Configuraci&oacute;n</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=NuevaCarta style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nueva Carta</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarCarta style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Editar Carta</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=NuevaBodega style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nueva Bodega</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=EditarBodega style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Editar Bodega</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=ItinerarioPendiente style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Pendientes de Asignar Itinerario</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=ItinerarioAsignado style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Itinerarios Asignados</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=ItinerarioLlegadas style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Llegadas de Carga</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=ItinerarioCobros style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Cobros</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=TextBlank style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0>
  <TR>
		<TD class=title vAlign=center width="100%" align=center>&nbsp;</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=NuevoEditor style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Nuevo Administrador</TD>
	</TR>
</TABLE>
</DIV>


<DIV id=EditarEditor style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Configurar Administrador</TD>
	</TR>
</TABLE>
</DIV>

<DIV id=mainMyData style="LEFT: 0px; VISIBILITY: hidden; WIDTH: auto; POSITION: absolute; TOP: 85px; HEIGHT: auto;">
<TABLE cellSpacing=0 cellPadding=0 width="100%">
  <TR>
		<TD class=title vAlign=center width="100%" align=left>&nbsp;Mis Datos</TD>
	</TR>
</TABLE>
</DIV>

</BODY></HTML>
<%
Else
Response.Redirect "redirect.asp?MS=4"
End if
%>
