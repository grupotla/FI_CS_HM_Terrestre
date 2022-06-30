<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #INCLUDE file="utils.asp" -->
<%
Checking "0|1|2"
Dim NumberPos, TC

NumberPos = CheckNum(Request("N"))
TC = Round(CheckNum(Request("TC")),4)

%>
<HTML>
<HEAD><TITLE>Aimar - Terrestre</TITLE>
<META http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<SCRIPT language="javascript" src="img/matchvalues.js"></SCRIPT>
<SCRIPT language="javascript" src="img/vals.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
	function validar(Action) {
	    var ConvertValue = 0;
        var Currency = 'USD';

        if (!valTxt(document.forma.elements["CostValue"], 1, 5)){return (false)};
        
        if (document.forma.Currency.value == 'NIO') {
	        ConvertValue = document.forma.CostValue.value / <%=TC%>;
        } else {
            ConvertValue = document.forma.CostValue.value;
        }
        
        top.opener.document.forma.elements["C"+<%=NumberPos%>].value = Currency;
        top.opener.document.forma.elements["CT"+<%=NumberPos%>].value = Math.round(ConvertValue*Math.pow(10,2))/Math.pow(10,2);
		this.close(); 
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
-->
</style>
<BODY text=#000000 vLink=#000000 aLink=#000000 link=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="self.focus();">
	<TABLE cellspacing=0 cellpadding=2 align=center>
	<FORM name="forma" method="post">
		<TD colspan="2" class=label align=center>
		<table width="80%" border="0">
			<TR><TD class=label align=center colspan=2><br><b>Ingreso de Costos:</b></TD></TR> 
		  <tr>
			<td align="right" class="style4">
				<select class='style10' name='Currency' id="Moneda">
				<option value='USD'>USD</option>
                <option value='NIO'>NIO</option>
				</select>
			</td>
			<td align="left" class="style4">
				<input type="text" size="10" class="style10" name="CostValue" value="" onKeyUp="res(this,numb);" id="Costos">
			</td>            				
		  </tr>
          <tr>
              <TD class=label align=center colspan=2><INPUT name=enviar type=button onClick="JavaScript:validar(1);" value="&nbsp;&nbsp;Ingresar&nbsp;&nbsp;" class=label></TD>
          </tr>
		</table>
		<TD>
		</TR>
	</FORM>
	</TABLE>
</BODY>
</HTML>